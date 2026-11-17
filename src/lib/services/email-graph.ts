import { readEnv, type EnvSource } from "../utils/env.ts";
import { extractEmailAddress } from "./email-addresses.ts";

// Sending through Microsoft 365 (Exchange Online) with the Microsoft Graph
// API, for organizations whose mail already lives there and that would rather
// not bring in another mail provider.
//
// The app signs in as itself (OAuth client credentials) and sends from one
// mailbox, usually a shared mailbox such as sign@acme.com, with
// POST /users/{mailbox}/sendMail. The app registration needs the Mail.Send
// application permission, and it should only be allowed to use that one
// mailbox: an Exchange "RBAC for Applications" role assignment scoped to the
// mailbox, or an application access policy. Without that, Mail.Send lets it
// send as anyone in the tenant. docs/self-hosting.md has the steps.
// https://learn.microsoft.com/graph/api/user-sendmail

export const GRAPH_ENV_KEYS = [
  "MICROSOFT_GRAPH_TENANT_ID",
  "MICROSOFT_GRAPH_CLIENT_ID",
  "MICROSOFT_GRAPH_CLIENT_SECRET",
  "MICROSOFT_GRAPH_SENDER",
] as const;

/**
 * Graph takes at most 4 MB in one sendMail request, and attachments are
 * base64 encoded in it, which makes them a third larger. Above this the
 * message goes out with links only, as it does with Resend above its limit.
 * https://learn.microsoft.com/graph/api/resources/fileattachment
 */
export const GRAPH_MAX_ATTACHMENT_BYTES = 2_900_000;

const DEFAULT_AUTHORITY_HOST = "https://login.microsoftonline.com";
const DEFAULT_GRAPH_HOST = "https://graph.microsoft.com";

/** How often a request Graph turned away for rate is tried again. */
const RATE_LIMIT_RETRIES = 3;
const DEFAULT_RATE_LIMIT_WAIT_MS = 1_000;
const MAX_RATE_LIMIT_WAIT_MS = 5_000;

/** A token is renewed this long before Microsoft says it runs out. */
const TOKEN_EXPIRY_MARGIN_MS = 60_000;

export type GraphConfig = {
  authorityHost: string;
  clientId: string;
  clientSecret: string;
  graphHost: string;
  /** The mailbox mail is sent from, as a bare address. */
  sender: string;
  tenantId: string;
};

/** The configuration, or null when any of the four variables is missing. */
export function readGraphConfig(env: EnvSource): GraphConfig | null {
  const tenantId = readEnv(env, "MICROSOFT_GRAPH_TENANT_ID");
  const clientId = readEnv(env, "MICROSOFT_GRAPH_CLIENT_ID");
  const clientSecret = readEnv(env, "MICROSOFT_GRAPH_CLIENT_SECRET");
  const sender = readEnv(env, "MICROSOFT_GRAPH_SENDER");

  if (!tenantId || !clientId || !clientSecret || !sender) {
    return null;
  }

  return {
    // Only national clouds need these two; see docs/self-hosting.md.
    authorityHost: (
      readEnv(env, "MICROSOFT_GRAPH_AUTHORITY_HOST") ?? DEFAULT_AUTHORITY_HOST
    ).replace(/\/+$/, ""),
    clientId,
    clientSecret,
    graphHost: (
      readEnv(env, "MICROSOFT_GRAPH_HOST") ?? DEFAULT_GRAPH_HOST
    ).replace(/\/+$/, ""),
    sender: extractEmailAddress(sender),
    tenantId,
  };
}

export type GraphAttachment = {
  content: Uint8Array;
  contentType: string;
  filename: string;
};

export type GraphMessage = {
  attachments?: readonly GraphAttachment[];
  /** "Acme <sign@acme.com>": the name is shown, the address is the mailbox. */
  from: string;
  htmlBody?: string;
  replyTo?: string;
  subject: string;
  textBody: string;
  to: string;
};

export type GraphFetch = (
  input: string,
  init: {
    body: string;
    headers: Record<string, string>;
    method: "POST";
  },
) => Promise<{
  headers: { get(name: string): string | null };
  ok: boolean;
  status: number;
  text(): Promise<string>;
}>;

export class GraphEmailError extends Error {
  readonly status: number;

  constructor(message: string, status: number) {
    super(message);
    this.name = "GraphEmailError";
    this.status = status;
  }
}

function displayNameOf(value: string): string | undefined {
  const match = /^\s*(.*?)\s*<[^<>]+>\s*$/.exec(value);

  return match?.[1] ? match[1] : undefined;
}

function recipient(value: string) {
  const name = displayNameOf(value);

  return {
    emailAddress: {
      address: extractEmailAddress(value),
      ...(name ? { name } : {}),
    },
  };
}

/** The sendMail request body for one message. */
export function buildGraphPayload(message: GraphMessage) {
  return {
    message: {
      ...(message.attachments?.length
        ? {
            attachments: message.attachments.map((attachment) => ({
              "@odata.type": "#microsoft.graph.fileAttachment",
              contentBytes: Buffer.from(attachment.content).toString("base64"),
              contentType: attachment.contentType,
              name: attachment.filename,
            })),
          }
        : {}),
      body: message.htmlBody
        ? { content: message.htmlBody, contentType: "HTML" }
        : { content: message.textBody, contentType: "Text" },
      from: recipient(message.from),
      ...(message.replyTo ? { replyTo: [recipient(message.replyTo)] } : {}),
      subject: message.subject,
      toRecipients: [recipient(message.to)],
    },
    // A copy in the mailbox's Sent Items is a record of what went out.
    saveToSentItems: true,
  };
}

// Graph and the identity platform both answer errors as JSON, with the
// useful words in different places.
function describeFailure(body: string, status: number): string {
  try {
    const parsed = JSON.parse(body) as {
      error?: { code?: string; message?: string } | string;
      error_description?: string;
    };

    if (typeof parsed.error === "object" && parsed.error?.message) {
      return parsed.error.code
        ? `${parsed.error.code}: ${parsed.error.message}`
        : parsed.error.message;
    }

    if (parsed.error_description) {
      // The identity platform adds trace and correlation ids on new lines.
      return parsed.error_description.split(/\r?\n/)[0] ?? "";
    }
  } catch {
    // Not JSON: fall through to the status.
  }

  return `Microsoft Graph answered with status ${String(status)}.`;
}

function retryAfterMs(value: string | null): number {
  const seconds = Number(value);

  return Number.isFinite(seconds) && seconds > 0
    ? Math.min(seconds * 1000, MAX_RATE_LIMIT_WAIT_MS)
    : DEFAULT_RATE_LIMIT_WAIT_MS;
}

type GraphSenderOptions = {
  config: GraphConfig;
  fetch?: GraphFetch;
  now?: () => number;
  wait: (milliseconds: number) => Promise<void>;
};

export type GraphSender = {
  send(message: GraphMessage): Promise<void>;
};

export function createGraphSender({
  config,
  fetch: fetchImpl = fetch as unknown as GraphFetch,
  now = Date.now,
  wait,
}: GraphSenderOptions): GraphSender {
  let token: { expiresAt: number; value: string } | null = null;

  async function getToken(): Promise<string> {
    if (token && token.expiresAt > now()) {
      return token.value;
    }

    const response = await fetchImpl(
      `${config.authorityHost}/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`,
      {
        body: new URLSearchParams({
          client_id: config.clientId,
          client_secret: config.clientSecret,
          grant_type: "client_credentials",
          scope: `${config.graphHost}/.default`,
        }).toString(),
        headers: { "content-type": "application/x-www-form-urlencoded" },
        method: "POST",
      },
    );
    const body = await response.text();

    if (!response.ok) {
      throw new GraphEmailError(
        `Microsoft sign-in failed: ${describeFailure(body, response.status)}`,
        response.status,
      );
    }

    const parsed = JSON.parse(body) as {
      access_token?: string;
      expires_in?: number;
    };

    if (!parsed.access_token) {
      throw new GraphEmailError(
        "Microsoft sign-in returned no access token.",
        response.status,
      );
    }

    token = {
      expiresAt:
        now() + (parsed.expires_in ?? 3600) * 1000 - TOKEN_EXPIRY_MARGIN_MS,
      value: parsed.access_token,
    };

    return token.value;
  }

  return {
    async send(message) {
      const url = `${config.graphHost}/v1.0/users/${encodeURIComponent(config.sender)}/sendMail`;
      const body = JSON.stringify(buildGraphPayload(message));

      for (let attempt = 0; ; attempt += 1) {
        const response = await fetchImpl(url, {
          body,
          headers: {
            authorization: `Bearer ${await getToken()}`,
            "content-type": "application/json",
          },
          method: "POST",
        });

        // 202 Accepted: Exchange has the message and will deliver it.
        if (response.ok) {
          return;
        }

        const text = await response.text();

        // A token revoked or rotated early: get a new one and try once more.
        if (response.status === 401 && attempt === 0) {
          token = null;
          continue;
        }

        // Throttled. Graph names the wait; the message is not queued, so it
        // would be lost if it were not tried again.
        if (
          (response.status === 429 || response.status === 503) &&
          attempt < RATE_LIMIT_RETRIES
        ) {
          await wait(retryAfterMs(response.headers.get("retry-after")));
          continue;
        }

        throw new GraphEmailError(
          describeFailure(text, response.status),
          response.status,
        );
      }
    },
  };
}
