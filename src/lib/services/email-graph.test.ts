import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildGraphPayload,
  createGraphSender,
  readGraphConfig,
  type GraphConfig,
  type GraphFetch,
} from "./email-graph.ts";
import { createEmailService, fitsAsAttachments } from "./email.ts";

// What is asked of Microsoft when mail goes out through Microsoft 365.
// Nothing here talks to Microsoft: fetch is replaced by a script of answers.

const ENV = {
  EMAIL_DRIVER: "graph",
  MICROSOFT_GRAPH_CLIENT_ID: "client-id",
  MICROSOFT_GRAPH_CLIENT_SECRET: "client-secret",
  MICROSOFT_GRAPH_SENDER: "sign@acme.test",
  MICROSOFT_GRAPH_TENANT_ID: "tenant-id",
};

const CONFIG = readGraphConfig(ENV) as GraphConfig;

type Answer = {
  body?: string;
  headers?: Record<string, string>;
  status: number;
};

const TOKEN: Answer = {
  body: JSON.stringify({ access_token: "token-1", expires_in: 3600 }),
  status: 200,
};
const ACCEPTED: Answer = { status: 202 };

type Request = {
  body: string;
  headers: Record<string, string>;
  url: string;
};

/** A fetch that gives these answers, in order, and keeps what it was asked. */
function scriptedFetch(answers: Answer[]) {
  const requests: Request[] = [];
  const fetch: GraphFetch = async (url, init) => {
    requests.push({ body: init.body, headers: init.headers, url });
    const answer = answers.shift();

    if (!answer) {
      throw new Error(`Unexpected request to ${url}`);
    }

    return {
      headers: {
        get: (name: string) => answer.headers?.[name.toLowerCase()] ?? null,
      },
      ok: answer.status >= 200 && answer.status < 300,
      status: answer.status,
      text: async () => answer.body ?? "",
    };
  };

  return { fetch, requests };
}

const MESSAGE = {
  from: "Acme <sign@acme.test>",
  htmlBody: "<p>Please sign.</p>",
  replyTo: "help@acme.test",
  subject: "Please sign",
  textBody: "Please sign.",
  to: "signer@example.test",
};

test("the configuration needs all four variables", () => {
  assert.equal(
    readGraphConfig({ ...ENV, MICROSOFT_GRAPH_SENDER: undefined }),
    null,
  );
  assert.equal(
    readGraphConfig({ ...ENV, MICROSOFT_GRAPH_TENANT_ID: "" }),
    null,
  );
  assert.equal(CONFIG.sender, "sign@acme.test");
  assert.equal(CONFIG.graphHost, "https://graph.microsoft.com");
});

test("a sender with a display name is used as its bare address", () => {
  const config = readGraphConfig({
    ...ENV,
    MICROSOFT_GRAPH_HOST: "https://graph.microsoft.us/",
    MICROSOFT_GRAPH_SENDER: "Acme <Sign@Acme.test>",
  });

  assert.equal(config?.sender, "sign@acme.test");
  assert.equal(config?.graphHost, "https://graph.microsoft.us");
});

test("the message is sent as HTML from the mailbox, with its display name", () => {
  const { message, saveToSentItems } = buildGraphPayload(MESSAGE);

  assert.equal(saveToSentItems, true);
  assert.deepEqual(message.body, {
    content: "<p>Please sign.</p>",
    contentType: "HTML",
  });
  assert.deepEqual(message.from, {
    emailAddress: { address: "sign@acme.test", name: "Acme" },
  });
  assert.deepEqual(message.replyTo, [
    { emailAddress: { address: "help@acme.test" } },
  ]);
  assert.deepEqual(message.toRecipients, [
    { emailAddress: { address: "signer@example.test" } },
  ]);
  assert.equal("attachments" in message, false);
});

test("a message without HTML is sent as plain text", () => {
  const { message } = buildGraphPayload({ ...MESSAGE, htmlBody: undefined });

  assert.deepEqual(message.body, {
    content: "Please sign.",
    contentType: "Text",
  });
});

test("attachments are sent as base64 file attachments", () => {
  const { message } = buildGraphPayload({
    ...MESSAGE,
    attachments: [
      {
        content: new TextEncoder().encode("%PDF-1.7"),
        contentType: "application/pdf",
        filename: "Contract - signed.pdf",
      },
    ],
  });

  assert.deepEqual(message.attachments, [
    {
      "@odata.type": "#microsoft.graph.fileAttachment",
      contentBytes: Buffer.from("%PDF-1.7").toString("base64"),
      contentType: "application/pdf",
      name: "Contract - signed.pdf",
    },
  ]);
});

test("the app signs in once and sends from the configured mailbox", async () => {
  const { fetch, requests } = scriptedFetch([TOKEN, ACCEPTED, ACCEPTED]);
  const sender = createGraphSender({
    config: CONFIG,
    fetch,
    wait: async () => undefined,
  });

  await sender.send(MESSAGE);
  await sender.send(MESSAGE);

  assert.equal(requests.length, 3);
  assert.equal(
    requests[0]?.url,
    "https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token",
  );

  const form = new URLSearchParams(requests[0]?.body);

  assert.equal(form.get("grant_type"), "client_credentials");
  assert.equal(form.get("scope"), "https://graph.microsoft.com/.default");
  assert.equal(form.get("client_id"), "client-id");
  assert.equal(
    requests[1]?.url,
    "https://graph.microsoft.com/v1.0/users/sign%40acme.test/sendMail",
  );
  assert.equal(requests[1]?.headers.authorization, "Bearer token-1");
  assert.equal(requests[2]?.headers.authorization, "Bearer token-1");
});

test("a token is renewed before it runs out", async () => {
  let clock = 0;
  const { fetch, requests } = scriptedFetch([
    TOKEN,
    ACCEPTED,
    {
      body: JSON.stringify({ access_token: "token-2", expires_in: 3600 }),
      status: 200,
    },
    ACCEPTED,
  ]);
  const sender = createGraphSender({
    config: CONFIG,
    fetch,
    now: () => clock,
    wait: async () => undefined,
  });

  await sender.send(MESSAGE);
  clock = 3_550_000;
  await sender.send(MESSAGE);

  assert.equal(requests[3]?.headers.authorization, "Bearer token-2");
});

test("a token Graph refuses is replaced once", async () => {
  const { fetch, requests } = scriptedFetch([
    TOKEN,
    { body: "{}", status: 401 },
    TOKEN,
    ACCEPTED,
  ]);
  const sender = createGraphSender({
    config: CONFIG,
    fetch,
    wait: async () => undefined,
  });

  await sender.send(MESSAGE);

  assert.equal(requests.length, 4);
});

test("a message Graph throttles is sent again after the wait it names", async () => {
  const waits: number[] = [];
  const { fetch, requests } = scriptedFetch([
    TOKEN,
    { headers: { "retry-after": "2" }, status: 429 },
    ACCEPTED,
  ]);
  const sender = createGraphSender({
    config: CONFIG,
    fetch,
    wait: async (milliseconds) => {
      waits.push(milliseconds);
    },
  });

  await sender.send(MESSAGE);

  assert.deepEqual(waits, [2000]);
  assert.equal(requests.length, 3);
});

test("throttling that does not clear is reported after a few tries", async () => {
  const throttled: Answer = {
    body: JSON.stringify({
      error: { code: "ApplicationThrottled", message: "Too many requests." },
    }),
    headers: { "retry-after": "60" },
    status: 429,
  };
  const waits: number[] = [];
  const { fetch } = scriptedFetch([
    TOKEN,
    throttled,
    throttled,
    throttled,
    throttled,
  ]);
  const sender = createGraphSender({
    config: CONFIG,
    fetch,
    wait: async (milliseconds) => {
      waits.push(milliseconds);
    },
  });

  await assert.rejects(sender.send(MESSAGE), {
    message: "ApplicationThrottled: Too many requests.",
    name: "GraphEmailError",
  });
  // Never longer than a few seconds, whatever Graph asks for.
  assert.deepEqual(waits, [5000, 5000, 5000]);
});

test("a mailbox the app may not use is reported in Graph's words", async () => {
  const { fetch } = scriptedFetch([
    TOKEN,
    {
      body: JSON.stringify({
        error: {
          code: "ErrorAccessDenied",
          message: "Access is denied. Check credentials and try again.",
        },
      }),
      status: 403,
    },
  ]);
  const sender = createGraphSender({
    config: CONFIG,
    fetch,
    wait: async () => undefined,
  });

  await assert.rejects(sender.send(MESSAGE), {
    message:
      "ErrorAccessDenied: Access is denied. Check credentials and try again.",
  });
});

test("a failed sign-in is reported without its trace ids", async () => {
  const { fetch } = scriptedFetch([
    {
      body: JSON.stringify({
        error: "invalid_client",
        error_description:
          "AADSTS7000215: Invalid client secret provided.\r\nTrace ID: 1\r\nCorrelation ID: 2",
      }),
      status: 401,
    },
  ]);
  const sender = createGraphSender({
    config: CONFIG,
    fetch,
    wait: async () => undefined,
  });

  await assert.rejects(sender.send(MESSAGE), {
    message:
      "Microsoft sign-in failed: AADSTS7000215: Invalid client secret provided.",
  });
});

test("EMAIL_DRIVER=graph sends through Graph from the resolved sender", async () => {
  const { fetch, requests } = scriptedFetch([TOKEN, ACCEPTED, ACCEPTED]);
  const service = createEmailService({
    env: ENV,
    graphFetch: fetch,
    resolveAddresses: async () => ({
      senderAddress: "Acme <sign@acme.test>",
      supportAddress: "help@acme.test",
    }),
  });

  const results = await service.sendBatch([
    {
      subject: "One",
      template: "signing_invitation",
      textBody: "1",
      to: "a@example.test",
    },
    {
      subject: "Two",
      template: "signing_invitation",
      textBody: "2",
      to: "b@example.test",
    },
  ]);

  assert.deepEqual(results, [
    { accepted: ["a@example.test"], rejected: [] },
    { accepted: ["b@example.test"], rejected: [] },
  ]);

  const sent = JSON.parse(requests[1]?.body ?? "{}") as {
    message: { replyTo: unknown };
  };

  assert.deepEqual(sent.message.replyTo, [
    { emailAddress: { address: "help@acme.test" } },
  ]);
});

test("EMAIL_DRIVER=graph without its settings names the missing ones", () => {
  assert.throws(
    () =>
      createEmailService({
        env: { ...ENV, MICROSOFT_GRAPH_CLIENT_SECRET: undefined },
      }),
    { message: /MICROSOFT_GRAPH_CLIENT_SECRET/ },
  );
});

test("Graph takes smaller attachments than Resend", () => {
  const attachments = [
    {
      content: new Uint8Array(5 * 1024 * 1024),
      contentType: "application/pdf",
      filename: "big.pdf",
    },
  ];

  assert.equal(fitsAsAttachments(attachments, ENV), false);
  assert.equal(
    fitsAsAttachments(attachments, { EMAIL_DRIVER: "resend" }),
    true,
  );
});
