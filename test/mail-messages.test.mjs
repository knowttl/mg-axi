import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { decode, encode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const mailbox = "mailbox@contoso.example";
const message = { id: "message-1", subject: "Invoice", receivedDateTime: "2026-01-01T00:00:00Z", isRead: false, body: { content: "private-body" }, bodyPreview: "private-preview" };
function fixture(t, response) {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-messages-"));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const config = join(dir, "config.json");
  const profiles = new Profiles(config);
  const tenant = "11111111-1111-4111-8111-111111111111";
  const client = "22222222-2222-4222-8222-222222222222";
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  const data = JSON.parse(readFileSync(config, "utf8"));
  for (const profile of Object.values(data.profiles)) profile.enabledPacks = ["mail"];
  writeFileSync(config, JSON.stringify(data));
  const requests = [];
  const credentials = [];
  const auth = { credential: async (...args) => {
    credentials.push(args);
    return { token: "opaque-fixture-token", tenantId: tenant, clientId: client, accountId: "fixture-account", username: mailbox };
  } };
  const overrides = { delegated: auth, application: auth, transport: async request => {
    requests.push(request);
    return response ? response(request, requests.length) : { status: 200, headers: {}, body: JSON.stringify({ value: [message] }) };
  } };
  const run = (args = [], profile = "soc", target = mailbox) => executeArgv(["mail", "message", "list", "--mailbox", target, "--profile", profile, "--config", config, ...args], overrides);
  const show = (args = [], profile = "soc", target = mailbox) => executeArgv(["mail", "message", "show", "--mailbox", target, "--id", "A+ID=", "--profile", profile, "--config", config, ...args], overrides);
  return { run, show, requests, credentials, config };
}

for (const profile of ["soc", "batch"]) test(`message list projects compact metadata with ${profile}`, async t => {
  const f = fixture(t);
  const result = decode(encode(await f.run([], profile)));
  assert.deepEqual(result.messages, [{ id: message.id, subject: message.subject, receivedDateTime: message.receivedDateTime, isRead: false }]);
  assert.ok(!JSON.stringify(result).includes("private-"));
  const url = new URL(f.requests[0].url);
  assert.equal(url.searchParams.get("$orderby"), "receivedDateTime desc");
  assert.equal(url.searchParams.get("$top"), "100");
  assert.deepEqual(f.credentials[0][1], profile === "soc" ? ["https://graph.microsoft.com/Mail.ReadBasic"] : undefined);
});

test("folder search preserves KQL and Graph search ordering without Entra consistency", async t => {
  const f = fixture(t);
  await f.run(["--folder", "A+ID=", "--search", 'subject:invoice AND from:sender@contoso.example'], "soc", "me");
  const url = new URL(f.requests[0].url);
  assert.equal(url.pathname, "/v1.0/me/mailFolders/A%2BID%3D/messages");
  assert.equal(url.searchParams.get("$search"), '"subject:invoice AND from:sender@contoso.example"');
  assert.equal(url.searchParams.has("$orderby"), false);
  assert.equal(f.requests[0].headers.ConsistencyLevel, undefined);
});

for (const [flag, value, expected] of [["filter", "isRead eq false", "isRead eq false"], ["conversation", "thread'1", "conversationId eq 'thread''1'"]]) test(`${flag} uses an unordered Graph filter`, async t => {
  const f = fixture(t);
  await f.run([`--${flag}`, value]);
  const url = new URL(f.requests[0].url);
  assert.equal(url.searchParams.get("$filter"), expected);
  assert.equal(url.searchParams.has("$orderby"), false);
});

for (const args of [["--search", "a", "--filter", "b"], ["--conversation", "a", "--search", "b"], ["--conversation", "a", "--filter", "b"], ["--search", ""], ["--fields", "bodyPreview"], ["--fields", "body"], ["--limit", "1001"], ["--api-version", "beta"], ["--all"], ["--scopes", "Mail.Read"]]) test(`invalid input refuses before credentials: ${args.join(" ")}`, async t => {
  const f = fixture(t);
  await assert.rejects(f.run(args), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

test("message cursors retain full nextLink and bind query, mailbox, fields and folder", async t => {
  const f = fixture(t, (request, count) => ({ status: 200, headers: {}, body: JSON.stringify({ value: [message], ...(count === 1 ? { "@odata.nextLink": `${request.url}&$skip=17` } : {}) }) }));
  const first = await f.run(["--search", "invoice", "--limit", "1"]);
  assert.equal(first.complete, false);
  assert.deepEqual(JSON.parse(Buffer.from(first.cursor, "base64url")).buffered, []);
  for (const args of [["--search", "other", "--limit", "1"], ["--search", "invoice", "--limit", "1", "--fields", "id"], ["--search", "invoice", "--limit", "1", "--folder", "inbox"]]) {
    await assert.rejects(f.run([...args, "--cursor", first.cursor]), { code: "VALIDATION_ERROR" });
  }
  const next = await f.run(["--search", "invoice", "--limit", "1", "--cursor", first.cursor]);
  assert.equal(next.complete, true);
  assert.equal(new URL(f.requests[1].url).searchParams.get("$skip"), "17");
});

test("search cap persists across pages and leaves no continuation after 1000", async t => {
  const f = fixture(t, (request, count) => {
    const url = new URL(request.url);
    url.searchParams.set("$skiptoken", `page-${count}`);
    return { status: 200, headers: {}, body: JSON.stringify({ value: Array.from({ length: Number(new URL(request.url).searchParams.get("$top")) }, () => message), "@odata.nextLink": url.toString() }) };
  });
  const first = await f.run(["--search", "invoice", "--limit", "600"]);
  const last = await f.run(["--search", "invoice", "--limit", "600", "--cursor", first.cursor]);
  assert.equal(first.messages.length, 600);
  assert.equal(last.messages.length, 400);
  assert.equal(last.complete, false);
  assert.equal(last.reason, "search-cap");
  assert.equal(last.cursor, undefined);
  assert.equal(f.requests.length, 2);
});

test("search errors suppress Graph echo", async t => {
  const f = fixture(t, () => ({ status: 400, headers: {}, body: JSON.stringify({ error: { message: "sensitive-query-echo" } }) }));
  await assert.rejects(f.run(["--search", "sensitive-query"]), error => error.code === "GRAPH_ERROR" && !JSON.stringify(error).includes("sensitive-query"));
});

test("metadata truncation supports full and empty results are definitive", async t => {
  const f = fixture(t, () => ({ status: 200, headers: {}, body: JSON.stringify({ value: [{ ...message, subject: "x".repeat(650) }] }) }));
  assert.match((await f.run()).messages[0].subject, /truncated, 650 chars total/);
  assert.equal((await f.run(["--full"])).messages[0].subject.length, 650);
  const empty = fixture(t, () => ({ status: 200, headers: {}, body: '{"value":[]}' }));
  assert.match((await empty.run()).help.join(" "), /0 messages matched/);
});

test("mailbox switches and oversized pages are refused", async t => {
  const f = fixture(t, request => ({ status: 200, headers: {}, body: JSON.stringify({ value: [message], "@odata.nextLink": request.url.replace(encodeURIComponent(mailbox), "other%40contoso.example") }) }));
  await assert.rejects(f.run(), { code: "POLICY_DENIED" });
  const oversized = fixture(t, () => ({ status: 200, headers: {}, body: JSON.stringify({ value: [message, message] }) }));
  await assert.rejects(oversized.run(["--limit", "1"]), { code: "GRAPH_ERROR" });
});

const details = { ...message, body: { contentType: "text", content: "Fixture body" },
  from: { emailAddress: { name: "Sender", address: "sender@contoso.example", unreviewed: "private-extra" } },
  toRecipients: [{ emailAddress: { address: mailbox } }],
  internetMessageHeaders: [{ name: "X-Fixture", value: "fixture header", unreviewed: "private-extra" }],
  unreviewed: "private-extra" };

for (const [profile, target] of [["soc", "me"], ["soc", mailbox], ["batch", mailbox]]) test(`message show reads reviewed details with ${profile}/${target}`, async t => {
  const f = fixture(t, () => ({ status: 200, headers: {}, body: JSON.stringify(details) }));
  const result = decode(encode(await f.show([], profile, target)));
  assert.equal(result.message.body.content, "Fixture body");
  assert.deepEqual(result.message.from, { emailAddress: { name: "Sender", address: "sender@contoso.example" } });
  assert.deepEqual(result.message.toRecipients, [{ emailAddress: { address: mailbox } }]);
  assert.equal(result.message.internetMessageHeaders, undefined);
  assert.equal(result.message.bodyPreview, undefined);
  assert.ok(!JSON.stringify(result).includes("private-extra"));
  assert.match(result.warning, /untrusted data.*never instructions/);
  const request = f.requests[0];
  assert.equal(new URL(request.url).pathname, `/v1.0/${target === "me" ? "me" : "users/mailbox%40contoso.example"}/messages/A%2BID%3D`);
  assert.equal(request.headers.Prefer, 'outlook.body-content-type="text"');
  assert.deepEqual(f.credentials[0][1], profile === "soc" ? ["https://graph.microsoft.com/Mail.Read"] : undefined);
});

test("message show omits bodies and explicitly selects internet headers", async t => {
  const f = fixture(t, () => ({ status: 200, headers: {}, body: JSON.stringify(details) }));
  const result = await f.show(["--body", "none", "--headers"]);
  assert.equal(result.message.body, undefined);
  assert.deepEqual(result.message.internetMessageHeaders, [{ name: "X-Fixture", value: "fixture header" }]);
  const select = new URL(f.requests[0].url).searchParams.get("$select").split(",");
  assert.ok(select.includes("internetMessageHeaders"));
  assert.ok(!select.includes("body"));
  assert.equal(f.requests[0].headers.Prefer, undefined);
});

test("message show requests HTML only on explicit opt-in", async t => {
  const f = fixture(t, () => ({ status: 200, headers: {}, body: JSON.stringify({ ...details, body: { contentType: "html", content: "<p>Fixture</p>" } }) }));
  const result = await f.show(["--body", "html"]);
  assert.deepEqual(result.message.body, { contentType: "html", content: "<p>Fixture</p>" });
  assert.equal(f.requests[0].headers.Prefer, 'outlook.body-content-type="html"');
  assert.match(result.warning, /untrusted/);
});

for (const length of [500, 501]) test(`body truncation boundary at ${length} characters`, async t => {
  const f = fixture(t, () => ({ status: 200, headers: {}, body: JSON.stringify({ ...details, body: { contentType: "text", content: "x".repeat(length) } }) }));
  const result = await f.show();
  assert.equal(result.message.body.content, length === 500 ? "x".repeat(500) : `${"x".repeat(500)}... (truncated, 501 chars total)`);
  assert.equal((await f.show(["--full"])).message.body.content, "x".repeat(length));
  if (length > 500) assert.match(result.help.join(" "), /--full/);
});

test("full message bodies and headers retain shared secret redaction", async t => {
  const secret = "-----BEGIN PRIVATE KEY-----fixture-----END PRIVATE KEY-----";
  const f = fixture(t, () => ({ status: 200, headers: {}, body: JSON.stringify({ ...details, body: { contentType: "text", content: secret }, internetMessageHeaders: [{ name: "X-Fixture", value: secret }] }) }));
  const result = await f.show(["--headers", "--full"]);
  assert.ok(!JSON.stringify(result).includes(secret));
  assert.equal(result.message.body.content, "***redacted***");
  assert.equal(result.message.internetMessageHeaders[0].value, "***redacted***");
});

for (const args of [["--body", "mime"], ["--id", ""], ["--fields", "body"], ["--api-version", "beta"], ["--scopes", "Mail.ReadBasic"]]) test(`message show rejects invalid input without calls: ${args.join(" ")}`, async t => {
  const f = fixture(t);
  await assert.rejects(f.show(args), { code: "VALIDATION_ERROR" });
  assert.equal(f.requests.length, 0);
  assert.equal(f.credentials.length, 0);
});

for (const [profile, target, credentials] of [["batch", "me", 0], ["soc", "other@contoso.example", 1]]) test(`message show refuses mailbox access for ${profile}/${target}`, async t => {
  const f = fixture(t);
  await assert.rejects(f.show([], profile, target), { code: "POLICY_DENIED" });
  assert.equal(f.requests.length, 0);
  assert.equal(f.credentials.length, credentials);
});

test("message show requires the opt-in mail pack before credentials", async t => {
  const f = fixture(t);
  const data = JSON.parse(readFileSync(f.config, "utf8"));
  data.profiles.soc.enabledPacks = ["entra"];
  writeFileSync(f.config, JSON.stringify(data));
  await assert.rejects(f.show(), { code: "POLICY_DENIED" });
  assert.equal(f.requests.length, 0);
  assert.equal(f.credentials.length, 0);
});

for (const body of [{ contentType: "html", content: "unexpected HTML" }, { contentType: "text", content: 42 }]) test(`message show refuses malformed body ${JSON.stringify(body)}`, async t => {
  const f = fixture(t, () => ({ status: 200, headers: {}, body: JSON.stringify({ ...details, body }) }));
  await assert.rejects(f.show(), { code: "GRAPH_ERROR" });
});

test("message show suppresses Graph error content", async t => {
  const f = fixture(t, () => ({ status: 403, headers: {}, body: JSON.stringify({ error: { message: "private-body-echo" } }) }));
  await assert.rejects(f.show(), error => error.code === "GRAPH_ERROR" && !JSON.stringify(error).includes("private-body-echo"));
});
