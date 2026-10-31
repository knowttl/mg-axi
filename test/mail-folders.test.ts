import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { encode, decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles, validateProfile } from "../dist/profiles.js";
import { GraphSession, resolveSessionOperation } from "../dist/graph-session.js";
import type { TransportRequest, TransportResponse } from "../dist/graph-session.js";
import type { TestContext } from "node:test";
import { record, records, text, strings, authFakes, type CredentialStub } from "./fixtures/typed.ts";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const user = "33333333-3333-4333-8333-333333333333";
const mailbox = "mailbox@contoso.example";
const folder = { id: "inbox", displayName: "Inbox", totalItemCount: 12, unreadItemCount: 2, parentFolderId: "root", childFolderCount: 0, isHidden: false, messages: [{ body: "never-output" }] };
function fixture(t: TestContext, { enabled = true, delegatedIdentity = { localAccountId: user, username: mailbox }, body = { value: [folder] }, response }: {
  enabled?: boolean;
  delegatedIdentity?: { localAccountId?: string; username?: string };
  body?: unknown;
  response?: (request: TransportRequest, count: number) => TransportResponse;
} = {}) {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-mail-"));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const config = join(dir, "config.json");
  const profiles = new Profiles(config);
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  const value: { profiles: Record<string, { enabledPacks: string[] }> } = JSON.parse(readFileSync(config, "utf8"));
  if (enabled) for (const profile of Object.values(value.profiles)) profile.enabledPacks = ["mail"];
  writeFileSync(config, JSON.stringify(value));
  const requests: TransportRequest[] = [];
  const credentials: { mode: string; args: Parameters<CredentialStub["credential"]> }[] = [];
  const auth = (mode: string): CredentialStub => ({ credential: async (...args) => {
    credentials.push({ mode, args });
    return { token: "opaque-fixture-token", tenantId: tenant, clientId: client, accountId: user, ...(mode === "delegated" ? delegatedIdentity : {}), expiresAt: Date.now() + 3600000 };
  } });
  const overrides = { delegated: authFakes(auth("delegated")).delegated, application: authFakes(auth("application")).application, transport: async (request: TransportRequest) => {
    requests.push(request);
    return response ? response(request, requests.length) : { status: 200, headers: {}, body: JSON.stringify(body) };
  } };
  const run = (args: string[] = [], profile = "soc", action = "list") => executeArgv(["mail", "folder", action, "--config", config, "--profile", profile, "--mailbox", mailbox, ...args], overrides);
  return { run, requests, credentials, config, profiles, overrides };
}

for (const profile of ["soc", "batch"]) test(`folder list uses least privilege and reviewed TOON fields with ${profile}`, async t => {
  const f = fixture(t);
  const result = record(decode(encode(record(await f.run([], profile)))));
  assert.equal(result.complete, true);
  assert.deepEqual(result.folders, [{ id: "inbox", displayName: "Inbox", totalItemCount: 12, unreadItemCount: 2 }]);
  const url = new URL(f.requests[0]!.url);
  assert.equal(url.pathname, `/v1.0/users/${encodeURIComponent(mailbox)}/mailFolders`);
  assert.equal(url.searchParams.get("$top"), "100");
  assert.equal(url.searchParams.get("$select"), "id,displayName,totalItemCount,unreadItemCount");
  assert.deepEqual(f.credentials[0]!.args[1], profile === "soc" ? ["https://graph.microsoft.com/Mail.ReadBasic"] : undefined);
  assert.ok(!JSON.stringify(result).includes("never-output"));
});

test("folder show projects all reviewed metadata and encodes folder IDs", async t => {
  const f = fixture(t, { body: folder });
  const result = record(await f.run(["--id", "A+ID="], "batch", "show"));
  assert.deepEqual(result.folder, { id: "inbox", displayName: "Inbox", totalItemCount: 12, unreadItemCount: 2, parentFolderId: "root", childFolderCount: 0, isHidden: false });
  assert.equal(new URL(f.requests[0]!.url).pathname, `/v1.0/users/${encodeURIComponent(mailbox)}/mailFolders/A%2BID%3D`);
});

test("delegated me uses the reviewed alternate route", async t => {
  const f = fixture(t);
  await executeArgv(["mail", "folder", "list", "--mailbox", "me", "--config", f.config, "--profile", "soc"], f.overrides);
  assert.equal(new URL(f.requests[0]!.url).pathname, "/v1.0/me/mailFolders");
});

for (const ownMailbox of [user, "MAILBOX@CONTOSO.EXAMPLE"]) test(`explicit delegated mailbox matches authenticated account metadata: ${ownMailbox}`, async t => {
  const f = fixture(t);
  await executeArgv(["mail", "folder", "list", "--mailbox", ownMailbox, "--config", f.config, "--profile", "soc"], f.overrides);
  assert.equal(f.requests.length, 1);
});

test("folder paging returns a row-free cursor and resumes on the same resource", async t => {
  const f = fixture(t, { response: (request, count) => ({ status: 200, headers: {}, body: JSON.stringify(count === 1 ? { value: [folder], "@odata.nextLink": `${request.url}&$skiptoken=next` } : { value: [{ ...folder, id: "sentitems" }] }) }) });
  const first = record(await f.run(["--limit", "1"]));
  assert.equal(first.complete, false);
  assert.deepEqual(JSON.parse(Buffer.from(text(first.cursor), "base64url").toString()).buffered, []);
  const next = record(await f.run(["--limit", "1", "--cursor", text(first.cursor)]));
  assert.equal(next.complete, true);
  assert.equal(records(next.folders)[0]!.id, "sentitems");
  assert.equal(new URL(f.requests[1]!.url).searchParams.get("$skiptoken"), "next");
});

test("folder text truncates at 500 and full retains redaction", async t => {
  const f = fixture(t, { body: { ...folder, displayName: "x".repeat(650), parentFolderId: "https://contoso.example/?sig=synthetic-secret" } });
  const result = record(await f.run(["--id", "inbox"], "soc", "show"));
  assert.match(text(record(result.folder).displayName), /truncated, 650 chars total/);
  assert.match(strings(result.help).join(" "), /--full/);
  const full = record(await f.run(["--id", "inbox", "--full"], "soc", "show"));
  assert.equal(text(record(full.folder).displayName).length, 650);
  assert.equal(record(full.folder).parentFolderId, "***redacted***");
});

test("empty folder list is definitive", async t => {
  const f = fixture(t, { body: { value: [] } });
  const result = record(await f.run());
  assert.equal(result.complete, true);
  assert.match(strings(result.help).join(" "), /0 folders returned/);
});

for (const args of [["--limit", "0"], ["--limit", "1001"], ["--limit", "1.5"], ["--fields", "messages"], ["--fields", ""], ["--api-version", "beta"], ["--search", "secret"], ["--scopes", "https://graph.microsoft.com/Mail.ReadWrite"], ["--all"]]) test(`invalid folder flags refuse before credentials: ${args.join(" ")}`, async t => {
  const f = fixture(t);
  await assert.rejects(f.run(args), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const action of ["list", "show"]) test(`folder ${action} reuses explicit Mail.Read consent`, async t => {
  const f = fixture(t, { response: () => ({ status: 200, headers: {}, body: JSON.stringify(action === "list" ? { value: [folder] } : folder) }) });
  const result = record(await f.run(["--scopes", "https://graph.microsoft.com/Mail.Read", ...(action === "show" ? ["--id", "inbox"] : [])], "soc", action));
  assert.equal(action === "list" ? records(result.folders)[0]!.displayName : record(result.folder).displayName, "Inbox");
  assert.deepEqual(f.credentials[0]!.args[1], ["https://graph.microsoft.com/Mail.Read"]);
});

test("mail pack stays opt-in and Entra remains the profile default", async t => {
  const f = fixture(t, { enabled: false });
  assert.deepEqual(record(f.profiles.resolve("soc").profile).enabledPacks, ["entra"]);
  await assert.rejects(f.run(), { code: "POLICY_DENIED" });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const bad of ["me", "other@contoso.example", "../messages"]) test(`application me and invalid or other delegated mailbox refusal: ${bad}`, async t => {
  const f = fixture(t);
  await assert.rejects(executeArgv(["mail", "folder", "list", "--mailbox", bad, "--config", f.config, "--profile", bad === "me" ? "batch" : "soc"], f.overrides));
  assert.equal(f.requests.length, 0);
});

for (const delegatedIdentity of [{}, { localAccountId: "other-id", username: "other@contoso.example" }]) test(`missing or mismatched authenticated mailbox metadata is refused: ${JSON.stringify(delegatedIdentity)}`, async t => {
  const f = fixture(t, { delegatedIdentity });
  await assert.rejects(f.run(), { code: "POLICY_DENIED" });
  assert.equal(f.requests.length, 0);
});

test("application mail credentials stay opaque and are never printed", async t => {
  const f = fixture(t);
  f.overrides.application = authFakes({ credential: async () => ({ token: "opaque-fixture-secret", tenantId: tenant, clientId: client }) }).application;
  const result = record(await f.run([], "batch"));
  assert.equal(f.requests.length, 1);
  assert.ok(!JSON.stringify(result).includes("opaque-fixture-secret"));
});

test("folder list hides all keys when Graph returns a malformed envelope", async t => {
  const f = fixture(t, { body: { PrivateMailSubject: { PrivateMailbox: "private-content" } } });
  await assert.rejects(f.run(), error => {
    assert.equal(record(error).code, "RESPONSE_INVALID");
    assert.deepEqual(record(record(error).diagnostics).shape, [{ path: "$", type: "object" }, { path: '$["<key>"]', type: "object" }, { path: '$["<key>"]["<key>"]', type: "string" }]);
    assert.doesNotMatch(JSON.stringify(error), /PrivateMailSubject|PrivateMailbox|private-content/);
    return true;
  });
});

test("folder metadata rejects nested objects instead of displaying unreviewed mail content", async t => {
  const f = fixture(t, { body: { value: [{ ...folder, displayName: { body: "private" } }] } });
  await assert.rejects(f.run(), { code: "GRAPH_ERROR" });
});

test("mail scope cannot be bypassed through direct collection calls", async t => {
  const f = fixture(t);
  const session = new GraphSession(f.overrides);
  await assert.rejects(session.collect({ profile: f.profiles.resolve("soc").profile, operation: resolveSessionOperation("v1.0", "GET", "/me/mailFolders"), scopes: ["https://graph.microsoft.com/User.Read.All"] }), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

test("mail pack does not open raw mail reads or sensitive expansion", async t => {
  const f = fixture(t);
  for (const args of [["/users/mailbox@contoso.example/mailFolders"], ["/users", "--odata", "$expand=mailFolders"]]) {
    await assert.rejects(executeArgv(["api", "get", ...args, "--config", f.config, "--profile", "soc", "--scopes", "https://graph.microsoft.com/User.Read.All"], f.overrides));
  }
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

test("a continuation cannot switch the bound mailbox", async t => {
  const f = fixture(t, { body: { value: [folder], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/other@contoso.example/mailFolders?$skiptoken=next" } });
  await assert.rejects(f.run(), { code: "POLICY_DENIED" });
  assert.equal(f.requests.length, 1);
});

test("oversized folder pages are refused instead of buffering mail rows", async t => {
  const f = fixture(t, { body: { value: [folder, folder] } });
  await assert.rejects(f.run(["--limit", "1"]), { code: "GRAPH_ERROR" });
  assert.equal(f.requests.length, 1);
});

test("mail-only profiles validate without a second sensitive-area switch", async t => {
  const f = fixture(t);
  assert.deepEqual(validateProfile(f.profiles.resolve("soc").profile).enabledPacks, ["mail"]);
  await assert.rejects(executeArgv(["entra", "user", "list", "--config", f.config, "--profile", "soc"], f.overrides), { code: "POLICY_DENIED" });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});
