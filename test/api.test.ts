import { record, records, text as outputText, strings, delegatedProvider, type RawResponseHandler } from "./fixtures/typed.ts";
import type { GraphTransport, TransportRequest, TransportResponse } from "../dist/graph-session.js";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { decode } from "@toon-format/toon";
import { ApplicationAuth } from "../dist/app-auth.js";
import { DelegatedAuth } from "../dist/auth.js";
import { Profiles } from "../dist/profiles.js";
import { fetchTransport, runApiGet } from "../dist/api.js";
import { GraphSession, MAX_CURSOR_BYTES, rawReadOperation } from "../dist/graph-session.js";
import { executeArgv } from "../dist/cli.js";

// API-01 through the lowest real interface observing behavior: runApiGet with
// real DelegatedAuth/ApplicationAuth and a fake credential/transport boundary
// (the same seam CORE-01/CORE-02 tests use), plus CLI subprocess checks for
// grammar and refusal precedence.

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, token: { stored: true } } as const;
const appProfile = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, auth: "federated", token: { env: "MG_AXI_TOKEN" } } as const;
const scopes = "https://graph.microsoft.com/User.Read.All";

function json(status: number, body: unknown, headers: Record<string, string> = {}): TransportResponse {
  return { status, headers, body: JSON.stringify(body) };
}

function fixture(handler?: RawResponseHandler) {
  const credentialCalls: unknown[][] = [];
  const requests: TransportRequest[] = [];
  const credential = <T extends object>(account: T) => ({ token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, ...account });
  const delegated = new DelegatedAuth(delegatedProvider({
    storage: "session-only",
    login: async () => credential({ accountId: "synthetic-account" }),
    silent: async (...args) => { credentialCalls.push(["silent", ...args]); return credential({ accountId: "synthetic-account" }); },
  }));
  const application = new ApplicationAuth({
    storage: "session-only",
    acquire: async (...args) => { credentialCalls.push(["acquire", ...args]); return credential({}); },
  });
  const transport: GraphTransport = async request => {
    requests.push(request);
    const response = typeof handler === "function" ? await handler(request, requests.length) : handler;
    return { headers: {}, body: "", ...response } as TransportResponse;
  };
  return { credentialCalls, requests, deps: { delegated, application, transport } };
}

const read = (overrides: Partial<import("../dist/api.js").ApiGetArgs> = {}, handler: RawResponseHandler = json(200, { value: [] })) => {
  const f = fixture(handler);
  return { ...f, run: (args: Partial<import("../dist/api.js").ApiGetArgs>) => runApiGet({ path: "/users", apiVersion: "v1.0", profile: delegatedProfile, scopes, ...overrides, ...args }, f.deps) };
};

test("any v1.0 path is sent as written to the Graph host", async () => {
  const f = read({ path: "/groups/abc/members/unknownSegment/deeper" }, json(200, { value: [] }));
  await f.run({});
  assert.equal(f.requests[0]!.url, "https://graph.microsoft.com/v1.0/groups/abc/members/unknownSegment/deeper");
});

test("beta reads go to the beta endpoint on a preview profile", async () => {
  const f = read({ path: "/users", apiVersion: "beta", profile: { ...delegatedProfile, preview: true } }, json(200, { value: [] }));
  await f.run({});
  assert.equal(new URL(f.requests[0]!.url).pathname, "/beta/users");
});

for (const path of ["/security/alerts_v2", "/%73ecurity/alerts_v2"]) test(`raw security route ${path} requires the security pack`, async () => {
  const denied = read({ path });
  await assert.rejects(denied.run({}), error => /Pack security/.test(outputText(record(error).message)));
  assert.equal(denied.credentialCalls.length, 0);
  const allowed = read({ path, profile: { ...delegatedProfile, enabledPacks: ["security"] } });
  assert.equal((await allowed.run({})).complete, true);
  const direct = fixture(json(200, { id: "alert" }));
  const session = new GraphSession(direct.deps);
  const operation = rawReadOperation("v1.0", path);
  await assert.rejects(session.execute({ profile: delegatedProfile, operation, raw: true, scopes: [scopes] }), error => /Pack security/.test(outputText(record(error).message)));
  assert.deepEqual(await session.execute({ profile: { ...delegatedProfile, enabledPacks: ["security"] }, operation, raw: true, scopes: [scopes] }), { id: "alert" });
});

test("raw security writes require the security pack", async () => {
  const f = fixture();
  const session = new GraphSession(f.deps);
  const profile = { ...appProfile, rawWrites: true, writes: { allowWrites: true, operations: [] } } as const;
  const args = { profile, method: "PATCH" as const, version: "v1.0", path: "/security/alerts_v2/a", transport: async () => ({ status: 204, headers: {}, body: "" }) };
  await assert.rejects(session.rawWrite(args), error => /Pack security/.test(outputText(record(error).message)));
  assert.equal((await session.rawWrite({ ...args, profile: { ...profile, enabledPacks: ["security"] } })).kind, "dry-run");
});

test("secret-bearing responses are redacted on the raw path", async () => {
  const f = read({ path: "/directory/deviceLocalCredentials" }, json(200, { value: [{ id: "a", credentials: [{ passwordBase64: "c2VjcmV0" }] }] }));
  const result = await f.run({});
  assert.doesNotMatch(JSON.stringify(result), /c2VjcmV0/);
});

test("raw credentials retain metadata while masking secrets in reads and writes", async () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-raw-credentials-"));
  try {
    const body = { keyCredentials: [{ keyId: "key-1", type: "AsymmetricX509Cert", usage: "Verify", key: "key-material" }], passwordCredentials: [{ keyId: "password-1", hint: "abc", secretText: "secret-value" }] };
    const expected = { keyCredentials: [{ keyId: "key-1", type: "AsymmetricX509Cert", usage: "Verify", key: "***redacted***" }], passwordCredentials: [{ keyId: "password-1", hint: "abc", secretText: "***redacted***" }] };
    const f = read({ path: "/applications/a" }, json(200, body));
    assert.deepEqual(await f.run({}), expected);
    const session = new GraphSession(f.deps);
    const args = { profile: { ...appProfile, rawWrites: true, writes: { allowWrites: true, operations: [] } }, method: "PATCH" as const, version: "v1.0", path: "/applications/a", payload: body, transport: async () => ({ status: 200, headers: {}, body: JSON.stringify(body) }), journalPath: join(dir, "writes.log") };
    const preview = await session.rawWrite(args);
    assert.deepEqual(JSON.parse(preview.preview.proposedChange!), expected);
    const result = await session.rawWrite({ ...args, execute: true, confirm: "/applications/a" });
    assert.equal(result.kind, "success");
    if (result.kind === "success") assert.deepEqual(result.response, expected);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("secret key-value pairs stay redacted in reads, cursors and writes", async () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-pair-secret-"));
  try {
    const pair = { key: "SecretToken", value: "private-value" };
    const f = read({ path: "/servicePrincipals/a/synchronization/secrets", limit: 1 }, json(200, { value: [{ id: "first" }, { id: "second", nested: [pair] }] }));
    const first = await f.run({});
    assert.doesNotMatch(JSON.stringify(first), /private-value/);
    const resumed = await f.run({ cursor: String(first.cursor) });
    assert.deepEqual(records(resumed.value)[0]!.nested, [{ key: "SecretToken", value: "***redacted***" }]);
    const session = new GraphSession(f.deps);
    const args = { profile: { ...appProfile, rawWrites: true, writes: { allowWrites: true, operations: [] } }, method: "PUT" as const, version: "v1.0", path: "/servicePrincipals/a/synchronization/secrets", payload: { value: [pair] }, transport: async () => ({ status: 200, headers: {}, body: JSON.stringify({ value: [pair] }) }), journalPath: join(dir, "writes.log") };
    const preview = await session.rawWrite(args);
    assert.deepEqual(JSON.parse(preview.preview.proposedChange!), { value: [{ key: "SecretToken", value: "***redacted***" }] });
    const result = await session.rawWrite({ ...args, execute: true, confirm: args.path });
    assert.equal(result.kind, "success");
    if (result.kind === "success") assert.deepEqual(result.response, { value: [{ key: "SecretToken", value: "***redacted***" }] });
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("raw nested expansions survive direct reads and continuation resumes", async () => {
  const expand = "members($select=id;$expand=owners($select=id))";
  const f = read({ path: "/groups", query: `$expand=${encodeURIComponent(expand)}`, limit: 1 }, (_request, count) => count === 1
    ? json(200, { value: [{ id: "a" }], "@odata.nextLink": `https://graph.microsoft.com/v1.0/groups?$skiptoken=next&$expand=${encodeURIComponent(expand)}` })
    : json(200, { value: [{ id: "b" }] }));
  const partial = await f.run({});
  assert.equal(new URL(f.requests[0]!.url).searchParams.get("$expand"), expand);
  assert.deepEqual((await f.run({ cursor: String(partial.cursor) })).value, [{ id: "b" }]);
  assert.equal(new URL(f.requests[1]!.url).searchParams.get("$expand"), expand);
  const direct = fixture(json(200, { id: "group" }));
  const session = new GraphSession(direct.deps);
  assert.deepEqual(await session.execute({ profile: delegatedProfile, operation: rawReadOperation("v1.0", "/groups/a"), raw: true, scopes: [scopes], query: { $expand: expand } }), { id: "group" });
});

for (const expand of ["members($expand=messages)", "members($expand=%6dessages)", "members($expand=drive)"]) test(`raw nested expansion ${expand} respects sensitive policy`, async () => {
  const f = read({ path: "/groups", query: `$expand=${encodeURIComponent(expand)}` });
  await assert.rejects(f.run({}), error => record(error).code === "POLICY_DENIED");
  assert.equal(f.requests.length, 0);
});

test("recovery keys remain redacted in singleton and buffered collection output", async () => {
  const singleton = read({ path: "/informationProtection/bitlocker/recoveryKeys/a" }, json(200, { id: "a", key: "recovery-secret" }));
  assert.deepEqual(await singleton.run({}), { id: "a", key: "***redacted***" });
  const collection = read({ path: "/informationProtection/bitlocker/recoveryKeys", limit: 1 }, json(200, { value: [{ id: "a", key: "first-secret" }, { id: "b", key: "second-secret" }] }));
  const partial = await collection.run({});
  assert.doesNotMatch(JSON.stringify(partial), /first-secret|second-secret/);
  const resumed = await collection.run({ cursor: String(partial.cursor) });
  assert.deepEqual(resumed.value, [{ id: "b", key: "***redacted***" }]);
});

test("oversized singleton returns a cursor and delivers the buffered object on resume", async () => {
  const f = read({ path: "/users/a" }, { ...json(200, { id: "a", displayName: "Adele" }), receivedBodyBytes: 5000001 });
  const partial = await f.run({});
  assert.equal(partial.complete, false);
  assert.match(String(partial.reason), /byte budget exceeded/);
  assert.equal(typeof partial.cursor, "string");
  assert.deepEqual(await f.run({ cursor: String(partial.cursor) }), { id: "a", displayName: "Adele" });
  assert.equal(f.requests.length, 1);
});

for (const value of [[], ["alpha", "beta"], [{ label: "first" }, { label: "second" }]] as const) test(`raw resource retains its array value ${JSON.stringify(value)}`, async () => {
  const f = read({ path: "/users/a/extensions/example", limit: 1 }, json(200, { id: "example", value, secretText: "hidden" }));
  const result = await f.run({});
  assert.equal(result.id, "example");
  assert.deepEqual(result.value, value);
  assert.equal(result.secretText, "***redacted***");
  assert.equal(result.cursor, undefined);
});

test("raw resource with an array value resumes intact after its byte budget", async () => {
  const body = { id: "example", value: ["alpha", "beta"] };
  const f = read({ path: "/users/a/extensions/example", limit: 1 }, { ...json(200, body), receivedBodyBytes: 5000001 });
  const partial = await f.run({});
  assert.equal(partial.complete, false);
  assert.equal(typeof partial.cursor, "string");
  assert.deepEqual(await f.run({ cursor: String(partial.cursor) }), { id: "example", value: ["alpha", "beta"] });
  assert.equal(f.requests.length, 1);
});

test("raw resource without an id retains its array value and other properties", async () => {
  const f = read({ path: "/users/a/extensions/example", limit: 1 }, json(200, { displayName: "example", value: ["alpha", "beta"] }));
  const result = await f.run({});
  assert.equal(result.displayName, "example");
  assert.deepEqual(result.value, ["alpha", "beta"]);
});

test("delegated raw mail refuses shared scopes and another mailbox", async () => {
  const profile = { ...delegatedProfile, enabledPacks: ["entra", "mail"] } as const;
  const shared = read({ path: "/users/other@example.com/messages", profile, scopes: "https://graph.microsoft.com/Mail.Read.Shared" });
  await assert.rejects(shared.run({}), error => record(error).code === "VALIDATION_ERROR");
  const other = read({ path: "/users/other@example.com/messages", profile, scopes: "https://graph.microsoft.com/Mail.Read" });
  await assert.rejects(other.run({}), error => record(error).code === "POLICY_DENIED");
  assert.equal(other.requests.length, 0);
  const own = read({ path: "/me/messages", profile, scopes: "https://graph.microsoft.com/Mail.Read" });
  assert.equal((await own.run({})).complete, true);
});

test("encoded raw mail paths retain mailbox restrictions", async () => {
  const f = read({ path: "/users/other@example.com/%6dessages", profile: { ...delegatedProfile, enabledPacks: ["entra", "mail"] }, scopes: "https://graph.microsoft.com/Mail.Read" });
  await assert.rejects(f.run({}), error => record(error).code === "POLICY_DENIED");
  assert.equal(f.requests.length, 0);
});

test("raw mail reads use the mail pack without the Entra pack", async () => {
  const f = read({ path: "/me/messages", profile: { ...delegatedProfile, enabledPacks: ["mail"] }, scopes: "https://graph.microsoft.com/Mail.Read" });
  assert.equal((await f.run({})).complete, true);
});

for (const action of ["sendMail", "%73endMail", "microsoft.graph.sendMail"]) test(`raw ${action} uses mail authorization`, async () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-raw-send-"));
  try {
    const f = fixture();
    const session = new GraphSession(f.deps);
    const path = `/users/other@example.com/${action}`;
    const profile = { ...delegatedProfile, enabledPacks: ["mail"], rawWrites: true, writes: { allowWrites: true, operations: [] } } as const;
    let sent = false;
    const args = { profile, method: "POST" as const, version: "v1.0", path, scopes: ["https://graph.microsoft.com/Mail.Send"], execute: true, confirm: path, transport: async () => { sent = true; return { status: 202, headers: {}, body: "" }; }, journalPath: join(dir, "writes.log") };
    await assert.rejects(session.rawWrite({ ...args, profile: { ...profile, enabledPacks: ["entra"] } }), error => /Pack mail/.test(outputText(record(error).message)));
    await assert.rejects(session.rawWrite({ ...args, scopes: ["https://graph.microsoft.com/Mail.Send.Shared"] }), error => record(error).code === "POLICY_DENIED");
    await assert.rejects(session.rawWrite(args), error => record(error).code === "POLICY_DENIED");
    const own = await session.rawWrite({ ...args, path: `/me/${action}`, confirm: `/me/${action}` });
    assert.equal(own.kind, "success");
    assert.equal(sent, true);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("temporary access passes are redacted in raw reads and write previews", async () => {
  const path = "/users/a/authentication/temporaryAccessPassMethods";
  const f = read({ path }, json(200, { value: [{ id: "a", temporaryAccessPass: "secret-pass" }] }));
  assert.deepEqual((await f.run({})).value, [{ id: "a", temporaryAccessPass: "***redacted***" }]);
  const session = new GraphSession(f.deps);
  const result = await session.rawWrite({ profile: { ...appProfile, rawWrites: true, writes: { allowWrites: true, operations: [] } }, method: "POST", version: "v1.0", path, payload: { temporaryAccessPass: "secret-pass" }, transport: async () => ({ status: 200, headers: {}, body: "{}" }) });
  assert.equal(result.kind, "dry-run");
  assert.doesNotMatch(JSON.stringify(result), /secret-pass/);
});

for (const path of ["/$batch", "/%24batch"]) test(`raw batch ${path} is refused before transport`, async () => {
  const f = fixture();
  const session = new GraphSession(f.deps);
  let sent = false;
  await assert.rejects(session.rawWrite({ profile: { ...delegatedProfile, rawWrites: true, writes: { allowWrites: true, operations: [] } }, method: "POST", version: "v1.0", path, scopes: ["https://graph.microsoft.com/Mail.ReadWrite.Shared"], payload: { requests: [{ id: "1", method: "PATCH", url: "/users/other/messages/a" }] }, execute: true, confirm: path, transport: async () => { sent = true; return { status: 200, headers: {}, body: "{}" }; } }), error => record(error).code === "POLICY_DENIED");
  assert.equal(sent, false);
  assert.equal(f.credentialCalls.length, 0);
});

for (const path of ["/$batch", "/%24batch"]) test(`raw batch ${path} reaches the preview when the profile sets rawBatch`, async () => {
  const f = fixture();
  const session = new GraphSession(f.deps);
  const profile = { ...appProfile, rawWrites: true, rawBatch: true, writes: { allowWrites: true, operations: [] } } as const;
  const preview = await session.rawWrite({ profile, method: "POST", version: "v1.0", path, payload: { requests: [{ id: "1", method: "GET", url: "/users" }] }, transport: async () => { throw new Error("preview must not send"); } });
  assert.equal(preview.kind, "dry-run");
  await assert.rejects(session.rawWrite({ profile: { ...profile, rawWrites: false }, method: "POST", version: "v1.0", path, transport: async () => ({ status: 200, headers: {}, body: "{}" }) }), error => record(error).code === "WRITES_DISABLED");
});

for (const [pack, path, otherPath] of [["mail", "/users/a/messages", "/security/alerts_v2"], ["security", "/security/alerts_v2", "/users/a/messages"]] as const) {
  for (const batchPath of ["/$batch", "/%24batch"]) test(`${pack} profile authorizes batch members independently for ${batchPath}`, async () => {
    const dir = mkdtempSync(join(tmpdir(), "mg-axi-batch-pack-"));
    try {
      const f = fixture();
      const session = new GraphSession(f.deps);
      const profile = { ...appProfile, enabledPacks: [pack], rawWrites: true, rawBatch: true, writes: { allowWrites: true, operations: [] } };
      let sends = 0;
      const args = { profile, method: "POST" as const, version: "v1.0", path: batchPath, execute: true, confirm: batchPath, journalPath: join(dir, "writes.log"), transport: async () => { sends++; return { status: 200, headers: {}, body: JSON.stringify({ responses: [{ id: "1", status: 200 }] }) }; } };
      assert.equal((await session.rawWrite({ ...args, payload: { requests: [{ id: "1", method: "GET", url: path }] } })).kind, "success");
      await assert.rejects(session.rawWrite({ ...args, payload: { requests: [{ id: "1", method: "GET", url: otherPath }] } }), { code: "POLICY_DENIED" });
      assert.equal(sends, 1);
    } finally { rmSync(dir, { recursive: true, force: true }); }
  });
}

for (const name of ["ConsistencyLevel", "consistencylevel", "CONSISTENCYLEVEL"]) test(`batch accepts ${name} header for advanced queries`, async () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-batch-header-"));
  try {
    const f = fixture();
    const session = new GraphSession(f.deps);
    const profile = { ...appProfile, rawWrites: true, rawBatch: true, writes: { allowWrites: true, operations: [] } };
    const result = await session.rawWrite({ profile, method: "POST", version: "v1.0", path: "/$batch", execute: true, confirm: "/$batch", journalPath: join(dir, "writes.log"), payload: { requests: [{ id: "1", method: "GET", url: "/users?$count=true", headers: { [name]: "eventual" } }] }, transport: async () => ({ status: 200, headers: {}, body: JSON.stringify({ responses: [{ id: "1", status: 200 }] }) }) });
    assert.equal(result.kind, "success");
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

for (const path of ["/drives/d/root", "/%64rives/d/root", "/shares/s/driveItem"]) test(`raw file route ${path} needs the files sensitive area`, async () => {
  const denied = read({ path });
  await assert.rejects(denied.run({}), error => record(error).code === "POLICY_DENIED" && /Sensitive area files/.test(outputText(record(error).message)));
  assert.equal(denied.credentialCalls.length, 0);
  const allowed = read({ path, profile: { ...delegatedProfile, sensitiveAreas: ["files"] } });
  assert.equal((await allowed.run({})).complete, true);
});

test("the files sensitive area does not admit raw mail routes", async () => {
  const f = read({ path: "/users/a/messages", profile: { ...delegatedProfile, sensitiveAreas: ["files"] } });
  await assert.rejects(f.run({}), error => /Pack mail/.test(outputText(record(error).message)));
  assert.equal(f.credentialCalls.length, 0);
});

test("raw collection preserves Graph metadata across pages", async () => {
  const f = read({ path: "/users/delta", limit: 1 }, json(200, { "@odata.count": 2, "@odata.deltaLink": "https://graph.microsoft.com/v1.0/users/delta?token=done", value: [{ id: "a" }, { id: "b" }] }));
  const partial = await f.run({});
  assert.equal(partial["@odata.count"], 2);
  const result = await f.run({ cursor: String(partial.cursor) });
  assert.equal(result["@odata.deltaLink"], "https://graph.microsoft.com/v1.0/users/delta?token=done");
  assert.deepEqual(result.value, [{ id: "b" }]);
});

test("raw delta queries reuse returned tokens and follow Graph continuation parameters", async () => {
  const deltaLink = "https://graph.microsoft.com/v1.0/users/delta?$deltatoken=next%2Btoken";
  const f = read({ path: "/users/delta" }, (_request, count) => count === 1
    ? json(200, { value: [{ id: "a" }], "@odata.deltaLink": deltaLink })
    : count === 2 ? json(200, { value: [{ id: "b" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/delta?$skiptoken=page2&custom=opaque" })
      : json(200, { value: [{ id: "c" }] }));
  const first = await f.run({});
  assert.equal(first["@odata.deltaLink"], deltaLink);
  const second = await f.run({ query: new URL(deltaLink).search.slice(1) });
  assert.deepEqual(second.value, [{ id: "b" }, { id: "c" }]);
  assert.equal(new URL(f.requests[1]!.url).searchParams.get("$deltatoken"), "next+token");
  assert.equal(new URL(f.requests[2]!.url).searchParams.get("$skiptoken"), "page2");
  assert.equal(new URL(f.requests[2]!.url).searchParams.get("custom"), "opaque");
});

test("raw query accepts an explicit skip token", async () => {
  const f = read({ path: "/users/delta", query: "$skiptoken=opaque%2Bvalue" });
  await f.run({});
  assert.equal(new URL(f.requests[0]!.url).searchParams.get("$skiptoken"), "opaque+value");
});

test("raw mail pages request only the remaining row allowance", async () => {
  const profile = { ...delegatedProfile, enabledPacks: ["mail"] } as const;
  const f = read({ path: "/me/messages", profile, scopes: "https://graph.microsoft.com/Mail.Read", limit: 100 }, (_request, count) => count === 1
    ? json(200, { value: Array.from({ length: 60 }, (_, i) => ({ id: String(i) })), "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=second&$top=100" })
    : count === 2 ? json(200, { value: Array.from({ length: 40 }, (_, i) => ({ id: String(i + 60) })), "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=third&$top=100" })
      : json(200, { value: [{ id: "100" }] }));
  const first = await f.run({});
  assert.equal(first.returned, 100);
  assert.equal(first.complete, false);
  assert.equal(new URL(f.requests[1]!.url).searchParams.get("$top"), "40");
  const resumed = await f.run({ cursor: String(first.cursor) });
  assert.deepEqual(resumed.value, [{ id: "100" }]);
  assert.equal(new URL(f.requests[2]!.url).searchParams.get("$top"), "100");
});

for (const odata of ["$count=true", '$search="displayName:x"']) test(`raw query ${odata} sends eventual consistency`, async () => {
  const f = read({ odata });
  assert.deepEqual(await f.run({}), { returned: 0, complete: true, value: [] });
  assert.equal(f.requests[0]!.headers.ConsistencyLevel, "eventual");
});

test("raw reads allow literal braces and uncatalogued roots", async () => {
  for (const path of ["/directoryObjects", "/domainDnsRecords", "/scopedRoleMemberships", "/users/{literal}"]) {
    const f = read({ path });
    await f.run({});
    assert.equal(new URL(f.requests[0]!.url).pathname, new URL(`https://graph.microsoft.com/v1.0${path}`).pathname);
  }
});

test("raw mail rejects oversized pages and protects attachment content", async () => {
  const profile = { ...delegatedProfile, enabledPacks: ["entra", "mail"] } as const;
  const args = { profile, scopes: "https://graph.microsoft.com/Mail.Read" };
  const f = read({ ...args, path: "/me/messages", limit: 1 }, json(200, { value: [{ id: "a" }, { id: "b" }] }));
  await assert.rejects(f.run({}), error => record(error).code === "GRAPH_ERROR");
  const attachment = read({ ...args, path: "/me/messages/a/attachments/b" }, json(200, { id: "b", contentBytes: "c2VjcmV0" }));
  assert.deepEqual(await attachment.run({}), { id: "b", contentBytes: "***redacted***" });
  const mime = read({ ...args, path: "/me/messages/a/$value" });
  await assert.rejects(mime.run({}), error => record(error).code === "POLICY_DENIED");
});

test("raw writes require opt-in, confirmation and a journal", async () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-raw-write-"));
  try {
    const f = fixture();
    const sent: string[] = [];
    const session = new GraphSession(f.deps);
    const args = { profile: { ...appProfile, rawWrites: true, writes: { allowWrites: true, operations: [] } }, method: "PATCH" as const, version: "v1.0", path: "/users/a", payload: { displayName: "Adele" }, transport: async (request: { method: string }) => { sent.push(request.method); return { status: 200, headers: {}, body: '{"id":"a","key":"secret","temporaryAccessPass":"pass-value"}' }; }, journalPath: join(dir, "writes.log") };
    assert.equal((await session.rawWrite(args)).kind, "dry-run");
    await assert.rejects(session.rawWrite({ ...args, execute: true }), error => record(error).code === "CONFIRM_REQUIRED");
    await assert.rejects(session.rawWrite({ ...args, profile: { ...args.profile, rawWrites: false }, execute: true, confirm: "/users/a" }), error => record(error).code === "WRITES_DISABLED");
    const result = await session.rawWrite({ ...args, execute: true, confirm: "/users/a" });
    assert.equal(result.kind, "success");
    if (result.kind === "success") assert.deepEqual(result.response, { id: "a", key: "***redacted***", temporaryAccessPass: "***redacted***" });
    assert.deepEqual(sent, ["PATCH"]);
    assert.deepEqual(readFileSync(join(dir, "writes.log"), "utf8").trim().split("\n").map(line => JSON.parse(line).kind), ["intent", "outcome"]);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("mail route without the mail pack is refused before credentials", async () => {
  const f = read({ path: "/me/messages" });
  await assert.rejects(f.run({}), error => /Pack mail is not enabled/.test(outputText(record(error).message)));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const pathname of ["https://graph.microsoft.com/v1.0/users", "/users?$top=2", "/users#frag", "users", ""] as const) test(`host/query/path escape ${pathname || "(empty)"} is refused`, async () => {
  const f = read({ path: pathname });
  await assert.rejects(f.run({}), error => record(error).code === "VALIDATION_ERROR" && /Refused raw path/.test(outputText(record(error).message)));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("an unsupported api version is refused before credentials", async () => {
  const f = read({ apiVersion: "v2" });
  await assert.rejects(f.run({}), error => record(error).code === "VALIDATION_ERROR" && /Unsupported raw api version v2/.test(outputText(record(error).message)));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const binding of ["..", "."] as const) test(`dot segment ${binding} never reaches credentials`, async () => {
  const f = read({ path: `/users/${binding}` });
  await assert.rejects(f.run({}), error => record(error).code === "VALIDATION_ERROR");
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("application profiles refuse delegated scopes before credentials", async () => {
  const f = fixture(json(200, { value: [] }));
  await assert.rejects(
    runApiGet({ path: "/users", apiVersion: "v1.0", profile: appProfile, scopes }, f.deps),
    error => record(error).code === "VALIDATION_ERROR" && /\.default audience/.test(outputText(record(error).message)),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const [path, body, result] of [
  ["/users/a", { id: "a" }, { id: "a" }],
  ["/users", { value: [{ id: "a" }] }, { returned: 1, complete: true, value: [{ id: "a" }] }],
] as const) test(`application read ${path} uses the configured audience without caller scopes`, async () => {
  const f = fixture(json(200, body));
  assert.deepEqual(await runApiGet({ path, apiVersion: "v1.0", profile: appProfile }, f.deps), result);
  assert.equal(f.credentialCalls.length, 1);
  assert.equal(f.credentialCalls[0]![0], "acquire");
  assert.deepEqual(f.credentialCalls[0]![2], ["https://graph.microsoft.com/.default"]);
  assert.equal(f.requests.length, 1);
  assert.equal(f.requests[0]!.headers.Authorization, "Bearer opaque-fixture-secret");
});

for (const path of ["/users/a", "/users"] as const) {
  for (const location of [
    "https://example.invalid/v1.0/users",
    "https://graph.microsoft.com/beta/users",
    "https://graph.microsoft.com/v1.0/groups",
  ] as const) test(`fetch-backed read ${path} refuses redirect to ${location}`, async t => {
    const requests: { url: string; options: RequestInit & { headers: Record<string, string> } }[] = [];
    t.mock.method(globalThis, "fetch", async (url: string, options: RequestInit & { headers: Record<string, string> }) => {
      requests.push({ url, options });
      return new Response("", { status: 302, headers: { Location: location } });
    });
    const f = fixture();
    await assert.rejects(
      runApiGet({ path, apiVersion: "v1.0", profile: delegatedProfile, scopes }, { ...f.deps, transport: fetchTransport }),
      { code: "POLICY_DENIED" },
    );
    assert.equal(requests.length, 1);
    assert.equal(requests[0]!.options.redirect, "manual");
  });
}

for (const [path, body, result] of [
  ["/users/a", { id: "a" }, { id: "a" }],
  ["/users", { value: [{ id: "a" }] }, { returned: 1, complete: true, value: [{ id: "a" }] }],
] as const) test(`fetch-backed read ${path} follows an authorized redirect`, async t => {
  const requests: { url: string; options: RequestInit & { headers: Record<string, string> } }[] = [];
  const target = `https://graph.microsoft.com/v1.0${path}?$select=id`;
  t.mock.method(globalThis, "fetch", async (url: string, options: RequestInit & { headers: Record<string, string> }) => {
    requests.push({ url, options });
    return requests.length === 1
      ? new Response("", { status: 302, headers: { Location: target } })
      : new Response(JSON.stringify(body), { status: 200 });
  });
  const f = fixture();
  assert.deepEqual(await runApiGet({ path, apiVersion: "v1.0", profile: delegatedProfile, scopes }, { ...f.deps, transport: fetchTransport }), result);
  assert.equal(requests.length, 2);
  assert.equal(new URL(requests[1]!.url).pathname, new URL(target).pathname);
  assert.equal(new URL(requests[1]!.url).searchParams.get("$select"), "id");
  assert.equal(requests[0]!.options.redirect, "manual");
  assert.equal(requests[1]!.options.redirect, "manual");
  assert.equal(requests[1]!.options.headers.Authorization, "Bearer opaque-fixture-secret");
});

test("collection read returns redacted rows with truthful completion", async () => {
  const f = read({}, json(200, { value: [{ id: "a", displayName: "AccountKey=fixture-secret" }] }));
  assert.deepEqual(await f.run({}), {
    returned: 1,
    complete: true,
    value: [{ id: "a", displayName: "***redacted***" }],
  });
  assert.equal(f.requests.length, 1);
  assert.equal(f.requests[0]!.method, "GET");
  assert.equal(new URL(f.requests[0]!.url).pathname, "/v1.0/users");
  assert.equal(f.requests[0]!.headers.Authorization, "Bearer opaque-fixture-secret");
});

test("$select travels on the authorized URL", async () => {
  const f = read({ odata: "$select=id,displayName" }, json(200, { value: [{ id: "a" }] }));
  await f.run({});
  assert.equal(f.requests[0]!.url, "https://graph.microsoft.com/v1.0/users?%24select=id%2CdisplayName");
});

for (const path of ["/users/a", "/users"] as const) test(`${path} returns the whole response with secret values redacted`, async () => {
  const row = { id: "a", jobTitle: "kept", ownedObjects: [{ id: "kept-too" }], passwordProfile: { password: "secret-value-123" } };
  const f = read({ path }, json(200, path === "/users" ? { value: [row] } : row));
  const result = await f.run({});
  assert.doesNotMatch(JSON.stringify(result), /secret-value-123/);
  const shown = path === "/users" ? (result.value as Record<string, unknown>[])[0]! : result;
  assert.equal(shown.jobTitle, "kept");
  assert.deepEqual(shown.ownedObjects, [{ id: "kept-too" }]);
});

test("a singleton route serves its whole body", async () => {
  const f = read({ path: "/me", scopes: "https://graph.microsoft.com/User.Read" }, json(200, { id: "a", jobTitle: "kept" }));
  assert.deepEqual(await f.run({}), { id: "a", jobTitle: "kept" });
  assert.equal(new URL(f.requests[0]!.url).pathname, "/v1.0/me");
});

test("a $count route returns the scalar count from its text body", async () => {
  const f = read({ path: "/users/$count" }, { status: 200, headers: { "content-type": "text/plain" }, body: "42" });
  assert.deepEqual(await f.run({}), { count: 42 });
  assert.equal(f.requests.length, 1);
});

test("--all follows @odata.nextLink continuations", async () => {
  const f = read({ limit: undefined }, (request, count) => count === 1
    ? json(200, { value: [{ id: "a" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=abc" })
    : json(200, { value: [{ id: "b" }] }));
  assert.deepEqual(await f.run({}), { returned: 2, complete: true, value: [{ id: "a" }, { id: "b" }] });
  assert.equal(f.requests.length, 2);
});

test("row cap ends partial with the buffered remainder reason", async () => {
  const f = read({ limit: 2 }, json(200, { value: [{ id: "a" }, { id: "b" }, { id: "c" }] }));
  const result = await f.run({});
  assert.equal(result.complete, false);
  assert.match(outputText(result.reason), /row limit reached/);
  assert.deepEqual(records(result.value).map(row => row.id), ["a", "b"]);
  assert.match(strings(result.help).join("\n"), /--all/);
  assert.equal(typeof result.cursor, "string");
});

test("cursor resumes buffered rows and the next page with its original query", async () => {
  const f = read({ path: "/groups/g/members", limit: 2 }, (request, count) => count === 1
    ? json(200, { value: [{ id: "a" }, { id: "b" }, { id: "c" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/groups/g/members?$skiptoken=next" })
    : json(200, { value: [{ id: "d" }] }));
  const first = await f.run({});
  assert.equal(first.complete, false);
  assert.equal(first["@odata.nextLink"], undefined);
  assert.deepEqual(first.value, [{ id: "a" }, { id: "b" }]);
  assert.deepEqual(JSON.parse(Buffer.from(outputText(first.cursor), "base64url").toString("utf8")).buffered, [{ id: "c" }]);
  assert.deepEqual(await f.run({ cursor: outputText(first.cursor), limit: undefined }), { returned: 2, complete: true, value: [{ id: "c" }, { id: "d" }] });
  assert.equal(f.requests.length, 2);
  assert.equal(new URL(f.requests[1]!.url).searchParams.get("$skiptoken"), "next");
});

test("cursor preserves untruncated buffered strings", async () => {
  const long = "x".repeat(5000);
  const f = read({ limit: 1 }, json(200, { value: [{ id: "a" }, { id: "b", displayName: long }] }));
  const first = await f.run({});
  const resumed = await f.run({ cursor: outputText(first.cursor), full: true });
  assert.deepEqual(resumed, { returned: 1, complete: true, value: [{ id: "b", displayName: long }] });
  assert.equal(f.requests.length, 1);
});

test("throttled partial output exposes a usable cursor", async () => {
  const f = read({}, (request, count) => count === 1
    ? { status: 429, headers: { "Retry-After": "3600" } }
    : json(200, { value: [{ id: "a" }] }));
  const first = await f.run({});
  assert.equal(first.complete, false);
  assert.match(outputText(first.reason), /throttled/);
  assert.deepEqual(await f.run({ cursor: outputText(first.cursor) }), { returned: 1, complete: true, value: [{ id: "a" }] });
});

for (const [name, extra, padding] of [
  ["whitespace", {}, " ".repeat(2_500_001)],
  ["large string properties", { unreviewed: "x".repeat(2_500_001) }, ""],
  ["UTF-8 properties", { unreviewed: "é".repeat(1_250_001) }, ""],
] as const) test(`raw ${name} bytes enforce the budget before rows are buffered`, async () => {
  const firstBody = JSON.stringify({ value: [{ id: "a", ...extra }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=next" }) + padding;
  const secondBody = JSON.stringify({ value: [{ id: "b", ...extra }] }) + padding;
  const f = read({ full: true }, (request, count) => ({ status: 200, body: count === 1 ? firstBody : secondBody }));
  const result = await f.run({});
  assert.equal(result.complete, false);
  assert.match(outputText(result.reason), /byte budget exceeded/);
  assert.deepEqual(result.value, [{ id: "a", ...extra }]);
  assert.equal(f.requests.length, 2);
  assert.deepEqual(await f.run({ cursor: outputText(result.cursor) }), { returned: 1, complete: true, value: [{ id: "b", ...extra }] });
  assert.equal(f.requests.length, 2);
});

test("raw response preserves an upstream transport's received byte count", async () => {
  const f = read({}, { ...json(200, { value: [{ id: "a" }] }), receivedBodyBytes: 5_000_001 });
  const result = await f.run({});
  assert.equal(result.complete, false);
  assert.match(outputText(result.reason), /byte budget exceeded after 5000001 bytes/);
  assert.deepEqual(result.value, []);
  assert.deepEqual(await f.run({ cursor: outputText(result.cursor) }), { returned: 1, complete: true, value: [{ id: "a" }] });
  assert.equal(f.requests.length, 1);
});

for (const overrides of [
  { path: "/users" },
  { path: "/groups/h/members" },
  { profile: { ...delegatedProfile, clientId: "44444444-4444-4444-8444-444444444444" } },
  { scopes: "https://graph.microsoft.com/GroupMember.Read.All" },
  { odata: "$select=displayName" },
]) test(`cursor refuses changed bindings ${JSON.stringify(overrides)}`, async () => {
  const f = read({ path: "/groups/g/members", limit: 1 }, json(200, { value: [{ id: "a" }, { id: "b" }] }));
  const first = await f.run({ odata: "$select=id" });
  await assert.rejects(f.run({ ...overrides, cursor: outputText(first.cursor) }), { code: "VALIDATION_ERROR" });
  assert.equal(f.requests.length, 1);
});

test("singleton reads refuse collection cursors before credentials", async () => {
  const f = read({ path: "/users/a", cursor: "invalid" });
  await assert.rejects(f.run({}), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("raw output retains long strings with or without --full", async () => {
  const long = "x".repeat(5000);
  const over = read({}, json(200, { value: [{ id: "a", displayName: long }] }));
  const output = await over.run({});
  assert.equal(records(output.value)[0]!.displayName, long);
  assert.equal(output.help, undefined);
  const f = fixture(json(200, { value: [{ id: "a", displayName: long }] }));
  const full = await runApiGet({ path: "/users", apiVersion: "v1.0", profile: delegatedProfile, scopes, full: true }, f.deps);
  assert.equal(records(full.value)[0]!.displayName, long);
  assert.equal(full.help, undefined);
  const singleton = read({ path: "/users/a" }, json(200, { id: "a", description: long }));
  assert.equal((await singleton.run({})).description, long);
  const metadata = read({}, json(200, { "@odata.context": long, value: [{ id: "a", nested: { description: long } }] }));
  const result = await metadata.run({});
  assert.equal(result["@odata.context"], long);
  assert.equal(record(records(result.value)[0]!.nested).description, long);
});

const bin = resolve("dist/bin/mg-axi.js");
const profileConfig = JSON.stringify({ version: 1, defaultProfile: "soc", profiles: { soc: delegatedProfile } });

function cli(args: readonly string[], config: string | null = null, input = "", bootstrap?: string) {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-api-"));
  try {
    if (config) writeFileSync(join(dir, "config.json"), config, { mode: 0o600 });
    return spawnSync(process.execPath, bootstrap === undefined ? [bin, ...args] : ["--input-type=module", "-e", bootstrap, ...args], {
      encoding: "utf8", input, timeout: 10000,
      env: { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot, HOME: dir, USERPROFILE: dir, XDG_CONFIG_HOME: join(dir, ".config"), MG_AXI_CONFIG: join(dir, "config.json") },
    });
  } finally { rmSync(dir, { recursive: true, force: true }); }
}

test("CLI renders safe response diagnostics on stdout for an HTML login page", () => {
  const body = "<html>private-response-value</html>";
  const bootstrap = `
    import { MsalProvider } from "./dist/msal-provider.js";
    import { Socket } from "node:net";
    Socket.prototype.connect = () => { throw new Error("Unexpected network request"); };
    MsalProvider.prototype.silent = async profile => ({
      token: "opaque-fixture-secret", expiresAt: Date.now() + 3600000,
      tenantId: profile.tenantId, clientId: profile.clientId, accountId: "synthetic-account",
    });
    globalThis.fetch = async () => new Response(${JSON.stringify(body)}, { status: 200, headers: { "content-type": "text/html; private=private-header-value" } });
    process.argv = [process.execPath, ${JSON.stringify(bin)}, ...process.argv.slice(1)];
    await (await import("./dist/cli.js")).main();
  `;
  const result = cli(["api", "get", "/users", "--scopes", scopes], profileConfig, "", bootstrap);
  assert.equal(result.status, 1);
  assert.equal(result.stderr, "");
  const output = record(decode(result.stdout));
  assert.equal(output.code, "RESPONSE_INVALID");
  assert.equal(output.status, 200);
  assert.equal(output.contentType, "text/html");
  assert.deepEqual(output.shape, [{ path: "$", type: "string" }]);
  assert.doesNotMatch(result.stdout, /private-response-value|private-header-value|opaque-fixture-secret/);
  assert.match(strings(output.help).join(" "), /API compatibility and proxy/);
});

test("api get help shows the positional path and flags", () => {
  const result = cli(["api", "get", "--help"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /mg-axi api get <path>/);
  assert.match(result.stdout, /--odata/);
  assert.match(result.stdout, /--cursor/);
  assert.match(result.stdout, /--scopes/);
  assert.match(result.stdout, /--all/);
});

for (const method of ["post", "put", "patch", "delete"]) test(`api ${method} accepts query parameters`, async () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-raw-query-"));
  const previous = process.env.MG_AXI_CONFIG;
  const config = join(dir, "config.json");
  const profile = { ...appProfile, rawWrites: true, writes: { allowWrites: true, operations: [] } };
  writeFileSync(config, JSON.stringify({ version: 1, defaultProfile: "soc", profiles: { soc: profile } }), { mode: 0o600 });
  process.env.MG_AXI_CONFIG = config;
  try {
    const f = fixture();
    const requests: { method: string; url: string }[] = [];
    const result = await executeArgv(["api", method, "/users/a", "--query", "$select=id&custom=a%20b", "--execute", "--confirm", "/users/a"], { ...f.deps, mutationTransport: async request => { requests.push(request); return { status: 204, headers: {}, body: "" }; }, journalPath: join(dir, "writes.log") });
    assert.equal(record(result).kind, "success");
    assert.equal(requests[0]!.method, method.toUpperCase());
    assert.equal(new URL(requests[0]!.url).searchParams.get("$select"), "id");
    assert.equal(new URL(requests[0]!.url).searchParams.get("custom"), "a b");
  } finally { process.env.MG_AXI_CONFIG = previous; rmSync(dir, { recursive: true, force: true }); }
});

test("missing api path fails with usage output", () => {
  const result = cli(["api", "get"]);
  assert.equal(result.status, 2);
  assert.equal(record(decode(result.stdout)).code, "VALIDATION_ERROR");
});

test("CLI rejects conflicting raw query aliases", () => {
  const result = cli(["api", "get", "/users", "--query", "$top=1", "--odata", "$top=2", "--scopes", scopes], profileConfig);
  assert.equal(result.status, 2);
  assert.match(outputText(record(decode(result.stdout)).error), /either --query or --odata/);
});

test("CLI passes --cursor to collection validation", () => {
  const result = cli(["api", "get", "/users", "--cursor", "invalid", "--scopes", scopes], profileConfig);
  assert.equal(result.status, 2);
  assert.match(outputText(record(decode(result.stdout)).error), /Invalid collection cursor/);
});

test("CLI resumes a cursor larger than the argv limit through stdin", async () => {
  const home = mkdtempSync(join(tmpdir(), "mg-axi-cursor-profile-"));
  const config = join(home, "config.json");
  writeFileSync(config, profileConfig, { mode: 0o600 });
  try {
    const rows = Array.from({ length: 200 }, (_, i) => ({ id: String(i), description: "x".repeat(1000) }));
    const f = read({ path: "/groups", limit: 1, odata: "$top=200&$select=id,description" }, json(200, { value: rows }));
    const first = await f.run({ profile: new Profiles(config).resolve("soc").profile });
    assert.ok(Buffer.byteLength(outputText(first.cursor)) > 265000);
    assert.match(strings(first.help).join("\n"), /--cursor -.*stdin/);
    const bootstrap = `
    import { MsalProvider } from "./dist/msal-provider.js";
    MsalProvider.prototype.silent = async profile => ({
      token: "opaque-fixture-secret", expiresAt: Date.now() + 3600000,
      tenantId: profile.tenantId, clientId: profile.clientId, accountId: "synthetic-account",
    });
    globalThis.fetch = async () => { throw new Error("Unexpected network request"); };
    process.argv = [process.execPath, ${JSON.stringify(bin)}, ...process.argv.slice(1)];
    await (await import("./dist/cli.js")).main();
    `;
    const result = cli(["api", "get", "/groups", "--cursor", "-", "--all", "--scopes", scopes], profileConfig, `${first.cursor}\n`, bootstrap);
    assert.equal(result.error, undefined);
    assert.equal(result.status, 0);
    assert.equal(result.stderr, "");
    assert.deepEqual(record(decode(result.stdout)), { returned: 199, complete: true, value: rows.slice(1) });
  } finally { rmSync(home, { recursive: true, force: true }); }
});

for (const token of ["", "invalid"] as const) test(`CLI stdin cursor refuses ${token || "empty input"}`, () => {
  const result = cli(["api", "get", "/users", "--cursor", "-", "--scopes", scopes], profileConfig, token);
  assert.equal(result.status, 2);
  assert.match(outputText(record(decode(result.stdout)).error), /Invalid collection cursor/);
});

test("stdin cursors retain resource binding checks", async () => {
  const f = read({ path: "/groups/g/members", limit: 1 }, json(200, { value: [{ id: "a" }, { id: "b" }] }));
  const first = await f.run({});
  const result = cli(["api", "get", "/groups/h/members", "--cursor", "-", "--scopes", scopes], profileConfig, outputText(first.cursor));
  assert.equal(result.status, 2);
  assert.match(outputText(record(decode(result.stdout)).error), /Invalid collection cursor|context does not match/);
});

test("stdin cursor input enforces the size ceiling", () => {
  const result = cli(["api", "get", "/users", "--cursor", "-", "--scopes", scopes], profileConfig, "a".repeat(MAX_CURSOR_BYTES + 1));
  assert.equal(result.status, 2);
  assert.match(outputText(record(decode(result.stdout)).error), /cursor exceeds 16000000 bytes/);
});

test("inline cursor input enforces the same size ceiling before credentials", async () => {
  const f = read({ cursor: "a".repeat(MAX_CURSOR_BYTES + 1) });
  await assert.rejects(f.run({}), error => record(error).code === "VALIDATION_ERROR" && /cursor exceeds 16000000 bytes/.test(outputText(record(error).message)));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("top help discovers the raw read leaf through its group", () => {
  const result = cli(["--help"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /mg-axi api --help/);
  const group = cli(["api", "--help"]);
  assert.equal(group.status, 0);
  assert.match(group.stdout, /mg-axi api get --help/);
});

const tenantLookupScope = "https://graph.microsoft.com/CrossTenantInformation.ReadBasic.All";
const commerceRawScope = "https://graph.microsoft.com/Organization.Read.All";
const lookupTenant = "55555555-5555-4555-8555-555555555555";
const lookupCommerce = "66666666-6666-4666-8666-666666666666";

for (const [profile, scopeArgs] of [[delegatedProfile, { scopes: tenantLookupScope }], [appProfile, {}]] as const) {
  test(`${profile.mode} raw tenant lookup by domain binds the canonical function URL`, async () => {
    const row = { tenantId: lookupTenant, defaultDomainName: "example.invalid", displayName: "Example", federationBrandName: null };
    const f = fixture(json(200, row));
    await runApiGet({ path: "/tenantRelationships/findTenantInformationByDomainName(domainName='example.invalid')", apiVersion: "v1.0", profile, ...scopeArgs }, f.deps);
    assert.equal(f.requests[0]!.url, "https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName='example.invalid')");
  });

  test(`${profile.mode} raw tenant lookup by tenant ID binds the canonical function URL`, async () => {
    const row = { tenantId: lookupTenant, defaultDomainName: "example.invalid", displayName: "Example", federationBrandName: null };
    const f = fixture(json(200, row));
    await runApiGet({ path: `/tenantRelationships/findTenantInformationByTenantId(tenantId='${lookupTenant}')`, apiVersion: "v1.0", profile, ...scopeArgs }, f.deps);
    assert.equal(f.requests[0]!.url, `https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByTenantId(tenantId='${lookupTenant}')`);
  });
}

for (const [profile, scopeArgs] of [[delegatedProfile, { scopes: commerceRawScope }], [appProfile, {}]] as const) {
  test(`${profile.mode} raw commerce-key lookup binds the canonical function URL`, async () => {
    const row = { id: "860697e3-b0aa-4196-a6c6-7ec361ed58f7", commerceSubscriptionId: lookupCommerce, status: "Enabled" };
    const f = fixture(json(200, row));
    await runApiGet({ path: `/directory/subscriptions(commerceSubscriptionId='${lookupCommerce}')`, apiVersion: "v1.0", profile, odata: "$select=id,commerceSubscriptionId,status", ...scopeArgs }, f.deps);
    assert.equal(f.requests[0]!.url, `https://graph.microsoft.com/v1.0/directory/subscriptions(commerceSubscriptionId='${lookupCommerce}')?%24select=id%2CcommerceSubscriptionId%2Cstatus`);
  });
}

for (const scopes of [undefined, "https://example.invalid/User.Read"] as const) test(`delegated raw read with scopes ${scopes ?? "(none)"} is refused before credentials`, async () => {
  const f = read({ path: "/users", scopes });
  await assert.rejects(f.run({}), error => record(error).code === "VALIDATION_ERROR" && /full https:\/\/graph.microsoft.com\/ names/.test(outputText(record(error).message)));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});
