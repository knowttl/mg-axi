import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, test } from "node:test";
import { ApplicationAuth } from "../dist/app-auth.js";
import { DelegatedAuth } from "../dist/auth.js";
import { LEAVES } from "../dist/catalogue.js";
import { runApiGet } from "../dist/api.js";
import { createMutationCoordinator, createMutationSender } from "../dist/mutations.js";

// WRITE-00 acceptance: the coordinator is entirely fixture-driven with no
// real mutation family enabled. Every test drives it through this synthetic
// fixture mutation; no live instance, real credential or customer data exists.
const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const scopes = ["https://graph.microsoft.com/Fixture.Write"];
const NOW = 1_700_000_000_000;
const TIME = new Date(NOW).toISOString();

const writes = { allowWrites: true, operations: ["mg.fixture.write"] };
const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key } };
const enabledDelegated = { ...delegatedProfile, writes };
const appProfile = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "federated", key } };
const enabledApp = { ...appProfile, writes };
const mutation = { operation: "mg.fixture.write", method: "POST", version: "v1.0", path: "/v1.0/fixture/notes", effect: "write", target: "fixture-note-1", payload: { text: "reviewed synthetic note" } };

const scratches = [];
afterEach(() => { while (scratches.length) rmSync(scratches.pop(), { recursive: true, force: true }); });

function fixture({ handler, credentialTenant = tenant, onCredential } = {}) {
  const scratch = mkdtempSync(join(tmpdir(), "mg-write00-"));
  scratches.push(scratch);
  const journalPath = join(scratch, "nested", "writes.log");
  const credentialCalls = [];
  const requests = [];
  const credential = account => {
    onCredential?.();
    return { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: credentialTenant, clientId: client, ...account };
  };
  const delegated = new DelegatedAuth({
    storage: "session-only",
    login: async () => credential({ accountId: "synthetic-account" }),
    silent: async (...args) => { credentialCalls.push(["silent", ...args]); return credential({ accountId: "synthetic-account" }); },
  });
  const application = new ApplicationAuth({
    storage: "session-only",
    acquire: async (...args) => { credentialCalls.push(["acquire", ...args]); return credential({}); },
  });
  const transport = async request => {
    requests.push(request);
    return typeof handler === "function" ? handler(request, requests.length) : handler ?? { status: 200, headers: {}, body: "{}" };
  };
  const coordinator = (profile = enabledDelegated, overrides = {}) => createMutationCoordinator({
    profile, delegated, application, transport, clock: () => NOW, journalPath, ...overrides,
  });
  return { scratch, journalPath, credentialCalls, requests, delegated, application, transport, coordinator };
}

function journal(path) {
  return readFileSync(path, "utf8").split("\n").filter(Boolean).map(line => JSON.parse(line));
}

async function withEnv(key, value, action) {
  const previous = process.env[key];
  if (value === undefined) delete process.env[key];
  else process.env[key] = value;
  try {
    return await action();
  } finally {
    if (previous === undefined) delete process.env[key];
    else process.env[key] = previous;
  }
}

const readState = (current = { notes: ["previous synthetic note"] }) => async () => current;

for (const path of ["/v1.0/fixture/$count", "/v1.0/fixture/%24value", "/v1.0/fixture/notes/$ref"]) {
  test(`mutation path rejects reserved resource binding ${path} before credentials`, async () => {
    const f = fixture();
    await assert.rejects(f.coordinator().execute({ ...mutation, path }, { readState: readState(), scopes }), { code: "VALIDATION_ERROR" });
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
  });
}

test("group-member mutation rejects a reserved group binding before credentials", async () => {
  const f = fixture();
  const operation = "mg.entra.group.member.add";
  const coordinator = f.coordinator({ ...enabledDelegated, writes: { allowWrites: true, operations: [operation] } });
  await assert.rejects(
    coordinator.execute({ ...mutation, operation, path: "/v1.0/groups/$count/members/$ref" }, { readState: readState(), scopes }),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("read-only by default: no writes key means WRITES_DISABLED with zero sends", async () => {
  const f = fixture();
  const coordinator = f.coordinator(delegatedProfile);
  await assert.rejects(coordinator.execute(mutation, { readState: readState() }), { code: "WRITES_DISABLED" });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("exposed scope is frozen: tampering throws and gates still hold", async () => {
  const f = fixture();
  const coordinator = f.coordinator(delegatedProfile);
  assert.ok(Object.isFrozen(coordinator.scope));
  assert.equal(coordinator.scope.allowWrites, false);
  assert.throws(() => coordinator.scope.operations.push("mg.other.write"), TypeError);
  await assert.rejects(coordinator.execute(mutation, { readState: readState() }), { code: "WRITES_DISABLED" });
  assert.equal(f.requests.length, 0);
});

test("explicit allowWrites false stays read-only with zero sends", async () => {
  const f = fixture();
  const coordinator = f.coordinator({ ...delegatedProfile, writes: { allowWrites: false, operations: ["mg.fixture.write"] } });
  await assert.rejects(coordinator.execute(mutation, { readState: readState() }), { code: "WRITES_DISABLED" });
  assert.equal(f.requests.length, 0);
});

test("forced read-only overrides a hand-enabled profile with zero sends", async () => {
  const f = fixture();
  await withEnv("MG_AXI_READ_ONLY", "1", async () => {
    const coordinator = f.coordinator();
    await assert.rejects(coordinator.execute(mutation, { readState: readState() }), { code: "WRITES_DISABLED" });
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
  });
});

for (const [mode, profile, credentialOptions] of [
  ["delegated", enabledDelegated, { scopes }],
  ["application", enabledApp, {}],
]) {
  test(`forced read-only during the fresh read blocks ${mode} credentials and send`, async () => {
    const f = fixture();
    let reads = 0;
    await withEnv("MG_AXI_READ_ONLY", undefined, async () => {
      await assert.rejects(f.coordinator(profile).execute(mutation, {
        execute: true, ...credentialOptions, intentId: "read-only-fresh-read",
        readState: async () => {
          reads += 1;
          if (reads === 2) process.env.MG_AXI_READ_ONLY = "1";
          return { notes: [] };
        },
      }), { code: "WRITES_DISABLED" });
    });
    assert.equal(reads, 2);
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
    assert.equal(journal(f.journalPath)[1].outcome, "NOT_SENT");
  });

  test(`forced read-only during ${mode} credential acquisition blocks the send`, async () => {
    const f = fixture({ onCredential: () => { process.env.MG_AXI_READ_ONLY = "1"; } });
    await withEnv("MG_AXI_READ_ONLY", undefined, async () => {
      await assert.rejects(f.coordinator(profile).execute(mutation, {
        execute: true, ...credentialOptions, readState: readState(), intentId: "read-only-credential",
      }), { code: "WRITES_DISABLED" });
    });
    assert.equal(f.credentialCalls.length, 1);
    assert.equal(f.requests.length, 0);
    assert.equal(journal(f.journalPath)[1].outcome, "NOT_SENT");
  });
}

test("operation outside the configured scope is refused with zero sends", async () => {
  const f = fixture();
  const coordinator = f.coordinator();
  await assert.rejects(
    coordinator.execute({ ...mutation, operation: "mg.other.write" }, { readState: readState(), scopes }),
    { code: "OPERATION_NOT_WRITABLE" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("read-time scopes never widen the write scope", async () => {
  const f = fixture();
  const coordinator = f.coordinator();
  await assert.rejects(
    coordinator.execute({ ...mutation, operation: "mg.other.write" }, { readState: readState(), scopes: [...scopes, "https://graph.microsoft.com/Group.Read.All"] }),
    { code: "OPERATION_NOT_WRITABLE" },
  );
  assert.equal(f.requests.length, 0);
});

for (const path of ["", "v1.0/fixture", "/beta/fixture/notes", "/v1.0/../tenants", "/v1.0/a b", "/v1.0/a?b=c", "/v1.0/a#b", "/v1.0/a\\b", "/v1.0/", "/v1.0//notes", "/v1.0/a/{b}", "/v1.0/%2e%2e/x", "/v1.0/a%2fb", "/v1.0/a%5cb", "/v1.0/a%252fb", "/v1.0/%20", "/v1.0/%00", "/v1.0/%7bb%7d", "/v1.0/%zz"]) {
  test(`hostile mutation path ${JSON.stringify(path)} is refused with zero sends`, async () => {
    const f = fixture();
    await assert.rejects(
      f.coordinator().execute({ ...mutation, path }, { readState: readState(), scopes }),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.requests.length, 0);
  });
}

test("beta version and empty operation or target fail closed with zero sends", async () => {
  const f = fixture();
  const coordinator = f.coordinator();
  for (const definition of [
    { ...mutation, version: "beta", path: "/beta/fixture/notes" },
    { ...mutation, operation: "  " },
    { ...mutation, target: "" },
  ]) {
    await assert.rejects(coordinator.execute(definition, { readState: readState(), scopes }), { code: "VALIDATION_ERROR" });
  }
  assert.equal(f.requests.length, 0);
});

for (const [name, profile] of [["enabled", enabledDelegated], ["disabled", delegatedProfile], ["application", enabledApp]]) {
  test(`forged sender authorization on a ${name} profile is refused before credentials or HTTP`, async () => {
    const f = fixture();
    const sender = createMutationSender({ profile, delegated: f.delegated, application: f.application, transport: f.transport });
    await assert.rejects(
      sender.send(Object.freeze({ nonce: "forged", method: "POST", url: "https://graph.microsoft.com/v1.0/fixture/notes" }), { scopes }),
      { code: "OPERATION_BLOCKED" },
    );
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
  });
}

test("forged sender authorization is refused under forced read-only", async () => {
  const f = fixture();
  const sender = createMutationSender({ profile: enabledDelegated, delegated: f.delegated, application: f.application, transport: f.transport });
  await withEnv("MG_AXI_READ_ONLY", "1", async () => {
    await assert.rejects(
      sender.send(Object.freeze({ nonce: "forged", method: "POST", url: "https://graph.microsoft.com/v1.0/fixture/notes" }), { scopes }),
      { code: "OPERATION_BLOCKED" },
    );
  });
  assert.equal(f.requests.length, 0);
});

test("preview shows the actual redacted change without sending or journaling", async () => {
  const f = fixture();
  const seen = f.coordinator().preview(
    { ...mutation, payload: { text: "reviewed", password: "fixture-secret" } },
    { notes: ["previous synthetic note"], password: "fixture-secret" },
  );
  assert.equal(seen.url, "https://graph.microsoft.com/v1.0/fixture/notes");
  assert.equal(seen.noop, false);
  assert.match(seen.currentState, /previous synthetic note/);
  assert.match(seen.currentState, /\*\*\*redacted\*\*\*/);
  assert.doesNotMatch(seen.currentState, /fixture-secret/);
  assert.match(seen.proposedChange, /\*\*\*redacted\*\*\*/);
  assert.doesNotMatch(seen.proposedChange, /fixture-secret/);
  assert.equal(f.requests.length, 0);
  assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
});

test("omitting execute previews without sending or journaling", async () => {
  const f = fixture();
  const result = await f.coordinator().execute(mutation, { readState: readState(), scopes });
  assert.equal(result.kind, "dry-run");
  assert.equal(result.preview.target, "fixture-note-1");
  assert.equal(f.requests.length, 0);
  assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
});

for (const confirm of [undefined, "wrong-target"]) {
  test(`disruptive mutation with confirm ${JSON.stringify(confirm)} is refused with zero sends`, async () => {
    const f = fixture();
    const definition = { ...mutation, effect: "disruptive" };
    await assert.rejects(
      f.coordinator().execute(definition, { execute: true, ...(confirm === undefined ? {} : { confirm }), readState: readState(), scopes }),
      { code: confirm === undefined ? "CONFIRM_REQUIRED" : "CONFIRM_MISMATCH" },
    );
    assert.equal(f.requests.length, 0);
  });
}

test("confirmed execution sends the previewed snapshot once with If-Match and journals intent plus outcome", async () => {
  const f = fixture({ handler: { status: 200, headers: {}, body: '{"id":"1"}' } });
  const definition = { ...mutation, effect: "disruptive" };
  const result = await f.coordinator().execute(definition, {
    execute: true, confirm: "fixture-note-1", ifMatch: '"fixture-etag"',
    readState: readState(), scopes, intentId: "approved-intent",
  });
  assert.equal(result.kind, "success");
  assert.equal(result.auditId, "approved-intent");
  assert.equal(result.status, 200);
  assert.deepEqual(result.response, { id: "1" });
  assert.equal(f.requests.length, 1);
  assert.equal(f.requests[0].method, "POST");
  assert.equal(f.requests[0].url, "https://graph.microsoft.com/v1.0/fixture/notes");
  assert.equal(f.requests[0].body, JSON.stringify(mutation.payload));
  assert.equal(f.requests[0].headers["If-Match"], '"fixture-etag"');
  assert.equal(f.requests[0].headers["Content-Type"], "application/json");
  assert.equal(f.requests[0].headers.Authorization, "Bearer opaque-fixture-secret");
  assert.match(f.requests[0].headers["client-request-id"], /^[0-9a-f-]{36}$/);
  const records = journal(f.journalPath);
  assert.equal(records.length, 2);
  assert.equal(records[0].kind, "intent");
  assert.equal(records[1].kind, "outcome");
  assert.equal(records[1].outcome, "SUCCESS");
  assert.equal(records[0].intentKey, records[1].intentKey);
  assert.equal(records[0].time, TIME);
  assert.equal(records[0].ifMatch, '"fixture-etag"');
  assert.ok(!("payload" in records[0]) && !("body" in records[0]));
});

test("non-disruptive write executes without target confirmation", async () => {
  const f = fixture({ handler: { status: 200, headers: {}, body: "{}" } });
  const result = await f.coordinator().execute(mutation, { execute: true, readState: readState(), scopes, intentId: "plain-write" });
  assert.equal(result.kind, "success");
  assert.equal(f.requests.length, 1);
  assert.equal(journal(f.journalPath)[1].outcome, "SUCCESS");
});

test("application profile executes through the .default audience without caller scopes", async () => {
  const f = fixture({ handler: { status: 200, headers: {}, body: "{}" } });
  const result = await f.coordinator(enabledApp).execute(mutation, { execute: true, readState: readState(), intentId: "app-write" });
  assert.equal(result.kind, "success");
  assert.deepEqual(f.credentialCalls[0][0], "acquire");
  assert.equal(journal(f.journalPath)[1].outcome, "SUCCESS");
});

test("application profile rejects caller scopes and delegated requires them", async () => {
  const f = fixture({ handler: { status: 200, headers: {}, body: "{}" } });
  await assert.rejects(
    f.coordinator(enabledApp).execute(mutation, { execute: true, readState: readState(), scopes, intentId: "app-scoped" }),
    { code: "VALIDATION_ERROR" },
  );
  await assert.rejects(
    f.coordinator().execute(mutation, { execute: true, readState: readState(), intentId: "delegated-unscoped" }),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.requests.length, 0);
});

test("credential for another tenant fails closed with zero sends", async () => {
  const f = fixture({ credentialTenant: "99999999-9999-4999-8999-999999999999" });
  await assert.rejects(
    f.coordinator().execute(mutation, { execute: true, readState: readState(), scopes, intentId: "wrong-tenant" }),
    { code: "AUTH_REQUIRED" },
  );
  assert.equal(f.requests.length, 0);
});

test("already-desired state is a no-op with zero sends and no journal", async () => {
  const f = fixture();
  const isNoop = current => current?.enabled === false;
  const result = await f.coordinator().execute(
    { ...mutation, payload: { enabled: false } },
    { execute: true, readState: readState({ enabled: false }), isNoop, scopes },
  );
  assert.equal(result.kind, "noop");
  assert.equal(result.preview.noop, true);
  assert.equal(f.requests.length, 0);
  assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
});

test("failed fresh re-read aborts without sending and records NOT_SENT", async () => {
  const f = fixture();
  let calls = 0;
  const readStateFlaky = async () => {
    calls += 1;
    if (calls > 1) throw Object.assign(new Error("Synthetic read failure"), { code: "FIXTURE_READ_FAILED" });
    return { notes: [] };
  };
  await assert.rejects(
    f.coordinator().execute(mutation, { execute: true, readState: readStateFlaky, scopes, intentId: "reread-fails" }),
    /Synthetic read failure/,
  );
  assert.equal(f.requests.length, 0);
  const records = journal(f.journalPath);
  assert.equal(records.length, 2);
  assert.equal(records[1].outcome, "NOT_SENT");
});

test("fresh read finding desired state cancels the send and records NOT_SENT", async () => {
  const f = fixture();
  let calls = 0;
  const readStateChanging = async () => (++calls === 1 ? { notes: [] } : { notes: ["reviewed synthetic note"] });
  const result = await f.coordinator().execute(mutation, {
    execute: true, readState: readStateChanging,
    isNoop: current => current?.notes?.includes("reviewed synthetic note") === true,
    scopes, intentId: "became-desired",
  });
  assert.equal(result.kind, "noop");
  assert.equal(f.requests.length, 0);
  assert.equal(journal(f.journalPath)[1].outcome, "NOT_SENT");
});

test("unavailable intent storage blocks the send", async () => {
  const f = fixture();
  const blocker = join(f.scratch, "blocker");
  writeFileSync(blocker, "not a directory");
  const coordinator = f.coordinator(enabledDelegated, { journalPath: join(blocker, "nested", "writes.log") });
  await assert.rejects(
    coordinator.execute(mutation, { execute: true, readState: readState(), scopes, intentId: "blocked-intent" }),
    { code: "INTENT_NOT_RECORDED" },
  );
  assert.equal(f.requests.length, 0);
});

test("outcome storage failure reports reconciliation with the send already made", async () => {
  // Between intent journaling and outcome journaling the transport swaps the
  // journal file for a directory, so only the outcome record fails.
  const f = fixture({ handler: () => {
    rmSync(f.journalPath);
    mkdirSync(f.journalPath);
    return { status: 200, headers: {}, body: "{}" };
  } });
  await assert.rejects(
    f.coordinator().execute(mutation, { execute: true, readState: readState(), scopes, intentId: "outcome-blocked" }),
    { code: "OUTCOME_NOT_RECORDED" },
  );
  assert.equal(f.requests.length, 1);
});

test("definitive server rejection reports failed and persists FAILED", async () => {
  const f = fixture({ handler: { status: 403, headers: {}, body: '{"error":{"code":"AuthorizationFailed"}}' } });
  const result = await f.coordinator().execute(mutation, { execute: true, readState: readState(), scopes, intentId: "denied-intent" });
  assert.equal(result.kind, "failed");
  assert.equal(result.status, 403);
  assert.equal(f.requests.length, 1);
  assert.equal(journal(f.journalPath)[1].outcome, "FAILED");
});

test("ambiguous transport failure reports unknown and the intent can never replay", async () => {
  const f = fixture({ handler: () => { throw new Error("socket hang up"); } });
  const coordinator = f.coordinator();
  const options = { execute: true, readState: readState(), scopes, intentId: "timeout-intent" };
  const result = await coordinator.execute(mutation, options);
  assert.equal(result.kind, "unknown");
  assert.equal(result.auditId, "timeout-intent");
  assert.match(result.guidance, /never replay this intent/);
  assert.equal(f.requests.length, 1);
  assert.equal(journal(f.journalPath)[1].outcome, "OUTCOME_UNKNOWN");
  await assert.rejects(f.coordinator().execute(mutation, options), { code: "ALREADY_EXECUTED" });
  assert.equal(f.requests.length, 1);
});

for (const status of [408, 500, 502, 503, 504, 599]) {
  test(`mutation applied before HTTP ${status} reports unknown and cannot replay`, async () => {
    const notes = [];
    const f = fixture({ handler: request => {
      notes.push(JSON.parse(request.body).text);
      return { status, headers: {}, body: "{}" };
    } });
    const options = { execute: true, readState: async () => ({ notes: [...notes] }), scopes, intentId: `http-${status}` };
    const result = await f.coordinator().execute(mutation, options);
    assert.deepEqual(notes, [mutation.payload.text]);
    assert.equal(result.kind, "unknown");
    assert.equal(result.httpStatus, status);
    assert.equal(result.auditId, options.intentId);
    assert.match(result.guidance, /read back target 'fixture-note-1'/);
    assert.match(result.guidance, /never replay this intent/);
    const records = journal(f.journalPath);
    assert.equal(records.length, 2);
    assert.equal(records[1].outcome, "OUTCOME_UNKNOWN");
    assert.equal(records[1].httpStatus, status);
    await assert.rejects(f.coordinator().execute(mutation, options), { code: "ALREADY_EXECUTED" });
    assert.equal(f.requests.length, 1);
    assert.deepEqual(notes, [mutation.payload.text]);
  });
}

test("recreated coordinator refuses an intent recorded without a terminal outcome", async () => {
  const f = fixture({ handler: { status: 200, headers: {}, body: "{}" } });
  const intentId = "crash-intent";
  const intentKey = createHash("sha256").update(intentId).digest("hex");
  mkdirSync(join(f.scratch, "nested"), { recursive: true });
  writeFileSync(f.journalPath, `${JSON.stringify({ kind: "intent", id: intentId, intentKey, time: TIME, profile: tenant, operation: mutation.operation, method: "POST", url: "https://graph.microsoft.com/v1.0/fixture/notes", target: mutation.target, effect: "write" })}\n`);
  await assert.rejects(
    f.coordinator().execute(mutation, { execute: true, readState: readState(), scopes, intentId }),
    { code: "ALREADY_EXECUTED" },
  );
  assert.equal(f.requests.length, 0);
});

test("every catalogue leaf stays a read except the reviewed named writes", () => {
  const operations = LEAVES.filter(leaf => leaf.operation).map(leaf => leaf.operation);
  assert.ok(operations.length > 0);
  assert.ok(operations.includes("POST:/groups/{group-id}/members/$ref"), "WRITE-01 leaf is catalogued");
  assert.ok(operations.includes("PATCH:/users/{user-id}"), "WRITE-02 leaf is catalogued");
  assert.ok(operations.includes("POST:/users/{user-id}/revokeSignInSessions"), "WRITE-03 leaf is catalogued");
  assert.ok(operations.includes("POST:/identityProtection/riskyUsers/dismiss"), "WRITE-05 leaf is catalogued");
  const reviewed = new Set(["POST:/groups/{group-id}/members/$ref", "PATCH:/users/{user-id}", "POST:/users/{user-id}/revokeSignInSessions", "POST:/identityProtection/riskyUsers/dismiss"]);
  for (const operation of operations) {
    if (reviewed.has(operation)) continue;
    assert.match(operation, /^GET:/, `write leaf in catalogue: ${operation}`);
  }
});

test("reviewed raw surface never serves a mutation route", async () => {
  const f = fixture({ handler: { status: 200, headers: {}, body: "{}" } });
  const delegated = new DelegatedAuth({ storage: "session-only", login: async () => { throw new Error("no login"); }, silent: async () => ({ token: "t", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "a" }) });
  const application = new ApplicationAuth({ storage: "session-only", acquire: async () => ({ token: "t", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client }) });
  await assert.rejects(
    runApiGet({ path: "/groups/fixture-id/members/$ref", apiVersion: "v1.0", profile: delegatedProfile, scopes: "https://graph.microsoft.com/Group.Read.All" }, { delegated, application, transport: f.transport }),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.requests.length, 0);
});
