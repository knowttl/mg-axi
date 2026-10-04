import assert from "node:assert/strict";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import { capabilityDocument, skillCommandTable } from "../dist/docs.js";
import { setupView } from "../dist/setup.js";

// WRITE-03 acceptance through the real CLI interface with fixture-only
// transports: exact POST target and empty request, user verification before
// preview and sending, action preview text with delay/external-user limits,
// confirmation required, refusal without opt-in or under forced read-only,
// unknown outcome on timeout or 5xx, journal intent/outcome. No live
// instance, real credential or customer data exists.

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const userId = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const WRITE_SCOPES = ["https://graph.microsoft.com/User.RevokeSessions.All"];
const READ_SCOPES = ["https://graph.microsoft.com/User.ReadBasic.All"];
const MUTATION_URL = `https://graph.microsoft.com/v1.0/users/${userId}/revokeSignInSessions`;

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-write-03-"));
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(dir, "config.json");
  const profiles = new Profiles();
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  return { dir, previous };
}

// Hand-enabling mirrors production: no command writes the writes object.
function enableWrites(dir, name = "soc") {
  const path = join(dir, "config.json");
  const config = JSON.parse(readFileSync(path, "utf8"));
  config.profiles[name].writes = { allowWrites: true, operations: ["entra.user.revokeSessions"] };
  writeFileSync(path, JSON.stringify(config, null, 2));
}

function teardown(state) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function json(status, body) {
  return { status, headers: {}, body: JSON.stringify(body) };
}

// reads: scripted user-verification answers ({ id } bodies, or { status }
// for read errors, or raw bodies for malformed/non-user cases); every read
// must hit the READ-01 user route with the id select. mutation: scripted
// POST answers or thrown errors.
function fixture({ reads = [{ id: userId }], mutation = [{ status: 200, body: { value: true } }] } = {}) {
  const readRequests = [];
  const mutRequests = [];
  const credCalls = [];
  let readIndex = 0;
  let mutIndex = 0;
  const transport = async request => {
    readRequests.push(request);
    const url = new URL(request.url);
    const match = /^\/v1\.0\/users\/([^/]+)$/.exec(url.pathname);
    assert.ok(match, `unexpected read route ${url.pathname}`);
    assert.equal(url.searchParams.get("$select"), "id");
    const next = readIndex < reads.length ? reads[readIndex++] : reads[reads.length - 1];
    if (next !== null && typeof next === "object" && "status" in next && !("id" in next) && !("@odata.type" in next)) {
      return json(next.status, { error: { code: "fixture-read-error", message: "synthetic read failure" } });
    }
    return json(200, next);
  };
  const mutationTransport = async request => {
    mutRequests.push(request);
    const next = mutIndex < mutation.length ? mutation[mutIndex++] : mutation[mutation.length - 1];
    if (next instanceof Error) throw next;
    const body = typeof next.body === "string" ? next.body : JSON.stringify(next.body ?? {});
    return { status: next.status, headers: {}, body };
  };
  const credential = mode => ({
    credential: async (...args) => {
      credCalls.push([mode, ...args]);
      return {
        token: `opaque-fixture-${mode}-token`,
        expiresAt: Date.now() + 3_600_000,
        tenantId: tenant,
        clientId: client,
        ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
      };
    },
  });
  const delegated = credential("delegated");
  const application = credential("application");
  const journalPath = join(mkdtempSync(join(tmpdir(), "mg-axi-write-03-journal-")), "writes.log");
  const overrides = { transport, delegated, application, mutationTransport, journalPath };
  return { readRequests, mutRequests, credCalls, journalPath, overrides };
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

function revokeArgs(user, extra = []) {
  return ["entra", "user", "revoke-sessions", "--user", user, "--profile", "soc", ...extra];
}

test("preview states the action, its limits and irreversibility, and sends nothing", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    const result = await executeArgv(revokeArgs(userId), f.overrides);
    assert.deepEqual(result.preview, {
      operation: "entra.user.revokeSessions",
      target: userId,
      method: "POST",
      url: MUTATION_URL,
      effect: "disruptive",
      action: result.preview.action,
      reversible: false,
      limitations: result.preview.limitations,
      noop: false,
    });
    assert.match(result.preview.action, /signInSessionsValidFromDateTime/);
    assert.match(result.preview.action, /sign in again/);
    const limits = result.preview.limitations.join("\n");
    assert.match(limits, /few minutes/);
    assert.match(limits, /home tenant/);
    assert.match(limits, /cannot be undone|no rollback/);
    assert.ok(result.help.some(hint => hint.includes("--execute") && hint.includes("--confirm") && hint.includes(userId)));
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 1);
    assert.ok(!existsSync(f.journalPath), "dry-run journals nothing");
  } finally { teardown(state); }
});

test("execute sends the exact POST with no body and records success", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    const result = await executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides);
    assert.deepEqual(result.user, { id: userId, sessionsRevoked: true });
    assert.equal(typeof result.auditId, "string");
    assert.equal(f.mutRequests.length, 1);
    const [sent] = f.mutRequests;
    assert.equal(sent.method, "POST");
    assert.equal(sent.url, MUTATION_URL);
    assert.equal(sent.body, undefined);
    assert.ok(!("Content-Type" in sent.headers), "an action with no body sends no content type");
    assert.match(sent.headers.Authorization, /^Bearer opaque-fixture-delegated-token$/);
    assert.equal(typeof sent.headers["client-request-id"], "string");
    assert.ok(f.credCalls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(WRITE_SCOPES)));
    assert.ok(f.credCalls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(READ_SCOPES)));
    assert.deepEqual(f.readRequests.map(request => new URL(request.url).pathname), [
      `/v1.0/users/${userId}`,
      `/v1.0/users/${userId}`,
      `/v1.0/users/${userId}`,
    ]);
    const records = journal(f.journalPath);
    assert.equal(records.length, 2);
    assert.equal(records[0].kind, "intent");
    assert.equal(records[0].operation, "entra.user.revokeSessions");
    assert.equal(records[0].method, "POST");
    assert.equal(records[0].url, MUTATION_URL);
    assert.equal(records[0].target, userId);
    assert.equal(records[0].effect, "disruptive");
    assert.equal(records[1].kind, "outcome");
    assert.equal(records[1].outcome, "SUCCESS");
    assert.equal(records[1].httpStatus, 200);
    assert.equal(records[1].intentKey, records[0].intentKey);
    const file = readFileSync(f.journalPath, "utf8");
    assert.ok(!file.includes("opaque-fixture"), "journal carries no credentials");
  } finally { teardown(state); }
});

test("execute without --confirm is refused before credentials", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(executeArgv(revokeArgs(userId, ["--execute"]), f.overrides), { code: "CONFIRM_REQUIRED" });
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("mismatched --confirm is refused before credentials", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(
      executeArgv(revokeArgs(userId, ["--execute", "--confirm", "someone-else"]), f.overrides),
      { code: "CONFIRM_MISMATCH" },
    );
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("profile without opt-in is refused with zero sends and zero reads", async () => {
  const state = setupProfiles();
  try {
    const f = fixture();
    await assert.rejects(executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides), { code: "WRITES_DISABLED" });
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("forced read-only overrides the hand-enabled profile", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await withEnv("MG_AXI_READ_ONLY", "1", async () => {
      await assert.rejects(executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides), { code: "WRITES_DISABLED" });
    });
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("unknown target sends nothing and journals no intent", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [{ status: 404 }] });
    await assert.rejects(executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides));
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath), "a failed preview read reserves no intent");
  } finally { teardown(state); }
});

test("a read without a Graph object ID blocks the revocation", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [{}] });
    await assert.rejects(executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides), { code: "GRAPH_ERROR" });
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath));
  } finally { teardown(state); }
});

test("a non-user object blocks the revocation", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [{ id: userId, "@odata.type": "#microsoft.graph.group" }] });
    await assert.rejects(
      executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides),
      error => error.code === "GRAPH_ERROR" && /user identity/.test(error.message),
    );
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath));
  } finally { teardown(state); }
});

test("failed fresh read journals NOT_SENT and sends nothing", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [{ id: userId }, { id: userId }, { status: 404 }] });
    await assert.rejects(executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides));
    assert.equal(f.mutRequests.length, 0);
    const records = journal(f.journalPath);
    assert.equal(records.length, 2);
    assert.equal(records[0].kind, "intent");
    assert.equal(records[1].outcome, "NOT_SENT");
  } finally { teardown(state); }
});

test("transport failure records an unknown outcome with no replay", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ mutation: [new Error("socket timeout")] });
    const error = await executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("timeout must throw"), caught => caught);
    assert.equal(error.code, "OUTCOME_UNKNOWN");
    assert.match(error.suggestions.join("\n"), /Never replay this intent/);
    assert.equal(f.mutRequests.length, 1);
    assert.equal(journal(f.journalPath)[1].outcome, "OUTCOME_UNKNOWN");
  } finally { teardown(state); }
});

for (const status of [408, 500, 503]) test(`uncertain status ${status} after send records an unknown outcome`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ mutation: [{ status }] });
    const error = await executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("uncertain outcome must throw"), caught => caught);
    assert.equal(error.code, "OUTCOME_UNKNOWN");
    assert.match(error.suggestions.join("\n"), /Never replay this intent/);
    assert.equal(f.mutRequests.length, 1);
    const records = journal(f.journalPath);
    assert.equal(records[1].outcome, "OUTCOME_UNKNOWN");
    assert.equal(records[1].httpStatus, status);
  } finally { teardown(state); }
});

test("denial surfaces the revoke permission without inventing a role verdict", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ mutation: [{ status: 403, body: { error: { code: "Authorization_RequestDenied", message: "denied" } } }] });
    const error = await executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("denial must throw"), caught => caught);
    assert.equal(error.code, "GRAPH_ERROR");
    assert.match(error.suggestions.join("\n"), /User\.RevokeSessions\.All/);
    assert.match(error.suggestions.join("\n"), /never replay/i);
    assert.equal(journal(f.journalPath)[1].outcome, "FAILED");
  } finally { teardown(state); }
});

for (const status of [400, 404]) test(`POST refusal ${status} preserves status without permission diagnosis`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ mutation: [{ status }] });
    const error = await executeArgv(revokeArgs(userId, ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("refusal must throw"), caught => caught);
    assert.equal(error.code, "GRAPH_ERROR");
    assert.equal(error.message, `Graph refused the session revocation (status ${status})`);
    assert.equal(error.suggestions.length, 1);
    assert.ok(error.suggestions[0].includes(`read back '${userId}' and never replay`));
    assert.equal(f.mutRequests.length, 1);
    assert.equal(journal(f.journalPath)[1].httpStatus, status);
  } finally { teardown(state); }
});

test("beta api-version is refused before credentials", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(
      executeArgv(["entra", "user", "revoke-sessions", "--user", userId, "--profile", "soc", "--api-version", "beta"], f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.credCalls.length, 0);
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 0);
  } finally { teardown(state); }
});

for (const user of ["AdeleV@contoso.com", "not-a-guid", "..", "a/b", "a b"]) test(`non-object-ID target ${user} is refused before credentials`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(executeArgv(revokeArgs(user, ["--execute", "--confirm", user]), f.overrides), { code: "VALIDATION_ERROR" });
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("caller scopes are rejected: only the documented scope is requested", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(
      executeArgv([...revokeArgs(userId), "--scopes", "https://graph.microsoft.com/User.Read.All"], f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("application mode uses the .default audience and succeeds", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir, "batch");
    const f = fixture();
    const result = await executeArgv(
      ["entra", "user", "revoke-sessions", "--user", userId, "--profile", "batch", "--execute", "--confirm", userId],
      f.overrides,
    );
    assert.deepEqual(result.user, { id: userId, sessionsRevoked: true });
    assert.equal(f.mutRequests.length, 1);
    assert.ok(f.credCalls.length > 0);
    for (const [mode, , scopes] of f.credCalls) {
      if (mode === "application") assert.equal(scopes, undefined, "application credentials use the .default audience, never delegated scopes");
    }
  } finally { teardown(state); }
});

test("generated records name the session revocation write", async () => {
  assert.ok(skillCommandTable().includes("`mg-axi entra user revoke-sessions` | native | write |"));
  const coverage = capabilityDocument();
  assert.ok(coverage.includes("## Named writes"));
  assert.ok(coverage.includes("`mg-axi entra user revoke-sessions` | `POST:/users/{user-id}/revokeSignInSessions` | WRITE-03 |"));
});

test("setup counts the write beside reads and local leaves", async () => {
  const state = setupProfiles();
  try {
    const output = setupView(new Profiles());
    assert.equal(output.capabilities.reads, 40);
    assert.equal(output.capabilities.writes, 3);
    assert.equal(output.capabilities.local, 6);
    assert.ok(output.capabilities.implemented.includes("entra user revoke-sessions"));
  } finally { teardown(state); }
});
