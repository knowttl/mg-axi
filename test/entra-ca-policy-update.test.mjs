import assert from "node:assert/strict";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import { capabilityDocument, skillCommandTable } from "../dist/docs.js";

// WRITE-04 acceptance through the real CLI interface with fixture-only
// transports: reviewed-field allowlist refusal, current-versus-proposed
// diff preview through the READ-03 route, lockout-analysis refusal and
// acknowledgement, disabled operation when analysis is unavailable, typed
// confirmation, refusal without opt-in or under forced read-only, unknown
// outcome on timeout, journal intent/outcome. No live instance, real
// credential or customer data exists.

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const policyId = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const WRITE_SCOPES = [
  "https://graph.microsoft.com/Policy.Read.All",
  "https://graph.microsoft.com/Policy.ReadWrite.ConditionalAccess",
];
const READ_SCOPES = ["https://graph.microsoft.com/Policy.Read.All"];
const MUTATION_URL = `https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/${policyId}`;
const POLICY_SELECT = "id,displayName,state,conditions,grantControls,sessionControls";

function basePolicy(overrides = {}) {
  return {
    id: policyId,
    displayName: "Baseline CA policy",
    state: "enabled",
    conditions: {
      users: { includeUsers: ["11111111-2222-4333-8444-555555555555"], excludeUsers: [], excludeGroups: [] },
    },
    grantControls: { operator: "OR", builtInControls: ["mfa"] },
    sessionControls: { applicationEnforcedRestrictions: null, cloudAppSecurity: null },
    ...overrides,
  };
}

function allUsersPolicy(overrides = {}) {
  return basePolicy({
    conditions: { users: { includeUsers: ["All"], excludeUsers: [], excludeGroups: [] } },
    ...overrides,
  });
}

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-write-04-"));
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
  config.profiles[name].writes = { allowWrites: true, operations: ["entra.conditional-access.policy.update"] };
  writeFileSync(path, JSON.stringify(config, null, 2));
}

function teardown(state) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function json(status, body) {
  return { status, headers: {}, body: JSON.stringify(body) };
}

// reads: scripted policy answers (object, or { status } for read errors);
// every read must hit the READ-03 policy show route with the reviewed
// select. mutation: scripted PATCH answers or thrown errors.
function fixture({ reads = [basePolicy()], mutation = [{ status: 204, body: "" }] } = {}) {
  const readRequests = [];
  const mutRequests = [];
  const credCalls = [];
  let readIndex = 0;
  let mutIndex = 0;
  const transport = async request => {
    readRequests.push(request);
    const url = new URL(request.url);
    const match = /^\/v1\.0\/identity\/conditionalAccess\/policies\/([^/]+)$/.exec(url.pathname);
    assert.ok(match, `unexpected read route ${url.pathname}`);
    assert.equal(url.searchParams.get("$select"), POLICY_SELECT);
    const next = readIndex < reads.length ? reads[readIndex++] : reads[reads.length - 1];
    if (next !== null && typeof next === "object" && "status" in next && !("id" in next)) {
      return json(next.status, { error: { code: "fixture-read-error", message: "synthetic read failure" } });
    }
    return json(200, next);
  };
  const mutationTransport = async request => {
    mutRequests.push(request);
    const next = mutIndex < mutation.length ? mutation[mutIndex++] : mutation[mutation.length - 1];
    if (next instanceof Error) throw next;
    return { status: next.status, headers: {}, body: next.body ?? "" };
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
  const journalPath = join(mkdtempSync(join(tmpdir(), "mg-axi-write-04-journal-")), "writes.log");
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

function updateArgs(extra = []) {
  return ["entra", "conditional-access", "policy", "update", "--id", policyId, "--profile", "soc", ...extra];
}

test("preview shows the current-versus-proposed diff and sends nothing", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    const result = await executeArgv(updateArgs(["--state", "disabled"]), f.overrides);
    assert.equal(result.preview.operation, "entra.conditional-access.policy.update");
    assert.equal(result.preview.target, policyId);
    assert.equal(result.preview.method, "PATCH");
    assert.equal(result.preview.url, MUTATION_URL);
    assert.equal(result.preview.effect, "disruptive");
    assert.deepEqual(result.preview.changes, [{ field: "state", current: "enabled", proposed: "disabled" }]);
    assert.equal(result.preview.noop, false);
    assert.deepEqual(result.preview.lockout, {
      available: true,
      level: "none",
      findings: ["policy state 'disabled' does not enforce access"],
    });
    assert.match(result.preview.concurrency, /no ETag|no If-Match|no concurrency/i);
    assert.ok(result.help.some(hint => hint.includes("--execute") && hint.includes("--confirm") && hint.includes(policyId)));
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath), "dry-run journals nothing");
  } finally { teardown(state); }
});

test("disable sends the exact PATCH target and body with the documented scopes", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    const result = await executeArgv(updateArgs(["--state", "disabled", "--execute", "--confirm", policyId]), f.overrides);
    assert.deepEqual(result.policy, { id: policyId });
    assert.equal(typeof result.auditId, "string");
    assert.equal(f.mutRequests.length, 1);
    const [sent] = f.mutRequests;
    assert.equal(sent.method, "PATCH");
    assert.equal(sent.url, MUTATION_URL);
    assert.equal(sent.body, JSON.stringify({ state: "disabled" }));
    assert.equal(sent.headers["Content-Type"], "application/json");
    assert.match(sent.headers.Authorization, /^Bearer opaque-fixture-delegated-token$/);
    assert.equal(typeof sent.headers["client-request-id"], "string");
    assert.ok(f.credCalls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(WRITE_SCOPES)));
    assert.ok(f.credCalls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(READ_SCOPES)));
    assert.equal(f.readRequests.length, 3);
    const records = journal(f.journalPath);
    assert.equal(records.length, 2);
    assert.equal(records[0].kind, "intent");
    assert.equal(records[0].operation, "entra.conditional-access.policy.update");
    assert.equal(records[0].method, "PATCH");
    assert.equal(records[0].target, policyId);
    assert.equal(records[0].effect, "disruptive");
    assert.equal(records[1].kind, "outcome");
    assert.equal(records[1].outcome, "SUCCESS");
    assert.equal(records[1].httpStatus, 204);
    assert.equal(records[1].intentKey, records[0].intentKey);
    const file = readFileSync(f.journalPath, "utf8");
    assert.ok(!file.includes("opaque-fixture"), "journal carries no credentials");
    assert.ok(!file.includes("disabled"), "journal carries no payload");
  } finally { teardown(state); }
});

test("the PATCH body carries only reviewed fields", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    const result = await executeArgv(updateArgs([
      "--display-name", "Renamed policy",
      "--state", "disabled",
      "--execute", "--confirm", policyId,
    ]), f.overrides);
    assert.deepEqual(result.policy, { id: policyId });
    assert.equal(f.mutRequests.length, 1);
    assert.equal(f.mutRequests[0].body, JSON.stringify({ displayName: "Renamed policy", state: "disabled" }));
    assert.deepEqual(Object.keys(JSON.parse(f.mutRequests[0].body)).sort(),
      ["displayName", "state"].sort());
  } finally { teardown(state); }
});

test("already-desired values are a no-op without sending", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    const result = await executeArgv(updateArgs([
      "--display-name", "Baseline CA policy",
      "--state", "enabled",
      "--execute", "--confirm", policyId,
    ]), f.overrides);
    assert.deepEqual(result, { noop: true, policy: { id: policyId } });
    assert.equal(f.mutRequests.length, 0);
  } finally { teardown(state); }
});

test("an unreviewed field flag is refused before credentials", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(
      executeArgv(updateArgs(["--created-date-time", "2026-01-01T00:00:00Z"]), f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("an update without any reviewed field is a usage error", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(executeArgv(updateArgs(), f.overrides), { code: "VALIDATION_ERROR" });
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

for (const value of ["ENABLED", "on", ""]) test(`state ${JSON.stringify(value)} is a usage error`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(executeArgv(updateArgs(["--state", value]), f.overrides), { code: "VALIDATION_ERROR" });
    assert.equal(f.credCalls.length, 0);
    assert.equal(f.mutRequests.length, 0);
  } finally { teardown(state); }
});

for (const value of ["not-json", "[1,2]", "null", "42"]) test(`conditions ${JSON.stringify(value)} must be a JSON object`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(executeArgv(updateArgs(["--conditions", value]), f.overrides), { code: "VALIDATION_ERROR" });
    assert.equal(f.credCalls.length, 0);
    assert.equal(f.mutRequests.length, 0);
  } finally { teardown(state); }
});

test("lockout refusal: enabling block for all users without exclusions is refused", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const blocking = allUsersPolicy({ grantControls: { operator: "OR", builtInControls: ["block"] } });
    const preview = await executeArgv(
      updateArgs(["--grant-controls", JSON.stringify({ operator: "OR", builtInControls: ["block", "mfa"] })]),
      fixture({ reads: [blocking] }).overrides,
    );
    assert.deepEqual(preview.preview.lockout, {
      available: true,
      level: "refused",
      findings: ["enabled policy would block all users with no exclusions: every admin including break-glass access would be locked out"],
    });
    assert.equal(preview.preview.noop, false);
    const f = fixture({ reads: [blocking] });
    const error = await executeArgv(updateArgs([
      "--grant-controls", JSON.stringify({ operator: "OR", builtInControls: ["block", "mfa"] }),
      "--execute", "--confirm", policyId,
    ]), f.overrides)
      .then(() => assert.fail("lockout must refuse"), caught => caught);
    assert.equal(error.code, "OPERATION_BLOCKED");
    assert.match(error.message, /locking out every admin including break-glass/i);
    assert.match(error.suggestions.join("\n"), /cannot be overridden/i);
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath), "a refused execute reserves no intent");
  } finally { teardown(state); }
});

test("lockout acknowledgement: all-users coverage without exclusions needs explicit acknowledgement", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const broad = allUsersPolicy({ state: "enabledForReportingButNotEnforced" });
    const preview = await executeArgv(
      updateArgs(["--state", "enabled"]),
      fixture({ reads: [broad] }).overrides,
    );
    assert.equal(preview.preview.lockout.level, "elevated");
    assert.ok(preview.help.some(hint => hint.includes("--acknowledge-lockout-risk")));
    const refused = fixture({ reads: [broad] });
    const error = await executeArgv(updateArgs(["--state", "enabled", "--execute", "--confirm", policyId]), refused.overrides)
      .then(() => assert.fail("unacknowledged risk must refuse"), caught => caught);
    assert.equal(error.code, "LOCKOUT_ACK_REQUIRED");
    assert.equal(refused.mutRequests.length, 0);
    const f = fixture({ reads: [broad, broad, broad, broad] });
    const result = await executeArgv(updateArgs([
      "--state", "enabled", "--execute", "--confirm", policyId, "--acknowledge-lockout-risk",
    ]), f.overrides);
    assert.deepEqual(result.policy, { id: policyId });
    assert.equal(f.mutRequests.length, 1);
    assert.equal(f.mutRequests[0].body, JSON.stringify({ state: "enabled" }));
    assert.equal(journal(f.journalPath)[1].outcome, "SUCCESS");
  } finally { teardown(state); }
});

for (const targeting of [
  { includeUsers: ["All"], excludeUsers: ["non-admin"] },
  { includeUsers: ["All"], excludeGroups: ["non-admin-group"] },
  { includeUsers: ["All"], excludeRoles: ["non-admin-role"] },
  { includeUsers: ["admin"] },
  { includeUsers: [], includeGroups: ["admins"] },
  { includeUsers: [], includeRoles: ["admin-role"] },
]) test(`unverified targeting ${JSON.stringify(targeting)} requires acknowledgement`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const policy = basePolicy({
      state: "disabled",
      conditions: { users: { excludeUsers: [], excludeGroups: [], ...targeting } },
      grantControls: { builtInControls: ["block"] },
    });
    const f = fixture({ reads: [policy] });
    const preview = await executeArgv(updateArgs(["--state", "enabled"]), f.overrides);
    assert.equal(preview.preview.lockout.level, "elevated");
    await assert.rejects(
      executeArgv(updateArgs(["--state", "enabled", "--execute", "--confirm", policyId]), f.overrides),
      { code: "LOCKOUT_ACK_REQUIRED" },
    );
    assert.equal(f.mutRequests.length, 0);
    const result = await executeArgv(updateArgs([
      "--state", "enabled", "--execute", "--confirm", policyId, "--acknowledge-lockout-risk",
    ]), f.overrides);
    assert.deepEqual(result.policy, { id: policyId });
    assert.equal(f.mutRequests.length, 1);
  } finally { teardown(state); }
});

for (const missing of ["excludeUsers", "excludeGroups"]) test(`omitted ${missing} disables conditions updates`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const conditions = { users: { includeUsers: ["All"], excludeUsers: [], excludeGroups: [] } };
    delete conditions.users[missing];
    const f = fixture();
    const fields = ["--conditions", JSON.stringify(conditions)];
    const preview = await executeArgv(updateArgs(fields), f.overrides);
    assert.equal(preview.preview.lockout.available, false);
    await assert.rejects(
      executeArgv(updateArgs([...fields, "--execute", "--confirm", policyId, "--acknowledge-lockout-risk"]), f.overrides),
      { code: "OPERATION_BLOCKED", message: /stays disabled without lockout analysis/ },
    );
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath));
  } finally { teardown(state); }
});

for (const field of ["includeUsers", "includeGroups", "includeRoles", "excludeUsers", "excludeGroups", "excludeRoles"]) test(`unreadable ${field} disables enforcement`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const policy = basePolicy({
      conditions: { users: { includeUsers: ["admin"], excludeUsers: [], excludeGroups: [], [field]: null } },
    });
    const f = fixture({ reads: [policy] });
    await assert.rejects(
      executeArgv(updateArgs(["--session-controls", "{}", "--execute", "--confirm", policyId, "--acknowledge-lockout-risk"]), f.overrides),
      { code: "OPERATION_BLOCKED" },
    );
    assert.equal(f.mutRequests.length, 0);
  } finally { teardown(state); }
});

for (const readIndex of [1, 2]) test(`fresh read ${readIndex} rechecks protected-account coverage`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const safe = basePolicy({ state: "disabled" });
    const reads = [safe, safe, safe];
    reads[readIndex] = basePolicy();
    const f = fixture({ reads });
    await assert.rejects(
      executeArgv(updateArgs(["--session-controls", "{}", "--execute", "--confirm", policyId]), f.overrides),
      { code: "LOCKOUT_ACK_REQUIRED" },
    );
    assert.equal(f.mutRequests.length, 0);
  } finally { teardown(state); }
});

test("disabled operation: unreadable conditions keep enforcement changes disabled", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const blind = basePolicy({ state: "disabled", conditions: null });
    const preview = await executeArgv(
      updateArgs(["--state", "enabled"]),
      fixture({ reads: [blind] }).overrides,
    );
    assert.deepEqual(preview.preview.lockout, {
      available: false,
      reason: "policy conditions are unreadable, so lockout risk cannot be assessed",
    });
    const f = fixture({ reads: [blind] });
    const error = await executeArgv(updateArgs(["--state", "enabled", "--execute", "--confirm", policyId]), f.overrides)
      .then(() => assert.fail("blind update must stay disabled"), caught => caught);
    assert.equal(error.code, "OPERATION_BLOCKED");
    assert.match(error.message, /stays disabled without lockout analysis/i);
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath), "a disabled operation reserves no intent");
  } finally { teardown(state); }
});

test("display-name-only changes need no lockout analysis", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const blind = basePolicy({ conditions: null });
    const f = fixture({ reads: [blind] });
    const result = await executeArgv(updateArgs([
      "--display-name", "Renamed policy", "--execute", "--confirm", policyId,
    ]), f.overrides);
    assert.deepEqual(result.policy, { id: policyId });
    assert.equal(f.mutRequests.length, 1);
    assert.equal(f.mutRequests[0].body, JSON.stringify({ displayName: "Renamed policy" }));
  } finally { teardown(state); }
});

for (const extra of [["--execute"], ["--execute", "--confirm", "someone-else"]]) test(`confirm ${JSON.stringify(extra)} is refused before reads`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    const error = await executeArgv(updateArgs(["--state", "disabled", ...extra]), f.overrides)
      .then(() => assert.fail("confirmation must be required"), caught => caught);
    assert.ok(["CONFIRM_REQUIRED", "CONFIRM_MISMATCH"].includes(error.code));
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("profile without opt-in is refused with zero sends and zero reads", async () => {
  const state = setupProfiles();
  try {
    const f = fixture();
    await assert.rejects(
      executeArgv(updateArgs(["--state", "disabled", "--execute", "--confirm", policyId]), f.overrides),
      { code: "WRITES_DISABLED" },
    );
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
      await assert.rejects(
        executeArgv(updateArgs(["--state", "disabled", "--execute", "--confirm", policyId]), f.overrides),
        { code: "WRITES_DISABLED" },
      );
    });
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("transport failure records an unknown outcome", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ mutation: [new Error("socket timeout")] });
    const error = await executeArgv(updateArgs(["--state", "disabled", "--execute", "--confirm", policyId]), f.overrides)
      .then(() => assert.fail("timeout must throw"), caught => caught);
    assert.equal(error.code, "OUTCOME_UNKNOWN");
    assert.match(error.suggestions.join("\n"), /Never replay this intent/);
    assert.equal(f.mutRequests.length, 1);
    assert.equal(journal(f.journalPath)[1].outcome, "OUTCOME_UNKNOWN");
  } finally { teardown(state); }
});

test("denial surfaces the permission pair and administrator roles", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ mutation: [{ status: 403, body: { error: { code: "Authorization_RequestDenied", message: "denied" } } }] });
    const error = await executeArgv(updateArgs(["--state", "disabled", "--execute", "--confirm", policyId]), f.overrides)
      .then(() => assert.fail("denial must throw"), caught => caught);
    assert.equal(error.code, "GRAPH_ERROR");
    const guidance = error.suggestions.join("\n");
    assert.match(guidance, /Policy\.ReadWrite\.ConditionalAccess/);
    assert.match(guidance, /Conditional Access Administrator/);
    assert.match(guidance, /never replay/i);
    assert.equal(journal(f.journalPath)[1].outcome, "FAILED");
  } finally { teardown(state); }
});

test("a read without a Graph object ID blocks the update", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    f.overrides.transport = async () => json(200, { displayName: "Baseline CA policy" });
    await assert.rejects(
      executeArgv(updateArgs(["--state", "disabled", "--execute", "--confirm", policyId]), f.overrides),
      { code: "GRAPH_ERROR" },
    );
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath));
  } finally { teardown(state); }
});

test("failed fresh read journals NOT_SENT and sends nothing", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [basePolicy(), basePolicy(), { status: 500 }] });
    await assert.rejects(
      executeArgv(updateArgs(["--state", "disabled", "--execute", "--confirm", policyId]), f.overrides),
    );
    assert.equal(f.mutRequests.length, 0);
    const records = journal(f.journalPath);
    assert.equal(records.length, 2);
    assert.equal(records[0].kind, "intent");
    assert.equal(records[1].outcome, "NOT_SENT");
  } finally { teardown(state); }
});

test("beta api-version is refused before credentials", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(
      executeArgv(["entra", "conditional-access", "policy", "update", "--id", policyId, "--state", "disabled",
        "--profile", "soc", "--api-version", "beta"], f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.credCalls.length, 0);
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 0);
  } finally { teardown(state); }
});

test("application mode uses the .default audience and succeeds", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir, "batch");
    const f = fixture();
    const result = await executeArgv(
      ["entra", "conditional-access", "policy", "update", "--id", policyId, "--state", "disabled",
        "--profile", "batch", "--execute", "--confirm", policyId],
      f.overrides,
    );
    assert.deepEqual(result.policy, { id: policyId });
    assert.equal(f.mutRequests.length, 1);
    assert.ok(f.credCalls.length > 0);
    for (const [mode, , scopes] of f.credCalls) {
      if (mode === "application") assert.equal(scopes, undefined, "application credentials use the .default audience, never delegated scopes");
    }
  } finally { teardown(state); }
});

test("generated records name the policy update", async () => {
  assert.ok(skillCommandTable().includes("`mg-axi entra conditional-access policy update` | native | write |"));
  const coverage = capabilityDocument();
  assert.ok(coverage.includes("## Named writes"));
  assert.ok(coverage.includes("`mg-axi entra conditional-access policy update` | `PATCH:/identity/conditionalAccess/policies/{conditionalAccessPolicy-id}` | WRITE-04 |"));
});
