import assert from "node:assert/strict";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import { capabilityDocument, skillCommandTable } from "../dist/docs.js";
import { setupView } from "../dist/setup.js";

// WRITE-05 acceptance through the real CLI interface with fixture-only
// transports: single-user POST body, bulk refusal, preview with the current
// risk state read through the READ-06 route, dismissal-vs-remediation
// wording, P2/permission/role guidance, typed confirmation, refusal without
// opt-in or under forced read-only, unknown outcome on timeout or 5xx, and
// journal intent/outcome. No live instance, real credential or customer
// data exists.

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const userId = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const WRITE_SCOPES = ["https://graph.microsoft.com/IdentityRiskyUser.ReadWrite.All"];
const READ_SCOPES = ["https://graph.microsoft.com/IdentityRiskyUser.Read.All"];
const READ_SELECT = "id,userPrincipalName,riskLevel,riskState,riskDetail,riskLastUpdatedDateTime";
const MUTATION_URL = "https://graph.microsoft.com/v1.0/identityProtection/riskyUsers/dismiss";

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-write-05-"));
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
  config.profiles[name].writes = { allowWrites: true, operations: ["entra.risky-user.dismiss"] };
  writeFileSync(path, JSON.stringify(config, null, 2));
}

function teardown(state) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function json(status, body) {
  return { status, headers: {}, body: JSON.stringify(body) };
}

function riskRow(riskState) {
  return {
    id: userId,
    userPrincipalName: "analyst@contoso.com",
    riskLevel: riskState === "dismissed" ? "none" : "high",
    riskState,
    riskDetail: riskState === "dismissed" ? "none" : "adminConfirmedUserCompromised",
    riskLastUpdatedDateTime: "2026-10-01T00:00:00Z",
  };
}

// reads: scripted riskState answers (string), { status } for read errors, or
// "malformed" for a row without an object ID. Every read must hit the
// READ-06 risky-user show route with the risk-state select. mutation:
// scripted POST answers or thrown errors.
function fixture({ reads = ["atRisk"], mutation = [{ status: 204, body: "" }] } = {}) {
  const readRequests = [];
  const mutRequests = [];
  const credCalls = [];
  let readIndex = 0;
  let mutIndex = 0;
  const transport = async request => {
    readRequests.push(request);
    const url = new URL(request.url);
    assert.equal(url.pathname, `/v1.0/identityProtection/riskyUsers/${userId}`);
    assert.equal(url.searchParams.get("$select"), READ_SELECT);
    const next = readIndex < reads.length ? reads[readIndex++] : reads[reads.length - 1];
    if (next !== null && typeof next === "object") {
      return json(next.status, { error: { code: "fixture-read-error", message: "synthetic read failure" } });
    }
    if (next === "malformed") return json(200, { riskState: "atRisk" });
    return json(200, riskRow(next));
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
  const journalPath = join(mkdtempSync(join(tmpdir(), "mg-axi-write-05-journal-")), "writes.log");
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

function dismissArgs(user, extra = []) {
  return ["entra", "risky-user", "dismiss", "--user", user, "--profile", "soc", ...extra];
}

test("preview shows the current risk state and sends nothing", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["atRisk"] });
    const result = await executeArgv(dismissArgs(userId), f.overrides);
    assert.deepEqual(result.preview, {
      operation: "entra.risky-user.dismiss",
      target: userId,
      method: "POST",
      url: MUTATION_URL,
      effect: "disruptive",
      current: riskRow("atRisk"),
      desired: "dismissed",
      noop: false,
    });
    assert.ok(result.help.some(hint => hint.includes("--execute") && hint.includes("--confirm") && hint.includes(userId)));
    assert.ok(result.help.some(hint => hint.includes("not remediation") && hint.includes("revoking sessions")));
    assert.ok(result.help.some(hint => hint.includes("P2")));
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath), "dry-run journals nothing");
  } finally { teardown(state); }
});

test("bulk user list is refused before credentials", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(
      executeArgv(dismissArgs(`${userId},bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb`, ["--execute", "--confirm", userId]), f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("a plural user-ids flag is an unknown flag", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(
      executeArgv(["entra", "risky-user", "dismiss", "--user-ids", userId, "--profile", "soc"], f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("dismiss sends the exact POST target and single-user body, then verifies", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["atRisk", "atRisk", "atRisk", "dismissed"] });
    const result = await executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides);
    assert.deepEqual(result.dismissal, { user: riskRow("dismissed") });
    assert.equal(typeof result.auditId, "string");
    assert.equal(f.mutRequests.length, 1);
    const [sent] = f.mutRequests;
    assert.equal(sent.method, "POST");
    assert.equal(sent.url, MUTATION_URL);
    assert.equal(sent.body, JSON.stringify({ userIds: [userId] }));
    assert.equal(sent.headers["Content-Type"], "application/json");
    assert.match(sent.headers.Authorization, /^Bearer opaque-fixture-delegated-token$/);
    assert.equal(typeof sent.headers["client-request-id"], "string");
    assert.ok(f.credCalls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(WRITE_SCOPES)));
    assert.ok(f.credCalls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(READ_SCOPES)));
    const records = journal(f.journalPath);
    assert.equal(records.length, 2);
    assert.equal(records[0].kind, "intent");
    assert.equal(records[0].operation, "entra.risky-user.dismiss");
    assert.equal(records[0].method, "POST");
    assert.equal(records[0].target, userId);
    assert.equal(records[0].effect, "disruptive");
    assert.equal(records[1].kind, "outcome");
    assert.equal(records[1].outcome, "SUCCESS");
    assert.equal(records[1].httpStatus, 204);
    assert.equal(records[1].intentKey, records[0].intentKey);
    const file = readFileSync(f.journalPath, "utf8");
    assert.ok(!file.includes("opaque-fixture"), "journal carries no credentials");
    assert.ok(!file.includes("userIds"), "journal carries no payload");
  } finally { teardown(state); }
});

test("already-dismissed user is a no-op without sending", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["dismissed", "dismissed"] });
    const result = await executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides);
    assert.deepEqual(result, { noop: true, dismissal: { user: { id: userId, riskState: "dismissed" } } });
    assert.equal(f.mutRequests.length, 0);
  } finally { teardown(state); }
});

for (const extra of [[], ["--execute"]]) test(`dismiss ${extra.join(" ") || "preview"} without --confirm is refused before credentials`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["atRisk"] });
    const args = extra.length ? dismissArgs(userId, ["--execute"]) : dismissArgs(userId);
    if (!extra.length) {
      const result = await executeArgv(args, f.overrides);
      assert.equal(result.preview.noop, false);
      return;
    }
    await assert.rejects(executeArgv(args, f.overrides), { code: "CONFIRM_REQUIRED" });
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("mismatched --confirm is refused before credentials", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["atRisk"] });
    await assert.rejects(
      executeArgv(dismissArgs(userId, ["--execute", "--confirm", "someone-else"]), f.overrides),
      { code: "CONFIRM_MISMATCH" },
    );
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("post-dismiss reread mismatch reports a conflict", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["atRisk", "atRisk", "atRisk", "atRisk"] });
    const error = await executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("conflict must throw"), caught => caught);
    assert.equal(error.code, "WRITE_CONFLICT");
    assert.match(error.suggestions.join("\n"), /[Rr]ead back/);
    assert.match(error.suggestions.join("\n"), /Never replay this intent/);
    assert.equal(f.mutRequests.length, 1);
    assert.equal(journal(f.journalPath)[1].outcome, "SUCCESS");
  } finally { teardown(state); }
});

for (const scenario of [
  { name: "preview", extra: [], reads: [{ status: 403 }], code: "GRAPH_ERROR", sends: 0, outcome: null },
  { name: "initial execution read", extra: ["--execute", "--confirm", userId], reads: [{ status: 403 }], code: "GRAPH_ERROR", sends: 0, outcome: null },
  { name: "coordinator preview read", extra: ["--execute", "--confirm", userId], reads: ["atRisk", { status: 403 }], code: "GRAPH_ERROR", sends: 0, outcome: null },
  { name: "fresh read", extra: ["--execute", "--confirm", userId], reads: ["atRisk", "atRisk", { status: 403 }], code: "GRAPH_ERROR", sends: 0, outcome: "NOT_SENT" },
  { name: "verification read", extra: ["--execute", "--confirm", userId], reads: ["atRisk", "atRisk", "atRisk", { status: 403 }], code: "OUTCOME_UNKNOWN", sends: 1, outcome: "SUCCESS" },
]) test(`${scenario.name} denial retains risky-user read guidance`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: scenario.reads });
    const error = await executeArgv(dismissArgs(userId, scenario.extra), f.overrides)
      .then(() => assert.fail("read denial must throw"), caught => caught);
    assert.equal(error.code, scenario.code);
    const guidance = error.suggestions.join("\n");
    assert.match(guidance, /IdentityRiskyUser\.Read\.All/);
    assert.match(guidance, /Global Reader, Security Operator, Security Reader or Security Administrator/);
    assert.match(guidance, /admin-consented/);
    assert.match(guidance, /P2/);
    assert.match(guidance, /A 403 never proves which prerequisite is missing/);
    assert.equal(f.mutRequests.length, scenario.sends);
    if (scenario.outcome === null) assert.ok(!existsSync(f.journalPath));
    else assert.equal(journal(f.journalPath)[1].outcome, scenario.outcome);
    if (scenario.code === "OUTCOME_UNKNOWN") assert.match(guidance, /never replay this intent/);
  } finally { teardown(state); }
});

test("denial surfaces P2, permission and role guidance", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["atRisk"], mutation: [{ status: 403, body: { error: { code: "Authorization_RequestDenied", message: "denied" } } }] });
    const error = await executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("denial must throw"), caught => caught);
    assert.equal(error.code, "GRAPH_ERROR");
    const guidance = error.suggestions.join("\n");
    assert.match(guidance, /IdentityRiskyUser\.ReadWrite\.All/);
    assert.match(guidance, /Security Administrator/);
    assert.match(guidance, /P2/);
    assert.match(guidance, /never replay/i);
    assert.equal(journal(f.journalPath)[1].outcome, "FAILED");
  } finally { teardown(state); }
});

for (const status of [400, 404, 429]) test(`dismissal refusal ${status} preserves status without permission or role diagnosis`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["atRisk"], mutation: [{ status }] });
    const error = await executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("refusal must throw"), caught => caught);
    assert.equal(error.code, "GRAPH_ERROR");
    assert.equal(error.message, `Graph refused the risk dismissal (status ${status})`);
    assert.equal(error.suggestions.length, 1);
    assert.ok(error.suggestions[0].includes(`read back '${userId}' and never replay`));
    assert.equal(f.mutRequests.length, 1);
    assert.equal(journal(f.journalPath)[1].httpStatus, status);
  } finally { teardown(state); }
});

test("profile without opt-in is refused with zero sends and zero reads", async () => {
  const state = setupProfiles();
  try {
    const f = fixture();
    await assert.rejects(executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides), { code: "WRITES_DISABLED" });
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
      await assert.rejects(executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides), { code: "WRITES_DISABLED" });
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
    const f = fixture({ reads: ["atRisk"], mutation: [new Error("socket timeout")] });
    const error = await executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("timeout must throw"), caught => caught);
    assert.equal(error.code, "OUTCOME_UNKNOWN");
    assert.match(error.suggestions.join("\n"), /Never replay this intent/);
    assert.equal(f.mutRequests.length, 1);
    assert.equal(journal(f.journalPath)[1].outcome, "OUTCOME_UNKNOWN");
  } finally { teardown(state); }
});

test("server error after send records an unknown outcome with no replay", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["atRisk"], mutation: [{ status: 503, body: { error: { code: "Service_Unavailable", message: "try again" } } }] });
    const error = await executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("5xx must throw"), caught => caught);
    assert.equal(error.code, "OUTCOME_UNKNOWN");
    assert.match(error.suggestions.join("\n"), /Never replay this intent/);
    assert.equal(f.mutRequests.length, 1);
    assert.equal(journal(f.journalPath)[1].outcome, "OUTCOME_UNKNOWN");
  } finally { teardown(state); }
});

test("unknown target sends nothing and journals no intent", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [{ status: 404 }] });
    await assert.rejects(executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides));
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath), "a failed preview read reserves no intent");
  } finally { teardown(state); }
});

test("a read without a Graph object ID blocks the dismissal", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["malformed"] });
    await assert.rejects(executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides), { code: "GRAPH_ERROR" });
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath));
  } finally { teardown(state); }
});

test("failed fresh read journals NOT_SENT and sends nothing", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: ["atRisk", "atRisk", { status: 404 }] });
    await assert.rejects(executeArgv(dismissArgs(userId, ["--execute", "--confirm", userId]), f.overrides));
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
      executeArgv(["entra", "risky-user", "dismiss", "--user", userId, "--profile", "soc", "--api-version", "beta"], f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.credCalls.length, 0);
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 0);
  } finally { teardown(state); }
});

test("caller scopes are rejected: only the documented scopes are requested", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(
      executeArgv([...dismissArgs(userId), "--scopes", "https://graph.microsoft.com/IdentityRiskyUser.Read.All"], f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("application mode uses the .default audience and succeeds", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir, "batch");
    const f = fixture({ reads: ["atRisk", "atRisk", "atRisk", "dismissed"] });
    const result = await executeArgv(
      ["entra", "risky-user", "dismiss", "--user", userId, "--profile", "batch", "--execute", "--confirm", userId],
      f.overrides,
    );
    assert.deepEqual(result.dismissal, { user: riskRow("dismissed") });
    assert.equal(f.mutRequests.length, 1);
    assert.ok(f.credCalls.length > 0);
    for (const [mode, , scopes] of f.credCalls) {
      if (mode === "application") assert.equal(scopes, undefined, "application credentials use the .default audience, never delegated scopes");
    }
  } finally { teardown(state); }
});

test("generated records name the risk dismissal write", async () => {
  assert.ok(skillCommandTable().includes("`mg-axi entra risky-user dismiss` | native | write |"));
  const coverage = capabilityDocument();
  assert.ok(coverage.includes("## Named writes"));
  assert.ok(coverage.includes("`mg-axi entra risky-user dismiss` | `POST:/identityProtection/riskyUsers/dismiss` | WRITE-05 |"));
});

test("setup counts the dismissal write beside reads and local leaves", async () => {
  const state = setupProfiles();
  try {
    const output = setupView(new Profiles());
    assert.equal(output.capabilities.reads, 52);
    assert.equal(output.capabilities.writes, 5);
    assert.equal(output.capabilities.local, 6);
    assert.ok(output.capabilities.implemented.includes("entra risky-user dismiss"));
  } finally { teardown(state); }
});
