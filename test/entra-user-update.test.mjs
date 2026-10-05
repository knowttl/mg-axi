import assert from "node:assert/strict";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import { capabilityDocument, skillCommandTable } from "../dist/docs.js";
import { setupView } from "../dist/setup.js";

// WRITE-02 acceptance through the real CLI interface with fixture-only
// transports: exact PATCH target and body, desired-state no-op, post-write
// reread conflict, permission/role-hierarchy guidance, refusal without opt-in
// or under forced read-only, unknown outcome on timeout, journal
// intent/outcome. No live instance, real credential or customer data exists.

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const userId = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const WRITE_SCOPES = ["https://graph.microsoft.com/User.EnableDisableAccount.All", "https://graph.microsoft.com/User.Read.All"];
const READ_SCOPES = ["https://graph.microsoft.com/User.Read.All"];
const MUTATION_URL = `https://graph.microsoft.com/v1.0/users/${userId}`;

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-write-02-"));
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
  config.profiles[name].writes = { allowWrites: true, operations: ["entra.user.update"] };
  writeFileSync(path, JSON.stringify(config, null, 2));
}

function teardown(state) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function json(status, body) {
  return { status, headers: {}, body: JSON.stringify(body) };
}

// reads: scripted accountEnabled answers (boolean, null, or { status } for
// read errors); every read must hit the READ-01 user route with the
// id and accountEnabled select. mutation: scripted PATCH answers or thrown errors.
function fixture({ reads = [true], mutation = [{ status: 204, body: "" }] } = {}) {
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
    assert.equal(url.searchParams.get("$select"), "id,accountEnabled");
    const next = readIndex < reads.length ? reads[readIndex++] : reads[reads.length - 1];
    if (next !== null && typeof next === "object") {
      return json(next.status, { error: { code: "fixture-read-error", message: "synthetic read failure" } });
    }
    return json(200, { id: userId, accountEnabled: next });
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
  const journalPath = join(mkdtempSync(join(tmpdir(), "mg-axi-write-02-journal-")), "writes.log");
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

function updateArgs(user, desired, extra = []) {
  return ["entra", "user", "update", "--user", user, "--account-enabled", desired, "--profile", "soc", ...extra];
}

test("preview shows the desired-state diff and sends nothing", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [true] });
    const result = await executeArgv(updateArgs(userId, "false"), f.overrides);
    assert.deepEqual(result.preview, {
      operation: "entra.user.update",
      target: userId,
      method: "PATCH",
      url: MUTATION_URL,
      effect: "disruptive",
      current: true,
      desired: false,
      noop: false,
    });
    assert.ok(result.help.some(hint => hint.includes("--execute") && hint.includes("--confirm") && hint.includes(userId)));
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath), "dry-run journals nothing");
  } finally { teardown(state); }
});

test("disable sends the exact PATCH target and body, then verifies", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [true, true, true, false] });
    const result = await executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides);
    assert.deepEqual(result.user, { id: userId, accountEnabled: false });
    assert.equal(typeof result.auditId, "string");
    assert.equal(f.mutRequests.length, 1);
    const [sent] = f.mutRequests;
    assert.equal(sent.method, "PATCH");
    assert.equal(sent.url, MUTATION_URL);
    assert.equal(sent.body, JSON.stringify({ accountEnabled: false }));
    assert.equal(sent.headers["Content-Type"], "application/json");
    assert.match(sent.headers.Authorization, /^Bearer opaque-fixture-delegated-token$/);
    assert.equal(typeof sent.headers["client-request-id"], "string");
    assert.ok(f.credCalls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(WRITE_SCOPES)));
    assert.ok(f.credCalls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(READ_SCOPES)));
    const records = journal(f.journalPath);
    assert.equal(records.length, 2);
    assert.equal(records[0].kind, "intent");
    assert.equal(records[0].operation, "entra.user.update");
    assert.equal(records[0].method, "PATCH");
    assert.equal(records[0].target, userId);
    assert.equal(records[0].effect, "disruptive");
    assert.equal(records[1].kind, "outcome");
    assert.equal(records[1].outcome, "SUCCESS");
    assert.equal(records[1].httpStatus, 204);
    assert.equal(records[1].intentKey, records[0].intentKey);
    const file = readFileSync(f.journalPath, "utf8");
    assert.ok(!file.includes("opaque-fixture"), "journal carries no credentials");
    assert.ok(!file.includes("accountEnabled"), "journal carries no payload");
  } finally { teardown(state); }
});

test("enable requires confirmation and records a disruptive write", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [false, false, false, true] });
    const preview = await executeArgv(updateArgs(userId, "true"), fixture({ reads: [false] }).overrides);
    assert.equal(preview.preview.effect, "disruptive");
    assert.ok(preview.help.some(hint => hint.includes("--execute") && hint.includes(`--confirm ${userId}`)));
    const result = await executeArgv(updateArgs(userId, "true", ["--execute", "--confirm", userId]), f.overrides);
    assert.deepEqual(result.user, { id: userId, accountEnabled: true });
    assert.equal(f.mutRequests.length, 1);
    assert.equal(f.mutRequests[0].body, JSON.stringify({ accountEnabled: true }));
    assert.ok(f.credCalls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(WRITE_SCOPES)));
    assert.equal(journal(f.journalPath)[0].effect, "disruptive");
  } finally { teardown(state); }
});

for (const [target, desired] of [
  [userId, true],
  [userId, false],
  ["AdeleV@contoso.com", true],
  ["AdeleV@contoso.com", false],
  ["alice_example.com#EXT#@tenant.onmicrosoft.com", true],
  ["alice_example.com#EXT#@tenant.onmicrosoft.com", false],
]) test(`already-desired ${desired} for ${target} returns the Graph object ID without sending`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [desired] });
    const result = await executeArgv(updateArgs(target, String(desired), ["--execute", "--confirm", target]), f.overrides);
    assert.deepEqual(result, { noop: true, user: { id: userId, accountEnabled: desired } });
    assert.equal(f.mutRequests.length, 0);
    assert.deepEqual(f.readRequests.map(request => new URL(request.url).pathname), [
      `/v1.0/users/${encodeURIComponent(target)}`,
      `/v1.0/users/${userId}`,
    ]);
  } finally { teardown(state); }
});

for (const [desired, current] of [["true", false], ["false", true], ["true", true], ["false", false]]) test(`${desired} from ${current} without --confirm is refused before credentials`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [current] });
    await assert.rejects(executeArgv(updateArgs(userId, desired, ["--execute"]), f.overrides), { code: "CONFIRM_REQUIRED" });
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

for (const desired of ["true", "false"]) test(`${desired} with mismatched --confirm is refused before credentials`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [true] });
    await assert.rejects(
      executeArgv(updateArgs(userId, desired, ["--execute", "--confirm", "someone-else"]), f.overrides),
      { code: "CONFIRM_MISMATCH" },
    );
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("post-write reread mismatch reports a conflict", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [true, true, true, true] });
    const error = await executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("conflict must throw"), caught => caught);
    assert.equal(error.code, "WRITE_CONFLICT");
    assert.match(error.suggestions.join("\n"), /[Rr]ead back/);
    assert.match(error.suggestions.join("\n"), /Never replay this intent/);
    assert.equal(f.mutRequests.length, 1);
    assert.equal(journal(f.journalPath)[1].outcome, "SUCCESS");
  } finally { teardown(state); }
});

test("denial surfaces the permission pair and role hierarchy", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [true], mutation: [{ status: 403, body: { error: { code: "Authorization_RequestDenied", message: "denied" } } }] });
    const error = await executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("denial must throw"), caught => caught);
    assert.equal(error.code, "GRAPH_ERROR");
    const guidance = error.suggestions.join("\n");
    assert.match(guidance, /User\.EnableDisableAccount\.All \+ User\.Read\.All|User\.EnableDisableAccount\.All plus User\.Read\.All/);
    assert.match(guidance, /Privileged Authentication Administrator/);
    assert.match(guidance, /never replay/i);
    assert.equal(journal(f.journalPath)[1].outcome, "FAILED");
  } finally { teardown(state); }
});

for (const status of [400, 404, 429]) test(`PATCH refusal ${status} preserves status without permission or role diagnosis`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [true], mutation: [{ status }] });
    const error = await executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("refusal must throw"), caught => caught);
    assert.equal(error.code, "GRAPH_ERROR");
    assert.equal(error.message, `Graph refused the account update (status ${status})`);
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
    await assert.rejects(executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides), { code: "WRITES_DISABLED" });
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
      await assert.rejects(executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides), { code: "WRITES_DISABLED" });
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
    const f = fixture({ reads: [true], mutation: [new Error("socket timeout")] });
    const error = await executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides)
      .then(() => assert.fail("timeout must throw"), caught => caught);
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
    await assert.rejects(executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides));
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath), "a failed preview read reserves no intent");
  } finally { teardown(state); }
});

test("a read without a Graph object ID blocks the update", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    f.overrides.transport = async () => json(200, { accountEnabled: true });
    await assert.rejects(executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides), { code: "GRAPH_ERROR" });
    assert.equal(f.mutRequests.length, 0);
    assert.ok(!existsSync(f.journalPath));
  } finally { teardown(state); }
});

test("failed fresh read journals NOT_SENT and sends nothing", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [true, true, { status: 404 }] });
    await assert.rejects(executeArgv(updateArgs(userId, "false", ["--execute", "--confirm", userId]), f.overrides));
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
      executeArgv(["entra", "user", "update", "--user", userId, "--account-enabled", "false", "--profile", "soc", "--api-version", "beta"], f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.credCalls.length, 0);
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.readRequests.length, 0);
  } finally { teardown(state); }
});

for (const value of ["maybe", "TRUE", "FALSE", " true ", " false "]) test(`account-enabled ${JSON.stringify(value)} is a usage error`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(executeArgv(updateArgs(userId, value), f.overrides), { code: "VALIDATION_ERROR" });
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
      executeArgv([...updateArgs(userId, "false"), "--scopes", "https://graph.microsoft.com/User.Read.All"], f.overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("application mode uses the .default audience and succeeds", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir, "batch");
    const f = fixture({ reads: [true, true, true, false] });
    const result = await executeArgv(
      ["entra", "user", "update", "--user", userId, "--account-enabled", "false", "--profile", "batch", "--execute", "--confirm", userId],
      f.overrides,
    );
    assert.deepEqual(result.user, { id: userId, accountEnabled: false });
    assert.equal(f.mutRequests.length, 1);
    assert.ok(f.credCalls.length > 0);
    for (const [mode, , scopes] of f.credCalls) {
      if (mode === "application") assert.equal(scopes, undefined, "application credentials use the .default audience, never delegated scopes");
    }
  } finally { teardown(state); }
});

for (const upn of ["AdeleV@contoso.com", "alice_example.com#EXT#@tenant.onmicrosoft.com"]) test(`UPN ${upn} resolves once and pins preview, PATCH and rereads to the object ID`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [true, true, true, false] });
    const preview = await executeArgv(updateArgs(upn, "false"), fixture().overrides);
    assert.equal(preview.preview.url, MUTATION_URL);
    assert.equal(preview.preview.target, upn);
    const result = await executeArgv(updateArgs(upn, "false", ["--execute", "--confirm", upn]), f.overrides);
    assert.deepEqual(result.user, { id: userId, accountEnabled: false });
    assert.equal(f.mutRequests[0].url, MUTATION_URL);
    assert.equal(f.mutRequests[0].body, JSON.stringify({ accountEnabled: false }));
    assert.equal(f.readRequests.length, 4);
    assert.deepEqual(f.readRequests.map(request => new URL(request.url).pathname), [
      `/v1.0/users/${encodeURIComponent(upn)}`,
      `/v1.0/users/${userId}`,
      `/v1.0/users/${userId}`,
      `/v1.0/users/${userId}`,
    ]);
    assert.equal(journal(f.journalPath)[0].target, upn);
    assert.equal(journal(f.journalPath)[0].url, MUTATION_URL);
  } finally { teardown(state); }
});

for (const [profileName, desired, current] of [
  ["soc", true, false],
  ["soc", false, true],
  ["batch", true, false],
  ["batch", false, true],
]) test(`${profileName} update to ${desired} keeps the original identity after UPN reassignment`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir, profileName);
    const upn = "alice_example.com#EXT#@tenant.onmicrosoft.com";
    const replacementId = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
    const f = fixture({ reads: [current, current, current, desired] });
    const originalTransport = f.overrides.transport;
    f.overrides.transport = async request => {
      const response = await originalTransport(request);
      if (f.readRequests.length > 1 && new URL(request.url).pathname === `/v1.0/users/${encodeURIComponent(upn)}`) {
        return json(200, { id: replacementId, accountEnabled: f.mutRequests.length > 0 ? desired : current });
      }
      return response;
    };
    const result = await executeArgv([
      "entra", "user", "update", "--user", upn, "--account-enabled", String(desired),
      "--profile", profileName, "--execute", "--confirm", upn,
    ], f.overrides);
    assert.deepEqual(result.user, { id: userId, accountEnabled: desired });
    assert.equal(f.mutRequests.length, 1);
    assert.equal(f.mutRequests[0].url, MUTATION_URL);
    assert.deepEqual(f.readRequests.map(request => new URL(request.url).pathname), [
      `/v1.0/users/${encodeURIComponent(upn)}`,
      `/v1.0/users/${userId}`,
      `/v1.0/users/${userId}`,
      `/v1.0/users/${userId}`,
    ]);
  } finally { teardown(state); }
});

for (const user of ["..", "a/b", "a\\b", "a?b", "a%23b", "a%2fb", "a b"]) test(`unsafe account target ${user} is refused before credentials`, async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture();
    await assert.rejects(executeArgv(updateArgs(user, "false", ["--execute", "--confirm", user]), f.overrides), { code: "VALIDATION_ERROR" });
    assert.equal(f.readRequests.length, 0);
    assert.equal(f.mutRequests.length, 0);
    assert.equal(f.credCalls.length, 0);
  } finally { teardown(state); }
});

test("unreadable current state never counts as desired", async () => {
  const state = setupProfiles();
  try {
    enableWrites(state.dir);
    const f = fixture({ reads: [null] });
    const result = await executeArgv(updateArgs(userId, "false"), f.overrides);
    assert.equal(result.preview.current, null);
    assert.equal(result.preview.noop, false);
    assert.equal(f.mutRequests.length, 0);
  } finally { teardown(state); }
});

test("generated records name the account write", async () => {
  assert.ok(skillCommandTable().includes("`mg-axi entra user update` | native | write |"));
  const coverage = capabilityDocument();
  assert.ok(coverage.includes("## Named writes"));
  assert.ok(coverage.includes("`mg-axi entra user update` | `PATCH:/users/{user-id}` | WRITE-02 |"));
});

test("setup counts the write beside reads and local leaves", async () => {
  const state = setupProfiles();
  try {
    const output = setupView(new Profiles());
    assert.equal(output.capabilities.reads, 115);
    assert.equal(output.capabilities.writes, 5);    assert.equal(output.capabilities.local, 6);
    assert.ok(output.capabilities.implemented.includes("entra user update"));
  } finally { teardown(state); }
});
