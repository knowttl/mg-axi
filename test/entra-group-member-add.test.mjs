import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, test } from "node:test";
import { ApplicationAuth } from "../dist/app-auth.js";
import { DelegatedAuth } from "../dist/auth.js";
import { GraphSession } from "../dist/graph-session.js";
import { executeArgv } from "../dist/cli.js";
import { skillCommandTable } from "../dist/docs.js";
import { addGroupMember, GROUP_MEMBER_ADD_OPERATION, GROUP_MEMBER_ADD_SCOPES } from "../dist/entra-group-member-add.js";

// WRITE-01 acceptance through real interfaces: the module drives a real
// GraphSession for preview reads and a real mutation coordinator for the
// POST, with fixture transports and credentials only. No live instance,
// real credential or customer data exists.
const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const gid = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const uid = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";
const other = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";
const NOW = 1_700_000_000_000;
const TIME = new Date(NOW).toISOString();
const WRITE_SCOPE = "https://graph.microsoft.com/GroupMember.ReadWrite.All";
const READ_SCOPE = "https://graph.microsoft.com/GroupMember.Read.All";

const writes = { allowWrites: true, operations: [GROUP_MEMBER_ADD_OPERATION] };
const delegatedBase = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key } };
const enabledDelegated = { ...delegatedBase, writes };
const appBase = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "federated", key } };
const enabledApp = { ...appBase, writes };

const plainGroup = {
  id: gid,
  displayName: "Engineering",
  isAssignableToRole: false,
  securityEnabled: true,
  groupTypes: [],
  mailEnabled: false,
  membershipRule: null,
  membershipRuleProcessingState: null,
};
const m365Group = { ...plainGroup, securityEnabled: false, groupTypes: ["Unified"] };

const scratches = [];
afterEach(() => { while (scratches.length) rmSync(scratches.pop(), { recursive: true, force: true }); });

function scratch() {
  const dir = mkdtempSync(join(tmpdir(), "mg-write01-"));
  scratches.push(dir);
  return dir;
}

function json(status, body, headers = {}) {
  return { status, headers, body: JSON.stringify(body) };
}

function fixture({ group = plainGroup, user = json(200, { id: uid }), members = [{ id: other }], post } = {}) {
  const dir = scratch();
  const journalPath = join(dir, "writes.log");
  const getRequests = [];
  const postRequests = [];
  const credentialCalls = [];
  let userReads = 0;
  const credential = account => ({
    token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, ...account,
  });
  const delegated = new DelegatedAuth({
    storage: "session-only",
    login: async () => credential({ accountId: "synthetic-account" }),
    silent: async (...args) => { credentialCalls.push(["silent", ...args]); return credential({ accountId: "synthetic-account" }); },
  });
  const application = new ApplicationAuth({
    storage: "session-only",
    acquire: async (...args) => { credentialCalls.push(["acquire", ...args]); return credential({}); },
  });
  const get = async request => {
    getRequests.push(request);
    const url = new URL(request.url);
    if (url.pathname === `/v1.0/users/${uid}`) {
      userReads += 1;
      return typeof user === "function" ? user(request, userReads) : user;
    }
    if (url.pathname === `/v1.0/groups/${gid}`) {
      if (group === null) return json(404, { error: { code: "Request_ResourceNotFound", message: "no such group" } });
      return json(200, group);
    }
    if (url.pathname === `/v1.0/groups/${gid}/members`) return json(200, { value: members });
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  };
  const transport = async request => {
    postRequests.push(request);
    if (typeof post === "function") return post(request, postRequests.length);
    return post ?? { status: 204, headers: {}, body: "" };
  };
  const session = new GraphSession({ delegated, application, transport: get });
  const run = (flags, profile = enabledDelegated) => addGroupMember({
    session, profile, delegated, application, transport,
    flags: { group: gid, user: uid, ...flags },
    profileName: "soc", help: "mg-axi entra group member add --help",
    journalPath, clock: () => NOW,
  });
  return { dir, journalPath, getRequests, postRequests, credentialCalls, delegated, application, session, getSend: get, postSend: transport, run };
}

function journal(path) {
  return readFileSync(path, "utf8").split("\n").filter(Boolean).map(line => JSON.parse(line));
}

async function withEnv(entries, action) {
  const previous = {};
  for (const [key, value] of Object.entries(entries)) {
    previous[key] = process.env[key];
    if (value === undefined) delete process.env[key];
    else process.env[key] = value;
  }
  try {
    return await action();
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
}

test("dry-run previews the exact $ref target without sending or journaling", async () => {
  const f = fixture();
  const result = await f.run({});
  assert.equal(result.membership.status, "preview");
  assert.equal(result.membership.group, `https://graph.microsoft.com/v1.0/groups/${gid}`);
  assert.equal(result.membership.user, `https://graph.microsoft.com/v1.0/directoryObjects/${uid}`);
  assert.equal(result.preview.operation, GROUP_MEMBER_ADD_OPERATION);
  assert.equal(result.preview.method, "POST");
  assert.equal(result.preview.url, `https://graph.microsoft.com/v1.0/groups/${gid}/members/$ref`);
  assert.equal(result.preview.target, gid);
  const change = JSON.parse(result.preview.proposedChange);
  assert.deepEqual(change, { "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${uid}` });
  assert.ok(!JSON.stringify(result).includes("/users/"), "the $ref target is directoryObjects, never users");
  assert.ok(result.help.some(hint => hint.includes("--execute --confirm")));
  assert.equal(f.postRequests.length, 0);
  assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
});

test("confirmed execution posts the directoryObjects $ref once and journals intent plus SUCCESS", async () => {
  const f = fixture();
  const result = await f.run({ execute: true, confirm: gid });
  assert.deepEqual(result.membership, {
    group: `https://graph.microsoft.com/v1.0/groups/${gid}`,
    user: `https://graph.microsoft.com/v1.0/directoryObjects/${uid}`,
    status: "added",
  });
  assert.match(result.auditId, /^[0-9a-f-]{36}$/);
  assert.ok(result.help.some(hint => hint.includes("entra group member list")));
  assert.equal(f.postRequests.length, 1);
  const sent = f.postRequests[0];
  assert.equal(sent.method, "POST");
  assert.equal(sent.url, `https://graph.microsoft.com/v1.0/groups/${gid}/members/$ref`);
  assert.equal(sent.body, JSON.stringify({ "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${uid}` }));
  assert.equal(sent.headers["Content-Type"], "application/json");
  assert.equal(sent.headers.Authorization, "Bearer opaque-fixture-secret");
  const userReads = f.getRequests.map(request => new URL(request.url)).filter(url => url.pathname === `/v1.0/users/${uid}`);
  assert.equal(userReads.length, 2);
  assert.deepEqual(userReads.map(url => url.searchParams.get("$select")), ["id", "id"]);
  const records = journal(f.journalPath);
  assert.equal(records.length, 2);
  assert.equal(records[0].kind, "intent");
  assert.equal(records[0].operation, GROUP_MEMBER_ADD_OPERATION);
  assert.equal(records[0].target, gid);
  assert.equal(records[0].time, TIME);
  assert.equal(records[1].kind, "outcome");
  assert.equal(records[1].outcome, "SUCCESS");
  assert.equal(records[1].httpStatus, 204);
  assert.ok(!readFileSync(f.journalPath, "utf8").includes("@odata.id"), "journal carries metadata, never the payload");
});

test("delegated execution requests only the documented write scope for the mutation", async () => {
  const f = fixture();
  await f.run({ execute: true, confirm: gid });
  const silent = f.credentialCalls.filter(call => call[0] === "silent");
  const scopeSets = silent.map(call => call[2]);
  assert.ok(scopeSets.some(scopes => scopes.length === 1 && scopes[0] === WRITE_SCOPE), "mutation credential uses GroupMember.ReadWrite.All only");
  assert.ok(scopeSets.some(scopes => scopes.includes(READ_SCOPE)), "preview reads use the READ-02 read scope");
  assert.ok(scopeSets.some(scopes => scopes.length === 1 && scopes[0] === "https://graph.microsoft.com/User.ReadBasic.All"));
  assert.deepEqual(GROUP_MEMBER_ADD_SCOPES, [WRITE_SCOPE]);
});

for (const profile of [enabledDelegated, enabledApp]) {
  for (const [name, response] of [
    ["non-user or missing object", json(404, { error: { code: "Request_ResourceNotFound", message: "not a user" } })],
    ["inaccessible user", json(403, { error: { code: "Authorization_RequestDenied", message: "denied" } })],
    ["null body", json(200, null)],
    ["array body", json(200, [{ id: uid }])],
    ["missing identity", json(200, {})],
    ["malformed identity", json(200, { id: 123 })],
    ["wrong identity", json(200, { id: other })],
    ["non-user type", json(200, { id: uid, "@odata.type": "#microsoft.graph.group" })],
  ]) {
    for (const [path, flags, members] of [
      ["preview", {}, []],
      ["execution", { execute: true, confirm: gid }, []],
      ["no-op", { execute: true, confirm: gid }, [{ id: uid }]],
    ]) {
      test(`${profile.mode} ${path} refuses ${name} before accepting membership`, async () => {
        const f = fixture({ user: response, members });
        await assert.rejects(f.run(flags, profile), { code: "GRAPH_ERROR" });
        assert.equal(f.postRequests.length, 0);
        assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
      });
    }
    test(`${profile.mode} fresh read refuses ${name} before sending`, async () => {
      const f = fixture({ user: (_request, count) => count === 1 ? json(200, { id: uid.toUpperCase() }) : response });
      await assert.rejects(f.run({ execute: true, confirm: gid }, profile), { code: "GRAPH_ERROR" });
      assert.equal(f.postRequests.length, 0);
      assert.equal(journal(f.journalPath)[1].outcome, "NOT_SENT");
    });
  }
}

test("already-member preview is a no-op with zero sends and no journal", async () => {
  const f = fixture({ members: [{ id: other }, { id: uid.toUpperCase() }] });
  const result = await f.run({ execute: true, confirm: gid });
  assert.equal(result.membership.status, "already-member");
  assert.equal(result.noop, true);
  assert.equal(f.postRequests.length, 0);
  assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
});

test("duplicate-reference 400 is a no-op with a NOOP outcome, not a failure", async () => {
  const f = fixture({ post: json(400, { error: { code: "Request_BadRequest", message: "One or more added object references already exist for the following modified properties: 'members'." } }) });
  const result = await f.run({ execute: true, confirm: gid });
  assert.equal(result.membership.status, "already-member");
  assert.equal(result.noop, true);
  assert.match(result.auditId, /^[0-9a-f-]{36}$/);
  assert.match(result.note, /already exists/);
  assert.equal(f.postRequests.length, 1);
  const records = journal(f.journalPath);
  assert.equal(records.length, 2);
  assert.equal(records[1].outcome, "NOOP");
  assert.equal(records[1].httpStatus, 400);
});

test("non-duplicate 400 stays a failure with replication guidance", async () => {
  const f = fixture({ post: json(400, { error: { code: "Request_BadRequest", message: "The source resource object or one of the objects being referenced don't exist." } }) });
  await assert.rejects(f.run({ execute: true, confirm: gid }), error => {
    assert.equal(error.code, "GRAPH_ERROR");
    return error.suggestions.some(hint => /recently created|replication/i.test(hint)) && error.suggestions.some(hint => /never replay/.test(hint));
  });
  assert.equal(journal(f.journalPath)[1].outcome, "FAILED");
});

test("403 denial names the write permission and delegated role requirement", async () => {
  const f = fixture({ post: json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }) });
  await assert.rejects(f.run({ execute: true, confirm: gid }), error => {
    assert.equal(error.code, "GRAPH_ERROR");
    return error.suggestions.some(hint => /GroupMember\.ReadWrite\.All/.test(hint))
      && error.suggestions.some(hint => /Groups Administrator/.test(hint))
      && error.suggestions.some(hint => /never proves/.test(hint));
  });
  assert.equal(journal(f.journalPath)[1].outcome, "FAILED");
});

test("404 reports the group or user as not found with identifier guidance", async () => {
  const f = fixture({ post: json(404, { error: { code: "Request_ResourceNotFound", message: "gone" } }) });
  await assert.rejects(f.run({ execute: true, confirm: gid }), error => {
    assert.equal(error.code, "GRAPH_ERROR");
    return /not found or inaccessible/.test(error.message) && error.suggestions.some(hint => /never a UPN/.test(hint));
  });
});

test("role-assignable groups are refused before sending with role-management guidance", async () => {
  const f = fixture({ group: { ...plainGroup, isAssignableToRole: true } });
  await assert.rejects(f.run({ execute: true, confirm: gid }), error => {
    assert.equal(error.code, "OPERATION_BLOCKED");
    return /role-assignable/i.test(error.message)
      && error.suggestions.some(hint => /RoleManagement\.ReadWrite\.Directory/.test(hint));
  });
  assert.equal(f.postRequests.length, 0);
  assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
});

test("dynamic-membership groups are refused before sending", async () => {
  const f = fixture({ group: { ...plainGroup, membershipRule: "user.department -eq \"Sales\"" } });
  await assert.rejects(f.run({ execute: true, confirm: gid }), error => {
    assert.equal(error.code, "OPERATION_BLOCKED");
    return /dynamic/i.test(error.message);
  });
  assert.equal(f.postRequests.length, 0);
});

test("distribution groups are refused as unmanageable before sending", async () => {
  const f = fixture({ group: { ...plainGroup, securityEnabled: false, groupTypes: [], mailEnabled: true } });
  await assert.rejects(f.run({ execute: true, confirm: gid }), error => {
    assert.equal(error.code, "OPERATION_BLOCKED");
    return /security and Microsoft 365/.test(error.message);
  });
  assert.equal(f.postRequests.length, 0);
});

test("Microsoft 365 groups are supported", async () => {
  const f = fixture({ group: m365Group });
  const result = await f.run({ execute: true, confirm: gid });
  assert.equal(result.membership.status, "added");
  assert.equal(f.postRequests.length, 1);
});

for (const [groupType, group] of [["security", plainGroup], ["Microsoft 365", m365Group]]) {
  for (const profile of [enabledDelegated, enabledApp]) {
    for (const [status, flags, members, sends] of [
      ["preview", {}, [{ id: other }], 0],
      ["added", { execute: true, confirm: gid }, [{ id: other }], 1],
      ["already-member", { execute: true, confirm: gid }, [{ id: uid }], 0],
    ]) {
      test(`null role-assignable state supports ${status} for ${profile.mode} ${groupType} groups`, async () => {
        const f = fixture({ group: { ...group, isAssignableToRole: null }, members });
        const result = await f.run(flags, profile);
        assert.equal(result.membership.status, status);
        assert.equal(f.postRequests.length, sends);
      });
    }
  }
}

for (const state of [undefined, "false", 0, {}, []]) {
  test(`unknown role-assignable state ${JSON.stringify(state)} fails closed before sending`, async () => {
    const f = fixture({ group: { ...plainGroup, isAssignableToRole: state } });
    await assert.rejects(f.run({ execute: true, confirm: gid }), { code: "OPERATION_BLOCKED" });
    assert.equal(f.postRequests.length, 0);
    assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
  });
}

test("missing group reads as not found without sending", async () => {
  const f = fixture({ group: null });
  await assert.rejects(f.run({ execute: true, confirm: gid }), error => {
    assert.equal(error.code, "GRAPH_ERROR");
    return /not found or inaccessible/.test(error.message);
  });
  assert.equal(f.postRequests.length, 0);
});

test("incomplete member window proceeds to the POST instead of claiming absence", async () => {
  const dir = scratch();
  const journalPath = join(dir, "writes.log");
  const postRequests = [];
  let pages = 0;
  const credential = account => ({ token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, ...account });
  const delegated = new DelegatedAuth({
    storage: "session-only",
    login: async () => credential({ accountId: "synthetic-account" }),
    silent: async () => credential({ accountId: "synthetic-account" }),
  });
  const application = new ApplicationAuth({ storage: "session-only", acquire: async () => credential({}) });
  const get = async request => {
    const url = new URL(request.url);
    if (url.pathname === `/v1.0/groups/${gid}`) return json(200, plainGroup);
    if (url.pathname === `/v1.0/users/${uid}`) return json(200, { id: uid });
    pages += 1;
    return json(200, { value: [{ id: other }], "@odata.nextLink": `https://graph.microsoft.com/v1.0/groups/${gid}/members?$skiptoken=${pages}` });
  };
  const session = new GraphSession({ delegated, application, transport: get });
  const result = await addGroupMember({
    session, profile: enabledDelegated, delegated, application,
    transport: async request => { postRequests.push(request); return { status: 204, headers: {}, body: "" }; },
    flags: { group: gid, user: uid, execute: true, confirm: gid },
    profileName: "soc", help: "mg-axi entra group member add --help", journalPath, clock: () => NOW,
  });
  assert.equal(result.membership.status, "added");
  assert.equal(postRequests.length, 1);
});

for (const confirm of [undefined, "wrong-target"]) {
  test(`confirm ${JSON.stringify(confirm)} is refused with zero sends and no journal`, async () => {
    const f = fixture();
    await assert.rejects(f.run({ execute: true, ...(confirm === undefined ? {} : { confirm }) }),
      { code: confirm === undefined ? "CONFIRM_REQUIRED" : "CONFIRM_MISMATCH" });
    assert.equal(f.postRequests.length, 0);
    assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
  });
}

test("beta api-version stays blocked before any credential or HTTP", async () => {
  const f = fixture();
  await assert.rejects(f.run({ "api-version": "beta", execute: true, confirm: gid }), { code: "VALIDATION_ERROR" });
  assert.equal(f.getRequests.length, 0);
  assert.equal(f.postRequests.length, 0);
  assert.equal(f.credentialCalls.length, 0);
});

for (const [name, flags] of [
  ["non-GUID group", { group: "Engineering", user: uid }],
  ["non-GUID user", { group: gid, user: "adele@contoso.com" }],
  ["empty group", { group: "", user: uid }],
]) {
  test(`invalid identifier (${name}) fails before any credential or HTTP`, async () => {
    const f = fixture();
    await assert.rejects(f.run(flags), { code: "VALIDATION_ERROR" });
    assert.equal(f.getRequests.length, 0);
    assert.equal(f.postRequests.length, 0);
    assert.equal(f.credentialCalls.length, 0);
  });
}

test("operation outside the configured scope is refused with zero mutation sends", async () => {
  const f = fixture();
  await assert.rejects(
    f.run({ execute: true, confirm: gid }, { ...enabledDelegated, writes: { allowWrites: true, operations: ["mg.other.write"] } }),
    { code: "OPERATION_NOT_WRITABLE" },
  );
  assert.equal(f.postRequests.length, 0);
});

test("missing hand-enabled writes key stays read-only with zero mutation sends", async () => {
  const f = fixture();
  await assert.rejects(f.run({ execute: true, confirm: gid }, delegatedBase), { code: "WRITES_DISABLED" });
  assert.equal(f.postRequests.length, 0);
  assert.throws(() => readFileSync(f.journalPath, "utf8"), /ENOENT/);
});

test("forced read-only overrides the hand-enabled profile with zero mutation sends", async () => {
  const f = fixture();
  await withEnv({ MG_AXI_READ_ONLY: "1" }, async () => {
    await assert.rejects(f.run({ execute: true, confirm: gid }), { code: "WRITES_DISABLED" });
  });
  assert.equal(f.postRequests.length, 0);
});

test("transport failure reports unknown with no-replay guidance and OUTCOME_UNKNOWN", async () => {
  const f = fixture({ post: async () => { throw new Error("socket hang up"); } });
  const result = await f.run({ execute: true, confirm: gid });
  assert.equal(result.membership.status, "unknown");
  assert.match(result.auditId, /^[0-9a-f-]{36}$/);
  assert.match(result.guidance, /never replay this intent/);
  assert.ok(result.help.some(hint => hint.includes("entra group member list")));
  assert.equal(journal(f.journalPath)[1].outcome, "OUTCOME_UNKNOWN");
});

test("server error reports unknown with the HTTP status", async () => {
  const f = fixture({ post: json(503, { error: { code: "ServiceUnavailable", message: "busy" } }) });
  const result = await f.run({ execute: true, confirm: gid });
  assert.equal(result.membership.status, "unknown");
  assert.equal(result.httpStatus, 503);
  assert.equal(journal(f.journalPath)[1].outcome, "OUTCOME_UNKNOWN");
});

test("application profile executes without caller scopes", async () => {
  const f = fixture();
  const result = await f.run({ execute: true, confirm: gid }, enabledApp);
  assert.equal(result.membership.status, "added");
  assert.deepEqual(f.credentialCalls[0][0], "acquire");
  assert.equal(journal(f.journalPath)[1].outcome, "SUCCESS");
});

test("application profile rejects caller scopes", async () => {
  const f = fixture();
  await assert.rejects(f.run({ execute: true, confirm: gid, scopes: WRITE_SCOPE }, enabledApp), { code: "VALIDATION_ERROR" });
  assert.equal(f.postRequests.length, 0);
});

test("member-add help names confirmation, scope and refusal without touching the network", async () => {
  const help = await executeArgv(["entra", "group", "member", "add", "--help"], {
    transport: async () => { throw new Error("transport must stay untouched"); },
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => { throw new Error("no login"); }, silent: async () => { throw new Error("no credential"); } }),
    application: new ApplicationAuth({ storage: "session-only", acquire: async () => { throw new Error("no credential"); } }),
  });
  assert.match(help, /--confirm/);
  assert.match(help, /GroupMember\.ReadWrite\.All/);
  assert.match(help, /Role-assignable/);
});

test("the skill table marks the membership write as a write", () => {
  assert.ok(skillCommandTable().split("\n").some(row => row.includes("`mg-axi entra group member add` | native | write |")));
});

function cliConfig(dir, profile) {
  writeFileSync(join(dir, "config.json"), JSON.stringify({ version: 1, profiles: { soc: profile } }));
}

test("CLI dry-run previews through fixtures with no send and no journal", async () => {
  const dir = scratch();
  cliConfig(dir, enabledDelegated);
  const logPath = join(dir, "cli-writes.log");
  const f = fixture();
  const output = await withEnv({ MG_AXI_CONFIG: join(dir, "config.json"), MG_AXI_WRITE_LOG: logPath }, async () => executeArgv(
    ["entra", "group", "member", "add", "--group", gid, "--user", uid, "--profile", "soc"],
    { transport: f.getSend, mutationTransport: f.postSend, delegated: f.delegated, application: f.application },
  ));
  assert.equal(output.membership.status, "preview");
  assert.equal(output.preview.url, `https://graph.microsoft.com/v1.0/groups/${gid}/members/$ref`);
  assert.equal(f.postRequests.length, 0);
  assert.throws(() => readFileSync(logPath, "utf8"), /ENOENT/);
});

test("CLI confirmed execution sends through fixtures and journals intent plus SUCCESS", async () => {
  const dir = scratch();
  cliConfig(dir, enabledDelegated);
  const logPath = join(dir, "cli-writes.log");
  const f = fixture();
  const output = await withEnv({ MG_AXI_CONFIG: join(dir, "config.json"), MG_AXI_WRITE_LOG: logPath }, async () => executeArgv(
    ["entra", "group", "member", "add", "--group", gid, "--user", uid, "--profile", "soc", "--execute", "--confirm", gid],
    { transport: f.getSend, mutationTransport: f.postSend, delegated: f.delegated, application: f.application },
  ));
  assert.equal(output.membership.status, "added");
  assert.equal(f.postRequests.length, 1);
  const records = journal(logPath);
  assert.equal(records.length, 2);
  assert.equal(records[0].kind, "intent");
  assert.equal(records[1].outcome, "SUCCESS");
});
