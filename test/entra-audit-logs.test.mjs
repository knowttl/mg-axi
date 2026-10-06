import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import { DelegatedAuth } from "../dist/auth.js";
import { MAX_CURSOR_BYTES } from "../dist/graph-session.js";
import { runApiGet } from "../dist/api.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const auditScopes = ["https://graph.microsoft.com/AuditLog.Read.All"];
const SINCE = "2026-09-01T00:00:00Z";
const UNTIL = "2026-09-08T00:00:00Z";

const s1 = {
  id: "11111111-1111-4111-8111-111111111111",
  createdDateTime: "2026-09-10T08:00:00Z",
  userId: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  userPrincipalName: "AdeleV@contoso.com",
  userDisplayName: "Adele Vance",
  appDisplayName: "Graph explorer",
  appId: "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
  ipAddress: "131.107.159.37",
  location: { city: "Redmond", countryOrRegion: "US" },
  status: { errorCode: 0, failureReason: null, additionalDetails: null },
  conditionalAccessStatus: "notApplied",
  riskDetail: "none",
  riskLevelAggregated: "none",
  riskLevelDuringSignIn: "none",
  riskState: "none",
  resourceDisplayName: "Microsoft Graph",
  resourceId: "00000003-0000-0000-c000-000000000000",
  clientAppUsed: "Browser",
};
const s2 = {
  id: "22222222-2222-4222-8222-222222222222",
  createdDateTime: "2026-09-11T09:30:00Z",
  userId: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  userPrincipalName: "AlexW@contoso.com",
  userDisplayName: "Alex Wilber",
  appDisplayName: "Azure Portal",
  appId: "c44b4083-3bb0-49c1-b47d-974e53cbdf3c",
  ipAddress: "131.107.159.38",
  location: { city: "Seattle", countryOrRegion: "US" },
  status: { errorCode: 50126, failureReason: "Invalid username or password", additionalDetails: null },
  conditionalAccessStatus: "notApplied",
  appliedConditionalAccessPolicies: [{ id: "de7e60eb-ed89-4d73-8205-2227def6b7c9", displayName: "Medium signin risk block", result: "notEnabled" }],
  riskDetail: "none",
  riskLevelAggregated: "none",
  riskLevelDuringSignIn: "none",
  riskState: "none",
  resourceDisplayName: "Microsoft Graph",
  resourceId: "00000003-0000-0000-c000-000000000000",
  clientAppUsed: "Browser",
};
const s3 = {
  id: "33333333-3333-4333-8333-333333333333",
  createdDateTime: "2026-09-12T10:00:00Z",
  userId: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  userPrincipalName: "GradyA@contoso.com",
  userDisplayName: "Grady Archie",
  appDisplayName: "Graph explorer",
  appId: "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
  ipAddress: null,
  location: null,
  status: { errorCode: 0, failureReason: null, additionalDetails: null },
  conditionalAccessStatus: null,
  riskDetail: null,
  riskLevelAggregated: null,
  riskLevelDuringSignIn: null,
  riskState: null,
  resourceDisplayName: "Microsoft Graph",
  resourceId: "00000003-0000-0000-c000-000000000000",
  clientAppUsed: "Mobile Apps and Desktop clients",
};
const signIns = [s1, s2, s3];

const a1 = {
  id: "SSGM_b662f17a-4e4d-4e1c-9248-cdec180024b_MCDC4_88453290",
  activityDateTime: "2026-09-10T21:20:02.7215374Z",
  activityDisplayName: "Add member to group",
  category: "GroupManagement",
  loggedByService: "Core Directory",
  operationType: "Assign",
  result: "success",
  resultReason: "Successfully added member to group",
  correlationId: "da159bfb-54fa-4092-8a38-6e1fa7870e30",
  initiatedBy: { user: { id: "728309ae-1a37-4937-9afe-e35d964db09b", displayName: "Audry Oliver", userPrincipalName: "bob@wingtiptoysonline.com" }, app: null },
  targetResources: [{ id: "ef7e527d-6c92-4234-8c6d-cf6fdfb57f95", displayName: "Example.com", type: "Group", modifiedProperties: [] }],
};
const a2 = {
  id: "SSGM_c773f28b-5f5e-5f2d-a359-dfed291135c_NDDD5_99564301",
  activityDateTime: "2026-09-11T07:05:11.1032219Z",
  activityDisplayName: "Update user",
  category: "UserManagement",
  loggedByService: "Core Directory",
  operationType: "Update",
  result: "failure",
  resultReason: "Insufficient privileges",
  correlationId: "eb260acc-65fb-5033-9f49-f70ba8981f41",
  initiatedBy: { user: null, app: { displayName: "Sync service", servicePrincipalId: "55555555-5555-4555-8555-555555555555" } },
  targetResources: [],
};
const audits = [a1, a2];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-05-"));
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(dir, "config.json");
  const profiles = new Profiles();
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  return { dir, previous };
}

function teardownProfiles(state) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function credentialService(mode, calls) {
  return {
    credential: async (...args) => {
      calls.push(args);
      return {
        token: `opaque-fixture-${mode}-token`,
        expiresAt: Date.now() + 3_600_000,
        tenantId: tenant,
        clientId: client,
        ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
      };
    },
  };
}

function transport(handler) {
  const requests = [];
  const send = async request => {
    requests.push(request);
    const response = await handler(request);
    return { headers: {}, body: "", ...response };
  };
  return { requests, send };
}

function json(status, body, headers = {}) {
  return { status, headers, body: JSON.stringify(body) };
}

function auditTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/auditLogs/signIns") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [s3] });
      return json(200, {
        value: [s1, s2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/auditLogs/signIns?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/auditLogs/directoryAudits") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [a2] });
      return json(200, {
        value: [a1],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?%24skiptoken=page2",
      });
    }
    const signIn = signIns.find(row => path === `/v1.0/auditLogs/signIns/${row.id}`);
    if (signIn) return json(200, signIn);
    const audit = audits.find(row => path === `/v1.0/auditLogs/directoryAudits/${row.id}`);
    if (audit) return json(200, audit);
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? auditTransport();
  const credential = credentialService(mode, calls);
  return {
    requests: fixture.requests,
    calls,
    overrides: {
      transport: fixture.send,
      delegated: mode === "delegated" ? credential : credentialService("delegated", []),
      application: mode === "application" ? credential : credentialService("application", []),
    },
  };
}

function hintArgv(hint) {
  const argv = [];
  let word = "";
  let quote = null;
  let escaped = false;
  let started = false;
  const push = () => { if (started) { argv.push(word); word = ""; started = false; } };
  for (const char of hint.slice("mg-axi ".length)) {
    if (escaped) { word += char; escaped = false; }
    else if (quote === "'") { if (char === "'") quote = null; else word += char; }
    else if (quote === '"') {
      if (char === '"') quote = null;
      else if (char === "\\") escaped = true;
      else word += char;
    }
    else if (char === "\\") escaped = true;
    else if (char === "'" || char === '"') { quote = char; started = true; }
    else if (/\s/.test(char)) push();
    else { word += char; started = true; }
  }
  push();
  return argv;
}

function runReadCli(args, state, mode, denied = false, input) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-audit-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["pipe", "pipe", "pipe"], input, timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, signIns, audits, denied }),
    },
  });
}

for (const [noun, key, field, plural] of [
  ["sign-in", "signIns", "userPrincipalName", "sign-ins"],
  ["directory-audit", "directoryAudits", "activityDisplayName", "directory audits"],
]) {
  for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
    test(`${mode} executable resumes large ${noun} cursors and recovers full text through stdin`, async () => {
      const state = setupProfiles();
      try {
        const rows = Array.from({ length: 1000 }, (_, i) => ({ id: String(i), [field]: "x".repeat(600) }));
        const fixture = transport(() => json(200, { value: rows }));
        const { overrides } = overridesFor(mode, fixture);
        const first = await executeArgv(["entra", noun, "list", "--profile", profile,
          "--since", SINCE, "--select", `id,${field}`], overrides);
        assert.equal(first[key].length, 100);
        assert.equal(first.complete, false);
        assert.ok(Buffer.byteLength(first.cursor) > 128 * 1024);
        assert.match(first.help.join("\n"), /--cursor -.*stdin/);
        const second = runReadCli(["entra", noun, "list", "--profile", profile, "--cursor", "-"],
          state, mode, true, `${first.cursor}\n`);
        assert.equal(second.error, undefined);
        assert.equal(second.status, 0, second.stdout);
        assert.equal(second.stderr, "");
        const partial = decode(second.stdout);
        assert.deepEqual(partial[key].map(row => row.id), rows.slice(100, 200).map(row => row.id));
        assert.match(partial[key][0][field], /truncated/);
        assert.ok(!partial.help[0].includes(first.cursor));
        assert.ok(partial.help[0].includes("--cursor -"));
        assert.ok(partial.help.some(hint => hint.includes("original input cursor on stdin")));
        const recovered = runReadCli(hintArgv(partial.help[0]), state, mode, true, first.cursor);
        assert.equal(recovered.status, 0, recovered.stdout);
        assert.deepEqual(decode(recovered.stdout)[key], rows.slice(100, 200));
        const last = runReadCli(["entra", noun, "list", "--profile", profile, "--cursor", "-", "--all", "--full"],
          state, mode, true, partial.cursor);
        assert.equal(last.status, 0, last.stdout);
        const complete = decode(last.stdout);
        assert.deepEqual(complete[key], rows.slice(200));
        assert.deepEqual(complete.count, `800 ${plural}`);
        assert.equal(complete.total, null);
        assert.equal(complete.complete, true);
      } finally { teardownProfiles(state); }
    });
  }

  for (const [input, expectedError] of [
    ["", /needs the opaque cursor/],
    ["invalid", /Invalid collection cursor/],
    ["a".repeat(MAX_CURSOR_BYTES + 1), /cursor exceeds 16000000 bytes/],
  ]) {
    test(`${noun} executable rejects ${input.length > MAX_CURSOR_BYTES ? "oversized" : input || "empty"} stdin cursors`, () => {
      const state = setupProfiles();
      try {
        const result = runReadCli(["entra", noun, "list", "--profile", "soc", "--cursor", "-"],
          state, "delegated", true, input);
        assert.equal(result.status, 2, result.stdout);
        assert.equal(result.stderr, "");
        const output = decode(result.stdout);
        assert.equal(output.code, "VALIDATION_ERROR");
        assert.match(output.error, expectedError);
      } finally { teardownProfiles(state); }
    });
  }

  test(`${noun} stdin cursor rejects another collection's binding`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor("delegated");
      const otherNoun = noun === "sign-in" ? "directory-audit" : "sign-in";
      const first = await executeArgv(["entra", otherNoun, "list", "--profile", "soc", "--since", SINCE, "--limit", "1"], overrides);
      const result = runReadCli(["entra", noun, "list", "--profile", "soc", "--cursor", "-"],
        state, "delegated", true, first.cursor);
      assert.equal(result.status, 2, result.stdout);
      assert.equal(decode(result.stdout).code, "VALIDATION_ERROR");
    } finally { teardownProfiles(state); }
  });
}

for (const [noun, path, dateField, key, filterField] of [
  ["sign-in", "/auditLogs/signIns", "createdDateTime", "signIns", "userDisplayName"],
  ["directory-audit", "/auditLogs/directoryAudits", "activityDateTime", "directoryAudits", "activityDisplayName"],
]) {
  for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
    for (const filter of [undefined, "id eq 'example'", `${dateField} le 2026-09-08T00:00:00.000Z`,
      `${dateField} ge 2026-09-01T00:00:00.000Z or id eq 'example'`]) {
      test(`${mode} ${noun} rejects raw cursors without a mandatory lower time bound: ${filter ?? "no filter"}`, async () => {
        const state = setupProfiles();
        try {
          const { requests, calls, overrides } = overridesFor(mode);
          const raw = await runApiGet({ path, apiVersion: "v1.0", profile: new Profiles().resolve(profile).profile,
            scopes: mode === "delegated" ? auditScopes[0] : undefined, limit: 1,
            odata: `$select=id${filter === undefined ? "" : `&$filter=${filter}`}` }, overrides);
          assert.equal(raw.complete, false);
          const requestCount = requests.length;
          const credentialCount = calls.length;
          await assert.rejects(executeArgv(["entra", noun, "list", "--profile", profile, "--cursor", raw.cursor], overrides), error => {
            assert.equal(error.code, "VALIDATION_ERROR");
            assert.match(error.message, /bounded in time/);
            assert.ok(error.suggestions.some(hint => hint.includes("Start a new query with --since")));
            return true;
          });
          await assert.rejects(executeArgv(["entra", noun, "list", "--profile", profile, "--cursor", raw.cursor,
            "--since", SINCE], overrides), { code: "VALIDATION_ERROR" });
          assert.equal(requests.length, requestCount);
          assert.equal(calls.length, credentialCount);
          const executable = runReadCli(["entra", noun, "list", "--profile", profile, "--cursor", "-"],
            state, mode, true, raw.cursor);
          assert.equal(executable.status, 2, executable.stdout);
          const output = decode(executable.stdout);
          assert.equal(output.code, "VALIDATION_ERROR");
          assert.ok(output.help.some(hint => hint.includes("Start a new query with --since")));
        } finally { teardownProfiles(state); }
      });
    }
  }

  test(`${noun} resumes precise bounds with grouped filters and quoted parentheses`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor("delegated");
      const first = await executeArgv(["entra", noun, "list", "--profile", "soc", "--limit", "1",
        "--since", "2026-09-10T21:20:02.7215373Z", "--until", "2026-09-10T21:20:02.7215374Z",
        "--filter", `(${filterField} eq 'O''Brien (west)' or ${filterField} eq 'other')`], overrides);
      const resumed = await executeArgv(["entra", noun, "list", "--profile", "soc", "--cursor", first.cursor], overrides);
      const expected = noun === "sign-in" ? signIns : audits;
      assert.deepEqual(resumed[key].map(row => row.id), expected.slice(1).map(row => row.id));
      assert.equal(resumed.complete, true);
    } finally { teardownProfiles(state); }
  });

  for (const filter of [") or id eq 'example' or (", "(id eq 'example'", "id eq 'unterminated"]) {
    test(`${noun} rejects filters that escape time-bound grouping: ${filter}`, async () => {
      const state = setupProfiles();
      try {
        const { requests, calls, overrides } = overridesFor("delegated");
        await assert.rejects(executeArgv(["entra", noun, "list", "--profile", "soc", "--since", SINCE,
          "--filter", filter], overrides), { code: "VALIDATION_ERROR" });
        assert.deepEqual(requests, []);
        assert.deepEqual(calls, []);
      } finally { teardownProfiles(state); }
    });
  }
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} executable lists bounded sign-ins with compact rows`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "sign-in", "list", "--profile", profile, "--since", SINCE], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.signIns, [
        { id: s1.id, createdDateTime: s1.createdDateTime, userPrincipalName: s1.userPrincipalName, appDisplayName: s1.appDisplayName },
        { id: s2.id, createdDateTime: s2.createdDateTime, userPrincipalName: s2.userPrincipalName, appDisplayName: s2.appDisplayName },
        { id: s3.id, createdDateTime: s3.createdDateTime, userPrincipalName: s3.userPrincipalName, appDisplayName: s3.appDisplayName },
      ]);
      assert.deepEqual(output.count, "3 sign-ins");
      assert.equal(output.total, null);
      assert.equal(output.complete, true);
      assert.ok(output.help.some(hint => hint.includes("entra sign-in show --id <sign-in-id>")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable shows one sign-in with unavailable CA detail`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "sign-in", "show", "--id", s1.id, "--profile", profile], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.signIn.id, s1.id);
      assert.equal(output.signIn.conditionalAccessStatus, "notApplied");
      assert.match(output.signIn.appliedConditionalAccessPolicies, /^unavailable: /);
      assert.match(output.signIn.appliedConditionalAccessPolicies, /Conditional Access Administrator/);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable lists bounded directory audits`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "directory-audit", "list", "--profile", profile, "--since", SINCE], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.directoryAudits, [
        { id: a1.id, activityDateTime: a1.activityDateTime, activityDisplayName: a1.activityDisplayName, result: a1.result },
        { id: a2.id, activityDateTime: a2.activityDateTime, activityDisplayName: a2.activityDisplayName, result: a2.result },
      ]);
      assert.deepEqual(output.count, "2 directory audits");
      assert.equal(output.total, null);
      assert.equal(output.complete, true);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied log list is an operational error on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "sign-in", "list", "--profile", profile, "--since", SINCE], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.ok(output.help.some(hint => hint.includes("AuditLog.Read.All")));
      assert.equal(output.signIns, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });
}

test("delegated sign-in list composes since, until and filter into one bounded $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "sign-in", "list", "--profile", "soc",
      "--since", SINCE, "--until", UNTIL, "--filter", "status/errorCode ne 0", "--all"], overrides);
    assert.equal(requests.length, 2);
    const params = new URL(requests[0].url).searchParams;
    assert.equal(params.get("$filter"),
      "createdDateTime ge 2026-09-01T00:00:00.000Z and createdDateTime le 2026-09-08T00:00:00.000Z and (status/errorCode ne 0)");
    assert.equal(params.get("$select"), "id,createdDateTime,userPrincipalName,appDisplayName");
    assert.deepEqual(result.signIns.map(row => row.id), [s1.id, s2.id, s3.id]);
    assert.deepEqual(result.count, "3 sign-ins");
    assert.equal(result.total, null);
    assert.equal(result.complete, true);
    assert.ok(!JSON.stringify(result).includes("opaque-fixture-delegated-token"));
  } finally {
    teardownProfiles(state);
  }
});

test("directory-audit list bounds activityDateTime and returns compact rows", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "directory-audit", "list", "--profile", "soc", "--since", SINCE, "--all"], overrides);
    const params = new URL(requests[0].url).searchParams;
    assert.equal(params.get("$filter"), "activityDateTime ge 2026-09-01T00:00:00.000Z");
    assert.deepEqual(result.directoryAudits, [
      { id: a1.id, activityDateTime: a1.activityDateTime, activityDisplayName: a1.activityDisplayName, result: a1.result },
      { id: a2.id, activityDateTime: a2.activityDateTime, activityDisplayName: a2.activityDisplayName, result: a2.result },
    ]);
    assert.deepEqual(result.count, "2 directory audits");
    assert.equal(result.total, null);
    assert.equal(result.complete, true);
    assert.ok(result.help.some(hint => hint.includes("entra directory-audit show --id <directory-audit-id>")));
  } finally {
    teardownProfiles(state);
  }
});

for (const [noun, dateField] of [["sign-in", "createdDateTime"], ["directory-audit", "activityDateTime"]]) {
  test(`${noun} resumes precise bounds canonically and rejects a changed fractional bound`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor("delegated");
      const first = await executeArgv(["entra", noun, "list", "--profile", "soc",
        "--since", "2026-09-10T21:20:02.7215374Z", "--limit", "1"], overrides);
      const resumed = await executeArgv(["entra", noun, "list", "--profile", "soc", "--cursor", first.cursor,
        "--since", "2026-09-10T23:20:02.721537400+02:00"], overrides);
      assert.equal(resumed.complete, true);
      await assert.rejects(executeArgv(["entra", noun, "list", "--profile", "soc", "--cursor", first.cursor,
        "--since", "2026-09-10T21:20:02.7215375Z"], overrides), { code: "VALIDATION_ERROR" });
    } finally { teardownProfiles(state); }
  });

  for (const [since, until, expectedSince, expectedUntil] of [
    ["2026-09-10T21:20:02.7215373Z", "2026-09-10T21:20:02.7215374Z",
      "2026-09-10T21:20:02.7215373Z", "2026-09-10T21:20:02.7215374Z"],
    ["2026-09-10T23:20:02.721+02:00", "2026-09-10T14:20:02.7215374-07:00",
      "2026-09-10T21:20:02.721Z", "2026-09-10T21:20:02.7215374Z"],
    ["2024-02-29T00:00:00Z", "2024-03-01T00:00:00Z",
      "2024-02-29T00:00:00.000Z", "2024-03-01T00:00:00.000Z"],
  ]) {
    test(`${noun} retains exact time bounds: ${since} to ${until}`, async () => {
      const state = setupProfiles();
      try {
        const fixture = transport(() => json(200, { value: [] }));
        const { requests, overrides } = overridesFor("delegated", fixture);
        await executeArgv(["entra", noun, "list", "--profile", "soc", "--since", since, "--until", until], overrides);
        assert.equal(new URL(requests[0].url).searchParams.get("$filter"),
          `${dateField} ge ${expectedSince} and ${dateField} le ${expectedUntil}`);
      } finally { teardownProfiles(state); }
    });
  }

  for (const flag of ["since", "until"]) {
    for (const invalid of ["2026-02-30T00:00:00Z", "2026-02-29T00:00:00Z", "2026-09-10T21:20:02",
      "2026-09-10", "2026-09-10T24:00:00Z", "2026-09-10T21:20:60Z", "2026-09-10T21:20:02+24:00"]) {
      test(`${noun} rejects invalid --${flag} ${invalid} before credentials`, async () => {
        const state = setupProfiles();
        try {
          const { requests, calls, overrides } = overridesFor("delegated");
          const bounds = flag === "since" ? ["--since", invalid] : ["--since", SINCE, "--until", invalid];
          await assert.rejects(executeArgv(["entra", noun, "list", "--profile", "soc", ...bounds], overrides),
            { code: "VALIDATION_ERROR" });
          assert.deepEqual(calls, []);
          assert.deepEqual(requests, []);
        } finally { teardownProfiles(state); }
      });
    }
  }

  for (const until of ["2026-09-10T21:20:02.7215373Z", "2026-09-10T23:20:02.721537400+02:00"]) {
    test(`${noun} rejects reversed or equal fractional bounds: ${until}`, async () => {
      const state = setupProfiles();
      try {
        const { requests, calls, overrides } = overridesFor("delegated");
        await assert.rejects(executeArgv(["entra", noun, "list", "--profile", "soc",
          "--since", "2026-09-10T21:20:02.7215374Z", "--until", until], overrides), { code: "VALIDATION_ERROR" });
        assert.deepEqual(calls, []);
        assert.deepEqual(requests, []);
      } finally { teardownProfiles(state); }
    });
  }
}

for (const args of [
  ["entra", "sign-in", "list", "--profile", "soc"],
  ["entra", "sign-in", "list", "--profile", "soc", "--until", UNTIL],
  ["entra", "directory-audit", "list", "--profile", "soc"],
]) {
  test(`list without --since fails before HTTP: ${args.slice(0, 2).join(" ")}${args.includes("--until") ? " with only --until" : ""}`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor("delegated");
      await assert.rejects(executeArgv(args, overrides), error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /bounded in time/.test(error.message);
      });
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

for (const args of [
  ["entra", "sign-in", "list", "--profile", "soc", "--since", "not-a-time"],
  ["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE, "--until", SINCE],
  ["entra", "sign-in", "list", "--profile", "soc", "--since", UNTIL, "--until", SINCE],
  ["entra", "directory-audit", "list", "--profile", "soc", "--since", SINCE, "--until", "2026-08-01T00:00:00Z"],
  ["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE, "--filter", "  "],
]) {
  test(`invalid time bounds fail before HTTP: ${args.join(" ")}`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor("delegated");
      await assert.rejects(executeArgv(args, overrides), { code: "VALIDATION_ERROR" });
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

test("show passes CA policy detail through when the caller can see it", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "sign-in", "show", "--id", s2.id, "--profile", "soc"], overrides);
    assert.deepEqual(result.signIn.appliedConditionalAccessPolicies, s2.appliedConditionalAccessPolicies);
    assert.equal(result.signIn.status.errorCode, 50126);
    assert.equal(result.help, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("show reports omitted CA policy detail as unavailable rather than empty", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "sign-in", "show", "--id", s1.id, "--profile", "soc",
      "--select", "id,conditionalAccessStatus,appliedConditionalAccessPolicies"], overrides);
    assert.deepEqual(result.signIn, {
      id: s1.id,
      conditionalAccessStatus: "notApplied",
      appliedConditionalAccessPolicies: result.signIn.appliedConditionalAccessPolicies,
    });
    assert.match(result.signIn.appliedConditionalAccessPolicies, /^unavailable: Graph omits CA policy detail/);
    assert.match(result.signIn.appliedConditionalAccessPolicies, /both modes need Policy.Read.All/);
    assert.match(result.signIn.appliedConditionalAccessPolicies, /delegated also needs Conditional Access Administrator, Global Reader/);
    assert.match(result.signIn.appliedConditionalAccessPolicies,
      /log in and repeat this read with --scopes https:\/\/graph.microsoft.com\/AuditLog.Read.All,https:\/\/graph.microsoft.com\/Policy.Read.All/);
    assert.equal(result.help, undefined);
  } finally {
    teardownProfiles(state);
  }
});

const longNestedText = "x".repeat(600);
const compactNestedText = `${"x".repeat(500)}... (truncated, 600 chars total)`;
for (const [noun, key, id, field, value, expected] of [
  ["sign-in", "signIn", s1.id, "status", { errorCode: 0, failureReason: longNestedText, additionalDetails: null },
    { errorCode: 0, failureReason: compactNestedText, additionalDetails: null }],
  ["sign-in", "signIn", s1.id, "location", { city: longNestedText, geoCoordinates: { latitude: 1, longitude: 2 } },
    { city: compactNestedText, geoCoordinates: { latitude: 1, longitude: 2 } }],
  ["sign-in", "signIn", s1.id, "appliedConditionalAccessPolicies", [{ displayName: longNestedText }, { displayName: "short" }],
    [{ displayName: compactNestedText }, { displayName: "short" }]],
  ["directory-audit", "directoryAudit", a1.id, "initiatedBy", { user: { displayName: longNestedText }, app: null },
    { user: { displayName: compactNestedText }, app: null }],
  ["directory-audit", "directoryAudit", a1.id, "targetResources",
    [{ modifiedProperties: [{ oldValue: null, newValue: longNestedText }], displayName: "short" }],
    [{ modifiedProperties: [{ oldValue: null, newValue: compactNestedText }], displayName: "short" }]],
]) {
  for (const action of ["list", "show"]) {
    test(`${noun} ${action} truncates nested ${field} and recovers with --full`, async () => {
      const state = setupProfiles();
      try {
        const row = { id, [field]: value };
        const fixture = transport(() => json(200, action === "list" ? { value: [row] } : row));
        const { overrides } = overridesFor("delegated", fixture);
        const binding = action === "list" ? ["--since", SINCE] : ["--id", id];
        const argv = ["entra", noun, action, "--profile", "soc", ...binding, "--select", `id,${field}`];
        const compact = await executeArgv(argv, overrides);
        const compactRow = action === "list" ? compact[`${key}s`][0] : compact[key];
        assert.deepEqual(compactRow, { id, [field]: expected });
        const full = await executeArgv(hintArgv(compact.help[0]), overrides);
        const fullRow = action === "list" ? full[`${key}s`][0] : full[key];
        assert.deepEqual(fullRow, row);
        assert.ok(!(full.help ?? []).some(hint => hint.includes("--full")));
      } finally { teardownProfiles(state); }
    });
  }
}

test("explicit nulls stay null and missing base properties stay absent", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "sign-in", "show", "--id", s3.id, "--profile", "soc",
      "--select", "id,ipAddress,location,conditionalAccessStatus"], overrides);
    assert.deepEqual(result.signIn, { id: s3.id, ipAddress: null, location: null, conditionalAccessStatus: null });
  } finally {
    teardownProfiles(state);
  }
});

test("show truncates long text unless --full is passed", async () => {
  const state = setupProfiles();
  try {
    const long = "x".repeat(600);
    const fixture = transport(() => json(200, { ...s1, userPrincipalName: long }));
    const credential = credentialService("delegated", []);
    const overrides = { transport: fixture.send, delegated: credential, application: credentialService("application", []) };
    const argv = ["entra", "sign-in", "show", "--id", s1.id, "--profile", "soc",
      "--select", "id,userPrincipalName", "--scopes", auditScopes[0]];
    const compact = await executeArgv(argv, overrides);
    assert.match(compact.signIn.userPrincipalName, /\.\.\. \(truncated, 600 chars total\)$/);
    assert.ok(compact.help.some(hint => hint.includes("--full")));
    const recovery = hintArgv(compact.help[0]);
    assert.deepEqual(recovery, [...argv, "--full"]);
    const full = await executeArgv(recovery, overrides);
    assert.equal(full.signIn.userPrincipalName, long);
    assert.equal(full.help, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("list truncation recovery replays time bounds with --full", async () => {
  const state = setupProfiles();
  try {
    const long = "x".repeat(600);
    const fixture = transport(() => json(200, { value: [{ ...s1, userPrincipalName: long }] }));
    const { overrides } = overridesFor("delegated", fixture);
    const argv = ["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE,
      "--filter", "status/errorCode ne 0", "--select", "id,userPrincipalName", "--scopes", auditScopes[0]];
    const compact = await executeArgv(argv, overrides);
    assert.match(compact.signIns[0].userPrincipalName, /truncated/);
    const recovery = hintArgv(compact.help[0]);
    assert.deepEqual(recovery, [...argv, "--full"]);
    const full = await executeArgv(recovery, overrides);
    assert.equal(full.signIns[0].userPrincipalName, long);
    assert.deepEqual(full.count, "1 sign-ins");
    assert.equal(full.total, null);
    assert.equal(full.complete, true);
  } finally {
    teardownProfiles(state);
  }
});

test("denied sign-in reads name sign-in roles and licensing", async () => {
  const state = setupProfiles();
  try {
    const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
    const credential = credentialService("delegated", []);
    const overrides = { transport: denied.send, delegated: credential, application: credentialService("application", []) };
    await assert.rejects(executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.match(error.message, /grant, role, licence or policy/);
      assert.ok(error.suggestions.some(hint => hint.includes("Security Operator")));
      assert.ok(error.suggestions.some(hint => hint.includes("P1/P2")));
      return true;
    });
    await assert.rejects(executeArgv(["entra", "sign-in", "show", "--id", s1.id, "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      return error.suggestions.some(hint => hint.includes("Global Reader"));
    });
  } finally {
    teardownProfiles(state);
  }
});

test("denied directory-audit reads name audit roles, not sign-in roles", async () => {
  const state = setupProfiles();
  try {
    const denied = transport(() => json(404, { error: { code: "Unknown", message: "gone" } }));
    const credential = credentialService("delegated", []);
    const notFound = { transport: denied.send, delegated: credential, application: credentialService("application", []) };
    await assert.rejects(executeArgv(["entra", "directory-audit", "show", "--id", a1.id, "--profile", "soc"], notFound), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      return /not found or inaccessible/.test(error.message);
    });
    const forbidden = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
    const forbiddenOverrides = { transport: forbidden.send, delegated: credential, application: credentialService("application", []) };
    await assert.rejects(executeArgv(["entra", "directory-audit", "list", "--profile", "soc", "--since", SINCE], forbiddenOverrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.ok(error.suggestions.some(hint => hint.includes("Reports Reader")));
      assert.ok(!error.suggestions.some(hint => hint.includes("Security Operator")));
      return true;
    });
  } finally {
    teardownProfiles(state);
  }
});

test("a capped sign-in list resumes losslessly through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE, "--limit", "2"], overrides);
    assert.deepEqual(first.signIns.map(row => row.id), [s1.id, s2.id]);
    assert.equal(first.complete, false);
    assert.match(first.reason, /row limit/);
    assert.equal(typeof first.cursor, "string");
    assert.ok(!first.cursor.includes(s1.id));
    const second = await executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.signIns.map(row => row.id), [s3.id]);
    assert.deepEqual(second.count, "1 sign-ins");
    assert.equal(second.total, null);
    assert.equal(second.complete, true);
    assert.ok(!JSON.stringify(second).includes("opaque-fixture-delegated-token"));
    assert.ok(requests.length >= 2);
  } finally {
    teardownProfiles(state);
  }
});

test("resume repeats identical time bounds but rejects a conflicting --since", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE, "--limit", "1"], overrides);
    assert.equal(first.complete, false);
    const resumed = await executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--cursor", first.cursor, "--since", SINCE], overrides);
    assert.deepEqual(resumed.signIns.map(row => row.id), [s2.id, s3.id]);
    await assert.rejects(
      executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--cursor", first.cursor, "--since", UNTIL], overrides),
      { code: "VALIDATION_ERROR" },
    );
  } finally {
    teardownProfiles(state);
  }
});

test("an empty log list is a definitive zero with widening guidance", async () => {
  const state = setupProfiles();
  try {
    const fixture = transport(() => json(200, { value: [] }));
    const credential = credentialService("delegated", []);
    const overrides = { transport: fixture.send, delegated: credential, application: credentialService("application", []) };
    const result = await executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE], overrides);
    assert.deepEqual(result.signIns, []);
    assert.deepEqual(result.count, "0 sign-ins");
    assert.equal(result.total, null);
    assert.equal(result.complete, true);
    assert.ok(result.help.some(hint => hint.includes("0 sign-ins matched in this window")));
  } finally {
    teardownProfiles(state);
  }
});

test("application mode lists and shows without delegated scopes", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("application");
    const listed = await executeArgv(["entra", "sign-in", "list", "--profile", "batch", "--since", SINCE, "--all"], overrides);
    assert.deepEqual(listed.signIns.map(row => row.id), [s1.id, s2.id, s3.id]);
    assert.equal(calls.length, 1);
    assert.ok(requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-application-token"));
    const shown = await executeArgv(["entra", "directory-audit", "show", "--id", a1.id, "--profile", "batch"], overrides);
    assert.equal(shown.directoryAudit.activityDisplayName, "Add member to group");
    assert.deepEqual(shown.directoryAudit.targetResources, a1.targetResources);
  } finally {
    teardownProfiles(state);
  }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "sign-in", "list", "--profile", "batch", "--since", SINCE, "--scopes", auditScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("unknown properties and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE, "--select", "id,aboutMe"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE, "--fields", "riskState"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "directory-audit", "show", "--id", a1.id, "--profile", "soc", "--select", "id,aboutMe"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

for (const scopes of [",", "https://example.invalid/AuditLog.Read.All", "https://graph.microsoft.com/.default"]) {
  test(`sign-in list rejects delegated scopes ${scopes} before provider acquisition`, async () => {
    const state = setupProfiles();
    try {
      const calls = [];
      const { overrides, requests } = overridesFor("delegated");
      overrides.delegated = new DelegatedAuth({
        storage: "session-only",
        login: async () => { throw new Error("Unexpected interactive login"); },
        silent: async (...args) => { calls.push(args); throw new Error("Unexpected credential acquisition"); },
      });
      await assert.rejects(executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE, "--scopes", scopes], overrides), { code: "VALIDATION_ERROR" });
      assert.deepEqual(calls, []);
      assert.deepEqual(requests, []);
    } finally { teardownProfiles(state); }
  });
}

test("richer selects dispatch the same mapping while beta stays preview-gated", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const richer = await executeArgv(
      ["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE, "--select", "id,riskState", "--all"],
      overrides,
    );
    assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,riskState");
    assert.deepEqual(richer.signIns[0], { id: s1.id, riskState: "none" });
    const beta = executeArgv(["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE, "--api-version", "beta", "--limit", "1"], overrides);
    await assert.rejects(beta, error => {
      assert.equal(error.code, "POLICY_DENIED");
      return /preview-enabled/.test(error.message);
    });
  } finally {
    teardownProfiles(state);
  }
});
