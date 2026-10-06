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

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const provisionScopes = [
  "https://graph.microsoft.com/AuditLog.Read.All",
  "https://graph.microsoft.com/Directory.Read.All",
];
const SINCE = "2026-09-01T00:00:00Z";
const UNTIL = "2026-09-08T00:00:00Z";

const p1 = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  activityDateTime: "2026-09-10T08:00:00Z",
  tenantId: tenant,
  jobId: "job-1111",
  cycleId: "cycle-1111",
  changeId: "chg-1111",
  action: "create",
  durationInMilliseconds: 1200,
  sourceSystem: { id: "source-system", displayName: "Microsoft Entra ID", details: {} },
  sourceIdentity: { id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb", displayName: "Adele Vance", identityType: "user", details: {} },
  targetSystem: { id: "target-system", displayName: "Salesforce", details: {} },
  targetIdentity: { id: "target-1111", displayName: "AdeleV", identityType: "user", details: {} },
  provisioningStatusInfo: { status: "success", statusInfo: "Provisioned", errorInformation: null },
  provisioningSteps: [{ name: "EntryImport", provisioningStepType: "import", status: "success", description: null, details: {} }],
  modifiedProperties: [{ displayName: "displayName", oldValue: null, newValue: "Adele Vance" }],
  servicePrincipal: { id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc", appId: "de8bc8b5-d9f9-48b1-a8ad-b748da725064", displayName: "Salesforce" },
  initiatedBy: { initiatingType: "system", initiatingSystem: { displayName: "Sync engine" } },
};
const p2 = {
  id: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
  activityDateTime: "2026-09-11T09:30:00Z",
  tenantId: tenant,
  jobId: "job-2222",
  cycleId: "cycle-2222",
  changeId: "chg-2222",
  action: "update",
  durationInMilliseconds: 3400,
  sourceSystem: { id: "source-system", displayName: "Microsoft Entra ID", details: {} },
  sourceIdentity: { id: "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee", displayName: "Alex Wilber", identityType: "user", details: {} },
  targetSystem: { id: "target-system", displayName: "Salesforce", details: {} },
  targetIdentity: { id: "target-2222", displayName: "AlexW", identityType: "user", details: {} },
  provisioningStatusInfo: { status: "failure", statusInfo: "Failed", errorInformation: "EntrySynchronizationSkipped" },
  provisioningSteps: [{ name: "EntryExportUpdate", provisioningStepType: "export", status: "failure", description: "Target rejected the update", details: {} }],
  modifiedProperties: [{ displayName: "department", oldValue: "Sales", newValue: "Marketing" }],
  servicePrincipal: { id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc", appId: "de8bc8b5-d9f9-48b1-a8ad-b748da725064", displayName: "Salesforce" },
  initiatedBy: { initiatingType: "user", initiatingSystem: null },
};
const events = [p1, p2];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-prov-"));
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

function provisioningTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/auditLogs/provisioning") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [p2] });
      return json(200, {
        value: [p1],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/auditLogs/provisioning?%24skiptoken=page2",
      });
    }
    const event = events.find(row => path === `/v1.0/auditLogs/provisioning/${row.id}`);
    if (event) return json(200, event);
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? provisioningTransport();
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

function runReadCli(args, state, mode, denied = false, input) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-provisioning-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["pipe", "pipe", "pipe"], input, timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, events, denied }),
    },
  });
}

const compact = row => ({ id: row.id, activityDateTime: row.activityDateTime, action: row.action, provisioningStatusInfo: row.provisioningStatusInfo });

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} executable lists bounded provisioning with compact rows`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "provisioning", "list", "--profile", profile, "--since", SINCE], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.provisioning, [compact(p1), compact(p2)]);
      assert.deepEqual(output.count, "2 provisioning events");
      assert.equal(output.total, null);
      assert.equal(output.complete, true);
      assert.ok(output.help.some(hint => hint.includes("entra provisioning show --id <provisioning-id>")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable shows one provisioning event with the full reviewed set`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "provisioning", "show", "--id", p1.id, "--profile", profile], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.provisioning, p1);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied provisioning list is an operational error on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "provisioning", "list", "--profile", profile, "--since", SINCE], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.ok(output.help.some(hint => hint.includes("AuditLog.Read.All") && hint.includes("Directory.Read.All")));
      assert.ok(output.help.some(hint => hint.includes("P1 or P2")));
      assert.equal(output.provisioning, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });
}

test("delegated provisioning list requests both directory scopes with a bounded filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "provisioning", "list", "--profile", "soc",
      "--since", SINCE, "--until", UNTIL, "--filter", "provisioningStatusInfo/status eq 'failure'", "--all"], overrides);
    assert.equal(requests.length, 2);
    assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/auditLogs/provisioning"));
    const params = new URL(requests[0].url).searchParams;
    assert.equal(params.get("$filter"),
      "activityDateTime ge 2026-09-01T00:00:00.000Z and activityDateTime le 2026-09-08T00:00:00.000Z and (provisioningStatusInfo/status eq 'failure')");
    assert.equal(params.get("$select"), "id,activityDateTime,action,provisioningStatusInfo");
    assert.deepEqual(calls[0][1], provisionScopes);
    assert.deepEqual(result.provisioning.map(row => row.id), [p1.id, p2.id]);
    assert.deepEqual(result.count, "2 provisioning events");
    assert.equal(result.total, null);
    assert.equal(result.complete, true);
    assert.ok(!JSON.stringify(result).includes("opaque-fixture-delegated-token"));
  } finally {
    teardownProfiles(state);
  }
});

test("provisioning list requires a time bound and stops before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "provisioning", "list", "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      assert.match(error.message, /bounded in time/);
      return true;
    });
    assert.deepEqual(requests, []);
    assert.deepEqual(calls, []);
  } finally {
    teardownProfiles(state);
  }
});

test("explicit scopes override the dual provisioning default", async () => {
  const state = setupProfiles();
  try {
    const { calls, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "provisioning", "list", "--profile", "soc",
      "--since", SINCE, "--all", "--scopes", "https://graph.microsoft.com/AuditLog.Read.All"], overrides);
    assert.deepEqual(calls[0][1], ["https://graph.microsoft.com/AuditLog.Read.All"]);
    assert.deepEqual(result.count, "2 provisioning events");
  } finally {
    teardownProfiles(state);
  }
});

test("application profiles reject delegated scopes", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("application");
    await assert.rejects(executeArgv(["entra", "provisioning", "list", "--profile", "batch",
      "--since", SINCE, "--scopes", "https://graph.microsoft.com/AuditLog.Read.All"], overrides), error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      assert.match(error.message, /delegated scopes are unavailable/);
      return true;
    });
  } finally {
    teardownProfiles(state);
  }
});

test("unreviewed select properties fail before credentials, never as server queries", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "provisioning", "show", "--profile", "soc",
      "--id", p1.id, "--select", "id,passwordProfile"], overrides), error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      assert.match(error.message, /Unknown property passwordProfile/);
      return true;
    });
    assert.deepEqual(requests, []);
    assert.deepEqual(calls, []);
  } finally {
    teardownProfiles(state);
  }
});

test("long nested status text truncates with a --full replay hint", async () => {
  const state = setupProfiles();
  try {
    const rows = [{ id: p1.id, activityDateTime: p1.activityDateTime, action: "update",
      provisioningStatusInfo: { status: "failure", errorInformation: "x".repeat(600) } }];
    const fixture = transport(() => json(200, { value: rows }));
    const { overrides } = overridesFor("delegated", fixture);
    const result = await executeArgv(["entra", "provisioning", "list", "--profile", "soc", "--since", SINCE], overrides);
    assert.match(result.provisioning[0].provisioningStatusInfo.errorInformation, /truncated/);
    assert.ok(result.help[0].includes("--full"));
  } finally {
    teardownProfiles(state);
  }
});

test("empty windows report absence as the answer", async () => {
  const state = setupProfiles();
  try {
    const fixture = transport(() => json(200, { value: [] }));
    const { overrides } = overridesFor("delegated", fixture);
    const result = await executeArgv(["entra", "provisioning", "list", "--profile", "soc", "--since", SINCE], overrides);
    assert.deepEqual(result.provisioning, []);
    assert.deepEqual(result.count, "0 provisioning events");
    assert.equal(result.complete, true);
    assert.ok(result.help.some(hint => hint.includes("absence of results is the answer")));
  } finally {
    teardownProfiles(state);
  }
});

test("provisioning show projects --fields locally from the fetched selection", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "provisioning", "show", "--profile", "soc",
      "--id", p1.id, "--select", "id,action", "--fields", "action"], overrides);
    assert.deepEqual(result.provisioning, { action: "create" });
  } finally {
    teardownProfiles(state);
  }
});

test("provisioning executable rejects invalid stdin cursors", () => {
  const state = setupProfiles();
  try {
    const result = runReadCli(["entra", "provisioning", "list", "--profile", "soc", "--cursor", "-"],
      state, "delegated", true, "invalid");
    assert.equal(result.status, 2, result.stdout);
    assert.equal(result.stderr, "");
    const output = decode(result.stdout);
    assert.equal(output.code, "VALIDATION_ERROR");
    assert.match(output.error, /Invalid collection cursor/);
  } finally { teardownProfiles(state); }
});
