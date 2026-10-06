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
const SINCE = "2026-09-01T00:00:00Z";

const s1 = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  displayName: "Contoso App",
  appId: "b55552fe-a272-4b56-990b-95038d917878",
  servicePrincipalType: "Application",
  riskLevel: "high",
  riskState: "atRisk",
  riskDetail: "none",
  riskLastUpdatedDateTime: "2026-09-12T10:00:00Z",
  isEnabled: true,
  isProcessing: false,
};
const s2 = {
  id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  displayName: "Fabrikam Sync",
  appId: "c66663ff-b383-4c67-a11c-06049e028989",
  servicePrincipalType: "Application",
  riskLevel: "medium",
  riskState: "atRisk",
  riskDetail: "anomalousServicePrincipalActivity",
  riskLastUpdatedDateTime: "2026-09-11T09:00:00Z",
  isEnabled: true,
  isProcessing: false,
};
const s3 = {
  id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  displayName: "Hidden Watcher",
  appId: "d77774aa-c494-4d78-b22d-17150f139090",
  servicePrincipalType: "Application",
  riskLevel: "hidden",
  riskState: "atRisk",
  riskDetail: "hidden",
  riskLastUpdatedDateTime: "2026-09-10T08:00:00Z",
  isEnabled: true,
  isProcessing: true,
};
const principals = [s1, s2, s3];

const h1 = {
  owner: s1.id,
  id: "h1111111-1111-4111-8111-111111111111",
  displayName: "Contoso App",
  appId: s1.appId,
  servicePrincipalType: "Application",
  servicePrincipalId: s1.id,
  riskLevel: "high",
  riskState: "atRisk",
  riskDetail: "none",
  riskLastUpdatedDateTime: "2026-09-12T10:00:00Z",
  isEnabled: true,
  isProcessing: false,
  initiatedBy: "system",
  activity: { riskEventTypes: ["leakedCredentials"], detail: "none" },
};
const h2 = {
  owner: s1.id,
  id: "h2222222-2222-4222-8222-222222222222",
  displayName: "Contoso App",
  appId: s1.appId,
  servicePrincipalType: "Application",
  servicePrincipalId: s1.id,
  riskLevel: "medium",
  riskState: "dismissed",
  riskDetail: "userPerformedSecuredPasswordReset",
  riskLastUpdatedDateTime: "2026-09-09T08:00:00Z",
  isEnabled: true,
  isProcessing: false,
  initiatedBy: "admin",
  activity: null,
};
const history = [h1, h2];

const sd1 = {
  id: "e1111111-1111-4111-8111-111111111111",
  requestId: null,
  correlationId: null,
  riskEventType: "leakedCredentials",
  riskState: "atRisk",
  riskLevel: "high",
  riskDetail: "leakedCredentials",
  source: "IdentityProtection",
  detectionTimingType: "offline",
  activity: "servicePrincipal",
  tokenIssuerType: "AzureAD",
  ipAddress: null,
  location: null,
  activityDateTime: "2026-09-10T08:00:00Z",
  detectedDateTime: "2026-09-10T08:11:27Z",
  lastUpdatedDateTime: "2026-09-10T08:11:27Z",
  servicePrincipalId: s1.id,
  servicePrincipalDisplayName: "Contoso App",
  appId: s1.appId,
  keyIds: ["9d9fea30-d8e3-481b-b57c-0ef569a989e5"],
  additionalInfo: "[]",
};
const sd2 = {
  id: "e2222222-2222-4222-8222-222222222222",
  requestId: "r2222222-2222-4222-8222-222222222222",
  correlationId: "c2222222-2222-4222-8222-222222222222",
  riskEventType: "suspiciousSignins",
  riskState: "atRisk",
  riskLevel: "hidden",
  riskDetail: "hidden",
  source: "IdentityProtection",
  detectionTimingType: "nearRealtime",
  activity: "servicePrincipal",
  ipAddress: "131.107.159.38",
  location: { city: "Redmond", state: "Washington", countryOrRegion: "US", geoCoordinates: null },
  activityDateTime: "2026-09-11T09:30:00Z",
  detectedDateTime: "2026-09-11T09:41:00Z",
  lastUpdatedDateTime: "2026-09-11T09:41:00Z",
  servicePrincipalId: s2.id,
  servicePrincipalDisplayName: "Fabrikam Sync",
  appId: s2.appId,
  keyIds: [],
  additionalInfo: "[]",
};
const sd3 = {
  id: "e3333333-3333-4333-8333-333333333333",
  requestId: null,
  correlationId: null,
  riskEventType: "generic",
  riskState: "dismissed",
  riskLevel: "low",
  riskDetail: "none",
  source: "IdentityProtection",
  detectionTimingType: "realtime",
  activity: "servicePrincipal",
  ipAddress: null,
  location: null,
  activityDateTime: "2026-09-12T10:00:00Z",
  detectedDateTime: "2026-09-12T11:00:00Z",
  lastUpdatedDateTime: "2026-09-12T11:00:00Z",
  servicePrincipalId: s3.id,
  servicePrincipalDisplayName: "Hidden Watcher",
  appId: s3.appId,
  keyIds: ["aa0fea30-d8e3-481b-b57c-0ef569a989e5"],
  additionalInfo: "[]",
};
const detections = [sd1, sd2, sd3];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-workload-risk-"));
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

function workloadTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/identityProtection/riskyServicePrincipals") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [s3] });
      return json(200, {
        value: [s1, s2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identityProtection/riskyServicePrincipals?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/identityProtection/servicePrincipalRiskDetections") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [sd3] });
      return json(200, {
        value: [sd1, sd2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identityProtection/servicePrincipalRiskDetections?%24skiptoken=page2",
      });
    }
    const principal = principals.find(row => path === `/v1.0/identityProtection/riskyServicePrincipals/${row.id}`);
    if (principal) return json(200, principal);
    const historyMatch = /^\/v1\.0\/identityProtection\/riskyServicePrincipals\/([^/]+)\/history(?:\/([^/]+))?$/.exec(path);
    if (historyMatch) {
      const items = history.filter(row => row.owner === historyMatch[1]).map(({ owner, ...row }) => row);
      const single = historyMatch[2] === undefined ? undefined : items.find(row => row.id === historyMatch[2]);
      return json(200, historyMatch[2] === undefined ? { value: items } : single);
    }
    const detection = detections.find(row => path === `/v1.0/identityProtection/servicePrincipalRiskDetections/${row.id}`);
    if (detection) return json(200, detection);
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? workloadTransport();
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
    "--import", pathToFileURL(resolve("test/fixtures/read-workload-risk-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["pipe", "pipe", "pipe"], input, timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, principals, history, detections, denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} executable lists risky service principals with compact rows`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "risky-service-principal", "list", "--profile", profile, "--limit", "10"], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.riskyServicePrincipals, [
        { id: s1.id, displayName: s1.displayName, riskLevel: s1.riskLevel, riskState: s1.riskState },
        { id: s2.id, displayName: s2.displayName, riskLevel: s2.riskLevel, riskState: s2.riskState },
        { id: s3.id, displayName: s3.displayName, riskLevel: s3.riskLevel, riskState: s3.riskState },
      ]);
      assert.deepEqual(output.count, { returned: 3, complete: true });
      assert.ok(output.help.some(hint => hint.includes("entra risky-service-principal show --id <risky-service-principal-id>")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable lists bounded service-principal detections with compact rows`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "service-principal-risk-detection", "list", "--profile", profile, "--since", SINCE], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.servicePrincipalRiskDetections, [
        { id: sd1.id, detectedDateTime: sd1.detectedDateTime, servicePrincipalDisplayName: sd1.servicePrincipalDisplayName, riskLevel: sd1.riskLevel },
        { id: sd2.id, detectedDateTime: sd2.detectedDateTime, servicePrincipalDisplayName: sd2.servicePrincipalDisplayName, riskLevel: sd2.riskLevel },
        { id: sd3.id, detectedDateTime: sd3.detectedDateTime, servicePrincipalDisplayName: sd3.servicePrincipalDisplayName, riskLevel: sd3.riskLevel },
      ]);
      assert.deepEqual(output.count, { returned: 3, complete: true });
      assert.ok(output.help.some(hint => hint.includes("entra service-principal-risk-detection show --id <service-principal-risk-detection-id>")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable lists one principal's risk history`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "risky-service-principal", "history", "list",
        "--service-principal", s1.id, "--profile", profile], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.riskyServicePrincipalHistory.map(row => row.id), [h1.id, h2.id]);
      assert.deepEqual(output.count, { returned: 2, complete: true });
      assert.ok(output.help.some(hint => hint.includes("entra risky-service-principal history show --service-principal <risky-service-principal-id> --id <history-item-id>")));
    } finally { teardownProfiles(state); }
  });
}

test("delegated executable shows one risky service principal with detection and history correlation", () => {
  const state = setupProfiles();
  try {
    const result = runReadCli(["entra", "risky-service-principal", "show", "--profile", "soc", "--id", s1.id], state, "delegated");
    assert.equal(result.status, 0, result.stdout);
    const output = decode(result.stdout);
    assert.deepEqual(output.riskyServicePrincipal, s1);
    assert.ok(output.help.some(hint => hint.includes("entra service-principal-risk-detection list")));
    assert.ok(output.help.some(hint => hint.includes("entra risky-service-principal history list")));
    assert.ok(output.help.some(hint => hint.includes("never confirms, dismisses or remediates")));
  } finally { teardownProfiles(state); }
});

test("delegated executable shows one history item with detection correlation", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const shown = await executeArgv(["entra", "risky-service-principal", "history", "show",
      "--profile", "soc", "--service-principal", s1.id, "--id", h1.id], overrides);
    const { owner, ...expected } = h1;
    assert.deepEqual(shown.riskyServicePrincipalHistoryItem, expected);
    assert.ok(new URL(requests[0].url).pathname.includes(`/riskyServicePrincipals/${s1.id}/history/${h1.id}`));
    assert.ok(shown.help.some(hint => hint.includes("entra service-principal-risk-detection list")));
  } finally { teardownProfiles(state); }
});

test("delegated executable shows one service-principal detection with key identifiers and hidden-value guidance", () => {
  const state = setupProfiles();
  try {
    const result = runReadCli(["entra", "service-principal-risk-detection", "show", "--profile", "soc", "--id", sd1.id], state, "delegated");
    assert.equal(result.status, 0, result.stdout);
    const output = decode(result.stdout);
    assert.equal(output.servicePrincipalRiskDetection.id, sd1.id);
    assert.deepEqual(output.servicePrincipalRiskDetection.keyIds, sd1.keyIds);
    assert.ok(!Object.hasOwn(output.servicePrincipalRiskDetection, "tokenIssuerType"));
    assert.ok(output.help.some(hint => hint.includes("hidden")));
  } finally { teardownProfiles(state); }
});

test("detection show without an associated sign-in says correlation does not apply", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const shown = await executeArgv(["entra", "service-principal-risk-detection", "show",
      "--profile", "soc", "--id", sd1.id], overrides);
    assert.equal(shown.servicePrincipalRiskDetection.correlationId, null);
    assert.ok(shown.help.some(hint => hint.includes("no associated sign-in")));
  } finally { teardownProfiles(state); }
});

test("hidden licence-bound values pass through instead of an invented level", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "service-principal-risk-detection", "list",
      "--profile", "soc", "--since", SINCE, "--select", "id,riskLevel,riskDetail"], overrides);
    assert.deepEqual(result.servicePrincipalRiskDetections.map(row => row.riskLevel), ["high", "hidden", "low"]);
    assert.ok(result.help.some(hint => hint.includes("hidden")));
  } finally { teardownProfiles(state); }
});

for (const noun of ["risky-service-principal", "service-principal-risk-detection"]) {
  test(`delegated executable denied ${noun} list is an operational error on stdout`, () => {
    const state = setupProfiles();
    try {
      const args = noun === "risky-service-principal"
        ? ["entra", "risky-service-principal", "list", "--profile", "soc"]
        : ["entra", "service-principal-risk-detection", "list", "--profile", "soc", "--since", SINCE];
      const result = runReadCli(args, state, "delegated", true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /\(403\)/);
      assert.ok(output.help.some(hint => hint.includes("Global Reader")));
      assert.ok(output.help.some(hint => hint.includes("Workload Identities Premium")));
    } finally { teardownProfiles(state); }
  });
}

test("history list without --service-principal fails before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "risky-service-principal", "history", "list", "--profile", "soc"], overrides),
      { code: "VALIDATION_ERROR" });
    assert.deepEqual(requests, []);
    assert.deepEqual(calls, []);
  } finally { teardownProfiles(state); }
});

test("history show without --service-principal fails before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "risky-service-principal", "history", "show",
      "--profile", "soc", "--id", h1.id], overrides), { code: "VALIDATION_ERROR" });
    assert.deepEqual(requests, []);
    assert.deepEqual(calls, []);
  } finally { teardownProfiles(state); }
});

test("service-principal detection list without --since fails before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "service-principal-risk-detection", "list", "--profile", "soc"], overrides),
      { code: "VALIDATION_ERROR" });
    assert.deepEqual(requests, []);
    assert.deepEqual(calls, []);
  } finally { teardownProfiles(state); }
});

test("service-principal detection list composes since and filter into one bounded $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "service-principal-risk-detection", "list", "--profile", "soc",
      "--since", SINCE, "--filter", "riskState eq 'atRisk'", "--all"], overrides);
    assert.equal(new URL(requests[0].url).searchParams.get("$filter"),
      "detectedDateTime ge 2026-09-01T00:00:00.000Z and (riskState eq 'atRisk')");
    assert.deepEqual(result.count, { returned: 3, complete: true });
    assert.equal(requests.length, 2);
  } finally { teardownProfiles(state); }
});

test("a capped risky-service-principal list resumes losslessly through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "risky-service-principal", "list", "--profile", "soc", "--limit", "2"], overrides);
    assert.deepEqual(first.riskyServicePrincipals.map(row => row.id), [s1.id, s2.id]);
    assert.equal(first.count.complete, false);
    const resumed = await executeArgv(["entra", "risky-service-principal", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(resumed.riskyServicePrincipals.map(row => row.id), [s3.id]);
    assert.deepEqual(resumed.count, { returned: 1, complete: true });
  } finally { teardownProfiles(state); }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(executeArgv(["entra", "risky-service-principal", "list", "--profile", "batch",
      "--scopes", "https://graph.microsoft.com/IdentityRiskyServicePrincipal.Read.All"], overrides), error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      return /Graph \.default audience/.test(error.message);
    });
    await assert.rejects(executeArgv(["entra", "service-principal-risk-detection", "list", "--profile", "batch", "--since", SINCE,
      "--scopes", "https://graph.microsoft.com/IdentityRiskEvent.Read.All"], overrides), { code: "VALIDATION_ERROR" });
    assert.deepEqual(requests, []);
  } finally {
    teardownProfiles(state);
  }
});

test("unknown workload properties and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "risky-service-principal", "list", "--profile", "soc", "--select", "id,aboutMe"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "service-principal-risk-detection", "list", "--profile", "soc", "--since", SINCE, "--fields", "riskState"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "service-principal-risk-detection", "show", "--profile", "soc", "--id", sd1.id, "--select", "id,tokenIssuerType"], overrides), error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      assert.match(error.message, /Unknown property tokenIssuerType/);
      return true;
    });
    await assert.rejects(executeArgv(["entra", "risky-service-principal", "history", "list",
      "--service-principal", s1.id, "--profile", "soc", "--select", "id,userPrincipalName"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("richer selects dispatch the same mapping while beta stays preview-gated", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const richer = await executeArgv(
      ["entra", "risky-service-principal", "list", "--profile", "soc", "--select", "id,riskState", "--all"],
      overrides,
    );
    assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,riskState");
    assert.deepEqual(richer.riskyServicePrincipals[0], { id: s1.id, riskState: "atRisk" });
    for (const argv of [
      ["entra", "risky-service-principal", "list", "--profile", "soc", "--api-version", "beta", "--limit", "1"],
      ["entra", "service-principal-risk-detection", "list", "--profile", "soc", "--since", SINCE, "--api-version", "beta", "--limit", "1"],
      ["entra", "risky-service-principal", "history", "list", "--service-principal", s1.id, "--profile", "soc", "--api-version", "beta"],
    ]) {
      await assert.rejects(executeArgv(argv, overrides), error => {
        assert.equal(error.code, "POLICY_DENIED");
        return /preview-enabled/.test(error.message);
      });
    }
  } finally {
    teardownProfiles(state);
  }
});

test("delegated scopes are validated before provider acquisition", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "risky-service-principal", "list", "--profile", "soc",
      "--scopes", "https://graph.microsoft.com/.default"], overrides), { code: "VALIDATION_ERROR" });
    assert.deepEqual(requests, []);
    assert.deepEqual(calls, []);
  } finally {
    teardownProfiles(state);
  }
});
