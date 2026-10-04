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
const SINCE = "2026-09-01T00:00:00Z";
const UNTIL = "2026-09-08T00:00:00Z";

const u1 = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  userPrincipalName: "AdeleV@contoso.com",
  userDisplayName: "Adele Vance",
  riskLevel: "high",
  riskState: "atRisk",
  riskDetail: "leakedCredentials",
  riskLastUpdatedDateTime: "2026-09-12T10:00:00Z",
  isDeleted: false,
  isProcessing: false,
};
const u2 = {
  id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  userPrincipalName: "AlexW@contoso.com",
  userDisplayName: "Alex Wilber",
  riskLevel: "medium",
  riskState: "dismissed",
  riskDetail: "userPerformedSecuredPasswordReset",
  riskLastUpdatedDateTime: "2026-09-11T09:00:00Z",
  isDeleted: false,
  isProcessing: false,
};
const u3 = {
  id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  userPrincipalName: "GradyA@contoso.com",
  userDisplayName: "Grady Archie",
  riskLevel: "hidden",
  riskState: "atRisk",
  riskDetail: "hidden",
  riskLastUpdatedDateTime: "2026-09-10T08:00:00Z",
  isDeleted: false,
  isProcessing: true,
};
const riskyUsers = [u1, u2, u3];

const LONG_INFO = `[{"Key":"userAgent","Value":"${"Mozilla/5.0 (Windows NT 10.0; Win64; x64) ".repeat(20)}"}]`;
const d1 = {
  id: "d1111111-1111-4111-8111-111111111111",
  requestId: "d1111111-1111-4111-8111-111111111111",
  correlationId: "c1111111-1111-4111-8111-111111111111",
  riskEventType: "leakedCredentials",
  riskState: "atRisk",
  riskLevel: "high",
  riskDetail: "leakedCredentials",
  source: "IdentityProtection",
  detectionTimingType: "realtime",
  activity: "signin",
  tokenIssuerType: "AzureAD",
  ipAddress: "131.107.159.37",
  location: { city: "Seattle", state: "Washington", countryOrRegion: "US", geoCoordinates: null },
  activityDateTime: "2026-09-10T08:00:00Z",
  detectedDateTime: "2026-09-10T08:11:27Z",
  lastUpdatedDateTime: "2026-09-10T08:11:27Z",
  userId: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  userDisplayName: "Adele Vance",
  userPrincipalName: "AdeleV@contoso.com",
  additionalInfo: LONG_INFO,
};
const d2 = {
  id: "d2222222-2222-4222-8222-222222222222",
  requestId: "d2222222-2222-4222-8222-222222222222",
  correlationId: "c2222222-2222-4222-8222-222222222222",
  riskEventType: "generic",
  riskState: "atRisk",
  riskLevel: "medium",
  riskDetail: "adminConfirmedUserCompromised",
  source: "IdentityProtection",
  detectionTimingType: "nearRealtime",
  activity: "signin",
  ipAddress: "131.107.159.38",
  location: { city: "Redmond", state: "Washington", countryOrRegion: "US", geoCoordinates: null },
  activityDateTime: "2026-09-11T09:30:00Z",
  detectedDateTime: "2026-09-11T09:41:00Z",
  lastUpdatedDateTime: "2026-09-11T09:41:00Z",
  userId: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  userDisplayName: "Alex Wilber",
  userPrincipalName: "AlexW@contoso.com",
  additionalInfo: "[]",
};
const d3 = {
  id: "d3333333-3333-4333-8333-333333333333",
  requestId: null,
  correlationId: null,
  riskEventType: "unfamiliarFeatures",
  riskState: "dismissed",
  riskLevel: "low",
  riskDetail: "userPerformedSecuredPasswordChange",
  source: "IdentityProtection",
  detectionTimingType: "offline",
  activity: "user",
  ipAddress: null,
  location: null,
  activityDateTime: "2026-09-12T10:00:00Z",
  detectedDateTime: "2026-09-12T11:00:00Z",
  lastUpdatedDateTime: "2026-09-12T11:00:00Z",
  userId: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  userDisplayName: "Grady Archie",
  userPrincipalName: "GradyA@contoso.com",
  additionalInfo: "[]",
};
const riskDetections = [d1, d2, d3];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-06-"));
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

function riskTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/identityProtection/riskyUsers") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [u3] });
      return json(200, {
        value: [u1, u2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identityProtection/riskyUsers?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/identityProtection/riskDetections") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [d3] });
      return json(200, {
        value: [d1, d2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identityProtection/riskDetections?%24skiptoken=page2",
      });
    }
    const user = riskyUsers.find(row => path === `/v1.0/identityProtection/riskyUsers/${row.id}`);
    if (user) return json(200, user);
    const detection = riskDetections.find(row => path === `/v1.0/identityProtection/riskDetections/${row.id}`);
    if (detection) return json(200, detection);
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? riskTransport();
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
    "--import", pathToFileURL(resolve("test/fixtures/read-risk-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["pipe", "pipe", "pipe"], input, timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, riskyUsers, riskDetections, denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} executable lists risky users with compact rows`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "risky-user", "list", "--profile", profile, "--limit", "10"], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.riskyUsers, [
        { id: u1.id, userPrincipalName: u1.userPrincipalName, riskLevel: u1.riskLevel, riskState: u1.riskState },
        { id: u2.id, userPrincipalName: u2.userPrincipalName, riskLevel: u2.riskLevel, riskState: u2.riskState },
        { id: u3.id, userPrincipalName: u3.userPrincipalName, riskLevel: u3.riskLevel, riskState: u3.riskState },
      ]);
      assert.deepEqual(output.count, { returned: 3, complete: true });
      assert.ok(output.help.some(hint => hint.includes("entra risky-user show --id <risky-user-id>")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable lists bounded risk detections with compact rows`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "risk-detection", "list", "--profile", profile, "--since", SINCE], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.riskDetections, [
        { id: d1.id, detectedDateTime: d1.detectedDateTime, userPrincipalName: d1.userPrincipalName, riskLevel: d1.riskLevel },
        { id: d2.id, detectedDateTime: d2.detectedDateTime, userPrincipalName: d2.userPrincipalName, riskLevel: d2.riskLevel },
        { id: d3.id, detectedDateTime: d3.detectedDateTime, userPrincipalName: d3.userPrincipalName, riskLevel: d3.riskLevel },
      ]);
      assert.deepEqual(output.count, { returned: 3, complete: true });
      assert.ok(output.help.some(hint => hint.includes("entra risk-detection show --id <risk-detection-id>")));
      assert.ok(output.help.some(hint => hint.includes("entra sign-in list")));
      assert.ok(output.help.some(hint => hint.includes("riskySignIns")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });
}

test("delegated executable shows one risky user with detection and sign-in correlation", () => {
  const state = setupProfiles();
  try {
    const result = runReadCli(["entra", "risky-user", "show", "--profile", "soc", "--id", u1.id], state, "delegated");
    assert.equal(result.status, 0, result.stdout);
    const output = decode(result.stdout);
    assert.deepEqual(output.riskyUser, u1);
    assert.ok(output.help.some(hint => hint.includes("entra risk-detection list")));
    assert.ok(output.help.some(hint => hint.includes("entra sign-in list")));
    assert.ok(output.help.some(hint => hint.includes("never confirms, dismisses or remediates")));
  } finally { teardownProfiles(state); }
});

test("delegated executable shows one risk detection with sign-in correlation and limited-view guidance", () => {
  const state = setupProfiles();
  try {
    const result = runReadCli(["entra", "risk-detection", "show", "--profile", "soc", "--id", d1.id], state, "delegated");
    assert.equal(result.status, 0, result.stdout);
    const output = decode(result.stdout);
    assert.equal(output.riskDetection.id, d1.id);
    assert.equal(output.riskDetection.correlationId, d1.correlationId);
    assert.ok(!Object.hasOwn(output.riskDetection, "tokenIssuerType"));
    assert.match(output.riskDetection.additionalInfo, /truncated/);
    assert.ok(output.help[0].includes("--full"));
    assert.ok(output.help.some(hint => hint.includes("entra sign-in list")));
    assert.ok(output.help.some(hint => hint.includes("generic")));
  } finally { teardownProfiles(state); }
});

test("detection show without an associated sign-in says correlation does not apply", () => {
  const state = setupProfiles();
  try {
    const result = runReadCli(["entra", "risk-detection", "show", "--profile", "soc", "--id", d3.id], state, "delegated");
    assert.equal(result.status, 0, result.stdout);
    const output = decode(result.stdout);
    assert.equal(output.riskDetection.correlationId, null);
    assert.ok(output.help.some(hint => hint.includes("no associated sign-in")));
  } finally { teardownProfiles(state); }
});

for (const noun of ["risky-user", "risk-detection"]) {
  test(`delegated executable denied ${noun} list is an operational error on stdout`, () => {
    const state = setupProfiles();
    try {
      const args = noun === "risky-user"
        ? ["entra", "risky-user", "list", "--profile", "soc"]
        : ["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE];
      const result = runReadCli(args, state, "delegated", true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /\(403\)/);
      assert.ok(output.help.some(hint => hint.includes("Global Reader")));
    } finally { teardownProfiles(state); }
  });
}

test("delegated executable resumes a capped detection list through a stdin cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "risk-detection", "list", "--profile", "soc",
      "--since", SINCE, "--limit", "2"], overrides);
    assert.deepEqual(first.riskDetections.map(row => row.id), [d1.id, d2.id]);
    assert.equal(first.count.complete, false);
    const resumed = runReadCli(["entra", "risk-detection", "list", "--profile", "soc", "--cursor", "-"],
      state, "delegated", false, first.cursor);
    assert.equal(resumed.status, 0, resumed.stdout);
    assert.equal(resumed.stderr, "");
    const output = decode(resumed.stdout);
    assert.deepEqual(output.riskDetections.map(row => row.id), [d3.id]);
    assert.deepEqual(output.count, { returned: 1, complete: true });
  } finally { teardownProfiles(state); }
});

test("delegated executable resumes a capped risky-user list through a stdin cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "risky-user", "list", "--profile", "soc", "--limit", "1"], overrides);
    assert.deepEqual(first.riskyUsers.map(row => row.id), [u1.id]);
    assert.equal(first.count.complete, false);
    const resumed = runReadCli(["entra", "risky-user", "list", "--profile", "soc", "--cursor", "-"],
      state, "delegated", false, first.cursor);
    assert.equal(resumed.status, 0, resumed.stdout);
    const output = decode(resumed.stdout);
    assert.deepEqual(output.riskyUsers.map(row => row.id), [u2.id, u3.id]);
    assert.deepEqual(output.count, { returned: 2, complete: true });
  } finally { teardownProfiles(state); }
});

for (const [input, expectedError] of [
  ["", /needs the opaque cursor/],
  ["invalid", /Invalid collection cursor/],
  ["a".repeat(MAX_CURSOR_BYTES + 1), /cursor exceeds 16000000 bytes/],
]) {
  test(`risk-detection executable rejects ${input.length > MAX_CURSOR_BYTES ? "oversized" : input || "empty"} stdin cursors`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "risk-detection", "list", "--profile", "soc", "--cursor", "-"],
        state, "delegated", true, input);
      assert.equal(result.status, 2, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "VALIDATION_ERROR");
      assert.match(output.error, expectedError);
    } finally { teardownProfiles(state); }
  });
}

test("risk stdin cursor rejects another collection's binding", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE, "--limit", "1"], overrides);
    const result = runReadCli(["entra", "risky-user", "list", "--profile", "soc", "--cursor", "-"],
      state, "delegated", true, first.cursor);
    assert.equal(result.status, 2, result.stdout);
    assert.equal(decode(result.stdout).code, "VALIDATION_ERROR");
  } finally { teardownProfiles(state); }
});

test("delegated detection list composes since, until and filter into one bounded $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "risk-detection", "list", "--profile", "soc",
      "--since", SINCE, "--until", UNTIL, "--filter", "riskState eq 'atRisk'", "--all"], overrides);
    assert.equal(new URL(requests[0].url).searchParams.get("$filter"),
      "detectedDateTime ge 2026-09-01T00:00:00.000Z and detectedDateTime le 2026-09-08T00:00:00.000Z and (riskState eq 'atRisk')");
    assert.deepEqual(result.count, { returned: 3, complete: true });
    assert.equal(requests.length, 2);
  } finally { teardownProfiles(state); }
});

test("delegated risky-user list passes a plain filter with no time bound", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "risky-user", "list", "--profile", "soc",
      "--filter", "riskState eq 'atRisk'"], overrides);
    assert.equal(new URL(requests[0].url).searchParams.get("$filter"), "riskState eq 'atRisk'");
    assert.deepEqual(result.count, { returned: 3, complete: true });
  } finally { teardownProfiles(state); }
});

for (const args of [["--until", UNTIL], []]) {
  test(`detection list without --since fails before HTTP${args.length ? " with only --until" : ""}`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor("delegated");
      await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "soc", ...args], overrides),
        { code: "VALIDATION_ERROR" });
      assert.deepEqual(requests, []);
      assert.deepEqual(calls, []);
    } finally { teardownProfiles(state); }
  });
}

for (const args of [["--since", "not-a-time"], ["--since", UNTIL, "--until", SINCE]]) {
  test(`invalid detection time bounds fail before HTTP: ${args.join(" ")}`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor("delegated");
      await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "soc", ...args], overrides),
        { code: "VALIDATION_ERROR" });
      assert.deepEqual(requests, []);
      assert.deepEqual(calls, []);
    } finally { teardownProfiles(state); }
  });
}

for (const filter of [undefined, "riskState eq 'atRisk'", `detectedDateTime le 2026-09-08T00:00:00.000Z`]) {
  test(`detection list rejects raw cursors without a mandatory lower time bound: ${filter ?? "no filter"}`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor("delegated");
      const raw = await runApiGet({ path: "/identityProtection/riskDetections", apiVersion: "v1.0",
        profile: new Profiles().resolve("soc").profile,
        scopes: "https://graph.microsoft.com/IdentityRiskEvent.Read.All", limit: 1,
        odata: `$select=id${filter === undefined ? "" : `&$filter=${filter}`}` }, overrides);
      assert.equal(raw.complete, false);
      const requestCount = requests.length;
      const credentialCount = calls.length;
      await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--cursor", raw.cursor], overrides), error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        assert.match(error.message, /bounded in time/);
        assert.ok(error.suggestions.some(hint => hint.includes("Start a new query with --since")));
        return true;
      });
      await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--cursor", raw.cursor,
        "--since", SINCE], overrides), { code: "VALIDATION_ERROR" });
      assert.equal(requests.length, requestCount);
      assert.equal(calls.length, credentialCount);
    } finally { teardownProfiles(state); }
  });
}

test("detection resume repeats identical time bounds but rejects a conflicting --since", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "risk-detection", "list", "--profile", "soc",
      "--since", SINCE, "--until", UNTIL, "--filter", "riskState eq 'atRisk'", "--limit", "1"], overrides);
    const resumed = await executeArgv(["entra", "risk-detection", "list", "--profile", "soc",
      "--since", SINCE, "--until", UNTIL, "--filter", "riskState eq 'atRisk'", "--cursor", first.cursor], overrides);
    assert.deepEqual(resumed.riskDetections.map(row => row.id), [d2.id, d3.id]);
    await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "soc",
      "--since", UNTIL, "--cursor", first.cursor], overrides), { code: "VALIDATION_ERROR" });
  } finally { teardownProfiles(state); }
});

test("a capped risky-user list resumes losslessly through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "risky-user", "list", "--profile", "soc", "--limit", "2"], overrides);
    assert.deepEqual(first.riskyUsers.map(row => row.id), [u1.id, u2.id]);
    assert.equal(first.count.complete, false);
    const resumed = await executeArgv(["entra", "risky-user", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(resumed.riskyUsers.map(row => row.id), [u3.id]);
    assert.deepEqual(resumed.count, { returned: 1, complete: true });
  } finally { teardownProfiles(state); }
});

test("risky-user list rejects a detection collection cursor binding", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE, "--limit", "1"], overrides);
    await assert.rejects(executeArgv(["entra", "risky-user", "list", "--profile", "soc", "--cursor", first.cursor], overrides),
      { code: "VALIDATION_ERROR" });
  } finally { teardownProfiles(state); }
});

test("an empty risk list is a definitive zero with honest guidance", async () => {
  const state = setupProfiles();
  try {
    const empty = transport(() => json(200, { value: [] }));
    const { overrides } = overridesFor("delegated", empty);
    const users = await executeArgv(["entra", "risky-user", "list", "--profile", "soc"], overrides);
    assert.deepEqual(users.riskyUsers, []);
    assert.deepEqual(users.count, { returned: 0, complete: true });
    assert.ok(users.help.some(hint => hint.includes("0 risky users matched")));
    assert.ok(users.help.some(hint => hint.includes("Limited results stay limited")));
    const detections = await executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE], overrides);
    assert.deepEqual(detections.count, { returned: 0, complete: true });
    assert.ok(detections.help.some(hint => hint.includes("0 risk detections matched")));
  } finally { teardownProfiles(state); }
});

test("denied risky-user reads name roles and the P2 requirement", async () => {
  const state = setupProfiles();
  try {
    const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants (403)" } }));
    const { overrides } = overridesFor("delegated", denied);
    await assert.rejects(executeArgv(["entra", "risky-user", "list", "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.ok(error.suggestions.some(hint => hint.includes("Security Administrator")));
      assert.ok(error.suggestions.some(hint => hint.includes("requires a Microsoft Entra ID P2 licence")));
      return true;
    });
  } finally { teardownProfiles(state); }
});

test("denied detection reads name roles and the P1-or-P2 boundary", async () => {
  const state = setupProfiles();
  try {
    const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants (403)" } }));
    const { overrides } = overridesFor("delegated", denied);
    await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.ok(error.suggestions.some(hint => hint.includes("Security Administrator")));
      assert.ok(error.suggestions.some(hint => hint.includes("P1 or P2")));
      return true;
    });
  } finally { teardownProfiles(state); }
});

test("premium detections keep the server's generic type instead of an invented one", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE,
      "--select", "id,riskEventType"], overrides);
    assert.deepEqual(result.riskDetections.map(row => row.riskEventType), ["leakedCredentials", "generic", "unfamiliarFeatures"]);
    assert.ok(result.help.some(hint => hint.includes("generic")));
  } finally { teardownProfiles(state); }
});

test("explicit nulls stay null and missing risk properties stay absent", async () => {
  const state = setupProfiles();
  try {
    const partial = transport(() => json(200, { id: d3.id, requestId: null, correlationId: null, ipAddress: null }));
    const { overrides } = overridesFor("delegated", partial);
    const shown = await executeArgv(["entra", "risk-detection", "show", "--profile", "soc", "--id", d3.id,
      "--select", "id,requestId,correlationId,ipAddress,source"], overrides);
    assert.deepEqual(shown.riskDetection, { id: d3.id, requestId: null, correlationId: null, ipAddress: null });
    assert.ok(!Object.hasOwn(shown.riskDetection, "source"));
  } finally { teardownProfiles(state); }
});

test("detection show truncates long text unless --full is passed", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const shown = await executeArgv(["entra", "risk-detection", "show", "--profile", "soc", "--id", d1.id,
      "--fields", "id,additionalInfo"], overrides);
    assert.match(shown.riskDetection.additionalInfo, /\.\.\. \(truncated, \d+ chars total\)/);
    assert.ok(shown.help[0].includes("--full"));
    const full = await executeArgv(["entra", "risk-detection", "show", "--profile", "soc", "--id", d1.id,
      "--fields", "id,additionalInfo", "--full"], overrides);
    assert.equal(full.riskDetection.additionalInfo, LONG_INFO);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally { teardownProfiles(state); }
});

test("detection list truncation recovery replays time bounds with --full", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const rows = [{ id: "e1", detectedDateTime: "2026-09-10T08:11:27Z", userPrincipalName: "AdeleV@contoso.com", riskLevel: "high", additionalInfo: "y".repeat(600) }];
    const wide = transport(() => json(200, { value: rows }));
    const { overrides: wideOverrides } = overridesFor("delegated", wide);
    const first = await executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE,
      "--select", "id,additionalInfo"], wideOverrides);
    assert.match(first.riskDetections[0].additionalInfo, /truncated/);
    const replayed = await executeArgv(hintArgv(first.help[0]), wideOverrides);
    assert.equal(replayed.riskDetections[0].additionalInfo, "y".repeat(600));
  } finally { teardownProfiles(state); }
});

test("application mode lists and shows risk without delegated scopes", async () => {
  const state = setupProfiles();
  try {
    const { calls, overrides } = overridesFor("application");
    const users = await executeArgv(["entra", "risky-user", "list", "--profile", "batch"], overrides);
    assert.equal(users.count.returned, 3);
    const shown = await executeArgv(["entra", "risk-detection", "show", "--profile", "batch", "--id", d2.id,
      "--fields", "id,riskEventType"], overrides);
    assert.deepEqual(shown.riskDetection, { id: d2.id, riskEventType: "generic" });
    assert.ok(calls.length > 0);
  } finally { teardownProfiles(state); }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(executeArgv(["entra", "risky-user", "list", "--profile", "batch",
      "--scopes", "https://graph.microsoft.com/IdentityRiskyUser.Read.All"], overrides), error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      return /Graph \.default audience/.test(error.message);
    });
    await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "batch", "--since", SINCE,
      "--scopes", "https://graph.microsoft.com/IdentityRiskEvent.Read.All"], overrides), { code: "VALIDATION_ERROR" });
    assert.deepEqual(requests, []);
  } finally {
    teardownProfiles(state);
  }
});

test("unknown risk properties and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "risky-user", "list", "--profile", "soc", "--select", "id,aboutMe"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE, "--fields", "riskState"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "risky-user", "show", "--profile", "soc", "--id", u1.id, "--select", "id,aboutMe"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "risk-detection", "show", "--profile", "soc", "--id", d1.id, "--select", "id,tokenIssuerType"], overrides), error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      assert.match(error.message, /Unknown property tokenIssuerType/);
      return true;
    });
    await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE, "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

for (const scopes of [",", "https://example.invalid/IdentityRiskEvent.Read.All", "https://graph.microsoft.com/.default"]) {
  test(`risk-detection list rejects delegated scopes ${scopes} before provider acquisition`, async () => {
    const state = setupProfiles();
    try {
      const calls = [];
      const { overrides, requests } = overridesFor("delegated");
      overrides.delegated = new DelegatedAuth({
        storage: "session-only",
        login: async () => { throw new Error("Unexpected interactive login"); },
        silent: async (...args) => { calls.push(args); throw new Error("Unexpected credential acquisition"); },
      });
      await assert.rejects(executeArgv(["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE, "--scopes", scopes], overrides), { code: "VALIDATION_ERROR" });
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
      ["entra", "risky-user", "list", "--profile", "soc", "--select", "id,riskState", "--all"],
      overrides,
    );
    assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,riskState");
    assert.deepEqual(richer.riskyUsers[0], { id: u1.id, riskState: "atRisk" });
    for (const argv of [
      ["entra", "risky-user", "list", "--profile", "soc", "--api-version", "beta", "--limit", "1"],
      ["entra", "risk-detection", "list", "--profile", "soc", "--since", SINCE, "--api-version", "beta", "--limit", "1"],
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
