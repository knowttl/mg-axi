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
const caScopes = ["https://graph.microsoft.com/Policy.Read.All"];

const p1 = {
  id: "11111111-1111-4111-8111-aaaaaaaaaaaa",
  displayName: "CA001: Require MFA for admins",
  state: "enabled",
  createdDateTime: "2021-11-02T14:17:09Z",
  modifiedDateTime: "2024-01-03T20:07:59Z",
  conditions: {
    clientAppTypes: ["all"],
    applications: {
      includeApplications: ["All"],
      excludeApplications: [],
      includeUserActions: [],
      includeAuthenticationContextClassReferences: [],
      applicationFilter: null,
    },
    users: {
      includeUsers: [],
      excludeUsers: [],
      includeGroups: [],
      excludeGroups: ["eedad040-3722-4bcb-bde5-bc7c857f4983"],
      includeRoles: ["62e90394-69f5-4237-9190-012177145e10"],
      excludeRoles: [],
      includeGuestsOrExternalUsers: null,
      excludeGuestsOrExternalUsers: null,
    },
  },
  grantControls: { operator: "OR", builtInControls: ["mfa"] },
  sessionControls: null,
};
// A long application-filter rule exercises nested truncation: one string far
// past the truncation ceiling inside conditions.
const longRule = `device.deviceOwnership -eq "Company" and (${"device.displayName -startsWith \"CORP-\" or ".repeat(20)}device.isCompliant -eq true)`;
const p2 = {
  id: "22222222-2222-4222-8222-bbbbbbbbbbbb",
  displayName: "CA002: Compliant devices for Corp network",
  state: "enabledForReportingButNotEnforced",
  createdDateTime: "2023-05-01T00:00:00Z",
  modifiedDateTime: "2024-02-01T00:00:00Z",
  conditions: {
    clientAppTypes: ["browser", "mobileAppsAndDesktopClients"],
    applications: {
      includeApplications: ["All"],
      excludeApplications: [],
      includeUserActions: [],
      includeAuthenticationContextClassReferences: [],
      applicationFilter: { mode: "include", rule: longRule },
    },
    users: {
      includeUsers: ["cccccccc-cccc-4ccc-8ccc-cccccccccccc"],
      excludeUsers: [],
      includeGroups: [],
      excludeGroups: [],
      includeRoles: [],
      excludeRoles: [],
      includeGuestsOrExternalUsers: null,
      excludeGuestsOrExternalUsers: null,
    },
  },
  grantControls: { operator: "AND", builtInControls: ["compliantDevice", "domainJoinedDevice"] },
  sessionControls: { applicationEnforcedRestrictions: { isEnabled: true }, persistentBrowser: null, signInFrequency: null },
};
const p3 = {
  id: "33333333-3333-4333-8333-cccccccccccc",
  displayName: "CA003: Block legacy auth",
  state: "disabled",
};
const policies = [p1, p2, p3];
assert.ok(longRule.length > 500, "The fixture rule must exceed the truncation ceiling");

const l1 = {
  "@odata.type": "#microsoft.graph.ipNamedLocation",
  id: "44444444-4444-4444-8444-dddddddddddd",
  displayName: "Corporate HQ",
  createdDateTime: "2019-07-26T18:00:43Z",
  modifiedDateTime: "2020-07-26T18:00:43Z",
  isTrusted: true,
  ipRanges: [{ "@odata.type": "#microsoft.graph.iPv4CidrRange", cidrAddress: "203.0.113.0/24" }],
};
const l2 = {
  "@odata.type": "#microsoft.graph.countryNamedLocation",
  id: "55555555-5555-4555-8555-eeeeeeeeeeee",
  displayName: "Blocked countries",
  createdDateTime: "2021-01-01T00:00:00Z",
  modifiedDateTime: "2021-06-01T00:00:00Z",
  countriesAndRegions: ["BY", "RU"],
  includeUnknownCountriesAndRegions: true,
};
const locations = [l1, l2];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-03-"));
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

function caTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/identity/conditionalAccess/policies") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [p3] });
      return json(200, {
        value: [p1, p2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/identity/conditionalAccess/namedLocations") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [l2] });
      return json(200, {
        value: [l1],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations?%24skiptoken=page2",
      });
    }
    const policy = policies.find(row => path === `/v1.0/identity/conditionalAccess/policies/${row.id}`);
    if (policy) return json(200, policy);
    const location = locations.find(row => path === `/v1.0/identity/conditionalAccess/namedLocations/${row.id}`);
    if (location) return json(200, location);
    return json(404, { error: { code: "Request_ResourceNotFound", message: "no such object" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? caTransport();
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

function runCaCli(args, state, mode, denied = false, input = "") {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-ca-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", input, timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, policies, locations, denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  for (const [family, key, field, detail] of [
    ["policy", "policies", "conditions", { applications: { applicationFilter: { mode: "include", rule: "x".repeat(4096) } } }],
    ["named-location", "namedLocations", "ipRanges", [{ cidrAddress: "x".repeat(4096) }]],
  ]) {
    test(`${mode} executable resumes large ${family} cursors through stdin`, async () => {
      const state = setupProfiles();
      try {
        const rows = Array.from({ length: 40 }, (_, index) => ({ id: `row-${index}`, [field]: detail }));
        const { overrides } = overridesFor(mode, transport(() => json(200, { value: rows })));
        const command = ["entra", "conditional-access", family, "list", "--profile", profile];
        const first = await executeArgv([...command, "--limit", "1", "--select", `id,${field}`], overrides);
        assert.ok(Buffer.byteLength(first.cursor) > 128 * 1024);
        assert.ok(first.help.some(hint => hint.includes("--cursor -") && hint.includes("stdin")));

        const resumed = runCaCli([...command, "--cursor", "-", "--all", "--full"], state, mode, false, first.cursor);
        assert.equal(resumed.status, 0, resumed.stdout);
        assert.equal(resumed.stderr, "");
        const output = decode(resumed.stdout);
        assert.deepEqual(output[key], rows.slice(1));
        assert.deepEqual(output.count, { returned: 39, complete: true });

        const truncated = runCaCli([...command, "--cursor", "-", "--limit", "1"], state, mode, false, first.cursor);
        assert.equal(truncated.status, 0, truncated.stdout);
        const hints = decode(truncated.stdout).help;
        assert.ok(hints.some(hint => hint.includes("--full") && hint.includes("--cursor -")));
        assert.ok(hints.some(hint => hint.includes("original input cursor on stdin")));
        assert.ok(hints.every(hint => !hint.includes(first.cursor)));
        const replayed = runCaCli([...command, "--cursor", "-", "--limit", "1", "--full"], state, mode, false, first.cursor);
        assert.equal(replayed.status, 0, replayed.stdout);
        assert.deepEqual(decode(replayed.stdout)[key], [rows[1]]);
      } finally { teardownProfiles(state); }
    });
  }

  test(`${mode} lists policies with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "conditional-access", "policy", "list", "--profile", profile], overrides);
      assert.deepEqual(result.policies, [
        { id: p1.id, displayName: p1.displayName, state: "enabled" },
        { id: p2.id, displayName: p2.displayName, state: "enabledForReportingButNotEnforced" },
        { id: p3.id, displayName: p3.displayName, state: "disabled" },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("conditional-access policy show --id <policy-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a policy with the full reviewed condition and control set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "conditional-access", "policy", "show", "--id", p1.id, "--profile", profile], overrides);
      assert.deepEqual(result.policy, p1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} truncated policy detail carries a --full hint that restores it`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const truncated = await executeArgv(["entra", "conditional-access", "policy", "show", "--id", p2.id, "--profile", profile], overrides);
      const rule = truncated.policy.conditions.applications.applicationFilter.rule;
      assert.match(rule, /\.\.\. \(truncated, \d+ chars total\)/);
      assert.ok(truncated.help.some(hint => hint.includes("--full")));
      assert.ok(truncated.help.every(hint => !hint.includes(longRule)));
      const full = await executeArgv(["entra", "conditional-access", "policy", "show", "--id", p2.id, "--profile", profile, "--full"], overrides);
      assert.deepEqual(full.policy, p2);
      assert.equal(full.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists named locations preserving the location kind`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "conditional-access", "named-location", "list", "--profile", profile], overrides);
      assert.deepEqual(result.namedLocations, [
        { id: l1.id, displayName: "Corporate HQ", "@odata.type": "#microsoft.graph.ipNamedLocation" },
        { id: l2.id, displayName: "Blocked countries", "@odata.type": "#microsoft.graph.countryNamedLocation" },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("conditional-access named-location show --id <named-location-id>")));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations?"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a named location with absent properties staying absent`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const ip = await executeArgv(["entra", "conditional-access", "named-location", "show", "--id", l1.id, "--profile", profile], overrides);
      assert.deepEqual(ip.namedLocation, l1);
      const country = await executeArgv(["entra", "conditional-access", "named-location", "show", "--id", l2.id, "--profile", profile], overrides);
      assert.deepEqual(country.namedLocation, l2);
      assert.ok(!Object.hasOwn(country.namedLocation, "isTrusted"));
      assert.ok(!Object.hasOwn(country.namedLocation, "ipRanges"));
      assert.equal(country.namedLocation["@odata.type"], "#microsoft.graph.countryNamedLocation");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped policy list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "conditional-access", "policy", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.policies.map(row => row.id), [p1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "conditional-access", "policy", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.policies.map(row => row.id), [p2.id, p3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped named-location list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "conditional-access", "named-location", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.namedLocations, [
        { id: l1.id, displayName: l1.displayName, "@odata.type": l1["@odata.type"] },
      ]);
      assert.equal(first.count.returned, 1);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "conditional-access", "named-location", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.namedLocations, [
        { id: l2.id, displayName: l2.displayName, "@odata.type": l2["@odata.type"] },
      ]);
      assert.deepEqual(second.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied policy reads surface role and P1/P2 guidance`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const missing = transport(() => json(404, { error: { code: "Request_ResourceNotFound", message: "gone" } }));
      const overrides = {
        transport: denied.send,
        delegated: credentialService("delegated", []),
        application: credentialService("application", []),
      };
      await assert.rejects(executeArgv(["entra", "conditional-access", "policy", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.match(error.message, /grant, role, licence or policy/);
        const guidance = error.suggestions.join("\n");
        assert.match(guidance, /Conditional Access Administrator/);
        assert.match(guidance, /P1/);
        return true;
      });
      await assert.rejects(executeArgv(["entra", "conditional-access", "named-location", "list", "--profile", profile],
        { ...overrides, transport: denied.send }), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        const guidance = error.suggestions.join("\n");
        assert.match(guidance, /Policy\.Read\.All/);
        assert.match(guidance, /P1/);
        return true;
      });
      await assert.rejects(executeArgv(["entra", "conditional-access", "policy", "show", "--id", p1.id, "--profile", profile],
        { ...overrides, transport: missing.send }), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty policy results report absence instead of an error`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const result = await executeArgv(["entra", "conditional-access", "policy", "list", "--profile", profile], overrides);
      assert.deepEqual(result.policies, []);
      assert.deepEqual(result.count, { returned: 0, complete: true });
      assert.ok(result.help.some(hint => hint.includes("0 conditional-access policies matched")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty named-location results report absence instead of an error`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const result = await executeArgv(["entra", "conditional-access", "named-location", "list", "--profile", profile], overrides);
      assert.deepEqual(result.namedLocations, []);
      assert.deepEqual(result.count, { returned: 0, complete: true });
      assert.ok(result.help.some(hint => hint.includes("0 named locations matched")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists and shows policies and named locations`, () => {
    const state = setupProfiles();
    try {
      const listed = runCaCli(["entra", "conditional-access", "policy", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const policiesOut = decode(listed.stdout);
      assert.deepEqual(policiesOut.policies.map(policy => policy.id), [p1.id, p2.id, p3.id]);
      assert.deepEqual(policiesOut.count, { returned: 3, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runCaCli(["entra", "conditional-access", "policy", "show", "--id", p1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).policy, p1);

      const locationsOut = runCaCli(["entra", "conditional-access", "named-location", "list", "--profile", profile], state, mode);
      assert.equal(locationsOut.status, 0, locationsOut.stdout);
      const locationsDecoded = decode(locationsOut.stdout);
      assert.deepEqual(locationsDecoded.namedLocations.map(location => location.id), [l1.id, l2.id]);

      const locationShown = runCaCli(["entra", "conditional-access", "named-location", "show", "--id", l1.id, "--profile", profile], state, mode);
      assert.equal(locationShown.status, 0, locationShown.stdout);
      assert.deepEqual(decode(locationShown.stdout).namedLocation, l1);
      assert.ok(!locationShown.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied policy reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runCaCli(["entra", "conditional-access", "policy", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.policies, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });
}

test("delegated policy reads default to Policy.Read.All without extra scopes", async () => {
  const state = setupProfiles();
  try {
    const calls = [];
    const { overrides } = overridesFor("delegated", undefined, calls);
    await executeArgv(["entra", "conditional-access", "policy", "list", "--profile", "soc"], overrides);
    assert.deepEqual(calls[0][1], caScopes);
  } finally {
    teardownProfiles(state);
  }
});

test("application profiles reject delegated scopes on conditional-access reads", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("application");
    await assert.rejects(executeArgv(["entra", "conditional-access", "policy", "list", "--profile", "batch",
      "--scopes", caScopes[0]], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "conditional-access", "named-location", "show", "--id", l1.id,
      "--profile", "batch", "--scopes", caScopes[0]], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
    assert.equal(calls.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

for (const [family, noun, known] of [
  ["policy", "policy", "id,displayName,state,createdDateTime,modifiedDateTime,conditions,grantControls,sessionControls"],
  ["named-location", "named-location", "id,displayName,createdDateTime,modifiedDateTime,isTrusted,ipRanges,countriesAndRegions,includeUnknownCountriesAndRegions"],
]) {
  test(`unknown ${noun} properties fail before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor("delegated");
      const showId = family === "policy" ? p1.id : l1.id;
      await assert.rejects(executeArgv(["entra", "conditional-access", family, "list", "--profile", "soc",
        "--select", "id,templateId"], overrides), error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        assert.match(error.message, /Unknown policy property|Unknown named-location property/);
        assert.ok(error.suggestions.join("\n").includes(known.split(",")[0]));
        return true;
      });
      await assert.rejects(executeArgv(["entra", "conditional-access", family, "show", "--id", showId,
        "--profile", "soc", "--select", "id", "--fields", "id,state"], overrides), { code: "VALIDATION_ERROR" });
      assert.equal(requests.length, 0);
      assert.equal(calls.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

// READ-03 leftover collections: authentication-strength policies, each
// policy's combination configurations, authentication method modes and
// templates. Fixtures mirror the reviewed v1.0 resource shapes; the fake
// transport below serves only these routes.
const strengthScopes = ["https://graph.microsoft.com/Policy.Read.AuthenticationMethod"];
const s1 = {
  id: "66666666-6666-4666-8666-ffffffffffff",
  displayName: "Multifactor authentication strength",
  description: "Built-in multifactor authentication strength.",
  policyType: "builtIn",
  requirementsSatisfied: "mfa",
  allowedCombinations: ["fido2", "x509CertificateMultiFactor"],
  createdDateTime: "2022-01-01T00:00:00Z",
  modifiedDateTime: "2022-01-01T00:00:00Z",
};
const longStrengthDescription = `${"Custom strength requiring phishing-resistant combinations. ".repeat(12)}Applies to all privileged roles.`;
assert.ok(longStrengthDescription.length > 500, "The strength description must exceed the truncation ceiling");
const s2 = {
  id: "77777777-7777-4777-8777-000000000000",
  displayName: "Phishing-resistant MFA",
  description: longStrengthDescription,
  policyType: "custom",
  requirementsSatisfied: "mfa",
  allowedCombinations: ["fido2"],
  createdDateTime: "2023-03-01T00:00:00Z",
  modifiedDateTime: "2024-04-01T00:00:00Z",
};
const strengths = [s1, s2];
const c1 = {
  "@odata.type": "#microsoft.graph.fido2CombinationConfiguration",
  id: "88888888-8888-4888-8888-111111111111",
  appliesToCombinations: ["fido2"],
};
const c2 = {
  "@odata.type": "#microsoft.graph.x509CertificateCombinationConfiguration",
  id: "99999999-9999-4999-8999-222222222222",
  appliesToCombinations: ["x509CertificateSingleFactor", "x509CertificateMultiFactor"],
};
const combos = [c1, c2];
const m1 = { id: "aaaaaaaa-aaaa-4aaa-8aaa-333333333333", displayName: "FIDO2 security key", authenticationMethod: "fido2" };
const m2 = { id: "bbbbbbbb-bbbb-4bbb-8bbb-444444444444", displayName: "Certificate-based authentication", authenticationMethod: "x509Certificate" };
const modes = [m1, m2];
const t1 = {
  id: "cccccccc-cccc-4ccc-8ccc-555555555555",
  name: "Block legacy authentication",
  description: "Block legacy authentication endpoints that bypass Conditional Access.",
  scenarios: ["secureFoundation", "zeroTrust"],
  details: { conditions: { clientAppTypes: ["exchangeActiveSync", "other"] }, grantControls: { builtInControls: ["block"] } },
};
const longTemplateDescription = `${"Template for requiring compliant devices on privileged access. ".repeat(10)}Review before creating a policy.`;
assert.ok(longTemplateDescription.length > 500, "The template description must exceed the truncation ceiling");
const t2 = {
  id: "dddddddd-dddd-4ddd-8ddd-666666666666",
  name: "Require compliant devices for admins",
  description: longTemplateDescription,
  scenarios: ["protectAdmins"],
  details: { conditions: { users: { includeRoles: ["62e90394-69f5-4237-9190-012177145e10"] } }, grantControls: { builtInControls: ["compliantDevice"] } },
};
const templates = [t1, t2];

function caLeftoverTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/identity/conditionalAccess/authenticationStrength/policies") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [s2] });
      return json(200, {
        value: [s1],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identity/conditionalAccess/authenticationStrength/policies?%24skiptoken=page2",
      });
    }
    if (path === `/v1.0/identity/conditionalAccess/authenticationStrength/policies/${s1.id}/combinationConfigurations`) {
      return json(200, { value: combos });
    }
    if (path === "/v1.0/identity/conditionalAccess/authenticationStrength/authenticationMethodModes") {
      return json(200, { value: modes });
    }
    if (path === "/v1.0/identity/conditionalAccess/templates") {
      return json(200, { value: templates });
    }
    const strength = strengths.find(row => path === `/v1.0/identity/conditionalAccess/authenticationStrength/policies/${row.id}`);
    if (strength) return json(200, strength);
    const combo = combos.find(row => path === `/v1.0/identity/conditionalAccess/authenticationStrength/policies/${s1.id}/combinationConfigurations/${row.id}`);
    if (combo) return json(200, combo);
    const mode = modes.find(row => path === `/v1.0/identity/conditionalAccess/authenticationStrength/authenticationMethodModes/${row.id}`);
    if (mode) return json(200, mode);
    const template = templates.find(row => path === `/v1.0/identity/conditionalAccess/templates/${row.id}`);
    if (template) return json(200, template);
    return json(404, { error: { code: "Request_ResourceNotFound", message: "no such object" } });
  });
}

function leftoverOverrides(mode) {
  return overridesFor(mode, caLeftoverTransport());
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists auth-strength policies with compact rows and totals`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = leftoverOverrides(mode);
      const result = await executeArgv(["entra", "conditional-access", "auth-strength-policy", "list", "--profile", profile, "--all"], overrides);
      assert.deepEqual(result.authStrengthPolicies, [
        { id: s1.id, displayName: s1.displayName, policyType: "builtIn" },
        { id: s2.id, displayName: s2.displayName, policyType: "custom" },
      ]);
      assert.equal(result.total, null);
      assert.equal(result.count, "2 auth-strength policies");
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("conditional-access auth-strength-policy show --id <auth-strength-policy-id>")));
      assert.ok(requests[0].url.includes("%24select=id%2CdisplayName%2CpolicyType"));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows an auth-strength policy with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = leftoverOverrides(mode);
      const result = await executeArgv(["entra", "conditional-access", "auth-strength-policy", "show", "--id", s1.id, "--profile", profile], overrides);
      assert.deepEqual(result.authStrengthPolicy, s1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists combination configurations under a policy preserving kind`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = leftoverOverrides(mode);
      const result = await executeArgv(["entra", "conditional-access", "combination-configuration", "list",
        "--policy", s1.id, "--profile", profile], overrides);
      assert.deepEqual(result.combinationConfigurations, [
        { id: c1.id, "@odata.type": "#microsoft.graph.fido2CombinationConfiguration" },
        { id: c2.id, "@odata.type": "#microsoft.graph.x509CertificateCombinationConfiguration" },
      ]);
      assert.equal(result.count, "2 combination configurations");
      assert.equal(result.complete, true);
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/identity/conditionalAccess/authenticationStrength/policies/${s1.id}/combinationConfigurations?`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} combination reads without a policy fail before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = leftoverOverrides(mode);
      await assert.rejects(executeArgv(["entra", "conditional-access", "combination-configuration", "list", "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--policy is required/.test(error.message));
      await assert.rejects(executeArgv(["entra", "conditional-access", "combination-configuration", "show",
        "--id", c1.id, "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--policy is required/.test(error.message));
      assert.equal(requests.length, 0);
      assert.equal(calls.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a combination configuration with its combinations`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = leftoverOverrides(mode);
      const result = await executeArgv(["entra", "conditional-access", "combination-configuration", "show",
        "--policy", s1.id, "--id", c2.id, "--profile", profile], overrides);
      assert.deepEqual(result.combinationConfiguration, c2);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists auth-method modes with compact rows and totals`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = leftoverOverrides(mode);
      const result = await executeArgv(["entra", "conditional-access", "auth-method-mode", "list", "--profile", profile], overrides);
      assert.deepEqual(result.authMethodModes, [
        { id: m1.id, displayName: m1.displayName },
        { id: m2.id, displayName: m2.displayName },
      ]);
      assert.equal(result.count, "2 auth-method modes");
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("conditional-access auth-method-mode show --id <auth-method-mode-id>")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows an auth-method mode with its method`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = leftoverOverrides(mode);
      const result = await executeArgv(["entra", "conditional-access", "auth-method-mode", "show", "--id", m1.id, "--profile", profile], overrides);
      assert.deepEqual(result.authMethodMode, m1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists templates with compact rows and totals`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = leftoverOverrides(mode);
      const result = await executeArgv(["entra", "conditional-access", "template", "list", "--profile", profile], overrides);
      assert.deepEqual(result.templates, [
        { id: t1.id, name: t1.name },
        { id: t2.id, name: t2.name },
      ]);
      assert.equal(result.count, "2 templates");
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("conditional-access template show --id <template-id>")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} truncated template detail carries a --full hint that restores it`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = leftoverOverrides(mode);
      const truncated = await executeArgv(["entra", "conditional-access", "template", "show", "--id", t2.id, "--profile", profile], overrides);
      assert.match(truncated.template.description, /\.\.\. \(truncated, \d+ chars total\)/);
      assert.ok(truncated.help.some(hint => hint.includes("--full")));
      assert.ok(truncated.help.every(hint => !hint.includes(longTemplateDescription)));
      const full = await executeArgv(["entra", "conditional-access", "template", "show", "--id", t2.id, "--profile", profile, "--full"], overrides);
      assert.deepEqual(full.template, t2);
      assert.equal(full.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped auth-strength-policy list through its cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = leftoverOverrides(mode);
      const first = await executeArgv(["entra", "conditional-access", "auth-strength-policy", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.authStrengthPolicies.map(row => row.id), [s1.id]);
      assert.equal(first.count, "1 auth-strength policies shown, more available");
      assert.equal(first.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "conditional-access", "auth-strength-policy", "list",
        "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.authStrengthPolicies.map(row => row.id), [s2.id]);
      assert.equal(second.count, "1 auth-strength policies");
      assert.equal(second.complete, true);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied strength reads surface role and scope guidance`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const overrides = {
        transport: denied.send,
        delegated: credentialService("delegated", []),
        application: credentialService("application", []),
      };
      await assert.rejects(executeArgv(["entra", "conditional-access", "auth-strength-policy", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        const guidance = error.suggestions.join("\n");
        assert.match(guidance, /Policy\.Read\.AuthenticationMethod/);
        assert.match(guidance, /Conditional Access Administrator/);
        assert.match(guidance, /P1/);
        return true;
      });
      await assert.rejects(executeArgv(["entra", "conditional-access", "template", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        const guidance = error.suggestions.join("\n");
        assert.match(guidance, /Policy\.Read\.All/);
        assert.match(guidance, /P1/);
        return true;
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty auth-strength results report absence instead of an error`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const result = await executeArgv(["entra", "conditional-access", "auth-strength-policy", "list", "--profile", profile], overrides);
      assert.deepEqual(result.authStrengthPolicies, []);
      assert.equal(result.count, "0 auth-strength policies");
      assert.ok(result.help.some(hint => hint.includes("0 authentication-strength policies matched")));
    } finally {
      teardownProfiles(state);
    }
  });
}

test("delegated strength reads default to Policy.Read.AuthenticationMethod without extra scopes", async () => {
  const state = setupProfiles();
  try {
    const calls = [];
    const { overrides } = overridesFor("delegated", caLeftoverTransport(), calls);
    await executeArgv(["entra", "conditional-access", "auth-strength-policy", "list", "--profile", "soc"], overrides);
    assert.deepEqual(calls[0][1], strengthScopes);
  } finally {
    teardownProfiles(state);
  }
});

test("application profiles reject delegated scopes on leftover conditional-access reads", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = leftoverOverrides("application");
    await assert.rejects(executeArgv(["entra", "conditional-access", "auth-strength-policy", "list", "--profile", "batch",
      "--scopes", strengthScopes[0]], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "conditional-access", "template", "show", "--id", t1.id,
      "--profile", "batch", "--scopes", caScopes[0]], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
    assert.equal(calls.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

for (const [family, noun, known, showId] of [
  ["auth-strength-policy", "auth-strength-policy", "id,displayName,description,policyType,requirementsSatisfied,allowedCombinations,createdDateTime,modifiedDateTime", s1.id],
  ["combination-configuration", "combination-configuration", "id,appliesToCombinations", c1.id],
  ["auth-method-mode", "auth-method-mode", "id,displayName,authenticationMethod", m1.id],
  ["template", "template", "id,name,description,scenarios,details", t1.id],
]) {
  test(`unknown ${noun} properties fail before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = leftoverOverrides("delegated");
      const extra = family === "combination-configuration" ? ["--policy", s1.id] : [];
      await assert.rejects(executeArgv(["entra", "conditional-access", family, "list", "--profile", "soc",
        "--select", "id,templateId", ...extra], overrides), error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        assert.match(error.message, new RegExp(`Unknown ${noun} property`));
        assert.ok(error.suggestions.join("\n").includes(known.split(",")[0]));
        return true;
      });
      await assert.rejects(executeArgv(["entra", "conditional-access", family, "show", "--id", showId,
        "--profile", "soc", "--select", "id", "--fields", "id,templateId", ...extra], overrides), { code: "VALIDATION_ERROR" });
      assert.equal(requests.length, 0);
      assert.equal(calls.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}
