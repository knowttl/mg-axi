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
const deviceScopes = ["https://graph.microsoft.com/Device.Read.All"];
const auScopes = ["https://graph.microsoft.com/AdministrativeUnit.Read.All"];

const d1 = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  deviceId: "11111111-1111-4111-8111-111111111111",
  displayName: "CONTOSO-WIN10",
  operatingSystem: "Windows",
  operatingSystemVersion: "10.0.19045",
  trustType: "AzureAD",
  isCompliant: true,
  isManaged: true,
  accountEnabled: true,
  createdDateTime: "2024-01-01T00:00:00Z",
  approximateLastSignInDateTime: "2024-06-01T00:00:00Z",
  manufacturer: "Contoso",
  model: "Virtual",
};
const d2 = {
  id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  deviceId: "22222222-2222-4222-8222-222222222222",
  displayName: "CONTOSO-ANDROID",
  operatingSystem: "Android",
  operatingSystemVersion: "13.0",
  trustType: "AzureAD",
  isCompliant: false,
  isManaged: false,
  accountEnabled: false,
  createdDateTime: "2024-02-01T00:00:00Z",
  approximateLastSignInDateTime: null,
  manufacturer: null,
  model: null,
};
const d3 = {
  id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  displayName: "Stale device",
};
const devices = [d1, d2, d3];

const au1 = {
  id: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
  displayName: "Seattle Schools",
  description: "Seattle district schools administration",
  visibility: "Public",
  membershipType: "Assigned",
  membershipRule: null,
};
const au2 = {
  id: "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee",
  displayName: "US Sales",
  description: `${"Dynamic sales scope. ".repeat(30)}(${30 * 22} chars)`,
  visibility: "HiddenMembership",
  membershipType: "Dynamic",
  membershipRule: "(user.country -eq \"United States\")",
};
const au3 = {
  id: "ffffffff-ffff-4fff-8fff-ffffffffffff",
  displayName: "Niche",
};
const units = [au1, au2, au3];

const mUser = { "@odata.type": "#microsoft.graph.user", id: "11111111-1111-4111-8111-111111111112", displayName: "Adele Vance", mail: "adele@contoso.com" };
const mGroup = { "@odata.type": "#microsoft.graph.group", id: "22222222-2222-4222-8222-222222222223", displayName: "Sales" };
const mDevice = { "@odata.type": "#microsoft.graph.device", id: d1.id, displayName: "CONTOSO-WIN10" };
const mLimited = { "@odata.type": "#microsoft.graph.servicePrincipal", id: "33333333-3333-4333-8333-333333333334" };
const mNullName = { "@odata.type": "#microsoft.graph.user", id: "44444444-4444-4444-8444-444444444445", displayName: null };
const members = [mUser, mGroup, mDevice, mLimited, mNullName];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-10-"));
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

function directoryTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/devices") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [d3] });
      return json(200, {
        value: [d1, d2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/devices?%24skiptoken=page2",
      });
    }
    const singleDevice = /^\/v1\.0\/devices\/([^/]+)$/.exec(path);
    if (singleDevice) {
      const found = devices.find(device => device.id === decodeURIComponent(singleDevice[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such device" } });
    }
    if (path === "/v1.0/directory/administrativeUnits") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [au3] });
      return json(200, {
        value: [au1, au2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/directory/administrativeUnits?%24skiptoken=page2",
      });
    }
    if (/^\/v1\.0\/directory\/administrativeUnits\/[^/]+\/members$/.test(path)) {
      return json(200, { value: members });
    }
    const singleAu = /^\/v1\.0\/directory\/administrativeUnits\/([^/]+)$/.exec(path);
    if (singleAu) {
      const found = units.find(unit => unit.id === decodeURIComponent(singleAu[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such administrative unit" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? directoryTransport();
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

function runDirectoryCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-directory-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, devices, units, members, denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists devices with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "device", "list", "--profile", profile], overrides);
      assert.deepEqual(result.devices, [
        { id: d1.id, displayName: "CONTOSO-WIN10", operatingSystem: "Windows", accountEnabled: true },
        { id: d2.id, displayName: "CONTOSO-ANDROID", operatingSystem: "Android", accountEnabled: false },
        { id: d3.id, displayName: "Stale device" },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra device show --id <device-id>")));
      assert.ok(result.help.some(hint => hint.includes("Intune managed devices are a separately authorized surface")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/devices?"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a device with richer properties`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "device", "show", "--id", d1.id, "--profile", profile], overrides);
      assert.deepEqual(result.device, d1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists administrative units with compact rows`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "administrative-unit", "list", "--profile", profile,
        "--select", "id,displayName,visibility,membershipType"], overrides);
      assert.deepEqual(result.administrativeUnits, [
        { id: au1.id, displayName: "Seattle Schools", visibility: "Public", membershipType: "Assigned" },
        { id: au2.id, displayName: "US Sales", visibility: "HiddenMembership", membershipType: "Dynamic" },
        { id: au3.id, displayName: "Niche" },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra administrative-unit show --id <administrative-unit-id>")));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/directory/administrativeUnits?"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} show marks dynamic membership licensing without mutating`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const assigned = await executeArgv(["entra", "administrative-unit", "show", "--id", au1.id, "--profile", profile], overrides);
      assert.deepEqual(assigned.administrativeUnit, au1);
      assert.equal(assigned.help, undefined);
      const dynamic = await executeArgv(["entra", "administrative-unit", "show", "--id", au2.id, "--profile", profile,
        "--select", "id,displayName,membershipType"], overrides);
      assert.equal(dynamic.administrativeUnit.membershipType, "Dynamic");
      assert.ok(dynamic.help.some(hint => hint.includes("Dynamic membership needs additional P1 licensing")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists AU members with kinds, hidden and limited-information hints`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "administrative-unit", "member", "list",
        "--administrative-unit", au1.id, "--profile", profile, "--select", "id,displayName"], overrides);
      assert.deepEqual(result.members, [
        { "@odata.type": "#microsoft.graph.user", id: mUser.id, displayName: "Adele Vance" },
        mGroup,
        mDevice,
        { "@odata.type": "#microsoft.graph.servicePrincipal", id: mLimited.id },
        mNullName,
      ]);
      assert.deepEqual(result.count, { returned: 5, complete: true });
      assert.ok(result.help.some(hint => hint.includes("Member.Read.Hidden")));
      assert.ok(result.help.some(hint => hint.includes("completion describes pagination, not visibility")));
      assert.ok(result.help.some(hint => hint.includes("2 of 5 rows have no non-null selected descriptive properties")));
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/directory/administrativeUnits/${au1.id}/members`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped device list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "device", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.devices.map(row => row.id), [d1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "device", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.devices.map(row => row.id), [d2.id, d3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} states an empty device list as the answer`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const result = await executeArgv(["entra", "device", "list", "--profile", profile], overrides);
      assert.deepEqual(result.devices, []);
      assert.deepEqual(result.count, { returned: 0, complete: true });
      assert.ok(result.help.some(hint => hint.includes("0 devices matched; the absence of results is the answer, not an error")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied directory reads surface scope, roles and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const missing = transport(() => json(404, { error: { code: "Request_ResourceNotFound", message: "gone" } }));
      const overrides = {
        transport: denied.send,
        delegated: credentialService("delegated", []),
        application: credentialService("application", []),
      };
      await assert.rejects(executeArgv(["entra", "device", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Device.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Intune managed devices")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "administrative-unit", "member", "list",
        "--administrative-unit", au1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("AdministrativeUnit.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Member.Read.Hidden")));
        assert.ok(error.suggestions.some(hint => hint.includes("P1")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "device", "show", "--id", d1.id, "--profile", profile],
        { ...overrides, transport: missing.send }), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists devices and units, shows one of each and lists members`, () => {
    const state = setupProfiles();
    try {
      const listed = runDirectoryCli(["entra", "device", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const devicesOut = decode(listed.stdout);
      assert.deepEqual(devicesOut.devices.map(device => device.id), [d1.id, d2.id, d3.id]);
      assert.deepEqual(devicesOut.count, { returned: 3, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runDirectoryCli(["entra", "device", "show", "--id", d1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).device, d1);

      const unitsOut = runDirectoryCli(["entra", "administrative-unit", "list", "--profile", profile], state, mode);
      assert.equal(unitsOut.status, 0, unitsOut.stdout);
      const unitsDecoded = decode(unitsOut.stdout);
      assert.equal(unitsDecoded.administrativeUnits.length, 3);
      assert.ok(!unitsOut.stdout.includes(`opaque-fixture-${mode}-token`));

      const unitShown = runDirectoryCli(["entra", "administrative-unit", "show", "--id", au2.id, "--profile", profile], state, mode);
      assert.equal(unitShown.status, 0, unitShown.stdout);
      const unitDecoded = decode(unitShown.stdout);
      assert.equal(unitDecoded.administrativeUnit.membershipType, "Dynamic");
      assert.ok(unitDecoded.help.some(hint => hint.includes("P1 licensing")));

      const membersOut = runDirectoryCli(["entra", "administrative-unit", "member", "list",
        "--administrative-unit", au1.id, "--profile", profile], state, mode);
      assert.equal(membersOut.status, 0, membersOut.stdout);
      const membersDecoded = decode(membersOut.stdout);
      assert.equal(membersDecoded.members.length, 5);
      assert.ok(!membersOut.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied directory reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runDirectoryCli(["entra", "device", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.devices, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });
}

test("device list filtering carries count and consistency through resume", async () => {
  const state = setupProfiles();
  try {
    const seen = [];
    const fixture = transport(request => {
      seen.push({ url: request.url, headers: request.headers });
      assert.equal(new URL(request.url).searchParams.get("$count"), "true");
      if (new URL(request.url).searchParams.has("$skiptoken")) {
        return json(200, { value: [{ id: d2.id, displayName: "CONTOSO-ANDROID", operatingSystem: "Android", accountEnabled: false }] });
      }
      return json(200, {
        value: [{ id: d1.id, displayName: "CONTOSO-WIN10", operatingSystem: "Windows", accountEnabled: true }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/devices?%24count=true&%24skiptoken=next",
      });
    });
    const { overrides } = overridesFor("delegated", fixture);
    const filtered = await executeArgv(["entra", "device", "list",
      "--profile", "soc", "--filter", "isCompliant eq true", "--limit", "1"], overrides);
    assert.equal(seen[0].headers.ConsistencyLevel, "eventual");
    assert.equal(filtered.count.complete, false);
    const resumed = await executeArgv(["entra", "device", "list",
      "--profile", "soc", "--cursor", filtered.cursor], overrides);
    assert.equal(seen[1].headers.ConsistencyLevel, "eventual");
    assert.deepEqual(resumed.count.complete, true);
    await assert.rejects(executeArgv(["entra", "device", "list",
      "--profile", "soc", "--cursor", filtered.cursor, "--filter", "isCompliant eq false"], overrides),
    { code: "VALIDATION_ERROR" });
    const plain = transport(request => {
      seen.push({ url: request.url, headers: request.headers });
      return json(200, { value: [{ id: d1.id, displayName: "CONTOSO-WIN10", operatingSystem: "Windows", accountEnabled: true }] });
    });
    const unfiltered = await executeArgv(["entra", "device", "list", "--profile", "soc", "--limit", "1"],
      overridesFor("delegated", plain).overrides);
    assert.equal(seen[seen.length - 1].headers.ConsistencyLevel, undefined);
    assert.equal(new URL(seen[seen.length - 1].url).searchParams.has("$count"), false);
    assert.deepEqual(unfiltered.count, { returned: 1, complete: true });
  } finally {
    teardownProfiles(state);
  }
});

test("AU member list filtering carries count and consistency through resume", async () => {
  const state = setupProfiles();
  try {
    const seen = [];
    const fixture = transport(request => {
      seen.push({ url: request.url, headers: request.headers });
      assert.equal(new URL(request.url).searchParams.get("$count"), "true");
      return json(200, { value: [mUser] });
    });
    const { overrides } = overridesFor("delegated", fixture);
    const result = await executeArgv(["entra", "administrative-unit", "member", "list",
      "--administrative-unit", au1.id, "--profile", "soc", "--filter", "startswith(displayName,'A')"], overrides);
    assert.equal(seen[0].headers.ConsistencyLevel, "eventual");
    assert.deepEqual(result.members, [{ "@odata.type": "#microsoft.graph.user", id: mUser.id, displayName: "Adele Vance" }]);
    assert.deepEqual(result.count, { returned: 1, complete: true });
  } finally {
    teardownProfiles(state);
  }
});

test("truncated AU descriptions carry a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const partial = await executeArgv(["entra", "administrative-unit", "list", "--profile", "soc", "--limit", "2",
      "--select", "id,displayName,description,visibility,membershipType"], overrides);
    const dynamic = partial.administrativeUnits.find(row => row.id === au2.id);
    assert.match(dynamic.description, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "administrative-unit", "list", "--profile", "soc", "--limit", "2", "--full",
      "--select", "id,displayName,description,visibility,membershipType"], overrides);
    assert.equal(full.administrativeUnits.find(row => row.id === au2.id).description, au2.description);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "device", "list", "--profile", "batch", "--scopes", deviceScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "administrative-unit", "member", "list",
        "--administrative-unit", au1.id, "--profile", "batch", "--scopes", auScopes[0]], overrides),
      { code: "VALIDATION_ERROR" },
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
    await assert.rejects(executeArgv(["entra", "device", "list", "--profile", "soc", "--select", "id,aboutMe"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "device", "list", "--profile", "soc", "--fields", "deviceVersion"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "administrative-unit", "list", "--profile", "soc", "--select", "id,isMemberManagementRestricted"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "administrative-unit", "member", "list",
      "--administrative-unit", au1.id, "--profile", "soc", "--select", "id,department"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "device", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("directory read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "device", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "device", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "device", "list", "--transitive"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "administrative-unit", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "administrative-unit", "member", "list"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "administrative-unit", "member", "list", "--administrative-unit="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "administrative-unit", "member", "list",
    "--administrative-unit", au1.id, "--id", au1.id]), { code: "VALIDATION_ERROR" });
});

test("directory reads stay preview-gated on beta", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "device", "list", "--profile", "soc", "--api-version", "beta", "--limit", "1"], overrides),
      error => {
        assert.equal(error.code, "POLICY_DENIED");
        return /preview-enabled/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "administrative-unit", "member", "list",
        "--administrative-unit", au1.id, "--profile", "soc", "--api-version", "beta"], overrides),
      { code: "POLICY_DENIED" },
    );
  } finally {
    teardownProfiles(state);
  }
});
