import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const tenant = "33333333-3333-4333-8333-333333333333";
const client = "44444444-4444-4222-8222-444444444444";
const userScopes = ["https://graph.microsoft.com/User.Read.All"];
const groupScopes = ["https://graph.microsoft.com/Group.Read.All"];
const appScopes = ["https://graph.microsoft.com/Application.Read.All"];
const auScopes = ["https://graph.microsoft.com/AdministrativeUnit.Read.All"];

// Fixture shapes follow the documented v1.0 list/get responses: typed rows
// sharing the @odata.type discriminator. The application row carries
// credential collections with a secret value and the user row carries
// personal data; projection must drop the former and keep defaults minimal.
const deletedUser = {
  "@odata.type": "#microsoft.graph.user",
  id: "6ea91a8d-e32e-41a1-b7bd-d2d185eed0e0",
  displayName: "Adele Vance",
  userPrincipalName: "AdeleV@contoso.com",
  mail: "AdeleV@contoso.com",
  userType: "Member",
  deletedDateTime: "2026-09-20T16:56:36Z",
};
const deletedGroup = {
  "@odata.type": "#microsoft.graph.group",
  id: "46cc6179-19d0-473e-97ad-6ff84347bbbb",
  displayName: "SampleGroup",
  mail: "example@contoso.com",
  mailNickname: "Example",
  groupTypes: ["Unified"],
  visibility: "Public",
  securityEnabled: false,
  deletedDateTime: "2026-09-21T10:00:00Z",
};
const deletedApp = {
  "@odata.type": "#microsoft.graph.application",
  id: "9b1c2d3e-4f5a-4b6c-8d7e-8f9a0b1c2d3e",
  appId: "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d",
  displayName: "Daemon app",
  publisherDomain: "contoso.com",
  signInAudience: "AzureADMyOrg",
  deletedDateTime: "2026-09-22T10:00:00Z",
  keyCredentials: [{ keyId: "k1", displayName: "key", secretText: "super-secret-value" }],
  passwordCredentials: [{ keyId: "p1", displayName: "daemon secret", secretText: "super-secret-value" }],
};
const deletedServicePrincipal = {
  "@odata.type": "#microsoft.graph.servicePrincipal",
  id: "7d8e9f0a-1b2c-4d5e-8f9a-0b1c2d3e4f5a",
  appId: "b2c3d4e5-f6a7-4b8c-9d0e-1f2a3b4c5d6e",
  displayName: "Daemon SP",
  servicePrincipalType: "Application",
  accountEnabled: false,
  deletedDateTime: "2026-09-23T10:00:00Z",
};
const deletedAdministrativeUnit = {
  "@odata.type": "#microsoft.graph.administrativeUnit",
  id: "1a2b3c4d-5e6f-4a7b-8c9d-0e1f2a3b4c5d",
  displayName: "West region",
  description: "Western sales administration",
  visibility: "Public",
  membershipType: "Assigned",
  deletedDateTime: "2026-09-24T10:00:00Z",
};
const longName = `group-${"g".repeat(600)}`;
const longGroup = { ...deletedGroup, id: "group-long-1", displayName: longName };

const byType = {
  user: { rows: [deletedUser], scopes: userScopes, leaf: "deleted-user", listPath: "/v1.0/directory/deletedItems/graph.user" },
  group: { rows: [deletedGroup], scopes: groupScopes, leaf: "deleted-group", listPath: "/v1.0/directory/deletedItems/graph.group" },
  application: { rows: [deletedApp], scopes: appScopes, leaf: "deleted-application", listPath: "/v1.0/directory/deletedItems/graph.application" },
  "service-principal": { rows: [deletedServicePrincipal], scopes: appScopes, leaf: "deleted-service-principal", listPath: "/v1.0/directory/deletedItems/graph.servicePrincipal" },
  "administrative-unit": { rows: [deletedAdministrativeUnit], scopes: auScopes, leaf: "deleted-administrative-unit", listPath: "/v1.0/directory/deletedItems/graph.administrativeUnit" },
};
const allRows = [deletedUser, deletedGroup, deletedApp, deletedServicePrincipal, deletedAdministrativeUnit];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-deleted-items-"));
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
      calls.push([mode, ...args]);
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

function deletedItemsTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    const count = /^\/v1\.0\/directory\/deletedItems\/(graph\.[^/]+)\/\$count$/.exec(path);
    if (count) {
      const entry = Object.values(byType).find(candidate => candidate.listPath === `/v1.0/directory/deletedItems/${count[1]}`);
      if (entry) return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(entry.rows.length) };
    }
    for (const entry of Object.values(byType)) {
      if (path === entry.listPath) {
        if (url.searchParams.has("$skiptoken")) return json(200, { value: [] });
        return json(200, { value: entry.rows });
      }
    }
    const single = /^\/v1\.0\/directory\/deletedItems\/([^/]+)$/.exec(path);
    if (single) {
      const found = allRows.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such deleted object" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? deletedItemsTransport();
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

function runDeletedItemCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-deleted-items-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, groups: [deletedGroup], users: [deletedUser], denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  for (const [type, entry] of Object.entries(byType)) {
    test(`${mode} lists deleted ${type} rows with compact defaults and the kind discriminator`, async () => {
      const state = setupProfiles();
      try {
        const { requests, calls, overrides } = overridesFor(mode);
        const result = await executeArgv(["entra", entry.leaf, "list", "--profile", profile], overrides);
        assert.ok(Array.isArray(result.deletedItems));
        assert.equal(result.deletedItems.length, entry.rows.length);
        for (const row of result.deletedItems) {
          assert.ok(typeof row["@odata.type"] === "string", "rows name their kind");
          assert.ok(!JSON.stringify(row).includes("super-secret-value"), "secret values never reach output");
        }
        assert.deepEqual(result.count, { returned: entry.rows.length, complete: true });
        assert.ok(result.help.some(hint => hint.includes("entra deleted-item show --id <object-id>")));
        assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
        assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com${entry.listPath}?`));
        assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
        assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
        if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(entry.scopes)));
      } finally {
        teardownProfiles(state);
      }
    });

    test(`${mode} counts deleted ${type} as one scalar without a collection query`, async () => {
      const state = setupProfiles();
      try {
        const { requests, overrides } = overridesFor(mode);
        const result = await executeArgv(["entra", entry.leaf, "count", "--profile", profile], overrides);
        assert.deepEqual(result, { count: { returned: entry.rows.length, complete: true } });
        assert.equal(requests.length, 1);
        assert.equal(requests[0].url, `https://graph.microsoft.com${entry.listPath}/$count`);
      } finally {
        teardownProfiles(state);
      }
    });
  }

  test(`${mode} deleted-user list defaults stay minimal for personal data`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "deleted-user", "list", "--profile", profile], overrides);
      assert.deepEqual(result.deletedItems, [{
        id: deletedUser.id,
        displayName: deletedUser.displayName,
        deletedDateTime: deletedUser.deletedDateTime,
        "@odata.type": "#microsoft.graph.user",
      }]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} deleted-application list never projects credential collections`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "deleted-application", "list", "--profile", profile,
          "--select", "id,appId,displayName,keyCredentials,passwordCredentials"], overrides),
        error => {
          assert.equal(error.code, "VALIDATION_ERROR");
          return /Unknown deleted-item property keyCredentials/.test(error.message);
        },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
      const allowed = await executeArgv(["entra", "deleted-application", "list", "--profile", profile], overrides);
      assert.ok(!JSON.stringify(allowed).includes("super-secret-value"), "secret values never reach output");
      assert.equal(allowed.deletedItems[0].appId, deletedApp.appId);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} deleted-group list names the groupTypes kind limitation`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "deleted-group", "list", "--profile", profile,
        "--select", "id,displayName,groupTypes,securityEnabled,deletedDateTime"], overrides);
      assert.deepEqual(result.deletedItems, [{
        id: deletedGroup.id,
        displayName: deletedGroup.displayName,
        groupTypes: ["Unified"],
        securityEnabled: false,
        deletedDateTime: deletedGroup.deletedDateTime,
        "@odata.type": "#microsoft.graph.group",
      }]);
      assert.ok(result.help.some(hint => hint.includes("groupTypes")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} deleted-item show returns the cross-type safe set with its kind`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const args = ["entra", "deleted-item", "show", "--id", deletedUser.id, "--profile", profile];
      if (mode === "delegated") args.push("--scopes", userScopes[0]);
      const result = await executeArgv(args, overrides);
      assert.deepEqual(result.deletedItem, {
        id: deletedUser.id,
        displayName: deletedUser.displayName,
        deletedDateTime: deletedUser.deletedDateTime,
        "@odata.type": "#microsoft.graph.user",
      });
      assert.ok(result.help.some(hint => hint.includes("no default scope")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} deleted-item show refuses unreviewed cross-type properties before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const args = ["entra", "deleted-item", "show", "--id", deletedUser.id, "--profile", profile,
        "--select", "id,appId"];
      if (mode === "delegated") args.push("--scopes", userScopes[0]);
      await assert.rejects(executeArgv(args, overrides), error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Unknown deleted-item property appId/.test(error.message);
      });
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} deleted-user list refuses --filter before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "deleted-user", "list", "--profile", profile, "--filter", "id eq 'x'"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped deleted-group list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const paged = transport(request => {
        const url = new URL(request.url);
        if (url.searchParams.has("$skiptoken")) return json(200, { value: [] });
        return json(200, {
          value: [deletedGroup],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/directory/deletedItems/graph.group?%24skiptoken=page2",
        });
      });
      const { overrides } = overridesFor(mode, paged);
      const first = await executeArgv(["entra", "deleted-group", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.deletedItems.map(row => row.id), [deletedGroup.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "deleted-group", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.deletedItems, []);
      assert.deepEqual(second.count, { returned: 0, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty deleted-user lists stay definitive`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "deleted-user", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.deletedItems, []);
      assert.deepEqual(listed.count, { returned: 0, complete: true });
      assert.ok(listed.help.some(hint => hint.includes("absence of results is the answer")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown deleted-item ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const args = ["entra", "deleted-item", "show", "--id", "object-missing", "--profile", profile];
      if (mode === "delegated") args.push("--scopes", userScopes[0]);
      await assert.rejects(executeArgv(args, overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied deleted-group reads surface scope and personal-account guidance`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "deleted-group", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Group.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose role or licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "deleted-group", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Group.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts deleted-item reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runDeletedItemCli(["entra", "deleted-group", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.deletedItems.map(row => row.id), [deletedGroup.id]);
      assert.deepEqual(listOut.deletedItems.map(row => row["@odata.type"]), ["#microsoft.graph.group"]);
      assert.deepEqual(listOut.count, { returned: 1, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const showArgs = ["entra", "deleted-item", "show", "--id", deletedUser.id, "--profile", profile];
      if (mode === "delegated") showArgs.push("--scopes", userScopes[0]);
      const shown = runDeletedItemCli(showArgs, state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).deletedItem, {
        id: deletedUser.id,
        displayName: deletedUser.displayName,
        deletedDateTime: deletedUser.deletedDateTime,
        "@odata.type": "#microsoft.graph.user",
      });

      const counted = runDeletedItemCli(["entra", "deleted-user", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).count, { returned: 1, complete: true });
      assert.ok(!counted.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied deleted-item reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runDeletedItemCli(["entra", "deleted-group", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.deletedItems, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["deleted-user", "list"], ["--limit", "1"]],
    [["deleted-group", "count"], []],
    [["deleted-item", "show"], ["--id", deletedUser.id, "--scopes", "https://graph.microsoft.com/User.Read.All"]],
  ]) {
    for (const preview of [false, true]) {
      test(`${mode} ${command.join(" ")} refuses beta before credentials with preview=${preview}`, async () => {
        const state = setupProfiles();
        try {
          const path = join(state.dir, "config.json");
          const config = JSON.parse(readFileSync(path, "utf8"));
          config.profiles[profile].preview = preview;
          writeFileSync(path, JSON.stringify(config));
          const { calls, requests, overrides } = overridesFor(mode);
          await assert.rejects(
            executeArgv(["entra", ...command, ...args, "--profile", profile, "--api-version", "beta"], overrides),
            { code: "VALIDATION_ERROR", message: "Deleted-item reads support v1.0 only; beta needs its own review" },
          );
          assert.equal(calls.length, 0);
          assert.equal(requests.length, 0);
        } finally {
          teardownProfiles(state);
        }
      });
    }
  }
}

test("delegated deleted-item show needs an explicit type-matching scope", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "deleted-item", "show", "--id", deletedUser.id, "--profile", "soc"], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /needs an explicit --scopes matching the object's type/.test(error.message);
      },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("application deleted-item reads refuse delegated scopes before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "deleted-user", "list", "--profile", "batch", "--scopes", userScopes[0]], overrides),
      { code: "VALIDATION_ERROR", message: "Application profiles use the configured Graph .default audience; delegated scopes are unavailable" },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("malformed deleted-item counts fail closed instead of inventing zero", async () => {
  const state = setupProfiles();
  try {
    const malformed = transport(() => ({ status: 200, headers: { "Content-Type": "text/plain" }, body: "not-a-count" }));
    const { overrides } = overridesFor("delegated", malformed);
    await assert.rejects(executeArgv(["entra", "deleted-user", "count", "--profile", "soc"], overrides),
      { code: "GRAPH_ERROR" });
  } finally {
    teardownProfiles(state);
  }
});

test("truncated deleted-group text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const longRow = transport(() => json(200, { value: [longGroup] }));
    const { overrides } = overridesFor("delegated", longRow);
    const partial = await executeArgv(["entra", "deleted-group", "list", "--profile", "soc",
      "--select", "id,displayName,deletedDateTime"], overrides);
    const truncated = partial.deletedItems.find(row => row.id === longGroup.id);
    assert.match(truncated.displayName, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "deleted-group", "list", "--profile", "soc", "--full",
      "--select", "id,displayName,deletedDateTime"], overrides);
    assert.equal(full.deletedItems.find(row => row.id === longGroup.id).displayName, longName);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("raw deleted-user reads keep the kind discriminator and refuse unreviewed fields", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const { ApplicationAuth } = await import("../dist/app-auth.js");
  const credentialCalls = [];
  const requests = [];
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async (...args) => { credentialCalls.push(args); return credential; } }),
    application: new ApplicationAuth({ storage: "session-only", acquire: async (...args) => { credentialCalls.push(args); return credential; } }),
    transport: async request => { requests.push(request); return { headers: {}, body: JSON.stringify({ value: [deletedUser] }) }; },
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const appProfile = { ...delegatedProfile, mode: "application", credentialRef: { provider: "federated", key: "55555555-5555-4555-8555-555555555555" } };
  for (const profile of [delegatedProfile, appProfile]) {
    const listed = await runApiGet({ path: "/directory/deletedItems/graph.user", apiVersion: "v1.0", profile,
      ...(profile.mode === "delegated" ? { scopes: userScopes[0] } : {}) }, deps);
    assert.equal(listed.returned, 1);
    assert.deepEqual(listed.value, [{
      id: deletedUser.id,
      displayName: deletedUser.displayName,
      deletedDateTime: deletedUser.deletedDateTime,
      "@odata.type": "#microsoft.graph.user",
    }]);
    await assert.rejects(
      runApiGet({ path: "/directory/deletedItems/graph.user", apiVersion: "v1.0", profile, odata: "$select=id,passwordProfile",
        ...(profile.mode === "delegated" ? { scopes: userScopes[0] } : {}) }, deps),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Unreviewed \$select field passwordProfile/.test(error.message);
      },
    );
  }
  assert.ok(requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-secret"));
});

test("raw deleted-item show keeps the discriminator without per-type detail", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async () => credential }),
    application: new (await import("../dist/app-auth.js")).ApplicationAuth({ storage: "session-only", acquire: async () => credential }),
    transport: async () => ({ status: 200, headers: {}, body: JSON.stringify(deletedApp) }),
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const shown = await runApiGet({ path: `/directory/deletedItems/${deletedApp.id}`, apiVersion: "v1.0", profile: delegatedProfile, scopes: appScopes[0] }, deps);
  assert.deepEqual(shown, {
    id: deletedApp.id,
    displayName: deletedApp.displayName,
    deletedDateTime: deletedApp.deletedDateTime,
    "@odata.type": "#microsoft.graph.application",
  });
});
