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
const directoryObjectScopes = ["https://graph.microsoft.com/Directory.Read.All"];

// Fixture shapes follow the documented v1.0 get response: polymorphic rows
// sharing the base-type properties plus the @odata.type discriminator. The
// service-principal row carries credential metadata and a secret value, and
// the user row carries extra subtype properties; none of them may be
// projected because only base properties are reviewed here. obj3 is a bare
// row exercising null/missing preservation.
const obj1 = {
  "@odata.type": "#microsoft.graph.user",
  id: "6ea91a8d-e32e-41a1-b7bd-d2d185eed0e0",
  deletedDateTime: null,
  displayName: "Conf Room Adams",
  userPrincipalName: "Adams@Contoso.com",
};
const obj2 = {
  "@odata.type": "#microsoft.graph.servicePrincipal",
  id: "9b1c2d3e-4f5a-4b6c-8d7e-8f9a0b1c2d3e",
  passwordCredentials: [{ displayName: "daemon secret", secretText: "super-secret-value" }],
  appId: "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d",
};
const obj3 = { id: "obj-bare-3" };
const directoryObjects = [obj1, obj2, obj3];
const longId = `obj-${"o".repeat(600)}`;

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-directory-objects-"));
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

function directoryObjectTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/directoryObjects/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(directoryObjects.length) };
    }
    if (path === "/v1.0/directoryObjects") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [obj3] });
      return json(200, {
        value: [obj1, obj2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/directoryObjects?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/directoryObjects\/([^/]+)$/.exec(path);
    if (single) {
      const found = directoryObjects.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such directory object" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? directoryObjectTransport();
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

function runDirectoryObjectCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-directory-objects-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, directoryObjects: [obj1, obj2], denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists directory objects with compact rows naming each subtype`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-object", "list", "--profile", profile], overrides);
      assert.deepEqual(result.directoryObjects, [
        { id: obj1.id, "@odata.type": "#microsoft.graph.user" },
        { id: obj2.id, "@odata.type": "#microsoft.graph.servicePrincipal" },
        { id: obj3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra directory-object show --id <object-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/directoryObjects?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      assert.ok(!JSON.stringify(result).includes("super-secret-value"), "subtype secrets never reach output");
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(directoryObjectScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} directory-object list refuses --filter before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "directory-object", "list", "--profile", profile, "--filter", "id eq 'x'"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped directory-object list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "directory-object", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.directoryObjects.map(row => row.id), [obj1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "directory-object", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.directoryObjects.map(row => row.id), [obj2.id, obj3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one directory object with the base set and its subtype`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-object", "show", "--id", obj1.id, "--profile", profile], overrides);
      assert.deepEqual(result.directoryObject, { id: obj1.id, deletedDateTime: null, "@odata.type": "#microsoft.graph.user" });
      const missing = await executeArgv(["entra", "directory-object", "show", "--id", obj3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.directoryObject, { id: obj3.id });
      const secret = await executeArgv(["entra", "directory-object", "show", "--id", obj2.id, "--profile", profile], overrides);
      assert.deepEqual(secret.directoryObject, { id: obj2.id, "@odata.type": "#microsoft.graph.servicePrincipal" });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} directory-object show refuses unreviewed subtype properties before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "directory-object", "show", "--id", obj1.id, "--profile", profile, "--select", "id,displayName"], overrides),
        error => {
          assert.equal(error.code, "VALIDATION_ERROR");
          return /Unknown directory-object property displayName/.test(error.message);
        },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts directory objects as one scalar without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-object", "count", "--profile", profile], overrides);
      assert.deepEqual(result, { count: { returned: 3, complete: true } });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/directoryObjects/$count");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} directory-object count refuses collection flags before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "directory-object", "count", "--profile", profile, "--filter", "id eq 'x'"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty directory-object lists stay definitive`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "directory-object", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.directoryObjects, []);
      assert.deepEqual(listed.count, { returned: 0, complete: true });
      assert.ok(listed.help.some(hint => hint.includes("verify the profile tenant")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown directory-object ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "directory-object", "show", "--id", "object-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied directory-object reads surface scope, role and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "directory-object", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Directory.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose role or licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "directory-object", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Directory.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts directory-object reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runDirectoryObjectCli(["entra", "directory-object", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.directoryObjects.map(row => row.id), [obj1.id, obj2.id]);
      assert.deepEqual(listOut.directoryObjects.map(row => row["@odata.type"]), ["#microsoft.graph.user", "#microsoft.graph.servicePrincipal"]);
      assert.deepEqual(listOut.count, { returned: 2, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runDirectoryObjectCli(["entra", "directory-object", "show", "--id", obj1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).directoryObject, { id: obj1.id, deletedDateTime: null, "@odata.type": "#microsoft.graph.user" });

      const counted = runDirectoryObjectCli(["entra", "directory-object", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).count, { returned: 2, complete: true });
      assert.ok(!counted.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied directory-object reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runDirectoryObjectCli(["entra", "directory-object", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.directoryObjects, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["directory-object", "list"], ["--limit", "1"]],
    [["directory-object", "show"], ["--id", obj1.id]],
    [["directory-object", "count"], []],
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
            { code: "VALIDATION_ERROR", message: "Directory-object reads support v1.0 only; beta needs its own review" },
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

test("application directory-object reads refuse delegated scopes before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "directory-object", "list", "--profile", "batch", "--scopes", directoryObjectScopes[0]], overrides),
      { code: "VALIDATION_ERROR", message: "Application profiles use the configured Graph .default audience; delegated scopes are unavailable" },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("truncated directory-object text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const longRow = transport(() => json(200, { value: [{ "@odata.type": "#microsoft.graph.group", id: longId }] }));
    const { overrides } = overridesFor("delegated", longRow);
    const partial = await executeArgv(["entra", "directory-object", "list", "--profile", "soc",
      "--select", "id"], overrides);
    const truncated = partial.directoryObjects.find(row => row.id.startsWith("obj-"));
    assert.match(truncated.id, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "directory-object", "list", "--profile", "soc", "--full",
      "--select", "id"], overrides);
    assert.equal(full.directoryObjects.find(row => row.id.startsWith("obj-")).id, longId);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("raw directory-object reads keep the subtype discriminator and refuse unreviewed fields", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const { ApplicationAuth } = await import("../dist/app-auth.js");
  const credentialCalls = [];
  const requests = [];
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async (...args) => { credentialCalls.push(args); return credential; } }),
    application: new ApplicationAuth({ storage: "session-only", acquire: async (...args) => { credentialCalls.push(args); return credential; } }),
    transport: async request => { requests.push(request); return { headers: {}, body: JSON.stringify({ value: [obj1, obj2] }) }; },
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const appProfile = { ...delegatedProfile, mode: "application", credentialRef: { provider: "federated", key: "55555555-5555-4555-8555-555555555555" } };
  for (const profile of [delegatedProfile, appProfile]) {
    const listed = await runApiGet({ path: "/directoryObjects", apiVersion: "v1.0", profile,
      ...(profile.mode === "delegated" ? { scopes: directoryObjectScopes[0] } : {}) }, deps);
    assert.equal(listed.returned, 2);
    assert.deepEqual(listed.value, [
      { id: obj1.id, "@odata.type": "#microsoft.graph.user" },
      { id: obj2.id, "@odata.type": "#microsoft.graph.servicePrincipal" },
    ]);
    await assert.rejects(
      runApiGet({ path: "/directoryObjects", apiVersion: "v1.0", profile, odata: "$select=id,displayName",
        ...(profile.mode === "delegated" ? { scopes: directoryObjectScopes[0] } : {}) }, deps),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Unreviewed \$select field displayName/.test(error.message);
      },
    );
  }
  assert.ok(requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-secret"));
});

test("raw directory-object show keeps the discriminator without subtype secrets", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async () => credential }),
    application: new (await import("../dist/app-auth.js")).ApplicationAuth({ storage: "session-only", acquire: async () => credential }),
    transport: async () => ({ status: 200, headers: {}, body: JSON.stringify(obj2) }),
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const shown = await runApiGet({ path: `/directoryObjects/${obj2.id}`, apiVersion: "v1.0", profile: delegatedProfile, scopes: directoryObjectScopes[0] }, deps);
  assert.deepEqual(shown, { id: obj2.id, "@odata.type": "#microsoft.graph.servicePrincipal" });
});
