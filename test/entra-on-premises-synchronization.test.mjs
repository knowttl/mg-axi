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
const syncScopes = ["https://graph.microsoft.com/OnPremDirectorySynchronization.Read.All"];

// Fixture shapes follow the documented v1.0 example response; sync3 is a
// synthetic overlong id exercising the truncation-marker convention, and the
// bare row below exercises null/missing preservation.
const sync1 = {
  id: "12cce4b4-4ab8-40b7-be4d-f5d7742ec185",
  configuration: {
    accidentalDeletionPrevention: {
      synchronizationPreventionType: "enabledForCount",
      alertThreshold: 500,
    },
  },
  features: {
    passwordSyncEnabled: false,
    passwordWritebackEnabled: false,
    directoryExtensionsEnabled: false,
    quarantineUponUpnConflictEnabled: true,
    quarantineUponProxyAddressesConflictEnabled: true,
    softMatchOnUpnEnabled: true,
    cloudPasswordPolicyForPasswordSyncedUsersEnabled: false,
    userWritebackEnabled: false,
    deviceWritebackEnabled: false,
    groupWriteBackEnabled: false,
  },
};
const sync2 = {
  id: "22cce4b4-4ab8-40b7-be4d-f5d7742ec186",
  configuration: null,
  features: {
    passwordSyncEnabled: true,
    groupWriteBackEnabled: true,
  },
};
const sync3 = { id: "sync-bare-3" };
const synchronizations = [sync1, sync2, sync3];
const longId = `sync-${"s".repeat(600)}`;

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-onprem-sync-"));
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

function syncTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/directory/onPremisesSynchronization") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [sync3] });
      return json(200, {
        value: [sync1, sync2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/directory\/onPremisesSynchronization\/([^/]+)$/.exec(path);
    if (single) {
      const found = synchronizations.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such synchronization" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? syncTransport();
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

function runSyncCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-on-premises-synchronization-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, synchronizations: [sync1, sync2], denied }),
    },
  });
}

test("delegated lists synchronizations with the reviewed set preserving null and missing", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "on-premises-synchronization", "list", "--profile", "soc"], overrides);
    assert.deepEqual(result.synchronizations, [sync1, { id: sync2.id, configuration: null, features: sync2.features }, sync3]);
    assert.deepEqual(result.count, "3 synchronizations");
    assert.equal(result.total, null);
    assert.equal(result.complete, true);
    assert.ok(result.help.some(hint => hint.includes("entra on-premises-synchronization show --id <synchronization-id>")));
    assert.ok(requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-delegated-token"));
    assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization?"));
    const sent = new URL(requests[0].url).searchParams;
    assert.equal(sent.get("$select"), "id,configuration,features");
    assert.ok(!sent.has("$filter"));
    assert.ok(!JSON.stringify(result).includes("opaque-fixture-delegated-token"));
    assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(syncScopes)));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated resumes a capped synchronization list through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "on-premises-synchronization", "list", "--profile", "soc", "--limit", "1"], overrides);
    assert.deepEqual(first.synchronizations.map(row => row.id), [sync1.id]);
    assert.equal(first.complete, false);
    assert.equal(typeof first.cursor, "string");
    const second = await executeArgv(["entra", "on-premises-synchronization", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.synchronizations.map(row => row.id), [sync2.id, sync3.id]);
    assert.deepEqual(second.count, "2 synchronizations");
    assert.equal(second.total, null);
    assert.equal(second.complete, true);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated shows one synchronization with the full reviewed set", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "on-premises-synchronization", "show", "--id", sync1.id, "--profile", "soc"], overrides);
    assert.deepEqual(result.synchronization, sync1);
    const missing = await executeArgv(["entra", "on-premises-synchronization", "show", "--id", sync3.id, "--profile", "soc"], overrides);
    assert.deepEqual(missing.synchronization, { id: sync3.id });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated show projects --fields locally from the fetched --select set", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "on-premises-synchronization", "show", "--id", sync1.id, "--profile", "soc",
      "--select", "id,features", "--fields", "features"], overrides);
    assert.deepEqual(result.synchronization, { features: sync1.features });
    await assert.rejects(
      executeArgv(["entra", "on-premises-synchronization", "show", "--id", sync1.id, "--profile", "soc",
        "--select", "id", "--fields", "features"], overrides),
      { code: "VALIDATION_ERROR", message: "--fields features was not fetched; request it with --select" },
    );
  } finally {
    teardownProfiles(state);
  }
});

for (const id of ["sync(id='x')", "sync)(", "sync%28"]) {
  test(`delegated synchronization show rejects function-style identifier ${id} before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor("delegated");
      await assert.rejects(
        executeArgv(["entra", "on-premises-synchronization", "show", "--id", id, "--profile", "soc"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

test("delegated synchronization list has no --filter flag to misuse", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "on-premises-synchronization", "list", "--profile", "soc", "--filter", "id eq 'x'"], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /unknown flag --filter/.test(error.message);
      },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

for (const command of [
  ["on-premises-synchronization", "list"],
  ["on-premises-synchronization", "show", "--id", sync1.id],
]) {
  test(`application ${command.join(" ")} is refused before credentials without a supported application permission`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor("application");
      await assert.rejects(
        executeArgv(["entra", ...command, "--profile", "batch"], overrides),
        error => {
          assert.equal(error.code, "VALIDATION_ERROR");
          return /need a delegated profile; Graph documents no supported application permission/.test(error.message);
        },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

test("delegated empty synchronization lists stay definitive", async () => {
  const state = setupProfiles();
  try {
    const empty = transport(() => json(200, { value: [] }));
    const { overrides } = overridesFor("delegated", empty);
    const listed = await executeArgv(["entra", "on-premises-synchronization", "list", "--profile", "soc"], overrides);
    assert.deepEqual(listed.synchronizations, []);
    assert.deepEqual(listed.count, "0 synchronizations");
    assert.equal(listed.total, null);
    assert.equal(listed.complete, true);
    assert.ok(listed.help.some(hint => hint.includes("not configured for the tenant")));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated unknown synchronization ids report absence, not emptiness", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "on-premises-synchronization", "show", "--id", "synchronization-missing", "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
      return /not found or inaccessible \(404\)/.test(error.message);
    });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated denied synchronization reads surface scope, role and licensing", async () => {
  const state = setupProfiles();
  try {
    const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
    const { overrides } = overridesFor("delegated", denied);
    await assert.rejects(executeArgv(["entra", "on-premises-synchronization", "list", "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.ok(error.suggestions.some(hint => hint.includes("OnPremDirectorySynchronization.Read.All")));
      assert.ok(error.suggestions.some(hint => hint.includes("Global Administrator")));
      assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
      return /grant, role, licence/.test(error.message);
    });
    await assert.rejects(executeArgv(["entra", "on-premises-synchronization", "show", "--id", sync1.id, "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.ok(error.suggestions.some(hint => hint.includes("OnPremDirectorySynchronization.Read.All")));
      return /grant, role, licence/.test(error.message);
    });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated executable lists and shows synchronization reads", () => {
  const state = setupProfiles();
  try {
    const listed = runSyncCli(["entra", "on-premises-synchronization", "list", "--profile", "soc"], state, "delegated");
    assert.equal(listed.status, 0, listed.stdout);
    assert.equal(listed.stderr, "");
    const listOut = decode(listed.stdout);
    assert.deepEqual(listOut.synchronizations.map(row => row.id), [sync1.id, sync2.id]);
    assert.deepEqual(listOut.count, "2 synchronizations");
    assert.equal(listOut.total, null);
    assert.equal(listOut.complete, true);
    assert.ok(!listed.stdout.includes("opaque-fixture-delegated-token"));

    const shown = runSyncCli(["entra", "on-premises-synchronization", "show", "--id", sync1.id, "--profile", "soc"], state, "delegated");
    assert.equal(shown.status, 0, shown.stdout);
    assert.deepEqual(decode(shown.stdout).synchronization, sync1);
    assert.ok(!shown.stdout.includes("opaque-fixture-delegated-token"));
  } finally { teardownProfiles(state); }
});

test("delegated executable denied synchronization reads fail operationally on stdout", () => {
  const state = setupProfiles();
  try {
    const result = runSyncCli(["entra", "on-premises-synchronization", "list", "--profile", "soc"], state, "delegated", true);
    assert.equal(result.status, 1, result.stdout);
    assert.equal(result.stderr, "");
    const output = decode(result.stdout);
    assert.equal(output.code, "GRAPH_ERROR");
    assert.match(output.error, /grant, role, licence or policy/);
    assert.equal(output.synchronizations, undefined);
    assert.ok(!result.stdout.includes("opaque-fixture-delegated-token"));
  } finally { teardownProfiles(state); }
});

for (const [command, args] of [
  [["on-premises-synchronization", "list"], []],
  [["on-premises-synchronization", "show"], ["--id", sync1.id]],
]) {
  for (const preview of [false, true]) {
    test(`delegated ${command.join(" ")} refuses beta before credentials with preview=${preview}`, async () => {
      const state = setupProfiles();
      try {
        const path = join(state.dir, "config.json");
        const config = JSON.parse(readFileSync(path, "utf8"));
        config.profiles.soc.preview = preview;
        writeFileSync(path, JSON.stringify(config));
        const { calls, requests, overrides } = overridesFor("delegated");
        await assert.rejects(
          executeArgv(["entra", ...command, ...args, "--profile", "soc", "--api-version", "beta"], overrides),
          { code: "VALIDATION_ERROR", message: "On-premises-synchronization reads support v1.0 only; beta needs its own review" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }
}

test("truncated synchronization text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const longRow = transport(() => json(200, { value: [{ id: longId }] }));
    const { overrides } = overridesFor("delegated", longRow);
    const partial = await executeArgv(["entra", "on-premises-synchronization", "list", "--profile", "soc",
      "--select", "id"], overrides);
    const truncated = partial.synchronizations.find(row => row.id.startsWith("sync-"));
    assert.match(truncated.id, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "on-premises-synchronization", "list", "--profile", "soc", "--full",
      "--select", "id"], overrides);
    assert.equal(full.synchronizations.find(row => row.id.startsWith("sync-")).id, longId);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("raw on-premises-synchronization reads serve delegated callers and refuse application profiles", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const { ApplicationAuth } = await import("../dist/app-auth.js");
  const credentialCalls = [];
  const requests = [];
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async (...args) => { credentialCalls.push(args); return credential; } }),
    application: new ApplicationAuth({ storage: "session-only", acquire: async (...args) => { credentialCalls.push(args); return credential; } }),
    transport: async request => { requests.push(request); return { headers: {}, body: JSON.stringify({ value: [sync1] }) }; },
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const appProfile = { ...delegatedProfile, mode: "application", credentialRef: { provider: "federated", key: "55555555-5555-4555-8555-555555555555" } };
  const listed = await runApiGet({ path: "/directory/onPremisesSynchronization", apiVersion: "v1.0", profile: delegatedProfile, scopes: syncScopes[0] }, deps);
  assert.equal(listed.returned, 1);
  await assert.rejects(
    runApiGet({ path: "/directory/onPremisesSynchronization", apiVersion: "v1.0", profile: appProfile }, deps),
    error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      return /needs a delegated profile/.test(error.message);
    },
  );
  assert.equal(credentialCalls.length, 1);
  assert.equal(requests.length, 1);
});
