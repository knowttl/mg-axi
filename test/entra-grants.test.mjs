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
const sp = "33333333-3333-4333-8333-333333333333";
const user = "44444444-4444-4444-8444-444444444444";
const graphResource = "00000003-0000-0000-c000-000000000000";

const grantScopes = "https://graph.microsoft.com/Directory.Read.All";
const appRoleScopes = "https://graph.microsoft.com/Application.Read.All";

const g1 = {
  id: "g1-grant-all-principals",
  clientId: sp,
  consentType: "AllPrincipals",
  principalId: null,
  resourceId: graphResource,
  scope: `openid profile User.Read ${"x".repeat(600)}`,
};
const g2 = {
  id: "g2-grant-one-principal",
  clientId: sp,
  consentType: "Principal",
  principalId: user,
  resourceId: graphResource,
  scope: "User.Read",
};
const g3 = {
  id: "g3-grant-minimal",
  consentType: "AllPrincipals",
  resourceId: graphResource,
};
const grants = [g1, g2, g3];

const a1 = {
  id: "a1-app-role-live",
  appRoleId: "e2a3a72e-5f79-4c64-b1b1-878b674786c9",
  createdDateTime: "2021-02-02T04:22:45.4980259Z",
  principalDisplayName: "fixture-client",
  principalId: sp,
  principalType: "ServicePrincipal",
  resourceDisplayName: "Microsoft Graph",
  resourceId: graphResource,
};
const a2 = {
  id: "a2-app-role-minimal",
  appRoleId: "00000000-0000-0000-0000-000000000000",
  resourceDisplayName: "fixture-api",
  resourceId: "55555555-5555-4555-8555-555555555555",
};
const appRoles = [a1, a2];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-08-"));
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

function grantTransport(denied = false) {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (denied) return json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } });
    if (path === `/v1.0/servicePrincipals/${sp}/oauth2PermissionGrants`) {
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("appRoleId"));
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [g3] });
      return json(200, {
        value: [g1, g2],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/servicePrincipals/${sp}/oauth2PermissionGrants?%24skiptoken=page2`,
      });
    }
    if (path === `/v1.0/servicePrincipals/${sp}/appRoleAssignments`) {
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("scope"));
      return json(200, { value: appRoles });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? grantTransport();
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

function runGrantsCli(args, state, mode, scopes) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-grants-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, scopes, client: sp, grants, appRoles }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists delegated grants for the named client preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "service-principal", "oauth2-grant", "list",
        "--service-principal", sp, "--profile", profile], overrides);
      assert.deepEqual(result.oauth2PermissionGrants, [
        { id: g1.id, consentType: "AllPrincipals", principalId: null, resourceId: graphResource,
          scope: `${g1.scope.slice(0, 500)}... (truncated, ${g1.scope.length} chars total)` },
        { id: g2.id, consentType: "Principal", principalId: user, resourceId: graphResource, scope: "User.Read" },
        { id: g3.id, consentType: "AllPrincipals", resourceId: graphResource },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("requiredResourceAccess")));
      assert.ok(result.help.some(hint => hint.includes(`entra service-principal show --id ${sp}`)));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/servicePrincipals/${sp}/oauth2PermissionGrants?`));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") {
        assert.deepEqual(calls[0][1], [grantScopes]);
      }
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists app-only role assignments for the named client`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "service-principal", "app-role-assignment", "list",
        "--service-principal", sp, "--profile", profile], overrides);
      assert.deepEqual(result.appRoleAssignments, [
        { id: a1.id, appRoleId: a1.appRoleId, resourceDisplayName: "Microsoft Graph", resourceId: graphResource },
        { id: a2.id, appRoleId: a2.appRoleId, resourceDisplayName: "fixture-api", resourceId: a2.resourceId },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("requiredResourceAccess")));
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/servicePrincipals/${sp}/appRoleAssignments`));
      assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,appRoleId,resourceDisplayName,resourceId");
      if (mode === "delegated") {
        assert.deepEqual(calls[0][1], [appRoleScopes]);
      }
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied grant reads name the read scope and role, never a write-consent scope`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, grantTransport(true));
      await assert.rejects(
        executeArgv(["entra", "service-principal", "oauth2-grant", "list",
          "--service-principal", sp, "--profile", profile], overrides),
        error => {
          assert.equal(error.code, "GRAPH_ERROR");
          const text = [error.message, ...error.suggestions].join("\n");
          assert.match(text, /Directory\.Read\.All/);
          assert.match(text, /Directory Readers/);
          assert.match(text, /never request \S*ReadWrite/);
          return true;
        },
      );
      await assert.rejects(
        executeArgv(["entra", "service-principal", "app-role-assignment", "list",
          "--service-principal", sp, "--profile", profile], overrides),
        error => {
          assert.equal(error.code, "GRAPH_ERROR");
          const text = [error.message, ...error.suggestions].join("\n");
          assert.match(text, /Application\.Read\.All/);
          assert.match(text, /Application Administrator/);
          assert.match(text, /never request \S*ReadWrite/);
          return true;
        },
      );
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists oauth2-grants end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runGrantsCli(
        ["entra", "service-principal", "oauth2-grant", "list", "--service-principal", sp, "--profile", profile],
        state, mode, grantScopes);
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.oauth2PermissionGrants.length, 3);
      assert.deepEqual(output.count, { returned: 3, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists app-role-assignments end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runGrantsCli(
        ["entra", "service-principal", "app-role-assignment", "list", "--service-principal", sp, "--profile", profile],
        state, mode, appRoleScopes);
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.appRoleAssignments, [
        { id: a1.id, appRoleId: a1.appRoleId, resourceDisplayName: "Microsoft Graph", resourceId: graphResource },
        { id: a2.id, appRoleId: a2.appRoleId, resourceDisplayName: "fixture-api", resourceId: a2.resourceId },
      ]);
      assert.deepEqual(output.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  for (const [command, property] of [
    [["entra", "service-principal", "oauth2-grant", "list"], "expiryTime"],
    [["entra", "service-principal", "app-role-assignment", "list"], "deletedDateTime"],
  ]) {
    for (const flag of ["select", "fields"]) {
      test(`${mode} ${command.join(" ")} rejects ${property} in --${flag} before credentials`, async () => {
        const state = setupProfiles();
        try {
          const { requests, calls, overrides } = overridesFor(mode);
          await assert.rejects(
            executeArgv([...command, "--service-principal", sp, "--profile", profile, `--${flag}`, `id,${property}`], overrides),
            error => {
              assert.equal(error.code, "VALIDATION_ERROR");
              assert.equal(error.message, `Unknown property ${property} in --${flag}`);
              return true;
            },
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

test("delegated grant lists truncate long scopes with a --full hint", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "service-principal", "oauth2-grant", "list",
      "--service-principal", sp, "--profile", "soc"], overrides);
    const grant = result.oauth2PermissionGrants.find(row => row.id === g1.id);
    assert.match(grant.scope, /truncated, \d+ chars total/);
    assert.ok(result.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "service-principal", "oauth2-grant", "list",
      "--service-principal", sp, "--profile", "soc", "--full"], overrides);
    assert.equal(full.oauth2PermissionGrants.find(row => row.id === g1.id).scope, g1.scope);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated grant filters pass through as plain $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "service-principal", "app-role-assignment", "list",
      "--service-principal", sp, "--profile", "soc", "--filter", `resourceId eq '${graphResource}'`], overrides);
    assert.equal(result.appRoleAssignments.length, 2);
    const url = new URL(requests[0].url);
    assert.equal(url.searchParams.get("$filter"), `resourceId eq '${graphResource}'`);
    assert.equal(url.searchParams.has("$count"), false);
    assert.equal(requests[0].headers.ConsistencyLevel, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated grant reads resume a capped list through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "service-principal", "oauth2-grant", "list",
      "--service-principal", sp, "--profile", "soc", "--limit", "2"], overrides);
    assert.equal(first.oauth2PermissionGrants.length, 2);
    assert.deepEqual(first.count, { returned: 2, complete: false, reason: first.count.reason });
    assert.ok(typeof first.cursor === "string" && first.cursor.length > 0);
    const second = await executeArgv(["entra", "service-principal", "oauth2-grant", "list",
      "--service-principal", sp, "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.oauth2PermissionGrants, [{ id: g3.id, consentType: "AllPrincipals", resourceId: graphResource }]);
    assert.deepEqual(second.count, { returned: 1, complete: true });
  } finally {
    teardownProfiles(state);
  }
});

test("grant creation routes are refused before credentials", async () => {
  const state = setupProfiles();
  try {
    const calls = [];
    const { requests, overrides } = overridesFor("delegated", undefined, calls);
    const { listAppRoleAssignments } = await import("../dist/entra-grants.js");
    const { operationFor, leafHelp, LEAVES } = await import("../dist/catalogue.js");
    const leaf = LEAVES.find(item => item.path === "entra service-principal app-role-assignment list");
    const grantOperation = {
      ...operationFor(leaf, "v1.0"),
      id: "v1.0:POST:/servicePrincipals/{servicePrincipal-id}/appRoleAssignments",
      method: "POST",
      path: "/servicePrincipals/{servicePrincipal-id}/appRoleAssignments",
    };
    const { GraphSession } = await import("../dist/graph-session.js");
    const session = new GraphSession(overrides);
    await assert.rejects(
      listAppRoleAssignments(session, { "service-principal": sp },
        { mode: "delegated", tenantId: tenant, clientId: client }, grantOperation, leafHelp(leaf), "soc"),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /two catalogued service-principal grant/.test(error.message);
      },
    );
    assert.equal(requests.length, 0);
    assert.equal(calls.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "service-principal", "oauth2-grant", "list",
        "--service-principal", sp, "--profile", "batch", "--scopes", grantScopes], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "service-principal", "app-role-assignment", "list",
        "--service-principal", sp, "--profile", "batch", "--scopes", appRoleScopes], overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});
