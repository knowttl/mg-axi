import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

const { mode, denied, scopes: expectedScopes, client, grants, appRoles } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const [major, minor] = process.versions.node.split(".").map(Number);
const exportOption = major >= 26 || (major === 25 && minor >= 9) || (major === 24 && minor >= 15)
  ? "exports" : "namedExports";
const noNetwork = () => { throw new Error("Network access is disabled in the offline read journey"); };
mock.method(globalThis, "fetch", noNetwork);
mock.method(Socket.prototype, "connect", noNetwork);

function credential(profile, scopes, expectedMode) {
  assert.equal(mode, expectedMode);
  assert.equal(profile.mode, expectedMode);
  assert.deepEqual(scopes, [mode === "application"
    ? "https://graph.microsoft.com/.default"
    : expectedScopes]);
  return {
    token: `opaque-fixture-${mode}-token`,
    expiresAt: Date.now() + 3_600_000,
    tenantId: profile.tenantId,
    clientId: profile.clientId,
    ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
  };
}

mock.module(new URL("../../dist/msal-provider.js", import.meta.url), {
  [exportOption]: { MsalProvider: class {
    storage = "session-only";
    async login() { throw new Error("Interactive authentication is disabled"); }
    async silent(profile, scopes) { return credential(profile, scopes, "delegated"); }
  } },
});
mock.module(new URL("../../dist/msal-app-provider.js", import.meta.url), {
  [exportOption]: { MsalApplicationProvider: class {
    storage = "session-only";
    async acquire(profile, scopes) { return credential(profile, scopes, "application"); }
  } },
});
mock.module(new URL("../../dist/api.js", import.meta.url), {
  [exportOption]: { ...await import("../../dist/api.js"), fetchTransport: async request => {
    assert.equal(request.headers.Authorization, `Bearer opaque-fixture-${mode}-token`);
    const url = new URL(request.url);
    assert.equal(url.origin, "https://graph.microsoft.com");
    // Grant collections document $select/$filter only, so a filter must
    // arrive as plain $filter with no advanced-query contract.
    if (url.searchParams.has("$filter")) {
      assert.equal(request.headers.ConsistencyLevel, undefined);
      assert.equal(url.searchParams.has("$count"), false);
    }
    let body;
    let status = 200;
    if (denied) {
      status = 403;
      body = { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } };
    } else if (url.pathname === `/v1.0/servicePrincipals/${client}/oauth2PermissionGrants`) {
      // Delegated-grant rows never carry app-role properties.
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("appRoleId"));
      body = url.searchParams.has("$skiptoken")
        ? { value: grants.slice(2) }
        : { value: grants.slice(0, 2), "@odata.nextLink": `https://graph.microsoft.com/v1.0/servicePrincipals/${client}/oauth2PermissionGrants?%24skiptoken=page2` };
    } else if (url.pathname === `/v1.0/servicePrincipals/${client}/appRoleAssignments`) {
      // App-role rows never carry delegated-scope properties.
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("scope"));
      body = { value: appRoles };
    } else {
      assert.fail(`Unexpected consent-grant route ${url.pathname}`);
    }
    return { status, headers: {}, body: JSON.stringify(body) };
  } },
});
