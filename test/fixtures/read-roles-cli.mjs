import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

const { mode, denied, scopes, roles, assignments, eligible, active } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const expectedDelegated = scopes ?? "https://graph.microsoft.com/RoleManagement.Read.Directory";
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
    : expectedDelegated]);
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
    // Role and PIM collections document $select/$filter/$expand only, so a
    // filter must arrive as plain $filter with no advanced-query contract.
    if (url.searchParams.has("$filter")) {
      assert.equal(request.headers.ConsistencyLevel, undefined);
      assert.equal(url.searchParams.has("$count"), false);
    }
    let body;
    let status = 200;
    if (denied) {
      status = 403;
      body = { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } };
    } else if (url.pathname === "/v1.0/directoryRoles") {
      body = url.searchParams.has("$skiptoken")
        ? { value: roles.slice(2) }
        : { value: roles.slice(0, 2), "@odata.nextLink": "https://graph.microsoft.com/v1.0/directoryRoles?%24skiptoken=page2" };
    } else if (url.pathname === "/v1.0/roleManagement/directory/roleAssignments") {
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("createdDateTime"));
      body = { value: assignments };
    } else if (url.pathname === "/v1.0/roleManagement/directory/roleEligibilityScheduleInstances") {
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("assignmentType"));
      body = { value: eligible };
    } else if (url.pathname === "/v1.0/roleManagement/directory/roleAssignmentScheduleInstances") {
      body = { value: active };
    } else {
      const role = roles.find(row => url.pathname === `/v1.0/directoryRoles/${row.id}`);
      assert.ok(role, "Unexpected directory-role route");
      body = role;
    }
    return { status, headers: {}, body: JSON.stringify(body) };
  } },
});
