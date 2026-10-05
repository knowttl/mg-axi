import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

const { mode, catalogs, accessPackages, policies, roleScopes, assignments, assignmentRequests, denied } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const allowedDelegated = new Set(["https://graph.microsoft.com/EntitlementManagement.Read.All"]);
const [major, minor] = process.versions.node.split(".").map(Number);
const exportOption = major >= 26 || (major === 25 && minor >= 9) || (major === 24 && minor >= 15)
  ? "exports" : "namedExports";
const noNetwork = () => { throw new Error("Network access is disabled in the offline read journey"); };
mock.method(globalThis, "fetch", noNetwork);
mock.method(Socket.prototype, "connect", noNetwork);

function credential(profile, scopes, expectedMode) {
  assert.equal(mode, expectedMode);
  assert.equal(profile.mode, expectedMode);
  if (mode === "application") {
    assert.deepEqual(scopes, ["https://graph.microsoft.com/.default"]);
  } else {
    assert.equal(scopes.length, 1);
    assert.ok(allowedDelegated.has(scopes[0]), `Unexpected delegated scope ${scopes[0]}`);
  }
  return {
    token: `opaque-fixture-${mode}-token`,
    expiresAt: Date.now() + 3_600_000,
    tenantId: profile.tenantId,
    clientId: profile.clientId,
    ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
  };
}

const base = "/v1.0/identityGovernance/entitlementManagement";

function single(rows, pathname, prefix) {
  const found = rows.find(row => pathname === `${prefix}/${row.id}`);
  assert.ok(found, `Unexpected entitlement route ${pathname}`);
  return found;
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
    // Entitlement lists document $select/$filter only, so a filter must
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
    } else if (url.pathname.endsWith("/$count")) {
      const count = url.pathname.includes("/assignmentPolicies") ? policies.length
        : url.pathname.includes("/resourceRoleScopes") ? roleScopes.length
        : url.pathname === `${base}/assignments/$count` ? assignments.length
        : url.pathname === `${base}/assignmentRequests/$count` ? assignmentRequests.length
        : url.pathname === `${base}/accessPackages/$count` ? accessPackages.length
        : catalogs.length;
      return { status, headers: { "Content-Type": "text/plain" }, body: String(count) };
    } else if (url.pathname === `${base}/catalogs`) {
      body = { value: catalogs };
    } else if (url.pathname === `${base}/accessPackages`) {
      body = { value: accessPackages };
    } else if (url.pathname === `${base}/assignments`) {
      body = { value: assignments };
    } else if (url.pathname === `${base}/assignmentRequests`) {
      body = { value: assignmentRequests };
    } else if (url.pathname === `${base}/accessPackages/${accessPackages[0].id}/assignmentPolicies`) {
      body = { value: policies };
    } else if (url.pathname === `${base}/accessPackages/${accessPackages[0].id}/resourceRoleScopes`) {
      body = { value: roleScopes };
    } else if (url.pathname.startsWith(`${base}/assignments/`)) {
      body = single(assignments, url.pathname, `${base}/assignments`);
    } else if (url.pathname.startsWith(`${base}/assignmentRequests/`)) {
      body = single(assignmentRequests, url.pathname, `${base}/assignmentRequests`);
    } else if (url.pathname.startsWith(`${base}/catalogs/`)) {
      body = single(catalogs, url.pathname, `${base}/catalogs`);
    } else if (url.pathname.includes("/assignmentPolicies/")) {
      body = single(policies, url.pathname, `${base}/accessPackages/${accessPackages[0].id}/assignmentPolicies`);
    } else if (url.pathname.includes("/resourceRoleScopes/")) {
      body = single(roleScopes, url.pathname, `${base}/accessPackages/${accessPackages[0].id}/resourceRoleScopes`);
    } else {
      body = single(accessPackages, url.pathname, `${base}/accessPackages`);
    }
    return { status, headers: {}, body: JSON.stringify(body) };
  } },
});
