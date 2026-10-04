import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

const { mode, denied, scopes, methods, report } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const expectedDelegated = scopes ?? "https://graph.microsoft.com/UserAuthenticationMethod.Read.All";
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
    if (url.searchParams.has("$filter")) {
      assert.equal(request.headers.ConsistencyLevel, undefined);
      assert.equal(url.searchParams.has("$count"), false);
    }
    let body;
    let status = 200;
    if (denied) {
      status = 403;
      body = { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } };
    } else if (/^\/v1\.0\/users\/[^/]+\/authentication\/methods$/.test(url.pathname)) {
      assert.equal(url.searchParams.has("$select"), false);
      assert.equal(url.searchParams.has("$filter"), false);
      body = { value: methods };
    } else if (url.pathname === "/v1.0/reports/authenticationMethods/userRegistrationDetails") {
      body = { value: report };
    } else {
      throw new Error(`Unexpected authentication route ${url.pathname}`);
    }
    return { status, headers: {}, body: JSON.stringify(body) };
  } },
});
