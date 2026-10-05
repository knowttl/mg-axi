import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

const { mode, agreements, acceptances, denied } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const allowedDelegated = new Set([
  "https://graph.microsoft.com/Agreement.Read.All",
  "https://graph.microsoft.com/AgreementAcceptance.Read",
  "https://graph.microsoft.com/AgreementAcceptance.Read.All",
]);
const [major, minor] = process.versions.node.split(".").map(Number);
const exportOption = major >= 26 || (major === 25 && minor >= 9) || (major === 24 && minor >= 15)
  ? "exports" : "namedExports";
const noNetwork = () => { throw new Error("Network access is disabled in the offline read journey"); };
mock.method(globalThis, "fetch", noNetwork);
mock.method(Socket.prototype, "connect", noNetwork);

function credential(profile, scopes, expectedMode) {
  assert.equal(mode, expectedMode);
  assert.equal(profile.mode, expectedMode);
  assert.equal(scopes.length, 1);
  assert.ok(allowedDelegated.has(scopes[0]), `Unexpected delegated scope ${scopes[0]}`);
  return {
    token: `opaque-fixture-${mode}-token`,
    expiresAt: Date.now() + 3_600_000,
    tenantId: profile.tenantId,
    clientId: profile.clientId,
    accountId: "synthetic-account",
  };
}

function single(rows, pathname, noun) {
  const found = rows.find(row => pathname.endsWith(`/${row.id}`));
  assert.ok(found, `Unexpected ${noun} route ${pathname}`);
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
    async acquire() { throw new Error("Application authentication is disabled for terms-of-use reads"); }
  } },
});
mock.module(new URL("../../dist/api.js", import.meta.url), {
  [exportOption]: { ...await import("../../dist/api.js"), fetchTransport: async request => {
    assert.equal(request.headers.Authorization, `Bearer opaque-fixture-${mode}-token`);
    const url = new URL(request.url);
    assert.equal(url.origin, "https://graph.microsoft.com");
    let body;
    let status = 200;
    if (denied) {
      status = 403;
      body = { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } };
    } else if (url.pathname === "/v1.0/agreements") {
      body = { value: agreements };
    } else if (url.pathname === "/v1.0/agreementAcceptances") {
      body = { value: acceptances };
    } else if (/^\/v1\.0\/agreements\/[^/]+\/acceptances$/.test(url.pathname)) {
      const agreementId = decodeURIComponent(url.pathname.split("/")[3]);
      body = { value: acceptances.filter(row => row.agreementId === agreementId) };
    } else if (/^\/v1\.0\/agreements\/[^/]+\/acceptances\/[^/]+$/.test(url.pathname)) {
      body = single(acceptances, url.pathname, "agreement acceptance");
    } else if (url.pathname.startsWith("/v1.0/agreements/")) {
      body = single(agreements, url.pathname, "agreement");
    } else {
      body = single(acceptances, url.pathname, "agreement-acceptance");
    }
    return { status, headers: {}, body: JSON.stringify(body) };
  } },
});
