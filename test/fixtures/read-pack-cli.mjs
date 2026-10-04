import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

// PACK-01 packaged offline journey: one fixture serving the critical SOC
// read path (user, group, Conditional Access policy, sign-in) plus the
// doctor bounded read, through the packaged executable with real profiles.
// MSAL providers resolve synthetic credentials and every HTTP exchange is a
// fixture response; unreachable code throws instead of touching a network.
const { mode, denied } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const [major, minor] = process.versions.node.split(".").map(Number);
const exportOption = major >= 26 || (major === 25 && minor >= 9) || (major === 24 && minor >= 15)
  ? "exports" : "namedExports";
const noNetwork = () => { throw new Error("Network access is disabled in the offline pack journey"); };
mock.method(globalThis, "fetch", noNetwork);
mock.method(Socket.prototype, "connect", noNetwork);

function credential(profile, scopes) {
  assert.equal(profile.mode, mode);
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
    async silent(profile, scopes) { return credential(profile, scopes); }
  } },
});
mock.module(new URL("../../dist/msal-app-provider.js", import.meta.url), {
  [exportOption]: { MsalApplicationProvider: class {
    storage = "session-only";
    async acquire(profile, scopes) { return credential(profile, scopes); }
  } },
});

const user = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  displayName: "Adele Vance",
  userPrincipalName: "AdeleV@contoso.com",
  mail: "AdeleV@contoso.com",
};
const group = {
  id: "gggggggg-gggg-4ggg-8ggg-gggggggggggg",
  displayName: "SOC Tier 1",
  mail: "soc@contoso.com",
  groupTypes: [],
};
const member = { id: user.id, displayName: user.displayName };
const policy = { id: "pppppppp-pppp-4ppp-8ppp-pppppppppppp", displayName: "Require MFA", state: "enabled" };
const signIn = {
  id: "ssssssss-ssss-4sss-8sss-ssssssssssss",
  createdDateTime: "2026-09-02T00:00:00Z",
  userPrincipalName: user.userPrincipalName,
  appDisplayName: "SOC console",
};

mock.module(new URL("../../dist/api.js", import.meta.url), {
  [exportOption]: { ...await import("../../dist/api.js"), fetchTransport: async request => {
    assert.equal(request.headers.Authorization, `Bearer opaque-fixture-${mode}-token`);
    const url = new URL(request.url);
    assert.equal(url.origin, "https://graph.microsoft.com");
    const body = () => {
      if (denied) return { status: 403, body: { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } } };
      if (url.pathname === "/v1.0/users") return { status: 200, body: { value: [user] } };
      if (url.pathname === `/v1.0/users/${user.id}`) return { status: 200, body: user };
      if (url.pathname === "/v1.0/groups") return { status: 200, body: { value: [group] } };
      if (url.pathname === `/v1.0/groups/${group.id}/members`) return { status: 200, body: { value: [member] } };
      if (url.pathname === "/v1.0/identity/conditionalAccess/policies") return { status: 200, body: { value: [policy] } };
      if (url.pathname === "/v1.0/auditLogs/signIns") return { status: 200, body: { value: [signIn] } };
      assert.fail(`Unexpected pack journey route ${url.pathname}`);
    };
    const response = body();
    return { status: response.status, headers: {}, body: JSON.stringify(response.body) };
  } },
});
