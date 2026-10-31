import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";
import { pathToFileURL } from "node:url";
import { join } from "node:path";

// PACK-01 packaged offline journey: one fixture serving the critical SOC
// read path (user, group, Conditional Access policy, sign-in) plus the
// doctor bounded read, through the packaged executable with real profiles.
// MSAL providers resolve synthetic credentials and every HTTP exchange is a
// fixture response; unreachable code throws instead of touching a network.
const { mode, denied, throttled, home, consent } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const moduleUrl = name => process.env.MG_AXI_PACK_ROOT
  ? pathToFileURL(join(process.env.MG_AXI_PACK_ROOT, "dist", name))
  : new URL(`../../dist/${name}`, import.meta.url);
const [major, minor] = process.versions.node.split(".").map(Number);
const exportOption = major >= 26 || (major === 25 && minor >= 9) || (major === 24 && minor >= 15)
  ? "exports" : "namedExports";
const noNetwork = () => { throw new Error("Network access is disabled in the offline pack journey"); };
mock.method(globalThis, "fetch", noNetwork);
mock.method(Socket.prototype, "connect", noNetwork);

const credentials = [];
function credential(profile, scopes) {
  credentials.push({ mode: profile.mode, scopes });
  if (home === "offline") assert.fail("Offline and invalid invocations must not authenticate");
  if (home === "missing") throw new Error("Synthetic missing consent for requested scope");
  if (consent && scopes.some(scope => !consent.includes(scope))) throw new Error("Synthetic scope not consented");
  return {
    token: `opaque-fixture-${mode}-token`,
    expiresAt: Date.now() + (home === "expired" ? -1 : 3_600_000),
    tenantId: profile.tenantId,
    clientId: profile.clientId,
    ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
  };
}

mock.module(moduleUrl("msal-provider.js"), {
  [exportOption]: { MsalProvider: class {
    storage = "session-only";
    async login() { throw new Error("Interactive authentication is disabled"); }
    async silent(profile, scopes) { return credential(profile, scopes); }
  } },
});
mock.module(moduleUrl("msal-app-provider.js"), {
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

const requests = [];
process.once("exit", () => {
  for (const credential of credentials) {
    assert.equal(credential.mode, mode);
    if (home) {
      assert.notEqual(home, "offline", "Offline invocations must not authenticate");
      assert.equal(credential.scopes.length, 1);
      assert.ok((mode === "application" ? ["https://graph.microsoft.com/.default"]
        : ["https://graph.microsoft.com/User.Read.All", "https://graph.microsoft.com/Mail.Read"]).includes(credential.scopes[0]));
    }
  }
  if (home) {
    assert.ok(requests.length <= (home === "offline" ? 0 : 2), "Home must never follow pages or retry");
    assert.equal(new Set(requests.map(request => new URL(request.url).pathname)).size, requests.length, "Home must read each pack at most once");
  }
  for (const request of requests) {
    assert.equal(request.headers.Authorization, `Bearer opaque-fixture-${mode}-token`);
    const url = new URL(request.url);
    assert.equal(url.origin, "https://graph.microsoft.com");
    if (home) {
      assert.equal(request.method, "GET");
      assert.equal(url.searchParams.get("$top"), "5");
      assert.equal(url.searchParams.get("$select"), "id");
      assert.equal(url.searchParams.has("$skiptoken"), false);
      assert.ok(["/v1.0/users", "/v1.0/me/mailFolders"].includes(url.pathname));
    }
  }
});
mock.module(moduleUrl("api.js"), {
  [exportOption]: { ...await import(moduleUrl("api.js")), fetchTransport: async request => {
    requests.push(request);
    const url = new URL(request.url);
    if (home) {
      if (home === "unavailable") throw new Error("Synthetic offline transport");
      if (home === "denied" || (home === "partial" && url.pathname.includes("mailFolders"))) {
        return { status: 403, headers: {}, body: JSON.stringify({ error: { code: "Authorization_RequestDenied" } }) };
      }
      if (home === "throttled") return { status: 503, headers: {}, body: "" };
      const users = url.pathname === "/v1.0/users";
      return { status: 200, headers: {}, body: JSON.stringify(home === "empty" ? { value: [] } : {
        value: users ? [user, { id: "second-user" }] : [{ id: "inbox" }],
        "@odata.count": users ? 12 : 1,
        ...(users ? { "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=next" } : {}),
      }) };
    }
    if (throttled) return { status: 429, headers: { "Retry-After": "60" }, body: JSON.stringify({ error: { code: "TooManyRequests" } }) };
    const body = () => {
      if (denied) return { status: 403, body: { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } } };
      if (url.pathname === "/v1.0/users") return { status: 200, body: { value: [user] } };
      if (url.pathname === "/v1.0/me/mailFolders") return { status: 200, body: { value: [{ id: "inbox", displayName: "Inbox" }] } };
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
