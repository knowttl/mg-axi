import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

const { mode, riskyUsers, riskDetections, denied, correlation } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const [major, minor] = process.versions.node.split(".").map(Number);
const exportOption = major >= 26 || (major === 25 && minor >= 9) || (major === 24 && minor >= 15)
  ? "exports" : "namedExports";
const noNetwork = () => { throw new Error("Network access is disabled in the offline read journey"); };
mock.method(globalThis, "fetch", noNetwork);
mock.method(Socket.prototype, "connect", noNetwork);

const RISK_SCOPES = [
  "https://graph.microsoft.com/IdentityRiskyUser.Read.All",
  "https://graph.microsoft.com/IdentityRiskEvent.Read.All",
  "https://graph.microsoft.com/AuditLog.Read.All",
];

function credential(profile, scopes, expectedMode) {
  assert.equal(mode, expectedMode);
  assert.equal(profile.mode, expectedMode);
  if (mode === "application") assert.deepEqual(scopes, ["https://graph.microsoft.com/.default"]);
  else assert.ok(scopes.length === 1 && RISK_SCOPES.includes(scopes[0]), `Unexpected delegated scopes ${scopes}`);
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
    assert.equal(request.method, "GET");
    let body;
    let status = 200;
    if (denied) {
      status = 403;
      body = { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } };
    } else if (url.pathname === "/v1.0/auditLogs/signIns") {
      assert.ok(correlation, "Sign-in queries require a correlation fixture");
      const detection = riskDetections[0];
      const page = Number(url.searchParams.get("$skiptoken") ?? 0);
      if (page === 0) {
        assert.equal(url.searchParams.get("$filter"), `createdDateTime ge 2026-09-10T07:55:00.000Z and createdDateTime le 2026-09-10T08:05:00.000Z and userId eq '${detection.userId}'`);
        assert.equal(url.searchParams.get("$select"), "id,createdDateTime,userId,correlationId");
      }
      const signIn = { id: detection.requestId, userId: detection.userId, createdDateTime: detection.activityDateTime, correlationId: detection.correlationId };
      if (correlation === "ambiguous") body = { value: [{ ...signIn, id: "candidate-one" }, { ...signIn, id: "candidate-two" }] };
      else if (correlation === "candidate") body = { value: [{ ...signIn, id: "candidate-one", correlationId: "different-correlation" }] };
      else if (correlation === "failure" && page > 0) {
        status = 403;
        body = { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } };
      } else if (correlation === "budget" || correlation === "failure" || correlation === "limit") {
        body = { value: [{ ...signIn, id: `candidate-${page}` }], "@odata.nextLink": `https://graph.microsoft.com/v1.0/auditLogs/signIns?%24skiptoken=${page + 1}` };
      } else body = { value: [signIn] };
    } else if (url.pathname === "/v1.0/identityProtection/riskyUsers") {
      body = url.searchParams.has("$skiptoken")
        ? { value: riskyUsers.slice(2) }
        : { value: riskyUsers.slice(0, 2), "@odata.nextLink": "https://graph.microsoft.com/v1.0/identityProtection/riskyUsers?%24skiptoken=page2" };
    } else if (url.pathname === "/v1.0/identityProtection/riskDetections") {
      body = url.searchParams.has("$skiptoken")
        ? { value: riskDetections.slice(2) }
        : { value: riskDetections.slice(0, 2), "@odata.nextLink": "https://graph.microsoft.com/v1.0/identityProtection/riskDetections?%24skiptoken=page2" };
    } else {
      const user = riskyUsers.find(row => url.pathname === `/v1.0/identityProtection/riskyUsers/${row.id}`);
      if (user) body = user;
      else {
        const detection = riskDetections.find(row => url.pathname === `/v1.0/identityProtection/riskDetections/${row.id}`);
        assert.ok(detection, "Unexpected risk route");
        body = detection;
      }
    }
    return { status, headers: {}, body: JSON.stringify(body) };
  } },
});
