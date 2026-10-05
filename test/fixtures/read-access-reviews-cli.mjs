import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

const { mode, denied, scopes, definitions, instances, decisions, reviewers = [], stages = [] } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const expectedDelegated = scopes ?? "https://graph.microsoft.com/AccessReview.Read.All";
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
    // Access-review collections document $select/$filter only, so a filter
    // must arrive as plain $filter with no advanced-query contract.
    if (url.searchParams.has("$filter")) {
      assert.equal(request.headers.ConsistencyLevel, undefined);
      assert.equal(url.searchParams.has("$count"), false);
    }
    const base = "/v1.0/identityGovernance/accessReviews/definitions";
    let body;
    let status = 200;
    if (denied) {
      status = 403;
      body = { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } };
    } else if (url.pathname.endsWith("/decisions")) {
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("justification"));
      body = { value: decisions };
    } else if (/\/decisions\/[^/]+$/.test(url.pathname)) {
      const decision = decisions.find(row => url.pathname.endsWith(`/decisions/${row.id}`));
      assert.ok(decision, "Unexpected access-review decision route");
      body = decision;
    } else if (url.pathname.endsWith("/contactedReviewers")) {
      body = { value: reviewers };
    } else if (/\/contactedReviewers\/[^/]+$/.test(url.pathname)) {
      const reviewer = reviewers.find(row => url.pathname.endsWith(`/contactedReviewers/${row.id}`));
      assert.ok(reviewer, "Unexpected contacted-reviewer route");
      body = reviewer;
    } else if (url.pathname.endsWith("/stages")) {
      body = { value: stages };
    } else if (/\/stages\/[^/]+$/.test(url.pathname)) {
      const stage = stages.find(row => url.pathname.endsWith(`/stages/${row.id}`));
      assert.ok(stage, "Unexpected stage route");
      body = stage;
    } else if (url.pathname === base) {
      body = { value: definitions };
    } else if (url.pathname === `${base}/${definitions[0].id}/instances`) {
      body = { value: instances };
    } else {
      const definition = definitions.find(row => url.pathname === `${base}/${row.id}`);
      if (definition) {
        body = definition;
      } else {
        const instance = instances.find(row => url.pathname === `${base}/${definitions[0].id}/instances/${row.id}`);
        assert.ok(instance, "Unexpected access-review route");
        body = instance;
      }
    }
    return { status, headers: {}, body: JSON.stringify(body) };
  } },
});
