import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

const { mode, workflows, workflowTemplates, taskDefinitions, settings, runs, userProcessingResults, subjectProcessingResults, denied } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const allowedDelegated = new Set([
  "https://graph.microsoft.com/LifecycleWorkflows-Workflow.ReadBasic.All",
  "https://graph.microsoft.com/LifecycleWorkflows.Read.All",
  "https://graph.microsoft.com/LifecycleWorkflows-Reports.Read.All",
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

const base = "/v1.0/identityGovernance/lifecycleWorkflows";

function single(rows, pathname, prefix) {
  const found = rows.find(row => pathname === `${prefix}/${row.id}`);
  assert.ok(found, `Unexpected lifecycle route ${pathname}`);
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
    // Lifecycle lists document $select/$filter only, so a filter must
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
      const count = url.pathname === `${base}/workflows/$count` ? workflows.length
        : url.pathname === `${base}/workflowTemplates/$count` ? workflowTemplates.length
        : url.pathname === `${base}/taskDefinitions/$count` ? taskDefinitions.length
        : url.pathname.endsWith("/runs/$count") ? runs.length
        : url.pathname.endsWith("/userProcessingResults/$count") ? userProcessingResults.length
        : url.pathname.endsWith("/subjectProcessingResults/$count") ? subjectProcessingResults.length
        : undefined;
      assert.ok(count !== undefined, `Unexpected lifecycle count route ${url.pathname}`);
      return { status, headers: { "Content-Type": "text/plain" }, body: String(count) };
    } else if (url.pathname === `${base}/workflows`) {
      body = { value: workflows };
    } else if (url.pathname === `${base}/workflowTemplates`) {
      body = { value: workflowTemplates };
    } else if (url.pathname === `${base}/taskDefinitions`) {
      body = { value: taskDefinitions };
    } else if (url.pathname === `${base}/settings`) {
      body = settings;
    } else if (url.pathname.endsWith("/runs")) {
      body = { value: runs };
    } else if (url.pathname.endsWith("/userProcessingResults")) {
      body = { value: userProcessingResults };
    } else if (url.pathname.endsWith("/subjectProcessingResults")) {
      body = { value: subjectProcessingResults };
    } else if (url.pathname.startsWith(`${base}/workflows/`)) {
      const nested = [...runs, ...userProcessingResults, ...subjectProcessingResults]
        .find(row => url.pathname.endsWith(`/${row.id}`));
      body = nested ?? single(workflows, url.pathname, `${base}/workflows`);
    } else if (url.pathname.startsWith(`${base}/workflowTemplates/`)) {
      body = single(workflowTemplates, url.pathname, `${base}/workflowTemplates`);
    } else {
      body = single(taskDefinitions, url.pathname, `${base}/taskDefinitions`);
    }
    return { status, headers: {}, body: JSON.stringify(body) };
  } },
});
