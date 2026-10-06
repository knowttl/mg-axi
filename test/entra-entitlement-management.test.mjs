import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const tenant = "33333333-3333-4333-8333-333333333333";
const client = "44444444-4444-4222-8222-444444444444";
const entitlementScopes = ["https://graph.microsoft.com/EntitlementManagement.Read.All"];

const longDescription = `General catalog holding every governed bundle${" with considerable detail" .repeat(30)}`;
const longApprovalNote = `Second-level approval${" with considerable detail".repeat(30)}`;
const cat1 = {
  id: "11111111-1111-4111-8111-111111111111",
  displayName: "General",
  description: longDescription,
  catalogType: "userManaged",
  state: "published",
  isExternallyVisible: false,
  createdDateTime: "2023-01-01T00:00:00Z",
  modifiedDateTime: "2023-02-01T00:00:00Z",
};
const cat2 = {
  id: "22222222-2222-4222-8222-222222222222",
  displayName: "Service default",
  description: null,
  catalogType: "serviceDefault",
  state: "published",
  isExternallyVisible: false,
  createdDateTime: "2023-03-01T00:00:00Z",
  modifiedDateTime: "2023-04-01T00:00:00Z",
};
const cat3 = { id: "cat-bare-3" };
const catalogs = [cat1, cat2, cat3];
const pkg1 = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  displayName: "Engineering bundle",
  description: "Laptops, repos and review duties",
  isHidden: false,
  createdDateTime: "2024-01-01T00:00:00Z",
  modifiedDateTime: "2024-02-01T00:00:00Z",
};
const pkg2 = {
  id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  displayName: "Quiet bundle",
  description: null,
  isHidden: true,
  createdDateTime: "2024-03-01T00:00:00Z",
  modifiedDateTime: "2024-04-01T00:00:00Z",
};
const accessPackages = [pkg1, pkg2];
const pol1 = {
  id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  displayName: "Engineering requests",
  description: "Manager approval then access",
  allowedTargetScope: "specificDirectoryUsers",
  automaticRequestSettings: null,
  createdDateTime: "2024-01-02T00:00:00Z",
  expiration: null,
  modifiedDateTime: "2024-02-02T00:00:00Z",
  notificationSettings: null,
  requestApprovalSettings: { isApprovalRequired: true, notes: longApprovalNote },
  requestorSettings: null,
  reviewSettings: null,
  specificAllowedTargets: [],
};
const pol2 = {
  id: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
  displayName: "Open requests",
  description: null,
  allowedTargetScope: "allMemberUsers",
  automaticRequestSettings: null,
  createdDateTime: "2024-03-02T00:00:00Z",
  expiration: null,
  modifiedDateTime: "2024-04-02T00:00:00Z",
  notificationSettings: null,
  requestApprovalSettings: null,
  requestorSettings: null,
  reviewSettings: null,
  specificAllowedTargets: [],
};
const policies = [pol1, pol2];
const rs1 = { id: "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee", createdDateTime: "2024-01-03T00:00:00Z" };
const rs2 = { id: "ffffffff-ffff-4fff-8fff-ffffffffffff", createdDateTime: "2024-02-03T00:00:00Z" };
const roleScopes = [rs1, rs2];
const asg1 = {
  id: "11111111-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  state: "delivered",
  status: "Delivered",
  expiredDateTime: "2024-06-01T00:00:00Z",
  schedule: { startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2024-12-31T00:00:00Z" },
  customExtensionCalloutInstances: [],
  target: {
    id: "99999999-9999-4999-8999-999999999999",
    objectId: "99999999-9999-4999-8999-999999999999",
    displayName: "Alice Example",
    email: "alice@example.invalid",
    principalName: "alice@example.invalid",
    subjectType: "user",
  },
  accessPackage: { id: pkg1.id, displayName: "Engineering bundle", description: "Laptops, repos and review duties" },
};
const asg2 = {
  id: "22222222-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  state: "expired",
  status: "ExpiredNotificationTriggered",
  expiredDateTime: "2024-07-01T00:00:00Z",
  schedule: null,
  customExtensionCalloutInstances: null,
  target: null,
  accessPackage: { id: pkg2.id },
};
const assignments = [asg1, asg2];
const longJustification = `Temporary cover for the release${" with considerable detail".repeat(30)}`;
const req1 = {
  id: "33333333-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  requestType: "userAdd",
  state: "delivered",
  status: "Delivered",
  justification: longJustification,
  schedule: { startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2024-12-31T00:00:00Z" },
  createdDateTime: "2024-01-02T00:00:00Z",
  completedDateTime: "2024-01-03T00:00:00Z",
  answers: [{ displayValue: "Covering the release", answeredQuestion: { id: "q1" } }],
  customExtensionCalloutInstances: [],
  accessPackage: { id: pkg1.id, displayName: "Engineering bundle" },
  assignment: { id: asg1.id },
  requestor: { id: "99999999-9999-4999-8999-999999999999", displayName: "Alice Example" },
};
const req2 = {
  id: "44444444-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  requestType: "adminAdd",
  state: "pendingApproval",
  status: "PendingApproval",
  justification: null,
  schedule: null,
  createdDateTime: "2024-02-02T00:00:00Z",
  completedDateTime: null,
  answers: [],
  customExtensionCalloutInstances: [],
  accessPackage: { id: pkg2.id },
  assignment: null,
  requestor: { id: "88888888-8888-4888-8888-888888888888", displayName: "Bob Example" },
};
const assignmentRequests = [req1, req2];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-entitlement-"));
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(dir, "config.json");
  const profiles = new Profiles();
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  return { dir, previous };
}

function teardownProfiles(state) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function credentialService(mode, calls) {
  return {
    credential: async (...args) => {
      calls.push([mode, ...args]);
      return {
        token: `opaque-fixture-${mode}-token`,
        expiresAt: Date.now() + 3_600_000,
        tenantId: tenant,
        clientId: client,
        ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
      };
    },
  };
}

function transport(handler) {
  const requests = [];
  const send = async request => {
    requests.push(request);
    const response = await handler(request);
    return { headers: {}, body: "", ...response };
  };
  return { requests, send };
}

function json(status, body, headers = {}) {
  return { status, headers, body: JSON.stringify(body) };
}

const emBase = "/v1.0/identityGovernance/entitlementManagement";

function entitlementTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === `${emBase}/catalogs/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(catalogs.length) };
    }
    if (path === `${emBase}/accessPackages/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(accessPackages.length) };
    }
    if (path === `${emBase}/accessPackages/${pkg1.id}/assignmentPolicies/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(policies.length) };
    }
    if (path === `${emBase}/accessPackages/${pkg1.id}/resourceRoleScopes/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(roleScopes.length) };
    }
    if (path === `${emBase}/assignments/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(assignments.length) };
    }
    if (path === `${emBase}/assignmentRequests/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(assignmentRequests.length) };
    }
    if (path === `${emBase}/catalogs`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [cat3] });
      return json(200, {
        value: [cat1, cat2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/catalogs?%24skiptoken=page2",
      });
    }
    if (path === `${emBase}/accessPackages`) return json(200, { value: accessPackages });
    if (path === `${emBase}/assignments`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [asg2] });
      return json(200, {
        value: [asg1],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignments?%24skiptoken=page2",
      });
    }
    if (path === `${emBase}/assignmentRequests`) return json(200, { value: assignmentRequests });
    if (path === `${emBase}/accessPackages/${pkg1.id}/assignmentPolicies`) return json(200, { value: policies });
    if (path === `${emBase}/accessPackages/${pkg1.id}/resourceRoleScopes`) return json(200, { value: roleScopes });
    for (const rows of [catalogs, accessPackages, policies, roleScopes]) {
      const found = rows.find(row => path === `${emBase}/catalogs/${row.id}`
        || path === `${emBase}/accessPackages/${row.id}`
        || path === `${emBase}/accessPackages/${pkg1.id}/assignmentPolicies/${row.id}`
        || path === `${emBase}/accessPackages/${pkg1.id}/resourceRoleScopes/${row.id}`);
      if (found) return json(200, found);
    }
    for (const [rows, base] of [[assignments, `${emBase}/assignments`], [assignmentRequests, `${emBase}/assignmentRequests`]]) {
      const found = rows.find(row => path === `${base}/${row.id}`);
      if (found) return json(200, found);
    }
    return json(404, { error: { code: "Request_ResourceNotFound", message: "no such entitlement object" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? entitlementTransport();
  const credential = credentialService(mode, calls);
  return {
    requests: fixture.requests,
    calls,
    overrides: {
      transport: fixture.send,
      delegated: mode === "delegated" ? credential : credentialService("delegated", []),
      application: mode === "application" ? credential : credentialService("application", []),
    },
  };
}

function runEntitlementCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-entitlement-management-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, catalogs: [cat1, cat2], accessPackages, policies, roleScopes, assignments, assignmentRequests, denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists catalogs with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "catalog", "list", "--profile", profile,
        "--select", "id,displayName,state,catalogType,description"], overrides);
      assert.deepEqual(result.catalogs, [
        { id: cat1.id, displayName: "General", state: "published", catalogType: "userManaged", description: `${longDescription.slice(0, 500)}... (truncated, ${longDescription.length} chars total)` },
        { id: cat2.id, displayName: "Service default", state: "published", catalogType: "serviceDefault", description: null },
        { id: cat3.id },
      ]);
      assert.deepEqual(result.count, "3 catalogs");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("entra entitlement catalog show --id <catalog-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/catalogs?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(entitlementScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists catalogs with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "catalog", "list", "--profile", profile,
        "--filter", "state eq 'published'"], overrides);
      assert.equal(result.catalogs.length, 3);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "state eq 'published'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped catalog list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "entitlement", "catalog", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.catalogs.map(row => row.id), [cat1.id]);
      assert.equal(first.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "entitlement", "catalog", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.catalogs.map(row => row.id), [cat2.id, cat3.id]);
      assert.deepEqual(second.count, "2 catalogs");
      assert.equal(second.total, null);
      assert.equal(second.complete, true);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one catalog with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "catalog", "show", "--id", cat1.id, "--profile", profile], overrides);
      assert.deepEqual(result.catalog, { ...cat1, description: `${longDescription.slice(0, 500)}... (truncated, ${longDescription.length} chars total)` });
      assert.ok(result.help.some(hint => hint.includes("--full")));
      const missing = await executeArgv(["entra", "entitlement", "catalog", "show", "--id", cat3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.catalog, { id: cat3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown catalog ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "entitlement", "catalog", "show", "--id", "catalog-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists access packages with compact rows`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "access-package", "list", "--profile", profile], overrides);
      assert.deepEqual(result.accessPackages, [
        { id: pkg1.id, displayName: "Engineering bundle", isHidden: false },
        { id: pkg2.id, displayName: "Quiet bundle", isHidden: true },
      ]);
      assert.deepEqual(result.count, "2 access packages");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("entra entitlement access-package show --id <access-package-id>")));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages?"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one access package with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "access-package", "show", "--id", pkg1.id, "--profile", profile], overrides);
      assert.deepEqual(result.accessPackage, pkg1);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts catalogs and access packages as scalars without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const catalogCount = await executeArgv(["entra", "entitlement", "catalog", "count", "--profile", profile], overrides);
      assert.deepEqual(catalogCount, { count: { returned: 3, complete: true } });
      const packageCount = await executeArgv(["entra", "entitlement", "access-package", "count", "--profile", profile], overrides);
      assert.deepEqual(packageCount, { count: { returned: 2, complete: true } });
      assert.deepEqual(requests.map(request => request.url), [
        "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/catalogs/$count",
        "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages/$count",
      ]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} entitlement counts refuse collection flags before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "entitlement", "catalog", "count", "--profile", profile, "--filter", "state eq 'published'"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists assignment policies of one access package`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "assignment-policy", "list",
        "--access-package", pkg1.id, "--profile", profile], overrides);
      assert.deepEqual(result.assignmentPolicies, [
        { id: pol1.id, displayName: "Engineering requests", allowedTargetScope: "specificDirectoryUsers" },
        { id: pol2.id, displayName: "Open requests", allowedTargetScope: "allMemberUsers" },
      ]);
      assert.deepEqual(result.count, "2 assignment policies");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages/${pkg1.id}/assignmentPolicies?`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} assignment-policy reads need their parent access package before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "entitlement", "assignment-policy", "list", "--profile", profile], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "entitlement", "resource-role-scope", "count", "--profile", profile], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one assignment policy with nested settings projected`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "assignment-policy", "show",
        "--access-package", pkg1.id, "--id", pol1.id, "--profile", profile], overrides);
      assert.deepEqual(result.assignmentPolicy, {
        ...pol1,
        requestApprovalSettings: { isApprovalRequired: true, notes: `${longApprovalNote.slice(0, 500)}... (truncated, ${longApprovalNote.length} chars total)` },
      });
      assert.ok(result.help.some(hint => hint.includes("--full")));
      const full = await executeArgv(["entra", "entitlement", "assignment-policy", "show",
        "--access-package", pkg1.id, "--id", pol1.id, "--profile", profile, "--full"], overrides);
      assert.deepEqual(full.assignmentPolicy, pol1);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts assignment policies of one access package as a scalar`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "assignment-policy", "count",
        "--access-package", pkg1.id, "--profile", profile], overrides);
      assert.deepEqual(result, { count: { returned: 2, complete: true } });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, `https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages/${pkg1.id}/assignmentPolicies/$count`);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists resource-role scopes of one access package`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "resource-role-scope", "list",
        "--access-package", pkg1.id, "--profile", profile], overrides);
      assert.deepEqual(result.resourceRoleScopes, [rs1, rs2]);
      assert.deepEqual(result.count, "2 resource-role scopes");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages/${pkg1.id}/resourceRoleScopes?`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows and counts one access package's resource-role scopes`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const shown = await executeArgv(["entra", "entitlement", "resource-role-scope", "show",
        "--access-package", pkg1.id, "--id", rs1.id, "--profile", profile], overrides);
      assert.deepEqual(shown.resourceRoleScope, rs1);
      const counted = await executeArgv(["entra", "entitlement", "resource-role-scope", "count",
        "--access-package", pkg1.id, "--profile", profile], overrides);
      assert.deepEqual(counted, { count: { returned: 2, complete: true } });
      assert.equal(requests[1].url, `https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages/${pkg1.id}/resourceRoleScopes/$count`);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists assignments with flattened linkage from the fixed expand`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "assignment", "list", "--profile", profile], overrides);
      assert.deepEqual(result.assignments, [
        {
          id: asg1.id,
          state: "delivered",
          targetId: "99999999-9999-4999-8999-999999999999",
          targetDisplayName: "Alice Example",
          accessPackageId: pkg1.id,
          expiredDateTime: "2024-06-01T00:00:00Z",
          schedule: { startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2024-12-31T00:00:00Z" },
        },
        {
          id: asg2.id,
          state: "expired",
          accessPackageId: pkg2.id,
          expiredDateTime: "2024-07-01T00:00:00Z",
          schedule: null,
        },
      ]);
      assert.deepEqual(result.count, "2 assignments");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("entra entitlement assignment show --id <assignment-id>")));
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$expand"), "target,accessPackage");
      assert.ok(!sent.get("$select").split(",").some(field => ["targetId", "targetDisplayName", "accessPackageId"].includes(field)));
      assert.ok(!JSON.stringify(result).includes("alice@example.invalid"));
      assert.ok(!JSON.stringify(result).includes("Laptops, repos and review duties"));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(entitlementScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists assignments with a plain documented filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "assignment", "list", "--profile", profile,
        "--filter", "state eq 'Delivered'"], overrides);
      assert.equal(result.assignments.length, 2);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "state eq 'Delivered'");
      assert.equal(sent.get("$expand"), "target,accessPackage");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped assignment list with linkage intact`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "entitlement", "assignment", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.assignments.map(row => row.id), [asg1.id]);
      assert.equal(first.assignments[0].targetDisplayName, "Alice Example");
      assert.equal(first.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "entitlement", "assignment", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.assignments.map(row => row.id), [asg2.id]);
      assert.equal(second.assignments[0].accessPackageId, pkg2.id);
      assert.ok(!Object.hasOwn(second.assignments[0], "targetId"));
      assert.deepEqual(second.count, "1 assignments");
      assert.equal(second.total, null);
      assert.equal(second.complete, true);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one assignment with the full reviewed grant set`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "assignment", "show", "--id", asg1.id, "--profile", profile], overrides);
      assert.deepEqual(result.assignment, {
        id: asg1.id,
        state: "delivered",
        status: "Delivered",
        expiredDateTime: "2024-06-01T00:00:00Z",
        schedule: { startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2024-12-31T00:00:00Z" },
        customExtensionCalloutInstances: [],
        targetId: "99999999-9999-4999-8999-999999999999",
        targetDisplayName: "Alice Example",
        accessPackageId: pkg1.id,
      });
      assert.equal(new URL(requests[0].url).searchParams.get("$expand"), "target,accessPackage");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown assignment ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "entitlement", "assignment", "show", "--id", "assignment-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists assignment requests with compact linkage and no justification`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "entitlement", "assignment-request", "list", "--profile", profile], overrides);
      assert.deepEqual(result.assignmentRequests, [
        {
          id: req1.id,
          requestType: "userAdd",
          state: "delivered",
          accessPackageId: pkg1.id,
          assignmentId: asg1.id,
          createdDateTime: "2024-01-02T00:00:00Z",
          completedDateTime: "2024-01-03T00:00:00Z",
          schedule: { startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2024-12-31T00:00:00Z" },
        },
        {
          id: req2.id,
          requestType: "adminAdd",
          state: "pendingApproval",
          accessPackageId: pkg2.id,
          createdDateTime: "2024-02-02T00:00:00Z",
          completedDateTime: null,
          schedule: null,
        },
      ]);
      assert.deepEqual(result.count, "2 assignment requests");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$expand"), "accessPackage,assignment");
      assert.ok(!JSON.stringify(result).includes(longJustification.slice(0, 50)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} request justification and answers ride only behind an explicit select`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const shown = await executeArgv(["entra", "entitlement", "assignment-request", "show",
        "--id", req1.id, "--profile", profile], overrides);
      assert.ok(!Object.hasOwn(shown.assignmentRequest, "justification"));
      assert.ok(!Object.hasOwn(shown.assignmentRequest, "answers"));
      const selected = await executeArgv(["entra", "entitlement", "assignment-request", "show",
        "--id", req1.id, "--profile", profile, "--select", "id,justification,answers"], overrides);
      assert.deepEqual(selected.assignmentRequest, {
        id: req1.id,
        justification: `${longJustification.slice(0, 500)}... (truncated, ${longJustification.length} chars total)`,
        answers: req1.answers,
      });
      assert.ok(selected.help.some(hint => hint.includes("--full")));
      assert.ok(new URL(requests[1].url).searchParams.get("$select").split(",").includes("justification"));
      const full = await executeArgv(["entra", "entitlement", "assignment-request", "show",
        "--id", req1.id, "--profile", profile, "--select", "id,justification", "--full"], overrides);
      assert.equal(full.assignmentRequest.justification, longJustification);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts assignments and assignment requests as scalars`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const assignmentCount = await executeArgv(["entra", "entitlement", "assignment", "count", "--profile", profile], overrides);
      assert.deepEqual(assignmentCount, { count: { returned: 2, complete: true } });
      const requestCount = await executeArgv(["entra", "entitlement", "assignment-request", "count", "--profile", profile], overrides);
      assert.deepEqual(requestCount, { count: { returned: 2, complete: true } });
      assert.deepEqual(requests.map(request => request.url), [
        "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignments/$count",
        "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignmentRequests/$count",
      ]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} assignment counts refuse collection flags before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "entitlement", "assignment", "count", "--profile", profile, "--limit", "1"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "entitlement", "assignment-request", "count", "--profile", profile, "--filter", "state eq 'Delivered'"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} navigation names fail as unknown properties on assignment reads`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "entitlement", "assignment", "list", "--profile", profile, "--select", "id,target"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property target/.test(error.message)
          && error.suggestions.join("\n").includes("Known properties: "),
      );
      await assert.rejects(
        executeArgv(["entra", "entitlement", "assignment-request", "list", "--profile", profile, "--select", "id,requestor"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property requestor/.test(error.message),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied assignment reads surface the delegation roles and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "entitlement", "assignment", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Catalog reader")));
        assert.ok(error.suggestions.some(hint => hint.includes("Identity Governance Administrator")));
        assert.ok(error.suggestions.some(hint => hint.includes("P2 or ID Governance")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "entitlement", "assignment-request", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("EntitlementManagement.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty assignment and request lists stay definitive`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "entitlement", "assignment", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.assignments, []);
      assert.deepEqual(listed.count, "0 assignments");
      assert.equal(listed.total, null);
      assert.equal(listed.complete, true);
      assert.ok(listed.help.some(hint => hint.includes("0 assignments matched")));
      const requested = await executeArgv(["entra", "entitlement", "assignment-request", "list", "--profile", profile], overrides);
      assert.deepEqual(requested.assignmentRequests, []);
      assert.ok(requested.help.some(hint => hint.includes("0 assignment requests matched")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown entitlement properties fail naming the known set`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "entitlement", "catalog", "list", "--profile", profile, "--select", "id,ownerId"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property ownerId/.test(error.message)
          && error.suggestions.join("\n").includes("Known properties: "),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied entitlement reads surface scope, role and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "entitlement", "catalog", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("EntitlementManagement.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Global Reader and Identity Governance Administrator")));
        assert.ok(error.suggestions.some(hint => hint.includes("P2 or ID Governance")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "entitlement", "assignment-policy", "count",
        "--access-package", pkg1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("EntitlementManagement.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty entitlement lists stay definitive`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "entitlement", "catalog", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.catalogs, []);
      assert.deepEqual(listed.count, "0 catalogs");
      assert.equal(listed.total, null);
      assert.equal(listed.complete, true);
      assert.ok(listed.help.some(hint => hint.includes("0 entitlement catalogs matched")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts entitlement reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runEntitlementCli(["entra", "entitlement", "catalog", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.catalogs.map(row => row.id), [cat1.id, cat2.id]);
      assert.deepEqual(listOut.count, "2 catalogs");
      assert.equal(listOut.total, null);
      assert.equal(listOut.complete, true);
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runEntitlementCli(["entra", "entitlement", "catalog", "show", "--id", cat1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.ok(decode(shown.stdout).catalog);

      const counted = runEntitlementCli(["entra", "entitlement", "catalog", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).count, { returned: 2, complete: true });

      const packages = runEntitlementCli(["entra", "entitlement", "access-package", "list", "--profile", profile], state, mode);
      assert.equal(packages.status, 0, packages.stdout);
      assert.deepEqual(decode(packages.stdout).accessPackages.map(row => row.id), [pkg1.id, pkg2.id]);

      const packageShown = runEntitlementCli(["entra", "entitlement", "access-package", "show", "--id", pkg1.id, "--profile", profile], state, mode);
      assert.equal(packageShown.status, 0, packageShown.stdout);
      assert.deepEqual(decode(packageShown.stdout).accessPackage, pkg1);

      const policiesListed = runEntitlementCli(["entra", "entitlement", "assignment-policy", "list",
        "--access-package", pkg1.id, "--profile", profile], state, mode);
      assert.equal(policiesListed.status, 0, policiesListed.stdout);
      assert.deepEqual(decode(policiesListed.stdout).assignmentPolicies.map(row => row.id), [pol1.id, pol2.id]);

      const policyShown = runEntitlementCli(["entra", "entitlement", "assignment-policy", "show",
        "--access-package", pkg1.id, "--id", pol1.id, "--profile", profile], state, mode);
      assert.equal(policyShown.status, 0, policyShown.stdout);

      const policyCounted = runEntitlementCli(["entra", "entitlement", "assignment-policy", "count",
        "--access-package", pkg1.id, "--profile", profile], state, mode);
      assert.equal(policyCounted.status, 0, policyCounted.stdout);
      assert.deepEqual(decode(policyCounted.stdout).count, { returned: 2, complete: true });

      const scopesListed = runEntitlementCli(["entra", "entitlement", "resource-role-scope", "list",
        "--access-package", pkg1.id, "--profile", profile], state, mode);
      assert.equal(scopesListed.status, 0, scopesListed.stdout);
      assert.deepEqual(decode(scopesListed.stdout).resourceRoleScopes.map(row => row.id), [rs1.id, rs2.id]);

      const scopeShown = runEntitlementCli(["entra", "entitlement", "resource-role-scope", "show",
        "--access-package", pkg1.id, "--id", rs1.id, "--profile", profile], state, mode);
      assert.equal(scopeShown.status, 0, scopeShown.stdout);
      assert.deepEqual(decode(scopeShown.stdout).resourceRoleScope, rs1);

      const scopeCounted = runEntitlementCli(["entra", "entitlement", "resource-role-scope", "count",
        "--access-package", pkg1.id, "--profile", profile], state, mode);
      assert.equal(scopeCounted.status, 0, scopeCounted.stdout);
      assert.deepEqual(decode(scopeCounted.stdout).count, { returned: 2, complete: true });
      assert.ok(!scopeCounted.stdout.includes(`opaque-fixture-${mode}-token`));

      const assignmentsListed = runEntitlementCli(["entra", "entitlement", "assignment", "list", "--profile", profile], state, mode);
      assert.equal(assignmentsListed.status, 0, assignmentsListed.stdout);
      assert.deepEqual(decode(assignmentsListed.stdout).assignments.map(row => row.id), [asg1.id, asg2.id]);

      const assignmentShown = runEntitlementCli(["entra", "entitlement", "assignment", "show",
        "--id", asg1.id, "--profile", profile], state, mode);
      assert.equal(assignmentShown.status, 0, assignmentShown.stdout);
      assert.deepEqual(decode(assignmentShown.stdout).assignment, {
        id: asg1.id,
        state: "delivered",
        status: "Delivered",
        expiredDateTime: "2024-06-01T00:00:00Z",
        schedule: { startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2024-12-31T00:00:00Z" },
        customExtensionCalloutInstances: [],
        targetId: "99999999-9999-4999-8999-999999999999",
        targetDisplayName: "Alice Example",
        accessPackageId: pkg1.id,
      });

      const assignmentCounted = runEntitlementCli(["entra", "entitlement", "assignment", "count", "--profile", profile], state, mode);
      assert.equal(assignmentCounted.status, 0, assignmentCounted.stdout);
      assert.deepEqual(decode(assignmentCounted.stdout).count, { returned: 2, complete: true });

      const requestsListed = runEntitlementCli(["entra", "entitlement", "assignment-request", "list", "--profile", profile], state, mode);
      assert.equal(requestsListed.status, 0, requestsListed.stdout);
      assert.deepEqual(decode(requestsListed.stdout).assignmentRequests.map(row => row.id), [req1.id, req2.id]);

      const requestShown = runEntitlementCli(["entra", "entitlement", "assignment-request", "show",
        "--id", req1.id, "--profile", profile], state, mode);
      assert.equal(requestShown.status, 0, requestShown.stdout);
      assert.ok(!Object.hasOwn(decode(requestShown.stdout).assignmentRequest, "justification"));

      const requestCounted = runEntitlementCli(["entra", "entitlement", "assignment-request", "count", "--profile", profile], state, mode);
      assert.equal(requestCounted.status, 0, requestCounted.stdout);
      assert.deepEqual(decode(requestCounted.stdout).count, { returned: 2, complete: true });
      assert.ok(!requestCounted.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied entitlement reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runEntitlementCli(["entra", "entitlement", "catalog", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.catalogs, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["entitlement", "catalog", "list"], ["--limit", "1"]],
    [["entitlement", "catalog", "show"], ["--id", cat1.id]],
    [["entitlement", "catalog", "count"], []],
    [["entitlement", "access-package", "list"], ["--limit", "1"]],
    [["entitlement", "access-package", "show"], ["--id", pkg1.id]],
    [["entitlement", "access-package", "count"], []],
    [["entitlement", "assignment-policy", "list"], ["--access-package", pkg1.id]],
    [["entitlement", "assignment-policy", "show"], ["--access-package", pkg1.id, "--id", pol1.id]],
    [["entitlement", "assignment-policy", "count"], ["--access-package", pkg1.id]],
    [["entitlement", "resource-role-scope", "list"], ["--access-package", pkg1.id]],
    [["entitlement", "resource-role-scope", "show"], ["--access-package", pkg1.id, "--id", rs1.id]],
    [["entitlement", "resource-role-scope", "count"], ["--access-package", pkg1.id]],
    [["entitlement", "assignment", "list"], []],
    [["entitlement", "assignment", "show"], ["--id", asg1.id]],
    [["entitlement", "assignment", "count"], []],
    [["entitlement", "assignment-request", "list"], []],
    [["entitlement", "assignment-request", "show"], ["--id", req1.id]],
    [["entitlement", "assignment-request", "count"], []],
  ]) {
    for (const preview of [false, true]) {
      test(`${mode} ${command.join(" ")} refuses beta before credentials with preview=${preview}`, async () => {
        const state = setupProfiles();
        try {
          const path = join(state.dir, "config.json");
          const config = JSON.parse(readFileSync(path, "utf8"));
          config.profiles[profile].preview = preview;
          writeFileSync(path, JSON.stringify(config));
          const { calls, requests, overrides } = overridesFor(mode);
          await assert.rejects(
            executeArgv(["entra", ...command, ...args, "--profile", profile, "--api-version", "beta"], overrides),
            { code: "VALIDATION_ERROR", message: "Entitlement-management reads support v1.0 only; beta needs its own review" },
          );
          assert.equal(calls.length, 0);
          assert.equal(requests.length, 0);
        } finally {
          teardownProfiles(state);
        }
      });
    }
  }
}
