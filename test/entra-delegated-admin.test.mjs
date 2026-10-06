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

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const delegatedAdminScopes = ["https://graph.microsoft.com/DelegatedAdminRelationship.Read.All"];

const longName = `Contoso subsidiary-${" very-important".repeat(40)}`;
const cu1 = {
  displayName: "Contoso Inc",
  id: "4fdbff88-9d6b-42e0-9713-45c922ba8001",
  tenantId: "4fdbff88-9d6b-42e0-9713-45c922ba8001",
};
const cu2 = {
  displayName: longName,
  id: "1c0fa218-5dec-49db-8247-cfa457af8116",
  tenantId: "1c0fa218-5dec-49db-8247-cfa457af8116",
};
const cu3 = { id: "cu-bare-3" };
const customers = [cu1, cu2, cu3];
const rel1 = {
  accessDetails: { unifiedRoles: [{ roleDefinitionId: "729827e3-9c14-49f7-bb1b-9608f156bbb8" }] },
  activatedDateTime: "2022-02-10T11:26:44.9941884Z",
  autoExtendDuration: "P180D",
  createdDateTime: "2022-02-10T11:24:42.3148266Z",
  customer: { tenantId: cu1.tenantId, displayName: "Contoso Inc" },
  displayName: "Contoso admin relationship",
  duration: "P730D",
  endDateTime: "2024-02-10T11:24:42.3148266Z",
  id: "5d027261-d21f-4aa9-b7db-7fa1f56fb163-8777b240-c6f0-4469-9e98-a3205431b836",
  lastModifiedDateTime: "2022-02-10T11:26:44.9941884Z",
  status: "active",
};
const rel2 = {
  displayName: longName,
  id: "1041ef52-a99b-4245-a3be-cbd3fa7c5ed1-8777b240-c6f0-4469-9e98-a3205431b836",
  status: "approvalPending",
};
const rel3 = { id: "rel-bare-3" };
const relationships = [rel1, rel2, rel3];
const aa1 = {
  accessContainer: { accessContainerId: "227a2f44-2682-4831-a021-f8d69a34bcba", accessContainerType: "securityGroup" },
  accessDetails: { unifiedRoles: [{ roleDefinitionId: "88d8e3e3-8f55-4a1e-953a-9b9898b8876b" }] },
  createdDateTime: "2022-03-07T22:55:18.6780449Z",
  id: "84c586df-0943-416e-b95f-7289cb8d3bd5",
  lastModifiedDateTime: "2022-03-11T23:50:35.8970153Z",
  status: "active",
};
const aa2 = {
  accessContainer: { accessContainerId: "869713c9-0b28-4d08-8949-ae07ae1bf528", accessContainerType: "securityGroup" },
  id: "8d56bce3-440f-4b4f-b5c2-cc0bcbd0199c",
  status: "pending",
};
const aa3 = { id: "aa-bare-3" };
const assignmentRows = [aa1, aa2, aa3];
const longData = `{"id":"a97a9b4c-f43e-4c47-bbd6-50d8d3c88d94","status":"active","padding":"${"p".repeat(600)}"}`;
const op1 = {
  createdDateTime: "2022-02-09T22:17:43.9821847Z",
  data: '{"id":"a97a9b4c-f43e-4c47-bbd6-50d8d3c88d94","status":"active"}',
  id: "e7de9158-df46-478e-820c-d6eff099d27b",
  lastModifiedDateTime: "2022-02-09T22:17:43.9821847Z",
  operationType: "delegatedAdminAccessAssignmentUpdate",
  status: "succeeded",
};
const op2 = {
  createdDateTime: "2022-02-11T19:27:31.4047395Z",
  data: longData,
  id: "f7a7dad4-8cc4-40d7-be44-dd3501b1f4e0",
  lastModifiedDateTime: "2022-02-11T19:27:31.4047395Z",
  operationType: "delegatedAdminAccessAssignmentUpdate",
  status: "running",
};
const op3 = { id: "op-bare-3" };
const operationRows = [op1, op2, op3];
const rq1 = {
  action: "lockForApproval",
  createdDateTime: "2022-02-01T06:14:55.5398865Z",
  id: "ae5a6b9e-6355-43dd-b708-48486b69c3ff",
  lastModifiedDateTime: "2022-02-01T06:14:55.5398865Z",
  status: "succeeded",
};
const rq2 = { action: "terminate", id: "8a1b6676-5c12-47ba-8d3a-1d38387b0909", status: "running" };
const rq3 = { id: "rq-bare-3" };
const requestRows = [rq1, rq2, rq3];
const longUrl = `https://admin.teams.microsoft.com/?delegatedOrg=contoso.com&padding=${"q".repeat(500)}`;
const smd1 = { id: "fa5fa04e-13df-4b7c-9e99-92573ba1fa55", serviceManagementUrl: "https://lighthouse.microsoft.com", serviceName: "Microsoft 365 Lighthouse" };
const smd2 = { id: "ce0b42f4-bfde-4abe-a5f7-add83f104b23", serviceManagementUrl: longUrl, serviceName: "Teams" };
const smd3 = { id: "smd-bare-3" };
const detailRows = [smd1, smd2, smd3];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-delegated-admin-"));
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

function delegatedAdminTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/tenantRelationships/delegatedAdminCustomers/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(customers.length) };
    }
    if (path === "/v1.0/tenantRelationships/delegatedAdminRelationships/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(relationships.length) };
    }
    if (/^\/v1\.0\/tenantRelationships\/delegatedAdminCustomers\/[^/]+\/serviceManagementDetails\/\$count$/.test(path)) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(detailRows.length) };
    }
    if (/^\/v1\.0\/tenantRelationships\/delegatedAdminRelationships\/[^/]+\/(accessAssignments|operations|requests)\/\$count$/.test(path)) {
      const segment = path.includes("/accessAssignments/") ? assignmentRows : path.includes("/operations/") ? operationRows : requestRows;
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(segment.length) };
    }
    if (path === "/v1.0/tenantRelationships/delegatedAdminCustomers") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [cu3] });
      return json(200, {
        value: [cu1, cu2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/tenantRelationships/delegatedAdminRelationships") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [rel3] });
      return json(200, {
        value: [rel1, rel2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships?%24skiptoken=page2",
      });
    }
    const customer = /^\/v1\.0\/tenantRelationships\/delegatedAdminCustomers\/([^/]+)$/.exec(path);
    if (customer) {
      const found = customers.find(row => row.id === decodeURIComponent(customer[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such customer" } });
    }
    const relationship = /^\/v1\.0\/tenantRelationships\/delegatedAdminRelationships\/([^/]+)$/.exec(path);
    if (relationship) {
      const found = relationships.find(row => row.id === decodeURIComponent(relationship[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such relationship" } });
    }
    const navigation = /^\/v1\.0\/tenantRelationships\/delegatedAdminRelationships\/([^/]+)\/(accessAssignments|operations|requests)(\/([^/]+))?$/.exec(path);
    if (navigation) {
      const rows = navigation[2] === "accessAssignments" ? assignmentRows : navigation[2] === "operations" ? operationRows : requestRows;
      if (navigation[4] !== undefined) {
        const found = rows.find(row => row.id === decodeURIComponent(navigation[4]));
        return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such navigation row" } });
      }
      if (navigation[2] === "accessAssignments") {
        if (url.searchParams.has("$skiptoken")) return json(200, { value: [aa3] });
        return json(200, {
          value: [aa1, aa2],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/${encodeURIComponent(navigation[1])}/accessAssignments?%24skiptoken=page2`,
        });
      }
      return json(200, { value: rows });
    }
    const serviceDetails = /^\/v1\.0\/tenantRelationships\/delegatedAdminCustomers\/([^/]+)\/serviceManagementDetails(\/([^/]+))?$/.exec(path);
    if (serviceDetails) {
      if (serviceDetails[3] !== undefined) {
        const found = detailRows.find(row => row.id === decodeURIComponent(serviceDetails[3]));
        return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such service-management detail" } });
      }
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [smd3] });
      return json(200, {
        value: [smd1, smd2],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers/${encodeURIComponent(serviceDetails[1])}/serviceManagementDetails?%24skiptoken=page2`,
      });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? delegatedAdminTransport();
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

function runDelegatedAdminCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-delegated-admin-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, customers: [cu1, cu2], relationships: [rel1, rel2],
        assignments: [aa1, aa2], operations: [op1, op2], requests: [rq1, rq2], details: [smd1, smd2], denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists delegated-admin customers with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-customer", "list", "--profile", profile,
        "--select", "id,displayName,tenantId"], overrides);
      assert.deepEqual(result.delegatedAdminCustomers, [
        { id: cu1.id, displayName: "Contoso Inc", tenantId: cu1.tenantId },
        { id: cu2.id, displayName: longName.slice(0, 500) + `... (truncated, ${longName.length} chars total)`, tenantId: cu2.tenantId },
        { id: cu3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra delegated-admin-customer show --id <customer-id>")));
      assert.ok(result.help.some(hint => hint.includes("partner tenant")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(delegatedAdminScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists delegated-admin relationships with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-relationship", "list", "--profile", profile,
        "--select", "id,displayName,status,customer,endDateTime"], overrides);
      assert.deepEqual(result.delegatedAdminRelationships, [
        { id: rel1.id, displayName: "Contoso admin relationship", status: "active", customer: rel1.customer, endDateTime: rel1.endDateTime },
        { id: rel2.id, displayName: longName.slice(0, 500) + `... (truncated, ${longName.length} chars total)`, status: "approvalPending" },
        { id: rel3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra delegated-admin-relationship show --id <relationship-id>")));
      assert.ok(result.help.some(hint => hint.includes("partner tenant")));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships?"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists delegated-admin customers with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-customer", "list", "--profile", profile,
        "--filter", "displayName eq 'Contoso Inc'"], overrides);
      assert.equal(result.count.returned, 3);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "displayName eq 'Contoso Inc'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped delegated-admin customer list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "delegated-admin-customer", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.delegatedAdminCustomers.map(row => row.id), [cu1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "delegated-admin-customer", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.delegatedAdminCustomers.map(row => row.id), [cu2.id, cu3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped delegated-admin relationship list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "delegated-admin-relationship", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.delegatedAdminRelationships.map(row => row.id), [rel1.id]);
      assert.equal(first.count.complete, false);
      const second = await executeArgv(["entra", "delegated-admin-relationship", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.delegatedAdminRelationships.map(row => row.id), [rel2.id, rel3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists relationship access assignments with compact rows bound to the parent`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-relationship", "list-access-assignments", "--id", rel1.id, "--profile", profile,
        "--select", "id,status,accessContainer,accessDetails"], overrides);
      assert.deepEqual(result.delegatedAdminAccessAssignments, [
        { id: aa1.id, status: "active", accessContainer: aa1.accessContainer, accessDetails: aa1.accessDetails },
        { id: aa2.id, status: "pending", accessContainer: aa2.accessContainer },
        { id: aa3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra delegated-admin-relationship show-access-assignment --id <relationship-id> --assignment-id <assignment-id>")));
      assert.ok(result.help.some(hint => hint.includes("partner tenant")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/${rel1.id}/accessAssignments?`));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists relationship operations with compact rows and truncates the data payload`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-relationship", "list-operations", "--id", rel1.id, "--profile", profile,
        "--select", "id,operationType,status,lastModifiedDateTime,data"], overrides);
      assert.deepEqual(result.delegatedAdminRelationshipOperations.map(row => row.id), [op1.id, op2.id, op3.id]);
      assert.equal(result.delegatedAdminRelationshipOperations[0].data, op1.data);
      assert.match(result.delegatedAdminRelationshipOperations[1].data, /truncated, \d+ chars total/);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra delegated-admin-relationship show-operation")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/${rel1.id}/operations?`));
      const full = await executeArgv(["entra", "delegated-admin-relationship", "list-operations", "--id", rel1.id, "--profile", profile, "--full",
        "--select", "id,data"], overrides);
      assert.equal(full.delegatedAdminRelationshipOperations[1].data, longData);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists relationship requests with compact rows`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-relationship", "list-requests", "--id", rel1.id, "--profile", profile,
        "--select", "id,action,status,lastModifiedDateTime"], overrides);
      assert.deepEqual(result.delegatedAdminRelationshipRequests, [
        { id: rq1.id, action: "lockForApproval", status: "succeeded", lastModifiedDateTime: rq1.lastModifiedDateTime },
        { id: rq2.id, action: "terminate", status: "running" },
        { id: rq3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra delegated-admin-relationship show-request")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/${rel1.id}/requests?`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists service-management details whole with no query parameters`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-customer", "list-service-management-details", "--id", cu1.id, "--profile", profile], overrides);
      assert.deepEqual(result.delegatedAdminServiceManagementDetails, [
        smd1,
        { ...smd2, serviceManagementUrl: smd2.serviceManagementUrl.slice(0, 500) + `... (truncated, ${smd2.serviceManagementUrl.length} chars total)` },
        smd3,
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra delegated-admin-customer show-service-management-detail --id <customer-id> --detail-id <detail-id>")));
      const sent = new URL(requests[0].url);
      assert.equal(sent.pathname, `/v1.0/tenantRelationships/delegatedAdminCustomers/${cu1.id}/serviceManagementDetails`);
      assert.equal(sent.search, "");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists relationship access assignments with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-relationship", "list-access-assignments", "--id", rel1.id, "--profile", profile,
        "--filter", "status eq 'active'"], overrides);
      assert.equal(result.count.returned, 3);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "status eq 'active'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped access-assignment list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "delegated-admin-relationship", "list-access-assignments", "--id", rel1.id, "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.delegatedAdminAccessAssignments.map(row => row.id), [aa1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "delegated-admin-relationship", "list-access-assignments", "--id", rel1.id, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.delegatedAdminAccessAssignments.map(row => row.id), [aa2.id, aa3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped service-management-detail list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "delegated-admin-customer", "list-service-management-details", "--id", cu1.id, "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.delegatedAdminServiceManagementDetails.map(row => row.id), [smd1.id]);
      assert.equal(first.count.complete, false);
      const second = await executeArgv(["entra", "delegated-admin-customer", "list-service-management-details", "--id", cu1.id, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.delegatedAdminServiceManagementDetails.map(row => row.id), [smd2.id, smd3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
      // The only query state ever sent is the server's own $skiptoken from
      // its nextLink; $select/$filter never leave this client.
      assert.ok(requests.every(request => {
        const params = new URL(request.url).searchParams;
        return !params.has("$select") && !params.has("$filter");
      }));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one access assignment with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-relationship", "show-access-assignment", "--id", rel1.id, "--assignment-id", aa1.id, "--profile", profile,
        "--select", "accessContainer,accessDetails,createdDateTime,id,lastModifiedDateTime,status"], overrides);
      assert.deepEqual(result.delegatedAdminAccessAssignment, aa1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one relationship operation with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-relationship", "show-operation", "--id", rel1.id, "--operation-id", op1.id, "--profile", profile,
        "--select", "createdDateTime,data,id,lastModifiedDateTime,operationType,status"], overrides);
      assert.deepEqual(result.delegatedAdminRelationshipOperation, op1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one relationship request with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-relationship", "show-request", "--id", rel1.id, "--request-id", rq1.id, "--profile", profile,
        "--select", "action,createdDateTime,id,lastModifiedDateTime,status"], overrides);
      assert.deepEqual(result.delegatedAdminRelationshipRequest, rq1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one service-management detail whole and projects fields locally`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-customer", "show-service-management-detail", "--id", cu1.id, "--detail-id", smd1.id, "--profile", profile], overrides);
      assert.deepEqual(result.delegatedAdminServiceManagementDetail, smd1);
      assert.equal(result.help, undefined);
      assert.equal(new URL(requests[0].url).search, "");
      const projected = await executeArgv(["entra", "delegated-admin-customer", "show-service-management-detail", "--id", cu1.id, "--detail-id", smd1.id, "--profile", profile,
        "--fields", "serviceName"], overrides);
      assert.deepEqual(projected.delegatedAdminServiceManagementDetail, { serviceName: smd1.serviceName });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown navigation ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "show-access-assignment", "--id", rel1.id, "--assignment-id", "aa-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "show-service-management-detail", "--id", cu1.id, "--detail-id", "smd-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied navigation reads surface scope and partner guidance`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list-access-assignments", "--id", rel1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("DelegatedAdminRelationship.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("partner tenant")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "show-service-management-detail", "--id", cu1.id, "--detail-id", smd1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("partner tenant")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists and shows navigation reads end to end`, () => {
    const state = setupProfiles();
    try {
      const listed = runDelegatedAdminCli(["entra", "delegated-admin-relationship", "list-access-assignments", "--id", rel1.id, "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      assert.deepEqual(decode(listed.stdout).delegatedAdminAccessAssignments.map(row => row.id), [aa1.id, aa2.id]);

      const shown = runDelegatedAdminCli(["entra", "delegated-admin-relationship", "show-operation", "--id", rel1.id, "--operation-id", op1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).delegatedAdminRelationshipOperation, op1);

      const listedDetails = runDelegatedAdminCli(["entra", "delegated-admin-customer", "list-service-management-details", "--id", cu1.id, "--profile", profile], state, mode);
      assert.equal(listedDetails.status, 0, listedDetails.stdout);
      assert.deepEqual(decode(listedDetails.stdout).delegatedAdminServiceManagementDetails.map(row => row.id), [smd1.id, smd2.id]);

      const shownDetail = runDelegatedAdminCli(["entra", "delegated-admin-customer", "show-service-management-detail", "--id", cu1.id, "--detail-id", smd1.id, "--profile", profile], state, mode);
      assert.equal(shownDetail.status, 0, shownDetail.stdout);
      assert.deepEqual(decode(shownDetail.stdout).delegatedAdminServiceManagementDetail, smd1);
      assert.ok(!shownDetail.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} shows one delegated-admin customer with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-customer", "show", "--id", cu1.id, "--profile", profile,
        "--select", "displayName,id,tenantId"], overrides);
      assert.deepEqual(result.delegatedAdminCustomer, cu1);
      assert.equal(result.help, undefined);
      const missing = await executeArgv(["entra", "delegated-admin-customer", "show", "--id", cu3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.delegatedAdminCustomer, { id: cu3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one delegated-admin relationship with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "delegated-admin-relationship", "show", "--id", rel1.id, "--profile", profile,
        "--select", "accessDetails,activatedDateTime,autoExtendDuration,createdDateTime,customer,displayName,duration,endDateTime,id,lastModifiedDateTime,status"], overrides);
      assert.deepEqual(result.delegatedAdminRelationship, rel1);
      assert.equal(result.help, undefined);
      const missing = await executeArgv(["entra", "delegated-admin-relationship", "show", "--id", rel3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.delegatedAdminRelationship, { id: rel3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty delegated-admin customer lists stay definitive with partner context`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const result = await executeArgv(["entra", "delegated-admin-customer", "list", "--profile", profile], overrides);
      assert.deepEqual(result.delegatedAdminCustomers, []);
      assert.deepEqual(result.count, { returned: 0, complete: true });
      assert.ok(result.help.some(hint => hint.includes("0 customers matched")));
      assert.ok(result.help.some(hint => hint.includes("partner tenant")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown delegated-admin ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "show", "--id", "cu-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "show", "--id", "rel-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied delegated-admin reads surface scope, partner and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("DelegatedAdminRelationship.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("partner tenant")));
        assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "show", "--id", rel1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("partner tenant")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts delegated-admin customers and relationships as scalars`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const customers = await executeArgv(["entra", "delegated-admin-customer", "count", "--profile", profile], overrides);
      assert.deepEqual(customers, { delegatedAdminCustomerCount: 3 });
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers/$count");
      const relationships = await executeArgv(["entra", "delegated-admin-relationship", "count", "--profile", profile], overrides);
      assert.deepEqual(relationships, { delegatedAdminRelationshipCount: 3 });
      assert.equal(requests[1].url, "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$count");
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(delegatedAdminScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts delegated-admin navigation collections as scalars bound to the parent`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const details = await executeArgv(["entra", "delegated-admin-customer", "count-service-management-details", "--id", cu1.id, "--profile", profile], overrides);
      assert.deepEqual(details, { delegatedAdminServiceManagementDetailCount: 3 });
      assert.equal(requests[0].url, `https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers/${cu1.id}/serviceManagementDetails/$count`);
      const assignments = await executeArgv(["entra", "delegated-admin-relationship", "count-access-assignments", "--id", rel1.id, "--profile", profile], overrides);
      assert.deepEqual(assignments, { delegatedAdminAccessAssignmentCount: 3 });
      assert.equal(requests[1].url, `https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/${encodeURIComponent(rel1.id)}/accessAssignments/$count`);
      const operations = await executeArgv(["entra", "delegated-admin-relationship", "count-operations", "--id", rel1.id, "--profile", profile], overrides);
      assert.deepEqual(operations, { delegatedAdminRelationshipOperationCount: 3 });
      assert.equal(requests[2].url, `https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/${encodeURIComponent(rel1.id)}/operations/$count`);
      const relationshipRequests = await executeArgv(["entra", "delegated-admin-relationship", "count-requests", "--id", rel1.id, "--profile", profile], overrides);
      assert.deepEqual(relationshipRequests, { delegatedAdminRelationshipRequestCount: 3 });
      assert.equal(requests[3].url, `https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/${encodeURIComponent(rel1.id)}/requests/$count`);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied delegated-admin counts surface scope and partner guidance`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("DelegatedAdminRelationship.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("partner tenant")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "count-access-assignments", "--id", rel1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("partner tenant")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable delegated-admin counts end to end`, () => {
    const state = setupProfiles();
    try {
      const customerCount = runDelegatedAdminCli(["entra", "delegated-admin-customer", "count", "--profile", profile], state, mode);
      assert.equal(customerCount.status, 0, customerCount.stdout);
      assert.deepEqual(decode(customerCount.stdout).delegatedAdminCustomerCount, 2);
      const detailsCount = runDelegatedAdminCli(["entra", "delegated-admin-customer", "count-service-management-details", "--id", cu1.id, "--profile", profile], state, mode);
      assert.equal(detailsCount.status, 0, detailsCount.stdout);
      assert.deepEqual(decode(detailsCount.stdout).delegatedAdminServiceManagementDetailCount, 2);
      const relationshipCount = runDelegatedAdminCli(["entra", "delegated-admin-relationship", "count", "--profile", profile], state, mode);
      assert.equal(relationshipCount.status, 0, relationshipCount.stdout);
      assert.deepEqual(decode(relationshipCount.stdout).delegatedAdminRelationshipCount, 2);
      const assignmentsCount = runDelegatedAdminCli(["entra", "delegated-admin-relationship", "count-access-assignments", "--id", rel1.id, "--profile", profile], state, mode);
      assert.equal(assignmentsCount.status, 0, assignmentsCount.stdout);
      assert.deepEqual(decode(assignmentsCount.stdout).delegatedAdminAccessAssignmentCount, 2);
      const operationsCount = runDelegatedAdminCli(["entra", "delegated-admin-relationship", "count-operations", "--id", rel1.id, "--profile", profile], state, mode);
      assert.equal(operationsCount.status, 0, operationsCount.stdout);
      assert.deepEqual(decode(operationsCount.stdout).delegatedAdminRelationshipOperationCount, 2);
      const requestsCount = runDelegatedAdminCli(["entra", "delegated-admin-relationship", "count-requests", "--id", rel1.id, "--profile", profile], state, mode);
      assert.equal(requestsCount.status, 0, requestsCount.stdout);
      assert.deepEqual(decode(requestsCount.stdout).delegatedAdminRelationshipRequestCount, 2);
      assert.ok(!requestsCount.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable lists and shows delegated-admin customers and relationships`, () => {
    const state = setupProfiles();
    try {
      const listed = runDelegatedAdminCli(["entra", "delegated-admin-customer", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.delegatedAdminCustomers.map(row => row.id), [cu1.id, cu2.id]);
      assert.deepEqual(listOut.count, { returned: 2, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runDelegatedAdminCli(["entra", "delegated-admin-customer", "show", "--id", cu1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).delegatedAdminCustomer, cu1);

      const listedRel = runDelegatedAdminCli(["entra", "delegated-admin-relationship", "list", "--profile", profile], state, mode);
      assert.equal(listedRel.status, 0, listedRel.stdout);
      assert.deepEqual(decode(listedRel.stdout).delegatedAdminRelationships.map(row => row.id), [rel1.id, rel2.id]);

      const shownRel = runDelegatedAdminCli(["entra", "delegated-admin-relationship", "show", "--id", rel1.id, "--profile", profile], state, mode);
      assert.equal(shownRel.status, 0, shownRel.stdout);
      assert.deepEqual(decode(shownRel.stdout).delegatedAdminRelationship, rel1);
      assert.ok(!shownRel.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied delegated-admin reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runDelegatedAdminCli(["entra", "delegated-admin-customer", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.delegatedAdminCustomers, undefined);
      assert.ok(output.help.some(hint => hint.includes("partner tenant")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["delegated-admin-customer", "list"], ["--limit", "1"]],
    [["delegated-admin-customer", "show"], ["--id", cu1.id]],
    [["delegated-admin-relationship", "list"], ["--limit", "1"]],
    [["delegated-admin-relationship", "show"], ["--id", rel1.id]],
    [["delegated-admin-relationship", "list-access-assignments"], ["--id", rel1.id, "--limit", "1"]],
    [["delegated-admin-relationship", "show-access-assignment"], ["--id", rel1.id, "--assignment-id", aa1.id]],
    [["delegated-admin-relationship", "list-operations"], ["--id", rel1.id, "--limit", "1"]],
    [["delegated-admin-relationship", "show-operation"], ["--id", rel1.id, "--operation-id", op1.id]],
    [["delegated-admin-relationship", "list-requests"], ["--id", rel1.id, "--limit", "1"]],
    [["delegated-admin-relationship", "show-request"], ["--id", rel1.id, "--request-id", rq1.id]],
    [["delegated-admin-customer", "list-service-management-details"], ["--id", cu1.id, "--limit", "1"]],
    [["delegated-admin-customer", "show-service-management-detail"], ["--id", cu1.id, "--detail-id", smd1.id]],
    [["delegated-admin-customer", "count"], []],
    [["delegated-admin-customer", "count-service-management-details"], ["--id", cu1.id]],
    [["delegated-admin-relationship", "count"], []],
    [["delegated-admin-relationship", "count-access-assignments"], ["--id", rel1.id]],
    [["delegated-admin-relationship", "count-operations"], ["--id", rel1.id]],
    [["delegated-admin-relationship", "count-requests"], ["--id", rel1.id]],
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
            { code: "VALIDATION_ERROR", message: "Delegated-admin reads support v1.0 only; beta needs its own review" },
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

test("truncated delegated-admin text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const partial = await executeArgv(["entra", "delegated-admin-customer", "list", "--profile", "soc",
      "--select", "id,displayName"], overrides);
    const truncated = partial.delegatedAdminCustomers.find(row => row.id === cu2.id);
    assert.match(truncated.displayName, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "delegated-admin-customer", "list", "--profile", "soc", "--full",
      "--select", "id,displayName"], overrides);
    assert.equal(full.delegatedAdminCustomers.find(row => row.id === cu2.id).displayName, longName);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "delegated-admin-customer", "list", "--profile", "batch", "--scopes", delegatedAdminScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "delegated-admin-relationship", "show", "--id", rel1.id, "--profile", "batch", "--scopes", delegatedAdminScopes[0]], overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("unknown properties and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "list", "--profile", "soc", "--select", "id,owner"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "list", "--profile", "soc", "--select", "id,displayName", "--fields", "tenantId"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list", "--profile", "soc", "--fields", "duration"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list", "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list-access-assignments", "--id", rel1.id, "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list-access-assignments", "--id", rel1.id, "--profile", "soc", "--select", "id,status", "--fields", "createdDateTime"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "list-service-management-details", "--id", cu1.id, "--profile", "soc", "--select", "id"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "list-service-management-details", "--id", cu1.id, "--profile", "soc", "--filter", "serviceName eq 'Teams'"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "show-service-management-detail", "--id", cu1.id, "--detail-id", smd1.id, "--profile", "soc", "--select", "id"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "show-service-management-detail", "--id", cu1.id, "--detail-id", smd1.id, "--profile", "soc", "--fields", "zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated-admin read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list-access-assignments"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "show-access-assignment", "--id", rel1.id]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list-access-assignments", "--bogus"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "show-operation", "--id", rel1.id]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "list-service-management-details"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "show-service-management-detail", "--id", cu1.id]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "count", "--limit", "5"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "count", "--filter", "displayName eq 'x'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "count", "--select", "id"]), /unknown flag --select/);
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "count", "--cursor", "x"]), /unknown flag --cursor/);
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "count-service-management-details"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "count-access-assignments"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "count-operations", "--id", rel1.id, "--limit", "5"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "count-requests", "--id", rel1.id, "--filter", "status eq 'succeeded'"]), /unknown flag --filter/);
});

test("malformed delegated-admin count bodies fail as unknown, not zero", async () => {
  const state = setupProfiles();
  try {
    for (const body of ["{}", "-1", "2.5", "\"3\""]) {
      const malformed = transport(() => ({ status: 200, headers: {}, body }));
      const scoped = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "count", "--profile", "soc"], scoped.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /malformed delegated-admin customer count|non-JSON success body/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "count-access-assignments", "--id", rel1.id, "--profile", "soc"], scoped.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /malformed delegated-admin access assignment count|non-JSON success body/.test(error.message);
      });
    }
  } finally {
    teardownProfiles(state);
  }
});
