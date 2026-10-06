import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";

const r1 = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  deletedDateTime: null,
  description: "Can manage all aspects of Microsoft Entra ID.",
  displayName: "Global Administrator",
  roleTemplateId: "62e90394-69f5-4237-9190-012177145e10",
};
const r2 = {
  id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  deletedDateTime: null,
  description: `Can reset passwords for non-administrators and Helpdesk Administrators. ${"x".repeat(600)}`,
  displayName: "Helpdesk Administrator",
  roleTemplateId: "729827e3-9c14-49f7-bb1b-9608f156bbb8",
};
const r3 = {
  id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  displayName: "Directory Readers",
};
const roles = [r1, r2, r3];

const a1 = {
  id: "lAPpYvVpN0KRkAEhdxReEMmO4KwRqtpKkUWt3wOYIz4-1",
  principalId: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
  roleDefinitionId: "62e90394-69f5-4237-9190-012177145e10",
  directoryScopeId: "/",
  appScopeId: null,
};
const assignments = [a1];

const e1 = {
  id: "8MYkhImhnkm70CbBdTyW1BbHHAdHgZdDpbqyEFlRzAs-1-e",
  principalId: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
  roleDefinitionId: "8424c6f0-a189-499e-bbd0-26c1753c96d4",
  directoryScopeId: "/",
  appScopeId: null,
  memberType: "Direct",
  startDateTime: "2024-01-01T00:00:00Z",
  endDateTime: "2025-01-01T00:00:00Z",
};
const eligible = [e1];

const vAssigned = {
  id: "lAPpYvVpN0KRkAEhdxReEAWz5Gtet_xOv8wxvTtTpfg-1",
  principalId: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
  roleDefinitionId: "62e90394-69f5-4237-9190-012177145e10",
  directoryScopeId: "/",
  appScopeId: null,
  assignmentType: "Assigned",
  memberType: "Direct",
  startDateTime: null,
  endDateTime: null,
};
const vActivated = {
  id: "lAPpYvVpN0KRkAEhdxReEBLS8lac5ONCgpgBiOW-8JQ-1",
  principalId: "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee",
  roleDefinitionId: "8424c6f0-a189-499e-bbd0-26c1753c96d4",
  directoryScopeId: "/",
  appScopeId: null,
  assignmentType: "Activated",
  memberType: "Direct",
  startDateTime: "2024-06-01T00:00:00Z",
  endDateTime: "2024-06-01T08:00:00Z",
};
const active = [vAssigned, vActivated];

const t1 = {
  id: "62e90394-69f5-4237-9190-012177145e10",
  displayName: "Global Administrator",
  description: "Can manage all aspects of Microsoft Entra ID.",
};
const t2 = {
  id: "729827e3-9c14-49f7-bb1b-9608f156bbb8",
  displayName: "Helpdesk Administrator",
  description: "Can reset passwords for non-administrators and Helpdesk Administrators.",
};
const templates = [t1, t2];

const m1 = {
  "@odata.type": "#microsoft.graph.user",
  id: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
  displayName: "Adele Vance",
  mail: "adele@contoso.example",
  userPrincipalName: "adele@contoso.example",
};
const m2 = {
  "@odata.type": "#microsoft.graph.group",
  id: "ffffffff-ffff-4fff-8fff-ffffffffffff",
  displayName: "Ops Admins",
};
const members = [m1, m2];

const s1 = {
  id: "11111111-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  principalId: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
  roleId: "62e90394-69f5-4237-9190-012177145e10",
  directoryScopeId: "/administrativeUnits/aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  administrativeUnitId: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
};
const scoped = [s1];

const g1 = {
  id: "22222222-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  principalId: "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee",
  roleId: "729827e3-9c14-49f7-bb1b-9608f156bbb8",
  directoryScopeId: "/administrativeUnits/bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  administrativeUnitId: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
};
const memberships = [g1];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-09-"));
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
      calls.push(args);
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

function roleTransport(denied = false) {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (denied) return json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } });
    if (path === "/v1.0/directoryRoles") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [r3] });
      return json(200, {
        value: [r1, r2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/directoryRoles?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/directoryRoles\/([^/]+)$/.exec(path);
    if (single) {
      const found = roles.find(role => role.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such role" } });
    }
    if (path === "/v1.0/roleManagement/directory/roleAssignments") {
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("createdDateTime"));
      return json(200, { value: assignments });
    }
    if (path === "/v1.0/roleManagement/directory/roleEligibilityScheduleInstances") {
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("assignmentType"));
      return json(200, { value: eligible });
    }
    if (path === "/v1.0/roleManagement/directory/roleAssignmentScheduleInstances") return json(200, { value: active });
    if (path === "/v1.0/directoryRoleTemplates") return json(200, { value: templates });
    const templateSingle = /^\/v1\.0\/directoryRoleTemplates\/([^/]+)$/.exec(path);
    if (templateSingle) {
      const found = templates.find(template => template.id === decodeURIComponent(templateSingle[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such template" } });
    }
    if (/^\/v1\.0\/directoryRoles\/[^/]+\/members$/.test(path)) {
      assert.equal(url.searchParams.get("$select"), "id,displayName");
      return json(200, { value: members });
    }
    if (/^\/v1\.0\/directoryRoles\/[^/]+\/scopedMembers$/.test(path)) return json(200, { value: scoped });
    const scopedSingle = /^\/v1\.0\/directoryRoles\/[^/]+\/scopedMembers\/([^/]+)$/.exec(path);
    if (scopedSingle) {
      const found = scoped.find(row => row.id === decodeURIComponent(scopedSingle[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such scoped member" } });
    }
    if (path === "/v1.0/scopedRoleMemberships") return json(200, { value: memberships });
    const membershipSingle = /^\/v1\.0\/scopedRoleMemberships\/([^/]+)$/.exec(path);
    if (membershipSingle) {
      const found = memberships.find(row => row.id === decodeURIComponent(membershipSingle[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such membership" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? roleTransport();
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

function runRolesCli(args, state, mode, scopes) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-roles-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, scopes, roles, assignments, eligible, active, templates, members, scoped, memberships }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists activated directory roles with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-role", "list", "--profile", profile], overrides);
      assert.deepEqual(result.directoryRoles, [
        { id: r1.id, displayName: "Global Administrator", description: r1.description, roleTemplateId: r1.roleTemplateId },
        { id: r2.id, displayName: "Helpdesk Administrator",
          description: `${r2.description.slice(0, 500)}... (truncated, ${r2.description.length} chars total)`,
          roleTemplateId: r2.roleTemplateId },
        { id: r3.id, displayName: "Directory Readers" },
      ]);
      assert.deepEqual(result.count, "3 directory roles");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("activated roles only")));
      assert.ok(result.help.some(hint => hint.includes("entra directory-role show --id <role-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/directoryRoles?"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a directory role with richer properties`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-role", "show", "--id", r1.id, "--profile", profile], overrides);
      assert.deepEqual(result.directoryRole, r1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists direct and PIM-activated role assignments in current inventory`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode, transport(() => json(200, { value: [a1, {
        id: vActivated.id,
        principalId: vActivated.principalId,
        roleDefinitionId: vActivated.roleDefinitionId,
        directoryScopeId: vActivated.directoryScopeId,
        appScopeId: vActivated.appScopeId,
      }] })));
      const result = await executeArgv(["entra", "role-assignment", "list", "--profile", profile], overrides);
      assert.deepEqual(result.roleAssignments, [
        { id: a1.id, principalId: a1.principalId, roleDefinitionId: a1.roleDefinitionId, directoryScopeId: "/" },
        { id: vActivated.id, principalId: vActivated.principalId, roleDefinitionId: vActivated.roleDefinitionId, directoryScopeId: "/" },
      ]);
      assert.deepEqual(result.count, "2 role assignments");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("pim active list for assignmentType Assigned versus Activated")));
      assert.ok(result.help.some(hint => hint.includes("roleTemplateId")));
      assert.ok(new URL(requests[0].url).pathname.endsWith("/roleManagement/directory/roleAssignments"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists eligible assignments as not active`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "pim", "eligible", "list", "--profile", profile], overrides);
      assert.deepEqual(result.eligibleAssignments, [
        { id: e1.id, principalId: e1.principalId, roleDefinitionId: e1.roleDefinitionId, memberType: "Direct" },
      ]);
      assert.deepEqual(result.count, "1 eligible assignments");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(new URL(requests[0].url).pathname.endsWith("/roleManagement/directory/roleEligibilityScheduleInstances"));
      assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,principalId,roleDefinitionId,memberType");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists active assignments with Assigned and Activated distinct`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "pim", "active", "list", "--profile", profile], overrides);
      assert.deepEqual(result.activeAssignments, [
        { id: vAssigned.id, principalId: vAssigned.principalId, roleDefinitionId: vAssigned.roleDefinitionId, assignmentType: "Assigned", memberType: "Direct" },
        { id: vActivated.id, principalId: vActivated.principalId, roleDefinitionId: vActivated.roleDefinitionId, assignmentType: "Activated", memberType: "Direct" },
      ]);
      assert.deepEqual(result.count, "2 active assignments");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("assignmentType Activated")));
      assert.ok(new URL(requests[0].url).pathname.endsWith("/roleManagement/directory/roleAssignmentScheduleInstances"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists directory-role templates with compact rows`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-role-template", "list", "--profile", profile], overrides);
      assert.deepEqual(result.directoryRoleTemplates, [t1, t2]);
      assert.deepEqual(result.count, "2 directory-role templates");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("activation state lives on directoryRoles")));
      assert.ok(result.help.some(hint => hint.includes("entra directory-role-template show --id <template-id>")));
      assert.ok(new URL(requests[0].url).pathname.endsWith("/directoryRoleTemplates"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a directory-role template`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-role-template", "show", "--id", t1.id, "--profile", profile], overrides);
      assert.deepEqual(result.directoryRoleTemplate, t1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists directory-role members preserving each member kind`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-role", "member", "list", "--role", r1.id, "--profile", profile], overrides);
      assert.deepEqual(result.roleMembers, [
        { "@odata.type": "#microsoft.graph.user", id: m1.id, displayName: "Adele Vance" },
        { "@odata.type": "#microsoft.graph.group", id: m2.id, displayName: "Ops Admins" },
      ]);
      assert.deepEqual(result.count, "2 role members");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("@odata.type names each member kind")));
      const url = new URL(requests[0].url);
      assert.equal(url.pathname, `/v1.0/directoryRoles/${r1.id}/members`);
      assert.equal(url.searchParams.get("$select"), "id,displayName");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists directory-role scoped members for one role`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-role", "scoped-member", "list", "--role", r1.id, "--profile", profile], overrides);
      assert.deepEqual(result.scopedMembers, [
        { id: s1.id, principalId: s1.principalId, roleId: s1.roleId, directoryScopeId: s1.directoryScopeId },
      ]);
      assert.deepEqual(result.count, "1 scoped members");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/directoryRoles/${r1.id}/scopedMembers`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a directory-role scoped member`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "directory-role", "scoped-member", "show",
        "--role", r1.id, "--id", s1.id, "--profile", profile], overrides);
      assert.deepEqual(result.scopedMember, s1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists scoped role memberships`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "scoped-role-membership", "list", "--profile", profile], overrides);
      assert.deepEqual(result.scopedRoleMemberships, [
        { id: g1.id, principalId: g1.principalId, roleId: g1.roleId, directoryScopeId: g1.directoryScopeId },
      ]);
      assert.deepEqual(result.count, "1 scoped role memberships");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("directory-wide assignments live on roleAssignments")));
      assert.ok(new URL(requests[0].url).pathname.endsWith("/scopedRoleMemberships"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a scoped role membership`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "scoped-role-membership", "show", "--id", g1.id, "--profile", profile], overrides);
      assert.deepEqual(result.scopedRoleMembership, g1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied member reads name the directory role requirement`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, roleTransport(true));
      await assert.rejects(
        executeArgv(["entra", "directory-role", "member", "list", "--role", r1.id, "--profile", profile], overrides),
        error => {
          assert.equal(error.code, "GRAPH_ERROR");
          const text = [error.message, ...error.suggestions].join("\n");
          assert.match(text, /RoleManagement\.Read\.Directory/);
          assert.match(text, /Privileged Role Administrator/);
          return true;
        },
      );
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists directory-role templates end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runRolesCli(
        ["entra", "directory-role-template", "list", "--profile", profile],
        state, mode, "https://graph.microsoft.com/RoleManagement.Read.Directory");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.directoryRoleTemplates, [t1, t2]);
      assert.deepEqual(output.count, "2 directory-role templates");
      assert.equal(output.total, null);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists directory-role members end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runRolesCli(
        ["entra", "directory-role", "member", "list", "--role", r1.id, "--profile", profile],
        state, mode, "https://graph.microsoft.com/RoleManagement.Read.Directory");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.roleMembers.length, 2);
      assert.deepEqual(output.count, "2 role members");
      assert.equal(output.total, null);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied PIM reads name the role and P2/Governance requirement`, async () => {    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, roleTransport(true));
      await assert.rejects(
        executeArgv(["entra", "pim", "active", "list", "--profile", profile], overrides),
        error => {
          assert.equal(error.code, "GRAPH_ERROR");
          const text = [error.message, ...error.suggestions].join("\n");
          assert.match(text, /RoleAssignmentSchedule\.Read\.Directory/);
          assert.match(text, /Privileged Role Administrator/);
          assert.match(text, /P2 or ID Governance/);
          return true;
        },
      );
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists pim active end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runRolesCli(
        ["entra", "pim", "active", "list", "--profile", profile],
        state, mode, "https://graph.microsoft.com/RoleAssignmentSchedule.Read.Directory");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.activeAssignments.length, 2);
      assert.deepEqual(output.count, "2 active assignments");
      assert.equal(output.total, null);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists pim eligible end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runRolesCli(
        ["entra", "pim", "eligible", "list", "--profile", profile],
        state, mode, "https://graph.microsoft.com/RoleEligibilitySchedule.Read.Directory");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.eligibleAssignments, [
        { id: e1.id, principalId: e1.principalId, roleDefinitionId: e1.roleDefinitionId, memberType: "Direct" },
      ]);
      assert.deepEqual(output.count, "1 eligible assignments");
      assert.equal(output.total, null);
    } finally {
      teardownProfiles(state);
    }
  });

  for (const [command, property] of [
    [["entra", "pim", "eligible", "list"], "assignmentType"],
    [["entra", "role-assignment", "list"], "createdDateTime"],
    [["entra", "directory-role-template", "list"], "owner"],
    [["entra", "scoped-role-membership", "list"], "owner"],
  ]) {
    for (const flag of ["select", "fields"]) {
      test(`${mode} ${command.join(" ")} rejects ${property} in --${flag} before credentials`, async () => {
        const state = setupProfiles();
        try {
          const { requests, calls, overrides } = overridesFor(mode);
          await assert.rejects(
            executeArgv([...command, "--profile", profile, `--${flag}`, `id,${property}`], overrides),
            error => {
              assert.equal(error.code, "VALIDATION_ERROR");
              assert.equal(error.message, `Unknown property ${property} in --${flag}`);
              return true;
            },
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

test("delegated role reads truncate long descriptions with a --full hint", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "directory-role", "list", "--profile", "soc"], overrides);
    const helpdesk = result.directoryRoles.find(row => row.id === r2.id);
    assert.match(helpdesk.description, /truncated, \d+ chars total/);
    assert.ok(result.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "directory-role", "list", "--profile", "soc", "--full"], overrides);
    assert.equal(full.directoryRoles.find(row => row.id === r2.id).description, r2.description);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated role filters pass through as plain $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "role-assignment", "list", "--profile", "soc",
      "--filter", "principalId eq 'dddddddd-dddd-4ddd-8ddd-dddddddddddd'"], overrides);
    assert.equal(result.roleAssignments.length, 1);
    const url = new URL(requests[0].url);
    assert.equal(url.searchParams.get("$filter"), "principalId eq 'dddddddd-dddd-4ddd-8ddd-dddddddddddd'");
    assert.equal(url.searchParams.has("$count"), false);
    assert.equal(requests[0].headers.ConsistencyLevel, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated role reads resume a capped list through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "directory-role", "list", "--profile", "soc", "--limit", "2"], overrides);
    assert.equal(first.directoryRoles.length, 2);
    assert.deepEqual(first.count, "2 directory roles shown, more available");
    assert.equal(first.total, null);
    assert.equal(first.complete, false);
    assert.ok(typeof first.cursor === "string" && first.cursor.length > 0);
    const second = await executeArgv(["entra", "directory-role", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.directoryRoles, [{ id: r3.id, displayName: "Directory Readers" }]);
    assert.deepEqual(second.count, "1 directory roles");
    assert.equal(second.total, null);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated role reads reject unknown properties before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "pim", "eligible", "list", "--profile", "soc", "--select", "id,owner"], overrides),
      /Unknown property owner in --select/,
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated role member lists reject a missing parent role before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "directory-role", "member", "list", "--profile", "soc"], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        assert.match(error.message, /--role is required/);
        return true;
      },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});
