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
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, scopes, roles, assignments, eligible, active }),
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
      assert.deepEqual(result.count, { returned: 3, complete: true });
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

  test(`${mode} lists direct role assignments with the direct-only distinction`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "role-assignment", "list", "--profile", profile], overrides);
      assert.deepEqual(result.roleAssignments, [
        { id: a1.id, principalId: a1.principalId, roleDefinitionId: a1.roleDefinitionId, directoryScopeId: "/" },
      ]);
      assert.deepEqual(result.count, { returned: 1, complete: true });
      assert.ok(result.help.some(hint => hint.includes("Direct persistent assignments only")));
      assert.ok(result.help.some(hint => hint.includes("pim active")));
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
      assert.deepEqual(result.count, { returned: 1, complete: true });
      assert.ok(result.help.some(hint => hint.includes("not active")));
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
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("assignmentType Assigned")));
      assert.ok(result.help.some(hint => hint.includes("role-assignment list")));
      assert.ok(new URL(requests[0].url).pathname.endsWith("/roleManagement/directory/roleAssignmentScheduleInstances"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied PIM reads name the role and P2/Governance requirement`, async () => {
    const state = setupProfiles();
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
      assert.deepEqual(output.count, { returned: 2, complete: true });
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
      assert.deepEqual(output.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  for (const [command, property] of [
    [["entra", "pim", "eligible", "list"], "assignmentType"],
    [["entra", "role-assignment", "list"], "createdDateTime"],
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
    assert.deepEqual(first.count, { returned: 2, complete: false, reason: first.count.reason });
    assert.ok(typeof first.cursor === "string" && first.cursor.length > 0);
    const second = await executeArgv(["entra", "directory-role", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.directoryRoles, [{ id: r3.id, displayName: "Directory Readers" }]);
    assert.deepEqual(second.count, { returned: 1, complete: true });
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
