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
const groupId = "33333333-3333-4333-8333-333333333333";
const principalId = "44444444-4444-4344-8344-444444444444";
const otherPrincipalId = "55555555-5555-4355-8355-555555555555";

const scheduleInfo = {
  startDateTime: "2024-01-01T00:00:00Z",
  recurrence: null,
  expiration: { type: "noExpiration", endDateTime: null, duration: null },
};

const as1 = {
  id: "a1a1a1a1-a1a1-4a1a-8a1a-a1a1a1a1a1a1",
  accessId: "member",
  assignmentType: "Assigned",
  memberType: "Direct",
  principalId,
  groupId,
  status: "Provisioned",
  scheduleInfo,
  createdDateTime: "2024-01-01T00:00:00Z",
  modifiedDateTime: null,
  createdUsing: "b1b1b1b1-b1b1-4b1b-8b1b-b1b1b1b1b1b1",
};
const as2 = {
  id: "a2a2a2a2-a2a2-4a2a-8a2a-a2a2a2a2a2a2",
  accessId: "owner",
  assignmentType: "Activated",
  memberType: "Direct",
  principalId: otherPrincipalId,
  groupId,
  status: "Provisioned",
  scheduleInfo,
  createdDateTime: "2024-01-02T00:00:00Z",
  modifiedDateTime: "2024-02-01T00:00:00Z",
  createdUsing: null,
};
const assignmentSchedules = [as1, as2];

const ai1 = {
  id: "c1c1c1c1-c1c1-4c1c-8c1c-c1c1c1c1c1c1",
  accessId: "member",
  assignmentType: "Activated",
  memberType: "Direct",
  principalId,
  groupId,
  assignmentScheduleId: as1.id,
  startDateTime: "2024-01-01T00:00:00Z",
  endDateTime: "2024-07-01T00:00:00Z",
};
const assignmentInstances = [ai1];

const es1 = {
  id: "d1d1d1d1-d1d1-4d1d-8d1d-d1d1d1d1d1d1",
  accessId: "member",
  memberType: "Direct",
  principalId,
  groupId,
  status: "Provisioned",
  scheduleInfo,
  createdDateTime: "2024-01-01T00:00:00Z",
  modifiedDateTime: null,
  createdUsing: "e1e1e1e1-e1e1-4e1e-8e1e-e1e1e1e1e1e1",
};
const eligibilitySchedules = [es1];

const ei1 = {
  id: "f1f1f1f1-f1f1-4f1f-8f1f-f1f1f1f1f1f1",
  accessId: "member",
  memberType: "Direct",
  principalId,
  groupId,
  eligibilityScheduleId: es1.id,
  startDateTime: "2024-01-01T00:00:00Z",
  endDateTime: null,
};
const eligibilityInstances = [ei1];

const longJustification = `Quarterly access review grant. ${"x".repeat(600)}`;
const er1 = {
  id: "11111111-2222-4333-8444-555555555555",
  accessId: "member",
  action: "adminAssign",
  status: "Provisioned",
  principalId,
  groupId,
  justification: longJustification,
  scheduleInfo,
  ticketInfo: { ticketNumber: null, ticketSystem: null },
  createdDateTime: "2024-01-01T00:00:00Z",
  completedDateTime: "2024-01-01T00:01:00Z",
  approvalId: null,
  targetScheduleId: es1.id,
  isValidationOnly: false,
};
const eligibilityRequests = [er1];

const base = "/v1.0/identityGovernance/privilegedAccess/group";
const routes = {
  assignmentSchedules: `${base}/assignmentSchedules`,
  assignmentSchedule: id => `${base}/assignmentSchedules/${id}`,
  assignmentInstances: `${base}/assignmentScheduleInstances`,
  assignmentInstance: id => `${base}/assignmentScheduleInstances/${id}`,
  eligibilitySchedules: `${base}/eligibilitySchedules`,
  eligibilitySchedule: id => `${base}/eligibilitySchedules/${id}`,
  eligibilityInstances: `${base}/eligibilityScheduleInstances`,
  eligibilityInstance: id => `${base}/eligibilityScheduleInstances/${id}`,
  eligibilityRequests: `${base}/eligibilityScheduleRequests`,
  eligibilityRequest: id => `${base}/eligibilityScheduleRequests/${id}`,
};

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-group-pim-"));
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

function scoped(url) {
  const filter = url.searchParams.get("$filter") ?? "";
  return filter.includes("groupId") || filter.includes("principalId");
}

function groupPimTransport(denied = false) {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (denied) return json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } });
    // Every one of these lists requires a groupId or principalId $filter;
    // singles carry no query filter. All filters arrive as plain $filter.
    if (path === routes.assignmentSchedules) {
      // Continuations carry server paging state, not the original filter.
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [as2] });
      assert.ok(scoped(url), "assignment schedules list must carry a groupId or principalId $filter");
      return json(200, {
        value: [as1],
        "@odata.nextLink": `https://graph.microsoft.com${routes.assignmentSchedules}?%24skiptoken=page2`,
      });
    }
    if (path === routes.assignmentSchedule(as1.id)) return json(200, as1);
    if (path === routes.assignmentSchedule(as2.id)) return json(200, as2);
    if (path === routes.assignmentInstances) {
      assert.ok(scoped(url), "assignment instances list must carry a groupId or principalId $filter");
      return json(200, { value: assignmentInstances });
    }
    if (path === routes.assignmentInstance(ai1.id)) return json(200, ai1);
    if (path === routes.eligibilitySchedules) {
      assert.ok(scoped(url), "eligibility schedules list must carry a groupId or principalId $filter");
      return json(200, { value: eligibilitySchedules });
    }
    if (path === routes.eligibilitySchedule(es1.id)) return json(200, es1);
    if (path === routes.eligibilityInstances) {
      assert.ok(scoped(url), "eligibility instances list must carry a groupId or principalId $filter");
      return json(200, { value: eligibilityInstances });
    }
    if (path === routes.eligibilityInstance(ei1.id)) return json(200, ei1);
    if (path === routes.eligibilityRequests) {
      assert.ok(scoped(url), "eligibility requests list must carry a groupId or principalId $filter");
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("createdBy"));
      return json(200, { value: eligibilityRequests });
    }
    if (path === routes.eligibilityRequest(er1.id)) return json(200, er1);
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? groupPimTransport();
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

function runGroupPimCli(args, state, mode, scopes) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-group-pim-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({
        mode, scopes, assignmentSchedules, assignmentInstances,
        eligibilitySchedules, eligibilityInstances, eligibilityRequests, groupId, principalId,
      }),
    },
  });
}

const filter = `groupId eq '${groupId}'`;

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists group assignment schedules with compact rows`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-assignment-schedule", "list", "--profile", profile, "--filter", filter, "--limit", "2"],
        overrides);
      assert.deepEqual(result.assignmentSchedules, [
        { id: as1.id, principalId, groupId, accessId: "member", assignmentType: "Assigned" },
        { id: as2.id, principalId: otherPrincipalId, groupId, accessId: "owner", assignmentType: "Activated" },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("never the schedule itself") || hint.includes("govern group membership")));
      assert.ok(new URL(requests[0].url).pathname.endsWith("/privilegedAccess/group/assignmentSchedules"));
      assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,principalId,groupId,accessId,assignmentType");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a group assignment schedule`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-assignment-schedule", "show", "--id", as1.id, "--profile", profile], overrides);
      assert.deepEqual(result.assignmentSchedule, as1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists group assignment instances with Assigned and Activated distinct`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-assignment-instance", "list", "--profile", profile, "--filter", filter], overrides);
      assert.deepEqual(result.assignmentScheduleInstances, [
        { id: ai1.id, principalId, groupId, accessId: "member", assignmentType: "Activated" },
      ]);
      assert.deepEqual(result.count, { returned: 1, complete: true });
      assert.ok(result.help.some(hint => hint.includes("provisioned membership or ownership windows")));
      assert.ok(new URL(requests[0].url).pathname.endsWith("/privilegedAccess/group/assignmentScheduleInstances"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a group assignment instance`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-assignment-instance", "show", "--id", ai1.id, "--profile", profile], overrides);
      assert.deepEqual(result.assignmentScheduleInstance, ai1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists group eligibility schedules as not active`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-eligibility-schedule", "list", "--profile", profile, "--filter", filter], overrides);
      assert.deepEqual(result.eligibilitySchedules, [
        { id: es1.id, principalId, groupId, accessId: "member", memberType: "Direct" },
      ]);
      assert.deepEqual(result.count, { returned: 1, complete: true });
      assert.ok(result.help.some(hint => hint.includes("not active grants") || hint.includes("may activate")));
      assert.ok(new URL(requests[0].url).pathname.endsWith("/privilegedAccess/group/eligibilitySchedules"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a group eligibility schedule`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-eligibility-schedule", "show", "--id", es1.id, "--profile", profile], overrides);
      assert.deepEqual(result.eligibilitySchedule, es1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists group eligibility instances`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-eligibility-instance", "list", "--profile", profile, "--filter", filter], overrides);
      assert.deepEqual(result.eligibilityScheduleInstances, [
        { id: ei1.id, principalId, groupId, accessId: "member", memberType: "Direct" },
      ]);
      assert.deepEqual(result.count, { returned: 1, complete: true });
      assert.ok(new URL(requests[0].url).pathname.endsWith("/privilegedAccess/group/eligibilityScheduleInstances"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a group eligibility instance`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-eligibility-instance", "show", "--id", ei1.id, "--profile", profile], overrides);
      assert.deepEqual(result.eligibilityScheduleInstance, ei1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists group eligibility requests without justification by default`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-eligibility-request", "list", "--profile", profile, "--filter", filter], overrides);
      assert.deepEqual(result.eligibilityScheduleRequests, [
        { id: er1.id, action: "adminAssign", status: "Provisioned", principalId, groupId, accessId: "member" },
      ]);
      assert.deepEqual(result.count, { returned: 1, complete: true });
      assert.ok(result.help.some(hint => hint.includes("explicit --select")));
      assert.ok(!new URL(requests[0].url).searchParams.get("$select")?.split(",").includes("justification"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a group eligibility request without justification by default`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(
        ["entra", "pim", "group-eligibility-request", "show", "--id", er1.id, "--profile", profile], overrides);
      assert.equal(result.eligibilityScheduleRequest.justification, undefined);
      assert.equal(result.eligibilityScheduleRequest.action, "adminAssign");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied assignment reads name the scope, role and P2/Governance requirement`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, groupPimTransport(true));
      await assert.rejects(
        executeArgv(["entra", "pim", "group-assignment-schedule", "list", "--profile", profile, "--filter", filter], overrides),
        error => {
          assert.equal(error.code, "GRAPH_ERROR");
          const text = [error.message, ...error.suggestions].join("\n");
          assert.match(text, /PrivilegedAssignmentSchedule\.Read\.AzureADGroup/);
          assert.match(text, /Privileged Role Administrator/);
          assert.match(text, /P2 or ID Governance/);
          return true;
        },
      );
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied eligibility reads name the scope, role and P2/Governance requirement`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, groupPimTransport(true));
      await assert.rejects(
        executeArgv(["entra", "pim", "group-eligibility-request", "list", "--profile", profile, "--filter", filter], overrides),
        error => {
          assert.equal(error.code, "GRAPH_ERROR");
          const text = [error.message, ...error.suggestions].join("\n");
          assert.match(text, /PrivilegedEligibilitySchedule\.Read\.AzureADGroup/);
          assert.match(text, /Privileged Role Administrator/);
          assert.match(text, /P2 or ID Governance/);
          return true;
        },
      );
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists group assignment schedules end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runGroupPimCli(
        ["entra", "pim", "group-assignment-schedule", "list", "--profile", profile, "--filter", filter],
        state, mode, "https://graph.microsoft.com/PrivilegedAssignmentSchedule.Read.AzureADGroup");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.ok(output.assignmentSchedules.length >= 1);
      assert.deepEqual(output.count, { returned: output.assignmentSchedules.length, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists group eligibility requests end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runGroupPimCli(
        ["entra", "pim", "group-eligibility-request", "list", "--profile", profile, "--filter", filter],
        state, mode, "https://graph.microsoft.com/PrivilegedEligibilitySchedule.Read.AzureADGroup");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.eligibilityScheduleRequests.length, 1);
      assert.deepEqual(output.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });
}

test("delegated group PIM lists reject a missing --filter before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "pim", "group-assignment-schedule", "list", "--profile", "soc"], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        assert.match(error.message, /--filter is required/);
        return true;
      },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated group PIM lists reject an unrelated --filter before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "pim", "group-eligibility-schedule", "list", "--profile", "soc",
        "--filter", "status eq 'Provisioned'"], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        assert.match(error.message, /groupId or principalId/);
        return true;
      },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated group PIM filters pass through as plain $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "pim", "group-assignment-instance", "list",
      "--profile", "soc", "--filter", filter], overrides);
    assert.equal(result.assignmentScheduleInstances.length, 1);
    const url = new URL(requests[0].url);
    assert.equal(url.searchParams.get("$filter"), filter);
    assert.equal(url.searchParams.has("$count"), false);
    assert.equal(requests[0].headers.ConsistencyLevel, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated group PIM reads resume a capped list through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "pim", "group-assignment-schedule", "list",
      "--profile", "soc", "--filter", filter, "--limit", "1"], overrides);
    assert.equal(first.assignmentSchedules.length, 1);
    assert.deepEqual(first.count, { returned: 1, complete: false, reason: first.count.reason });
    assert.ok(typeof first.cursor === "string" && first.cursor.length > 0);
    assert.ok(first.help.some(hint => hint.includes("--filter")));
    const second = await executeArgv(["entra", "pim", "group-assignment-schedule", "list",
      "--profile", "soc", "--filter", filter, "--cursor", first.cursor], overrides);
    assert.deepEqual(second.assignmentSchedules, [
      { id: as2.id, principalId: otherPrincipalId, groupId, accessId: "owner", assignmentType: "Activated" },
    ]);
    assert.deepEqual(second.count, { returned: 1, complete: true });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated group PIM reads reject unknown properties before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "pim", "group-assignment-schedule", "list", "--profile", "soc",
        "--filter", filter, "--select", "id,group"], overrides),
      /Unknown property group in --select/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated group PIM reads reject navigation names before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "pim", "group-eligibility-request", "show", "--id", er1.id,
        "--profile", "soc", "--fields", "id,targetSchedule"], overrides),
      /Unknown property targetSchedule in --fields/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated eligibility requests truncate long justification with a --full hint", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "pim", "group-eligibility-request", "list",
      "--profile", "soc", "--filter", filter, "--select", "id,action,status,justification"], overrides);
    const row = result.eligibilityScheduleRequests[0];
    assert.match(row.justification, /truncated, \d+ chars total/);
    assert.ok(result.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "pim", "group-eligibility-request", "list",
      "--profile", "soc", "--filter", filter, "--select", "id,action,status,justification", "--full"], overrides);
    assert.equal(full.eligibilityScheduleRequests[0].justification, longJustification);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated group PIM show rejects a malformed body as unknown", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated", transport(() => json(200, [ai1])));
    await assert.rejects(
      executeArgv(["entra", "pim", "group-assignment-instance", "show", "--id", ai1.id, "--profile", "soc"], overrides),
      error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.match(error.message, /malformed assignmentScheduleInstance body/);
        return true;
      },
    );
  } finally {
    teardownProfiles(state);
  }
});
