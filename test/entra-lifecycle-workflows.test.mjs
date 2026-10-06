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

const tenant = "33333333-3333-4333-8333-333333333333";
const client = "44444444-4444-4222-8222-444444444444";
const workflowScopes = ["https://graph.microsoft.com/LifecycleWorkflows-Workflow.ReadBasic.All"];
const lifecycleScopes = ["https://graph.microsoft.com/LifecycleWorkflows.Read.All"];
const reportsScopes = ["https://graph.microsoft.com/LifecycleWorkflows-Reports.Read.All"];

const longDescription = `Configure pre-hire tasks for onboarding${" with considerable detail".repeat(30)}`;
const wf1 = {
  id: "11111111-1111-4111-8111-111111111111",
  category: "joiner",
  createdDateTime: "2023-01-10T20:04:30.619368Z",
  deletedDateTime: null,
  description: longDescription,
  displayName: "Global pre hire",
  executionConditions: { scope: { rule: "(department eq 'Marketing')" } },
  isEnabled: true,
  isSchedulingEnabled: true,
  lastModifiedDateTime: "2023-08-16T20:05:51.4618603Z",
  nextScheduleRunDateTime: null,
  version: 12,
  quarantineDetails: null,
  settings: null,
  tasks: [{ id: "task-expanded-by-default", displayName: "Enable User Account" }],
  createdBy: { id: "creator-id" },
};
const wf2 = {
  id: "22222222-2222-4222-8222-222222222222",
  category: "leaver",
  displayName: "Post-Offboarding of an employee",
  isEnabled: true,
  isSchedulingEnabled: false,
};
const workflows = [wf1, wf2];
const tmpl1 = {
  id: "77179007-8114-41b5-922e-2e22109df41f",
  category: "joiner",
  description: "Configure pre-hire tasks for onboarding employees before their first day",
  displayName: "Onboard pre-hire employee",
  executionConditions: { scope: { rule: "department eq 'Marketing'" } },
  tasks: [{ id: "template-task-expanded-by-default", displayName: "Generate TAP And Send Email" }],
};
const tmpl2 = { id: "4e06785d-7c1d-4b24-b9e1-bba2b890c58b", category: "leaver", displayName: "Offboard an employee" };
const workflowTemplates = [tmpl1, tmpl2];
const td1 = {
  id: "22085229-5809-45e8-97fd-270d28d66910",
  category: "joiner,leaver,mover",
  continueOnError: false,
  description: "Add user to selected groups",
  displayName: "Add user to groups",
  parameters: [{ name: "groupID", values: [], valueType: "string" }],
  version: 1,
};
const td2 = { id: "1dfdfcc7-52fa-4c2e-bf3a-e3919cc12950", displayName: "Disable user account" };
const taskDefinitions = [td1, td2];
const settings = {
  workflowScheduleIntervalInHours: 3,
  emailSettings: { senderDomain: "ContosoIndustries.net", useCompanyBranding: true },
  quarantineConfiguration: null,
};
const run1 = {
  id: "dad77a47-6eda-4de7-bc37-fe8eb5aaf17d",
  completedDateTime: "2022-08-24T23:28:11.1348863Z",
  failedTasksCount: 0,
  failedUsersCount: 0,
  lastUpdatedDateTime: "2022-08-24T23:28:12.1348863Z",
  processingStatus: "completed",
  scheduledDateTime: "2022-08-24T23:28:01.6476553Z",
  startedDateTime: "2022-08-24T23:28:04.490313Z",
  successfulUsersCount: 2,
  totalTasksCounts: 4,
  totalUsersCount: 2,
  totalUnprocessedTasksCount: 0,
  workflowExecutionType: "scheduled",
  userProcessingResults: [{ id: "run-nav-expanded" }],
  taskProcessingResults: [{ id: "run-task-nav-expanded" }],
};
const run2 = {
  id: "e65e08a0-d68d-41dc-915b-8c4019af5cc2",
  processingStatus: "inProgress",
  totalUsersCount: 1,
  failedUsersCount: 0,
  successfulUsersCount: 0,
};
const runs = [run1, run2];
const upr1 = {
  id: "78b83505-6967-4168-a7ea-4921c0543ce9",
  completedDateTime: "2022-08-24T23:28:11.1348863Z",
  failedTasksCount: 0,
  processingStatus: "completed",
  scheduledDateTime: "2022-08-24T23:28:01.6476553Z",
  startedDateTime: "2022-08-24T23:28:04.490313Z",
  totalTasksCount: 2,
  totalUnprocessedTasksCount: 0,
  workflowExecutionType: "onDemand",
  workflowVersion: 1,
  subject: { id: "ea09ac2e-77e3-4134-85f2-25ccf3c33387" },
  taskProcessingResults: [{ id: "user-task-nav-expanded" }],
};
const upr2 = {
  id: "40efc576-840f-47d0-ab95-5abca800f8a2",
  processingStatus: "completed",
  failedTasksCount: 0,
  totalTasksCount: 3,
  totalUnprocessedTasksCount: 0,
};
const userProcessingResults = [upr1, upr2];
const spr1 = {
  id: "6c6d9550-4adc-117f-b307-ee188d511d67",
  completedDateTime: "2026-05-15T10:35:00Z",
  failedTasksCount: 0,
  processingStatus: "completed",
  scheduledDateTime: "2026-05-15T10:30:00Z",
  startedDateTime: "2026-05-15T10:30:15Z",
  totalTasksCount: 3,
  totalUnprocessedTasksCount: 0,
  workflowExecutionType: "extensibilityOnDemand",
  workflowVersion: 1,
  subjectType: "provisioningObject",
  subject: { "@odata.type": "#microsoft.graph.identityGovernance.provisioningObjectWorkflowSubject", id: "b74f0fae-b1f3-4c96-9bf0-d4d8a8e37cbe" },
  taskProcessingResults: [{ id: "subject-task-nav-expanded" }],
};
const spr2 = {
  id: "7d7d9661-5bed-228f-c418-ff299622e78",
  subjectType: "user",
  processingStatus: "queued",
  failedTasksCount: 0,
  totalTasksCount: 1,
  totalUnprocessedTasksCount: 1,
};
const subjectProcessingResults = [spr1, spr2];
const tr1 = {
  id: "3a3bea11-99ca-462d-86fb-d283db8d734a",
  runId: "dad77a47-6eda-4de7-bc37-fe8eb5aaf17d",
  processingStatus: "completed",
  successfulUsersCount: 2,
  failedUsersCount: 0,
  unprocessedUsersCount: 0,
  totalUsersCount: 2,
  startedDateTime: "2022-08-24T23:28:04.5785337Z",
  completedDateTime: "2022-08-24T23:28:11.1348863Z",
  lastUpdatedDateTime: "2022-08-24T23:33:09.1980357Z",
  task: { id: "report-task-expanded" },
  taskDefinition: { id: "report-definition-expanded" },
  taskProcessingResults: [{ id: "report-task-nav-expanded" }],
};
const tr2 = {
  id: "23f37fcb-040d-4ee9-91df-1234700ebeb6",
  runId: "dad77a47-6eda-4de7-bc37-fe8eb5aaf17d",
  processingStatus: "inProgress",
  totalUsersCount: 1,
  failedUsersCount: 0,
  successfulUsersCount: 0,
};
const taskReports = [tr1, tr2];
const longFailureReason = `Access denied for the target group${" with considerable detail".repeat(30)}`;
const rupr1 = {
  id: "a1b2c3d4-1111-4111-8111-111111111111",
  completedDateTime: "2022-08-24T23:28:11.1348863Z",
  failedTasksCount: 1,
  processingStatus: "failed",
  scheduledDateTime: "2022-08-24T23:28:01.6476553Z",
  startedDateTime: "2022-08-24T23:28:04.490313Z",
  totalTasksCount: 2,
  totalUnprocessedTasksCount: 0,
  workflowExecutionType: "scheduled",
  workflowVersion: 12,
  subject: { id: "ea09ac2e-77e3-4134-85f2-25ccf3c33387" },
  taskProcessingResults: [{ id: "run-user-task-nav-expanded" }],
};
const rupr2 = {
  id: "b2c3d4e5-2222-4222-8222-222222222222",
  processingStatus: "completed",
  failedTasksCount: 0,
  totalTasksCount: 1,
  totalUnprocessedTasksCount: 0,
};
const runUserProcessingResults = [rupr1, rupr2];
const rspr1 = {
  id: "c3d4e5f6-3333-4333-8333-333333333333",
  completedDateTime: "2026-05-15T10:35:00Z",
  failedTasksCount: 0,
  processingStatus: "completed",
  scheduledDateTime: "2026-05-15T10:30:00Z",
  startedDateTime: "2026-05-15T10:30:15Z",
  totalTasksCount: 2,
  totalUnprocessedTasksCount: 0,
  workflowExecutionType: "scheduled",
  workflowVersion: 12,
  subjectType: "user",
  subject: { id: "d4e5f6a7-4444-4222-8222-444444444444" },
  taskProcessingResults: [{ id: "run-subject-task-nav-expanded" }],
};
const rspr2 = {
  id: "e5f6a7b8-5555-4222-8222-555555555555",
  subjectType: "provisioningObject",
  processingStatus: "queued",
  failedTasksCount: 0,
  totalTasksCount: 1,
  totalUnprocessedTasksCount: 1,
};
const runSubjectProcessingResults = [rspr1, rspr2];
const rtpr1 = {
  id: "f6a7b8c9-6666-4222-8222-666666666666",
  completedDateTime: "2023-01-20T17:16:03.4863553Z",
  createdDateTime: "2023-01-20T17:16:00.9095011Z",
  failureReason: longFailureReason,
  processingInfo: "User was already a member of all requested groups",
  processingStatus: "failed",
  startedDateTime: "2023-01-20T17:16:02.8025169Z",
  workflowSubject: { id: "a7b8c9d0-7777-4222-8222-777777777777" },
  subject: { id: "1baa57fa-3c4e-4526-ba5a-db47a9df95f0" },
  task: { id: "c8dbaed8-3d23-4e5a-8f65-130767639667" },
};
const rtpr2 = {
  id: "a7b8c9d0-8888-4222-8222-888888888888",
  processingStatus: "completed",
  failureReason: null,
};
const runTaskProcessingResults = [rtpr1, rtpr2];
const reportTask = {
  id: "fafa2189-cd62-4643-a825-06cab8817086",
  arguments: [{ name: "groupID", value: "e5659cb0-bcbb-4a9f-9092-90f72bd19028" }],
  category: "joiner,leaver",
  continueOnError: false,
  description: "Enable user account in the directory",
  displayName: "Enable User Account",
  executionSequence: 1,
  isEnabled: true,
  taskDefinitionId: "6fc52c9d-398b-4305-9763-15f42c1676fc",
};

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-lifecycle-"));
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

const lcBase = "/v1.0/identityGovernance/lifecycleWorkflows";

function lifecycleTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === `${lcBase}/workflows/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(workflows.length) };
    }
    if (path === `${lcBase}/workflowTemplates/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(workflowTemplates.length) };
    }
    if (path === `${lcBase}/taskDefinitions/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(taskDefinitions.length) };
    }
    if (path === `${lcBase}/workflows`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [wf2] });
      return json(200, {
        value: [wf1],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows?%24skiptoken=page2",
      });
    }
    if (path === `${lcBase}/workflowTemplates`) return json(200, { value: workflowTemplates });
    if (path === `${lcBase}/taskDefinitions`) return json(200, { value: taskDefinitions });
    if (path === `${lcBase}/settings`) return json(200, settings);
    const nestedBase = `${lcBase}/workflows/${wf1.id}`;
    if (path === `${nestedBase}/runs/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(runs.length) };
    }
    if (path === `${nestedBase}/userProcessingResults/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(userProcessingResults.length) };
    }
    if (path === `${nestedBase}/subjectProcessingResults/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(subjectProcessingResults.length) };
    }
    if (path === `${nestedBase}/taskReports/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(taskReports.length) };
    }
    const runNestedBase = `${nestedBase}/runs/${run1.id}`;
    if (path === `${runNestedBase}/userProcessingResults/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(runUserProcessingResults.length) };
    }
    if (path === `${runNestedBase}/subjectProcessingResults/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(runSubjectProcessingResults.length) };
    }
    if (path === `${runNestedBase}/taskProcessingResults/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(runTaskProcessingResults.length) };
    }
    if (path === `${nestedBase}/runs`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [run2] });
      return json(200, {
        value: [run1],
        "@odata.nextLink": `https://graph.microsoft.com${nestedBase}/runs?%24skiptoken=page2`,
      });
    }
    if (path === `${nestedBase}/userProcessingResults`) return json(200, { value: userProcessingResults });
    if (path === `${nestedBase}/subjectProcessingResults`) return json(200, { value: subjectProcessingResults });
    if (path === `${nestedBase}/taskReports`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [tr2] });
      return json(200, {
        value: [tr1],
        "@odata.nextLink": `https://graph.microsoft.com${nestedBase}/taskReports?%24skiptoken=page2`,
      });
    }
    if (path === `${runNestedBase}/userProcessingResults`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [rupr2] });
      return json(200, {
        value: [rupr1],
        "@odata.nextLink": `https://graph.microsoft.com${runNestedBase}/userProcessingResults?%24skiptoken=page2`,
      });
    }
    if (path === `${runNestedBase}/subjectProcessingResults`) return json(200, { value: runSubjectProcessingResults });
    if (path === `${runNestedBase}/taskProcessingResults`) return json(200, { value: runTaskProcessingResults });
    for (const report of taskReports) {
      if (path === `${nestedBase}/taskReports/${report.id}`) return json(200, report);
      if (path === `${nestedBase}/taskReports/${report.id}/task`) return json(200, reportTask);
      if (path === `${nestedBase}/taskReports/${report.id}/taskDefinition`) return json(200, td1);
    }
    for (const [rows, base] of [[runUserProcessingResults, `${runNestedBase}/userProcessingResults`], [runSubjectProcessingResults, `${runNestedBase}/subjectProcessingResults`], [runTaskProcessingResults, `${runNestedBase}/taskProcessingResults`]]) {
      const nested = rows.find(row => path === `${base}/${row.id}`);
      if (nested) return json(200, nested);
    }
    for (const [rows, base] of [[runs, `${nestedBase}/runs`], [userProcessingResults, `${nestedBase}/userProcessingResults`], [subjectProcessingResults, `${nestedBase}/subjectProcessingResults`]]) {
      const found = rows.find(row => path === `${base}/${row.id}`);
      if (found) return json(200, found);
    }
    for (const [rows, base] of [[workflows, `${lcBase}/workflows`], [workflowTemplates, `${lcBase}/workflowTemplates`], [taskDefinitions, `${lcBase}/taskDefinitions`]]) {
      const found = rows.find(row => path === `${base}/${row.id}`);
      if (found) return json(200, found);
    }
    return json(404, { error: { code: "Request_ResourceNotFound", message: "no such lifecycle object" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? lifecycleTransport();
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

function runLifecycleCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-lifecycle-workflows-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, workflows, workflowTemplates, taskDefinitions, settings, runs, userProcessingResults, subjectProcessingResults, taskReports, reportTask, runUserProcessingResults, runSubjectProcessingResults, runTaskProcessingResults, denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists workflows with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "workflow", "list", "--profile", profile,
        "--select", "id,displayName,category,isEnabled,isSchedulingEnabled,description,deletedDateTime"], overrides);
      assert.deepEqual(result.workflows, [
        { id: wf1.id, displayName: "Global pre hire", category: "joiner", isEnabled: true, isSchedulingEnabled: true, description: `${longDescription.slice(0, 500)}... (truncated, ${longDescription.length} chars total)`, deletedDateTime: null },
        { id: wf2.id, displayName: "Post-Offboarding of an employee", category: "leaver", isEnabled: true, isSchedulingEnabled: false },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra lifecycle workflow show --id <workflow-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows?"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(workflowScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists workflows with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "workflow", "list", "--profile", profile,
        "--filter", "category eq 'leaver'"], overrides);
      assert.equal(result.count.returned, 2);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "category eq 'leaver'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped workflow list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "lifecycle", "workflow", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.workflows.map(row => row.id), [wf1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "lifecycle", "workflow", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.workflows.map(row => row.id), [wf2.id]);
      assert.deepEqual(second.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one workflow with the full scalar set and no relationships`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "workflow", "show", "--id", wf1.id, "--profile", profile], overrides);
      assert.ok(!Object.hasOwn(result.workflow, "tasks"));
      assert.ok(!Object.hasOwn(result.workflow, "createdBy"));
      assert.deepEqual(result.workflow, {
        id: wf1.id,
        category: "joiner",
        createdDateTime: "2023-01-10T20:04:30.619368Z",
        deletedDateTime: null,
        description: `${longDescription.slice(0, 500)}... (truncated, ${longDescription.length} chars total)`,
        displayName: "Global pre hire",
        executionConditions: { scope: { rule: "(department eq 'Marketing')" } },
        isEnabled: true,
        isSchedulingEnabled: true,
        lastModifiedDateTime: "2023-08-16T20:05:51.4618603Z",
        nextScheduleRunDateTime: null,
        version: 12,
        quarantineDetails: null,
        settings: null,
      });
      assert.ok(result.help.some(hint => hint.includes("--full")));
      assert.ok(!new URL(requests[0].url).searchParams.has("$expand"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown workflow ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "lifecycle", "workflow", "show", "--id", "workflow-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists runs with compact rows and the reports scope`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "list", "--workflow", wf1.id, "--profile", profile], overrides);
      assert.deepEqual(result.runs, [
        { id: run1.id, processingStatus: "completed", totalUsersCount: 2, failedUsersCount: 0, successfulUsersCount: 2 },
        { id: run2.id, processingStatus: "inProgress", totalUsersCount: 1, failedUsersCount: 0, successfulUsersCount: 0 },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra lifecycle run show --workflow <workflow-id> --id <run-id>")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/runs?`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(reportsScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists runs with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "list", "--workflow", wf1.id, "--profile", profile,
        "--filter", "processingStatus eq 'completed'"], overrides);
      assert.equal(result.count.returned, 2);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "processingStatus eq 'completed'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped run list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "lifecycle", "run", "list", "--workflow", wf1.id, "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.runs.map(row => row.id), [run1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "lifecycle", "run", "list", "--workflow", wf1.id, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.runs.map(row => row.id), [run2.id]);
      assert.deepEqual(second.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one run with the full scalar set and no processing-result relationships`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "show", "--workflow", wf1.id, "--id", run1.id, "--profile", profile], overrides);
      assert.ok(!Object.hasOwn(result.run, "userProcessingResults"));
      assert.ok(!Object.hasOwn(result.run, "taskProcessingResults"));
      assert.deepEqual(result.run, {
        id: run1.id,
        completedDateTime: "2022-08-24T23:28:11.1348863Z",
        failedTasksCount: 0,
        failedUsersCount: 0,
        lastUpdatedDateTime: "2022-08-24T23:28:12.1348863Z",
        processingStatus: "completed",
        scheduledDateTime: "2022-08-24T23:28:01.6476553Z",
        startedDateTime: "2022-08-24T23:28:04.490313Z",
        successfulUsersCount: 2,
        totalTasksCounts: 4,
        totalUsersCount: 2,
        totalUnprocessedTasksCount: 0,
        workflowExecutionType: "scheduled",
      });
      assert.ok(!new URL(requests[0].url).searchParams.has("$expand"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown run ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "lifecycle", "run", "show", "--workflow", wf1.id, "--id", "run-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists user processing results with minimal personal-data defaults`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "user-processing-result", "list", "--workflow", wf1.id, "--profile", profile], overrides);
      assert.deepEqual(result.userProcessingResults, [
        { id: upr1.id, processingStatus: "completed", failedTasksCount: 0, totalTasksCount: 2, totalUnprocessedTasksCount: 0 },
        { id: upr2.id, processingStatus: "completed", failedTasksCount: 0, totalTasksCount: 3, totalUnprocessedTasksCount: 0 },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra lifecycle user-processing-result show")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/userProcessingResults?`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(reportsScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one user processing result without the subject link`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "user-processing-result", "show", "--workflow", wf1.id, "--id", upr1.id, "--profile", profile], overrides);
      assert.ok(!Object.hasOwn(result.userProcessingResult, "subject"));
      assert.ok(!Object.hasOwn(result.userProcessingResult, "taskProcessingResults"));
      assert.deepEqual(result.userProcessingResult, {
        id: upr1.id,
        processingStatus: "completed",
        failedTasksCount: 0,
        totalTasksCount: 2,
        totalUnprocessedTasksCount: 0,
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists subject processing results whole with local field projection`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "subject-processing-result", "list", "--workflow", wf1.id, "--profile", profile], overrides);
      assert.ok(!new URL(requests[0].url).searchParams.has("$select"));
      assert.deepEqual(result.subjectProcessingResults, [
        { id: spr1.id, subjectType: "provisioningObject", processingStatus: "completed", failedTasksCount: 0, totalTasksCount: 3, totalUnprocessedTasksCount: 0 },
        { id: spr2.id, subjectType: "user", processingStatus: "queued", failedTasksCount: 0, totalTasksCount: 1, totalUnprocessedTasksCount: 1 },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      const narrowed = await executeArgv(["entra", "lifecycle", "subject-processing-result", "list", "--workflow", wf1.id, "--profile", profile,
        "--fields", "id,subjectType"], overrides);
      assert.deepEqual(narrowed.subjectProcessingResults, [
        { id: spr1.id, subjectType: "provisioningObject" },
        { id: spr2.id, subjectType: "user" },
      ]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one subject processing result whole with no $select`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "subject-processing-result", "show", "--workflow", wf1.id, "--id", spr1.id, "--profile", profile], overrides);
      assert.ok(!new URL(requests[0].url).searchParams.has("$select"));
      assert.ok(!Object.hasOwn(result.subjectProcessingResult, "subject"));
      assert.deepEqual(result.subjectProcessingResult, {
        id: spr1.id,
        subjectType: "provisioningObject",
        processingStatus: "completed",
        failedTasksCount: 0,
        totalTasksCount: 3,
        totalUnprocessedTasksCount: 0,
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} subject reads refuse --select before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "subject-processing-result", "list", "--workflow", wf1.id, "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "subject-processing-result", "show", "--workflow", wf1.id, "--id", spr1.id, "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} run and processing-result reads require --workflow before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "list", "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--workflow/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "user-processing-result", "show", "--id", upr1.id, "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--workflow/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "subject-processing-result", "count", "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--workflow/.test(error.message),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts runs and processing results as scalars`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const runCount = await executeArgv(["entra", "lifecycle", "run", "count", "--workflow", wf1.id, "--profile", profile], overrides);
      assert.deepEqual(runCount, { count: { returned: 2, complete: true } });
      const userCount = await executeArgv(["entra", "lifecycle", "user-processing-result", "count", "--workflow", wf1.id, "--profile", profile], overrides);
      assert.deepEqual(userCount, { count: { returned: 2, complete: true } });
      const subjectCount = await executeArgv(["entra", "lifecycle", "subject-processing-result", "count", "--workflow", wf1.id, "--profile", profile], overrides);
      assert.deepEqual(subjectCount, { count: { returned: 2, complete: true } });
      assert.deepEqual(requests.map(request => request.url), [
        `https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/runs/$count`,
        `https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/userProcessingResults/$count`,
        `https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/subjectProcessingResults/$count`,
      ]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists run user processing results with minimal personal-data defaults`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "user-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], overrides);
      assert.deepEqual(result.userProcessingResults, [
        { id: rupr1.id, processingStatus: "failed", failedTasksCount: 1, totalTasksCount: 2, totalUnprocessedTasksCount: 0 },
        { id: rupr2.id, processingStatus: "completed", failedTasksCount: 0, totalTasksCount: 1, totalUnprocessedTasksCount: 0 },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra lifecycle run user-processing-result show --workflow <workflow-id> --run <run-id> --id <result-id>")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/runs/${run1.id}/userProcessingResults?`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(reportsScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists run user processing results with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "user-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile,
        "--filter", "processingStatus eq 'failed'"], overrides);
      assert.equal(result.count.returned, 2);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "processingStatus eq 'failed'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped run user list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "lifecycle", "run", "user-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.userProcessingResults.map(row => row.id), [rupr1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "lifecycle", "run", "user-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.userProcessingResults.map(row => row.id), [rupr2.id]);
      assert.deepEqual(second.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one run user processing result without the subject link`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "user-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rupr1.id, "--profile", profile], overrides);
      assert.ok(!Object.hasOwn(result.userProcessingResult, "subject"));
      assert.ok(!Object.hasOwn(result.userProcessingResult, "taskProcessingResults"));
      assert.deepEqual(result.userProcessingResult, {
        id: rupr1.id,
        processingStatus: "failed",
        failedTasksCount: 1,
        totalTasksCount: 2,
        totalUnprocessedTasksCount: 0,
      });
      assert.ok(new URL(requests[0].url).searchParams.has("$select"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown run user processing result ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "lifecycle", "run", "user-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", "result-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists run subject processing results whole with local field projection`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "subject-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], overrides);
      assert.ok(!new URL(requests[0].url).searchParams.has("$select"));
      assert.equal(requests[0].url, `https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/runs/${run1.id}/subjectProcessingResults`);
      assert.deepEqual(result.subjectProcessingResults, [
        { id: rspr1.id, subjectType: "user", processingStatus: "completed", failedTasksCount: 0, totalTasksCount: 2, totalUnprocessedTasksCount: 0 },
        { id: rspr2.id, subjectType: "provisioningObject", processingStatus: "queued", failedTasksCount: 0, totalTasksCount: 1, totalUnprocessedTasksCount: 1 },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      const narrowed = await executeArgv(["entra", "lifecycle", "run", "subject-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile,
        "--fields", "id,subjectType"], overrides);
      assert.deepEqual(narrowed.subjectProcessingResults, [
        { id: rspr1.id, subjectType: "user" },
        { id: rspr2.id, subjectType: "provisioningObject" },
      ]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one run subject processing result whole with no $select`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "subject-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rspr1.id, "--profile", profile], overrides);
      assert.ok(!new URL(requests[0].url).searchParams.has("$select"));
      assert.ok(!Object.hasOwn(result.subjectProcessingResult, "subject"));
      assert.deepEqual(result.subjectProcessingResult, {
        id: rspr1.id,
        subjectType: "user",
        processingStatus: "completed",
        failedTasksCount: 0,
        totalTasksCount: 2,
        totalUnprocessedTasksCount: 0,
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} run subject reads refuse --select before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "subject-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "subject-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rspr1.id, "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists run task processing results with compact rows and the reports scope`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "task-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], overrides);
      assert.deepEqual(result.taskProcessingResults, [
        { id: rtpr1.id, processingStatus: "failed", failureReason: `${longFailureReason.slice(0, 500)}... (truncated, ${longFailureReason.length} chars total)` },
        { id: rtpr2.id, processingStatus: "completed", failureReason: null },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra lifecycle run task-processing-result show --workflow <workflow-id> --run <run-id> --id <result-id>")));
      assert.ok(result.help.some(hint => hint.includes("--full")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/runs/${run1.id}/taskProcessingResults?`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(reportsScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists run task processing results with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "task-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile,
        "--filter", "processingStatus eq 'failed'"], overrides);
      assert.equal(result.count.returned, 2);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "processingStatus eq 'failed'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one run task processing result whole with failure detail`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "run", "task-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rtpr1.id, "--profile", profile, "--full"], overrides);
      assert.ok(!new URL(requests[0].url).searchParams.has("$select"));
      assert.ok(!Object.hasOwn(result.taskProcessingResult, "subject"));
      assert.ok(!Object.hasOwn(result.taskProcessingResult, "task"));
      assert.ok(!Object.hasOwn(result.taskProcessingResult, "workflowSubject"));
      assert.deepEqual(result.taskProcessingResult, {
        id: rtpr1.id,
        completedDateTime: "2023-01-20T17:16:03.4863553Z",
        createdDateTime: "2023-01-20T17:16:00.9095011Z",
        failureReason: longFailureReason,
        processingInfo: "User was already a member of all requested groups",
        processingStatus: "failed",
        startedDateTime: "2023-01-20T17:16:02.8025169Z",
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} run task reads refuse --select before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "task-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rtpr1.id, "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} run-nested reads require --workflow and --run before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "user-processing-result", "list", "--run", run1.id, "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--workflow/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "user-processing-result", "list", "--workflow", wf1.id, "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--run/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "subject-processing-result", "count", "--workflow", wf1.id, "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--run/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "task-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--id/.test(error.message),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts run-nested processing results as scalars`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const userCount = await executeArgv(["entra", "lifecycle", "run", "user-processing-result", "count", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], overrides);
      assert.deepEqual(userCount, { count: { returned: 2, complete: true } });
      const subjectCount = await executeArgv(["entra", "lifecycle", "run", "subject-processing-result", "count", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], overrides);
      assert.deepEqual(subjectCount, { count: { returned: 2, complete: true } });
      const taskCount = await executeArgv(["entra", "lifecycle", "run", "task-processing-result", "count", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], overrides);
      assert.deepEqual(taskCount, { count: { returned: 2, complete: true } });
      assert.deepEqual(requests.map(request => request.url), [
        `https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/runs/${run1.id}/userProcessingResults/$count`,
        `https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/runs/${run1.id}/subjectProcessingResults/$count`,
        `https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/runs/${run1.id}/taskProcessingResults/$count`,
      ]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied run-nested reads surface scope, role and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "lifecycle", "run", "task-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("LifecycleWorkflows-Reports.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Lifecycle Workflows Administrator")));
        assert.ok(error.suggestions.some(hint => hint.includes("Entra ID Governance or Microsoft Entra Suite")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} navigation names fail as unknown properties on run-nested reads`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "user-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile, "--select", "id,subject"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property subject/.test(error.message)
          && error.suggestions.join("\n").includes("Known properties: "),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "task-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rtpr1.id, "--profile", profile, "--fields", "id,task"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property task/.test(error.message),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} run-nested reads refuse beta before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "user-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile, "--api-version", "beta"], overrides),
        error => error.code === "VALIDATION_ERROR" && /v1\.0 only/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "task-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rtpr1.id, "--profile", profile, "--api-version", "beta"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists task reports with compact rows and the reports scope`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "task-report", "list", "--workflow", wf1.id, "--profile", profile], overrides);
      assert.deepEqual(result.taskReports, [
        { id: tr1.id, runId: tr1.runId, processingStatus: "completed", totalUsersCount: 2, failedUsersCount: 0, successfulUsersCount: 2 },
        { id: tr2.id, runId: tr2.runId, processingStatus: "inProgress", totalUsersCount: 1, failedUsersCount: 0, successfulUsersCount: 0 },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra lifecycle task-report show --workflow <workflow-id> --id <report-id>")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/taskReports?`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(reportsScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists task reports with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "task-report", "list", "--workflow", wf1.id, "--profile", profile,
        "--filter", "processingStatus eq 'completed'"], overrides);
      assert.equal(result.count.returned, 2);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "processingStatus eq 'completed'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped task-report list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "lifecycle", "task-report", "list", "--workflow", wf1.id, "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.taskReports.map(row => row.id), [tr1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "lifecycle", "task-report", "list", "--workflow", wf1.id, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.taskReports.map(row => row.id), [tr2.id]);
      assert.deepEqual(second.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one task report whole with no $select`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "task-report", "show", "--workflow", wf1.id, "--id", tr1.id, "--profile", profile], overrides);
      assert.ok(!new URL(requests[0].url).searchParams.has("$select"));
      assert.ok(!Object.hasOwn(result.taskReport, "task"));
      assert.ok(!Object.hasOwn(result.taskReport, "taskDefinition"));
      assert.ok(!Object.hasOwn(result.taskReport, "taskProcessingResults"));
      assert.deepEqual(result.taskReport, {
        id: tr1.id,
        runId: tr1.runId,
        processingStatus: "completed",
        successfulUsersCount: 2,
        failedUsersCount: 0,
        unprocessedUsersCount: 0,
        totalUsersCount: 2,
        startedDateTime: "2022-08-24T23:28:04.5785337Z",
        completedDateTime: "2022-08-24T23:28:11.1348863Z",
        lastUpdatedDateTime: "2022-08-24T23:33:09.1980357Z",
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown task report ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "lifecycle", "task-report", "show", "--workflow", wf1.id, "--id", "report-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows the task behind a task report whole with arguments`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "task-report", "task", "show", "--workflow", wf1.id, "--report", tr1.id, "--profile", profile], overrides);
      assert.ok(!new URL(requests[0].url).searchParams.has("$select"));
      assert.ok(!Object.hasOwn(result.task, "taskProcessingResults"));
      assert.deepEqual(result.task, reportTask);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows the task definition behind a task report whole`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "task-report", "task-definition", "show", "--workflow", wf1.id, "--report", tr1.id, "--profile", profile], overrides);
      assert.ok(!new URL(requests[0].url).searchParams.has("$select"));
      assert.deepEqual(result.taskDefinition, td1);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} task-report reads refuse --select before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "show", "--workflow", wf1.id, "--id", tr1.id, "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "task", "show", "--workflow", wf1.id, "--report", tr1.id, "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "task-definition", "show", "--workflow", wf1.id, "--report", tr1.id, "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} task-report reads require --workflow and --report before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "list", "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--workflow/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "count", "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--workflow/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "task", "show", "--workflow", wf1.id, "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--report/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "task-definition", "show", "--report", tr1.id, "--profile", profile], overrides),
        error => error.code === "VALIDATION_ERROR" && /--workflow/.test(error.message),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts task reports as scalars`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const reportCount = await executeArgv(["entra", "lifecycle", "task-report", "count", "--workflow", wf1.id, "--profile", profile], overrides);
      assert.deepEqual(reportCount, { count: { returned: 2, complete: true } });
      assert.deepEqual(requests.map(request => request.url), [
        `https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/${wf1.id}/taskReports/$count`,
      ]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied task-report reads surface scope, role and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "lifecycle", "task-report", "list", "--workflow", wf1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("LifecycleWorkflows-Reports.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Lifecycle Workflows Administrator")));
        assert.ok(error.suggestions.some(hint => hint.includes("Entra ID Governance or Microsoft Entra Suite")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} task navigation names fail as unknown properties on task-report reads`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "list", "--workflow", wf1.id, "--profile", profile, "--select", "id,taskProcessingResults"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property taskProcessingResults/.test(error.message)
          && error.suggestions.join("\n").includes("Known properties: "),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "show", "--workflow", wf1.id, "--id", tr1.id, "--profile", profile, "--fields", "id,task"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property task/.test(error.message),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} task-report reads refuse beta before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "list", "--workflow", wf1.id, "--profile", profile, "--api-version", "beta"], overrides),
        error => error.code === "VALIDATION_ERROR" && /v1\.0 only/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-report", "task", "show", "--workflow", wf1.id, "--report", tr1.id, "--profile", profile, "--api-version", "beta"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists workflow templates with compact rows`, async () => {
    const state = setupProfiles();
    try {
      const { calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "workflow-template", "list", "--profile", profile], overrides);
      assert.deepEqual(result.workflowTemplates, [
        { id: tmpl1.id, displayName: "Onboard pre-hire employee", category: "joiner" },
        { id: tmpl2.id, displayName: "Offboard an employee", category: "leaver" },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(lifecycleScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one workflow template whole with no $select`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "workflow-template", "show", "--id", tmpl1.id, "--profile", profile], overrides);
      assert.deepEqual(result.workflowTemplate, {
        id: tmpl1.id,
        category: "joiner",
        description: "Configure pre-hire tasks for onboarding employees before their first day",
        displayName: "Onboard pre-hire employee",
        executionConditions: { scope: { rule: "department eq 'Marketing'" } },
      });
      assert.ok(!new URL(requests[0].url).searchParams.has("$select"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} workflow-template show refuses --select before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "workflow-template", "show", "--id", tmpl1.id, "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists task definitions with compact rows`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "task-definition", "list", "--profile", profile], overrides);
      assert.deepEqual(result.taskDefinitions, [
        { id: td1.id, displayName: "Add user to groups", category: "joiner,leaver,mover", version: 1 },
        { id: td2.id, displayName: "Disable user account" },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one task definition with its parameters`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "task-definition", "show", "--id", td1.id, "--profile", profile], overrides);
      assert.deepEqual(result.taskDefinition, {
        id: td1.id,
        category: "joiner,leaver,mover",
        continueOnError: false,
        description: "Add user to selected groups",
        displayName: "Add user to groups",
        parameters: [{ name: "groupID", values: [], valueType: "string" }],
        version: 1,
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows the tenant lifecycle settings singleton`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "lifecycle", "settings", "show", "--profile", profile], overrides);
      assert.deepEqual(result.settings, settings);
      assert.ok(new URL(requests[0].url).searchParams.get("$select").split(",").includes("workflowScheduleIntervalInHours"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts workflows, templates and task definitions as scalars`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const workflowCount = await executeArgv(["entra", "lifecycle", "workflow", "count", "--profile", profile], overrides);
      assert.deepEqual(workflowCount, { count: { returned: 2, complete: true } });
      const templateCount = await executeArgv(["entra", "lifecycle", "workflow-template", "count", "--profile", profile], overrides);
      assert.deepEqual(templateCount, { count: { returned: 2, complete: true } });
      const definitionCount = await executeArgv(["entra", "lifecycle", "task-definition", "count", "--profile", profile], overrides);
      assert.deepEqual(definitionCount, { count: { returned: 2, complete: true } });
      assert.deepEqual(requests.map(request => request.url), [
        "https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflows/$count",
        "https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/workflowTemplates/$count",
        "https://graph.microsoft.com/v1.0/identityGovernance/lifecycleWorkflows/taskDefinitions/$count",
      ]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lifecycle counts refuse collection flags before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "workflow", "count", "--profile", profile, "--limit", "1"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-definition", "count", "--profile", profile, "--filter", "category eq 'joiner'"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "count", "--workflow", wf1.id, "--profile", profile, "--limit", "1"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} navigation names fail as unknown properties on lifecycle reads`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "workflow", "list", "--profile", profile, "--select", "id,tasks"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property tasks/.test(error.message)
          && error.suggestions.join("\n").includes("Known properties: "),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "workflow", "show", "--id", wf1.id, "--profile", profile, "--select", "id,createdBy"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property createdBy/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "list", "--workflow", wf1.id, "--profile", profile, "--select", "id,userProcessingResults"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property userProcessingResults/.test(error.message)
          && error.suggestions.join("\n").includes("Known properties: "),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied lifecycle reads surface scope, role and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "lifecycle", "workflow", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("LifecycleWorkflows-Workflow.ReadBasic.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Lifecycle Workflows Administrator")));
        assert.ok(error.suggestions.some(hint => hint.includes("Entra ID Governance or Microsoft Entra Suite")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "lifecycle", "task-definition", "show", "--id", td1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("LifecycleWorkflows.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Global Reader")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "lifecycle", "run", "list", "--workflow", wf1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("LifecycleWorkflows-Reports.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Lifecycle Workflows Administrator")));
        assert.ok(error.suggestions.some(hint => hint.includes("Entra ID Governance or Microsoft Entra Suite")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty lifecycle lists stay definitive`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "lifecycle", "workflow", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.workflows, []);
      assert.deepEqual(listed.count, { returned: 0, complete: true });
      assert.ok(listed.help.some(hint => hint.includes("0 workflows matched")));
      const templated = await executeArgv(["entra", "lifecycle", "workflow-template", "list", "--profile", profile], overrides);
      assert.deepEqual(templated.workflowTemplates, []);
      assert.ok(templated.help.some(hint => hint.includes("0 workflow templates matched")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown lifecycle properties fail naming the known set`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "task-definition", "list", "--profile", profile, "--select", "id,ownerId"], overrides),
        error => error.code === "VALIDATION_ERROR" && /Unknown property ownerId/.test(error.message)
          && error.suggestions.join("\n").includes("Known properties: "),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lifecycle reads refuse beta before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "workflow", "list", "--profile", profile, "--api-version", "beta"], overrides),
        error => error.code === "VALIDATION_ERROR" && /v1\.0 only/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "settings", "show", "--profile", profile, "--api-version", "beta"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "lifecycle", "run", "list", "--workflow", wf1.id, "--profile", profile, "--api-version", "beta"], overrides),
        error => error.code === "VALIDATION_ERROR" && /v1\.0 only/.test(error.message),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts lifecycle reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runLifecycleCli(["entra", "lifecycle", "workflow", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.workflows.map(row => row.id), [wf1.id, wf2.id]);
      assert.deepEqual(listOut.count, { returned: 2, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runLifecycleCli(["entra", "lifecycle", "workflow", "show", "--id", wf1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.equal(decode(shown.stdout).workflow.id, wf1.id);
      assert.ok(!Object.hasOwn(decode(shown.stdout).workflow, "tasks"));

      const counted = runLifecycleCli(["entra", "lifecycle", "workflow", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).count, { returned: 2, complete: true });

      const templates = runLifecycleCli(["entra", "lifecycle", "workflow-template", "list", "--profile", profile], state, mode);
      assert.equal(templates.status, 0, templates.stdout);
      assert.deepEqual(decode(templates.stdout).workflowTemplates.map(row => row.id), [tmpl1.id, tmpl2.id]);

      const templateShown = runLifecycleCli(["entra", "lifecycle", "workflow-template", "show", "--id", tmpl1.id, "--profile", profile], state, mode);
      assert.equal(templateShown.status, 0, templateShown.stdout);
      assert.deepEqual(decode(templateShown.stdout).workflowTemplate.id, tmpl1.id);

      const templateCounted = runLifecycleCli(["entra", "lifecycle", "workflow-template", "count", "--profile", profile], state, mode);
      assert.equal(templateCounted.status, 0, templateCounted.stdout);
      assert.deepEqual(decode(templateCounted.stdout).count, { returned: 2, complete: true });

      const definitions = runLifecycleCli(["entra", "lifecycle", "task-definition", "list", "--profile", profile], state, mode);
      assert.equal(definitions.status, 0, definitions.stdout);
      assert.deepEqual(decode(definitions.stdout).taskDefinitions.map(row => row.id), [td1.id, td2.id]);

      const definitionShown = runLifecycleCli(["entra", "lifecycle", "task-definition", "show", "--id", td1.id, "--profile", profile], state, mode);
      assert.equal(definitionShown.status, 0, definitionShown.stdout);
      assert.deepEqual(decode(definitionShown.stdout).taskDefinition, td1);

      const definitionCounted = runLifecycleCli(["entra", "lifecycle", "task-definition", "count", "--profile", profile], state, mode);
      assert.equal(definitionCounted.status, 0, definitionCounted.stdout);
      assert.deepEqual(decode(definitionCounted.stdout).count, { returned: 2, complete: true });

      const settingsShown = runLifecycleCli(["entra", "lifecycle", "settings", "show", "--profile", profile], state, mode);
      assert.equal(settingsShown.status, 0, settingsShown.stdout);
      assert.deepEqual(decode(settingsShown.stdout).settings, settings);
      assert.ok(!settingsShown.stdout.includes(`opaque-fixture-${mode}-token`));

      const runsListed = runLifecycleCli(["entra", "lifecycle", "run", "list", "--workflow", wf1.id, "--profile", profile], state, mode);
      assert.equal(runsListed.status, 0, runsListed.stdout);
      assert.equal(runsListed.stderr, "");
      assert.deepEqual(decode(runsListed.stdout).runs.map(row => row.id), [run1.id, run2.id]);
      assert.deepEqual(decode(runsListed.stdout).count, { returned: 2, complete: true });

      const runShown = runLifecycleCli(["entra", "lifecycle", "run", "show", "--workflow", wf1.id, "--id", run1.id, "--profile", profile], state, mode);
      assert.equal(runShown.status, 0, runShown.stdout);
      assert.deepEqual(decode(runShown.stdout).run.id, run1.id);
      assert.ok(!Object.hasOwn(decode(runShown.stdout).run, "userProcessingResults"));

      const runsCounted = runLifecycleCli(["entra", "lifecycle", "run", "count", "--workflow", wf1.id, "--profile", profile], state, mode);
      assert.equal(runsCounted.status, 0, runsCounted.stdout);
      assert.deepEqual(decode(runsCounted.stdout).count, { returned: 2, complete: true });

      const usersListed = runLifecycleCli(["entra", "lifecycle", "user-processing-result", "list", "--workflow", wf1.id, "--profile", profile], state, mode);
      assert.equal(usersListed.status, 0, usersListed.stdout);
      assert.deepEqual(decode(usersListed.stdout).userProcessingResults.map(row => row.id), [upr1.id, upr2.id]);
      assert.ok(!Object.hasOwn(decode(usersListed.stdout).userProcessingResults[0], "subject"));

      const userShown = runLifecycleCli(["entra", "lifecycle", "user-processing-result", "show", "--workflow", wf1.id, "--id", upr1.id, "--profile", profile], state, mode);
      assert.equal(userShown.status, 0, userShown.stdout);
      assert.deepEqual(decode(userShown.stdout).userProcessingResult.id, upr1.id);

      const usersCounted = runLifecycleCli(["entra", "lifecycle", "user-processing-result", "count", "--workflow", wf1.id, "--profile", profile], state, mode);
      assert.equal(usersCounted.status, 0, usersCounted.stdout);
      assert.deepEqual(decode(usersCounted.stdout).count, { returned: 2, complete: true });

      const subjectsListed = runLifecycleCli(["entra", "lifecycle", "subject-processing-result", "list", "--workflow", wf1.id, "--profile", profile], state, mode);
      assert.equal(subjectsListed.status, 0, subjectsListed.stdout);
      assert.deepEqual(decode(subjectsListed.stdout).subjectProcessingResults.map(row => row.id), [spr1.id, spr2.id]);

      const subjectShown = runLifecycleCli(["entra", "lifecycle", "subject-processing-result", "show", "--workflow", wf1.id, "--id", spr1.id, "--profile", profile], state, mode);
      assert.equal(subjectShown.status, 0, subjectShown.stdout);
      assert.deepEqual(decode(subjectShown.stdout).subjectProcessingResult, {
        id: spr1.id,
        subjectType: "provisioningObject",
        processingStatus: "completed",
        failedTasksCount: 0,
        totalTasksCount: 3,
        totalUnprocessedTasksCount: 0,
      });

      const subjectsCounted = runLifecycleCli(["entra", "lifecycle", "subject-processing-result", "count", "--workflow", wf1.id, "--profile", profile], state, mode);
      assert.equal(subjectsCounted.status, 0, subjectsCounted.stdout);
      assert.deepEqual(decode(subjectsCounted.stdout).count, { returned: 2, complete: true });

      const runUsersListed = runLifecycleCli(["entra", "lifecycle", "run", "user-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], state, mode);
      assert.equal(runUsersListed.status, 0, runUsersListed.stdout);
      assert.deepEqual(decode(runUsersListed.stdout).userProcessingResults.map(row => row.id), [rupr1.id, rupr2.id]);
      assert.ok(!Object.hasOwn(decode(runUsersListed.stdout).userProcessingResults[0], "subject"));

      const runUserShown = runLifecycleCli(["entra", "lifecycle", "run", "user-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rupr1.id, "--profile", profile], state, mode);
      assert.equal(runUserShown.status, 0, runUserShown.stdout);
      assert.deepEqual(decode(runUserShown.stdout).userProcessingResult.id, rupr1.id);

      const runUsersCounted = runLifecycleCli(["entra", "lifecycle", "run", "user-processing-result", "count", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], state, mode);
      assert.equal(runUsersCounted.status, 0, runUsersCounted.stdout);
      assert.deepEqual(decode(runUsersCounted.stdout).count, { returned: 2, complete: true });

      const runSubjectsListed = runLifecycleCli(["entra", "lifecycle", "run", "subject-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], state, mode);
      assert.equal(runSubjectsListed.status, 0, runSubjectsListed.stdout);
      assert.deepEqual(decode(runSubjectsListed.stdout).subjectProcessingResults.map(row => row.id), [rspr1.id, rspr2.id]);

      const runSubjectShown = runLifecycleCli(["entra", "lifecycle", "run", "subject-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rspr1.id, "--profile", profile], state, mode);
      assert.equal(runSubjectShown.status, 0, runSubjectShown.stdout);
      assert.deepEqual(decode(runSubjectShown.stdout).subjectProcessingResult, {
        id: rspr1.id,
        subjectType: "user",
        processingStatus: "completed",
        failedTasksCount: 0,
        totalTasksCount: 2,
        totalUnprocessedTasksCount: 0,
      });

      const runSubjectsCounted = runLifecycleCli(["entra", "lifecycle", "run", "subject-processing-result", "count", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], state, mode);
      assert.equal(runSubjectsCounted.status, 0, runSubjectsCounted.stdout);
      assert.deepEqual(decode(runSubjectsCounted.stdout).count, { returned: 2, complete: true });

      const runTasksListed = runLifecycleCli(["entra", "lifecycle", "run", "task-processing-result", "list", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], state, mode);
      assert.equal(runTasksListed.status, 0, runTasksListed.stdout);
      assert.deepEqual(decode(runTasksListed.stdout).taskProcessingResults.map(row => row.id), [rtpr1.id, rtpr2.id]);
      assert.ok(!Object.hasOwn(decode(runTasksListed.stdout).taskProcessingResults[0], "subject"));
      assert.ok(!Object.hasOwn(decode(runTasksListed.stdout).taskProcessingResults[0], "task"));

      const runTaskShown = runLifecycleCli(["entra", "lifecycle", "run", "task-processing-result", "show", "--workflow", wf1.id, "--run", run1.id, "--id", rtpr1.id, "--profile", profile, "--full"], state, mode);
      assert.equal(runTaskShown.status, 0, runTaskShown.stdout);
      assert.deepEqual(decode(runTaskShown.stdout).taskProcessingResult.failureReason, longFailureReason);
      assert.ok(!Object.hasOwn(decode(runTaskShown.stdout).taskProcessingResult, "workflowSubject"));

      const runTasksCounted = runLifecycleCli(["entra", "lifecycle", "run", "task-processing-result", "count", "--workflow", wf1.id, "--run", run1.id, "--profile", profile], state, mode);
      assert.equal(runTasksCounted.status, 0, runTasksCounted.stdout);
      assert.deepEqual(decode(runTasksCounted.stdout).count, { returned: 2, complete: true });

      const reportsListed = runLifecycleCli(["entra", "lifecycle", "task-report", "list", "--workflow", wf1.id, "--profile", profile], state, mode);
      assert.equal(reportsListed.status, 0, reportsListed.stdout);
      assert.deepEqual(decode(reportsListed.stdout).taskReports.map(row => row.id), [tr1.id, tr2.id]);
      assert.ok(!Object.hasOwn(decode(reportsListed.stdout).taskReports[0], "task"));

      const reportShown = runLifecycleCli(["entra", "lifecycle", "task-report", "show", "--workflow", wf1.id, "--id", tr1.id, "--profile", profile], state, mode);
      assert.equal(reportShown.status, 0, reportShown.stdout);
      assert.deepEqual(decode(reportShown.stdout).taskReport, {
        id: tr1.id,
        runId: tr1.runId,
        processingStatus: "completed",
        successfulUsersCount: 2,
        failedUsersCount: 0,
        unprocessedUsersCount: 0,
        totalUsersCount: 2,
        startedDateTime: "2022-08-24T23:28:04.5785337Z",
        completedDateTime: "2022-08-24T23:28:11.1348863Z",
        lastUpdatedDateTime: "2022-08-24T23:33:09.1980357Z",
      });

      const reportsCounted = runLifecycleCli(["entra", "lifecycle", "task-report", "count", "--workflow", wf1.id, "--profile", profile], state, mode);
      assert.equal(reportsCounted.status, 0, reportsCounted.stdout);
      assert.deepEqual(decode(reportsCounted.stdout).count, { returned: 2, complete: true });

      const reportTaskShown = runLifecycleCli(["entra", "lifecycle", "task-report", "task", "show", "--workflow", wf1.id, "--report", tr1.id, "--profile", profile], state, mode);
      assert.equal(reportTaskShown.status, 0, reportTaskShown.stdout);
      assert.deepEqual(decode(reportTaskShown.stdout).task, reportTask);

      const reportDefinitionShown = runLifecycleCli(["entra", "lifecycle", "task-report", "task-definition", "show", "--workflow", wf1.id, "--report", tr1.id, "--profile", profile], state, mode);
      assert.equal(reportDefinitionShown.status, 0, reportDefinitionShown.stdout);
      assert.deepEqual(decode(reportDefinitionShown.stdout).taskDefinition, td1);
    } finally {
      teardownProfiles(state);
    }
  });
}
