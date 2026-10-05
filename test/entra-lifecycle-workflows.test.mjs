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
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, workflows, workflowTemplates, taskDefinitions, settings, denied }),
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
    } finally {
      teardownProfiles(state);
    }
  });
}
