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
const definitionId = "98dcebed-c7f6-46f4-bcf3-4a3fccdb3e2a";
const otherDefinitionId = "cc701697-762c-439a-81f5-f58d680fde76";
const instanceId = "7bc18cf4-3d70-4009-bc8e-a7c5adb30849";
const otherInstanceId = "8d035c9d-798d-47fa-beb4-f986a4b8126f";

const d1 = {
  id: definitionId,
  displayName: `Quarterly guest access review ${"x".repeat(600)}`,
  status: "InProgress",
  createdDateTime: "2021-04-27T00:00:00Z",
  lastModifiedDateTime: "2021-05-01T00:00:00Z",
  descriptionForAdmins: "Quarterly review of guest access.",
  descriptionForReviewers: "Please review guest access.",
  scope: {
    "@odata.type": "#microsoft.graph.accessReviewQueryScope",
    query: "./members/microsoft.graph.user/?$count=true&$filter=(userType eq 'Guest')",
    queryType: "MicrosoftGraph",
  },
  instanceEnumerationScope: {
    "@odata.type": "#microsoft.graph.accessReviewQueryScope",
    query: "/groups?$filter=(groupTypes/any(c:c+eq+'Unified'))&$count=true",
    queryType: "MicrosoftGraph",
  },
  reviewers: [{ query: "./manager", queryType: "MicrosoftGraph", queryRoot: "decisions" }],
  fallbackReviewers: [],
  settings: { mailNotificationsEnabled: true, instanceDurationInDays: 25, autoApplyDecisionsEnabled: true, recommendationsEnabled: true },
};
const d2 = {
  id: otherDefinitionId,
  displayName: "Monthly role review",
  status: "Completed",
};
const definitions = [d1, d2];

const i1 = {
  id: instanceId,
  startDateTime: "2021-03-09T23:10:28.83Z",
  endDateTime: "2021-04-09T23:10:28.83Z",
  status: "InProgress",
  scope: {
    "@odata.type": "#microsoft.graph.accessReviewQueryScope",
    query: "/groups/f661fdd0-f0f7-42c0-8281-e89c6527ac63/members/microsoft.graph.user/",
    queryType: "MicrosoftGraph",
    queryRoot: null,
  },
  reviewers: [{ query: "./manager", queryType: "MicrosoftGraph", queryRoot: "decisions" }],
  fallbackReviewers: [],
};
const i2 = {
  id: otherInstanceId,
  startDateTime: "2021-01-09T23:10:28.83Z",
  endDateTime: "2021-02-09T23:10:28.83Z",
  status: "Completed",
};
const instances = [i1, i2];

const dec1 = {
  id: "139166ec-d214-4835-95aa-3c1d89581e51",
  accessReviewId: instanceId,
  decision: "NotReviewed",
  recommendation: "Deny",
  reviewedDateTime: null,
  reviewedBy: { id: "00000000-0000-0000-0000-000000000000", displayName: "", userPrincipalName: "" },
  appliedDateTime: null,
  applyResult: "New",
  principal: { "@odata.type": "#microsoft.graph.userIdentity", id: "1800bb2c-955d-4205-8471-3a6c3116435d", displayName: "guest example" },
  resourceLink: null,
  target: { "@odata.type": "#microsoft.graph.accessReviewInstanceDecisionItemUserTarget", userId: "04777c4b-4d43-4d32-a2e7-1eba5d03f8cf", userDisplayName: "Diego Siciliani" },
};
const dec2 = {
  id: "4bde8d40-9224-4aa3-936b-08d73e1baf47",
  accessReviewId: instanceId,
  decision: "Approve",
  recommendation: "Approve",
  reviewedDateTime: "2021-03-10T10:00:00Z",
};
const decisions = [dec1, dec2];

const r1 = {
  id: "4a2499b9-88e7-48ff-92f6-1c8b0f8d6b1a",
  displayName: "Adele Vance",
  userPrincipalName: "AdeleV@contoso.com",
  createdDateTime: "2021-03-09T23:10:28.83Z",
};
const r2 = {
  id: "7c3b2e1a-5d6f-4a8b-9c0d-2e3f4a5b6c7d",
  displayName: "Diego Siciliani",
  userPrincipalName: "DiegoS@contoso.com",
  createdDateTime: "2021-03-10T10:00:00Z",
};
const reviewers = [r1, r2];

const s1 = {
  id: "9ac05ca6-396a-469c-8a8b-bcb98fceb2dd",
  startDateTime: "2021-03-11T16:44:59.337Z",
  endDateTime: "2021-04-09T23:10:28.83Z",
  status: "InProgress",
  reviewers: [{ query: "./manager", queryType: "MicrosoftGraph", queryRoot: "decisions" }],
  fallbackReviewers: [],
};
const s2 = {
  id: "03266a48-8731-4cfc-8a60-b2fa6648a14c",
  startDateTime: "2021-04-10T00:00:00Z",
  endDateTime: "2021-05-10T00:00:00Z",
  status: "NotStarted",
};
const stages = [s1, s2];

const base = "/v1.0/identityGovernance/accessReviews/definitions";

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-ar-"));
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

function accessReviewTransport(denied = false) {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (denied) return json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } });
    if (path === base) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [d2] });
      return json(200, {
        value: [d1],
        "@odata.nextLink": `https://graph.microsoft.com${base}?%24skiptoken=page2`,
      });
    }
    const singleDefinition = new RegExp(`^${base}/([^/]+)$`).exec(path);
    if (singleDefinition) {
      const found = definitions.find(row => row.id === decodeURIComponent(singleDefinition[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such definition" } });
    }
    if (path === `${base}/${definitionId}/instances`) return json(200, { value: instances });
    const singleInstance = new RegExp(`^${base}/([^/]+)/instances/([^/]+)$`).exec(path);
    if (singleInstance) {
      const found = instances.find(row => row.id === decodeURIComponent(singleInstance[2]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such instance" } });
    }
    if (path === `${base}/${definitionId}/instances/${instanceId}/decisions`) {
      assert.ok(!url.searchParams.get("$select")?.split(",").includes("justification"));
      return json(200, { value: decisions });
    }
    const singleDecision = new RegExp(`^${base}/([^/]+)/instances/([^/]+)/decisions/([^/]+)$`).exec(path);
    if (singleDecision) {
      const found = decisions.find(row => row.id === decodeURIComponent(singleDecision[3]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such decision" } });
    }
    if (path === `${base}/${definitionId}/instances/${instanceId}/contactedReviewers`) return json(200, { value: reviewers });
    const singleReviewer = new RegExp(`^${base}/([^/]+)/instances/([^/]+)/contactedReviewers/([^/]+)$`).exec(path);
    if (singleReviewer) {
      const found = reviewers.find(row => row.id === decodeURIComponent(singleReviewer[3]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such reviewer" } });
    }
    if (path === `${base}/${definitionId}/instances/${instanceId}/stages`) return json(200, { value: stages });
    const singleStage = new RegExp(`^${base}/([^/]+)/instances/([^/]+)/stages/([^/]+)$`).exec(path);
    if (singleStage) {
      const found = stages.find(row => row.id === decodeURIComponent(singleStage[3]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such stage" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? accessReviewTransport();
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

function runAccessReviewCli(args, state, mode, scopes, input) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-access-reviews-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", input, stdio: [input === undefined ? "ignore" : "pipe", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, scopes, definitions, instances, decisions, reviewers, stages }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists access-review definitions with compact schedule rows`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode, transport(request => {
        const url = new URL(request.url);
        if (url.pathname === base) return json(200, { value: definitions });
        return json(404, { error: { code: "Unknown", message: "unexpected route" } });
      }));
      const result = await executeArgv(["entra", "access-review", "definition", "list", "--profile", profile], overrides);
      assert.deepEqual(result.definitions, [
        { id: d1.id, displayName: `${d1.displayName.slice(0, 500)}... (truncated, ${d1.displayName.length} chars total)`, status: "InProgress" },
        { id: d2.id, displayName: "Monthly role review", status: "Completed" },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("review schedules (a series)")));
      assert.ok(result.help.some(hint => hint.includes("entra access-review definition show --id <definition-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com${base}?`));
      assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,displayName,status");
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows an access-review definition without its instances`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "access-review", "definition", "show", "--id", d2.id, "--profile", profile], overrides);
      assert.deepEqual(result.definition, d2);
      assert.equal(result.help, undefined);
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/definitions/${d2.id}`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists instances of one definition as occurrences, not schedules`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "access-review", "instance", "list", "--definition", definitionId, "--profile", profile], overrides);
      assert.deepEqual(result.instances, [
        { id: i1.id, status: "InProgress", startDateTime: i1.startDateTime, endDateTime: i1.endDateTime },
        { id: i2.id, status: "Completed", startDateTime: i2.startDateTime, endDateTime: i2.endDateTime },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("never schedules themselves")));
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/definitions/${definitionId}/instances`));
      assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,status,startDateTime,endDateTime");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows an access-review instance with occurrence state`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "access-review", "instance", "show",
        "--definition", definitionId, "--id", i1.id, "--profile", profile], overrides);
      assert.deepEqual(result.instance, i1);
      assert.equal(result.help, undefined);
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/definitions/${definitionId}/instances/${i1.id}`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists decisions as read-only outcomes`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "access-review", "decision", "list",
        "--definition", definitionId, "--instance", instanceId, "--profile", profile], overrides);
      assert.deepEqual(result.decisions, [
        { id: dec1.id, accessReviewId: instanceId, decision: "NotReviewed", recommendation: "Deny" },
        { id: dec2.id, accessReviewId: instanceId, decision: "Approve", recommendation: "Approve" },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("never approves, denies or applies anything")));
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/instances/${instanceId}/decisions`));
      assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,accessReviewId,decision,recommendation");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows an access-review decision without approving anything`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "access-review", "decision", "show",
        "--definition", definitionId, "--instance", instanceId, "--id", dec1.id, "--profile", profile], overrides);
      assert.deepEqual(result.decision, dec1);
      assert.equal(result.help, undefined);
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/decisions/${dec1.id}`));
      assert.equal(new URL(requests[0].url).searchParams.get("$select"),
        "id,accessReviewId,decision,recommendation,reviewedDateTime,reviewedBy,appliedDateTime,applyResult,principal,resourceLink,target");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists contacted reviewers as recorded identities, not outcomes`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "access-review", "contacted-reviewer", "list",
        "--definition", definitionId, "--instance", instanceId, "--profile", profile], overrides);
      assert.deepEqual(result.contactedReviewers, [
        { id: r1.id, displayName: "Adele Vance", userPrincipalName: "AdeleV@contoso.com" },
        { id: r2.id, displayName: "Diego Siciliani", userPrincipalName: "DiegoS@contoso.com" },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("whether or not they were notified")));
      assert.ok(result.help.some(hint => hint.includes("entra access-review contacted-reviewer show --definition <definition-id>")));
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/instances/${instanceId}/contactedReviewers`));
      assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,displayName,userPrincipalName");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a contacted reviewer with the identity set`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "access-review", "contacted-reviewer", "show",
        "--definition", definitionId, "--instance", instanceId, "--id", r1.id, "--profile", profile], overrides);
      assert.deepEqual(result.contactedReviewer, r1);
      assert.equal(result.help, undefined);
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/contactedReviewers/${r1.id}`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists stages as sequential phases of one instance`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "access-review", "stage", "list",
        "--definition", definitionId, "--instance", instanceId, "--profile", profile], overrides);
      assert.deepEqual(result.stages, [
        { id: s1.id, status: "InProgress", startDateTime: s1.startDateTime, endDateTime: s1.endDateTime },
        { id: s2.id, status: "NotStarted", startDateTime: s2.startDateTime, endDateTime: s2.endDateTime },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("sequential phases")));
      assert.ok(result.help.some(hint => hint.includes("entra access-review stage show --definition <definition-id>")));
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/instances/${instanceId}/stages`));
      assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,status,startDateTime,endDateTime");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a stage with reviewer scopes and no decisions`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "access-review", "stage", "show",
        "--definition", definitionId, "--instance", instanceId, "--id", s1.id, "--profile", profile], overrides);
      assert.deepEqual(result.stage, s1);
      assert.equal(result.help, undefined);
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/stages/${s1.id}`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied access-review reads name the scope, role and P2/Governance requirement`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, accessReviewTransport(true));
      await assert.rejects(
        executeArgv(["entra", "access-review", "definition", "list", "--profile", profile], overrides),
        error => {
          assert.equal(error.code, "GRAPH_ERROR");
          const text = [error.message, ...error.suggestions].join("\n");
          assert.match(text, /AccessReview\.Read\.All/);
          assert.match(text, /Identity Governance Administrator/);
          assert.match(text, /P2 or ID Governance/);
          return true;
        },
      );
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists definitions end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runAccessReviewCli(
        ["entra", "access-review", "definition", "list", "--profile", profile],
        state, mode, "https://graph.microsoft.com/AccessReview.Read.All");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.definitions.length, 2);
      assert.deepEqual(output.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists decisions end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runAccessReviewCli(
        ["entra", "access-review", "decision", "list", "--definition", definitionId, "--instance", instanceId, "--profile", profile],
        state, mode, "https://graph.microsoft.com/AccessReview.Read.All");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.decisions, [
        { id: dec1.id, accessReviewId: instanceId, decision: "NotReviewed", recommendation: "Deny" },
        { id: dec2.id, accessReviewId: instanceId, decision: "Approve", recommendation: "Approve" },
      ]);
      assert.deepEqual(output.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists contacted reviewers end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runAccessReviewCli(
        ["entra", "access-review", "contacted-reviewer", "list", "--definition", definitionId, "--instance", instanceId, "--profile", profile],
        state, mode, "https://graph.microsoft.com/AccessReview.Read.All");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.contactedReviewers, [
        { id: r1.id, displayName: "Adele Vance", userPrincipalName: "AdeleV@contoso.com" },
        { id: r2.id, displayName: "Diego Siciliani", userPrincipalName: "DiegoS@contoso.com" },
      ]);
      assert.deepEqual(output.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable shows a stage end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runAccessReviewCli(
        ["entra", "access-review", "stage", "show", "--definition", definitionId, "--instance", instanceId, "--id", s1.id, "--profile", profile],
        state, mode, "https://graph.microsoft.com/AccessReview.Read.All");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      assert.deepEqual(decode(result.stdout).stage, s1);
    } finally {
      teardownProfiles(state);
    }
  });

  for (const [command, property] of [
    [["entra", "access-review", "definition", "list"], "createdBy"],
    [["entra", "access-review", "definition", "list"], "stageSettings"],
    [["entra", "access-review", "instance", "list"], "errors"],
    [["entra", "access-review", "decision", "list"], "justification"],
    [["entra", "access-review", "decision", "show"], "justification"],
    [["entra", "access-review", "contacted-reviewer", "list"], "justification"],
    [["entra", "access-review", "contacted-reviewer", "show"], "displayname"],
    [["entra", "access-review", "stage", "list"], "durationInDays"],
    [["entra", "access-review", "stage", "show"], "decisions"],
  ]) {
    for (const flag of ["select", "fields"]) {
      test(`${mode} ${command.join(" ")} rejects ${property} in --${flag} before credentials`, async () => {
        const state = setupProfiles();
        try {
          const args = [...command, "--profile", profile, `--${flag}`, `id,${property}`];
          if (command[2] === "instance") args.push("--definition", definitionId);
          if (["decision", "contacted-reviewer", "stage"].includes(command[2])) args.push("--definition", definitionId, "--instance", instanceId);
          if (command[3] === "show") args.push("--id", "00000000-0000-0000-0000-000000000000");
          const { requests, calls, overrides } = overridesFor(mode);
          await assert.rejects(
            executeArgv(args, overrides),
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

test("delegated definition reads truncate long display names with a --full hint", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated", transport(request => {
      const url = new URL(request.url);
      if (url.pathname === base) return json(200, { value: definitions });
      return json(404, { error: { code: "Unknown", message: "unexpected route" } });
    }));
    const result = await executeArgv(["entra", "access-review", "definition", "list", "--profile", "soc"], overrides);
    const quarterly = result.definitions.find(row => row.id === d1.id);
    assert.match(quarterly.displayName, /truncated, \d+ chars total/);
    assert.ok(result.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "access-review", "definition", "list", "--profile", "soc", "--full"], overrides);
    assert.equal(full.definitions.find(row => row.id === d1.id).displayName, d1.displayName);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated definition filters pass through as plain $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated", transport(request => {
      const url = new URL(request.url);
      if (url.pathname === base) return json(200, { value: definitions });
      return json(404, { error: { code: "Unknown", message: "unexpected route" } });
    }));
    const result = await executeArgv(["entra", "access-review", "definition", "list", "--profile", "soc",
      "--filter", "status eq 'InProgress'"], overrides);
    assert.equal(result.definitions.length, 2);
    const url = new URL(requests[0].url);
    assert.equal(url.searchParams.get("$filter"), "status eq 'InProgress'");
    assert.equal(url.searchParams.has("$count"), false);
    assert.equal(requests[0].headers.ConsistencyLevel, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated definition reads resume a capped list through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "access-review", "definition", "list", "--profile", "soc", "--limit", "1"], overrides);
    assert.equal(first.definitions.length, 1);
    assert.deepEqual(first.count, { returned: 1, complete: false, reason: first.count.reason });
    assert.ok(typeof first.cursor === "string" && first.cursor.length > 0);
    const second = await executeArgv(["entra", "access-review", "definition", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.definitions, [{ id: d2.id, displayName: "Monthly role review", status: "Completed" }]);
    assert.deepEqual(second.count, { returned: 1, complete: true });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated instance reads resume under the same definition binding", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated", transport(request => {
      const url = new URL(request.url);
      if (url.pathname === `${base}/${definitionId}/instances`) {
        if (url.searchParams.has("$skiptoken")) return json(200, { value: [i2] });
        return json(200, {
          value: [i1],
          "@odata.nextLink": `https://graph.microsoft.com${base}/${definitionId}/instances?%24skiptoken=page2`,
        });
      }
      return json(404, { error: { code: "Unknown", message: "unexpected route" } });
    }));
    const first = await executeArgv(["entra", "access-review", "instance", "list",
      "--definition", definitionId, "--profile", "soc", "--limit", "1"], overrides);
    assert.deepEqual(first.count, { returned: 1, complete: false, reason: first.count.reason });
    const second = await executeArgv(["entra", "access-review", "instance", "list",
      "--definition", definitionId, "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.instances, [
      { id: i2.id, status: "Completed", startDateTime: i2.startDateTime, endDateTime: i2.endDateTime },
    ]);
    await assert.rejects(
      executeArgv(["entra", "access-review", "instance", "list",
        "--definition", otherDefinitionId, "--profile", "soc", "--cursor", first.cursor], overrides),
      /resource or authentication context does not match/,
    );
  } finally {
    teardownProfiles(state);
  }
});

test("delegated instance list without its definition fails before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "access-review", "instance", "list", "--profile", "soc"], overrides),
      /--definition is required/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated decision list without its instance fails before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "access-review", "decision", "list", "--definition", definitionId, "--profile", "soc"], overrides),
      /--instance is required/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated contacted-reviewer filters pass through as plain $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated", transport(request => {
      const url = new URL(request.url);
      if (url.pathname.endsWith("/contactedReviewers")) return json(200, { value: reviewers });
      return json(404, { error: { code: "Unknown", message: "unexpected route" } });
    }));
    const result = await executeArgv(["entra", "access-review", "contacted-reviewer", "list",
      "--definition", definitionId, "--instance", instanceId, "--profile", "soc",
      "--filter", "displayName eq 'Adele Vance'"], overrides);
    assert.equal(result.contactedReviewers.length, 2);
    const url = new URL(requests[0].url);
    assert.equal(url.searchParams.get("$filter"), "displayName eq 'Adele Vance'");
    assert.equal(url.searchParams.has("$count"), false);
    assert.equal(requests[0].headers.ConsistencyLevel, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated stage filters pass through as plain $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated", transport(request => {
      const url = new URL(request.url);
      if (url.pathname.endsWith("/stages")) return json(200, { value: stages });
      return json(404, { error: { code: "Unknown", message: "unexpected route" } });
    }));
    const result = await executeArgv(["entra", "access-review", "stage", "list",
      "--definition", definitionId, "--instance", instanceId, "--profile", "soc",
      "--filter", "status eq 'InProgress'"], overrides);
    assert.equal(result.stages.length, 2);
    const url = new URL(requests[0].url);
    assert.equal(url.searchParams.get("$filter"), "status eq 'InProgress'");
    assert.equal(url.searchParams.has("$count"), false);
    assert.equal(requests[0].headers.ConsistencyLevel, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated contacted-reviewer reads resume under the same instance binding", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated", transport(request => {
      const url = new URL(request.url);
      if (url.pathname.endsWith("/contactedReviewers")) {
        if (url.searchParams.has("$skiptoken")) return json(200, { value: [r2] });
        return json(200, {
          value: [r1],
          "@odata.nextLink": `https://graph.microsoft.com${base}/${definitionId}/instances/${instanceId}/contactedReviewers?%24skiptoken=page2`,
        });
      }
      return json(404, { error: { code: "Unknown", message: "unexpected route" } });
    }));
    const first = await executeArgv(["entra", "access-review", "contacted-reviewer", "list",
      "--definition", definitionId, "--instance", instanceId, "--profile", "soc", "--limit", "1"], overrides);
    assert.deepEqual(first.count, { returned: 1, complete: false, reason: first.count.reason });
    const second = await executeArgv(["entra", "access-review", "contacted-reviewer", "list",
      "--definition", definitionId, "--instance", instanceId, "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.contactedReviewers, [
      { id: r2.id, displayName: "Diego Siciliani", userPrincipalName: "DiegoS@contoso.com" },
    ]);
    await assert.rejects(
      executeArgv(["entra", "access-review", "contacted-reviewer", "list",
        "--definition", definitionId, "--instance", otherInstanceId, "--profile", "soc", "--cursor", first.cursor], overrides),
      /resource or authentication context does not match/,
    );
  } finally {
    teardownProfiles(state);
  }
});

test("delegated contacted-reviewer list without its instance fails before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "access-review", "contacted-reviewer", "list", "--definition", definitionId, "--profile", "soc"], overrides),
      /--instance is required/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated stage show without its parents fails before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "access-review", "stage", "show", "--id", s1.id, "--profile", "soc"], overrides),
      /--definition is required/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated decision show without its instance fails before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "access-review", "decision", "show", "--definition", definitionId, "--id", dec1.id, "--profile", "soc"], overrides),
      /--instance is required/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated stage reads reject --fields outside the fetched selection", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "access-review", "stage", "list", "--definition", definitionId, "--instance", instanceId,
        "--profile", "soc", "--select", "id,status", "--fields", "id,reviewers"], overrides),
      /--fields reviewers was not fetched/,
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("application stage reads reject delegated scopes", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "access-review", "stage", "list", "--definition", definitionId, "--instance", instanceId, "--profile", "batch",
        "--scopes", "https://graph.microsoft.com/AccessReview.Read.All"], overrides),
      /configured Graph .default audience/,
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated decision reads reject --fields outside the fetched selection", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "access-review", "decision", "list", "--definition", definitionId, "--instance", instanceId,
        "--profile", "soc", "--select", "id,decision", "--fields", "id,recommendation"], overrides),
      /--fields recommendation was not fetched/,
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("application access-review reads reject delegated scopes", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "access-review", "definition", "list", "--profile", "batch",
        "--scopes", "https://graph.microsoft.com/AccessReview.Read.All"], overrides),
      /configured Graph .default audience/,
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

for (const [noun, key, rows] of [
  ["contacted-reviewer", "contactedReviewers", reviewers],
  ["stage", "stages", stages],
]) {
  for (const cursorArgs of [["--cursor", "-"], ["--cursor=-"]]) {
    test(`${noun} list resumes from stdin with ${cursorArgs.join(" ")}`, async () => {
      const state = setupProfiles();
      try {
        const { overrides } = overridesFor("delegated", transport(() => json(200, { value: rows })));
        const args = ["entra", "access-review", noun, "list", "--definition", definitionId,
          "--instance", instanceId, "--profile", "soc", "--limit", "1", "--select", "id"];
        const first = await executeArgv(args, overrides);
        assert.deepEqual(first[key], [{ id: rows[0].id }]);
        assert.equal(first.count.complete, false);
        assert.ok(first.help.some(hint => hint.includes("--cursor -") && hint.includes("stdin")));
        const resumed = runAccessReviewCli([...args, ...cursorArgs], state, "delegated", undefined, first.cursor);
        assert.equal(resumed.error, undefined);
        assert.equal(resumed.status, 0, resumed.stdout);
        const output = decode(resumed.stdout);
        assert.deepEqual(output[key], [{ id: rows[1].id }]);
        assert.deepEqual(output.count, { returned: 1, complete: true });
      } finally {
        teardownProfiles(state);
      }
    });
  }
}

for (const [noun, parents, key, field] of [
  ["definition", [], "definitions", "scope"],
  ["instance", ["--definition", definitionId], "instances", "scope"],
  ["decision", ["--definition", definitionId, "--instance", instanceId], "decisions", "principal"],
]) {
  test(`${noun} list resumes large cursors from stdin and replays full output`, async () => {
    const state = setupProfiles();
    try {
      const text = "x".repeat(140_000);
      const rows = [{ id: "first" }, { id: "second", [field]: { query: text } }, { id: "third" }];
      const { overrides } = overridesFor("delegated", transport(() => json(200, { value: rows })));
      const args = ["entra", "access-review", noun, "list", ...parents, "--profile", "soc", "--limit", "1", "--select", `id,${field}`];
      const first = await executeArgv(args, overrides);
      assert.ok(first.cursor.length > 128 * 1024);
      assert.ok(first.help.some(hint => hint.includes("--cursor -") && hint.includes("stdin")));
      const resumed = runAccessReviewCli([...args, "--cursor", "-"], state, "delegated", undefined, first.cursor);
      assert.equal(resumed.error, undefined);
      assert.equal(resumed.status, 0, resumed.stdout);
      const output = decode(resumed.stdout);
      assert.deepEqual(output[key], [{ id: "second", [field]: { query: `${"x".repeat(500)}... (truncated, 140000 chars total)` } }]);
      assert.equal(output.count.complete, false);
      assert.ok(output.help.some(hint => hint.includes("--full") && hint.includes("--cursor -")));
      assert.ok(output.help.some(hint => hint.includes("original input cursor on stdin")));
      assert.ok(output.help.every(hint => !hint.includes(first.cursor) && !hint.includes(output.cursor)));
      const replay = runAccessReviewCli([...args, "--cursor", "-", "--full"], state, "delegated", undefined, first.cursor);
      assert.equal(replay.error, undefined);
      assert.equal(replay.status, 0, replay.stdout);
      assert.deepEqual(decode(replay.stdout)[key], [rows[1]]);
    } finally {
      teardownProfiles(state);
    }
  });
}

for (const [noun, action, parents, key, fields] of [
  ["definition", "list", [], "definitions", ["scope", "instanceEnumerationScope", "reviewers", "fallbackReviewers", "settings"]],
  ["definition", "show", ["--id", definitionId], "definition", ["scope", "instanceEnumerationScope", "reviewers", "fallbackReviewers", "settings"]],
  ["instance", "list", ["--definition", definitionId], "instances", ["scope", "reviewers", "fallbackReviewers"]],
  ["instance", "show", ["--definition", definitionId, "--id", instanceId], "instance", ["scope", "reviewers", "fallbackReviewers"]],
  ["decision", "list", ["--definition", definitionId, "--instance", instanceId], "decisions", ["reviewedBy", "principal", "target"]],
]) {
  test(`${noun} ${action} truncates nested text while preserving structure and full output`, async () => {
    const state = setupProfiles();
    try {
      const value = { query: "x".repeat(700), nested: [null, { text: "y".repeat(501), boundary: "z".repeat(500), enabled: true, count: 2 }, ["short"]] };
      const compact = { query: `${"x".repeat(500)}... (truncated, 700 chars total)`, nested: [null, { text: `${"y".repeat(500)}... (truncated, 501 chars total)`, boundary: "z".repeat(500), enabled: true, count: 2 }, ["short"]] };
      const row = Object.fromEntries(fields.map(field => [field, value]));
      const expected = Object.fromEntries(fields.map(field => [field, compact]));
      const { overrides } = overridesFor("delegated", transport(() => json(200, action === "list" ? { value: [row] } : row)));
      const args = ["entra", "access-review", noun, action, ...parents, "--profile", "soc", "--select", fields.join(",")];
      const result = await executeArgv(args, overrides);
      assert.deepEqual(result[key], action === "list" ? [expected] : expected);
      assert.ok(result.help.some(hint => hint.includes("--full")));
      const full = await executeArgv([...args, "--full"], overrides);
      assert.deepEqual(full[key], action === "list" ? [row] : row);
      assert.ok(!(full.help ?? []).some(hint => hint.includes("--full")));
      assert.equal(value.query.length, 700);
    } finally {
      teardownProfiles(state);
    }
  });
}
