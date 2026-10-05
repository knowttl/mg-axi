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
const lifecycleScopes = ["https://graph.microsoft.com/Directory.Read.All"];
const templateScopes = ["https://graph.microsoft.com/GroupSettings.Read.All"];

const longEmails = `admin@contoso.com${";notify-extra-long-admin-address".repeat(30)}@contoso.com`;
const policy1 = {
  id: "policy-selected-1",
  groupLifetimeInDays: 180,
  managedGroupTypes: "Selected",
  alternateNotificationEmails: "admin@contoso.com",
};
const policy2 = {
  id: "policy-all-2",
  groupLifetimeInDays: 365,
  managedGroupTypes: "All",
  alternateNotificationEmails: longEmails,
};
const policy3 = { id: "policy-bare-3" };
const policies = [policy1, policy2, policy3];

const longDescription = `Tenant-wide Microsoft 365 group settings${" with considerable detail".repeat(30)}`;
const template1 = {
  id: "08d542b9-071f-4e16-94b0-74abb372e3d9",
  deletedDateTime: null,
  displayName: "Group.Unified.Guest",
  description: "Settings for a specific Unified Group",
  values: [
    {
      name: "AllowToAddGuests",
      type: "System.Boolean",
      defaultValue: "true",
      description: "Flag indicating if guests are allowed in a specific Unified Group.",
    },
  ],
};
const template2 = {
  id: "6225ff8b-3645-4c7e-8842-01d0e4b0f8d1",
  deletedDateTime: null,
  displayName: "Group.Unified",
  description: longDescription,
  values: [],
};
const template3 = { id: "template-bare-3" };
const templates = [template1, template2, template3];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-group-lifecycle-"));
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

function lifecycleTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/groupLifecyclePolicies/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(policies.length) };
    }
    if (path === "/v1.0/groupSettingTemplates/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(templates.length) };
    }
    if (path === "/v1.0/groupLifecyclePolicies") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [policy3] });
      return json(200, {
        value: [policy1, policy2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/groupLifecyclePolicies?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/groupSettingTemplates") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [template3] });
      return json(200, {
        value: [template1, template2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/groupSettingTemplates?%24skiptoken=page2",
      });
    }
    const policy = /^\/v1\.0\/groupLifecyclePolicies\/([^/]+)$/.exec(path);
    if (policy) {
      const found = policies.find(row => row.id === decodeURIComponent(policy[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such policy" } });
    }
    const template = /^\/v1\.0\/groupSettingTemplates\/([^/]+)$/.exec(path);
    if (template) {
      const found = templates.find(row => row.id === decodeURIComponent(template[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such template" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
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
    "--import", pathToFileURL(resolve("test/fixtures/read-group-lifecycle-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, policies: [policy1, policy2], templates: [template1, template2], denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists lifecycle policies with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", profile,
        "--select", "id,groupLifetimeInDays,managedGroupTypes,alternateNotificationEmails"], overrides);
      assert.deepEqual(result.groupLifecyclePolicies, [
        { id: policy1.id, groupLifetimeInDays: 180, managedGroupTypes: "Selected", alternateNotificationEmails: "admin@contoso.com" },
        { id: policy2.id, groupLifetimeInDays: 365, managedGroupTypes: "All", alternateNotificationEmails: longEmails.slice(0, 500) + `... (truncated, ${longEmails.length} chars total)` },
        { id: policy3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra group-lifecycle-policy show --id <policy-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/groupLifecyclePolicies?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(lifecycleScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists lifecycle policies with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", profile,
        "--filter", "managedGroupTypes eq 'Selected'"], overrides);
      assert.equal(result.count.returned, 3);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "managedGroupTypes eq 'Selected'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped lifecycle-policy list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.groupLifecyclePolicies.map(row => row.id), [policy1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.groupLifecyclePolicies.map(row => row.id), [policy2.id, policy3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one lifecycle policy with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group-lifecycle-policy", "show", "--id", policy1.id, "--profile", profile,
        "--select", "id,alternateNotificationEmails,groupLifetimeInDays,managedGroupTypes"], overrides);
      assert.deepEqual(result.groupLifecyclePolicy, policy1);
      assert.equal(result.help, undefined);
      const missing = await executeArgv(["entra", "group-lifecycle-policy", "show", "--id", policy3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.groupLifecyclePolicy, { id: policy3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  for (const id of ["delta()", "DELTA()", "delta(", "delta)", "delta%28%29"]) {
    test(`${mode} lifecycle-policy show rejects function-style identifier ${id} before credentials`, async () => {
      const state = setupProfiles();
      try {
        const delta = transport(() => json(200, { value: [policy1], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/groupLifecyclePolicies/delta()" }));
        const { requests, calls, overrides } = overridesFor(mode, delta);
        await assert.rejects(
          executeArgv(["entra", "group-lifecycle-policy", "show", "--id", id, "--profile", profile], overrides),
          { code: "VALIDATION_ERROR" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }

  test(`${mode} counts lifecycle policies as one scalar without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group-lifecycle-policy", "count", "--profile", profile], overrides);
      assert.deepEqual(result, { groupLifecyclePolicyCount: 3 });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/groupLifecyclePolicies/$count");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists setting templates with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group-setting-template", "list", "--profile", profile,
        "--select", "id,displayName,description"], overrides);
      assert.deepEqual(result.groupSettingTemplates, [
        { id: template1.id, displayName: "Group.Unified.Guest", description: "Settings for a specific Unified Group" },
        { id: template2.id, displayName: "Group.Unified", description: longDescription.slice(0, 500) + `... (truncated, ${longDescription.length} chars total)` },
        { id: template3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra group-setting-template show --id <template-id>")));
      assert.ok(result.help.some(hint => hint.includes("--filter is unsupported") || hint.includes("offer no --filter")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/groupSettingTemplates?"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(templateScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped setting-template list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "group-setting-template", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.groupSettingTemplates.map(row => row.id), [template1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "group-setting-template", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.groupSettingTemplates.map(row => row.id), [template2.id, template3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one setting template with values and the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group-setting-template", "show", "--id", template1.id, "--profile", profile,
        "--select", "id,deletedDateTime,displayName,description,values"], overrides);
      assert.deepEqual(result.groupSettingTemplate, template1);
      assert.equal(result.help, undefined);
      const missing = await executeArgv(["entra", "group-setting-template", "show", "--id", template3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.groupSettingTemplate, { id: template3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  for (const id of ["delta()", "DELTA()", "delta(", "delta)", "delta%28%29"]) {
    test(`${mode} setting-template show rejects function-style identifier ${id} before credentials`, async () => {
      const state = setupProfiles();
      try {
        const delta = transport(() => json(200, { value: [template1], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/groupSettingTemplates/delta()" }));
        const { requests, calls, overrides } = overridesFor(mode, delta);
        await assert.rejects(
          executeArgv(["entra", "group-setting-template", "show", "--id", id, "--profile", profile], overrides),
          { code: "VALIDATION_ERROR" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }

  test(`${mode} counts setting templates as one scalar without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group-setting-template", "count", "--profile", profile], overrides);
      assert.deepEqual(result, { groupSettingTemplateCount: 3 });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/groupSettingTemplates/$count");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty group lifecycle lists stay definitive`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.groupLifecyclePolicies, []);
      assert.deepEqual(listed.count, { returned: 0, complete: true });
      assert.ok(listed.help.some(hint => hint.includes("0 lifecycle policies matched")));
      const templated = await executeArgv(["entra", "group-setting-template", "list", "--profile", profile], overrides);
      assert.deepEqual(templated.groupSettingTemplates, []);
      assert.deepEqual(templated.count, { returned: 0, complete: true });
      assert.ok(templated.help.some(hint => hint.includes("0 setting templates matched")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown lifecycle and template ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "show", "--id", "policy-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "group-setting-template", "show", "--id", "template-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied group lifecycle reads surface scope, role and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Directory.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose role or licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "group-setting-template", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("GroupSettings.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Directory Readers or Global Reader are the least-privileged roles")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Directory.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "group-setting-template", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("GroupSettings.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts group lifecycle reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runLifecycleCli(["entra", "group-lifecycle-policy", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.groupLifecyclePolicies.map(row => row.id), [policy1.id, policy2.id]);
      assert.deepEqual(listOut.count, { returned: 2, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runLifecycleCli(["entra", "group-lifecycle-policy", "show", "--id", policy1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).groupLifecyclePolicy, policy1);

      const counted = runLifecycleCli(["entra", "group-lifecycle-policy", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).groupLifecyclePolicyCount, 2);
      assert.ok(!counted.stdout.includes(`opaque-fixture-${mode}-token`));

      const templated = runLifecycleCli(["entra", "group-setting-template", "list", "--profile", profile], state, mode);
      assert.equal(templated.status, 0, templated.stdout);
      const templateOut = decode(templated.stdout);
      assert.deepEqual(templateOut.groupSettingTemplates.map(row => row.id), [template1.id, template2.id]);
      assert.deepEqual(templateOut.count, { returned: 2, complete: true });

      const templateShown = runLifecycleCli(["entra", "group-setting-template", "show", "--id", template1.id, "--profile", profile], state, mode);
      assert.equal(templateShown.status, 0, templateShown.stdout);
      assert.deepEqual(decode(templateShown.stdout).groupSettingTemplate, template1);

      const templateCounted = runLifecycleCli(["entra", "group-setting-template", "count", "--profile", profile], state, mode);
      assert.equal(templateCounted.status, 0, templateCounted.stdout);
      assert.deepEqual(decode(templateCounted.stdout).groupSettingTemplateCount, 2);
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied group lifecycle reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runLifecycleCli(["entra", "group-lifecycle-policy", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.groupLifecyclePolicies, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["group-lifecycle-policy", "list"], ["--limit", "1"]],
    [["group-lifecycle-policy", "show"], ["--id", policy1.id]],
    [["group-lifecycle-policy", "count"], []],
    [["group-setting-template", "list"], ["--limit", "1"]],
    [["group-setting-template", "show"], ["--id", template1.id]],
    [["group-setting-template", "count"], []],
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
            { code: "VALIDATION_ERROR", message: "Group lifecycle policy and setting-template reads support v1.0 only; beta needs its own review" },
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

test("truncated lifecycle text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const partial = await executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", "soc",
      "--select", "id,alternateNotificationEmails"], overrides);
    const truncated = partial.groupLifecyclePolicies.find(row => row.id === policy2.id);
    assert.match(truncated.alternateNotificationEmails, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", "soc", "--full",
      "--select", "id,alternateNotificationEmails"], overrides);
    assert.equal(full.groupLifecyclePolicies.find(row => row.id === policy2.id).alternateNotificationEmails, longEmails);
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
      executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", "batch", "--scopes", lifecycleScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "group-setting-template", "count", "--profile", "batch", "--scopes", templateScopes[0]], overrides),
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
    await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", "soc", "--select", "id,owner"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "group-setting-template", "list", "--profile", "soc", "--fields", "values"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "group-setting-template", "show", "--id", template1.id, "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("malformed group lifecycle count bodies fail as unknown, not zero", async () => {
  const state = setupProfiles();
  try {
    for (const body of ["{}", "-1", "2.5", "\"3\""]) {
      const malformed = transport(() => ({ status: 200, headers: {}, body }));
      const scoped = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "count", "--profile", "soc"], scoped.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /malformed lifecycle-policy count|non-JSON success body/.test(error.message);
      });
      const templated = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "group-setting-template", "count", "--profile", "soc"], templated.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /malformed setting-template count|non-JSON success body/.test(error.message);
      });
    }
  } finally {
    teardownProfiles(state);
  }
});

test("group lifecycle read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "group-setting-template", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "count", "--limit", "5"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "count", "--filter", "managedGroupTypes eq 'All'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "group-setting-template", "count", "--select", "id"]), /unknown flag --select/);
  await assert.rejects(executeArgv(["entra", "group-setting-template", "count", "--cursor", "x"]), /unknown flag --cursor/);
  await assert.rejects(executeArgv(["entra", "group-setting-template", "list", "--filter", "displayName eq 'x'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "group-lifecycle-policy", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
});
