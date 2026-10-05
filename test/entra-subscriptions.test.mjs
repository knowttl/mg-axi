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
const subscriptionScopes = ["https://graph.microsoft.com/Organization.Read.All"];

const longSkuPartNumber = `CONTOSO_ENTERPRISE_SUBSCRIPTION_SKU${"_with_considerable_detail".repeat(30)}`;
const sub1 = {
  id: "860697e3-b0aa-4196-a6c6-7ec361ed58f7",
  commerceSubscriptionId: "f9c1ea2d-2c6e-4717-8c3b-7130812d70ba",
  createdDateTime: "2023-01-01T00:00:00Z",
  isTrial: false,
  nextLifecycleDateTime: "2023-02-01T00:00:00Z",
  ownerId: "fe04f19f-d924-42b7-9dee-edf4e3fab7f6",
  ownerTenantId: "331af819-4e0b-49f7-a6bf-14e1165ad3a0",
  ownerType: "User",
  serviceStatus: [
    {
      appliesTo: "User",
      provisioningStatus: "Success",
      servicePlanId: "8b8269e5-f841-416c-ab3a-f5dfb9737986",
      servicePlanName: "MyPlanName",
    },
  ],
  skuId: "0816ccb9-3785-4d19-bf78-6c53e2106509",
  skuPartNumber: "MyPartNumber",
  status: "Enabled",
  totalLicenses: 25,
};
const sub2 = {
  id: "9b1c2d3e-4f5a-4b6c-8d7e-8f9a0b1c2d3e",
  commerceSubscriptionId: "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d",
  createdDateTime: "2024-06-01T00:00:00Z",
  isTrial: true,
  nextLifecycleDateTime: null,
  ownerId: "aa04f19f-d924-42b7-9dee-edf4e3fab7f6",
  ownerTenantId: null,
  ownerType: "User",
  serviceStatus: [],
  skuId: "1916ccb9-3785-4d19-bf78-6c53e2106509",
  skuPartNumber: longSkuPartNumber,
  status: "Suspended",
  totalLicenses: 10,
};
const sub3 = { id: "sub-bare-3" };
const subscriptions = [sub1, sub2, sub3];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-subscriptions-"));
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

function subscriptionTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/directory/subscriptions/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(subscriptions.length) };
    }
    if (path === "/v1.0/directory/subscriptions") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [sub3] });
      return json(200, {
        value: [sub1, sub2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/directory/subscriptions?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/directory\/subscriptions\/([^/]+)$/.exec(path);
    if (single) {
      const found = subscriptions.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such subscription" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? subscriptionTransport();
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

function runSubscriptionCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-subscriptions-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, subscriptions: [sub1, sub2], denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists subscriptions with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "subscription", "list", "--profile", profile,
        "--select", "id,skuPartNumber,status,totalLicenses"], overrides);
      assert.deepEqual(result.subscriptions, [
        { id: sub1.id, skuPartNumber: "MyPartNumber", status: "Enabled", totalLicenses: 25 },
        { id: sub2.id, skuPartNumber: `${longSkuPartNumber.slice(0, 500)}... (truncated, ${longSkuPartNumber.length} chars total)`, status: "Suspended", totalLicenses: 10 },
        { id: sub3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra subscription show --id <subscription-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/directory/subscriptions?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(subscriptionScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists subscriptions with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "subscription", "list", "--profile", profile,
        "--filter", "status eq 'Enabled'"], overrides);
      assert.equal(result.count.returned, 3);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "status eq 'Enabled'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped subscription list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "subscription", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.subscriptions.map(row => row.id), [sub1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "subscription", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.subscriptions.map(row => row.id), [sub2.id, sub3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one subscription with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "subscription", "show", "--id", sub1.id, "--profile", profile], overrides);
      assert.deepEqual(result.subscription, sub1);
      const missing = await executeArgv(["entra", "subscription", "show", "--id", sub3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.subscription, { id: sub3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  for (const id of ["(commerceSubscriptionId='9f8e7d6c-5b4a-4f3e-8d2c-1b0a9f8e7d6c')", "sub(commerceSubscriptionId='x')", "sub)(", "sub%28"]) {
    test(`${mode} subscription show rejects function-style identifier ${id} before credentials`, async () => {
      const state = setupProfiles();
      try {
        const { requests, calls, overrides } = overridesFor(mode);
        await assert.rejects(
          executeArgv(["entra", "subscription", "show", "--id", id, "--profile", profile], overrides),
          { code: "VALIDATION_ERROR" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }

  test(`${mode} counts subscriptions as one scalar without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "subscription", "count", "--profile", profile], overrides);
      assert.deepEqual(result, { count: { returned: 3, complete: true } });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/directory/subscriptions/$count");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} subscription count refuses collection flags before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "subscription", "count", "--profile", profile, "--filter", "status eq 'Enabled'"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty subscription lists stay definitive`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "subscription", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.subscriptions, []);
      assert.deepEqual(listed.count, { returned: 0, complete: true });
      assert.ok(listed.help.some(hint => hint.includes("no commercial subscriptions")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} filtered empty subscription lists report no match, not an empty tenant`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "subscription", "list", "--profile", profile,
        "--filter", "status eq 'Suspended'"], overrides);
      assert.deepEqual(listed.help, ["0 subscriptions matched; the absence of results is the answer, not an error"]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown subscription ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "subscription", "show", "--id", "subscription-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied subscription reads surface scope, role and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "subscription", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Organization.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Global Reader, Directory Readers")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "subscription", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Organization.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts subscription reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runSubscriptionCli(["entra", "subscription", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.subscriptions.map(row => row.id), [sub1.id, sub2.id]);
      assert.deepEqual(listOut.count, { returned: 2, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runSubscriptionCli(["entra", "subscription", "show", "--id", sub1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).subscription, sub1);

      const counted = runSubscriptionCli(["entra", "subscription", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).count, { returned: 2, complete: true });
      assert.ok(!counted.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied subscription reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runSubscriptionCli(["entra", "subscription", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.subscriptions, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["subscription", "list"], ["--limit", "1"]],
    [["subscription", "show"], ["--id", sub1.id]],
    [["subscription", "count"], []],
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
            { code: "VALIDATION_ERROR", message: "Subscription reads support v1.0 only; beta needs its own review" },
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

test("truncated subscription text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const partial = await executeArgv(["entra", "subscription", "list", "--profile", "soc",
      "--select", "id,skuPartNumber"], overrides);
    const truncated = partial.subscriptions.find(row => row.id === sub2.id);
    assert.match(truncated.skuPartNumber, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "subscription", "list", "--profile", "soc", "--full",
      "--select", "id,skuPartNumber"], overrides);
    assert.equal(full.subscriptions.find(row => row.id === sub2.id).skuPartNumber, longSkuPartNumber);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});
