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
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, customers: [cu1, cu2], relationships: [rel1, rel2], denied }),
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
    await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "show", "--id", rel1.id, "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
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
  await assert.rejects(executeArgv(["entra", "delegated-admin-customer", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "delegated-admin-relationship", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
});
