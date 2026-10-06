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
const mtoScopes = ["https://graph.microsoft.com/MultiTenantOrganization.Read.All"];

const longDescription = `Alliance of Contoso, Fabrikam and Woodgrove${" collaborating closely".repeat(40)}`;
const org = {
  id: "6d8b39e5-039a-4034-bf3a-e0b4f8cd60b6",
  createdDateTime: "2023-05-26T22:05:23Z",
  displayName: "Contoso organization",
  description: "Multitenant organization between Contoso, Fabrikam, and Woodgrove Bank",
  state: "active",
};
const inactiveOrg = {
  createdDateTime: null,
  displayName: null,
  description: null,
  state: "inactive",
};
const joinRequest = {
  id: "3e536776-7489-43e9-9637-742d62ec5",
  addedByTenantId: "1fd6544e-e994-4de2-9f1b-787b51c7d325",
  memberState: "pending",
  role: null,
  transitionDetails: { desiredMemberState: "active", status: "notStarted", details: "" },
};
const t1 = {
  tenantId: "1fd6544e-e994-4de2-9f1b-787b51c7d325",
  displayName: "Contoso",
  addedDateTime: "2023-05-26T22:05:23Z",
  joinedDateTime: null,
  addedByTenantId: "1fd6544e-e994-4de2-9f1b-787b51c7d325",
  role: "owner",
  state: "active",
  transitionDetails: null,
};
const t2 = {
  tenantId: "4a12efe6-aa14-4d03-8dff-88fc89e2e2ad",
  displayName: longDescription,
  addedDateTime: "2023-05-27T19:24:32Z",
  joinedDateTime: null,
  addedByTenantId: "1fd6544e-e994-4de2-9f1b-787b51c7d325",
  role: "member",
  state: "pending",
  transitionDetails: null,
};
const t3 = { tenantId: "5036a0a0-a7a4-4933-9086-5dd54535dd6e" };
const tenants = [t1, t2, t3];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-mto-"));
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

function mtoTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/tenantRelationships/multiTenantOrganization/tenants/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(tenants.length) };
    }
    if (path === "/v1.0/tenantRelationships/multiTenantOrganization/tenants") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [t3] });
      return json(200, {
        value: [t1, t2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/tenantRelationships/multiTenantOrganization/joinRequest") return json(200, joinRequest);
    if (path === "/v1.0/tenantRelationships/multiTenantOrganization") return json(200, org);
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? mtoTransport();
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

function runMtoCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-mto-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, organization: org, joinRequest, tenants: [t1, t2], denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists multi-tenant-organization tenants with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", profile,
        "--select", "tenantId,displayName,role,state"], overrides);
      assert.deepEqual(result.multiTenantOrganizationTenants, [
        { tenantId: t1.tenantId, displayName: "Contoso", role: "owner", state: "active" },
        { tenantId: t2.tenantId, displayName: longDescription.slice(0, 500) + `... (truncated, ${longDescription.length} chars total)`, role: "member", state: "pending" },
        { tenantId: t3.tenantId },
      ]);
      assert.deepEqual(result.count, "3 tenants");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("entra multi-tenant-organization tenant count")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(mtoScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists multi-tenant-organization tenants with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", profile,
        "--filter", "state eq 'active'"], overrides);
      assert.equal(result.multiTenantOrganizationTenants.length, 3);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "state eq 'active'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped multi-tenant-organization tenant list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.multiTenantOrganizationTenants.map(row => row.tenantId), [t1.tenantId]);
      assert.equal(first.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.multiTenantOrganizationTenants.map(row => row.tenantId), [t2.tenantId, t3.tenantId]);
      assert.deepEqual(second.count, "2 tenants");
      assert.equal(second.total, null);
      assert.equal(second.complete, true);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty tenant lists stay definitive with multitenant context`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const result = await executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", profile], overrides);
      assert.deepEqual(result.multiTenantOrganizationTenants, []);
      assert.deepEqual(result.count, "0 tenants");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("0 tenants matched")));
      assert.ok(result.help.some(hint => hint.includes("outside any multitenant organization")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows the multitenant organization with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "multi-tenant-organization", "show", "--profile", profile,
        "--select", "createdDateTime,description,displayName,id,state"], overrides);
      assert.deepEqual(result.multiTenantOrganization, org);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} inactive organizations read as absence, not error`, async () => {
    const state = setupProfiles();
    try {
      const inactive = transport(request => {
        assert.ok(new URL(request.url).pathname === "/v1.0/tenantRelationships/multiTenantOrganization");
        return json(200, inactiveOrg);
      });
      const { overrides } = overridesFor(mode, inactive);
      const result = await executeArgv(["entra", "multi-tenant-organization", "show", "--profile", profile], overrides);
      assert.deepEqual(result.multiTenantOrganization, inactiveOrg);
      assert.ok(result.help.some(hint => hint.includes("not part of any multitenant organization")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows the join request with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "multi-tenant-organization", "join-request", "show", "--profile", profile,
        "--select", "addedByTenantId,id,memberState,role,transitionDetails"], overrides);
      assert.deepEqual(result.multiTenantOrganizationJoinRequest, joinRequest);
      assert.equal(result.help, undefined);
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/joinRequest?"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts tenants as one scalar without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "multi-tenant-organization", "tenant", "count", "--profile", profile], overrides);
      assert.deepEqual(result, { multiTenantOrganizationTenantCount: 3 });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants/$count");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} malformed tenant counts stay unknown, not zero`, async () => {
    const state = setupProfiles();
    try {
      const malformed = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, malformed);
      await assert.rejects(executeArgv(["entra", "multi-tenant-organization", "tenant", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /malformed multi-tenant-organization tenant count body/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied multi-tenant-organization reads surface scope, roles and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("MultiTenantOrganization.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Security Reader or Global Reader")));
        assert.ok(error.suggestions.some(hint => hint.includes("Global service")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "multi-tenant-organization", "show", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("MultiTenantOrganization.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "multi-tenant-organization", "tenant", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("MultiTenantOrganization.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts multi-tenant-organization reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runMtoCli(["entra", "multi-tenant-organization", "tenant", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.multiTenantOrganizationTenants.map(row => row.tenantId), [t1.tenantId, t2.tenantId]);
      assert.deepEqual(listOut.count, "2 tenants");
      assert.equal(listOut.total, null);
      assert.equal(listOut.complete, true);
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runMtoCli(["entra", "multi-tenant-organization", "show", "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).multiTenantOrganization, org);

      const joinShown = runMtoCli(["entra", "multi-tenant-organization", "join-request", "show", "--profile", profile], state, mode);
      assert.equal(joinShown.status, 0, joinShown.stdout);
      assert.deepEqual(decode(joinShown.stdout).multiTenantOrganizationJoinRequest, joinRequest);

      const counted = runMtoCli(["entra", "multi-tenant-organization", "tenant", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).multiTenantOrganizationTenantCount, 2);
      assert.ok(!counted.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied multi-tenant-organization reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runMtoCli(["entra", "multi-tenant-organization", "tenant", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.multiTenantOrganizationTenants, undefined);
      assert.ok(output.help.some(hint => hint.includes("Global service")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["multi-tenant-organization", "show"], []],
    [["multi-tenant-organization", "join-request", "show"], []],
    [["multi-tenant-organization", "tenant", "list"], ["--limit", "1"]],
    [["multi-tenant-organization", "tenant", "count"], []],
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
            { code: "VALIDATION_ERROR", message: "Multi-tenant-organization reads support v1.0 only; beta needs its own review" },
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

test("truncated tenant text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const partial = await executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", "soc",
      "--select", "tenantId,displayName"], overrides);
    const truncated = partial.multiTenantOrganizationTenants.find(row => row.tenantId === t2.tenantId);
    assert.match(truncated.displayName, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", "soc", "--full",
      "--select", "tenantId,displayName"], overrides);
    assert.equal(full.multiTenantOrganizationTenants.find(row => row.tenantId === t2.tenantId).displayName, longDescription);
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
      executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", "batch", "--scopes", mtoScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "multi-tenant-organization", "tenant", "count", "--profile", "batch", "--scopes", mtoScopes[0]], overrides),
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
    await assert.rejects(executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", "soc", "--select", "tenantId,owner"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", "soc", "--fields", "addedDateTime"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "multi-tenant-organization", "show", "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "multi-tenant-organization", "tenant", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});
