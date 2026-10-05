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
const contractScopes = ["https://graph.microsoft.com/Directory.Read.All"];

const longName = `Fabrikam customer-${" very-important".repeat(40)}`;
const ct1 = {
  contractType: "ResellerPartner",
  customerId: "aaaaaaaa-1111-4111-8111-111111111111",
  defaultDomainName: "fabrikam.com",
  displayName: "Fabrikam",
  id: "ct-reseller-1",
};
const ct2 = {
  contractType: "BreadthPartner",
  customerId: "bbbbbbbb-2222-4222-8222-222222222222",
  defaultDomainName: "contoso.com",
  displayName: longName,
  id: "ct-breadth-2",
};
const ct3 = { id: "ct-bare-3" };
const contracts = [ct1, ct2, ct3];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-contracts-"));
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

function contractTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/contracts/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(contracts.length) };
    }
    if (path === "/v1.0/contracts") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [ct3] });
      return json(200, {
        value: [ct1, ct2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/contracts?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/contracts\/([^/]+)$/.exec(path);
    if (single) {
      const found = contracts.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such contract" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? contractTransport();
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

function runContractCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-contracts-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, contracts: [ct1, ct2], denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists contracts with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "contract", "list", "--profile", profile,
        "--select", "id,displayName,contractType,defaultDomainName"], overrides);
      assert.deepEqual(result.contracts, [
        { id: ct1.id, displayName: "Fabrikam", contractType: "ResellerPartner", defaultDomainName: "fabrikam.com" },
        { id: ct2.id, displayName: longName.slice(0, 500) + `... (truncated, ${longName.length} chars total)`, contractType: "BreadthPartner", defaultDomainName: "contoso.com" },
        { id: ct3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra contract show --id <contract-id>")));
      assert.ok(result.help.some(hint => hint.includes("partner tenants only")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/contracts?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(contractScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists contracts with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "contract", "list", "--profile", profile,
        "--filter", "contractType eq 'ResellerPartner'"], overrides);
      assert.equal(result.count.returned, 3);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "contractType eq 'ResellerPartner'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped contract list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "contract", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.contracts.map(row => row.id), [ct1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "contract", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.contracts.map(row => row.id), [ct2.id, ct3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one contract with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "contract", "show", "--id", ct1.id, "--profile", profile,
        "--select", "contractType,customerId,defaultDomainName,displayName,id"], overrides);
      assert.deepEqual(result.contract, ct1);
      assert.equal(result.help, undefined);
      const missing = await executeArgv(["entra", "contract", "show", "--id", ct3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.contract, { id: ct3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  for (const id of ["delta()", "DELTA()", "delta(", "delta)", "delta%28%29"]) {
    test(`${mode} contract show rejects function-style identifier ${id} before credentials`, async () => {
      const state = setupProfiles();
      try {
        const delta = transport(() => json(200, { value: [ct1], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/contracts/delta()" }));
        const { requests, calls, overrides } = overridesFor(mode, delta);
        await assert.rejects(
          executeArgv(["entra", "contract", "show", "--id", id, "--profile", profile], overrides),
          { code: "VALIDATION_ERROR" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }

  test(`${mode} counts contracts as one scalar without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "contract", "count", "--profile", profile], overrides);
      assert.deepEqual(result, { contractCount: 3 });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/contracts/$count");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty contract lists stay definitive with partner context`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const result = await executeArgv(["entra", "contract", "list", "--profile", profile], overrides);
      assert.deepEqual(result.contracts, []);
      assert.deepEqual(result.count, { returned: 0, complete: true });
      assert.ok(result.help.some(hint => hint.includes("0 contracts matched")));
      assert.ok(result.help.some(hint => hint.includes("partner tenants only")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown contract ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "contract", "show", "--id", "ct-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied contract reads surface scope, roles, partner and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "contract", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Directory.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Directory Readers is the least-privileged role")));
        assert.ok(error.suggestions.some(hint => hint.includes("partner tenants only")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "contract", "show", "--id", ct1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("partner tenants only")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "contract", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Directory.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts contracts`, () => {
    const state = setupProfiles();
    try {
      const listed = runContractCli(["entra", "contract", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.contracts.map(row => row.id), [ct1.id, ct2.id]);
      assert.deepEqual(listOut.count, { returned: 2, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runContractCli(["entra", "contract", "show", "--id", ct1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).contract, ct1);

      const counted = runContractCli(["entra", "contract", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).contractCount, 2);
      assert.ok(!counted.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied contract reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runContractCli(["entra", "contract", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.contracts, undefined);
      assert.ok(output.help.some(hint => hint.includes("partner tenants only")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["contract", "list"], ["--limit", "1"]],
    [["contract", "show"], ["--id", ct1.id]],
    [["contract", "count"], []],
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
            { code: "VALIDATION_ERROR", message: "Contract reads support v1.0 only; beta needs its own review" },
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

test("truncated contract text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const partial = await executeArgv(["entra", "contract", "list", "--profile", "soc",
      "--select", "id,displayName"], overrides);
    const truncated = partial.contracts.find(row => row.id === ct2.id);
    assert.match(truncated.displayName, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "contract", "list", "--profile", "soc", "--full",
      "--select", "id,displayName"], overrides);
    assert.equal(full.contracts.find(row => row.id === ct2.id).displayName, longName);
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
      executeArgv(["entra", "contract", "list", "--profile", "batch", "--scopes", contractScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "contract", "count", "--profile", "batch", "--scopes", contractScopes[0]], overrides),
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
    await assert.rejects(executeArgv(["entra", "contract", "list", "--profile", "soc", "--select", "id,owner"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "contract", "list", "--profile", "soc", "--fields", "customerId"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "contract", "show", "--id", ct1.id, "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "contract", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("malformed contract count bodies fail as unknown, not zero", async () => {
  const state = setupProfiles();
  try {
    for (const body of ["{}", "-1", "2.5", "\"3\""]) {
      const malformed = transport(() => ({ status: 200, headers: {}, body }));
      const scoped = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "contract", "count", "--profile", "soc"], scoped.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /malformed contract count|non-JSON success body/.test(error.message);
      });
    }
  } finally {
    teardownProfiles(state);
  }
});

test("contract read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "contract", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "contract", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "contract", "count", "--limit", "5"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "contract", "count", "--filter", "displayName eq 'x'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "contract", "count", "--select", "id"]), /unknown flag --select/);
  await assert.rejects(executeArgv(["entra", "contract", "count", "--cursor", "x"]), /unknown flag --cursor/);
  await assert.rejects(executeArgv(["entra", "contract", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
});
