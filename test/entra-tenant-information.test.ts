import { record, text, strings, authFakes, type CredentialStub, type CredentialArgs, type TransportFixture } from "./fixtures/typed.ts";
import type { GraphTransport, TransportRequest, TransportResponse } from "../dist/graph-session.js";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const tenant = "33333333-3333-4333-8333-333333333333";
const client = "44444444-4444-4222-8222-444444444444";
const tenantInformationScopes = ["https://graph.microsoft.com/CrossTenantInformation.ReadBasic.All"];
const lookupTenantId = "55555555-5555-4555-8555-555555555555";

const infoByDomain = {
  tenantId: lookupTenantId,
  defaultDomainName: "example.invalid",
  displayName: "Example",
  federationBrandName: null,
  unreviewed: "must-not-escape",
};
const infoByTenant = {
  tenantId: lookupTenantId,
  defaultDomainName: "example.invalid",
  displayName: "Example",
  federationBrandName: "Example brand",
};
const projectedByDomain = {
  tenantId: lookupTenantId,
  defaultDomainName: "example.invalid",
  displayName: "Example",
  federationBrandName: null,
};

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-tenant-information-"));
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(dir, "config.json");
  const profiles = new Profiles();
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  return { dir, previous };
}

function teardownProfiles(state: ReturnType<typeof setupProfiles>) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function credentialService(mode: string, calls: [string, ...CredentialArgs][]): CredentialStub {
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

function json(status: number, body: unknown, headers: Record<string, string> = {}): TransportResponse {
  return { status, headers, body: JSON.stringify(body) };
}

function tenantInformationTransport() {
  const requests: TransportRequest[] = [];
  const send: GraphTransport = async request => {
    requests.push(request);
    const url = new URL(request.url);
    if (url.pathname === "/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName=%27example.invalid%27)") return json(200, infoByDomain);
    if (url.pathname === `/v1.0/tenantRelationships/findTenantInformationByTenantId(tenantId=%27${lookupTenantId}%27)`) return json(200, infoByTenant);
    return json(404, { error: { code: "Request_ResourceNotFound", message: "no such tenant" } });
  };
  return { requests, send };
}

function overridesFor(mode: string, fixture?: TransportFixture, calls: [string, ...CredentialArgs][] = []) {
  const transport = fixture ?? tenantInformationTransport();
  const credential = credentialService(mode, calls);
  return {
    requests: transport.requests,
    calls,
    overrides: {
      transport: transport.send,
      delegated: authFakes(mode === "delegated" ? credential : credentialService("delegated", [])).delegated,
      application: authFakes(mode === "application" ? credential : credentialService("application", [])).application,
    },
  };
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]] as const) {
  test(`${mode} shows tenant information by domain name with the reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = record(await executeArgv(["entra", "tenant-information", "show", "--domain-name", "example.invalid", "--profile", profile], overrides));
      assert.deepEqual(result.tenantInformation, projectedByDomain);
      assert.equal(requests[0]!.url, "https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName=%27example.invalid%27)");
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(tenantInformationScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows tenant information by tenant ID with the reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = record(await executeArgv(["entra", "tenant-information", "show", "--tenant-id", lookupTenantId, "--profile", profile], overrides));
      assert.deepEqual(result.tenantInformation, infoByTenant);
      assert.equal(requests[0]!.url, `https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByTenantId(tenantId=%27${lookupTenantId}%27)`);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} tenant-information show needs exactly one lookup key`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "tenant-information", "show", "--profile", profile], overrides),
        error => record(error).code === "VALIDATION_ERROR" && /exactly one of --domain-name or --tenant-id/.test(text(record(error).message)),
      );
      await assert.rejects(
        executeArgv(["entra", "tenant-information", "show", "--domain-name", "example.invalid", "--tenant-id", lookupTenantId, "--profile", profile], overrides),
        error => record(error).code === "VALIDATION_ERROR" && /cannot be combined/.test(text(record(error).message)),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  for (const [flags, pattern] of [
    [["--domain-name", "x') OR '1'='1"], /Invalid domain name/],
    [["--domain-name", "a%2Fb"], /Invalid domain name/],
    [["--domain-name", "$batch"], /Invalid domain name/],
    [["--domain-name", "a\nb"], /Invalid domain name/],
    [["--tenant-id", "example.invalid"], /Invalid tenant ID/],
    [["--tenant-id", "x') OR '1'='1"], /Invalid tenant ID/],
    [["--tenant-id", "a%2Fb"], /Invalid tenant ID/],
    [["--tenant-id", "$batch"], /Invalid tenant ID/],
  ] as const) test(`${mode} tenant-information show refuses ${flags[0]} injection before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "tenant-information", "show", ...flags, "--profile", profile], overrides),
        error => record(error).code === "VALIDATION_ERROR" && pattern.test(text(record(error).message)),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown tenant lookups report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "tenant-information", "show", "--domain-name", "missing.invalid", "--profile", profile], overrides), error => {
        assert.equal(record(error).code, "GRAPH_ERROR");
        assert.deepEqual(record(error).suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(text(record(error).message));
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied tenant-information lookups surface scope and role guidance`, async () => {
    const state = setupProfiles();
    try {
      const denied = (() => {
        const requests: TransportRequest[] = [];
        return { requests, send: async (request: TransportRequest) => { requests.push(request); return json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }); } };
      })();
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "tenant-information", "show", "--domain-name", "example.invalid", "--profile", profile], overrides), error => {
        assert.equal(record(error).code, "GRAPH_ERROR");
        assert.ok(strings(record(error).suggestions).some(hint => hint.includes("CrossTenantInformation.ReadBasic.All")));
        assert.ok(strings(record(error).suggestions).some(hint => hint.includes("No Entra role is required")));
        assert.ok(strings(record(error).suggestions).some(hint => hint.includes("Personal Microsoft accounts are not supported")));
        assert.ok(strings(record(error).suggestions).some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(text(record(error).message));
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} tenant-information lookups refuse beta before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "tenant-information", "show", "--domain-name", "example.invalid", "--profile", profile, "--api-version", "beta"], overrides),
        error => record(error).code === "VALIDATION_ERROR" && /v1.0 only/.test(text(record(error).message)),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} tenant-information lookups refuse unreviewed query flags before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "tenant-information", "show", "--domain-name", "example.invalid", "--profile", profile, "--select", "id"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

test("application tenant-information lookups reject caller scopes before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "tenant-information", "show", "--domain-name", "example.invalid", "--profile", "batch", "--scopes", tenantInformationScopes.join(",")], overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});
