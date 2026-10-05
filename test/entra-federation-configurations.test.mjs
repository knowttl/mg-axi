import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import { ApplicationAuth } from "../dist/app-auth.js";
import { DelegatedAuth } from "../dist/auth.js";
import { runApiGet } from "../dist/api.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const federationScopes = ["https://graph.microsoft.com/Domain.Read.All"];
const typesScopes = ["https://graph.microsoft.com/IdentityProvider.Read.All"];
const certBlob = "MIIC6DCCAdCgAwIBAgIQQ6vYJIVKQ".repeat(40);
const unsupportedFederationFields = ["domains", "issuer", "thumbprint", "activeSignInUri", "signOutUri"];

const contoso = {
  "@odata.type": "#microsoft.graph.samlOrWsFedExternalDomainFederation",
  id: "96db02e2-80c1-5555-bc3a-de92ffb8c5be",
  displayName: "Contoso",
  issuerUri: "http://contoso.com/adfs/services/trust",
  metadataExchangeUri: null,
  signingCertificate: certBlob,
  passiveSignInUri: "https://contoso.com/adfs/ls/",
  preferredAuthenticationProtocol: "saml",
};
const fabrikam = {
  "@odata.type": "#microsoft.graph.samlOrWsFedExternalDomainFederation",
  id: "fa421032-5d40-5555-a428-a304b4bc18b6",
  displayName: "Fabrikam",
  issuerUri: "https://fabrikam.com/o/saml2?idpid=C018555d",
  metadataExchangeUri: null,
  passiveSignInUri: "https://fabrikam.com/o/saml2/saml2?idpid=C018555d",
  preferredAuthenticationProtocol: "wsFed",
};
const builtin = { id: "builtin-config" };
const configurations = [contoso, fabrikam, builtin];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-federation-"));
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

function federationTransport({ denied = false, countBody = "2", countStatus = 200, typesBody = { value: ["saml", "wsFed"] } } = {}) {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (denied) return json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } });
    if (path === "/v1.0/directory/federationConfigurations/availableProviderTypes()") return json(200, typesBody);
    if (path === "/v1.0/directory/federationConfigurations/$count") {
      return { status: countStatus, headers: {}, body: countBody };
    }
    if (path === "/v1.0/directory/federationConfigurations") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [fabrikam, builtin] });
      return json(200, {
        value: [contoso],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/directory/federationConfigurations?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/directory\/federationConfigurations\/([^/]+)$/.exec(path);
    if (single) {
      const found = configurations.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such configuration" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, fixture, calls = []) {
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

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists federation configurations with compact rows, kinds and a count aggregate`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode, federationTransport());
      const result = await executeArgv(["entra", "federation-configuration", "list", "--profile", profile], overrides);
      assert.deepEqual(result.federationConfigurations, [
        { id: contoso.id, displayName: "Contoso", "@odata.type": "#microsoft.graph.samlOrWsFedExternalDomainFederation" },
        { id: fabrikam.id, displayName: "Fabrikam", "@odata.type": "#microsoft.graph.samlOrWsFedExternalDomainFederation" },
        { id: builtin.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra federation-configuration show --id <configuration-id>")));
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      assert.ok(result.help.some(hint => hint.includes("--select signingCertificate")));
      assert.ok(!JSON.stringify(result).includes(certBlob.slice(0, 32)));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/directory/federationConfigurations?"));
      if (mode === "delegated") assert.deepEqual(calls[0][1], federationScopes);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists federation configurations with an explicit signing-certificate select`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, federationTransport());
      const result = await executeArgv(
        ["entra", "federation-configuration", "list", "--profile", profile, "--select", "id,signingCertificate", "--full"], overrides);
      assert.equal(result.federationConfigurations[0].signingCertificate, certBlob);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} reads available federation provider types through the v1.0 function`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode, federationTransport());
      const result = await executeArgv(["entra", "federation-configuration", "available-types", "--profile", profile], overrides);
      assert.deepEqual(result.availableProviderTypes, ["saml", "wsFed"]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/directory/federationConfigurations/availableProviderTypes()");
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      assert.ok(result.help.some(hint => hint.includes("licensing")));
      assert.equal(calls.length, 1);
      if (mode === "delegated") assert.deepEqual(calls[0][1], typesScopes);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} accepts an empty available federation provider types result`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, federationTransport({ typesBody: { value: [] } }));
      const result = await executeArgv(["entra", "federation-configuration", "available-types", "--profile", profile], overrides);
      assert.deepEqual(result.availableProviderTypes, []);
      assert.deepEqual(result.count, { returned: 0, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped federation-configuration list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, federationTransport());
      const first = await executeArgv(["entra", "federation-configuration", "list", "--profile", profile, "--limit", "2"], overrides);
      assert.equal(first.federationConfigurations.length, 2);
      assert.deepEqual(first.count, { returned: 2, complete: false, reason: "row limit reached; buffered remainder is preserved in the cursor" });
      assert.equal(typeof first.cursor, "string");
      const { overrides: resumeOverrides } = overridesFor(mode, federationTransport());
      const second = await executeArgv(["entra", "federation-configuration", "list", "--profile", profile, "--cursor", first.cursor], resumeOverrides);
      assert.equal(second.federationConfigurations.length, 1);
      assert.deepEqual(second.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one federation configuration without the certificate blob by default`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, federationTransport());
      const result = await executeArgv(["entra", "federation-configuration", "show", "--id", contoso.id, "--profile", profile], overrides);
      assert.deepEqual(result.federationConfiguration, {
        id: contoso.id,
        displayName: "Contoso",
        issuerUri: "http://contoso.com/adfs/services/trust",
        metadataExchangeUri: null,
        passiveSignInUri: "https://contoso.com/adfs/ls/",
        preferredAuthenticationProtocol: "saml",
        "@odata.type": "#microsoft.graph.samlOrWsFedExternalDomainFederation",
      });
      assert.ok(!("signingCertificate" in result.federationConfiguration));
      assert.ok(result.help.some(hint => hint.includes("--select signingCertificate")));
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      const blob = await executeArgv(
        ["entra", "federation-configuration", "show", "--id", contoso.id, "--profile", profile, "--select", "id,signingCertificate", "--full"], overrides);
      assert.equal(blob.federationConfiguration.signingCertificate, certBlob);
      const wsFed = await executeArgv(["entra", "federation-configuration", "show", "--id", fabrikam.id, "--profile", profile], overrides);
      assert.equal(wsFed.federationConfiguration.preferredAuthenticationProtocol, "wsFed");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts federation configurations as a scalar number`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode, federationTransport());
      const result = await executeArgv(["entra", "federation-configuration", "count", "--profile", profile], overrides);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/directory/federationConfigurations/$count"));
      const filtered = await executeArgv(
        ["entra", "federation-configuration", "count", "--profile", profile, "--filter", "displayName eq 'Contoso'"], overrides);
      assert.equal(filtered.count.returned, 2);
      assert.ok(new URL(requests[requests.length - 1].url).searchParams.has("$filter"));
      assert.ok(filtered.help.some(hint => hint.includes("drop --filter for the tenant total")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied federation reads surface scope, role and licence guidance`, async () => {
    const state = setupProfiles();
    try {
      const { calls, requests, overrides } = overridesFor(mode, federationTransport({ denied: true }));
      for (const args of [
        ["entra", "federation-configuration", "list", "--profile", profile],
        ["entra", "federation-configuration", "show", "--id", contoso.id, "--profile", profile],
        ["entra", "federation-configuration", "count", "--profile", profile],
      ]) {
        await assert.rejects(executeArgv(args, overrides), error => {
          assert.equal(error.code, "GRAPH_ERROR");
          assert.match(error.message, /\(403\)/);
          assert.match(error.message, /grant, role, licence/);
          assert.ok(error.suggestions.some(hint => hint.includes("Domain.Read.All")));
          assert.ok(error.suggestions.some(hint => hint.includes("External Identity Provider Administrator")));
          assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
          assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
          return true;
        });
      }
      await assert.rejects(executeArgv(["entra", "federation-configuration", "available-types", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("IdentityProvider.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("External Identity Provider Administrator")));
        return true;
      });
      assert.ok(calls.length > 0);
      assert.ok(requests.length > 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown federation-configuration id reports not found, never empty`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, federationTransport());
      await assert.rejects(
        executeArgv(["entra", "federation-configuration", "show", "--id", "no-such-configuration", "--profile", profile], overrides),
        error => error.code === "GRAPH_ERROR" && /not found or inaccessible/.test(error.message),
      );
    } finally {
      teardownProfiles(state);
    }
  });
}

test("truncated federation text carries a --full hint without lifting row caps", async () => {
  const state = setupProfiles();
  try {
    const displayName = "x".repeat(600);
    const { overrides } = overridesFor("delegated", transport(() => json(200, { ...contoso, displayName })));
    const result = await executeArgv(["entra", "federation-configuration", "show", "--id", contoso.id, "--profile", "soc"], overrides);
    assert.match(result.federationConfiguration.displayName, /\.\.\. \(truncated, \d+ chars total\)/);
    assert.ok(result.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "federation-configuration", "show", "--id", contoso.id, "--profile", "soc", "--full"], overrides);
    assert.equal(full.federationConfiguration.displayName, displayName);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("a non-numeric federation count body is malformed, never a count", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated", federationTransport({ countBody: JSON.stringify({ value: [] }) }));
    await assert.rejects(
      executeArgv(["entra", "federation-configuration", "count", "--profile", "soc"], overrides),
      error => error.code === "GRAPH_ERROR" && /non-numeric success body/.test(error.message),
    );
  } finally {
    teardownProfiles(state);
  }
});

test("unknown federation properties and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { calls, requests, overrides } = overridesFor("delegated", federationTransport());
    await assert.rejects(
      executeArgv(["entra", "federation-configuration", "list", "--profile", "soc", "--select", "id,bogus"], overrides),
      /Unknown federation-configuration property bogus/,
    );
    await assert.rejects(
      executeArgv(["entra", "federation-configuration", "list", "--profile", "soc", "--fields", "issuerUri"], overrides),
      /was not fetched; request it with --select/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("federation-configuration read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "federation-configuration", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "federation-configuration", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "federation-configuration", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "federation-configuration", "count", "--select", "id"]), /unknown flag --select/);
  await assert.rejects(executeArgv(["entra", "federation-configuration", "count", "--limit", "10"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "federation-configuration", "list", "--limit", "10", "--all"]), { code: "VALIDATION_ERROR" });
});

for (const field of unsupportedFederationFields) {
  for (const args of [
    ["list", "--select"], ["list", "--fields"],
    ["show", "--id", contoso.id, "--select"], ["show", "--id", contoso.id, "--fields"],
  ]) {
    test(`named federation-configuration ${args.join(" ")} refuses ${field} before credentials`, async () => {
      const state = setupProfiles();
      try {
        const { calls, requests, overrides } = overridesFor("delegated", federationTransport());
        await assert.rejects(
          executeArgv(["entra", "federation-configuration", ...args, field, "--profile", "soc"], overrides),
          error => error.code === "VALIDATION_ERROR" && error.message.includes(`Unknown federation-configuration property ${field}`),
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }
}

for (const body of [null, [], {}, { value: null }, { value: ["saml", 1] }, { value: [{ displayName: "saml" }] }]) {
  test(`available federation provider types rejects malformed body ${JSON.stringify(body)}`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor("delegated", federationTransport({ typesBody: body }));
      await assert.rejects(
        executeArgv(["entra", "federation-configuration", "available-types", "--profile", "soc"], overrides),
        { code: "GRAPH_ERROR", message: "Graph returned a malformed available federation provider types body" },
      );
    } finally {
      teardownProfiles(state);
    }
  });
}

for (const id of ["$count", "$value", "$ref"]) {
  test(`federation-configuration show rejects reserved resource binding ${id} before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { calls, requests, overrides } = overridesFor("delegated", federationTransport());
      await assert.rejects(
        executeArgv(["entra", "federation-configuration", "show", "--id", id, "--profile", "soc"], overrides),
        { code: "VALIDATION_ERROR", message: "OData reserved segments cannot be resource identifiers" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  for (const args of [
    ["entra", "federation-configuration", "list", "--profile", profile],
    ["entra", "federation-configuration", "show", "--id", contoso.id, "--profile", profile],
    ["entra", "federation-configuration", "count", "--profile", profile],
    ["entra", "federation-configuration", "available-types", "--profile", profile],
  ]) {
    test(`${mode} ${args[1]} ${args[2]} refuses beta before credentials`, async () => {
      const state = setupProfiles();
      try {
        const { calls, requests, overrides } = overridesFor(mode, federationTransport());
        await assert.rejects(
          executeArgv([...args, "--api-version", "beta"], overrides),
          { code: "VALIDATION_ERROR", message: "Federation-configuration reads support v1.0 only; beta needs its own review" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }
}

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { calls, requests, overrides } = overridesFor("application", federationTransport());
    await assert.rejects(
      executeArgv(["entra", "federation-configuration", "list", "--profile", "batch", "--scopes", federationScopes[0]], overrides),
      { code: "VALIDATION_ERROR", message: "Application profiles use the configured Graph .default audience; delegated scopes are unavailable" },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("available federation types application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { calls, requests, overrides } = overridesFor("application", federationTransport());
    await assert.rejects(
      executeArgv(["entra", "federation-configuration", "available-types", "--profile", "batch", "--scopes", typesScopes[0]], overrides),
      { code: "VALIDATION_ERROR", message: "Application profiles use the configured Graph .default audience; delegated scopes are unavailable" },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

function rawFixture(handler) {
  const credentialCalls = [];
  const requests = [];
  const credential = account => ({ token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, ...account });
  const delegated = new DelegatedAuth({
    storage: "session-only",
    login: async () => credential({ accountId: "synthetic-account" }),
    silent: async (...args) => { credentialCalls.push(["silent", ...args]); return credential({ accountId: "synthetic-account" }); },
  });
  const application = new ApplicationAuth({
    storage: "session-only",
    acquire: async (...args) => { credentialCalls.push(["acquire", ...args]); return credential({}); },
  });
  const transport = async request => {
    requests.push(request);
    const response = typeof handler === "function" ? await handler(request, requests.length) : handler;
    return { headers: {}, body: "", ...response };
  };
  return { credentialCalls, requests, deps: { delegated, application, transport } };
}

const delegatedRawProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key } };

function rawBody(request) {
  const url = new URL(request.url);
  if (url.pathname === "/v1.0/directory/federationConfigurations") return { status: 200, headers: {}, body: JSON.stringify({ value: configurations }) };
  if (url.pathname === `/v1.0/directory/federationConfigurations/${contoso.id}`) return { status: 200, headers: {}, body: JSON.stringify(contoso) };
  return { status: 404, headers: {}, body: JSON.stringify({ error: { code: "Unknown", message: "unexpected route" } }) };
}

test("raw api serves reviewed federation list and show with the certificate blob", async () => {
  const f = rawFixture(rawBody);
  const list = await runApiGet({ path: "/directory/federationConfigurations", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: federationScopes[0] }, f.deps);
  assert.equal(list.returned, 3);
  assert.equal(list.complete, true);
  const show = await runApiGet({ path: `/directory/federationConfigurations/${contoso.id}`, apiVersion: "v1.0", profile: delegatedRawProfile, scopes: federationScopes[0] }, f.deps);
  assert.equal(show.displayName, "Contoso");
  assert.ok(typeof show.signingCertificate === "string");
});

test("raw api refuses the federation $count scalar before credentials", async () => {
  const f = rawFixture(rawBody);
  await assert.rejects(
    runApiGet({ path: "/directory/federationConfigurations/$count", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: federationScopes[0] }, f.deps),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("raw api refuses unreviewed federation $select fields before credentials", async () => {
  const f = rawFixture(rawBody);
  await assert.rejects(
    runApiGet({ path: "/directory/federationConfigurations", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: federationScopes[0], odata: "$select=id,domains" }, f.deps),
    error => error.code === "VALIDATION_ERROR" && /Unreviewed \$select field domains/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const field of unsupportedFederationFields) {
  for (const path of ["/directory/federationConfigurations", `/directory/federationConfigurations/${contoso.id}`]) {
    test(`raw federation ${path} refuses ${field} before credentials`, async () => {
      const f = rawFixture(rawBody);
      await assert.rejects(
        runApiGet({ path, apiVersion: "v1.0", profile: delegatedRawProfile, scopes: federationScopes[0], odata: `$select=id,${field}` }, f.deps),
        error => error.code === "VALIDATION_ERROR" && error.message.includes(`Unreviewed $select field ${field}`),
      );
      assert.equal(f.credentialCalls.length, 0);
      assert.equal(f.requests.length, 0);
    });
  }
}
