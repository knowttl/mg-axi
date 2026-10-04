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
const providerScopes = ["https://graph.microsoft.com/IdentityProvider.Read.All"];

const google = {
  "@odata.type": "#microsoft.graph.socialIdentityProvider",
  id: "Google-OAUTH",
  displayName: "Google",
  identityProviderType: "Google",
  clientId: "google-client-id",
  clientSecret: "must-never-surface",
};
const longMetadata = `https://login.contoso.example.com/metadata/${"x".repeat(600)}.xml`;
const saml = {
  "@odata.type": "#microsoft.graph.samlOrWsFedProvider",
  id: "saml-contoso",
  displayName: "Contoso SAML",
  issuerUri: "https://login.contoso.example.com/issuer",
  metadataExchangeUri: longMetadata,
  passiveSignInUri: null,
  preferredAuthenticationProtocol: "saml",
  activeSignInUri: "https://login.contoso.example.com/active",
  signOutUri: "https://login.contoso.example.com/signout",
};
const apple = {
  "@odata.type": "#microsoft.graph.appleManagedIdentityProvider",
  id: "Apple-Managed",
  displayName: "Apple",
  developerId: "DEV123",
  serviceId: "com.example.service",
  keyId: "KEY123",
  certificateData: "must-never-surface",
};
const builtin = { id: "builtin-aad" };
const providers = [google, saml, apple, builtin];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-providers-"));
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

function providerTransport({ denied = false, countBody = "4", countStatus = 200 } = {}) {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (denied) return json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } });
    if (path === "/v1.0/identity/identityProviders/$count") {
      return { status: countStatus, headers: {}, body: countBody };
    }
    if (path === "/v1.0/identity/identityProviders") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [apple, builtin] });
      return json(200, {
        value: [google, saml],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identity/identityProviders?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/identity\/identityProviders\/([^/]+)$/.exec(path);
    if (single) {
      const found = providers.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such provider" } });
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
  test(`${mode} lists providers with compact rows, kinds and a count aggregate`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode, providerTransport());
      const result = await executeArgv(["entra", "identity-provider", "list", "--profile", profile], overrides);
      assert.deepEqual(result.identityProviders, [
        { id: google.id, displayName: "Google", "@odata.type": "#microsoft.graph.socialIdentityProvider" },
        { id: saml.id, displayName: "Contoso SAML", "@odata.type": "#microsoft.graph.samlOrWsFedProvider" },
        { id: apple.id, displayName: "Apple", "@odata.type": "#microsoft.graph.appleManagedIdentityProvider" },
        { id: builtin.id },
      ]);
      assert.deepEqual(result.count, { returned: 4, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra identity-provider show --id <provider-id>")));
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      assert.ok(!JSON.stringify(result).includes("must-never-surface"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/identity/identityProviders?"));
      if (mode === "delegated") assert.deepEqual(calls[0][1], providerScopes);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped provider list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, providerTransport());
      const first = await executeArgv(["entra", "identity-provider", "list", "--profile", profile, "--limit", "3"], overrides);
      assert.equal(first.identityProviders.length, 3);
      assert.deepEqual(first.count, { returned: 3, complete: false, reason: "row limit reached; buffered remainder is preserved in the cursor" });
      assert.equal(typeof first.cursor, "string");
      const { overrides: resumeOverrides } = overridesFor(mode, providerTransport());
      const second = await executeArgv(["entra", "identity-provider", "list", "--profile", profile, "--cursor", first.cursor], resumeOverrides);
      assert.equal(second.identityProviders.length, 1);
      assert.deepEqual(second.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one provider with the full reviewed set and scrubbed secrets`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, providerTransport());
      const result = await executeArgv(["entra", "identity-provider", "show", "--id", saml.id, "--profile", profile], overrides);
      assert.equal(result.identityProvider.id, saml.id);
      assert.equal(result.identityProvider.issuerUri, saml.issuerUri);
      assert.equal(result.identityProvider["@odata.type"], "#microsoft.graph.samlOrWsFedProvider");
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      const social = await executeArgv(["entra", "identity-provider", "show", "--id", google.id, "--profile", profile], overrides);
      assert.equal(social.identityProvider.clientId, "google-client-id");
      assert.ok(!("clientSecret" in social.identityProvider));
      const managed = await executeArgv(["entra", "identity-provider", "show", "--id", apple.id, "--profile", profile], overrides);
      assert.equal(managed.identityProvider.keyId, "KEY123");
      assert.ok(!("certificateData" in managed.identityProvider));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts providers as a scalar number`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode, providerTransport());
      const result = await executeArgv(["entra", "identity-provider", "count", "--profile", profile], overrides);
      assert.deepEqual(result.count, { returned: 4, complete: true });
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/identity/identityProviders/$count"));
      const filtered = await executeArgv(
        ["entra", "identity-provider", "count", "--profile", profile, "--filter", "identityProviderType eq 'Google'"], overrides);
      assert.equal(filtered.count.returned, 4);
      assert.ok(new URL(requests[requests.length - 1].url).searchParams.has("$filter"));
      assert.ok(filtered.help.some(hint => hint.includes("drop --filter for the tenant total")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied provider reads surface scope, role and licence guidance`, async () => {
    const state = setupProfiles();
    try {
      const { calls, requests, overrides } = overridesFor(mode, providerTransport({ denied: true }));
      for (const args of [
        ["entra", "identity-provider", "list", "--profile", profile],
        ["entra", "identity-provider", "show", "--id", saml.id, "--profile", profile],
        ["entra", "identity-provider", "count", "--profile", profile],
      ]) {
        await assert.rejects(executeArgv(args, overrides), error => {
          assert.equal(error.code, "GRAPH_ERROR");
          assert.match(error.message, /\(403\)/);
          assert.match(error.message, /grant, role, licence/);
          assert.ok(error.suggestions.some(hint => hint.includes("IdentityProvider.Read.All")));
          assert.ok(error.suggestions.some(hint => hint.includes("Global Reader")));
          assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
          assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
          return true;
        });
      }
      assert.ok(calls.length > 0);
      assert.ok(requests.length > 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown provider id reports not found, never empty`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, providerTransport());
      await assert.rejects(
        executeArgv(["entra", "identity-provider", "show", "--id", "no-such-provider", "--profile", profile], overrides),
        error => error.code === "GRAPH_ERROR" && /not found or inaccessible/.test(error.message),
      );
    } finally {
      teardownProfiles(state);
    }
  });
}

test("truncated provider text carries a --full hint without lifting redaction or caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated", providerTransport());
    const result = await executeArgv(["entra", "identity-provider", "show", "--id", saml.id, "--profile", "soc"], overrides);
    assert.match(result.identityProvider.metadataExchangeUri, /\.\.\. \(truncated, \d+ chars total\)/);
    assert.ok(result.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "identity-provider", "show", "--id", saml.id, "--profile", "soc", "--full"], overrides);
    assert.equal(full.identityProvider.metadataExchangeUri, longMetadata);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
    const social = await executeArgv(["entra", "identity-provider", "show", "--id", google.id, "--profile", "soc", "--full"], overrides);
    assert.ok(!("clientSecret" in social.identityProvider));
  } finally {
    teardownProfiles(state);
  }
});

test("a non-numeric count body is malformed, never a count", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated", providerTransport({ countBody: JSON.stringify({ value: [] }) }));
    await assert.rejects(
      executeArgv(["entra", "identity-provider", "count", "--profile", "soc"], overrides),
      error => error.code === "GRAPH_ERROR" && /non-numeric success body/.test(error.message),
    );
  } finally {
    teardownProfiles(state);
  }
});

test("unknown properties, secret fields and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { calls, requests, overrides } = overridesFor("delegated", providerTransport());
    await assert.rejects(
      executeArgv(["entra", "identity-provider", "list", "--profile", "soc", "--select", "id,bogus"], overrides),
      /Unknown provider property bogus/,
    );
    await assert.rejects(
      executeArgv(["entra", "identity-provider", "list", "--profile", "soc", "--select", "id,clientSecret"], overrides),
      /secret-bearing provider fields are never returned/,
    );
    await assert.rejects(
      executeArgv(["entra", "identity-provider", "show", "--id", google.id, "--profile", "soc", "--select", "id,certificateData"], overrides),
      /secret-bearing provider fields are never returned/,
    );
    await assert.rejects(
      executeArgv(["entra", "identity-provider", "list", "--profile", "soc", "--fields", "issuerUri"], overrides),
      /was not fetched; request it with --select/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("provider read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "identity-provider", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "identity-provider", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "identity-provider", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "identity-provider", "count", "--select", "id"]), /unknown flag --select/);
  await assert.rejects(executeArgv(["entra", "identity-provider", "count", "--limit", "10"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "identity-provider", "list", "--limit", "10", "--all"]), { code: "VALIDATION_ERROR" });
});

for (const id of ["$count", "$value", "$ref"]) {
  test(`identity-provider show rejects reserved resource binding ${id} before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { calls, requests, overrides } = overridesFor("delegated", providerTransport());
      await assert.rejects(
        executeArgv(["entra", "identity-provider", "show", "--id", id, "--profile", "soc"], overrides),
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
    ["entra", "identity-provider", "list", "--profile", profile],
    ["entra", "identity-provider", "show", "--id", saml.id, "--profile", profile],
    ["entra", "identity-provider", "count", "--profile", profile],
  ]) {
    test(`${mode} ${args[1]} ${args[2]} refuses beta before credentials`, async () => {
      const state = setupProfiles();
      try {
        const { calls, requests, overrides } = overridesFor(mode, providerTransport());
        await assert.rejects(
          executeArgv([...args, "--api-version", "beta"], overrides),
          { code: "VALIDATION_ERROR", message: "Identity-provider reads support v1.0 only; beta needs its own review" },
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
    const { calls, requests, overrides } = overridesFor("application", providerTransport());
    await assert.rejects(
      executeArgv(["entra", "identity-provider", "list", "--profile", "batch", "--scopes", providerScopes[0]], overrides),
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
  if (url.pathname === "/v1.0/identity/identityProviders") return { status: 200, headers: {}, body: JSON.stringify({ value: providers }) };
  if (url.pathname === "/v1.0/identity/identityProviders/Google-OAUTH") return { status: 200, headers: {}, body: JSON.stringify(google) };
  return { status: 404, headers: {}, body: JSON.stringify({ error: { code: "Unknown", message: "unexpected route" } }) };
}

test("raw api serves reviewed provider list and show without secret fields", async () => {
  const f = rawFixture(rawBody);
  const list = await runApiGet({ path: "/identity/identityProviders", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: providerScopes[0] }, f.deps);
  assert.equal(list.returned, 4);
  assert.equal(list.complete, true);
  assert.ok(!JSON.stringify(list.value).includes("must-never-surface"));
  assert.ok(list.value[0].clientId === "google-client-id");
  const show = await runApiGet({ path: "/identity/identityProviders/Google-OAUTH", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: providerScopes[0] }, f.deps);
  assert.equal(show.displayName, "Google");
  assert.ok(!("clientSecret" in show));
});

test("raw api refuses the provider $count scalar before credentials", async () => {
  const f = rawFixture(rawBody);
  await assert.rejects(
    runApiGet({ path: "/identity/identityProviders/$count", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: providerScopes[0] }, f.deps),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("raw api refuses unreviewed provider $select fields before credentials", async () => {
  const f = rawFixture(rawBody);
  await assert.rejects(
    runApiGet({ path: "/identity/identityProviders", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: providerScopes[0], odata: "$select=id,clientSecret" }, f.deps),
    error => error.code === "VALIDATION_ERROR" && /Unreviewed \$select field clientSecret/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});
