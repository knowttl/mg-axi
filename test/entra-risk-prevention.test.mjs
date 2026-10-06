import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import { runApiGet } from "../dist/api.js";
import { DelegatedAuth } from "../dist/auth.js";
import { ApplicationAuth } from "../dist/app-auth.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const riskScopes = ["https://graph.microsoft.com/RiskPreventionProviders.Read.All"];

const arkose = {
  "@odata.type": "#microsoft.graph.arkoseFraudProtectionProvider",
  id: "fraud-arkose-1",
  displayName: "Arkose Sign-Up Protection",
  publicKey: "A1EE42E0-C88B-47FE-A176-5E9FB8F116FB",
  privateKey: "fixture-arkose-private-key",
  clientSubDomain: "client-api",
  verifySubDomain: "verify-api",
};
const human = {
  "@odata.type": "#microsoft.graph.humanSecurityFraudProtectionProvider",
  id: "fraud-human-2",
  displayName: "HUMAN Sign-Up Protection",
  appId: "PXfixtureapp",
  serverToken: "fixture-human-server-token",
};
const fraudBare = { id: "fraud-bare-3" };
const fraudProviders = [arkose, human, fraudBare];

const cloudflare = {
  "@odata.type": "#microsoft.graph.cloudFlareWebApplicationFirewallProvider",
  id: "waf-cloudflare-1",
  displayName: "Cloudflare Provider Example",
  zoneId: "11111111111111111111111111111111",
  apiToken: "fixture-cloudflare-api-token",
};
const akamai = {
  "@odata.type": "#microsoft.graph.akamaiWebApplicationFirewallProvider",
  id: "waf-akamai-2",
  displayName: "Akamai Provider Example",
  hostPrefix: "akab-exampleprefix",
  clientSecret: "fixture-akamai-client-secret",
  clientToken: "fixture-akamai-client-token",
  accessToken: "fixture-akamai-access-token",
};
const wafBare = { id: "waf-bare-3" };
const wafProviders = [cloudflare, akamai, wafBare];

const verification1 = {
  id: "verif-1",
  verifiedHost: "www.contoso.com",
  providerType: "cloudflare",
  verificationResult: { status: "success", verifiedOnDateTime: "2025-10-04T00:50:26.4909654Z", errors: [], warnings: [] },
  verifiedDetails: {
    "@odata.type": "#microsoft.graph.cloudFlareVerifiedDetailsModel",
    zoneId: "11111111111111111111111111111111",
    dnsConfiguration: { name: "www.contoso.com", isProxied: true, recordType: "cname", value: "contoso.azurefd.net", isDomainVerified: true },
    enabledRecommendedRulesets: [{ rulesetId: "22222222222222222222222222222222", name: "CloudFlare Managed Ruleset", phaseName: "http_request_firewall_managed" }],
    enabledCustomRules: [{ ruleId: "33333333333333333333333333333333", name: "Block SQL Injection", action: "block" }],
  },
};
const verificationBare = { id: "verif-bare-2" };
const verifications = [verification1, verificationBare];
const longHost = `www.${"h".repeat(600)}.contoso.com`;

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-risk-prevention-"));
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

function riskTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    const base = "/v1.0/identity/riskPrevention";
    if (path === `${base}/fraudProtectionProviders/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(fraudProviders.length) };
    }
    if (path === `${base}/webApplicationFirewallProviders/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(wafProviders.length) };
    }
    if (path === `${base}/webApplicationFirewallVerifications/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(verifications.length) };
    }
    if (path === `${base}/fraudProtectionProviders`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [fraudBare] });
      return json(200, {
        value: [arkose, human],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identity/riskPrevention/fraudProtectionProviders?%24skiptoken=page2",
      });
    }
    if (path === `${base}/webApplicationFirewallProviders`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [wafBare] });
      return json(200, {
        value: [cloudflare, akamai],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identity/riskPrevention/webApplicationFirewallProviders?%24skiptoken=page2",
      });
    }
    if (path === `${base}/webApplicationFirewallVerifications`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [verificationBare] });
      return json(200, {
        value: [verification1],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/identity/riskPrevention/webApplicationFirewallVerifications?%24skiptoken=page2",
      });
    }
    const provider = new RegExp(`^${base}/webApplicationFirewallVerifications/([^/]+)/provider$`).exec(path);
    if (provider) {
      if (decodeURIComponent(provider[1]) === verification1.id) return json(200, cloudflare);
      return json(404, { error: { code: "Request_ResourceNotFound", message: "no such verification" } });
    }
    const single = new RegExp(`^${base}/([^/]+)/([^/]+)$`).exec(path);
    if (single) {
      const rows = single[1] === "fraudProtectionProviders" ? fraudProviders
        : single[1] === "webApplicationFirewallProviders" ? wafProviders
        : single[1] === "webApplicationFirewallVerifications" ? verifications : [];
      const found = rows.find(row => row.id === decodeURIComponent(single[2]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such resource" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, fixture, calls = []) {
  const transportFixture = fixture ?? riskTransport();
  const credential = credentialService(mode, calls);
  return {
    requests: transportFixture.requests,
    calls,
    overrides: {
      transport: transportFixture.send,
      delegated: mode === "delegated" ? credential : credentialService("delegated", []),
      application: mode === "application" ? credential : credentialService("application", []),
    },
  };
}

const SECRET_VALUES = [
  "fixture-arkose-private-key",
  "A1EE42E0-C88B-47FE-A176-5E9FB8F116FB",
  "fixture-human-server-token",
  "fixture-cloudflare-api-token",
  "fixture-akamai-client-secret",
  "fixture-akamai-client-token",
  "fixture-akamai-access-token",
];

test("delegated lists fraud providers with compact rows and scrubbed key material", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "fraud-protection-provider", "list", "--profile", "soc"], overrides);
    assert.deepEqual(result.fraudProtectionProviders, [
      { id: arkose.id, displayName: "Arkose Sign-Up Protection", "@odata.type": "#microsoft.graph.arkoseFraudProtectionProvider" },
      { id: human.id, displayName: "HUMAN Sign-Up Protection", "@odata.type": "#microsoft.graph.humanSecurityFraudProtectionProvider" },
      { id: fraudBare.id },
    ]);
    assert.deepEqual(result.count, { returned: 3, complete: true });
    assert.ok(result.help.some(hint => hint.includes("entra fraud-protection-provider show --id <fraud-protection-provider-id>")));
    assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
    for (const secret of SECRET_VALUES) assert.ok(!JSON.stringify(result).includes(secret), `leaked ${secret}`);
    assert.ok(!JSON.stringify(result).includes("opaque-fixture-delegated-token"));
    assert.ok(requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-delegated-token"));
    assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/identity/riskPrevention/fraudProtectionProviders?"));
    const sent = new URL(requests[0].url).searchParams;
    assert.equal(sent.get("$select"), "id,displayName");
    assert.ok(!sent.has("$filter"));
    assert.deepEqual(calls[0][1], riskScopes);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated shows one fraud provider with the full set and scrubbed secrets", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "fraud-protection-provider", "show", "--id", arkose.id, "--profile", "soc"], overrides);
    assert.deepEqual(result.fraudProtectionProvider, {
      id: arkose.id,
      displayName: "Arkose Sign-Up Protection",
      clientSubDomain: "client-api",
      verifySubDomain: "verify-api",
      "@odata.type": "#microsoft.graph.arkoseFraudProtectionProvider",
    });
    assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
    for (const secret of SECRET_VALUES) assert.ok(!JSON.stringify(result).includes(secret), `leaked ${secret}`);
    const app = await executeArgv(["entra", "fraud-protection-provider", "show", "--id", human.id, "--profile", "soc"], overrides);
    assert.deepEqual(app.fraudProtectionProvider, {
      id: human.id,
      displayName: "HUMAN Sign-Up Protection",
      appId: "PXfixtureapp",
      "@odata.type": "#microsoft.graph.humanSecurityFraudProtectionProvider",
    });
    const bare = await executeArgv(["entra", "fraud-protection-provider", "show", "--id", fraudBare.id, "--profile", "soc"], overrides);
    assert.deepEqual(bare.fraudProtectionProvider, { id: fraudBare.id });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated lists WAF providers with compact rows, filter support and scrubbed credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "web-application-firewall-provider", "list", "--profile", "soc"], overrides);
    assert.deepEqual(result.webApplicationFirewallProviders, [
      { id: cloudflare.id, displayName: "Cloudflare Provider Example", "@odata.type": "#microsoft.graph.cloudFlareWebApplicationFirewallProvider" },
      { id: akamai.id, displayName: "Akamai Provider Example", "@odata.type": "#microsoft.graph.akamaiWebApplicationFirewallProvider" },
      { id: wafBare.id },
    ]);
    assert.deepEqual(result.count, { returned: 3, complete: true });
    assert.ok(result.help.some(hint => hint.includes("entra web-application-firewall-provider show --id <web-application-firewall-provider-id>")));
    for (const secret of SECRET_VALUES) assert.ok(!JSON.stringify(result).includes(secret), `leaked ${secret}`);
    const seen = requests.length;
    const filtered = await executeArgv(
      ["entra", "web-application-firewall-provider", "list", "--profile", "soc", "--filter", "displayName eq 'Akamai Provider Example'"], overrides);
    const sent = new URL(requests[seen].url).searchParams;
    assert.equal(sent.get("$filter"), "displayName eq 'Akamai Provider Example'");
    assert.equal(sent.get("$select"), "id,displayName");
    assert.deepEqual(filtered.webApplicationFirewallProviders.map(row => row.id), [cloudflare.id, akamai.id, wafBare.id]);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated shows one WAF provider with the full set and scrubbed credentials", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "web-application-firewall-provider", "show", "--id", akamai.id, "--profile", "soc"], overrides);
    assert.deepEqual(result.webApplicationFirewallProvider, {
      id: akamai.id,
      displayName: "Akamai Provider Example",
      hostPrefix: "akab-exampleprefix",
      "@odata.type": "#microsoft.graph.akamaiWebApplicationFirewallProvider",
    });
    for (const secret of SECRET_VALUES) assert.ok(!JSON.stringify(result).includes(secret), `leaked ${secret}`);
    const zone = await executeArgv(["entra", "web-application-firewall-provider", "show", "--id", cloudflare.id, "--profile", "soc"], overrides);
    assert.deepEqual(zone.webApplicationFirewallProvider, {
      id: cloudflare.id,
      displayName: "Cloudflare Provider Example",
      zoneId: "11111111111111111111111111111111",
      "@odata.type": "#microsoft.graph.cloudFlareWebApplicationFirewallProvider",
    });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated lists verifications with compact rows and whole complex detail on show", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "web-application-firewall-verification", "list", "--profile", "soc"], overrides);
    assert.deepEqual(result.webApplicationFirewallVerifications, [
      { id: verification1.id, verifiedHost: "www.contoso.com", providerType: "cloudflare" },
      { id: verificationBare.id },
    ]);
    assert.deepEqual(result.count, { returned: 2, complete: true });
    assert.ok(result.help.some(hint => hint.includes("entra web-application-firewall-verification show --id <web-application-firewall-verification-id>")));
    assert.ok(!("provider" in result.webApplicationFirewallVerifications[0]));
    const sent = new URL(requests[0].url).searchParams;
    assert.equal(sent.get("$select"), "id,verifiedHost,providerType");
    const shown = await executeArgv(["entra", "web-application-firewall-verification", "show", "--id", verification1.id, "--profile", "soc"], overrides);
    assert.deepEqual(shown.webApplicationFirewallVerification, verification1);
    assert.ok(!("provider" in shown.webApplicationFirewallVerification));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated shows a verification provider with the reviewed provider set and scrubbed credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "web-application-firewall-verification", "provider", "show", "--id", verification1.id, "--profile", "soc"], overrides);
    assert.deepEqual(result.webApplicationFirewallProvider, {
      id: cloudflare.id,
      displayName: "Cloudflare Provider Example",
      zoneId: "11111111111111111111111111111111",
      "@odata.type": "#microsoft.graph.cloudFlareWebApplicationFirewallProvider",
    });
    assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
    for (const secret of SECRET_VALUES) assert.ok(!JSON.stringify(result).includes(secret), `leaked ${secret}`);
    assert.equal(requests[0].url.split("?")[0],
      `https://graph.microsoft.com/v1.0/identity/riskPrevention/webApplicationFirewallVerifications/${verification1.id}/provider`);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated resumes a capped fraud list through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "fraud-protection-provider", "list", "--profile", "soc", "--limit", "1"], overrides);
    assert.deepEqual(first.fraudProtectionProviders.map(row => row.id), [arkose.id]);
    assert.equal(first.count.complete, false);
    assert.equal(typeof first.cursor, "string");
    const second = await executeArgv(["entra", "fraud-protection-provider", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.fraudProtectionProviders.map(row => row.id), [human.id, fraudBare.id]);
    assert.deepEqual(second.count, { returned: 2, complete: true });
  } finally {
    teardownProfiles(state);
  }
});

for (const [command, expected, route] of [
  [["fraud-protection-provider", "count"], 3, "fraudProtectionProviders/$count"],
  [["web-application-firewall-provider", "count"], 3, "webApplicationFirewallProviders/$count"],
  [["web-application-firewall-verification", "count"], 2, "webApplicationFirewallVerifications/$count"],
]) {
  test(`delegated ${command.join(" ")} counts as one scalar without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor("delegated");
      const result = await executeArgv(["entra", ...command, "--profile", "soc"], overrides);
      assert.deepEqual(result.count, { returned: expected, complete: true });
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, `https://graph.microsoft.com/v1.0/identity/riskPrevention/${route}`);
    } finally {
      teardownProfiles(state);
    }
  });
}

test("delegated denied risk-prevention reads surface scope, role and licence guidance", async () => {
  const state = setupProfiles();
  try {
    const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
    const { overrides } = overridesFor("delegated", denied);
    for (const args of [
      ["entra", "fraud-protection-provider", "list", "--profile", "soc"],
      ["entra", "fraud-protection-provider", "show", "--id", arkose.id, "--profile", "soc"],
      ["entra", "fraud-protection-provider", "count", "--profile", "soc"],
      ["entra", "web-application-firewall-provider", "list", "--profile", "soc"],
      ["entra", "web-application-firewall-provider", "show", "--id", akamai.id, "--profile", "soc"],
      ["entra", "web-application-firewall-provider", "count", "--profile", "soc"],
      ["entra", "web-application-firewall-verification", "list", "--profile", "soc"],
      ["entra", "web-application-firewall-verification", "show", "--id", verification1.id, "--profile", "soc"],
      ["entra", "web-application-firewall-verification", "count", "--profile", "soc"],
      ["entra", "web-application-firewall-verification", "provider", "show", "--id", verification1.id, "--profile", "soc"],
    ]) {
      await assert.rejects(executeArgv(args, overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.match(error.message, /\(403\)/);
        assert.match(error.message, /grant, role, licence/);
        assert.ok(error.suggestions.some(hint => hint.includes("RiskPreventionProviders.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Security Reader")));
        assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return true;
      });
    }
  } finally {
    teardownProfiles(state);
  }
});

test("delegated unknown ids report absence, not emptiness", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    for (const args of [
      ["entra", "fraud-protection-provider", "show", "--id", "fraud-missing", "--profile", "soc"],
      ["entra", "web-application-firewall-provider", "show", "--id", "waf-missing", "--profile", "soc"],
      ["entra", "web-application-firewall-verification", "show", "--id", "verif-missing", "--profile", "soc"],
      ["entra", "web-application-firewall-verification", "provider", "show", "--id", "verif-missing", "--profile", "soc"],
    ]) {
      await assert.rejects(
        executeArgv(args, overrides),
        error => error.code === "GRAPH_ERROR" && /not found or inaccessible/.test(error.message),
      );
    }
  } finally {
    teardownProfiles(state);
  }
});

for (const [command, args] of [
  [["fraud-protection-provider", "list"], []],
  [["fraud-protection-provider", "show"], ["--id", arkose.id]],
  [["fraud-protection-provider", "count"], []],
  [["web-application-firewall-provider", "list"], []],
  [["web-application-firewall-provider", "show"], ["--id", akamai.id]],
  [["web-application-firewall-provider", "count"], []],
  [["web-application-firewall-verification", "list"], []],
  [["web-application-firewall-verification", "show"], ["--id", verification1.id]],
  [["web-application-firewall-verification", "count"], []],
  [["web-application-firewall-verification", "provider", "show"], ["--id", verification1.id]],
]) {
  for (const preview of [false, true]) {
    test(`delegated ${command.join(" ")} refuses beta before credentials with preview=${preview}`, async () => {
      const state = setupProfiles();
      try {
        const path = join(state.dir, "config.json");
        const config = JSON.parse(readFileSync(path, "utf8"));
        config.profiles.soc.preview = preview;
        writeFileSync(path, JSON.stringify(config));
        const { calls, requests, overrides } = overridesFor("delegated");
        await assert.rejects(
          executeArgv(["entra", ...command, ...args, "--profile", "soc", "--api-version", "beta"], overrides),
          { code: "VALIDATION_ERROR", message: "Risk-prevention reads support v1.0 only; beta needs its own review" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }
}

test("truncated verification text carries a --full hint without lifting redaction or caps", async () => {
  const state = setupProfiles();
  try {
    const longRow = transport(() => json(200, { value: [{ ...verification1, id: "verif-long-9", verifiedHost: longHost }] }));
    const { overrides } = overridesFor("delegated", longRow);
    const partial = await executeArgv(["entra", "web-application-firewall-verification", "list", "--profile", "soc"], overrides);
    const truncated = partial.webApplicationFirewallVerifications.find(row => row.id === "verif-long-9");
    assert.match(truncated.verifiedHost, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "web-application-firewall-verification", "list", "--profile", "soc", "--full"], overrides);
    assert.equal(full.webApplicationFirewallVerifications.find(row => row.id === "verif-long-9").verifiedHost, longHost);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("unknown properties, secret fields and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "fraud-protection-provider", "list", "--profile", "soc", "--select", "id,bogus"], overrides),
      /Unknown fraud-protection-provider property bogus/,
    );
    await assert.rejects(
      executeArgv(["entra", "fraud-protection-provider", "list", "--profile", "soc", "--select", "id,serverToken"], overrides),
      /cannot request serverToken: secret-bearing provider fields are never returned/,
    );
    await assert.rejects(
      executeArgv(["entra", "web-application-firewall-provider", "show", "--id", akamai.id, "--profile", "soc", "--select", "id,apiToken"], overrides),
      /cannot request apiToken: secret-bearing provider fields are never returned/,
    );
    await assert.rejects(
      executeArgv(["entra", "web-application-firewall-verification", "list", "--profile", "soc", "--fields", "verificationResult"], overrides),
      /was not fetched; request it with --select/,
    );
    await assert.rejects(
      executeArgv(["entra", "web-application-firewall-verification", "provider", "show", "--id", verification1.id, "--profile", "soc", "--select", "id,clientSecret"], overrides),
      /cannot request clientSecret: secret-bearing provider fields are never returned/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("risk-prevention read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "fraud-protection-provider", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "fraud-protection-provider", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "fraud-protection-provider", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "fraud-protection-provider", "list", "--filter", "displayName eq 'x'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "fraud-protection-provider", "count", "--filter", "displayName eq 'x'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "fraud-protection-provider", "count", "--select", "id"]), /unknown flag --select/);
  await assert.rejects(executeArgv(["entra", "fraud-protection-provider", "count", "--limit", "10"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "web-application-firewall-verification", "provider", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "web-application-firewall-verification", "count", "--cursor", "abc"]), /unknown flag --cursor/);
  await assert.rejects(executeArgv(["entra", "web-application-firewall-provider", "list", "--limit", "10", "--all"]), { code: "VALIDATION_ERROR" });
});

for (const command of [
  ["fraud-protection-provider", "list"],
  ["fraud-protection-provider", "show", "--id", arkose.id],
  ["fraud-protection-provider", "count"],
  ["web-application-firewall-provider", "list"],
  ["web-application-firewall-provider", "show", "--id", akamai.id],
  ["web-application-firewall-provider", "count"],
  ["web-application-firewall-verification", "list"],
  ["web-application-firewall-verification", "show", "--id", verification1.id],
  ["web-application-firewall-verification", "count"],
  ["web-application-firewall-verification", "provider", "show", "--id", verification1.id],
]) {
  test(`application ${command.join(" ")} is refused before credentials without a supported application permission`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor("application");
      await assert.rejects(
        executeArgv(["entra", ...command, "--profile", "batch"], overrides),
        error => {
          assert.equal(error.code, "VALIDATION_ERROR");
          return /need a delegated profile; Graph documents no supported application permission/.test(error.message);
        },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

test("malformed risk-prevention count bodies fail as unknown, not zero", async () => {
  const state = setupProfiles();
  try {
    for (const body of ["{}", "-1", "2.5", "\"3\""]) {
      const malformed = transport(() => ({ status: 200, headers: {}, body }));
      const scoped = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "fraud-protection-provider", "count", "--profile", "soc"], scoped.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /non-numeric success body/.test(error.message);
      });
    }
  } finally {
    teardownProfiles(state);
  }
});

test("malformed single risk-prevention bodies fail as unknown, not empty", async () => {
  const state = setupProfiles();
  try {
    for (const body of [null, [], "3"]) {
      const malformed = transport(() => json(200, body));
      const scoped = overridesFor("delegated", malformed);
      for (const args of [
        ["entra", "fraud-protection-provider", "show", "--id", arkose.id, "--profile", "soc"],
        ["entra", "web-application-firewall-provider", "show", "--id", akamai.id, "--profile", "soc"],
        ["entra", "web-application-firewall-verification", "show", "--id", verification1.id, "--profile", "soc"],
        ["entra", "web-application-firewall-verification", "provider", "show", "--id", verification1.id, "--profile", "soc"],
      ]) {
        await assert.rejects(
          executeArgv(args, scoped.overrides),
          error => error.code === "GRAPH_ERROR" && /Graph returned a malformed web-application-firewall-(provider|verification) body|Graph returned a malformed fraud-protection-provider body/.test(error.message),
        );
      }
    }
  } finally {
    teardownProfiles(state);
  }
});

for (const [command, id] of [
  [["fraud-protection-provider", "show"], "$count"],
  [["web-application-firewall-provider", "show"], "$value"],
  [["web-application-firewall-verification", "show"], "$ref"],
  [["web-application-firewall-verification", "provider", "show"], "$count"],
]) {
  test(`${command.join(" ")} rejects reserved resource binding ${id} before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { calls, requests, overrides } = overridesFor("delegated");
      await assert.rejects(
        executeArgv(["entra", ...command, "--id", id, "--profile", "soc"], overrides),
        { code: "VALIDATION_ERROR", message: "OData reserved segments cannot be resource identifiers" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

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
    acquire: async (...args) => { credentialCalls.push(args); return credential({}); },
  });
  const transport = async request => {
    requests.push(request);
    const response = typeof handler === "function" ? await handler(request, requests.length) : handler;
    return { headers: {}, body: "", ...response };
  };
  return { credentialCalls, requests, deps: { delegated, application, transport } };
}

const delegatedRawProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key } };
const appRawProfile = { ...delegatedRawProfile, mode: "application", credentialRef: { provider: "federated", key } };
const rawScopes = riskScopes.join(",");
const riskBase = "/v1.0/identity/riskPrevention";

function rawBody(request) {
  const url = new URL(request.url);
  if (url.pathname === `${riskBase}/fraudProtectionProviders`) return { status: 200, headers: {}, body: JSON.stringify({ value: fraudProviders }) };
  if (url.pathname === `${riskBase}/fraudProtectionProviders/${arkose.id}`) return { status: 200, headers: {}, body: JSON.stringify(arkose) };
  if (url.pathname === `${riskBase}/webApplicationFirewallProviders`) return { status: 200, headers: {}, body: JSON.stringify({ value: wafProviders }) };
  if (url.pathname === `${riskBase}/webApplicationFirewallProviders/${akamai.id}`) return { status: 200, headers: {}, body: JSON.stringify(akamai) };
  if (url.pathname === `${riskBase}/webApplicationFirewallVerifications`) return { status: 200, headers: {}, body: JSON.stringify({ value: verifications }) };
  if (url.pathname === `${riskBase}/webApplicationFirewallVerifications/${verification1.id}`) return { status: 200, headers: {}, body: JSON.stringify(verification1) };
  if (url.pathname === `${riskBase}/webApplicationFirewallVerifications/${verification1.id}/provider`) {
    return { status: 200, headers: {}, body: JSON.stringify(cloudflare) };
  }
  return { status: 404, headers: {}, body: JSON.stringify({ error: { code: "Unknown", message: "unexpected route" } }) };
}

test("raw api serves reviewed risk-prevention lists and singles with scrubbed key material", async () => {
  const f = rawFixture(rawBody);
  const fraud = await runApiGet({ path: "/identity/riskPrevention/fraudProtectionProviders", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes }, f.deps);
  assert.equal(fraud.returned, 3);
  assert.equal(fraud.complete, true);
  for (const secret of SECRET_VALUES) assert.ok(!JSON.stringify(fraud).includes(secret), `leaked ${secret}`);
  const shown = await runApiGet({ path: `/identity/riskPrevention/fraudProtectionProviders/${arkose.id}`, apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes }, f.deps);
  assert.equal(shown.id, arkose.id);
  assert.equal(shown.clientSubDomain, "client-api");
  assert.ok(!("privateKey" in shown) && !("publicKey" in shown));
  const waf = await runApiGet({ path: "/identity/riskPrevention/webApplicationFirewallProviders", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes }, f.deps);
  assert.equal(waf.returned, 3);
  for (const secret of SECRET_VALUES) assert.ok(!JSON.stringify(waf).includes(secret), `leaked ${secret}`);
  const verifications = await runApiGet({ path: "/identity/riskPrevention/webApplicationFirewallVerifications", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes }, f.deps);
  assert.equal(verifications.returned, 2);
  const provider = await runApiGet({ path: `/identity/riskPrevention/webApplicationFirewallVerifications/${verification1.id}/provider`, apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes }, f.deps);
  assert.equal(provider.id, cloudflare.id);
  assert.ok(!("apiToken" in provider));
});

test("raw api refuses the risk-prevention $count scalars before credentials", async () => {
  const f = rawFixture(rawBody);
  for (const path of [
    "/identity/riskPrevention/fraudProtectionProviders/$count",
    "/identity/riskPrevention/webApplicationFirewallProviders/$count",
    "/identity/riskPrevention/webApplicationFirewallVerifications/$count",
  ]) {
    await assert.rejects(
      runApiGet({ path, apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes }, f.deps),
      { code: "VALIDATION_ERROR" },
    );
  }
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("raw api refuses unreviewed risk-prevention $select fields before credentials", async () => {
  const f = rawFixture(rawBody);
  await assert.rejects(
    runApiGet({ path: "/identity/riskPrevention/fraudProtectionProviders", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes, odata: "$select=id,serverToken" }, f.deps),
    error => error.code === "VALIDATION_ERROR" && /Unreviewed \$select field serverToken/.test(error.message),
  );
  await assert.rejects(
    runApiGet({ path: "/identity/riskPrevention/webApplicationFirewallProviders", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes, odata: "$select=id,apiToken" }, f.deps),
    error => error.code === "VALIDATION_ERROR" && /Unreviewed \$select field apiToken/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("raw risk-prevention reads refuse application profiles before credentials", async () => {
  const f = rawFixture(rawBody);
  for (const path of [
    "/identity/riskPrevention/fraudProtectionProviders",
    `/identity/riskPrevention/fraudProtectionProviders/${arkose.id}`,
    "/identity/riskPrevention/webApplicationFirewallProviders",
    `/identity/riskPrevention/webApplicationFirewallProviders/${akamai.id}`,
    "/identity/riskPrevention/webApplicationFirewallVerifications",
    `/identity/riskPrevention/webApplicationFirewallVerifications/${verification1.id}`,
    `/identity/riskPrevention/webApplicationFirewallVerifications/${verification1.id}/provider`,
  ]) {
    await assert.rejects(
      runApiGet({ path, apiVersion: "v1.0", profile: appRawProfile }, f.deps),
      error => error.code === "VALIDATION_ERROR" && /needs a delegated profile/.test(error.message),
    );
  }
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});
