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

// EXT-01 organization acceptance through the real CLI interface with
// fixture-only transports: tenant list/show with compact and full reviewed
// sets, default-branding show with its documented Accept-Language behavior
// and unconfigured-404 guidance, branding-localization list/show, Stream
// property refusal, denial guidance and the offline executable journey. No
// live tenant, real credential or customer data exists.

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const orgId = "84841066-274d-4ec0-a5c1-276be684bdd3";
const orgScopes = ["https://graph.microsoft.com/Organization.Read.All"];
const brandingScopes = ["https://graph.microsoft.com/User.Read"];

const org1 = {
  id: orgId,
  deletedDateTime: null,
  businessPhones: ["425-555-0100"],
  city: null,
  country: null,
  countryLetterCode: "NL",
  createdDateTime: "2021-08-02T10:30:06Z",
  defaultUsageLocation: "NL",
  displayName: "Contoso",
  isMultipleDataLocationsForServicesEnabled: null,
  marketingNotificationEmails: [],
  onPremisesLastSyncDateTime: null,
  onPremisesSyncEnabled: null,
  partnerTenantType: null,
  postalCode: null,
  preferredLanguage: "en",
  privacyProfile: { contactEmail: "", statementUrl: "" },
  provisionedPlans: [],
  securityComplianceNotificationMails: [],
  securityComplianceNotificationPhones: [],
  state: null,
  street: null,
  technicalNotificationMails: ["admin@contoso.com"],
  tenantType: "AAD",
  assignedPlans: [],
  verifiedDomains: [
    { capabilities: "Email, OfficeCommunicationsOnline", isDefault: true, isInitial: true, name: "contoso.com", type: "Managed" },
  ],
};
// A second synthetic row exercises the generic capped-collection path; a
// real tenant carries exactly one organization.
const org2 = { id: "dcd219dd-bc68-4b9b-bf0b-4a33a796be35", displayName: "Contoso Subsidiary" };
const organizations = [org1, org2];

const branding = {
  id: "0",
  backgroundColor: "",
  backgroundImageRelativeUrl: "c1c6b6c8/logintenantbranding/0/illustration?ts=637535563816027796",
  bannerLogoRelativeUrl: "c1c6b6c8/logintenantbranding/0/bannerlogo?ts=637535563824629275",
  cdnList: ["secure.aadcdn.microsoftonline-p.com", "aadcdn.msftauthimages.net"],
  signInPageText: "Contoso",
  squareLogoRelativeUrl: "c1c6b6c8/logintenantbranding/0/tilelogo?ts=637535563832888580",
  usernameHintText: "",
};

const longWelcome = `Welcome to Contoso France ${"s'il vous plait ".repeat(40)}`;
const locDefault = {
  id: "0",
  backgroundColor: "",
  backgroundImageRelativeUrl: null,
  bannerLogoRelativeUrl: null,
  cdnList: [],
  signInPageText: "Contoso",
  squareLogoRelativeUrl: null,
  usernameHintText: "",
};
const locFr = {
  id: "fr-FR",
  backgroundColor: "#FFFF33",
  backgroundImageRelativeUrl: null,
  bannerLogoRelativeUrl: null,
  cdnList: [],
  signInPageText: longWelcome,
  squareLogoRelativeUrl: null,
  usernameHintText: "Welcome to Contoso France",
};
const localizations = [locDefault, locFr];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-organization-"));
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

function organizationTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/organization") return json(200, { value: organizations });
    if (path === `/v1.0/organization/${orgId}/branding/localizations`) return json(200, { value: localizations });
    if (path === `/v1.0/organization/${orgId}/branding`) {
      assert.equal(request.headers["Accept-Language"], "0");
      return json(200, branding);
    }
    const singleLocalization = new RegExp(`^/v1\\.0/organization/${orgId}/branding/localizations/([^/]+)$`).exec(path);
    if (singleLocalization) {
      const found = localizations.find(row => row.id === decodeURIComponent(singleLocalization[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such locale" } });
    }
    const singleOrganization = /^\/v1\.0\/organization\/([^/]+)$/.exec(path);
    if (singleOrganization) {
      const found = organizations.find(row => row.id === decodeURIComponent(singleOrganization[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such organization" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? organizationTransport();
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

function runOrganizationCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-organization-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, organizations, branding, localizations, denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists organizations with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "organization", "list", "--profile", profile], overrides);
      assert.deepEqual(result.organizations, [
        { id: org1.id, displayName: "Contoso", tenantType: "AAD", verifiedDomains: org1.verifiedDomains },
        { id: org2.id, displayName: org2.displayName },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra organization show --id <organization-id>")));
      assert.ok(result.help.some(hint => hint.includes("no --filter")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/organization?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(orgScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} organization list has no --filter flag to misuse`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "organization", "list", "--profile", profile, "--filter", "displayName eq 'Contoso'"], overrides),
        error => {
          assert.equal(error.code, "VALIDATION_ERROR");
          return /unknown flag --filter/.test(error.message);
        },
      );
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped organization list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "organization", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.organizations.map(row => row.id), [org1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "organization", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.organizations.map(row => row.id), [org2.id]);
      assert.deepEqual(second.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one organization with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "organization", "show", "--id", orgId, "--profile", profile], overrides);
      assert.deepEqual(result.organization, org1);
      assert.equal(result.help, undefined);
      const technical = await executeArgv(["entra", "organization", "show", "--id", orgId, "--profile", profile,
        "--select", "id,technicalNotificationMails"], overrides);
      assert.deepEqual(technical.organization, { id: orgId, technicalNotificationMails: ["admin@contoso.com"] });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows default branding with the documented locale behavior`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "organization", "branding", "show",
        "--organization", orgId, "--profile", profile], overrides);
      assert.deepEqual(result.branding, branding);
      assert.ok(result.help.some(hint => hint.includes("entra organization branding-localization list --organization")));
      assert.equal(requests[0].headers["Accept-Language"], "0");
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(brandingScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists branding localizations with compact rows`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "organization", "branding-localization", "list",
        "--organization", orgId, "--profile", profile,
        "--select", "id,signInPageText,usernameHintText,backgroundColor"], overrides);
      assert.deepEqual(result.brandingLocalizations, [
        { id: locDefault.id, signInPageText: "Contoso", usernameHintText: "", backgroundColor: "" },
        { id: locFr.id, signInPageText: longWelcome.slice(0, 500) + `... (truncated, ${longWelcome.length} chars total)`, usernameHintText: locFr.usernameHintText, backgroundColor: locFr.backgroundColor },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra organization branding-localization show")));
      assert.ok(result.help.some(hint => hint.includes("--full")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one branding localization`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "organization", "branding-localization", "show",
        "--organization", orgId, "--id", locDefault.id, "--profile", profile,
        "--select", "id,signInPageText"], overrides);
      assert.deepEqual(result.brandingLocalization, { id: locDefault.id, signInPageText: "Contoso" });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unconfigured branding reads as absent, not denied`, async () => {
    const state = setupProfiles();
    try {
      const missing = transport(() => json(404, { error: { code: "Request_ResourceNotFound", message: "no branding" } }));
      const overrides = {
        transport: missing.send,
        delegated: credentialService("delegated", []),
        application: credentialService("application", []),
      };
      await assert.rejects(executeArgv(["entra", "organization", "branding", "show",
        "--organization", orgId, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /no default branding is configured/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "organization", "branding-localization", "show",
        "--organization", orgId, "--id", "xx-XX", "--profile", profile],
        { ...overrides, transport: organizationTransport().send }), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /no branding localization xx-XX/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied organization reads surface scope, roles and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const overrides = {
        transport: denied.send,
        delegated: credentialService("delegated", []),
        application: credentialService("application", []),
      };
      await assert.rejects(executeArgv(["entra", "organization", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Organization.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Directory Readers and Global Reader")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "organization", "branding", "show",
        "--organization", orgId, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("OrganizationalBranding.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Global Reader or Organizational Branding Administrator")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists and shows organizations, branding and localizations`, () => {
    const state = setupProfiles();
    try {
      const listed = runOrganizationCli(["entra", "organization", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const organizationsOut = decode(listed.stdout);
      assert.deepEqual(organizationsOut.organizations.map(row => row.id), [org1.id, org2.id]);
      assert.deepEqual(organizationsOut.count, { returned: 2, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runOrganizationCli(["entra", "organization", "show", "--id", orgId, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).organization, org1);

      const branded = runOrganizationCli(["entra", "organization", "branding", "show",
        "--organization", orgId, "--profile", profile], state, mode);
      assert.equal(branded.status, 0, branded.stdout);
      assert.deepEqual(decode(branded.stdout).branding, branding);

      const locales = runOrganizationCli(["entra", "organization", "branding-localization", "list",
        "--organization", orgId, "--profile", profile,
        "--select", "id,signInPageText,usernameHintText,backgroundColor"], state, mode);
      assert.equal(locales.status, 0, locales.stdout);
      assert.equal(decode(locales.stdout).brandingLocalizations.length, 2);

      const locale = runOrganizationCli(["entra", "organization", "branding-localization", "show",
        "--organization", orgId, "--id", locDefault.id, "--profile", profile,
        "--select", "id,signInPageText"], state, mode);
      assert.equal(locale.status, 0, locale.stdout);
      assert.deepEqual(decode(locale.stdout).brandingLocalization, { id: locDefault.id, signInPageText: "Contoso" });
      assert.ok(!locale.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied organization reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runOrganizationCli(["entra", "organization", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.organizations, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });
}

test("truncated localization text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const partial = await executeArgv(["entra", "organization", "branding-localization", "show",
      "--organization", orgId, "--id", locFr.id, "--profile", "soc",
      "--select", "id,signInPageText"], overrides);
    assert.match(partial.brandingLocalization.signInPageText, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "organization", "branding-localization", "show",
      "--organization", orgId, "--id", locFr.id, "--profile", "soc", "--full",
      "--select", "id,signInPageText"], overrides);
    assert.equal(full.brandingLocalization.signInPageText, longWelcome);
    assert.ok(!full.help || !full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("Stream image properties fail with a later-piece hint before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "organization", "branding", "show",
        "--organization", orgId, "--profile", "soc", "--select", "id,bannerLogo"], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /later piece/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "organization", "branding-localization", "show",
        "--organization", orgId, "--id", locFr.id, "--profile", "soc", "--fields", "backgroundImage"], overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "organization", "list", "--profile", "batch", "--scopes", orgScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "organization", "branding", "show",
        "--organization", orgId, "--profile", "batch", "--scopes", brandingScopes[0]], overrides),
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
    await assert.rejects(executeArgv(["entra", "organization", "list", "--profile", "soc", "--select", "id,owner"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "organization", "list", "--profile", "soc", "--fields", "city"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "organization", "branding-localization", "list",
      "--organization", orgId, "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "organization", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("organization read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "organization", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "organization", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "organization", "list", "--filter", "displayName eq 'Contoso'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "organization", "branding", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "organization", "branding", "show", "--organization="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "organization", "branding-localization", "list"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "organization", "branding-localization", "show",
    "--organization", orgId]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "organization", "branding-localization", "show",
    "--organization", orgId, "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "organization", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
});

test("organization reads stay preview-gated on beta", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "organization", "list", "--profile", "soc", "--api-version", "beta", "--limit", "1"], overrides),
      error => {
        assert.equal(error.code, "POLICY_DENIED");
        return /preview-enabled/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "organization", "branding", "show",
        "--organization", orgId, "--profile", "soc", "--api-version", "beta"], overrides),
      { code: "POLICY_DENIED" },
    );
  } finally {
    teardownProfiles(state);
  }
});
