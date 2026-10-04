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

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const appScopes = ["https://graph.microsoft.com/Application.Read.All"];

const a1 = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  appId: "99999999-9999-4999-8999-999999999999",
  displayName: "Contoso Web",
  signInAudience: "AzureADMyOrg",
  publisherDomain: "contoso.com",
  createdDateTime: "2024-01-01T00:00:00Z",
  keyCredentials: [
    {
      keyId: "key-1",
      displayName: "Signing key",
      startDateTime: "2024-01-01T00:00:00Z",
      endDateTime: "2026-01-01T00:00:00Z",
      key: "AQAB key material that must never surface",
      customKeyIdentifier: "b3BhcXVl",
    },
  ],
  passwordCredentials: [
    {
      keyId: "secret-1",
      displayName: "Client secret",
      startDateTime: "2024-06-01T00:00:00Z",
      endDateTime: "2025-06-01T00:00:00Z",
      hint: "abc",
      secretText: "super-secret-value that must never surface",
    },
  ],
};
const a2 = {
  id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  appId: "88888888-8888-4888-8888-888888888888",
  displayName: "Daemon Batch",
  signInAudience: "AzureADMyOrg",
  publisherDomain: "contoso.com",
  createdDateTime: "2024-02-01T00:00:00Z",
  keyCredentials: [],
  passwordCredentials: [],
};
const a3 = {
  id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  appId: "77777777-7777-4777-8777-777777777777",
  displayName: "Niche Tool",
};
const applications = [a1, a2, a3];

const s1 = {
  id: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
  appId: a1.appId,
  displayName: "Contoso Web",
  servicePrincipalType: "Application",
  accountEnabled: true,
  appOwnerOrganizationId: tenant,
  appRoleAssignmentRequired: false,
  preferredSingleSignOnMode: null,
  loginUrl: null,
};
const s2 = {
  id: "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee",
  appId: a2.appId,
  displayName: "Daemon Batch",
  servicePrincipalType: "Application",
  accountEnabled: true,
  appOwnerOrganizationId: tenant,
  appRoleAssignmentRequired: false,
  preferredSingleSignOnMode: null,
  loginUrl: null,
};
const s3 = {
  id: "ffffffff-ffff-4fff-8fff-ffffffffffff",
  appId: a3.appId,
  displayName: "Niche Tool",
  servicePrincipalType: "Application",
  accountEnabled: false,
  appOwnerOrganizationId: tenant,
  appRoleAssignmentRequired: false,
  preferredSingleSignOnMode: null,
  loginUrl: null,
};
const servicePrincipals = [s1, s2, s3];

const appOwners = [
  { "@odata.type": "#microsoft.graph.user", id: "11111111-1111-4111-8111-111111111112", displayName: "Adele Vance", mail: "adele@contoso.com" },
  { "@odata.type": "#microsoft.graph.servicePrincipal", id: s2.id, displayName: "Daemon Batch" },
];
const spOwners = [
  { "@odata.type": "#microsoft.graph.user", id: "11111111-1111-4111-8111-111111111112", displayName: "Adele Vance", mail: "adele@contoso.com" },
];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-07-"));
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

function appTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/applications") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [a3] });
      return json(200, {
        value: [a1, a2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/applications?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/servicePrincipals") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [s3] });
      return json(200, {
        value: [s1, s2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/servicePrincipals?%24skiptoken=page2",
      });
    }
    const appSingle = /^\/v1\.0\/applications\/([^/]+)$/.exec(path);
    if (appSingle) {
      const found = applications.find(row => row.id === decodeURIComponent(appSingle[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such application" } });
    }
    const spSingle = /^\/v1\.0\/servicePrincipals\/([^/]+)$/.exec(path);
    if (spSingle) {
      const found = servicePrincipals.find(row => row.id === decodeURIComponent(spSingle[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such service principal" } });
    }
    if (/^\/v1\.0\/applications\/[^/]+\/owners$/.test(path)) return json(200, { value: appOwners });
    if (/^\/v1\.0\/servicePrincipals\/[^/]+\/owners$/.test(path)) return json(200, { value: spOwners });
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? appTransport();
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

function runAppCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, directory: [], denied, scopes: appScopes[0], applications, servicePrincipals, appOwners, spOwners }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists applications keeping appId distinct from the object id`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "application", "list", "--profile", profile], overrides);
      assert.deepEqual(result.applications, [
        { id: a1.id, appId: a1.appId, displayName: "Contoso Web" },
        { id: a2.id, appId: a2.appId, displayName: "Daemon Batch" },
        { id: a3.id, appId: a3.appId, displayName: "Niche Tool" },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      for (const row of result.applications) assert.notEqual(row.id, row.appId);
      assert.ok(result.help.some(hint => hint.includes("entra application show --id <application-object-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/applications?"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows an application with expiry metadata and no secret material`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "application", "show", "--id", a1.id, "--profile", profile], overrides);
      assert.deepEqual(result.application, {
        id: a1.id,
        appId: a1.appId,
        displayName: "Contoso Web",
        signInAudience: "AzureADMyOrg",
        publisherDomain: "contoso.com",
        createdDateTime: "2024-01-01T00:00:00Z",
        keyCredentials: [
          { keyId: "key-1", displayName: "Signing key", startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2026-01-01T00:00:00Z" },
        ],
        passwordCredentials: [
          { keyId: "secret-1", displayName: "Client secret", startDateTime: "2024-06-01T00:00:00Z", endDateTime: "2025-06-01T00:00:00Z" },
        ],
      });
      const text = JSON.stringify(result);
      assert.ok(!text.includes("super-secret-value"));
      assert.ok(!text.includes("key material"));
      assert.ok(!text.includes("customKeyIdentifier"));
      assert.ok(!text.includes("abc"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists service principals keeping appId distinct from the object id`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "service-principal", "list", "--profile", profile], overrides);
      assert.deepEqual(result.servicePrincipals, [
        { id: s1.id, appId: s1.appId, displayName: "Contoso Web" },
        { id: s2.id, appId: s2.appId, displayName: "Daemon Batch" },
        { id: s3.id, appId: s3.appId, displayName: "Niche Tool" },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      for (const row of result.servicePrincipals) assert.notEqual(row.id, row.appId);
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/servicePrincipals?"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a service principal with richer properties`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "service-principal", "show", "--id", s1.id, "--profile", profile], overrides);
      assert.equal(result.servicePrincipal.id, s1.id);
      assert.equal(result.servicePrincipal.appId, a1.appId);
      assert.equal(result.servicePrincipal.servicePrincipalType, "Application");
      assert.equal(result.servicePrincipal.accountEnabled, true);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists application owners with kinds preserved`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "application", "owner", "list", "--application", a1.id, "--profile", profile], overrides);
      assert.deepEqual(result.owners, appOwners);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/applications/${a1.id}/owners`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists service-principal owners with kinds preserved`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "service-principal", "owner", "list", "--service-principal", s1.id, "--profile", profile], overrides);
      assert.deepEqual(result.owners, spOwners);
      assert.deepEqual(result.count, { returned: 1, complete: true });
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/servicePrincipals/${s1.id}/owners`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped application list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "application", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.applications.map(row => row.id), [a1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "application", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.applications.map(row => row.id), [a2.id, a3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  for (const [resource, key] of [["application", "applications"], ["service-principal", "servicePrincipals"]]) {
    for (const field of ["passwordCredentials", "keyCredentials"]) {
      test(`${mode} capped ${resource} ${field} retain only expiry metadata through resume`, async () => {
        const state = setupProfiles();
        try {
          const metadata = field === "passwordCredentials"
            ? { keyId: "secret-1", displayName: "Client secret", startDateTime: "2024-06-01T00:00:00Z", endDateTime: "2025-06-01T00:00:00Z" }
            : { keyId: "key-1", displayName: "Signing key", startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2026-01-01T00:00:00Z" };
          const rows = [a1.id, a2.id, a3.id].map(id => ({ id, [field]: a1[field] }));
          const fixture = transport(() => json(200, { value: rows }));
          const { overrides } = overridesFor(mode, fixture);
          const args = ["entra", resource, "list", "--profile", profile];
          const first = await executeArgv([...args, "--select", `id,${field}`, "--limit", "1"], overrides);
          assert.deepEqual(first[key], [{ id: a1.id, [field]: [metadata] }]);
          const cursor = JSON.parse(Buffer.from(first.cursor, "base64url").toString("utf8"));
          assert.deepEqual(cursor.buffered, rows.slice(1).map(row => ({ id: row.id, [field]: [metadata] })));
          cursor.buffered = rows.slice(1);
          const legacyCursor = Buffer.from(JSON.stringify(cursor)).toString("base64url");
          const second = await executeArgv([...args, "--cursor", legacyCursor, "--limit", "1"], overrides);
          assert.deepEqual(second[key], [{ id: a2.id, [field]: [metadata] }]);
          assert.deepEqual(JSON.parse(Buffer.from(second.cursor, "base64url").toString("utf8")).buffered,
            [{ id: a3.id, [field]: [metadata] }]);
          const third = await executeArgv([...args, "--cursor", second.cursor], overrides);
          assert.deepEqual(third[key], [{ id: a3.id, [field]: [metadata] }]);
          assert.deepEqual(third.count, { returned: 1, complete: true });
          assert.equal(fixture.requests.length, 1);
        } finally {
          teardownProfiles(state);
        }
      });
    }
  }

  test(`${mode} denied application reads surface as operational failures`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const missing = transport(() => json(404, { error: { code: "Request_ResourceNotFound", message: "gone" } }));
      const overrides = {
        transport: denied.send,
        delegated: credentialService("delegated", []),
        application: credentialService("application", []),
      };
      await assert.rejects(executeArgv(["entra", "application", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "application", "show", "--id", a1.id, "--profile", profile],
        { ...overrides, transport: missing.send }), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists applications, shows one and lists owners`, () => {
    const state = setupProfiles();
    try {
      const listed = runAppCli(["entra", "application", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listedOut = decode(listed.stdout);
      assert.deepEqual(listedOut.applications.map(row => row.id), [a1.id, a2.id, a3.id]);
      assert.deepEqual(listedOut.count, { returned: 3, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runAppCli(["entra", "application", "show", "--id", a1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      const shownOut = decode(shown.stdout);
      assert.equal(shownOut.application.appId, a1.appId);
      assert.equal(shownOut.application.passwordCredentials[0].keyId, "secret-1");
      assert.ok(!shown.stdout.includes("super-secret-value"));

      const owners = runAppCli(["entra", "application", "owner", "list", "--application", a1.id, "--profile", profile], state, mode);
      assert.equal(owners.status, 0, owners.stdout);
      assert.deepEqual(decode(owners.stdout).owners, appOwners);

      const sps = runAppCli(["entra", "service-principal", "list", "--profile", profile], state, mode);
      assert.equal(sps.status, 0, sps.stdout);
      assert.deepEqual(decode(sps.stdout).servicePrincipals.map(row => row.id), [s1.id, s2.id, s3.id]);

      const spShown = runAppCli(["entra", "service-principal", "show", "--id", s1.id, "--profile", profile], state, mode);
      assert.equal(spShown.status, 0, spShown.stdout);
      assert.equal(decode(spShown.stdout).servicePrincipal.servicePrincipalType, "Application");

      const spOwnersOut = runAppCli(["entra", "service-principal", "owner", "list", "--service-principal", s1.id, "--profile", profile], state, mode);
      assert.equal(spOwnersOut.status, 0, spOwnersOut.stdout);
      assert.deepEqual(decode(spOwnersOut.stdout).owners, spOwners);
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied application reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runAppCli(["entra", "application", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.applications, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });
}

test("application list $filter carries count and consistency through resume", async () => {
  const state = setupProfiles();
  try {
    const seen = [];
    const fixture = transport(request => {
      seen.push({ url: request.url, headers: request.headers });
      assert.equal(new URL(request.url).searchParams.get("$count"), "true");
      if (new URL(request.url).searchParams.has("$skiptoken")) {
        return json(200, { value: [{ id: a2.id, appId: a2.appId, displayName: "Daemon Batch" }] });
      }
      return json(200, {
        value: [{ id: a1.id, appId: a1.appId, displayName: "Contoso Web" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/applications?%24count=true&%24skiptoken=next",
      });
    });
    const { overrides } = overridesFor("delegated", fixture);
    const filtered = await executeArgv(["entra", "application", "list",
      "--profile", "soc", "--filter", "startswith(displayName,'C')", "--limit", "1"], overrides);
    assert.equal(seen[0].headers.ConsistencyLevel, "eventual");
    assert.equal(filtered.count.complete, false);
    const resumed = await executeArgv(["entra", "application", "list",
      "--profile", "soc", "--cursor", filtered.cursor], overrides);
    assert.equal(seen[1].headers.ConsistencyLevel, "eventual");
    assert.deepEqual(resumed.count.complete, true);
    await assert.rejects(executeArgv(["entra", "application", "list",
      "--profile", "soc", "--cursor", filtered.cursor, "--filter", "startswith(displayName,'D')"], overrides),
    { code: "VALIDATION_ERROR" });
  } finally {
    teardownProfiles(state);
  }
});

test("secret-minting routes are refused before credentials", async () => {
  const state = setupProfiles();
  try {
    const calls = [];
    const { requests, overrides } = overridesFor("delegated", undefined, calls);
    const { showApplication } = await import("../dist/entra-apps.js");
    const { operationFor, leafHelp, LEAVES } = await import("../dist/catalogue.js");
    const leaf = LEAVES.find(item => item.path === "entra application show");
    const secretOperation = {
      ...operationFor(leaf, "v1.0"),
      id: "v1.0:POST:/applications/{application-id}/addPassword",
      method: "POST",
      path: "/applications/{application-id}/addPassword",
    };
    const { GraphSession } = await import("../dist/graph-session.js");
    const session = new GraphSession(overrides);
    await assert.rejects(
      showApplication(session, { id: a1.id }, { mode: "delegated", tenantId: tenant, clientId: client }, secretOperation, leafHelp(leaf), "soc"),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /six catalogued application/.test(error.message);
      },
    );
    assert.equal(requests.length, 0);
    assert.equal(calls.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "application", "list", "--profile", "batch", "--scopes", appScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "service-principal", "owner", "list", "--service-principal", s1.id, "--profile", "batch", "--scopes", appScopes[0]], overrides),
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
    await assert.rejects(executeArgv(["entra", "application", "list", "--profile", "soc", "--select", "id,aboutMe"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "application", "list", "--profile", "soc", "--fields", "signInAudience"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "service-principal", "show", "--id", s1.id, "--profile", "soc", "--select", "id,hint"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "application", "owner", "list", "--application", a1.id, "--profile", "soc", "--select", "id,department"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "application", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("application relationship flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "application", "owner", "list"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "application", "owner", "list", "--application="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "service-principal", "owner", "list"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "service-principal", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "application", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "application", "list", "--transitive"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "application", "show", "--id", a1.id, "--group", a1.id]), { code: "VALIDATION_ERROR" });
});

test("application reads stay preview-gated on beta", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "application", "list", "--profile", "soc", "--api-version", "beta", "--limit", "1"], overrides),
      error => {
        assert.equal(error.code, "POLICY_DENIED");
        return /preview-enabled/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "service-principal", "owner", "list", "--service-principal", s1.id, "--profile", "soc", "--api-version", "beta"], overrides),
      { code: "POLICY_DENIED" },
    );
  } finally {
    teardownProfiles(state);
  }
});
