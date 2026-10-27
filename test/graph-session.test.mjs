import assert from "node:assert/strict";
import { test } from "node:test";
import { readFileSync } from "node:fs";
import { ApplicationAuth } from "../dist/app-auth.js";
import { DelegatedAuth } from "../dist/auth.js";
import { authorizeUrl, GraphSession, resolveSessionOperation } from "../dist/graph-session.js";
import { listScopedMembers, showScopedMember } from "../dist/entra-roles.js";
import { listAuthContexts, showAuthStrengthPolicy } from "../dist/entra-conditional-access.js";
import { listLifecyclePolicies } from "../dist/entra-group-lifecycle.js";
import { listProvisioning } from "../dist/entra-audit-logs.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const audience = "https://graph.microsoft.com/.default";
const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, token: { stored: true } };
const appProfile = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, auth: "federated", token: { env: "MG_AXI_TOKEN" } };
const scopes = ["https://graph.microsoft.com/User.Read"];
const users = resolveSessionOperation("v1.0", "GET", "/users");
const userById = resolveSessionOperation("v1.0", "GET", "/users/{user-id}");
// Repository-only discovery rows exercise callers attempting to bypass the runtime index.
const discovery = new Map(JSON.parse(readFileSync(new URL("../inventory/operations.json", import.meta.url))).operations.map(row => [row.id, row]));
function discoveryOperation(version, method, path) {
  const operation = discovery.get(`${version}:${method}:${path}`);
  assert.ok(operation);
  return { ...operation };
}
const blockedDeviceCredentials = discoveryOperation("v1.0", "GET", "/directory/deviceLocalCredentials/{deviceLocalCredentialInfo-id}");
const me = discoveryOperation("v1.0", "GET", "/me");
const betaUsers = discoveryOperation("beta", "GET", "/users");
const revokeSessions = resolveSessionOperation("v1.0", "POST", "/users/{user-id}/revokeSignInSessions");

function json(status, body, headers = {}) {
  return { status, headers, body: JSON.stringify(body) };
}

function fixture(handler) {
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
  return { credentialCalls, requests, session: new GraphSession({ delegated, application, transport }) };
}

test("PIM reads send a supported culture through both request builders", async () => {
  const operation = resolveSessionOperation("v1.0", "GET", "/roleManagement/directory/roleEligibilityScheduleInstances");
  const f = fixture(request => request.headers["Accept-Language"] === "en-US"
    ? json(400, { error: { code: "AadPremiumLicenseRequired", message: "Premium license required" } })
    : json(400, { error: { code: "CultureNotFoundException", message: "* is an invalid culture identifier" } }));
  await assert.rejects(f.session.execute({ profile: appProfile, operation }), /AadPremiumLicenseRequired/);
  await assert.rejects(f.session.collect({ profile: appProfile, operation }), /AadPremiumLicenseRequired/);
});

test("PIM licence errors preserve the diagnosis without suggesting retries", async () => {
  const operation = resolveSessionOperation("v1.0", "GET", "/roleManagement/directory/roleEligibilityScheduleInstances");
  const message = "The tenant needs to have Microsoft Entra ID P2 or Microsoft Entra ID Governance license.";
  const f = fixture(json(400, { error: { code: "AadPremiumLicenseRequired", message } }));
  for (const method of ["execute", "collect"]) {
    await assert.rejects(f.session[method]({ profile: appProfile, operation }), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.match(error.message, /\(400\): graph AadPremiumLicenseRequired/);
      assert.ok(error.message.includes(message));
      assert.deepEqual(error.suggestions, []);
      return true;
    });
  }
  assert.equal(f.requests.length, 2);
});

test("Graph preserves a query rejection message when the error code is empty", async () => {
  const f = fixture(json(400, { error: { code: "", message: "Query option 'Select' is not allowed" } }));
  const operation = resolveSessionOperation("v1.0", "GET", "/identity/conditionalAccess/authenticationContextClassReferences");
  await assert.rejects(f.session.execute({ profile: appProfile, operation, query: { $select: "id,displayName" } }), /Query option 'Select' is not allowed/);
});

test("unsupported root collections fail before credentials with supported alternatives", async () => {
  const f = fixture(json(200, { value: [] }));
  for (const [path, hint] of [["/domainDnsRecords", /verification-dns-record/], ["/directoryObjects", /directory-object show/]]) {
    const operation = discoveryOperation("v1.0", "GET", path);
    await assert.rejects(f.session.collect({ profile: appProfile, operation }), error => {
      assert.match(error.message, /not supported/);
      assert.ok(error.suggestions.some(value => hint.test(value)));
      return true;
    });
  }
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("scoped membership reads select the v1.0 fields and refuse direct root queries", async () => {
  const row = { id: "membership", roleId: "role", administrativeUnitId: "unit", roleMemberInfo: { id: "member", displayName: "Member" } };
  const f = fixture(request => {
    const url = new URL(request.url);
    const select = url.searchParams.get("$select").split(",");
    if (select.some(field => !Object.hasOwn(row, field))) return json(400, { error: { code: "Request_BadRequest", message: "Property does not exist" } });
    return json(200, url.pathname.endsWith("/scopedMembers") ? { value: [row] } : row);
  });
  const list = resolveSessionOperation("v1.0", "GET", "/directoryRoles/{directoryRole-id}/scopedMembers");
  const show = resolveSessionOperation("v1.0", "GET", "/directoryRoles/{directoryRole-id}/scopedMembers/{scopedRoleMembership-id}");
  assert.deepEqual((await listScopedMembers(f.session, { role: "role" }, appProfile, list, "help", "batch")).scopedMembers, [row]);
  assert.deepEqual((await showScopedMember(f.session, { role: "role", id: "membership" }, appProfile, show, "help", "batch")).scopedMember, row);
  for (const path of ["/scopedRoleMemberships", "/scopedRoleMemberships/{scopedRoleMembership-id}"]) {
    await assert.rejects(f.session.execute({ profile: appProfile, operation: discoveryOperation("v1.0", "GET", path), params: { "scopedRoleMembership-id": "membership" } }), /not supported/);
  }
  assert.equal(f.requests.length, 2);
});

test("select-rejecting policy reads project compact fields without sending $select", async () => {
  for (const [run, path, flags, body, key, expected] of [
    [listAuthContexts, "/identity/conditionalAccess/authenticationContextClassReferences", {}, { value: [{ id: "c1", displayName: "Context", description: "Detail" }] }, "authContexts", [{ id: "c1", displayName: "Context" }]],
    [showAuthStrengthPolicy, "/identity/conditionalAccess/authenticationStrength/policies/{authenticationStrengthPolicy-id}", { id: "strength" }, { id: "strength", displayName: "Strength" }, "authStrengthPolicy", { id: "strength", displayName: "Strength" }],
    [listLifecyclePolicies, "/groupLifecyclePolicies", {}, { value: [{ id: "policy", groupLifetimeInDays: 180, managedGroupTypes: "All", alternateNotificationEmails: "admin@example.com" }] }, "groupLifecyclePolicies", [{ id: "policy", groupLifetimeInDays: 180, managedGroupTypes: "All" }]],
  ]) {
    const f = fixture(request => new URL(request.url).searchParams.has("$select")
      ? json(400, { error: { code: "", message: "Query option 'Select' is not allowed" } })
      : json(200, body));
    const result = await run(f.session, flags, appProfile, resolveSessionOperation("v1.0", "GET", path), "help", "batch");
    assert.deepEqual(result[key], expected);
    assert.equal(f.requests.length, 1);
  }
});

test("provisioning list keeps its time bound and projects provisioningAction without $select", async () => {
  const row = { id: "log", activityDateTime: "2026-09-01T01:00:00Z", provisioningAction: "create", provisioningStatusInfo: { status: "success" }, jobId: "job" };
  const f = fixture(request => new URL(request.url).searchParams.has("$select")
    ? json(400, { error: { code: "", message: "Query option 'Select' is not allowed" } })
    : json(200, { value: [row] }));
  const operation = resolveSessionOperation("v1.0", "GET", "/auditLogs/provisioning");
  const result = await listProvisioning(f.session, { since: "2026-09-01T00:00:00Z" }, appProfile, operation, "help", "batch");
  assert.deepEqual(result.provisioning, [{ id: row.id, activityDateTime: row.activityDateTime, provisioningAction: "create", provisioningStatusInfo: row.provisioningStatusInfo }]);
  assert.equal(new URL(f.requests[0].url).searchParams.get("$filter"), "activityDateTime ge 2026-09-01T00:00:00.000Z");
});

test("delegated read returns parsed JSON through the authorized URL", async () => {
  const f = fixture(json(200, { value: [{ id: "a" }] }));
  const result = await f.session.execute({ profile: delegatedProfile, operation: users, query: { $top: "2" }, scopes });
  assert.deepEqual(result, { value: [{ id: "a" }] });
  assert.equal(f.requests.length, 1);
  assert.equal(f.requests[0].url, "https://graph.microsoft.com/v1.0/users?%24top=2");
  assert.equal(f.requests[0].headers.Authorization, "Bearer opaque-fixture-secret");
  assert.equal(f.requests[0].headers.Accept, "application/json");
  assert.match(f.requests[0].headers["client-request-id"], /^[0-9a-f-]{36}$/);
  assert.equal(f.credentialCalls.length, 1);
  assert.equal(f.credentialCalls[0][0], "silent");
});

for (const [profile, scopeArgs] of [[delegatedProfile, { scopes }], [appProfile, {}]]) {
  for (const [name, body, expected] of [
    ["root string", "AccountKey=fixture-secret", "***redacted***"],
    ["root array", ["SharedAccessKey=fixture-secret", "ordinary", 42, null], ["***redacted***", "ordinary", 42, null]],
    ["provider secrets", { id: "provider", certificateData: "YmFzZTY0LWtleQ==", clientSecret: "ordinary-secret", nested: { certificateData: ["key-material"], clientSecret: { value: "secret" }, clientId: "provider-client" } },
      { id: "provider", nested: { clientId: "provider-client" } }],
    ["nested object", { value: [{ id: "a", displayName: "AccountKey=fixture-secret", details: { password: "fixture-secret", enabled: true } }] },
      { value: [{ id: "a", displayName: "***redacted***", details: { password: "***redacted***", enabled: true } }] }],
    ["public branding metadata", { customAccountResetCredentialsUrl: "https://contoso.com/reset", customForgotMyPasswordText: "Forgot your password?", password: "fixture-secret", customForgotMyPasswordTextSecret: "fixture-secret" },
      { customAccountResetCredentialsUrl: "https://contoso.com/reset", customForgotMyPasswordText: "Forgot your password?", password: "***redacted***", customForgotMyPasswordTextSecret: "***redacted***" }],
    ["branding metadata with secret values", { value: [{ customAccountResetCredentialsUrl: "https://contoso.com/reset?sig=fixture-secret", customForgotMyPasswordText: "AccountKey=fixture-secret" }] },
      { value: [{ customAccountResetCredentialsUrl: "***redacted***", customForgotMyPasswordText: "***redacted***" }] }],
  ]) test(`${profile.mode} success redacts sentinels in a ${name}`, async () => {
    const f = fixture(json(200, body));
    assert.deepEqual(await f.session.execute({ profile, operation: users, ...scopeArgs }), expected);
  });

  test(`${profile.mode} redirected success redacts sentinels`, async () => {
    const f = fixture((request, count) => count === 1
      ? { status: 302, headers: { location: "/v1.0/users?$top=2" }, body: "" }
      : json(200, { value: [{ id: "a", displayName: "AccountKey=fixture-secret" }] }));
    assert.deepEqual(await f.session.execute({ profile, operation: users, ...scopeArgs }), {
      value: [{ id: "a", displayName: "***redacted***" }],
    });
    assert.equal(f.requests.length, 2);
  });
}

for (const [profile, scopeArgs] of [[delegatedProfile, { scopes: ["https://graph.microsoft.com/IdentityProvider.Read.All"] }], [appProfile, {}]]) {
  const operation = resolveSessionOperation("v1.0", "GET", "/identity/identityProviders");
  const rows = [{ id: "Google", clientSecret: "ordinary-secret" }, { id: "buffered-provider", certificateData: "YmFzZTY0LWtleQ==", clientId: "provider-client" }, { id: "builtin" }];

  test(`${profile.mode} provider cursor excludes secrets from capped pages`, async () => {
    const f = fixture(json(200, { value: rows }));
    const first = await f.session.collect({ profile, operation, ...scopeArgs, limit: 1 });
    assert.deepEqual(first.value, [{ id: "Google" }]);
    const cursor = JSON.parse(Buffer.from(first.cursor, "base64url").toString("utf8"));
    assert.deepEqual(cursor.buffered, [{ id: "buffered-provider", clientId: "provider-client" }, { id: "builtin" }]);
    const resumed = await f.session.collect({ profile, operation, ...scopeArgs, cursor: first.cursor });
    assert.deepEqual(resumed.value, [{ id: "buffered-provider", clientId: "provider-client" }, { id: "builtin" }]);
    assert.equal(resumed.complete, true);
    assert.equal(f.requests.length, 1);
  });

  test(`${profile.mode} provider cursor scrubs legacy buffered secrets before returning or reserializing`, async () => {
    const f = fixture(json(200, { value: rows }));
    const first = await f.session.collect({ profile, operation, ...scopeArgs, limit: 1 });
    const cursor = JSON.parse(Buffer.from(first.cursor, "base64url").toString("utf8"));
    cursor.buffered = [rows[1], rows[0]];
    const legacyCursor = Buffer.from(JSON.stringify(cursor)).toString("base64url");
    const resumed = await f.session.collect({ profile, operation, ...scopeArgs, cursor: legacyCursor, limit: 1 });
    assert.deepEqual(resumed.value, [{ id: "buffered-provider", clientId: "provider-client" }]);
    const next = JSON.parse(Buffer.from(resumed.cursor, "base64url").toString("utf8"));
    assert.deepEqual(next.buffered, [{ id: "Google" }]);
    assert.equal(f.requests.length, 1);
  });
}

test("path parameters bind one encoded resource per placeholder", async () => {
  const f = fixture(json(200, { id: "a" }));
  await f.session.execute({ profile: delegatedProfile, operation: userById, params: { "user-id": "analyst@example.invalid" }, scopes });
  assert.equal(f.requests[0].url, "https://graph.microsoft.com/v1.0/users/analyst%40example.invalid");
});

for (const [profile, scopeArgs] of [[delegatedProfile, { scopes }], [appProfile, {}]]) {
  for (const id of ["$count", "$value", "$ref", "$custom", "%24count", "delta()", "DELTA()", "delta(", "delta)", "delta%28%29"]) {
    for (const [method, operation, params] of [
      ["execute", userById, { "user-id": id }],
      ["collect", resolveSessionOperation("v1.0", "GET", "/groups/{group-id}/members"), { "group-id": id }],
    ]) test(`${profile.mode} ${method} rejects unsafe resource binding ${id} before credentials`, async () => {
      const f = fixture(json(200, {}));
      await assert.rejects(f.session[method]({ profile, operation, params, ...scopeArgs }), { code: "VALIDATION_ERROR" });
      assert.equal(f.credentialCalls.length, 0);
      assert.equal(f.requests.length, 0);
    });
  }
}

for (const id of ["$count", "delta()", "DELTA()", "delta(", "delta)"]) {
  test(`continuation authorization rejects unsafe resource binding ${id}`, () => {
    assert.throws(() => authorizeUrl(userById, { "user-id": id }, `/v1.0/users/${encodeURIComponent(id)}`), { code: "VALIDATION_ERROR" });
  });
}

test("guest UPN remains bound across a same-resource redirect", async () => {
  const user = "alice_example.com#EXT#@tenant.onmicrosoft.com";
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(user)}`;
  const redirectedUrl = `${url}?$select=id`;
  const f = fixture((request, count) => count === 1
    ? { status: 302, headers: { location: redirectedUrl }, body: "" }
    : json(200, { id: user }));
  const result = await f.session.execute({ profile: delegatedProfile, operation: userById, params: { "user-id": user }, scopes });
  assert.deepEqual(result, { id: user });
  assert.deepEqual(f.requests.map(request => request.url), [url, redirectedUrl]);
});

for (const [profile, scopeArgs] of [[delegatedProfile, { scopes }], [appProfile, {}]]) {
  for (const params of [{}, { appId: client }]) test(`${profile.mode} embedded path templates with ${Object.keys(params).length} bindings fail before credentials`, async () => {
    const operation = discoveryOperation("beta", "GET", "/applications(appId='{appId}')");
    const f = fixture(json(200, {}));
    await assert.rejects(
      f.session.execute({ profile: { ...profile, preview: true }, operation, params, ...scopeArgs }),
      error => error.code === "VALIDATION_ERROR" && /Unknown catalogued/.test(error.message),
    );
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
  });
}

for (const target of ["/v1.0/applications(appId='{appId}')", `/v1.0/applications(appId='${client}')`]) test(`embedded-template continuation ${target} is denied`, () => {
  const operation = discoveryOperation("v1.0", "GET", "/applications(appId='{appId}')");
  assert.throws(() => authorizeUrl(operation, { appId: client }, target), { code: "POLICY_DENIED" });
});

for (const method of ["execute", "collect"]) test(`${method} refuses uncatalogued operation ids before credential or HTTP`, async () => {
  assert.throws(() => resolveSessionOperation("v1.0", "GET", "/nope"), { code: "VALIDATION_ERROR" });
  const f = fixture(json(200, {}));
  await assert.rejects(
    f.session[method]({ profile: delegatedProfile, operation: { id: "v1.0:GET:/nope", method: "GET", path: "/nope", version: "v1.0", disposition: "scheduled", owningSlice: "READ-01" }, scopes }),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("a fabricated executable disposition cannot widen a blocked operation", async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: { ...blockedDeviceCredentials, disposition: "scheduled" }, params: { "deviceLocalCredentialInfo-id": "fixture-id" }, scopes }),
    error => error.code === "VALIDATION_ERROR" && /Unknown catalogued/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("mutating a resolved operation cannot change the authoritative inventory", async () => {
  const operation = resolveSessionOperation("v1.0", "GET", "/users");
  operation.disposition = "scheduled";
  operation.path = blockedDeviceCredentials.path;
  const f = fixture(json(200, {}));
  await f.session.execute({ profile: delegatedProfile, operation, scopes });
  assert.deepEqual(resolveSessionOperation("v1.0", "GET", "/users"), users);
  assert.equal(f.requests[0].url, "https://graph.microsoft.com/v1.0/users");
});

for (const [name, operation, params] of [
  ["intentionally-blocked route", blockedDeviceCredentials, { "deviceLocalCredentialInfo-id": "fixture-id" }],
  ["excluded route", discoveryOperation("v1.0", "GET", "/me/messages/{message-id}/$value"), { "message-id": "fixture-message" }],
]) test(`${name} is denied with no credential and no HTTP call`, async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(f.session.execute({ profile: delegatedProfile, operation, params, scopes }), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("disabled pack is denied before credential or HTTP", async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(
    f.session.execute({ profile: { ...delegatedProfile, enabledPacks: [] }, operation: users, scopes }),
    error => error.code === "POLICY_DENIED" && /entra/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const [area, id] of [
  ["mail", "v1.0:GET:/me/messages"],
  ["files", "v1.0:GET:/drives"],
]) test(`sensitive ${area} area is denied without enablement`, async () => {
  const [version, method, ...route] = id.split(":");
  const operation = discoveryOperation(version, method, route.join(":"));
  const f = fixture(json(200, {}));
  await assert.rejects(f.session.execute({ profile: delegatedProfile, operation, scopes }), { code: area === "mail" ? "POLICY_DENIED" : "VALIDATION_ERROR" });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("unshipped beta reads stay absent even with a preview-enabled profile", async () => {
  const denied = fixture(json(200, {}));
  await assert.rejects(denied.session.execute({ profile: delegatedProfile, operation: betaUsers, scopes }), { code: "VALIDATION_ERROR" });
  assert.equal(denied.credentialCalls.length, 0);
  assert.equal(denied.requests.length, 0);
  const allowed = fixture(json(200, { value: [] }));
  await assert.rejects(allowed.session.execute({ profile: { ...delegatedProfile, preview: true }, operation: betaUsers, scopes }), { code: "VALIDATION_ERROR" });
  assert.equal(allowed.credentialCalls.length, 0);
  assert.equal(allowed.requests.length, 0);
});

for (const [version, resource] of [["beta", "accessPackageResources"], ["v1.0", "resources"]]) {
  for (const [suffix, fileParams] of [
    ["", {}],
    ["/$count", {}],
    ["/{customDataProvidedResourceFile-id}", { "customDataProvidedResourceFile-id": "file-id" }],
    ["/{customDataProvidedResourceFile-id}/$value", { "customDataProvidedResourceFile-id": "file-id" }],
  ]) {
    const operation = discoveryOperation(version, "GET", `/identityGovernance/entitlementManagement/${resource}/{accessPackageResource-id}/uploadSessions/{customDataProvidedResourceUploadSession-id}/files${suffix}`);
    for (const [profile, scopeArgs] of [[delegatedProfile, { scopes }], [appProfile, {}]]) test(`${profile.mode} governance file route ${operation.id} is denied before credentials`, async () => {
      const f = fixture(json(200, {}));
      await assert.rejects(
        f.session.execute({
          profile: { ...profile, preview: true }, operation,
          params: {
            "accessPackageResource-id": "resource-id",
            "customDataProvidedResourceUploadSession-id": "upload-id",
            ...fileParams,
          },
          ...scopeArgs,
        }),
        error => error.code === "VALIDATION_ERROR" && /Unknown catalogued/.test(error.message),
      );
      assert.equal(f.credentialCalls.length, 0);
      assert.equal(f.requests.length, 0);
    });
  }
}

for (const version of ["beta", "v1.0"]) {
  for (const root of ["/agreements", "/identityGovernance/termsOfUse/agreements"]) {
    for (const [suffix, params] of [
      ["/file", { "agreement-id": "agreement-id" }],
      ["/file/localizations/{agreementFileLocalization-id}/versions/{agreementFileVersion-id}", {
        "agreement-id": "agreement-id", "agreementFileLocalization-id": "localization-id", "agreementFileVersion-id": "version-id",
      }],
    ]) {
      const operation = discoveryOperation(version, "GET", `${root}/{agreement-id}${suffix}`);
      for (const [profile, scopeArgs] of [[delegatedProfile, { scopes }], [appProfile, {}]]) test(`${profile.mode} agreement file route ${operation.id} is denied before credentials`, async () => {
        const f = fixture(json(200, {}));
        await assert.rejects(
          f.session.execute({ profile: { ...profile, preview: true }, operation, params, ...scopeArgs }),
          error => error.code === "VALIDATION_ERROR" && /Unknown catalogued/.test(error.message),
        );
        assert.equal(f.credentialCalls.length, 0);
        assert.equal(f.requests.length, 0);
      });
    }
  }
}

test("application profiles cannot use /me", async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(f.session.execute({ profile: appProfile, operation: me }), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("application reads use the configured .default audience", async () => {
  const f = fixture(json(200, { value: [] }));
  await f.session.execute({ profile: appProfile, operation: users });
  assert.equal(f.requests[0].headers.Authorization, "Bearer opaque-fixture-secret");
  assert.deepEqual(f.credentialCalls, [["acquire", { ...appProfile }, [audience]]]);
});

test("delegated reads require explicit scopes before HTTP", async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(f.session.execute({ profile: delegatedProfile, operation: users }), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const method of ["execute", "collect"]) {
  for (const scope of [
    "https://graph.microsoft.com/Directory.ReadWrite.All",
    "https://graph.microsoft.com/Application.ReadWrite.All",
    "https://graph.microsoft.com/Policy.ReadWrite.ConditionalAccess",
    "https://graph.microsoft.com/Mail.Write.All",
    "https://graph.microsoft.com/User.EnableDisableAccount.All",
    "https://graph.microsoft.com/Mail.Read.Shared",
    "https://example.invalid/User.Read.All",
    "User.Read.All",
  ]) test(`${method} rejects unsupported read scope ${scope} before credentials`, async () => {
    const f = fixture(json(200, { value: [] }));
    await assert.rejects(f.session[method]({ profile: delegatedProfile, operation: users, scopes: [...scopes, scope] }), error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      assert.match(error.suggestions[0], /Supported read scopes:.*https:\/\/graph\.microsoft\.com\/Directory\.Read\.All/);
      assert.match(error.suggestions[0], /https:\/\/graph\.microsoft\.com\/Application\.Read\.All/);
      return true;
    });
    assert.deepEqual(f.credentialCalls, []);
    assert.deepEqual(f.requests, []);
  });

  test(`${method} accepts documented read scope combinations`, async () => {
    const f = fixture(json(200, { value: [] }));
    const requested = [
      "https://graph.microsoft.com/User.ReadBasic.All",
      "https://graph.microsoft.com/GroupMember.ReadBasic.All",
      "https://graph.microsoft.com/Member.Read.Hidden",
      "https://graph.microsoft.com/Policy.Read.ConditionalAccess",
    ];
    await f.session[method]({ profile: delegatedProfile, operation: users, scopes: requested });
    assert.deepEqual(f.credentialCalls[0][2], [...requested].sort());
    assert.equal(f.requests.length, 1);
  });
}

test("application reads reject caller scopes before credential or HTTP", async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(f.session.execute({ profile: appProfile, operation: users, scopes }), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("mutations are rejected before credential or HTTP", async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: revokeSessions, params: { "user-id": "fixture-id" }, scopes }),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("supported query keys pass through; unknown keys fail before credential or HTTP", async () => {
  const allowed = fixture(json(200, { value: [] }));
  await allowed.session.execute({ profile: delegatedProfile, operation: users, query: { $select: "id,displayName", $top: "5" }, scopes });
  assert.match(allowed.requests[0].url, /%24select=id%2CdisplayName/);
  const denied = fixture(json(200, {}));
  await assert.rejects(denied.session.execute({ profile: delegatedProfile, operation: users, query: { $bogus: "1" }, scopes }), { code: "VALIDATION_ERROR" });
  assert.equal(denied.credentialCalls.length, 0);
  assert.equal(denied.requests.length, 0);
});

for (const value of ["a/b", "..", "a%40b", "", "a?b", "a b", "a\\b"]) test(`hostile path parameter ${JSON.stringify(value)} is rejected before credential or HTTP`, async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(f.session.execute({ profile: delegatedProfile, operation: userById, params: { "user-id": value }, scopes }), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("same-route redirect re-attaches the credential to the final response", async () => {
  const f = fixture((request, count) => count === 1
    ? { status: 302, headers: { location: "/v1.0/users?$top=2" }, body: "" }
    : json(200, { value: [{ id: "a" }] }));
  const result = await f.session.execute({ profile: delegatedProfile, operation: users, query: { $top: "2" }, scopes });
  assert.deepEqual(result, { value: [{ id: "a" }] });
  assert.equal(f.requests.length, 2);
  assert.ok(f.requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-secret"));
  assert.equal(f.requests[1].url, "https://graph.microsoft.com/v1.0/users?$top=2");
});

test("redirect to a blocked route sends no credential and no second request", async () => {
  const f = fixture({ status: 302, headers: { location: "https://graph.microsoft.com/v1.0/directory/deviceLocalCredentials/fixture-id" }, body: "" });
  await assert.rejects(f.session.execute({ profile: delegatedProfile, operation: users, scopes }), { code: "POLICY_DENIED" });
  assert.equal(f.requests.length, 1);
  assert.ok(f.requests.every(request => !request.url.includes("deviceLocalCredentials")));
});

test("redirect off the commercial host is denied after one request", async () => {
  const f = fixture({ status: 302, headers: { location: "https://login.microsoftonline.com/v1.0/users" }, body: "" });
  await assert.rejects(f.session.execute({ profile: delegatedProfile, operation: users, scopes }), { code: "POLICY_DENIED" });
  assert.equal(f.requests.length, 1);
});

test("redirect swapping the bound resource is denied after one request", async () => {
  const f = fixture((request, count) => count === 1
    ? { status: 302, headers: { location: "https://graph.microsoft.com/v1.0/users/other-id" }, body: "" }
    : json(200, {}));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: userById, params: { "user-id": "fixture-id" }, scopes }),
    { code: "POLICY_DENIED" },
  );
  assert.equal(f.requests.length, 1);
});

test("redirects stop after the hop budget", async () => {
  const f = fixture((request, count) => ({ status: 302, headers: { location: `/v1.0/users?$top=${count}` }, body: "" }));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: users, scopes }),
    error => error.code === "POLICY_DENIED" && /exceeds 3 hops/.test(error.message),
  );
  assert.equal(f.requests.length, 4);
});

for (const [status, body, code, pattern] of [
  [401, { error: { code: "InvalidAuthenticationToken", message: "Access token has expired" } }, "AUTH_REQUIRED", /rejected the credential/],
  [403, { error: { code: "Authorization_RequestDenied", message: "Insufficient privileges" } }, "GRAPH_ERROR", /grant, role, licence or policy/],
  [404, { error: { code: "Request_ResourceNotFound", message: "Resource not found" } }, "GRAPH_ERROR", /not found or inaccessible/],
  [429, { error: { code: "TooManyRequests", message: "Throttled" } }, "GRAPH_ERROR", /retry after 120/],
  [500, { error: { code: "ServiceUnavailable", message: "Boom" } }, "GRAPH_ERROR", /ServiceUnavailable/],
]) test(`${status} translates without leaking the credential`, async () => {
  const headers = status === 429 ? { "retry-after": "120" } : {};
  const f = fixture(json(status, body, headers));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: users, scopes }),
    error => error.code === code && pattern.test(error.message) && !`${error.message} ${JSON.stringify(error)}`.includes("opaque-fixture-secret"),
  );
});

test("sentinel values are redacted from translated errors", async () => {
  const f = fixture(json(500, { error: { code: "Error", message: "boom eyJmaWtpZS1zZWNyZXQ.e30.e30", accessToken: "opaque-fixture-secret" } }));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: users, scopes }),
    error => {
      const text = `${error.message} ${JSON.stringify(error)}`;
      return error.code === "GRAPH_ERROR" && !text.includes("eyJmaWtpZS1zZWNyZXQ") && !text.includes("opaque-fixture-secret") && text.includes("***redacted***");
    },
  );
});

test("transport failure before a response becomes a credential-free error", async () => {
  const f = fixture(async () => { throw new Error("socket hang up: opaque-fixture-secret"); });
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: users, scopes }),
    error => error.code === "GRAPH_ERROR" && !`${error.message} ${JSON.stringify(error)}`.includes("opaque-fixture-secret") && !`${error.message} ${JSON.stringify(error)}`.includes("socket hang up"),
  );
});

test("204 returns null; empty 200 fails closed", async () => {
  const empty = fixture({ status: 204, headers: {}, body: "" });
  assert.equal(await empty.session.execute({ profile: delegatedProfile, operation: users, scopes }), null);
  const malformed = fixture({ status: 200, headers: {}, body: "" });
  await assert.rejects(malformed.session.execute({ profile: delegatedProfile, operation: users, scopes }), { code: "GRAPH_ERROR" });
});

for (const [name, target, ok] of [
  ["nextLink with skiptoken", "https://graph.microsoft.com/v1.0/users?$skiptoken=abc123&$top=5", true],
  ["off-host continuation", "https://example.invalid/v1.0/users", false],
  ["version switch", "https://graph.microsoft.com/beta/users", false],
  ["route switch", "https://graph.microsoft.com/v1.0/groups", false],
  ["unsupported query", "https://graph.microsoft.com/v1.0/users?$bogus=1", false],
]) test(`continuation authorization ${name} ${ok ? "passes" : "is denied"}`, () => {
  if (ok) assert.equal(authorizeUrl(users, {}, target), "https://graph.microsoft.com/v1.0/users?$skiptoken=abc123&$top=5");
  else assert.throws(() => authorizeUrl(users, {}, target), { code: "POLICY_DENIED" });
});

test("continuation authorization keeps the bound resource", () => {
  assert.equal(
    authorizeUrl(userById, { "user-id": "Fixture-ID" }, "https://graph.microsoft.com/v1.0/users/fixture-id"),
    "https://graph.microsoft.com/v1.0/users/fixture-id",
  );
  assert.throws(() => authorizeUrl(userById, { "user-id": "fixture-id" }, "https://graph.microsoft.com/v1.0/users/other-id"), { code: "POLICY_DENIED" });
});

// EXT-04e: allowlisted function-argument binding. Only the three catalogue
// operations bind parenthesised arguments; values arrive as explicit CLI
// arguments and the session validates, OData-quotes and encodes them before
// credentials. Every other parenthesised template keeps the refusal.
const tenantByDomain = resolveSessionOperation("v1.0", "GET", "/tenantRelationships/findTenantInformationByDomainName(domainName='{domainName}')");
const tenantById = resolveSessionOperation("v1.0", "GET", "/tenantRelationships/findTenantInformationByTenantId(tenantId='{tenantId}')");
const subscriptionByCommerceId = resolveSessionOperation("v1.0", "GET", "/directory/subscriptions(commerceSubscriptionId='{commerceSubscriptionId}')");
const tenantLookupScopes = ["https://graph.microsoft.com/CrossTenantInformation.ReadBasic.All"];
const commerceScopes = ["https://graph.microsoft.com/Organization.Read.All"];
const lookupTenantId = "55555555-5555-4555-8555-555555555555";
const lookupCommerceId = "66666666-6666-4666-8666-666666666666";

const functionCases = [
  {
    name: "domain-name lookup",
    operation: tenantByDomain,
    params: { domainName: "example.invalid" },
    canonical: "https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName=%27example.invalid%27)",
    scopes: tenantLookupScopes,
    bad: ["x') OR '1'='1", "')", "a/b", "a%2Fb", "$batch", "..", "a\nb", "a b", "a?b", "a#b", "", "x".repeat(300), "not_a_domain!", "-bad-start.invalid", "bad-end-.invalid", "a..invalid", ".invalid", "invalid."],
  },
  {
    name: "tenant-ID lookup",
    operation: tenantById,
    params: { tenantId: lookupTenantId },
    canonical: `https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByTenantId(tenantId=%27${lookupTenantId}%27)`,
    scopes: tenantLookupScopes,
    bad: ["example.invalid", "x') OR '1'='1", "')", "a/b", "a%2Fb", "$batch", "..", "a\nb", "zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz", "5555555545554555855555555555555", "", "x".repeat(300)],
  },
  {
    name: "commerce-key lookup",
    operation: subscriptionByCommerceId,
    params: { commerceSubscriptionId: lookupCommerceId },
    canonical: `https://graph.microsoft.com/v1.0/directory/subscriptions(commerceSubscriptionId=%27${lookupCommerceId}%27)`,
    scopes: commerceScopes,
    bad: ["x') OR '1'='1", "')", "a/b", "a%2Fb", "$batch", "..", "a\nb", "a b", "", "x".repeat(65), "a+b", "a.b", "a:b"],
  },
];

for (const { name, operation, params, canonical, scopes: scopeArgs } of functionCases) {
  for (const [profile, args] of [[delegatedProfile, { scopes: scopeArgs }], [appProfile, {}]]) {
    test(`${profile.mode} binds the ${name} to its canonical function URL`, async () => {
      const f = fixture(json(200, { id: "a" }));
      await f.session.execute({ profile, operation, params, ...args });
      assert.equal(f.requests[0].url, canonical);
    });
  }
}

for (const { name, operation, params, scopes: scopeArgs, bad } of functionCases) {
  const key = Object.keys(params)[0];
  for (const value of bad) {
    test(`${name} refuses injection ${JSON.stringify(String(value).slice(0, 24))} before credentials`, async () => {
      const f = fixture(json(200, {}));
      await assert.rejects(
        f.session.execute({ profile: delegatedProfile, operation, params: { [key]: value }, scopes: scopeArgs }),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(f.credentialCalls.length, 0);
      assert.equal(f.requests.length, 0);
    });
  }
}

test("function routes refuse missing and unknown parameters before credentials", async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: tenantByDomain, params: {}, scopes: tenantLookupScopes }),
    error => error.code === "VALIDATION_ERROR" && /Missing path parameter/.test(error.message),
  );
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: tenantByDomain, params: { domainName: "example.invalid", tenantId: lookupTenantId }, scopes: tenantLookupScopes }),
    error => error.code === "VALIDATION_ERROR" && /Unknown path parameter/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("a non-allowlisted parenthesised operation is still refused before credentials", async () => {
  const operation = discoveryOperation("beta", "GET", "/servicePrincipals(appId='{appId}')");
  const f = fixture(json(200, {}));
  await assert.rejects(
    f.session.execute({ profile: { ...delegatedProfile, preview: true }, operation, params: { appId: "00000000-0000-4000-8000-000000000000" }, scopes: ["https://graph.microsoft.com/Application.Read.All"] }),
    error => error.code === "VALIDATION_ERROR" && /Unknown catalogued/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
  assert.throws(() => authorizeUrl(operation, { appId: "00000000-0000-4000-8000-000000000000" }, "/v1.0/servicePrincipals(appId='00000000-0000-4000-8000-000000000000')"), { code: "POLICY_DENIED" });
});

test("continuation authorization re-authorizes the bound function argument", () => {
  assert.equal(
    authorizeUrl(tenantByDomain, { domainName: "example.invalid" }, "https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName=%27example.invalid%27)"),
    "https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName=%27example.invalid%27)",
  );
  assert.throws(
    () => authorizeUrl(tenantByDomain, { domainName: "example.invalid" }, "https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName=%27other.invalid%27)"),
    { code: "POLICY_DENIED" },
  );
  assert.throws(
    () => authorizeUrl(tenantByDomain, { domainName: "a/b" }, "https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName=%27a%2Fb%27)"),
    { code: "VALIDATION_ERROR" },
  );
});
