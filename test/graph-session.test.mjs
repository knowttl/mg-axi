import assert from "node:assert/strict";
import { test } from "node:test";
import { ApplicationAuth } from "../dist/app-auth.js";
import { DelegatedAuth } from "../dist/auth.js";
import { authorizeUrl, GraphSession, resolveSessionOperation } from "../dist/graph-session.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const audience = "https://graph.microsoft.com/.default";
const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key } };
const appProfile = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "federated", key } };
const scopes = ["https://graph.microsoft.com/User.Read"];
const users = resolveSessionOperation("v1.0", "GET", "/users");
const userById = resolveSessionOperation("v1.0", "GET", "/users/{user-id}");
const blockedDeviceCredentials = resolveSessionOperation("v1.0", "GET", "/directory/deviceLocalCredentials/{deviceLocalCredentialInfo-id}");
const me = resolveSessionOperation("v1.0", "GET", "/me");
const betaUsers = resolveSessionOperation("beta", "GET", "/users");
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
  test(`${profile.mode} embedded path templates without bindings fail before credentials`, async () => {
    const operation = resolveSessionOperation("v1.0", "GET", "/applications(appId='{appId}')");
    const f = fixture(json(200, {}));
    await assert.rejects(
      f.session.execute({ profile, operation, params: {}, ...scopeArgs }),
      error => error.code === "VALIDATION_ERROR" && /Missing path parameter/.test(error.message),
    );
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
  });
  test(`${profile.mode} embedded path templates bind one function argument`, async () => {
    const operation = resolveSessionOperation("v1.0", "GET", "/applications(appId='{appId}')");
    const f = fixture(json(200, { id: "a" }));
    const result = await f.session.execute({ profile, operation, params: { appId: client }, ...scopeArgs });
    assert.deepEqual(result, { id: "a" });
    assert.equal(f.requests[0].url, `https://graph.microsoft.com/v1.0/applications(appId='${client}')`);
    assert.equal(f.credentialCalls.length, 1);
  });
}

test("embedded-template continuation with the bound resource passes", () => {
  const operation = resolveSessionOperation("v1.0", "GET", "/applications(appId='{appId}')");
  assert.equal(authorizeUrl(operation, { appId: client }, `/v1.0/applications(appId='${client}')`), `https://graph.microsoft.com/v1.0/applications(appId='${client}')`);
});

for (const target of ["/v1.0/applications(appId='{appId}')", "/v1.0/applications(appId='other-id')"]) test(`embedded-template continuation ${target} is denied`, () => {
  const operation = resolveSessionOperation("v1.0", "GET", "/applications(appId='{appId}')");
  assert.throws(() => authorizeUrl(operation, { appId: client }, target), { code: "POLICY_DENIED" });
});

test("uncatalogued operation ids fail before credential or HTTP", async () => {
  assert.throws(() => resolveSessionOperation("v1.0", "GET", "/nope"), { code: "VALIDATION_ERROR" });
  const f = fixture(json(200, {}));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: { id: "v1.0:GET:/nope", method: "GET", path: "/nope", version: "v1.0", disposition: "scheduled", owningSlice: "READ-01" }, scopes }),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("a fabricated executable disposition cannot widen a blocked operation", async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: { ...blockedDeviceCredentials, disposition: "scheduled" }, params: { "deviceLocalCredentialInfo-id": "fixture-id" }, scopes }),
    error => error.code === "POLICY_DENIED" && /intentionally-blocked/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("mutating a resolved operation cannot change the authoritative inventory", async () => {
  const operation = resolveSessionOperation("v1.0", "GET", blockedDeviceCredentials.path);
  operation.disposition = "scheduled";
  operation.path = "/users";
  const f = fixture(json(200, {}));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation, params: { "deviceLocalCredentialInfo-id": "fixture-id" }, scopes }),
    error => error.code === "POLICY_DENIED" && /intentionally-blocked/.test(error.message),
  );
  assert.deepEqual(resolveSessionOperation("v1.0", "GET", blockedDeviceCredentials.path), blockedDeviceCredentials);
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const [name, operation, params] of [
  ["intentionally-blocked route", blockedDeviceCredentials, { "deviceLocalCredentialInfo-id": "fixture-id" }],
  ["excluded route", resolveSessionOperation("v1.0", "GET", "/me/messages"), {}],
]) test(`${name} is denied with no credential and no HTTP call`, async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(f.session.execute({ profile: delegatedProfile, operation, params, scopes }), { code: "POLICY_DENIED" });
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
  const operation = resolveSessionOperation(version, method, route.join(":"));
  const f = fixture(json(200, {}));
  await assert.rejects(f.session.execute({ profile: delegatedProfile, operation, scopes }), error => error.code === "POLICY_DENIED" && new RegExp(area).test(error.message));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("beta reads require a preview-enabled profile", async () => {
  const denied = fixture(json(200, {}));
  await assert.rejects(denied.session.execute({ profile: delegatedProfile, operation: betaUsers, scopes }), { code: "POLICY_DENIED" });
  assert.equal(denied.credentialCalls.length, 0);
  assert.equal(denied.requests.length, 0);
  const allowed = fixture(json(200, { value: [] }));
  await allowed.session.execute({ profile: { ...delegatedProfile, preview: true }, operation: betaUsers, scopes });
  assert.equal(allowed.requests[0].url, "https://graph.microsoft.com/beta/users");
});

for (const [version, resource] of [["beta", "accessPackageResources"], ["v1.0", "resources"]]) {
  for (const [suffix, fileParams] of [
    ["", {}],
    ["/$count", {}],
    ["/{customDataProvidedResourceFile-id}", { "customDataProvidedResourceFile-id": "file-id" }],
    ["/{customDataProvidedResourceFile-id}/$value", { "customDataProvidedResourceFile-id": "file-id" }],
  ]) {
    const operation = resolveSessionOperation(version, "GET", `/identityGovernance/entitlementManagement/${resource}/{accessPackageResource-id}/uploadSessions/{customDataProvidedResourceUploadSession-id}/files${suffix}`);
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
        error => error.code === "POLICY_DENIED" && /Sensitive area files/.test(error.message),
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
      const operation = resolveSessionOperation(version, "GET", `${root}/{agreement-id}${suffix}`);
      for (const [profile, scopeArgs] of [[delegatedProfile, { scopes }], [appProfile, {}]]) test(`${profile.mode} agreement file route ${operation.id} is denied before credentials`, async () => {
        const f = fixture(json(200, {}));
        await assert.rejects(
          f.session.execute({ profile: { ...profile, preview: true }, operation, params, ...scopeArgs }),
          error => error.code === "POLICY_DENIED" && /Sensitive area files/.test(error.message),
        );
        assert.equal(f.credentialCalls.length, 0);
        assert.equal(f.requests.length, 0);
      });
    }
  }
}

test("application profiles cannot use /me", async () => {
  const f = fixture(json(200, {}));
  await assert.rejects(f.session.execute({ profile: appProfile, operation: me }), { code: "POLICY_DENIED" });
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
    "https://graph.microsoft.com/Mail.Read",
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
