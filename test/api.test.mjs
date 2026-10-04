import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { decode } from "@toon-format/toon";
import { ApplicationAuth } from "../dist/app-auth.js";
import { DelegatedAuth } from "../dist/auth.js";
import { REVIEWED_ROUTES, fetchTransport, matchReviewed, runApiGet } from "../dist/api.js";

// API-01 through the lowest real interface observing behavior: runApiGet with
// real DelegatedAuth/ApplicationAuth and a fake credential/transport boundary
// (the same seam CORE-01/CORE-02 tests use), plus CLI subprocess checks for
// grammar and refusal precedence. CLI checks stop at validation: execution
// would need a real tenant, which offline verification never touches.

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key } };
const appProfile = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "federated", key } };
const scopes = "https://graph.microsoft.com/User.Read.All";

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
  return { credentialCalls, requests, deps: { delegated, application, transport } };
}

const read = (overrides = {}, handler = json(200, { value: [] })) => {
  const f = fixture(handler);
  return { ...f, run: args => runApiGet({ path: "/users", apiVersion: "v1.0", profile: delegatedProfile, scopes, ...overrides }, f.deps) };
};

test("every reviewed id binds a real inventory row", async () => {
  const { resolveSessionOperation } = await import("../dist/graph-session.js");
  assert.ok(REVIEWED_ROUTES.length > 0);
  for (const route of REVIEWED_ROUTES) {
    const [version, method, ...rest] = route.id.split(":");
    const operation = resolveSessionOperation(version, method, rest.join(":"));
    assert.equal(operation.id, route.id);
    assert.equal(operation.method, "GET");
  }
});

test("template matching is case-insensitive with single-segment bindings", () => {
  assert.deepEqual({ ...matchReviewed("/USERS/")?.params }, {});
  assert.deepEqual({ ...matchReviewed("/groups/abc/members")?.params }, { "group-id": "abc" });
  assert.equal(matchReviewed("/users/abc/extra/depth")?.route, undefined);
  assert.equal(matchReviewed("/me/messages")?.route, undefined);
});

test("unreviewed mail route is refused before credentials", async () => {
  const f = read({ path: "/me/messages" });
  await assert.rejects(f.run({}), error => error.code === "VALIDATION_ERROR" && /not in the reviewed raw inventory/.test(error.message));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("secret-value route stays refused before credentials", async () => {
  const f = read({ path: "/directory/deviceLocalCredentials/abc" });
  await assert.rejects(f.run({}), error => error.code === "VALIDATION_ERROR" && /never served raw/.test(error.suggestions.join("\n")));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const pathname of ["https://graph.microsoft.com/v1.0/users", "/users?$top=2", "/users#frag", "users", ""]) test(`host/query/path escape ${pathname || "(empty)"} is refused`, async () => {
  const f = read({ path: pathname });
  await assert.rejects(f.run({}), error => error.code === "VALIDATION_ERROR" && /Refused raw path/.test(error.message));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("beta version is refused: no reviewed beta raw surface", async () => {
  const f = read({ apiVersion: "beta" });
  await assert.rejects(f.run({}), error => error.code === "VALIDATION_ERROR" && /No reviewed beta raw reads/.test(error.message));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const [name, query, pattern] of [
  ["advanced $search", "$search=\"displayName:x\"", /not reviewed for raw reads/],
  ["$count", "$count=true", /not reviewed for raw reads/],
  ["server $skiptoken", "$skiptoken=abc", /page with --all/],
  ["collection $expand", "$expand=members", /single objects only/],
  ["unknown key", "$foo=1", /Unsupported query key/],
]) test(`${name} is refused before credentials`, async () => {
  const f = read({ query });
  await assert.rejects(f.run({}), error => error.code === "VALIDATION_ERROR" && pattern.test(error.message));
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("unreviewed $select fields are refused naming the reviewed set", async () => {
  const f = read({ query: "$select=id,passwordProfile" });
  await assert.rejects(f.run({}), error => {
    assert.equal(error.code, "VALIDATION_ERROR");
    assert.match(error.message, /Unreviewed \$select field passwordProfile/);
    assert.match(error.suggestions.join("\n"), /Reviewed fields: id, displayName/);
    return true;
  });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const binding of ["..", "a%2Fb", "a/b"]) test(`binding ${binding} never reaches credentials`, async () => {
  const f = read({ path: `/users/${binding}` });
  await assert.rejects(f.run({}), error => error.code === "VALIDATION_ERROR");
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("application profiles refuse delegated scopes before credentials", async () => {
  const f = fixture(json(200, { value: [] }));
  await assert.rejects(
    runApiGet({ path: "/users", apiVersion: "v1.0", profile: appProfile, scopes }, f.deps),
    error => error.code === "VALIDATION_ERROR" && /\.default audience/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const [path, body, result] of [
  ["/users/a", { id: "a" }, { id: "a" }],
  ["/users", { value: [{ id: "a" }] }, { returned: 1, complete: true, value: [{ id: "a" }] }],
]) test(`application read ${path} uses the configured audience without caller scopes`, async () => {
  const f = fixture(json(200, body));
  assert.deepEqual(await runApiGet({ path, apiVersion: "v1.0", profile: appProfile }, f.deps), result);
  assert.equal(f.credentialCalls.length, 1);
  assert.equal(f.credentialCalls[0][0], "acquire");
  assert.deepEqual(f.credentialCalls[0][2], ["https://graph.microsoft.com/.default"]);
  assert.equal(f.requests.length, 1);
  assert.equal(f.requests[0].headers.Authorization, "Bearer opaque-fixture-secret");
});

for (const path of ["/users/a", "/users"]) {
  for (const location of [
    "https://example.invalid/v1.0/users",
    "https://graph.microsoft.com/beta/users",
    "https://graph.microsoft.com/v1.0/groups",
  ]) test(`fetch-backed read ${path} refuses redirect to ${location}`, async t => {
    const requests = [];
    t.mock.method(globalThis, "fetch", async (url, options) => {
      requests.push({ url, options });
      return new Response("", { status: 302, headers: { Location: location } });
    });
    const f = fixture();
    await assert.rejects(
      runApiGet({ path, apiVersion: "v1.0", profile: delegatedProfile, scopes }, { ...f.deps, transport: fetchTransport }),
      { code: "POLICY_DENIED" },
    );
    assert.equal(requests.length, 1);
    assert.equal(requests[0].options.redirect, "manual");
  });
}

for (const [path, body, result] of [
  ["/users/a", { id: "a" }, { id: "a" }],
  ["/users", { value: [{ id: "a" }] }, { returned: 1, complete: true, value: [{ id: "a" }] }],
]) test(`fetch-backed read ${path} follows an authorized redirect`, async t => {
  const requests = [];
  const target = `https://graph.microsoft.com/v1.0${path}?$select=id`;
  t.mock.method(globalThis, "fetch", async (url, options) => {
    requests.push({ url, options });
    return requests.length === 1
      ? new Response("", { status: 302, headers: { Location: target } })
      : new Response(JSON.stringify(body), { status: 200 });
  });
  const f = fixture();
  assert.deepEqual(await runApiGet({ path, apiVersion: "v1.0", profile: delegatedProfile, scopes }, { ...f.deps, transport: fetchTransport }), result);
  assert.equal(requests.length, 2);
  assert.equal(requests[1].url, target);
  assert.equal(requests[0].options.redirect, "manual");
  assert.equal(requests[1].options.redirect, "manual");
  assert.equal(requests[1].options.headers.Authorization, "Bearer opaque-fixture-secret");
});

test("collection read returns redacted rows with truthful completion", async () => {
  const f = read({}, json(200, { value: [{ id: "a", displayName: "AccountKey=fixture-secret" }] }));
  assert.deepEqual(await f.run({}), {
    returned: 1,
    complete: true,
    value: [{ id: "a", displayName: "***redacted***" }],
  });
  assert.equal(f.requests.length, 1);
  assert.equal(f.requests[0].method, "GET");
  assert.equal(f.requests[0].url, "https://graph.microsoft.com/v1.0/users");
  assert.equal(f.requests[0].headers.Authorization, "Bearer opaque-fixture-secret");
});

test("reviewed $select travels on the authorized URL", async () => {
  const f = read({ query: "$select=id,displayName" }, json(200, { value: [{ id: "a" }] }));
  await f.run({});
  assert.equal(f.requests[0].url, "https://graph.microsoft.com/v1.0/users?%24select=id%2CdisplayName");
});

test("single read returns the redacted object itself", async () => {
  const f = read({ path: "/users/a", profile: delegatedProfile }, json(200, { id: "a", password: "fixture-secret" }));
  assert.deepEqual(await f.run({}), { id: "a", password: "***redacted***" });
  assert.equal(f.requests[0].url, "https://graph.microsoft.com/v1.0/users/a");
});

test("--all follows @odata.nextLink continuations", async () => {
  const f = read({ limit: undefined }, (request, count) => count === 1
    ? json(200, { value: [{ id: "a" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=abc" })
    : json(200, { value: [{ id: "b" }] }));
  assert.deepEqual(await f.run({}), { returned: 2, complete: true, value: [{ id: "a" }, { id: "b" }] });
  assert.equal(f.requests.length, 2);
});

test("row cap ends partial with the buffered remainder reason", async () => {
  const f = read({ limit: 2 }, json(200, { value: [{ id: "a" }, { id: "b" }, { id: "c" }] }));
  const result = await f.run({});
  assert.equal(result.complete, false);
  assert.match(result.reason, /row limit reached/);
  assert.deepEqual(result.value.map(row => row.id), ["a", "b"]);
  assert.match(result.help.join("\n"), /--all/);
});

test("long strings truncate unless --full", async () => {
  const long = "x".repeat(5000);
  const over = read({}, json(200, { value: [{ id: "a", displayName: long }] }));
  const truncated = await over.run({});
  assert.match(truncated.value[0].displayName, /\.\.\. \(truncated, 5000 chars total\)$/);
  assert.match(truncated.help.join("\n"), /--full/);
  const f = fixture(json(200, { value: [{ id: "a", displayName: long }] }));
  const full = await runApiGet({ path: "/users", apiVersion: "v1.0", profile: delegatedProfile, scopes, full: true }, f.deps);
  assert.equal(full.value[0].displayName, long);
  assert.equal(full.help, undefined);
});

const bin = resolve("dist/bin/mg-axi.js");
const profileConfig = JSON.stringify({ version: 1, defaultProfile: "soc", profiles: { soc: delegatedProfile } });

function cli(args, config = null) {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-api-"));
  try {
    if (config) writeFileSync(join(dir, "config.json"), config);
    return spawnSync(process.execPath, [bin, ...args], {
      encoding: "utf8", input: "", timeout: 10000,
      env: { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot, MG_AXI_CONFIG: join(dir, "config.json") },
    });
  } finally { rmSync(dir, { recursive: true, force: true }); }
}

test("api get help shows the positional path and reviewed flags", () => {
  const result = cli(["api", "get", "--help"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /mg-axi api get <path>/);
  assert.match(result.stdout, /--query/);
  assert.match(result.stdout, /--scopes/);
  assert.match(result.stdout, /--all/);
});

test("non-GET api verbs are refused before credentials", () => {
  const result = cli(["api", "delete", "/users"]);
  assert.equal(result.status, 2);
  assert.equal(result.stderr, "");
  const output = decode(result.stdout);
  assert.equal(output.code, "VALIDATION_ERROR");
  assert.match(output.error, /reviewed GET reads only/);
});

test("missing api path fails with usage output", () => {
  const result = cli(["api", "get"]);
  assert.equal(result.status, 2);
  assert.equal(decode(result.stdout).code, "VALIDATION_ERROR");
});

test("unreviewed CLI route is refused before credentials", () => {
  const result = cli(["api", "get", "/me/messages", "--scopes", scopes], profileConfig);
  assert.equal(result.status, 2);
  assert.equal(result.stderr, "");
  assert.match(decode(result.stdout).error, /not in the reviewed raw inventory/);
});

test("beta CLI reads are refused before credentials", () => {
  const result = cli(["api", "get", "/users", "--api-version", "beta", "--scopes", scopes], profileConfig);
  assert.equal(result.status, 2);
  assert.match(decode(result.stdout).error, /No reviewed beta raw reads/);
});

test("unreviewed CLI $select is refused before credentials", () => {
  const result = cli(["api", "get", "/users", "--query", "$select=id,passwordProfile", "--scopes", scopes], profileConfig);
  assert.equal(result.status, 2);
  assert.match(decode(result.stdout).error, /Unreviewed \$select field/);
});

test("top help lists the raw read leaf", () => {
  const result = cli(["--help"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /api get/);
});
