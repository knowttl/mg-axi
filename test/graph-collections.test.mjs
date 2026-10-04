import assert from "node:assert/strict";
import { test } from "node:test";
import { ApplicationAuth } from "../dist/app-auth.js";
import { DelegatedAuth } from "../dist/auth.js";
import { GraphSession, resolveSessionOperation } from "../dist/graph-session.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key } };
const scopes = ["https://graph.microsoft.com/User.Read"];
const users = resolveSessionOperation("v1.0", "GET", "/users");

function json(status, body, headers = {}) {
  return { status, headers, body: JSON.stringify(body) };
}

function fixture(handler) {
  const credentialCalls = [];
  const requests = [];
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const delegated = new DelegatedAuth({
    storage: "session-only",
    login: async () => credential,
    silent: async (...args) => { credentialCalls.push(["silent", ...args]); return credential; },
  });
  const application = new ApplicationAuth({ storage: "session-only", acquire: async () => credential });
  const transport = async request => {
    requests.push(request);
    const response = typeof handler === "function" ? await handler(request, requests.length) : handler;
    return { headers: {}, body: "", ...response };
  };
  return { credentialCalls, requests, session: new GraphSession({ delegated, application, transport }) };
}

function fakeClock(start = 1_000_000) {
  let t = start;
  const sleeps = [];
  return {
    sleeps,
    clock: {
      now: () => t,
      sleep: async (ms, signal) => {
        signal?.throwIfAborted();
        sleeps.push(ms);
        t += ms;
      },
    },
  };
}

function decodeCursor(cursor) {
  return JSON.parse(Buffer.from(cursor, "base64url").toString("utf8"));
}

test("collect follows exact nextLink and preserves required headers", async () => {
  const next = "https://graph.microsoft.com/v1.0/users?$skiptoken=abc123&$top=2";
  const f = fixture((request, count) => count === 1
    ? json(200, { value: [{ id: "a" }], "@odata.nextLink": next })
    : json(200, { value: [{ id: "b" }] }));
  const result = await f.session.collect({ profile: delegatedProfile, operation: users, query: { $top: "2" }, scopes });
  assert.deepEqual(result.value, [{ id: "a" }, { id: "b" }]);
  assert.equal(result.complete, true);
  assert.equal(result.cursor, undefined);
  assert.equal(result.requests, 2);
  assert.ok(result.bytes > 0);
  assert.equal(f.requests.length, 2);
  assert.equal(f.requests[1].url, next);
  assert.ok(f.requests.every(r => r.headers.Authorization === "Bearer opaque-fixture-secret"));
  assert.ok(f.requests.every(r => r.headers.Accept === "application/json"));
  assert.match(f.requests[0].headers["client-request-id"], /^[0-9a-f-]{36}$/);
  assert.notEqual(f.requests[0].headers["client-request-id"], f.requests[1].headers["client-request-id"]);
});

test("collect redacts sentinels on every page", async () => {
  const next = "https://graph.microsoft.com/v1.0/users?$skiptoken=x";
  const f = fixture((request, count) => count === 1
    ? json(200, { value: [{ id: "a", note: "AccountKey=fixture-secret" }], "@odata.nextLink": next })
    : json(200, { value: [{ password: "fixture-secret" }] }));
  const result = await f.session.collect({ profile: delegatedProfile, operation: users, scopes });
  assert.deepEqual(result.value, [{ id: "a", note: "***redacted***" }, { password: "***redacted***" }]);
  assert.equal(result.complete, true);
});

test("hostile nextLink is denied without a second request", async () => {
  const f = fixture(json(200, { value: [{ id: "a" }], "@odata.nextLink": "https://example.invalid/v1.0/users" }));
  await assert.rejects(f.session.collect({ profile: delegatedProfile, operation: users, scopes }), { code: "POLICY_DENIED" });
  assert.equal(f.requests.length, 1);
});

test("row cap buffers the remainder instead of discarding it", async () => {
  const next = "https://graph.microsoft.com/v1.0/users?$skiptoken=page2";
  const f = fixture((request, count) => count === 1
    ? json(200, { value: [{ id: "a" }, { id: "b" }, { id: "c" }], "@odata.nextLink": next })
    : json(200, { value: [{ id: "d" }] }));
  const first = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, limit: 2 });
  assert.deepEqual(first.value, [{ id: "a" }, { id: "b" }]);
  assert.equal(first.complete, false);
  assert.ok(first.cursor);
  assert.equal(first.requests, 1);
  const cursor = decodeCursor(first.cursor);
  assert.deepEqual(cursor.buffered, [{ id: "c" }]);
  const second = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, limit: 10, cursor: first.cursor });
  assert.deepEqual(second.value, [{ id: "c" }, { id: "d" }]);
  assert.equal(second.complete, true);
});

test("request budget ends as partial with a resumable cursor", async () => {
  const next = "https://graph.microsoft.com/v1.0/users?$skiptoken=page2";
  const f = fixture((request, count) => count === 1
    ? json(200, { value: [{ id: "a" }], "@odata.nextLink": next })
    : json(200, { value: [{ id: "b" }] }));
  const first = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, budget: { maxRequests: 1, maxBytes: 5_000_000, deadlineMs: 30_000 } });
  assert.deepEqual(first.value, [{ id: "a" }]);
  assert.equal(first.complete, false);
  assert.match(first.reason ?? "", /request budget/);
  assert.ok(first.cursor);
  const second = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor });
  assert.deepEqual(second.value, [{ id: "b" }]);
  assert.equal(second.complete, true);
});

test("byte budget preserves unreturned rows in the cursor", async () => {
  const next = "https://graph.microsoft.com/v1.0/users?$skiptoken=page2";
  const f = fixture((request, count) => count === 1
    ? json(200, { value: [{ id: "a" }], "@odata.nextLink": next })
    : json(200, { value: [{ id: "b" }] }));
  const first = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, budget: { maxRequests: 20, maxBytes: 10, deadlineMs: 30_000 } });
  assert.equal(first.complete, false);
  assert.match(first.reason ?? "", /byte budget/);
  assert.ok(first.cursor);
  const cursor = decodeCursor(first.cursor);
  assert.deepEqual(cursor.buffered, [{ id: "a" }]);
  const second = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor });
  assert.ok(second.value.some(row => row.id === "a"));
  assert.equal(second.complete, true);
});

test("deadline exceeded by a Retry-After wait ends as partial", async () => {
  const fake = fakeClock();
  const f = fixture(() => json(429, { error: { code: "TooManyRequests", message: "throttled" } }, { "retry-after": "5" }));
  const result = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, clock: fake.clock, budget: { maxRequests: 20, maxBytes: 5_000_000, deadlineMs: 1000 } });
  assert.equal(result.complete, false);
  assert.match(result.reason ?? "", /deadline/);
  assert.deepEqual(fake.sleeps, []);
  assert.ok(result.cursor);
});

test("Retry-After seconds are honored through the fake clock", async () => {
  const fake = fakeClock();
  const f = fixture((request, count) => count === 1
    ? json(429, { error: { code: "TooManyRequests", message: "slow" } }, { "retry-after": "2" })
    : json(200, { value: [{ id: "a" }] }));
  const result = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, clock: fake.clock });
  assert.deepEqual(result.value, [{ id: "a" }]);
  assert.equal(result.complete, true);
  assert.deepEqual(fake.sleeps, [2000]);
  assert.equal(f.requests.length, 2);
});

test("Retry-After HTTP dates resolve through the fake clock", async () => {
  const start = 1_700_000_000_000;
  const fake = fakeClock(start);
  const date = new Date(start + 3000).toUTCString();
  const f = fixture((request, count) => count === 1
    ? json(429, { error: { code: "TooManyRequests", message: "slow" } }, { "retry-after": date })
    : json(200, { value: [{ id: "a" }] }));
  const result = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, clock: fake.clock });
  assert.equal(result.complete, true);
  assert.equal(fake.sleeps.length, 1);
  assert.ok(Math.abs(fake.sleeps[0] - 3000) < 5);
});

test("absent Retry-After uses bounded backoff, not real time", async () => {
  const fake = fakeClock();
  const f = fixture((request, count) => count === 1
    ? json(503, { error: { code: "ServiceUnavailable", message: "busy" } })
    : json(200, { value: [{ id: "a" }] }));
  const result = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, clock: fake.clock });
  assert.equal(result.complete, true);
  assert.deepEqual(fake.sleeps, [1000]);
});

test("oversized Retry-After never retries a single read", async () => {
  const fake = fakeClock();
  const f = fixture(json(429, { error: { code: "TooManyRequests", message: "slow" } }, { "retry-after": "120" }));
  await assert.rejects(
    f.session.execute({ profile: delegatedProfile, operation: users, scopes, clock: fake.clock }),
    error => error.code === "GRAPH_ERROR" && /retry after 120/.test(error.message),
  );
  assert.deepEqual(fake.sleeps, []);
  assert.equal(f.requests.length, 1);
});

test("oversized Retry-After returns partial with a retry cursor for collections", async () => {
  const fake = fakeClock();
  const f = fixture(json(429, { error: { code: "TooManyRequests", message: "slow" } }, { "retry-after": "120" }));
  const result = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, clock: fake.clock });
  assert.equal(result.complete, false);
  assert.match(result.reason ?? "", /throttled/);
  assert.deepEqual(result.value, []);
  assert.ok(result.cursor);
  assert.deepEqual(fake.sleeps, []);
});

test("execute retries a small Retry-After through the fake clock", async () => {
  const fake = fakeClock();
  const f = fixture((request, count) => count === 1
    ? json(429, { error: { code: "TooManyRequests", message: "slow" } }, { "retry-after": "1" })
    : json(200, { value: [{ id: "a" }] }));
  const result = await f.session.execute({ profile: delegatedProfile, operation: users, scopes, clock: fake.clock });
  assert.deepEqual(result, { value: [{ id: "a" }] });
  assert.deepEqual(fake.sleeps, [1000]);
  assert.equal(f.requests.length, 2);
});

test("continuation cycles end as partial and never complete across resumes", async () => {
  const firstUrl = "https://graph.microsoft.com/v1.0/users?$skiptoken=one";
  const f = fixture(request => request.url.includes("skiptoken=one")
    ? json(200, { value: [{ id: "b" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=two" })
    : json(200, { value: [{ id: "a" }], "@odata.nextLink": firstUrl }));
  const first = await f.session.collect({ profile: delegatedProfile, operation: users, scopes });
  assert.equal(first.complete, false);
  assert.match(first.reason ?? "", /cycle/);
  assert.ok(first.cursor);
  const cursor = decodeCursor(first.cursor);
  assert.ok(cursor.seen.length <= 64);
  const second = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor });
  assert.equal(second.complete, false);
  assert.match(second.reason ?? "", /cycle/);
});

test("invalid cursors fail before credential or HTTP", async () => {
  const f = fixture(json(200, { value: [] }));
  await assert.rejects(f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: "not-a-cursor" }), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("cursors from another operation are rejected before HTTP", async () => {
  const groups = resolveSessionOperation("v1.0", "GET", "/groups");
  const maker = fixture(json(200, { value: [{ id: "a" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=x" }));
  const partial = await maker.session.collect({ profile: delegatedProfile, operation: users, scopes, budget: { maxRequests: 1, maxBytes: 5_000_000, deadlineMs: 30_000 } });
  assert.equal(partial.complete, false);
  const f = fixture(json(200, { value: [] }));
  await assert.rejects(f.session.collect({ profile: delegatedProfile, operation: groups, scopes, cursor: partial.cursor }), { code: "VALIDATION_ERROR" });
  assert.equal(f.requests.length, 0);
});

test("$search without eventual consistency fails before credential or HTTP", async () => {
  const f = fixture(json(200, { value: [] }));
  await assert.rejects(
    f.session.collect({ profile: delegatedProfile, operation: users, query: { $search: "\"displayName:ana\"" }, scopes }),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("$search with eventual consistency sends the header on every page", async () => {
  const next = "https://graph.microsoft.com/v1.0/users?$skiptoken=x";
  const f = fixture((request, count) => count === 1
    ? json(200, { value: [{ id: "a" }], "@odata.nextLink": next })
    : json(200, { value: [{ id: "b" }] }));
  const result = await f.session.collect({ profile: delegatedProfile, operation: users, query: { $search: "\"displayName:ana\"", $select: "id" }, consistencyLevel: "eventual", scopes });
  assert.equal(result.complete, true);
  assert.ok(f.requests.every(r => r.headers.ConsistencyLevel === "eventual"));
});

test("$count=true without eventual consistency fails before HTTP", async () => {
  const f = fixture(json(200, { value: [] }));
  await assert.rejects(
    f.session.collect({ profile: delegatedProfile, operation: users, query: { $count: "true" }, scopes }),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.requests.length, 0);
});

test("aborted signals fail fast before HTTP", async () => {
  const f = fixture(json(200, { value: [] }));
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(
    f.session.collect({ profile: delegatedProfile, operation: users, scopes, signal: controller.signal }),
    error => error.name === "AbortError" || /abort/i.test(error.message),
  );
  assert.equal(f.requests.length, 0);
});

test("cancellation while waiting for Retry-After aborts the collection", async () => {
  const f = fixture(json(429, { error: { code: "TooManyRequests", message: "slow" } }, { "retry-after": "5" }));
  const controller = new AbortController();
  const hanging = { now: () => 1_000_000, sleep: (ms, signal) => new Promise((resolve, reject) => {
    if (signal?.aborted) return reject(signal.reason ?? new DOMException("aborted", "AbortError"));
    signal?.addEventListener("abort", () => reject(signal.reason ?? new DOMException("aborted", "AbortError")), { once: true });
  }) };
  const pending = f.session.collect({ profile: delegatedProfile, operation: users, scopes, clock: hanging, signal: controller.signal });
  setImmediate(() => controller.abort());
  await assert.rejects(pending, error => error.name === "AbortError" || /abort/i.test(error.message));
  assert.equal(f.requests.length, 1);
});

test("invalid budgets fail before credential or HTTP", async () => {
  const f = fixture(json(200, { value: [] }));
  await assert.rejects(
    f.session.collect({ profile: delegatedProfile, operation: users, scopes, budget: { maxRequests: 0, maxBytes: 5_000_000, deadlineMs: 30_000 } }),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});
