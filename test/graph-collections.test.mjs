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

function fixture(handler, credentialHandler) {
  const credentialCalls = [];
  const requests = [];
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const delegated = new DelegatedAuth({
    storage: "session-only",
    login: async () => credential,
    silent: async (...args) => { credentialCalls.push(["silent", ...args]); return credentialHandler ? credentialHandler(credential) : credential; },
  });
  const application = new ApplicationAuth({ storage: "session-only", acquire: async (...args) => {
    credentialCalls.push(["acquire", ...args]);
    return credentialHandler ? credentialHandler(credential) : credential;
  } });
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
      sleep: (ms, signal) => new Promise((resolve, reject) => {
        signal?.throwIfAborted();
        const onAbort = () => {
          clearImmediate(timer);
          reject(signal.reason);
        };
        const timer = setImmediate(() => {
          signal?.removeEventListener("abort", onAbort);
          sleeps.push(ms);
          t += ms;
          resolve();
        });
        signal?.addEventListener("abort", onAbort, { once: true });
      }),
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

for (const next of ["https://example.invalid/", "https://graph.microsoft.com/v1.0/applications", "https://graph.microsoft.com/v1.0/users/other"]) {
  test(`cursor destination is authorized before credentials: ${next}`, async () => {
    const f = fixture(json(200, { value: [] }));
    const cursor = Buffer.from(JSON.stringify({ v: 2, op: users.id, next, buffered: [], seen: [], query: {}, consistencyLevel: null })).toString("base64url");
    await assert.rejects(f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor }), { code: "POLICY_DENIED" });
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
  });
}

for (const [name, response, budget] of [
  ["throttle", json(429, {}, { "retry-after": "120" }), {}],
  ["throttle byte budget", json(503, { message: "busy" }), { maxBytes: 1 }],
  ["redirect request budget", json(302, {}, { location: "/v1.0/users?$skiptoken=target" }), { maxRequests: 1 }],
  ["redirect byte budget", json(302, {}, { location: "/v1.0/users?$skiptoken=target" }), { maxBytes: 1 }],
  ["retry request budget", json(429, {}, { "retry-after": "1" }), { maxRequests: 1 }],
  ["retry deadline", json(429, {}, { "retry-after": "5" }), { deadlineMs: 1000 }],
]) {
  test(`unfinished ${name} request can resume`, async () => {
    const f = fixture((request, count) => count === 1 ? response : json(200, { value: [{ id: "resumed" }] }));
    const first = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, budget, clock: fakeClock().clock });
    assert.equal(first.complete, false);
    const second = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor });
    assert.deepEqual(second.value, [{ id: "resumed" }]);
    assert.equal(second.complete, true);
    assert.equal(f.requests.length, 2);
  });
}

for (const method of ["collect", "execute"]) {
  test(`${method} keeps continuation token case distinct`, async () => {
    const f = fixture((request, count) => method === "collect"
      ? json(200, { value: [{ id: count }], ...(count < 3 ? { "@odata.nextLink": `https://graph.microsoft.com/v1.0/users?$skiptoken=${count === 1 ? "AbC" : "abc"}` } : {}) })
      : count < 3 ? json(302, {}, { location: `/v1.0/users?$skiptoken=${count === 1 ? "AbC" : "abc"}` }) : json(200, { value: [{ id: 3 }] }));
    const result = await f.session[method]({ profile: delegatedProfile, operation: users, scopes });
    assert.equal(f.requests.length, 3);
    assert.equal(result.value.at(-1).id, 3);
  });

  test(`${method} bounds an unresponsive transport and aborts it`, async () => {
    const f = fixture(() => new Promise(() => {}));
    const pending = f.session[method]({ profile: delegatedProfile, operation: users, scopes, clock: fakeClock().clock });
    if (method === "collect") {
      const result = await pending;
      assert.equal(result.complete, false);
      assert.match(result.reason, /deadline/);
      assert.equal(decodeCursor(result.cursor).seen.length, 0);
    } else {
      await assert.rejects(pending, error => error.code === "GRAPH_ERROR" && /deadline/.test(error.message));
    }
    assert.equal(f.requests[0].signal.aborted, true);
  });

  test(`${method} rejects a success received after its deadline`, async () => {
    let now = 0;
    const clock = { now: () => now, sleep: () => new Promise(() => {}) };
    const f = fixture(() => { now = 30_001; return json(200, { value: [{ id: "late" }] }); });
    const pending = f.session[method]({ profile: delegatedProfile, operation: users, scopes, clock });
    if (method === "collect") {
      const result = await pending;
      assert.equal(result.complete, false);
      assert.match(result.reason, /deadline/);
      assert.deepEqual(result.value, []);
    } else {
      await assert.rejects(pending, error => error.code === "GRAPH_ERROR" && /deadline/.test(error.message));
    }
  });

  for (const [query, code] of [
    [{ $expand: "files" }, "POLICY_DENIED"],
    [{ $expand: "manager/messages" }, "POLICY_DENIED"],
    [{ $expand: "manager($expand=messages)" }, "VALIDATION_ERROR"],
    [{ $expand: "*" }, "VALIDATION_ERROR"],
    [{ $search: '"displayName:ana"', $expand: "manager" }, "VALIDATION_ERROR"],
    [{ $count: "true", $expand: "manager" }, "VALIDATION_ERROR"],
  ]) {
    test(`${method} rejects unsafe query ${JSON.stringify(query)} before credentials`, async () => {
      const f = fixture(json(200, { value: [] }));
      await assert.rejects(f.session[method]({ profile: delegatedProfile, operation: users, scopes, query, consistencyLevel: "eventual" }), { code });
      assert.equal(f.credentialCalls.length, 0);
      assert.equal(f.requests.length, 0);
    });
  }
}

for (const [status, code] of [[401, "AUTH_REQUIRED"], [403, "GRAPH_ERROR"], [404, "GRAPH_ERROR"], [500, "GRAPH_ERROR"]]) {
  test(`byte budget preserves terminal ${status} error`, async () => {
    const f = fixture(json(status, { error: { code: "Denied", message: "terminal failure" } }));
    await assert.rejects(f.session.collect({ profile: delegatedProfile, operation: users, scopes, budget: { maxBytes: 1 } }), error => error.code === code && error.message.includes(String(status)));
    assert.equal(f.requests.length, 1);
  });
}

for (const [query, code] of [
  ["$expand=files", "POLICY_DENIED"],
  ["$search=x&$expand=manager", "VALIDATION_ERROR"],
  ["$count=true&$expand=manager", "VALIDATION_ERROR"],
  ["$expand=files&$expand=manager", "POLICY_DENIED"],
]) {
  const target = `https://graph.microsoft.com/v1.0/users?${query}`;
  test(`cursor validates query context: ${query}`, async () => {
    const f = fixture(json(200, { value: [] }));
    const cursor = Buffer.from(JSON.stringify({ v: 2, op: users.id, next: target, buffered: [], seen: [], query: {}, consistencyLevel: "eventual" })).toString("base64url");
    await assert.rejects(f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor, consistencyLevel: "eventual" }), { code });
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
  });
  for (const [method, response] of [
    ["collect", json(200, { value: [], "@odata.nextLink": target })],
    ["collect", json(302, {}, { location: target })],
    ["execute", json(302, {}, { location: target })],
  ]) {
    test(`${method} validates ${response.status} target query: ${query}`, async () => {
      const f = fixture(response);
      await assert.rejects(f.session[method]({ profile: delegatedProfile, operation: users, scopes, consistencyLevel: "eventual" }), { code });
      assert.equal(f.requests.length, 1);
    });
  }
}

for (const method of ["collect", "execute"]) {
  test(`${method} still accepts a metadata expansion`, async () => {
    const f = fixture(json(200, { value: [{ id: "a", manager: { id: "b" } }] }));
    const result = await f.session[method]({ profile: delegatedProfile, operation: users, scopes, query: { $expand: "manager" } });
    assert.deepEqual(result.value, [{ id: "a", manager: { id: "b" } }]);
  });

  test(`${method} cancels an in-flight transport that ignores the signal`, async () => {
    const controller = new AbortController();
    const f = fixture(() => { controller.abort(); return new Promise(() => {}); });
    await assert.rejects(f.session[method]({ profile: delegatedProfile, operation: users, scopes, signal: controller.signal }), { name: "AbortError" });
    assert.equal(f.requests.length, 1);
  });
}

for (const retainedSearch of [false, true]) {
  for (const [name, options, throttled] of [
    ["row cap", { limit: 1 }, false],
    ["request budget", { budget: { maxRequests: 1 } }, false],
    ["byte budget", { budget: { maxBytes: 1 } }, false],
    ["throttle", {}, true],
    ["deadline", { budget: { deadlineMs: 1000 } }, true],
  ]) {
    test(`${name} preserves search context across repeated resumes, retained search: ${retainedSearch}`, async () => {
      const next = `https://graph.microsoft.com/v1.0/users?$skiptoken=opaque${retainedSearch ? '&$search=%22displayName%3Aana%22' : ""}`;
      const f = fixture((request, count) => throttled && count === 1
        ? json(429, {}, { "retry-after": "120" })
        : count === (throttled ? 2 : 1)
          ? json(200, { value: [{ id: "a" }, { id: "b" }], "@odata.nextLink": next })
          : json(200, { value: [{ id: "c" }, { id: "d" }] }));
      const first = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, query: { $search: '"displayName:ana"' }, consistencyLevel: "eventual", ...options });
      assert.equal(first.complete, false);
      const second = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor, limit: 1 });
      assert.equal(second.complete, false);
      const third = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: second.cursor });
      assert.equal(third.complete, true);
      assert.deepEqual([...first.value, ...second.value, ...third.value], [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }]);
      assert.ok(f.requests.every(request => request.headers.ConsistencyLevel === "eventual"));
      assert.equal(f.requests.at(-1).url, next);
    });
  }
}

for (const query of [{ $count: "true" }, {}]) {
  test(`resume preserves eventual consistency through a redirect for ${JSON.stringify(query)}`, async () => {
    const f = fixture((request, count) => count === 1
      ? json(200, { value: [{ id: "a" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=one" })
      : count === 2 ? json(302, {}, { location: "/v1.0/users?$skiptoken=two" })
        : json(200, { value: [{ id: "b" }] }));
    const first = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, query, consistencyLevel: "eventual", limit: 1 });
    const second = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor });
    assert.equal(second.complete, true);
    assert.deepEqual(second.value, [{ id: "b" }]);
    assert.ok(f.requests.every(request => request.headers.ConsistencyLevel === "eventual"));
  });
}

for (const overrides of [{ query: { $filter: "id eq 'other'" } }, { consistencyLevel: "eventual" }]) {
  test(`conflicting resume context fails before credentials: ${JSON.stringify(overrides)}`, async () => {
    const maker = fixture(json(200, { value: [{ id: "a" }, { id: "b" }] }));
    const first = await maker.session.collect({ profile: delegatedProfile, operation: users, scopes, query: { $select: "id" }, limit: 1 });
    const f = fixture(json(200, { value: [] }));
    await assert.rejects(f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor, ...overrides }), { code: "VALIDATION_ERROR" });
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
  });
}

test("matching resume context accepts query keys in a different order", async () => {
  const f = fixture(json(200, { value: [{ id: "a" }, { id: "b" }] }));
  const first = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, query: { $search: "ana", $select: "id" }, consistencyLevel: "eventual", limit: 1 });
  const second = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor, query: { $select: "id", $search: "ana" }, consistencyLevel: "eventual" });
  assert.deepEqual(second.value, [{ id: "b" }]);
  assert.equal(second.complete, true);
});

for (const context of [
  { query: { $search: "ana" }, consistencyLevel: null },
  { query: { $top: 5 }, consistencyLevel: null },
  { query: { unsupported: "value" }, consistencyLevel: null },
  { query: {}, consistencyLevel: "strong" },
  { query: [], consistencyLevel: null },
]) {
  test(`invalid saved cursor context fails before credentials: ${JSON.stringify(context)}`, async () => {
    const f = fixture(json(200, { value: [] }));
    const cursor = Buffer.from(JSON.stringify({ v: 2, op: users.id, next: null, buffered: [{ id: "a" }], seen: [], ...context })).toString("base64url");
    await assert.rejects(f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor }), { code: "VALIDATION_ERROR" });
    assert.equal(f.credentialCalls.length, 0);
    assert.equal(f.requests.length, 0);
  });
}

for (const budget of [undefined, { maxBytes: 1 }]) {
  test(`only the exact nextLink annotation is followed, byte cap: ${Boolean(budget)}`, async () => {
    const next = "https://graph.microsoft.com/v1.0/users?$skiptoken=exact";
    const f = fixture((request, count) => count === 1
      ? json(200, { value: [{ id: "a" }], "@ODATA.NEXTLINK": "https://example.invalid/", "@odata.nextLink": next })
      : json(200, { value: [{ id: "b" }] }));
    const first = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, budget });
    const result = budget ? await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor }) : first;
    assert.equal(result.complete, true);
    assert.deepEqual(result.value, [{ id: "a" }, { id: "b" }]);
    assert.equal(f.requests[1].url, next);
  });
}

test("alternate nextLink capitalization is ignored", async () => {
  const f = fixture(json(200, { value: [{ id: "a" }], "@odata.nextlink": "https://example.invalid/" }));
  const result = await f.session.collect({ profile: delegatedProfile, operation: users, scopes });
  assert.equal(result.complete, true);
  assert.deepEqual(result.value, [{ id: "a" }]);
  assert.equal(f.requests.length, 1);
});

for (const profile of [delegatedProfile, { ...delegatedProfile, mode: "application", credentialRef: { provider: "federated", key } }]) {
  for (const method of ["collect", "execute"]) {
    const args = { profile, operation: users, ...(profile.mode === "delegated" ? { scopes } : {}) };

    test(`${method} with ${profile.mode} credentials rejects pre-aborted calls before acquisition`, async () => {
      const controller = new AbortController();
      controller.abort();
      const f = fixture(json(200, { value: [] }));
      await assert.rejects(f.session[method]({ ...args, signal: controller.signal }), { name: "AbortError" });
      assert.equal(f.credentialCalls.length, 0);
      assert.equal(f.requests.length, 0);
    });

    test(`${method} with ${profile.mode} credentials cancels pending acquisition`, async () => {
      const controller = new AbortController();
      const f = fixture(json(200, { value: [] }), () => { controller.abort(); return new Promise(() => {}); });
      await assert.rejects(f.session[method]({ ...args, signal: controller.signal }), { name: "AbortError" });
      assert.equal(f.credentialCalls.length, 1);
      assert.equal(f.requests.length, 0);
    });

    test(`${method} with ${profile.mode} credentials bounds pending acquisition through the clock`, async () => {
      const f = fixture(json(200, { value: [] }), () => new Promise(() => {}));
      const pending = f.session[method]({ ...args, clock: fakeClock().clock, budget: { deadlineMs: 1000 } });
      if (method === "collect") {
        const result = await pending;
        assert.equal(result.complete, false);
        assert.match(result.reason, /deadline/);
        assert.equal(result.requests, 0);
        assert.equal(result.bytes, 0);
        assert.equal(decodeCursor(result.cursor).next, "https://graph.microsoft.com/v1.0/users");
      } else {
        await assert.rejects(pending, error => error.code === "GRAPH_ERROR" && /deadline/.test(error.message));
      }
      assert.equal(f.credentialCalls.length, 1);
      assert.equal(f.requests.length, 0);
    });

    test(`${method} with ${profile.mode} credentials rejects acquisition that completes too late`, async () => {
      let now = 0;
      const clock = { now: () => now, sleep: () => new Promise(() => {}) };
      const f = fixture(json(200, { value: [] }), credential => { now = 30_001; return credential; });
      const pending = f.session[method]({ ...args, clock });
      if (method === "collect") {
        const result = await pending;
        assert.equal(result.complete, false);
        assert.match(result.reason, /deadline/);
        assert.equal(result.requests, 0);
      } else {
        await assert.rejects(pending, error => error.code === "GRAPH_ERROR" && /deadline/.test(error.message));
      }
      assert.equal(f.requests.length, 0);
    });

    test(`${method} with ${profile.mode} credentials shares its deadline with transport`, async () => {
      let now = 0;
      const clock = { now: () => now, sleep: () => new Promise(() => {}) };
      const f = fixture(() => { now = 30_001; return json(200, { value: [{ id: "late" }] }); }, credential => { now = 29_990; return credential; });
      const pending = f.session[method]({ ...args, clock });
      if (method === "collect") {
        const result = await pending;
        assert.equal(result.complete, false);
        assert.match(result.reason, /deadline/);
      } else {
        await assert.rejects(pending, error => error.code === "GRAPH_ERROR" && /deadline/.test(error.message));
      }
      assert.equal(f.requests.length, 1);
    });

    test(`${method} with ${profile.mode} credentials still succeeds before the shared deadline`, async () => {
      let now = 0;
      const clock = { now: () => now, sleep: () => new Promise(() => {}) };
      const f = fixture(() => { now = 200; return json(200, { value: [{ id: "a" }] }); }, credential => { now = 100; return credential; });
      const result = await f.session[method]({ ...args, clock });
      assert.deepEqual(result.value, [{ id: "a" }]);
      assert.equal(f.requests.length, 1);
      assert.equal(f.requests[0].headers.Authorization, "Bearer opaque-fixture-secret");
    });

    test(`${method} with ${profile.mode} credentials preserves authentication failures`, async () => {
      const f = fixture(json(200, { value: [] }), () => { throw new Error("provider unavailable"); });
      await assert.rejects(f.session[method](args), { code: "AUTH_REQUIRED" });
      assert.equal(f.requests.length, 0);
    });
  }
}

test("credential deadline preserves buffered rows and context on a resumed collection", async () => {
  const next = "https://graph.microsoft.com/v1.0/users?$skiptoken=next";
  const maker = fixture(json(200, { value: [{ id: "a" }, { id: "b" }], "@odata.nextLink": next }));
  const first = await maker.session.collect({ profile: delegatedProfile, operation: users, scopes, query: { $search: "ana" }, consistencyLevel: "eventual", limit: 1 });
  const blocked = fixture(json(200, { value: [] }), () => new Promise(() => {}));
  const partial = await blocked.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: first.cursor, clock: fakeClock().clock });
  assert.equal(partial.complete, false);
  assert.deepEqual(partial.value, []);
  assert.equal(partial.requests, 0);
  const f = fixture(json(200, { value: [{ id: "c" }] }));
  const resumed = await f.session.collect({ profile: delegatedProfile, operation: users, scopes, cursor: partial.cursor });
  assert.equal(resumed.complete, true);
  assert.deepEqual(resumed.value, [{ id: "b" }, { id: "c" }]);
  assert.equal(f.requests[0].url, next);
  assert.equal(f.requests[0].headers.ConsistencyLevel, "eventual");
});
