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
const policyScopes = [
  "https://graph.microsoft.com/User.Export.All",
  "https://graph.microsoft.com/User.Read.All",
];

const blobUrl = "https://contoso.blob.core.windows.net/exports/op-1.zip?sv=2024-01-01&sig=fixture-signature";
const op1 = {
  id: "op-complete-1",
  completedDateTime: "2026-10-01T00:00:00Z",
  status: "complete",
  storageLocation: blobUrl,
  userId: "user-1",
  submittedDateTime: "2026-09-30T00:00:00Z",
  progress: "100",
};
const op2 = {
  id: "op-running-2",
  completedDateTime: null,
  status: "running",
  storageLocation: null,
  userId: "user-2",
  submittedDateTime: "2026-10-02T00:00:00Z",
  progress: "42",
};
const op3 = { id: "op-bare-3" };
const operations = [op1, op2, op3];
const longUserId = `user-${"u".repeat(600)}`;

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-data-policy-"));
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

function policyTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/dataPolicyOperations/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(operations.length) };
    }
    if (path === "/v1.0/dataPolicyOperations") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [op3] });
      return json(200, {
        value: [op1, op2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/dataPolicyOperations?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/dataPolicyOperations\/([^/]+)$/.exec(path);
    if (single) {
      const found = operations.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such operation" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, fixture, calls = []) {
  const transportFixture = fixture ?? policyTransport();
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

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists operations with compact rows and redacted storage locations`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "data-policy-operation", "list", "--profile", profile], overrides);
      assert.deepEqual(result.dataPolicyOperations, [
        { id: op1.id, status: "complete", userId: "user-1", submittedDateTime: "2026-09-30T00:00:00Z" },
        { id: op2.id, status: "running", userId: "user-2", submittedDateTime: "2026-10-02T00:00:00Z" },
        { id: op3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra data-policy-operation show --id <operation-id>")));
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      assert.ok(!JSON.stringify(result).includes("contoso.blob.core.windows.net"));
      assert.ok(!JSON.stringify(result).includes("fixture-signature"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/dataPolicyOperations?"));
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$select"), "id,status,userId,submittedDateTime");
      assert.ok(!sent.has("$filter"));
      if (mode === "delegated") assert.deepEqual(calls[0][1], policyScopes);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one operation with the full set and a redacted blob URL`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "data-policy-operation", "show", "--id", op1.id, "--profile", profile], overrides);
      assert.deepEqual(result.dataPolicyOperation, {
        id: op1.id,
        completedDateTime: "2026-10-01T00:00:00Z",
        status: "complete",
        storageLocation: "***redacted***",
        userId: "user-1",
        submittedDateTime: "2026-09-30T00:00:00Z",
        progress: "100",
      });
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      assert.ok(!JSON.stringify(result).includes("contoso.blob.core.windows.net"));
      const running = await executeArgv(["entra", "data-policy-operation", "show", "--id", op2.id, "--profile", profile], overrides);
      assert.equal(running.dataPolicyOperation.storageLocation, null);
      const bare = await executeArgv(["entra", "data-policy-operation", "show", "--id", op3.id, "--profile", profile], overrides);
      assert.deepEqual(bare.dataPolicyOperation, { id: op3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} redacts storage locations without a sig sentinel`, async () => {
    const state = setupProfiles();
    try {
      const plain = transport(request => {
        const url = new URL(request.url);
        if (url.pathname === "/v1.0/dataPolicyOperations/plain-1") {
          return json(200, { ...op1, id: "plain-1", storageLocation: "https://contoso.blob.core.windows.net/exports/plain.zip" });
        }
        return json(404, { error: { code: "Unknown", message: "unexpected route" } });
      });
      const { overrides } = overridesFor(mode, plain);
      const result = await executeArgv(["entra", "data-policy-operation", "show", "--id", "plain-1", "--profile", profile], overrides);
      assert.equal(result.dataPolicyOperation.storageLocation, "***redacted***");
      assert.ok(!JSON.stringify(result).includes("contoso.blob.core.windows.net"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped operation list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "data-policy-operation", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.dataPolicyOperations.map(row => row.id), [op1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "data-policy-operation", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.dataPolicyOperations.map(row => row.id), [op2.id, op3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts operations as one scalar without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "data-policy-operation", "count", "--profile", profile], overrides);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("Workforce tenant context only")));
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/dataPolicyOperations/$count");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied operation reads surface scope, role and licence guidance`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      for (const args of [
        ["entra", "data-policy-operation", "list", "--profile", profile],
        ["entra", "data-policy-operation", "show", "--id", op1.id, "--profile", profile],
        ["entra", "data-policy-operation", "count", "--profile", profile],
      ]) {
        await assert.rejects(executeArgv(args, overrides), error => {
          assert.equal(error.code, "GRAPH_ERROR");
          assert.match(error.message, /\(403\)/);
          assert.match(error.message, /grant, role, licence/);
          assert.ok(error.suggestions.some(hint => hint.includes("User.Export.All")));
          assert.ok(error.suggestions.some(hint => hint.includes("User.Read.All")));
          assert.ok(error.suggestions.some(hint => hint.includes("Company Administrator")));
          assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
          assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
          return true;
        });
      }
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown operation ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "data-policy-operation", "show", "--id", "op-missing", "--profile", profile], overrides),
        error => error.code === "GRAPH_ERROR" && /not found or inaccessible/.test(error.message),
      );
    } finally {
      teardownProfiles(state);
    }
  });

  for (const [command, args] of [
    [["data-policy-operation", "list"], []],
    [["data-policy-operation", "show"], ["--id", op1.id]],
    [["data-policy-operation", "count"], []],
  ]) {
    for (const preview of [false, true]) {
      test(`${mode} ${command.join(" ")} refuses beta before credentials with preview=${preview}`, async () => {
        const state = setupProfiles();
        try {
          const path = join(state.dir, "config.json");
          const config = JSON.parse(readFileSync(path, "utf8"));
          config.profiles[profile].preview = preview;
          writeFileSync(path, JSON.stringify(config));
          const { calls, requests, overrides } = overridesFor(mode);
          await assert.rejects(
            executeArgv(["entra", ...command, ...args, "--profile", profile, "--api-version", "beta"], overrides),
            { code: "VALIDATION_ERROR", message: "Data-policy-operation reads support v1.0 only; beta needs its own review" },
          );
          assert.equal(calls.length, 0);
          assert.equal(requests.length, 0);
        } finally {
          teardownProfiles(state);
        }
      });
    }
  }
}

test("truncated operation text carries a --full hint without lifting redaction or caps", async () => {
  const state = setupProfiles();
  try {
    const longRow = transport(() => json(200, { value: [{ id: op1.id, status: "complete", userId: longUserId, submittedDateTime: op1.submittedDateTime }] }));
    const { overrides } = overridesFor("delegated", longRow);
    const partial = await executeArgv(["entra", "data-policy-operation", "list", "--profile", "soc"], overrides);
    const truncated = partial.dataPolicyOperations.find(row => row.id === op1.id);
    assert.match(truncated.userId, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "data-policy-operation", "list", "--profile", "soc", "--full"], overrides);
    assert.equal(full.dataPolicyOperations.find(row => row.id === op1.id).userId, longUserId);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("unknown properties and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "data-policy-operation", "list", "--profile", "soc", "--select", "id,bogus"], overrides),
      /Unknown data-policy-operation property bogus/,
    );
    await assert.rejects(
      executeArgv(["entra", "data-policy-operation", "list", "--profile", "soc", "--fields", "progress"], overrides),
      /was not fetched; request it with --select/,
    );
    await assert.rejects(
      executeArgv(["entra", "data-policy-operation", "show", "--id", op1.id, "--profile", "soc", "--select", "id,zone"], overrides),
      /Unknown data-policy-operation property zone/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("operation read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "data-policy-operation", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "data-policy-operation", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "data-policy-operation", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "data-policy-operation", "list", "--filter", "status eq 'complete'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "data-policy-operation", "count", "--filter", "status eq 'complete'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "data-policy-operation", "count", "--select", "id"]), /unknown flag --select/);
  await assert.rejects(executeArgv(["entra", "data-policy-operation", "count", "--limit", "10"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "data-policy-operation", "list", "--limit", "10", "--all"]), { code: "VALIDATION_ERROR" });
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { calls, requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "data-policy-operation", "list", "--profile", "batch", "--scopes", policyScopes[0]], overrides),
      { code: "VALIDATION_ERROR", message: "Application profiles use the configured Graph .default audience; delegated scopes are unavailable" },
    );
    await assert.rejects(
      executeArgv(["entra", "data-policy-operation", "count", "--profile", "batch", "--scopes", policyScopes[0]], overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("malformed operation count bodies fail as unknown, not zero", async () => {
  const state = setupProfiles();
  try {
    for (const body of ["{}", "-1", "2.5", "\"3\""]) {
      const malformed = transport(() => ({ status: 200, headers: {}, body }));
      const scoped = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "data-policy-operation", "count", "--profile", "soc"], scoped.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /non-numeric success body/.test(error.message);
      });
    }
  } finally {
    teardownProfiles(state);
  }
});

test("malformed single-operation bodies fail as unknown, not empty", async () => {
  const state = setupProfiles();
  try {
    for (const body of [null, [], "3"]) {
      const malformed = transport(() => json(200, body));
      const scoped = overridesFor("delegated", malformed);
      await assert.rejects(
        executeArgv(["entra", "data-policy-operation", "show", "--id", op1.id, "--profile", "soc"], scoped.overrides),
        { code: "GRAPH_ERROR", message: "Graph returned a malformed data-policy-operation body" },
      );
    }
  } finally {
    teardownProfiles(state);
  }
});

for (const id of ["$count", "$value", "$ref"]) {
  test(`data-policy-operation show rejects reserved resource binding ${id} before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { calls, requests, overrides } = overridesFor("delegated");
      await assert.rejects(
        executeArgv(["entra", "data-policy-operation", "show", "--id", id, "--profile", "soc"], overrides),
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
const rawScopes = policyScopes.join(",");

function rawBody(request) {
  const url = new URL(request.url);
  if (url.pathname === "/v1.0/dataPolicyOperations") return { status: 200, headers: {}, body: JSON.stringify({ value: operations }) };
  if (url.pathname === `/v1.0/dataPolicyOperations/${op1.id}`) return { status: 200, headers: {}, body: JSON.stringify(op1) };
  return { status: 404, headers: {}, body: JSON.stringify({ error: { code: "Unknown", message: "unexpected route" } }) };
}

test("raw api serves reviewed operation list and show with redacted blob URLs", async () => {
  const f = rawFixture(rawBody);
  const list = await runApiGet({ path: "/dataPolicyOperations", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes }, f.deps);
  assert.equal(list.returned, 3);
  assert.equal(list.complete, true);
  assert.ok(!JSON.stringify(list.value).includes("contoso.blob.core.windows.net"));
  const show = await runApiGet({ path: `/dataPolicyOperations/${op1.id}`, apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes }, f.deps);
  assert.equal(show.id, op1.id);
  assert.equal(show.storageLocation, "***redacted***");
  assert.ok(!JSON.stringify(show).includes("fixture-signature"));
});

test("raw api refuses the operation $count scalar before credentials", async () => {
  const f = rawFixture(rawBody);
  await assert.rejects(
    runApiGet({ path: "/dataPolicyOperations/$count", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes }, f.deps),
    { code: "VALIDATION_ERROR" },
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});

test("raw api refuses unreviewed operation $select fields before credentials", async () => {
  const f = rawFixture(rawBody);
  await assert.rejects(
    runApiGet({ path: "/dataPolicyOperations", apiVersion: "v1.0", profile: delegatedRawProfile, scopes: rawScopes, odata: "$select=id,clientSecret" }, f.deps),
    error => error.code === "VALIDATION_ERROR" && /Unreviewed \$select field clientSecret/.test(error.message),
  );
  assert.equal(f.credentialCalls.length, 0);
  assert.equal(f.requests.length, 0);
});
