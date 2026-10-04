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
import { DelegatedAuth } from "../dist/auth.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const delegatedScopes = ["https://graph.microsoft.com/User.Read.All"];

const u1 = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  displayName: "Adele Vance",
  userPrincipalName: "AdeleV@contoso.com",
  mail: "AdeleV@contoso.com",
  givenName: null,
  jobTitle: "Engineer",
  officeLocation: "18/2111",
  businessPhones: ["425-555-0100"],
  mobilePhone: null,
  preferredLanguage: "en-US",
  surname: "Vance",
};
const u2 = {
  id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  displayName: "Alex Wilber",
  userPrincipalName: "AlexW@contoso.com",
  mail: null,
  givenName: "Alex",
  jobTitle: null,
  officeLocation: null,
  businessPhones: [],
  mobilePhone: "425-555-0101",
  preferredLanguage: null,
  surname: "Wilber",
};
const u3 = {
  id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  displayName: "Grady Archie",
  userPrincipalName: "GradyA@contoso.com",
  givenName: "Grady",
  jobTitle: "Designer",
  officeLocation: "19/3122",
  businessPhones: [],
  mobilePhone: null,
  preferredLanguage: "en-US",
  surname: "Archie",
};
const directory = [u1, u2, u3];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-01-"));
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

function directoryTransport(requests) {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/users" || path === "/beta/users") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [u3] });
      const host = "https://graph.microsoft.com";
      return json(200, {
        value: [u1, u2],
        "@odata.nextLink": `${host}${path}?%24skiptoken=page2`,
      });
    }
    const match = /^\/(v1\.0|beta)\/users\/([^/]+)$/.exec(path);
    if (match) {
      const found = directory.find(user => user.id === decodeURIComponent(match[2]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such user" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? directoryTransport();
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

// Parse a recovery hint back into argv without a shell: hints render values
// with POSIX single-quote escaping, and Windows has no /bin/sh to delegate
// to. The child receives the words as an array, so no quoting remains.
function hintArgv(hint) {
  const argv = [];
  let word = "";
  let quote = null;
  let escaped = false;
  let started = false;
  const push = () => { if (started) { argv.push(word); word = ""; started = false; } };
  for (const char of hint.slice("mg-axi ".length)) {
    if (escaped) { word += char; escaped = false; }
    else if (quote === "'") { if (char === "'") quote = null; else word += char; }
    else if (quote === '"') {
      if (char === '"') quote = null;
      else if (char === "\\") escaped = true;
      else word += char;
    }
    else if (char === "\\") escaped = true;
    else if (char === "'" || char === '"') { quote = char; started = true; }
    else if (/\s/.test(char)) push();
    else { word += char; started = true; }
  }
  push();
  return argv;
}

function runReadCli(args, state, mode, rows = directory, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, directory: rows, denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} executable lists basic users with null and missing values`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "user", "list", "--profile", profile], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.users, [
        { id: u1.id, displayName: u1.displayName, userPrincipalName: u1.userPrincipalName, mail: u1.mail },
        { id: u2.id, displayName: u2.displayName, userPrincipalName: u2.userPrincipalName, mail: null },
        { id: u3.id, displayName: u3.displayName, userPrincipalName: u3.userPrincipalName },
      ]);
      assert.deepEqual(output.count, { returned: 3, complete: true });
      assert.ok(output.help.length);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable shows richer user properties`, () => {
    const state = setupProfiles();
    try {
      const result = runReadCli(["entra", "user", "show", "--id", u1.id, "--profile", profile], state, mode);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      assert.deepEqual(decode(result.stdout), { user: u1 });
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const full of [false, true]) {
    test(`${mode} executable ${full ? "full view" : "compact view"} preserves scalar and phone-array values`, () => {
      const state = setupProfiles();
      try {
        const long = "x".repeat(600);
        const rows = [{ ...u1, department: long, businessPhones: ["425-555-0100", long] }, u2, u3];
        const args = ["entra", "user", "show", "--id", u1.id, "--profile", profile,
          "--select", "id,department,businessPhones", ...(full ? ["--full"] : [])];
        const result = runReadCli(args, state, mode, rows);
        assert.equal(result.status, 0, result.stdout);
        assert.equal(result.stderr, "");
        const output = decode(result.stdout);
        const expected = full ? long : `${"x".repeat(500)}... (truncated, 600 chars total)`;
        assert.deepEqual(output.user, { id: u1.id, department: expected, businessPhones: ["425-555-0100", expected] });
        assert.equal(output.help?.length ?? 0, full ? 0 : 1);
        assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
      } finally { teardownProfiles(state); }
    });
  }

  for (const [action, args] of [["list", []], ["show", ["--id", u1.id]]]) {
    test(`${mode} executable denied ${action} is an operational error on stdout`, () => {
      const state = setupProfiles();
      try {
        const result = runReadCli(["entra", "user", action, ...args, "--profile", profile], state, mode, directory, true);
        assert.equal(result.status, 1, result.stdout);
        assert.equal(result.stderr, "");
        const output = decode(result.stdout);
        assert.equal(output.code, "GRAPH_ERROR");
        assert.match(output.error, /grant, role, licence or policy/);
        assert.ok(output.help.length);
        assert.equal(output.users, undefined);
        assert.equal(output.user, undefined);
        assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
      } finally { teardownProfiles(state); }
    });
  }

  test(`${mode} executable resumes a capped list through its opaque cursor`, () => {
    const state = setupProfiles();
    try {
      const first = runReadCli(["entra", "user", "list", "--profile", profile, "--limit", "1"], state, mode);
      assert.equal(first.status, 0, first.stdout);
      assert.equal(first.stderr, "");
      const partial = decode(first.stdout);
      assert.equal(partial.count.complete, false);
      assert.equal(typeof partial.cursor, "string");
      const second = runReadCli(["entra", "user", "list", "--profile", profile, "--cursor", partial.cursor], state, mode);
      assert.equal(second.status, 0, second.stdout);
      assert.equal(second.stderr, "");
      const resumed = decode(second.stdout);
      assert.deepEqual([...partial.users, ...resumed.users].map(user => user.id), [u1.id, u2.id, u3.id]);
      assert.deepEqual(resumed.count, { returned: 2, complete: true });
      assert.ok(!second.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable rejects invalid scope arguments as a usage error`, () => {
    const state = setupProfiles();
    try {
      const scopes = mode === "delegated" ? "https://example.invalid/User.Read" : delegatedScopes[0];
      const result = runReadCli(["entra", "user", "list", "--profile", profile, "--scopes", scopes], state, mode);
      assert.equal(result.status, 2, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "VALIDATION_ERROR");
      assert.ok(output.help.length);
    } finally { teardownProfiles(state); }
  });
}

for (const [action, args] of [["list", []], ["show", ["--id", u1.id]]]) {
  for (const scopes of [",", "https://example.invalid/User.Read", "https://graph.microsoft.com/.default"]) {
    test(`${action} rejects delegated scopes ${scopes} before provider acquisition`, async () => {
      const state = setupProfiles();
      try {
        const calls = [];
        const { overrides, requests } = overridesFor("delegated");
        overrides.delegated = new DelegatedAuth({
          storage: "session-only",
          login: async () => { throw new Error("Unexpected interactive login"); },
          silent: async (...args) => { calls.push(args); throw new Error("Unexpected credential acquisition"); },
        });
        await assert.rejects(executeArgv(["entra", "user", action, ...args, "--profile", "soc", "--scopes", scopes], overrides), { code: "VALIDATION_ERROR" });
        assert.deepEqual(calls, []);
        assert.deepEqual(requests, []);
      } finally { teardownProfiles(state); }
    });
  }
}

test("delegated list returns compact basic rows with a complete count and show hint", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "user", "list", "--profile", "soc"], overrides);
    assert.deepEqual(result.users, [
      { id: u1.id, displayName: "Adele Vance", userPrincipalName: "AdeleV@contoso.com", mail: "AdeleV@contoso.com" },
      { id: u2.id, displayName: "Alex Wilber", userPrincipalName: "AlexW@contoso.com", mail: null },
      { id: u3.id, displayName: "Grady Archie", userPrincipalName: "GradyA@contoso.com" },
    ]);
    assert.deepEqual(result.count, { returned: 3, complete: true });
    assert.ok(result.help.some(hint => hint.includes("entra user show --id <user-id-or-upn> --profile soc")));
    assert.equal(requests.length, 2);
    assert.ok(requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-delegated-token"));
    assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/users?"));
    assert.ok(!JSON.stringify(result).includes("opaque-fixture-delegated-token"));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated show returns the suggested user with richer properties", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const listed = await executeArgv(["entra", "user", "list", "--profile", "soc", "--limit", "1"], overrides);
    assert.equal(listed.count.complete, false);
    assert.equal(typeof listed.cursor, "string");
    const result = await executeArgv(["entra", "user", "show", "--id", u1.id, "--profile", "soc"], overrides);
    assert.equal(result.user.id, u1.id);
    assert.equal(result.user.displayName, "Adele Vance");
    assert.equal(result.user.jobTitle, "Engineer");
    assert.equal(result.user.officeLocation, "18/2111");
    assert.equal(result.user.givenName, null);
    assert.equal(result.help, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("show truncates long text unless --full is passed", async () => {
  const state = setupProfiles();
  try {
    const long = "x".repeat(600);
    const fixture = transport(() => json(200, { ...u1, department: long }));
    const credential = credentialService("delegated", []);
    const overrides = { transport: fixture.send, delegated: credential, application: credentialService("application", []) };
    const argv = ["entra", "user", "show", "--id", u1.id, "--profile", "soc", "--select", "id,department",
      "--fields", "department", "--api-version", "v1.0", "--scopes", delegatedScopes[0]];
    const compact = await executeArgv(argv, overrides);
    assert.match(compact.user.department, /\.\.\. \(truncated, 600 chars total\)$/);
    assert.ok(compact.help.some(hint => hint.includes("--full")));
    const recovery = hintArgv(compact.help[0]);
    assert.deepEqual(recovery, [...argv, "--full"]);
    const full = await executeArgv(recovery, overrides);
    assert.equal(full.user.department, long);
    assert.equal(full.help, undefined);
  } finally {
    teardownProfiles(state);
  }
});

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  for (const repeated of [[], ["--select", "id,department"], ["--filter", "department eq 'R&D'"]]) {
    test(`${mode} resume restores omitted query flags with ${JSON.stringify(repeated)}`, async () => {
      const state = setupProfiles();
      try {
        const fixture = transport(request => new URL(request.url).searchParams.has("$skiptoken")
          ? json(200, { value: [{ id: "c", department: "R&D" }] })
          : json(200, {
            value: [{ id: "a", department: "R&D" }, { id: "b", department: null }],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=next",
          }));
        const { overrides } = overridesFor(mode, fixture);
        const first = await executeArgv(["entra", "user", "list", "--profile", profile,
          "--select", "id,department", "--filter", "department eq 'R&D'", "--limit", "1"], overrides);
        const second = await executeArgv(["entra", "user", "list", "--profile", profile,
          "--cursor", first.cursor, ...repeated], overrides);
        assert.deepEqual(second.users, [{ id: "b", department: null }, { id: "c", department: "R&D" }]);
        assert.equal(second.count.complete, true);
        assert.equal(fixture.requests[1].url, "https://graph.microsoft.com/v1.0/users?$skiptoken=next");
        await assert.rejects(executeArgv(["entra", "user", "list", "--profile", profile,
          "--cursor", first.cursor, "--select", "id"], overrides), { code: "VALIDATION_ERROR" });
      } finally {
        teardownProfiles(state);
      }
    });
  }
}

for (const cap of [["--all"], ["--limit", "1"]]) {
  test(`list truncation recovery preserves query and projection with ${cap.join(" ")}`, async () => {
    const state = setupProfiles();
    try {
      const long = "x".repeat(600);
      const fixture = transport(() => json(200, { value: [{ id: "a", department: long }, { id: "b", department: long }] }));
      const { overrides } = overridesFor("delegated", fixture);
      const argv = ["entra", "user", "list", "--profile", "soc", "--select", "id,department",
        "--fields", "department", "--filter", "department eq 'R&D'", "--scopes", delegatedScopes[0], "--api-version", "v1.0", ...cap];
      const compact = await executeArgv(argv, overrides);
      assert.match(compact.users[0].department, /truncated/);
      const recovery = hintArgv(compact.help[0]);
      assert.deepEqual(recovery, [...argv, "--full"]);
      const full = await executeArgv(recovery, overrides);
      assert.deepEqual(full.users, cap[0] === "--all" ? [{ department: long }, { department: long }] : [{ department: long }]);
      assert.equal(full.count.complete, compact.count.complete);
      if (compact.cursor) {
        const resumed = await executeArgv(["entra", "user", "list", "--profile", "soc", "--cursor", compact.cursor], overrides);
        const recovered = await executeArgv(hintArgv(resumed.help[0]), overrides);
        assert.deepEqual(recovered.users, [{ id: "b", department: long }]);
      }
    } finally {
      teardownProfiles(state);
    }
  });
}

test("a capped delegated list resumes losslessly through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "user", "list", "--profile", "soc", "--limit", "2"], overrides);
    assert.deepEqual(first.users.map(user => user.id), [u1.id, u2.id]);
    assert.deepEqual(first.count, { returned: 2, complete: false, reason: first.count.reason });
    assert.match(first.count.reason, /row limit/);
    assert.equal(typeof first.cursor, "string");
    assert.ok(!first.cursor.includes(u1.id));
    const second = await executeArgv(["entra", "user", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.users.map(user => user.id), [u3.id]);
    assert.deepEqual(second.count, { returned: 1, complete: true });
    assert.deepEqual([...first.users, ...second.users].map(user => user.id), [u1.id, u2.id, u3.id]);
    assert.ok(!JSON.stringify(second).includes("opaque-fixture-delegated-token"));
    assert.ok(requests.length >= 2);
  } finally {
    teardownProfiles(state);
  }
});

test("list --all follows pages within budget", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "user", "list", "--profile", "soc", "--all"], overrides);
    assert.deepEqual(result.users.map(user => user.id), [u1.id, u2.id, u3.id]);
    assert.deepEqual(result.count, { returned: 3, complete: true });
  } finally {
    teardownProfiles(state);
  }
});

test("an empty list is a definitive zero, not an error", async () => {
  const state = setupProfiles();
  try {
    const fixture = transport(() => json(200, { value: [] }));
    const credential = credentialService("delegated", []);
    const overrides = { transport: fixture.send, delegated: credential, application: credentialService("application", []) };
    const result = await executeArgv(["entra", "user", "list", "--profile", "soc"], overrides);
    assert.deepEqual(result.users, []);
    assert.deepEqual(result.count, { returned: 0, complete: true });
    assert.ok(result.help.some(hint => hint.includes("0 users matched")));
  } finally {
    teardownProfiles(state);
  }
});

test("denied reads surface as operational failures, never as empty results", async () => {
  const state = setupProfiles();
  try {
    const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
    const credential = credentialService("delegated", []);
    const overrides = { transport: denied.send, delegated: credential, application: credentialService("application", []) };
    await assert.rejects(executeArgv(["entra", "user", "list", "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      return /grant, role, licence/.test(error.message);
    });
    const missing = transport(() => json(404, { error: { code: "Request_ResourceNotFound", message: "gone" } }));
    const missingOverrides = { transport: missing.send, delegated: credential, application: credentialService("application", []) };
    await assert.rejects(executeArgv(["entra", "user", "show", "--id", u1.id, "--profile", "soc"], missingOverrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      return /not found or inaccessible/.test(error.message);
    });
  } finally {
    teardownProfiles(state);
  }
});

test("application mode lists and shows without delegated scopes", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("application");
    const listed = await executeArgv(["entra", "user", "list", "--profile", "batch"], overrides);
    assert.deepEqual(listed.users.map(user => user.id), [u1.id, u2.id, u3.id]);
    assert.equal(calls.length, 1);
    assert.ok(requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-application-token"));
    const shown = await executeArgv(["entra", "user", "show", "--id", u2.id, "--profile", "batch"], overrides);
    assert.equal(shown.user.userPrincipalName, "AlexW@contoso.com");
    assert.equal(shown.user.mail, null);
  } finally {
    teardownProfiles(state);
  }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "user", "list", "--profile", "batch", "--scopes", delegatedScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
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
    await assert.rejects(executeArgv(["entra", "user", "list", "--profile", "soc", "--select", "id,aboutMe"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "user", "list", "--profile", "soc", "--fields", "department"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "user", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  for (const [name, body] of [
    ["buffered-only", { value: [{ id: "a" }, { id: "b" }] }],
    ["nextLink", { value: [{ id: "a" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=next" }],
  ]) {
    test(`${mode} ${name} resume rejects unfetched fields before credentials or HTTP`, async () => {
      const state = setupProfiles();
      try {
        const fixture = transport(() => json(200, body));
        const { calls, requests, overrides } = overridesFor(mode, fixture);
        const first = await executeArgv(["entra", "user", "list", "--profile", profile,
          "--select", "id", "--limit", "1"], overrides);
        assert.equal(first.count.complete, false);
        overrides[mode] = { credential: async (...args) => {
          calls.push(args);
          throw new Error("Authentication is unavailable");
        } };
        await assert.rejects(executeArgv(["entra", "user", "list", "--profile", profile,
          "--cursor", first.cursor, "--fields", "department"], overrides),
        error => error.code === "VALIDATION_ERROR" && /department was not fetched/.test(error.message));
        assert.equal(calls.length, 1);
        assert.equal(requests.length, 1);
      } finally { teardownProfiles(state); }
    });
  }
  for (const [name, body, expectedRequests] of [
    ["buffered-only", { value: [{ id: "a", department: "Sales" }, { id: "b", department: "R&D" }] }, 1],
    ["nextLink", { value: [{ id: "a", department: "Sales" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=next" }, 2],
  ]) {
    test(`${mode} ${name} resume projects fields from the saved selection`, async () => {
      const state = setupProfiles();
      try {
        const fixture = transport(request => json(200, new URL(request.url).searchParams.has("$skiptoken")
          ? { value: [{ id: "b", department: "R&D" }] } : body));
        const { calls, requests, overrides } = overridesFor(mode, fixture);
        const first = await executeArgv(["entra", "user", "list", "--profile", profile,
          "--select", "id,department", "--limit", "1"], overrides);
        const resumed = await executeArgv(["entra", "user", "list", "--profile", profile,
          "--cursor", first.cursor, "--fields", "department"], overrides);
        assert.deepEqual(resumed.users, [{ department: "R&D" }]);
        assert.deepEqual(resumed.count, { returned: 1, complete: true });
        assert.equal(calls.length, 2);
        assert.equal(requests.length, expectedRequests);
      } finally { teardownProfiles(state); }
    });
  }
}

test("richer selects dispatch the same mapping while beta stays preview-gated", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const richer = await executeArgv(
      ["entra", "user", "list", "--profile", "soc", "--select", "id,displayName,accountEnabled", "--all"],
      overrides,
    );
    assert.equal(new URL(requests[0].url).searchParams.get("$select"), "id,displayName,accountEnabled");
    assert.deepEqual(richer.users[0], { id: u1.id, displayName: "Adele Vance" });
    const beta = executeArgv(["entra", "user", "list", "--profile", "soc", "--api-version", "beta", "--limit", "1"], overrides);
    await assert.rejects(beta, error => {
      assert.equal(error.code, "POLICY_DENIED");
      return /preview-enabled/.test(error.message);
    });
  } finally {
    teardownProfiles(state);
  }
});
