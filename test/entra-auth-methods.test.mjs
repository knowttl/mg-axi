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
const userId = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";

const m1 = {
  "@odata.type": "#microsoft.graph.phoneAuthenticationMethod",
  id: "b5e01f81-1f81-b5e0-811f-e0b5811fe0b5",
  createdDateTime: "2024-11-27T23:12:49Z",
  displayName: "Mobile phone",
  phoneNumber: "+1 5550100",
  phoneType: "mobile",
  smsSignInState: "allowedByPolicy",
};
const m2 = {
  "@odata.type": "#microsoft.graph.fido2AuthenticationMethod",
  id: "-2_GRUg2-HYz6_1YG4YRAQ2",
  createdDateTime: "2024-08-10T06:44:09Z",
  displayName: `Red key ${"x".repeat(600)}`,
  aaGuid: "2fc0579f-8113-47ea-b116-555a8db9202a",
};
const m3 = {
  "@odata.type": "#microsoft.graph.emailAuthenticationMethod",
  id: "e3e01f81-1f81-b5e0-811f-e0b5811fe0b6",
  createdDateTime: null,
  emailAddress: "soc-analyst@example.com",
};
const methods = [m1, m2, m3];

const d1 = {
  id: "86462606-fde0-4fc4-9e0c-a20eb73e54c6",
  userPrincipalName: "alexw@example.com",
  userDisplayName: "Alex Wilber",
  userType: "member",
  isAdmin: false,
  isMfaRegistered: true,
  isMfaCapable: true,
  isPasswordlessCapable: false,
  isSsprRegistered: false,
  isSsprEnabled: false,
  isSsprCapable: false,
  userPreferredMethodForSecondaryAuthentication: "push",
  lastUpdatedDateTime: "2024-03-13T19:15:41.6195833Z",
};
const d2 = {
  id: "c6ad1942-4afa-47f8-8d48-afb5d8d69d2f",
  userPrincipalName: "alland@example.com",
};
const report = [d1, d2];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-04-"));
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

function authTransport(denied = false) {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (denied) return json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } });
    if (/^\/v1\.0\/users\/[^/]+\/authentication\/methods$/.test(path)) {
      if (url.searchParams.has("$select")) return json(400, { error: { code: "BadRequest", message: "Unsupported $select" } });
      return json(200, { value: methods });
    }
    if (path === "/v1.0/reports/authenticationMethods/userRegistrationDetails") return json(200, { value: report });
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? authTransport();
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

function runAuthCli(args, state, mode, scopes) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-auth-methods-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, scopes, methods, report }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists one user's methods with compact rows naming the method kind`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "user", "authentication-method", "list", "--user", userId, "--profile", profile], overrides);
      assert.deepEqual(result.authenticationMethods, [
        { id: m1.id, displayName: "Mobile phone", createdDateTime: m1.createdDateTime, "@odata.type": m1["@odata.type"] },
        { id: m2.id, displayName: `${m2.displayName.slice(0, 500)}... (truncated, ${m2.displayName.length} chars total)`,
          createdDateTime: m2.createdDateTime, "@odata.type": m2["@odata.type"] },
        { id: m3.id, createdDateTime: null, "@odata.type": m3["@odata.type"] },
      ]);
      assert.deepEqual(result.count, "3 authentication methods");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("never a scan across users")));
      assert.ok(result.help.some(hint => hint.includes("entra registration list")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      const url = new URL(requests[0].url);
      assert.ok(url.pathname.endsWith(`/users/${userId}/authentication/methods`));
      assert.equal(url.searchParams.has("$select"), false);
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} redacts method phone numbers in output and resume cursors`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "user", "authentication-method", "list",
        "--user", userId, "--profile", profile, "--select", "id,displayName,phoneNumber,phoneType", "--limit", "2"], overrides);
      assert.equal(first.authenticationMethods.length, 2);
      assert.equal(first.authenticationMethods[0].phoneNumber, "***redacted***");
      assert.equal(first.authenticationMethods[0].phoneType, "mobile");
      assert.deepEqual(first.count, "2 authentication methods shown, more available");
      assert.equal(first.total, null);
      assert.equal(first.complete, false);
      assert.ok(typeof first.cursor === "string" && first.cursor.length > 0);
      assert.ok(!JSON.stringify(first).includes("+1 5550100"));
      assert.equal(new URL(requests[0].url).searchParams.has("$select"), false);
      const second = await executeArgv(["entra", "user", "authentication-method", "list",
        "--user", userId, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.ok(!JSON.stringify(second).includes("+1 5550100"));
      assert.deepEqual(second.count, "1 authentication methods");
      assert.equal(second.total, null);
      assert.equal(second.complete, true);
      assert.deepEqual(second.authenticationMethods, [{ id: m3.id, "@odata.type": m3["@odata.type"] }]);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes method selection locally across a continuation request`, async () => {
    const state = setupProfiles();
    try {
      const path = `https://graph.microsoft.com/v1.0/users/${userId}/authentication/methods`;
      const fixture = transport(request => {
        const url = new URL(request.url);
        assert.equal(url.searchParams.has("$select"), false);
        return url.searchParams.has("$skiptoken")
          ? json(200, { value: [m3] })
          : json(200, { value: [m1], "@odata.nextLink": `${path}?$skiptoken=next&$select=id,emailAddress` });
      });
      const { requests, overrides } = overridesFor(mode, fixture);
      const first = await executeArgv(["entra", "user", "authentication-method", "list",
        "--user", userId, "--profile", profile, "--select", "id,emailAddress", "--limit", "1"], overrides);
      const second = await executeArgv(["entra", "user", "authentication-method", "list",
        "--user", userId, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.equal(requests.length, 2);
      assert.deepEqual(second.authenticationMethods, [{ id: m3.id, emailAddress: m3.emailAddress, "@odata.type": m3["@odata.type"] }]);
      assert.deepEqual(second.count, "1 authentication methods");
      assert.equal(second.total, null);
      assert.equal(second.complete, true);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} method reads reject --filter before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "user", "authentication-method", "list",
        "--user", userId, "--profile", profile, "--filter", "phoneType eq 'mobile'"], overrides),
      error => error.code === "VALIDATION_ERROR");
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists the registration report with the disabled-user gap explicit`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "registration", "list", "--profile", profile], overrides);
      assert.deepEqual(result.registrationDetails, [
        { id: d1.id, userPrincipalName: d1.userPrincipalName, userDisplayName: d1.userDisplayName, isMfaRegistered: true },
        { id: d2.id, userPrincipalName: d2.userPrincipalName },
      ]);
      assert.deepEqual(result.count, "2 registration rows");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("does not cover disabled users")));
      assert.ok(result.help.some(hint => hint.includes("not proof of no MFA")));
      const url = new URL(requests[0].url);
      assert.equal(url.pathname, "/v1.0/reports/authenticationMethods/userRegistrationDetails");
      assert.equal(url.searchParams.get("$select"), "id,userPrincipalName,userDisplayName,isMfaRegistered");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied method reads name the authentication role requirement`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, authTransport(true));
      await assert.rejects(
        executeArgv(["entra", "user", "authentication-method", "list", "--user", userId, "--profile", profile], overrides),
        error => {
          assert.equal(error.code, "GRAPH_ERROR");
          const text = [error.message, ...error.suggestions].join("\n");
          assert.match(text, /UserAuthenticationMethod\.Read\.All/);
          assert.match(text, /Authentication Administrator/);
          return true;
        },
      );
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied report reads name the report role requirement`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, authTransport(true));
      await assert.rejects(
        executeArgv(["entra", "registration", "list", "--profile", profile], overrides),
        error => {
          assert.equal(error.code, "GRAPH_ERROR");
          const text = [error.message, ...error.suggestions].join("\n");
          assert.match(text, /AuditLog\.Read\.All/);
          assert.match(text, /Reports Reader/);
          return true;
        },
      );
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists user methods end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runAuthCli(
        ["entra", "user", "authentication-method", "list", "--user", userId, "--profile", profile],
        state, mode, "https://graph.microsoft.com/UserAuthenticationMethod.Read.All");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.authenticationMethods.length, 3);
      assert.deepEqual(output.count, "3 authentication methods");
      assert.equal(output.total, null);
      assert.equal(output.complete, true);
      assert.ok(!result.stdout.includes("+1 5550100"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists the registration report end to end on the fake transport`, async () => {
    const state = setupProfiles();
    try {
      const result = runAuthCli(
        ["entra", "registration", "list", "--profile", profile],
        state, mode, "https://graph.microsoft.com/AuditLog.Read.All");
      assert.equal(result.error, undefined);
      assert.equal(result.status, 0, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.deepEqual(output.registrationDetails, [
        { id: d1.id, userPrincipalName: d1.userPrincipalName, userDisplayName: d1.userDisplayName, isMfaRegistered: true },
        { id: d2.id, userPrincipalName: d2.userPrincipalName },
      ]);
      assert.deepEqual(output.count, "2 registration rows");
      assert.equal(output.total, null);
      assert.equal(output.complete, true);
      assert.ok(JSON.stringify(output.help).includes("disabled users"));
    } finally {
      teardownProfiles(state);
    }
  });

  for (const [command, property] of [
    [["entra", "user", "authentication-method", "list", "--user", userId], "aaGuid"],
    [["entra", "registration", "list"], "methodsRegistered"],
  ]) {
    for (const flag of ["select", "fields"]) {
      test(`${mode} ${command.join(" ")} rejects ${property} in --${flag} before credentials`, async () => {
        const state = setupProfiles();
        try {
          const { requests, calls, overrides } = overridesFor(mode);
          await assert.rejects(
            executeArgv([...command, "--profile", profile, `--${flag}`, `id,${property}`], overrides),
            error => {
              assert.equal(error.code, "VALIDATION_ERROR");
              assert.equal(error.message, `Unknown property ${property} in --${flag}`);
              return true;
            },
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

test("delegated method reads truncate long display names with a --full hint", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "user", "authentication-method", "list", "--user", userId, "--profile", "soc"], overrides);
    const key = result.authenticationMethods.find(row => row.id === m2.id);
    assert.match(key.displayName, /truncated, \d+ chars total/);
    assert.ok(result.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "user", "authentication-method", "list", "--user", userId, "--profile", "soc", "--full"], overrides);
    assert.equal(full.authenticationMethods.find(row => row.id === m2.id).displayName, m2.displayName);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated registration filters pass through as plain $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "registration", "list", "--profile", "soc",
      "--filter", "isMfaRegistered eq false"], overrides);
    assert.equal(result.registrationDetails.length, 2);
    const url = new URL(requests[0].url);
    assert.equal(url.searchParams.get("$filter"), "isMfaRegistered eq false");
    assert.equal(url.searchParams.has("$count"), false);
    assert.equal(requests[0].headers.ConsistencyLevel, undefined);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated method reads require --user before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "user", "authentication-method", "list", "--profile", "soc"], overrides),
      /--user is required/,
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated registration reads resume a capped list through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "registration", "list", "--profile", "soc", "--limit", "1"], overrides);
    assert.equal(first.registrationDetails.length, 1);
    assert.deepEqual(first.count, "1 registration rows shown, more available");
    assert.equal(first.total, null);
    assert.equal(first.complete, false);
    assert.ok(typeof first.cursor === "string" && first.cursor.length > 0);
    const second = await executeArgv(["entra", "registration", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.registrationDetails, [{ id: d2.id, userPrincipalName: d2.userPrincipalName }]);
    assert.deepEqual(second.count, "1 registration rows");
    assert.equal(second.total, null);
    assert.equal(second.complete, true);
    assert.ok(second.help.some(hint => hint.includes("does not cover disabled users")));
  } finally {
    teardownProfiles(state);
  }
});
