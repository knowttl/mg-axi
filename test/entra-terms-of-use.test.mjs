import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const tenant = "33333333-3333-4333-8333-333333333333";
const client = "44444444-4444-4222-8222-444444444444";
const agreementScopes = ["https://graph.microsoft.com/Agreement.Read.All"];
const acceptanceScopes = ["https://graph.microsoft.com/AgreementAcceptance.Read"];

// Fixture shapes follow the documented v1.0 example responses; agreement3
// is a synthetic bare row exercising null/missing preservation, and the
// long display name below exercises the truncation-marker convention.
// Acceptance identifying fields (userEmail and friends) stay out of default
// rows because acceptance records are personal data.
const agreement1 = {
  id: "0ec9f6a6-159d-4dd8-a563-1f0b5935e80b",
  displayName: "All users terms of use",
  isPerDeviceAcceptanceRequired: false,
  isViewingBeforeAcceptanceRequired: false,
  termsExpiration: null,
  userReacceptRequiredFrequency: "P90D",
};
const agreement2 = {
  id: "920f5775-d5d7-454b-861f-14686b244e2c",
  displayName: "Guest users terms of use",
  isPerDeviceAcceptanceRequired: false,
  isViewingBeforeAcceptanceRequired: true,
  termsExpiration: { frequency: "monthly", startDateTime: "2022-03-04T14:11:22.6658376Z" },
  userReacceptRequiredFrequency: null,
};
const agreement3 = { id: "94410bbf-3d3e-4683-8149-f034e55c39dd" };
const agreements = [agreement1, agreement2, agreement3];
const acceptance1 = {
  id: "94410bbf-3d3e-4683-8149-f034e55c39dd_d4bb5206-77bf-4d5c-96b4-cf7b0ed3be98",
  agreementId: agreement1.id,
  agreementFileId: "08033369-8972-42a3-8533-90bbd2757a01",
  userId: "d4bb5206-77bf-4d5c-96b4-cf7b0ed3be98",
  deviceId: "00000000-0000-0000-0000-000000000000",
  deviceDisplayName: null,
  deviceOSType: null,
  deviceOSVersion: null,
  userDisplayName: "Megan Bowen",
  userPrincipalName: "MeganB@contoso.example",
  userEmail: "MeganB@contoso.example",
  recordedDateTime: "2022-03-04T14:11:22.6658376Z",
  expirationDateTime: null,
  state: "accepted",
};
const acceptance2 = {
  id: "920f5775-d5d7-454b-861f-14686b244e2c_5b5b5206-77bf-4d5c-96b4-cf7b0ed3be99",
  agreementId: agreement2.id,
  state: "declined",
  recordedDateTime: "2022-04-05T10:00:00Z",
};
const acceptances = [acceptance1, acceptance2];
const minimalAcceptance1 = {
  id: acceptance1.id,
  agreementId: acceptance1.agreementId,
  state: acceptance1.state,
  recordedDateTime: acceptance1.recordedDateTime,
};
const minimalAcceptance2 = {
  id: acceptance2.id,
  agreementId: acceptance2.agreementId,
  state: acceptance2.state,
  recordedDateTime: acceptance2.recordedDateTime,
};
const longName = `terms-${"t".repeat(600)}`;

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-terms-of-use-"));
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

function termsOfUseTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/agreements") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [agreement3] });
      return json(200, {
        value: [agreement1, agreement2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/agreements?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/agreementAcceptances") return json(200, { value: acceptances });
    const nestedList = /^\/v1\.0\/agreements\/([^/]+)\/acceptances$/.exec(path);
    if (nestedList) {
      return json(200, { value: acceptances.filter(row => row.agreementId === decodeURIComponent(nestedList[1])) });
    }
    const nestedShow = /^\/v1\.0\/agreements\/([^/]+)\/acceptances\/([^/]+)$/.exec(path);
    if (nestedShow) {
      const found = acceptances.find(row => row.agreementId === decodeURIComponent(nestedShow[1]) && row.id === decodeURIComponent(nestedShow[2]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such acceptance" } });
    }
    const agreement = /^\/v1\.0\/agreements\/([^/]+)$/.exec(path);
    if (agreement) {
      const found = agreements.find(row => row.id === decodeURIComponent(agreement[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such agreement" } });
    }
    const acceptance = /^\/v1\.0\/agreementAcceptances\/([^/]+)$/.exec(path);
    if (acceptance) {
      const found = acceptances.find(row => row.id === decodeURIComponent(acceptance[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such acceptance" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? termsOfUseTransport();
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

function runTermsCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-terms-of-use-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, agreements: [agreement1, agreement2], acceptances, denied }),
    },
  });
}

test("delegated lists agreements with compact rows preserving null and missing", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "agreement", "list", "--profile", "soc"], overrides);
    assert.deepEqual(result.agreements, [
      { id: agreement1.id, displayName: agreement1.displayName },
      { id: agreement2.id, displayName: agreement2.displayName },
      { id: agreement3.id },
    ]);
    assert.deepEqual(result.count, { returned: 3, complete: true });
    assert.ok(result.help.some(hint => hint.includes("entra agreement show --id <agreement-id>")));
    assert.ok(requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-delegated-token"));
    assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/agreements?"));
    const sent = new URL(requests[0].url).searchParams;
    assert.equal(sent.get("$select"), "id,displayName");
    assert.ok(!sent.has("$filter"));
    assert.ok(!JSON.stringify(result).includes("opaque-fixture-delegated-token"));
    assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(agreementScopes)));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated agreement list passes --filter through as plain $filter", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "agreement", "list", "--profile", "soc", "--filter", "displayName eq 'All users terms of use'"], overrides);
    assert.deepEqual(result.count, { returned: 3, complete: true });
    assert.equal(new URL(requests[0].url).searchParams.get("$filter"), "displayName eq 'All users terms of use'");
  } finally {
    teardownProfiles(state);
  }
});

test("delegated resumes a capped agreement list through its opaque cursor", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "agreement", "list", "--profile", "soc", "--limit", "1"], overrides);
    assert.deepEqual(first.agreements.map(row => row.id), [agreement1.id]);
    assert.equal(first.count.complete, false);
    assert.equal(typeof first.cursor, "string");
    const second = await executeArgv(["entra", "agreement", "list", "--profile", "soc", "--cursor", first.cursor], overrides);
    assert.deepEqual(second.agreements.map(row => row.id), [agreement2.id, agreement3.id]);
    assert.deepEqual(second.count, { returned: 2, complete: true });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated shows one agreement with the full reviewed metadata set", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "agreement", "show", "--id", agreement1.id, "--profile", "soc"], overrides);
    assert.deepEqual(result.agreement, agreement1);
    assert.ok(result.help.some(hint => hint.includes(`entra agreement acceptance list --agreement ${agreement1.id}`)));
    const missing = await executeArgv(["entra", "agreement", "show", "--id", agreement3.id, "--profile", "soc"], overrides);
    assert.deepEqual(missing.agreement, { id: agreement3.id });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated agreement show has no --filter flag to misuse", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "agreement", "show", "--id", agreement1.id, "--profile", "soc", "--filter", "id eq 'x'"], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /unknown flag --filter/.test(error.message);
      },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated lists one agreement's acceptances with minimal personal-data rows", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "agreement", "acceptance", "list", "--agreement", agreement1.id, "--profile", "soc"], overrides);
    assert.deepEqual(result.agreementAcceptances, [minimalAcceptance1]);
    assert.deepEqual(result.count, { returned: 1, complete: true });
    assert.ok(!JSON.stringify(result).includes("MeganB@contoso.example"));
    assert.ok(result.help.some(hint => hint.includes("entra agreement acceptance show")));
    assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/agreements/${agreement1.id}/acceptances?`));
    const sent = new URL(requests[0].url).searchParams;
    assert.equal(sent.get("$select"), "id,agreementId,state,recordedDateTime");
    assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(acceptanceScopes)));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated lists tenant-wide acceptances with minimal personal-data rows", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "agreement-acceptance", "list", "--profile", "soc"], overrides);
    assert.deepEqual(result.agreementAcceptances, [minimalAcceptance1, minimalAcceptance2]);
    assert.deepEqual(result.count, { returned: 2, complete: true });
    assert.ok(!JSON.stringify(result).includes("MeganB@contoso.example"));
    assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/agreementAcceptances?"));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated shows one acceptance with minimal rows unless identifying fields are selected", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const nested = await executeArgv(["entra", "agreement", "acceptance", "show",
      "--agreement", agreement1.id, "--id", acceptance1.id, "--profile", "soc"], overrides);
    assert.deepEqual(nested.agreementAcceptance, minimalAcceptance1);
    const root = await executeArgv(["entra", "agreement-acceptance", "show", "--id", acceptance1.id, "--profile", "soc"], overrides);
    assert.deepEqual(root.agreementAcceptance, minimalAcceptance1);
    const identified = await executeArgv(["entra", "agreement-acceptance", "show", "--id", acceptance1.id, "--profile", "soc",
      "--select", "id,userEmail,userPrincipalName"], overrides);
    assert.deepEqual(identified.agreementAcceptance, {
      id: acceptance1.id,
      userEmail: acceptance1.userEmail,
      userPrincipalName: acceptance1.userPrincipalName,
    });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated nested acceptance list requires --agreement before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "agreement", "acceptance", "list", "--profile", "soc"], overrides),
      { code: "VALIDATION_ERROR", message: "--agreement is required" },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("delegated show projects --fields locally from the fetched --select set", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const result = await executeArgv(["entra", "agreement", "show", "--id", agreement1.id, "--profile", "soc",
      "--select", "id,displayName", "--fields", "displayName"], overrides);
    assert.deepEqual(result.agreement, { displayName: agreement1.displayName });
    await assert.rejects(
      executeArgv(["entra", "agreement", "show", "--id", agreement1.id, "--profile", "soc",
        "--select", "id", "--fields", "displayName"], overrides),
      { code: "VALIDATION_ERROR", message: "--fields displayName was not fetched; request it with --select" },
    );
  } finally {
    teardownProfiles(state);
  }
});

for (const id of ["agreement(id='x')", "agreement)(", "agreement%28"]) {
  test(`delegated agreement show rejects function-style identifier ${id} before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor("delegated");
      await assert.rejects(
        executeArgv(["entra", "agreement", "show", "--id", id, "--profile", "soc"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

for (const command of [
  ["agreement", "list"],
  ["agreement", "show", "--id", agreement1.id],
  ["agreement", "acceptance", "list", "--agreement", agreement1.id],
  ["agreement", "acceptance", "show", "--agreement", agreement1.id, "--id", acceptance1.id],
  ["agreement-acceptance", "list"],
  ["agreement-acceptance", "show", "--id", acceptance1.id],
]) {
  test(`application ${command.join(" ")} is refused before credentials without a supported application permission`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor("application");
      await assert.rejects(
        executeArgv(["entra", ...command, "--profile", "batch"], overrides),
        error => {
          assert.equal(error.code, "VALIDATION_ERROR");
          return /need a delegated profile; Graph documents no supported application permission/.test(error.message);
        },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });
}

test("delegated empty agreement lists stay definitive", async () => {
  const state = setupProfiles();
  try {
    const empty = transport(() => json(200, { value: [] }));
    const { overrides } = overridesFor("delegated", empty);
    const listed = await executeArgv(["entra", "agreement", "list", "--profile", "soc"], overrides);
    assert.deepEqual(listed.agreements, []);
    assert.deepEqual(listed.count, { returned: 0, complete: true });
    assert.ok(listed.help.some(hint => hint.includes("no terms-of-use agreements are configured")));
    const accepted = await executeArgv(["entra", "agreement-acceptance", "list", "--profile", "soc"], overrides);
    assert.deepEqual(accepted.agreementAcceptances, []);
    assert.ok(accepted.help.some(hint => hint.includes("0 agreement acceptances matched")));
  } finally {
    teardownProfiles(state);
  }
});

test("delegated unknown agreement and acceptance ids report absence, not emptiness", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "agreement", "show", "--id", "agreement-missing", "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
      return /not found or inaccessible \(404\)/.test(error.message);
    });
    await assert.rejects(executeArgv(["entra", "agreement-acceptance", "show", "--id", "acceptance-missing", "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      return /not found or inaccessible \(404\)/.test(error.message);
    });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated denied agreement reads surface scope, role and P1 licensing", async () => {
  const state = setupProfiles();
  try {
    const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
    const { overrides } = overridesFor("delegated", denied);
    await assert.rejects(executeArgv(["entra", "agreement", "list", "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.ok(error.suggestions.some(hint => hint.includes("Agreement.Read.All")));
      assert.ok(error.suggestions.some(hint => hint.includes("Security Reader")));
      assert.ok(error.suggestions.some(hint => hint.includes("Microsoft Entra ID P1")));
      assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
      return /grant, role, licence/.test(error.message);
    });
    await assert.rejects(executeArgv(["entra", "agreement", "show", "--id", agreement1.id, "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.ok(error.suggestions.some(hint => hint.includes("Agreement.Read.All")));
      return /grant, role, licence/.test(error.message);
    });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated denied acceptance reads surface scope, role and P1 licensing", async () => {
  const state = setupProfiles();
  try {
    const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
    const { overrides } = overridesFor("delegated", denied);
    await assert.rejects(executeArgv(["entra", "agreement", "acceptance", "list", "--agreement", agreement1.id, "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.ok(error.suggestions.some(hint => hint.includes("AgreementAcceptance.Read")));
      assert.ok(error.suggestions.some(hint => hint.includes("AgreementAcceptance.Read.All")));
      assert.ok(error.suggestions.some(hint => hint.includes("Security Reader")));
      assert.ok(error.suggestions.some(hint => hint.includes("Microsoft Entra ID P1")));
      return /grant, role, licence/.test(error.message);
    });
    await assert.rejects(executeArgv(["entra", "agreement-acceptance", "show", "--id", acceptance1.id, "--profile", "soc"], overrides), error => {
      assert.equal(error.code, "GRAPH_ERROR");
      assert.ok(error.suggestions.some(hint => hint.includes("AgreementAcceptance.Read")));
      return /grant, role, licence/.test(error.message);
    });
  } finally {
    teardownProfiles(state);
  }
});

test("delegated executable lists and shows agreement reads", () => {
  const state = setupProfiles();
  try {
    const listed = runTermsCli(["entra", "agreement", "list", "--profile", "soc"], state, "delegated");
    assert.equal(listed.status, 0, listed.stdout);
    assert.equal(listed.stderr, "");
    const listOut = decode(listed.stdout);
    assert.deepEqual(listOut.agreements.map(row => row.id), [agreement1.id, agreement2.id]);
    assert.deepEqual(listOut.count, { returned: 2, complete: true });
    assert.ok(!listed.stdout.includes("opaque-fixture-delegated-token"));

    const shown = runTermsCli(["entra", "agreement", "show", "--id", agreement1.id, "--profile", "soc"], state, "delegated");
    assert.equal(shown.status, 0, shown.stdout);
    assert.deepEqual(decode(shown.stdout).agreement, {
      id: agreement1.id,
      displayName: agreement1.displayName,
      isPerDeviceAcceptanceRequired: false,
      isViewingBeforeAcceptanceRequired: false,
      termsExpiration: null,
      userReacceptRequiredFrequency: "P90D",
    });
    assert.ok(!shown.stdout.includes("opaque-fixture-delegated-token"));

    const nested = runTermsCli(["entra", "agreement", "acceptance", "list", "--agreement", agreement1.id, "--profile", "soc"], state, "delegated");
    assert.equal(nested.status, 0, nested.stdout);
    assert.deepEqual(decode(nested.stdout).agreementAcceptances, [minimalAcceptance1]);

    const rooted = runTermsCli(["entra", "agreement-acceptance", "list", "--profile", "soc"], state, "delegated");
    assert.equal(rooted.status, 0, rooted.stdout);
    assert.equal(decode(rooted.stdout).count.returned, 2);
  } finally { teardownProfiles(state); }
});

test("delegated executable denied agreement reads fail operationally on stdout", () => {
  const state = setupProfiles();
  try {
    const result = runTermsCli(["entra", "agreement", "list", "--profile", "soc"], state, "delegated", true);
    assert.equal(result.status, 1, result.stdout);
    assert.equal(result.stderr, "");
    const output = decode(result.stdout);
    assert.equal(output.code, "GRAPH_ERROR");
    assert.match(output.error, /grant, role, licence or policy/);
    assert.equal(output.agreements, undefined);
    assert.ok(!result.stdout.includes("opaque-fixture-delegated-token"));
  } finally { teardownProfiles(state); }
});

for (const [command, args] of [
  [["agreement", "list"], []],
  [["agreement", "show"], ["--id", agreement1.id]],
  [["agreement", "acceptance", "list"], ["--agreement", agreement1.id]],
  [["agreement", "acceptance", "show"], ["--agreement", agreement1.id, "--id", acceptance1.id]],
  [["agreement-acceptance", "list"], []],
  [["agreement-acceptance", "show"], ["--id", acceptance1.id]],
]) {
  for (const preview of [false, true]) {
    test(`delegated ${command.join(" ")} refuses beta before credentials with preview=${preview}`, async () => {
      const state = setupProfiles();
      try {
        const path = join(state.dir, "config.json");
        const config = JSON.parse(readFileSync(path, "utf8"));
        config.profiles.soc.preview = preview;
        writeFileSync(path, JSON.stringify(config));
        const { calls, requests, overrides } = overridesFor("delegated");
        await assert.rejects(
          executeArgv(["entra", ...command, ...args, "--profile", "soc", "--api-version", "beta"], overrides),
          { code: "VALIDATION_ERROR", message: "Terms-of-use reads support v1.0 only; beta needs its own review" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }
}

test("truncated agreement text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const longRow = transport(() => json(200, { value: [{ id: "agreement-long", displayName: longName }] }));
    const { overrides } = overridesFor("delegated", longRow);
    const partial = await executeArgv(["entra", "agreement", "list", "--profile", "soc"], overrides);
    const truncated = partial.agreements.find(row => row.displayName.startsWith("terms-"));
    assert.match(truncated.displayName, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "agreement", "list", "--profile", "soc", "--full"], overrides);
    assert.equal(full.agreements.find(row => row.displayName.startsWith("terms-")).displayName, longName);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("raw agreement reads serve delegated callers and refuse application profiles", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const { ApplicationAuth } = await import("../dist/app-auth.js");
  const credentialCalls = [];
  const requests = [];
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async (...args) => { credentialCalls.push(args); return credential; } }),
    application: new ApplicationAuth({ storage: "session-only", acquire: async (...args) => { credentialCalls.push(args); return credential; } }),
    transport: async request => { requests.push(request); return { headers: {}, body: JSON.stringify({ value: [agreement1] }) }; },
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const appProfile = { ...delegatedProfile, mode: "application", credentialRef: { provider: "federated", key: "55555555-5555-4555-8555-555555555555" } };
  const listed = await runApiGet({ path: "/agreements", apiVersion: "v1.0", profile: delegatedProfile, scopes: agreementScopes[0] }, deps);
  assert.equal(listed.returned, 1);
  for (const path of ["/agreements", `/agreements/${agreement1.id}`, `/agreements/${agreement1.id}/acceptances`, `/agreements/${agreement1.id}/acceptances/${acceptance1.id}`, "/agreementAcceptances", `/agreementAcceptances/${acceptance1.id}`]) {
    await assert.rejects(
      runApiGet({ path, apiVersion: "v1.0", profile: appProfile }, deps),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /needs a delegated profile/.test(error.message);
      },
    );
  }
  assert.equal(credentialCalls.length, 1);
  assert.equal(requests.length, 1);
});

test("raw acceptance reads project only reviewed fields with minimal defaults", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const { ApplicationAuth } = await import("../dist/app-auth.js");
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async () => credential }),
    application: new ApplicationAuth({ storage: "session-only", acquire: async () => credential }),
    transport: async () => ({ status: 200, headers: {}, body: JSON.stringify({ ...acceptance1, fileData: "QUJD", unknownField: true }) }),
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const shown = await runApiGet({ path: `/agreementAcceptances/${acceptance1.id}`, apiVersion: "v1.0", profile: delegatedProfile, scopes: acceptanceScopes[0] }, deps);
  assert.deepEqual(shown, minimalAcceptance1);
});
