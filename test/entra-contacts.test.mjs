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
const contactScopes = ["https://graph.microsoft.com/OrgContact.Read.All"];

// Fixture shapes follow the documented v1.0 get response: flat scalar
// properties plus unreviewed nested data (the manager navigation object and
// the phones collection) that must never be projected because $expand and
// nested collections stay out of this slice. c3 is a bare row exercising
// null/missing preservation.
const c1 = {
  id: "25caf6a2-d5cb-470d-8940-20ba795ef62d",
  deletedDateTime: null,
  companyName: "Adatum Corporation",
  department: "Sales",
  displayName: "Adele Vance",
  givenName: "Adele",
  surname: "Vance",
  mail: "AdeleVance@adatum.com",
  mailNickname: "AdeleVance",
  jobTitle: "Engagement manager",
  onPremisesSyncEnabled: null,
  proxyAddresses: ["SMTP:AdeleVance@adatum.com"],
  manager: { id: "manager-object-id", displayName: "A manager" },
  phones: [{ number: "555-0100", type: "business" }],
};
const c2 = {
  id: "7b2c9d4e-1f5a-4b6c-8d7e-8f9a0b1c2d3e",
  displayName: "Alex Wilber",
  mail: "AlexW@adatum.com",
  companyName: "Adatum Corporation",
};
const c3 = { id: "contact-bare-3" };
const contacts = [c1, c2, c3];
const longName = `Contact ${"C".repeat(600)}`;

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-contacts-"));
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

function contactTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/contacts/$count") {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(contacts.length) };
    }
    if (path === "/v1.0/contacts") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [c3] });
      return json(200, {
        value: [c1, c2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/contacts?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/contacts\/([^/]+)$/.exec(path);
    if (single) {
      const found = contacts.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such contact" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? contactTransport();
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

function runContactCli(args, state, mode, denied = false, nav = undefined) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-contacts-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, contacts: [c1, c2], denied, ...nav }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists contacts with minimal personal-data rows`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "contact", "list", "--profile", profile], overrides);
      assert.deepEqual(result.contacts, [
        { id: c1.id, displayName: c1.displayName, mail: c1.mail, companyName: c1.companyName },
        { id: c2.id, displayName: c2.displayName, mail: c2.mail, companyName: c2.companyName },
        { id: c3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra contact show --id <contact-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/contacts?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      assert.ok(!JSON.stringify(result).includes("A manager"), "navigation objects never reach output");
      assert.ok(!JSON.stringify(result).includes("555-0100"), "nested collections never reach output");
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(contactScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} contact list filters with the documented eventual contract`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "contact", "list", "--profile", profile, "--filter", "startswith(displayName,'A')"], overrides);
      assert.equal(result.count.complete, true);
      const url = new URL(requests[0].url);
      assert.equal(url.searchParams.get("$filter"), "startswith(displayName,'A')");
      assert.equal(url.searchParams.get("$count"), "true");
      assert.equal(requests[0].headers.ConsistencyLevel, "eventual");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped contact list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "contact", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.contacts.map(row => row.id), [c1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "contact", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.contacts.map(row => row.id), [c2.id, c3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one contact with the minimal set and wider selects`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "contact", "show", "--id", c1.id, "--profile", profile], overrides);
      assert.deepEqual(result.contact, { id: c1.id, displayName: c1.displayName, mail: c1.mail, companyName: c1.companyName });
      const wide = await executeArgv(["entra", "contact", "show", "--id", c1.id, "--profile", profile,
        "--select", "id,displayName,jobTitle,proxyAddresses"], overrides);
      assert.deepEqual(wide.contact, { id: c1.id, displayName: c1.displayName, jobTitle: c1.jobTitle, proxyAddresses: c1.proxyAddresses });
      const missing = await executeArgv(["entra", "contact", "show", "--id", c3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.contact, { id: c3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} contact show refuses nested collections before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "contact", "show", "--id", c1.id, "--profile", profile, "--select", "id,phones"], overrides),
        error => {
          assert.equal(error.code, "VALIDATION_ERROR");
          return /Unknown contact property phones/.test(error.message);
        },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts contacts as one scalar with the documented header`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "contact", "count", "--profile", profile], overrides);
      assert.deepEqual(result, { count: { returned: 3, complete: true } });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/contacts/$count");
      assert.equal(requests[0].headers.ConsistencyLevel, "eventual");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} contact count refuses collection flags before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "contact", "count", "--profile", profile, "--filter", "id eq 'x'"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty contact lists stay definitive`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "contact", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.contacts, []);
      assert.deepEqual(listed.count, { returned: 0, complete: true });
      assert.ok(listed.help.some(hint => hint.includes("absence of results is the answer")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown contact ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "contact", "show", "--id", "contact-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied contact reads surface scope, role and account guidance`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "contact", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("OrgContact.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Directory Readers")));
        assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose role or licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "contact", "count", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("OrgContact.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts contact reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runContactCli(["entra", "contact", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.contacts.map(row => row.id), [c1.id, c2.id]);
      assert.deepEqual(listOut.count, { returned: 2, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runContactCli(["entra", "contact", "show", "--id", c1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).contact, { id: c1.id, displayName: c1.displayName, mail: c1.mail, companyName: c1.companyName });

      const counted = runContactCli(["entra", "contact", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).count, { returned: 2, complete: true });
      assert.ok(!counted.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied contact reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runContactCli(["entra", "contact", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.contacts, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["contact", "list"], ["--limit", "1"]],
    [["contact", "show"], ["--id", c1.id]],
    [["contact", "count"], []],
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
            { code: "VALIDATION_ERROR", message: "Contact reads support v1.0 only; beta needs its own review" },
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

test("application contact reads refuse delegated scopes before credentials", async () => {
  const state = setupProfiles();
  try {
    const { requests, calls, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "contact", "list", "--profile", "batch", "--scopes", contactScopes[0]], overrides),
      { code: "VALIDATION_ERROR", message: "Application profiles use the configured Graph .default audience; delegated scopes are unavailable" },
    );
    assert.equal(calls.length, 0);
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("truncated contact text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const longRow = transport(() => json(200, { value: [{ id: "contact-long-1", displayName: longName }] }));
    const { overrides } = overridesFor("delegated", longRow);
    const partial = await executeArgv(["entra", "contact", "list", "--profile", "soc",
      "--select", "id,displayName"], overrides);
    const truncated = partial.contacts.find(row => row.id === "contact-long-1");
    assert.match(truncated.displayName, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "contact", "list", "--profile", "soc", "--full",
      "--select", "id,displayName"], overrides);
    assert.equal(full.contacts.find(row => row.id === "contact-long-1").displayName, longName);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("raw contact reads project reviewed fields and refuse unreviewed ones", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const { ApplicationAuth } = await import("../dist/app-auth.js");
  const credentialCalls = [];
  const requests = [];
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async (...args) => { credentialCalls.push(args); return credential; } }),
    application: new ApplicationAuth({ storage: "session-only", acquire: async (...args) => { credentialCalls.push(args); return credential; } }),
    transport: async request => { requests.push(request); return { headers: {}, body: JSON.stringify({ value: [c1, c2] }) }; },
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const appProfile = { ...delegatedProfile, mode: "application", credentialRef: { provider: "federated", key: "55555555-5555-4555-8555-555555555555" } };
  for (const profile of [delegatedProfile, appProfile]) {
    const listed = await runApiGet({ path: "/contacts", apiVersion: "v1.0", profile,
      ...(profile.mode === "delegated" ? { scopes: contactScopes[0] } : {}) }, deps);
    assert.equal(listed.returned, 2);
    assert.deepEqual(listed.value, [
      { id: c1.id, displayName: c1.displayName, mail: c1.mail, companyName: c1.companyName },
      { id: c2.id, displayName: c2.displayName, mail: c2.mail, companyName: c2.companyName },
    ]);
    await assert.rejects(
      runApiGet({ path: "/contacts", apiVersion: "v1.0", profile, odata: "$select=id,phones",
        ...(profile.mode === "delegated" ? { scopes: contactScopes[0] } : {}) }, deps),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Unreviewed \$select field phones/.test(error.message);
      },
    );
  }
  assert.ok(requests.every(request => request.headers.Authorization === "Bearer opaque-fixture-secret"));
});

test("raw contact show projects the minimal set without navigation", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async () => credential }),
    application: new (await import("../dist/app-auth.js")).ApplicationAuth({ storage: "session-only", acquire: async () => credential }),
    transport: async () => ({ status: 200, headers: {}, body: JSON.stringify(c1) }),
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const shown = await runApiGet({ path: `/contacts/${c1.id}`, apiVersion: "v1.0", profile: delegatedProfile, scopes: contactScopes[0] }, deps);
  assert.deepEqual(shown, { id: c1.id, displayName: c1.displayName, mail: c1.mail, companyName: c1.companyName });
});

// Navigation fixtures follow the documented v1.0 shapes: directory-object
// rows of mixed user/contact type with the @odata.type discriminator. r3 is
// a limited-information row carrying only type and id.
const mgr = {
  "@odata.type": "#microsoft.graph.user",
  id: "11111111-1111-4111-8111-111111111111",
  displayName: "Megan Manager",
  mail: "MeganM@adatum.com",
};
const r1 = {
  "@odata.type": "#microsoft.graph.user",
  id: "22222222-2222-4222-8222-222222222222",
  displayName: "Rene Report",
  mail: "ReneR@adatum.com",
};
const r2 = {
  "@odata.type": "#microsoft.graph.orgContact",
  id: "9b2c9d4e-1f5a-4b6c-8d7e-8f9a0b1c2d3e",
  displayName: "Robin Report",
  mail: "RobinR@adatum.com",
};
const r3 = { "@odata.type": "#microsoft.graph.user", id: "limited-report-1" };
const navReports = [r1, r2, r3];

function navTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const match = /^\/v1\.0\/contacts\/([^/]+)\/(manager|directReports(?:\/.*)?)$/.exec(url.pathname);
    if (!match || match[1] !== c1.id) {
      return json(404, { error: { code: "Request_ResourceNotFound", message: "no such contact" } });
    }
    const [, , nav] = match;
    const castOf = segment => segment === "graph.user" ? "#microsoft.graph.user"
      : segment === "graph.orgContact" ? "#microsoft.graph.orgContact" : undefined;
    if (nav === "manager") return json(200, mgr);
    const tail = nav.split("/").slice(1);
    if (tail.length === 0) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [r2, r3] });
      return json(200, {
        value: [r1],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/contacts/${c1.id}/directReports?%24skiptoken=page2`,
      });
    }
    if (tail.length === 1 && tail[0] === "$count") {
      assert.equal(request.headers.ConsistencyLevel, "eventual");
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(navReports.length) };
    }
    if (tail.length === 1 && tail[0].startsWith("graph.")) {
      const type = castOf(tail[0]);
      if (!type) return json(404, { error: { code: "Unknown", message: "unexpected cast" } });
      return json(200, { value: navReports.filter(row => row["@odata.type"] === type) });
    }
    if (tail.length === 2 && tail[1] === "$count" && tail[0].startsWith("graph.")) {
      const type = castOf(tail[0]);
      if (!type) return json(404, { error: { code: "Unknown", message: "unexpected cast" } });
      assert.equal(request.headers.ConsistencyLevel, "eventual");
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(navReports.filter(row => row["@odata.type"] === type).length) };
    }
    if (tail.length === 1 || (tail.length === 2 && tail[1].startsWith("graph."))) {
      const [reportId, cast] = tail;
      const type = cast === undefined ? undefined : castOf(cast);
      if (cast !== undefined && !type) return json(404, { error: { code: "Unknown", message: "unexpected cast" } });
      const found = navReports.find(row => row.id === reportId && (type === undefined || row["@odata.type"] === type));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such report" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} shows a contact manager with type plus minimal rows`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode, navTransport());
      const result = await executeArgv(["entra", "contact", "show-manager", "--id", c1.id, "--profile", profile], overrides);
      assert.deepEqual(result.manager, { "@odata.type": "#microsoft.graph.user", id: mgr.id, displayName: mgr.displayName });
      assert.equal(requests.length, 1);
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com/v1.0/contacts/${c1.id}/manager?`));
      const wide = await executeArgv(["entra", "contact", "show-manager", "--id", c1.id, "--profile", profile,
        "--select", "id,displayName,mail"], overrides);
      assert.deepEqual(wide.manager, { "@odata.type": "#microsoft.graph.user", id: mgr.id, displayName: mgr.displayName, mail: mgr.mail });
      await assert.rejects(executeArgv(["entra", "contact", "show-manager", "--id", "contact-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists direct reports with type discriminators and limited-info preservation`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, navTransport());
      const result = await executeArgv(["entra", "contact", "list-direct-reports", "--id", c1.id, "--profile", profile], overrides);
      assert.deepEqual(result.directReports, [
        { "@odata.type": "#microsoft.graph.user", id: r1.id, displayName: r1.displayName },
        { "@odata.type": "#microsoft.graph.orgContact", id: r2.id, displayName: r2.displayName },
        { "@odata.type": "#microsoft.graph.user", id: r3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("show-direct-report --id <contact-id> --report-id <report-id>")));
      assert.ok(result.help.some(hint => hint.includes("only type and id")), "limited-information rows are flagged");
      assert.ok(!JSON.stringify(result).includes("ReneR@adatum.com"), "mail stays behind --select");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} direct-report casts select the typed route`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode, navTransport());
      const users = await executeArgv(["entra", "contact", "list-direct-reports", "--id", c1.id, "--profile", profile, "--as", "user"], overrides);
      assert.ok(requests[0].url.includes("/directReports/graph.user?"));
      assert.deepEqual(users.directReports.map(row => row.id), [r1.id, r3.id]);
      assert.ok(users.directReports.every(row => row["@odata.type"] === "#microsoft.graph.user"));
      const contacts = await executeArgv(["entra", "contact", "list-direct-reports", "--id", c1.id, "--profile", profile, "--as", "orgContact"], overrides);
      assert.ok(requests[1].url.includes("/directReports/graph.orgContact?"));
      assert.deepEqual(contacts.directReports.map(row => row.id), [r2.id]);
      const shown = await executeArgv(["entra", "contact", "show-direct-report", "--id", c1.id, "--report-id", r1.id, "--profile", profile, "--as", "user"], overrides);
      assert.ok(requests[2].url.includes(`/directReports/${r1.id}/graph.user?`));
      assert.equal(shown.directReport.displayName, r1.displayName);
      const counted = await executeArgv(["entra", "contact", "count-direct-reports", "--id", c1.id, "--profile", profile, "--as", "orgContact"], overrides);
      assert.ok(requests[3].url.endsWith("/directReports/graph.orgContact/$count"));
      assert.deepEqual(counted, { count: { returned: 1, complete: true } });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped direct-report list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode, navTransport());
      const first = await executeArgv(["entra", "contact", "list-direct-reports", "--id", c1.id, "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.directReports.map(row => row.id), [r1.id]);
      assert.equal(first.count.complete, false);
      const second = await executeArgv(["entra", "contact", "list-direct-reports", "--id", c1.id, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.directReports.map(row => row.id), [r2.id, r3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one direct report and validates report identity before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode, navTransport());
      const result = await executeArgv(["entra", "contact", "show-direct-report", "--id", c1.id, "--report-id", r2.id, "--profile", profile], overrides);
      assert.deepEqual(result.directReport, { "@odata.type": "#microsoft.graph.orgContact", id: r2.id, displayName: r2.displayName });
      const sent = requests.length;
      await assert.rejects(
        executeArgv(["entra", "contact", "show-direct-report", "--id", c1.id, "--profile", profile], overrides),
        { code: "VALIDATION_ERROR" },
      );
      assert.equal(requests.length, sent, "missing --report-id fails before HTTP");
      await assert.rejects(
        executeArgv(["entra", "contact", "show-direct-report", "--id", c1.id, "--report-id", "report-missing", "--profile", profile], overrides),
        error => error.code === "GRAPH_ERROR" && /not found or inaccessible \(404\)/.test(error.message),
      );
      assert.equal(requests.length, sent + 1);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts direct reports as one scalar with the documented header`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode, navTransport());
      const result = await executeArgv(["entra", "contact", "count-direct-reports", "--id", c1.id, "--profile", profile], overrides);
      assert.deepEqual(result, { count: { returned: 3, complete: true } });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, `https://graph.microsoft.com/v1.0/contacts/${c1.id}/directReports/$count`);
      assert.equal(requests[0].headers.ConsistencyLevel, "eventual");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} direct-report reads refuse misuse before credentials`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode, navTransport());
      await assert.rejects(
        executeArgv(["entra", "contact", "list-direct-reports", "--id", c1.id, "--profile", profile, "--as", "group"], overrides),
        error => error.code === "VALIDATION_ERROR" && /--as group is not a direct-report cast/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "contact", "show-manager", "--id", c1.id, "--profile", profile, "--as", "user"], overrides),
        { code: "VALIDATION_ERROR" },
      );
      await assert.rejects(
        executeArgv(["entra", "contact", "list-direct-reports", "--id", c1.id, "--profile", profile, "--filter", "id eq 'x'"], overrides),
        error => error.code === "VALIDATION_ERROR" && /unknown flag --filter/.test(error.message),
      );
      await assert.rejects(
        executeArgv(["entra", "contact", "count-direct-reports", "--id", c1.id, "--profile", profile, "--select", "id"], overrides),
        error => error.code === "VALIDATION_ERROR" && /unknown flag --select/.test(error.message),
      );
      assert.equal(calls.length, 0);
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied navigation reads surface scope, role and account guidance`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "contact", "show-manager", "--id", c1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("OrgContact.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Directory Readers")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "contact", "list-direct-reports", "--id", c1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable shows, lists and counts direct-report reads`, () => {
    const state = setupProfiles();
    const nav = { manager: mgr, reports: [r1, r2] };
    try {
      const shown = runContactCli(["entra", "contact", "show-manager", "--id", c1.id, "--profile", profile], state, mode, false, nav);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).manager, { "@odata.type": "#microsoft.graph.user", id: mgr.id, displayName: mgr.displayName });

      const listed = runContactCli(["entra", "contact", "list-direct-reports", "--id", c1.id, "--profile", profile, "--as", "user"], state, mode, false, nav);
      assert.equal(listed.status, 0, listed.stdout);
      assert.deepEqual(decode(listed.stdout).directReports.map(row => row.id), [r1.id]);

      const single = runContactCli(["entra", "contact", "show-direct-report", "--id", c1.id, "--report-id", r2.id, "--profile", profile], state, mode, false, nav);
      assert.equal(single.status, 0, single.stdout);
      assert.deepEqual(decode(single.stdout).directReport, { "@odata.type": "#microsoft.graph.orgContact", id: r2.id, displayName: r2.displayName });

      const counted = runContactCli(["entra", "contact", "count-direct-reports", "--id", c1.id, "--profile", profile], state, mode, false, nav);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).count, { returned: 2, complete: true });
      assert.ok(!counted.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["contact", "show-manager"], ["--id", c1.id]],
    [["contact", "list-direct-reports"], ["--id", c1.id]],
    [["contact", "show-direct-report"], ["--id", c1.id, "--report-id", r1.id]],
    [["contact", "count-direct-reports"], ["--id", c1.id]],
  ]) {
    for (const preview of [false, true]) {
      test(`${mode} ${command.join(" ")} refuses beta before credentials with preview=${preview}`, async () => {
        const state = setupProfiles();
        try {
          const path = join(state.dir, "config.json");
          const config = JSON.parse(readFileSync(path, "utf8"));
          config.profiles[profile].preview = preview;
          writeFileSync(path, JSON.stringify(config));
          const { calls, requests, overrides } = overridesFor(mode, navTransport());
          await assert.rejects(
            executeArgv(["entra", ...command, ...args, "--profile", profile, "--api-version", "beta"], overrides),
            { code: "VALIDATION_ERROR", message: "Contact reads support v1.0 only; beta needs its own review" },
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

test("raw navigation reads project reviewed fields and refuse unreviewed ones", async () => {
  const { runApiGet } = await import("../dist/api.js");
  const { DelegatedAuth } = await import("../dist/auth.js");
  const credential = { token: "opaque-fixture-secret", expiresAt: Date.now() + 3_600_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" };
  const row = { "@odata.type": "#microsoft.graph.user", id: r1.id, displayName: r1.displayName, mail: r1.mail, jobTitle: "unused" };
  const deps = {
    delegated: new DelegatedAuth({ storage: "session-only", login: async () => credential, silent: async () => credential }),
    application: new (await import("../dist/app-auth.js")).ApplicationAuth({ storage: "session-only", acquire: async () => credential }),
    transport: async () => ({ status: 200, headers: {}, body: JSON.stringify({ value: [row] }) }),
  };
  const delegatedProfile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "55555555-5555-4555-8555-555555555555" } };
  const listed = await runApiGet({ path: `/contacts/${c1.id}/directReports`, apiVersion: "v1.0", profile: delegatedProfile, scopes: contactScopes[0] }, deps);
  assert.deepEqual(listed.value, [{ "@odata.type": "#microsoft.graph.user", id: r1.id, displayName: r1.displayName }]);
  const cast = await runApiGet({ path: `/contacts/${c1.id}/directReports/graph.user`, apiVersion: "v1.0", profile: delegatedProfile, scopes: contactScopes[0] }, deps);
  assert.deepEqual(cast.value, [{ "@odata.type": "#microsoft.graph.user", id: r1.id, displayName: r1.displayName }]);
  await assert.rejects(
    runApiGet({ path: `/contacts/${c1.id}/directReports`, apiVersion: "v1.0", profile: delegatedProfile, odata: "$select=id,jobTitle", scopes: contactScopes[0] }, deps),
    error => {
      assert.equal(error.code, "VALIDATION_ERROR");
      return /Unreviewed \$select field jobTitle/.test(error.message);
    },
  );
  await assert.rejects(
    runApiGet({ path: `/contacts/${c1.id}/directReports/$count`, apiVersion: "v1.0", profile: delegatedProfile, scopes: contactScopes[0] }, deps),
    error => error.code === "VALIDATION_ERROR",
  );
});
