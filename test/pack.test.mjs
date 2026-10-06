import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { doctorTargets, runDoctor } from "../dist/doctor.js";
import { capabilityDocument, skillCommandTable, skillDescription, skillDocument, skillHomeHints } from "../dist/docs.js";
import { GraphSession, systemClock } from "../dist/graph-session.js";
import { Profiles } from "../dist/profiles.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const bin = resolve("dist/bin/mg-axi.js");
const userId = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const groupId = "gggggggg-gggg-4ggg-8ggg-gggggggggggg";

function dir() {
  return mkdtempSync(join(tmpdir(), "mg-axi-pack-"));
}

function runPlain(args, home) {
  return spawnSync(process.execPath, [bin, ...args], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      HOME: home, USERPROFILE: home, MG_AXI_CONFIG: join(home, "config.json"),
    },
  });
}

function runPack(args, home, mode, denied = false, throttled = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-pack-cli.mjs")).href, bin, ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      HOME: home, USERPROFILE: home, MG_AXI_CONFIG: join(home, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, denied, throttled }),
    },
  });
}

function createProfiles(home) {
  for (const args of [
    ["profile", "create", "--name", "soc", "--tenant", tenant, "--client", client, "--cloud", "commercial"],
    ["profile", "create", "--name", "batch", "--tenant", tenant, "--client", client, "--cloud", "commercial",
      "--mode", "application", "--federated"],
  ]) {
    const created = runPlain(args, home);
    assert.equal(created.status, 0, created.stdout);
  }
}

test("setup help is available offline", () => {
  const home = dir();
  try {
    const result = runPlain(["setup", "--help"], home);
    assert.equal(result.status, 0);
    assert.equal(result.stderr, "");
    assert.match(result.stdout, /without signing in or writing anything/);
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test("doctor help is available offline", () => {
  const home = dir();
  try {
    const result = runPlain(["doctor", "--help"], home);
    assert.equal(result.status, 0);
    assert.equal(result.stderr, "");
    assert.match(result.stdout, /never signs in interactively/);
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test("setup reports unconfigured state without writing configuration", () => {
  const home = dir();
  try {
    const config = join(home, "config.json");
    const result = runPlain(["setup"], home);
    assert.equal(result.status, 0, result.stdout);
    assert.equal(result.stderr, "");
    const output = decode(result.stdout);
    assert.equal(output.command, "mg-axi setup");
    assert.equal(output.config, config);
    assert.equal(output.state, "unconfigured");
    assert.equal(output.profiles, "0 profiles configured");
    assert.ok(output.capabilities.implemented.includes("mg-axi setup") || output.capabilities.implemented.includes("setup"));
    assert.ok(output.capabilities.implemented.includes("entra user list"));
    assert.equal(output.capabilities.reads, 192);
    assert.equal(output.capabilities.writes, 5);    assert.equal(output.capabilities.local, 6);
    assert.ok(!existsSync(config), "setup writes nothing");
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test("setup capability prose lists the shipped read families from the catalogue", () => {
  const home = dir();
  try {
    const output = decode(runPlain(["setup"], home).stdout);
    assert.equal(output.capabilities.api,
      "Entra user, registration, group, directory-role/PIM, device, administrative-unit, organization and branding, domain and domain DNS, certificate authentication, commercial subscriptions, on-premises synchronization, terms-of-use agreements and acceptances, directory objects, organizational contacts, group lifecycle policies, group setting templates, attribute sets and custom security attributes, partner contracts, delegated administration, multi-tenant organization, tenant information, sign-in/directory-audit, application/service-principal and consent grants, risk, Conditional Access, identity-provider, federation configurations, deleted directory items, data policy operations, access-review, entitlement catalogs and access packages, entitlement assignments and requests and lifecycle workflows reads plus reviewed raw api get, one gated group-membership write, one gated account enable/disable write, one gated session-revocation write and one gated risky-user dismissal write; every other operation is scheduled, blocked, deprecated or excluded");
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test("setup reports configured profiles without credential material", () => {
  const home = dir();
  try {
    createProfiles(home);
    const result = runPlain(["setup"], home);
    assert.equal(result.status, 0, result.stdout);
    assert.equal(result.stderr, "");
    const output = decode(result.stdout);
    assert.equal(output.state, "configured");
    assert.deepEqual(output.profiles, [
      { name: "soc", mode: "delegated" },
      { name: "batch", mode: "application" },
    ]);
    assert.ok(!result.stdout.includes("credentialRef"));
    assert.ok(!result.stdout.includes(tenant));
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test("setup suggests runnable checkout commands", () => {
  const home = dir();
  try {
    const output = decode(runPlain(["setup"], home).stdout);
    for (const command of [
      ...output.guidance.slice(1).map(command => `${command} --help`),
      `${output.help[0]} --help`, ...output.help.slice(1),
    ]) {
      const [executable, ...args] = command.split(" ");
      const result = spawnSync(executable, args, {
        encoding: "utf8", timeout: 30000,
        env: { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
          HOME: home, USERPROFILE: home, MG_AXI_CONFIG: join(home, "config.json") },
      });
      assert.equal(result.status, 0, `${command}: ${result.error ?? result.stdout}`);
      assert.match(result.stdout, /mg-axi/);
    }
    assert.ok(!existsSync(join(home, "config.json")));
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test("doctor rejects unknown flags before credentials", () => {
  const home = dir();
  try {
    const result = runPlain(["doctor", "--cursor", "anything"], home);
    assert.equal(result.status, 2);
    assert.equal(result.stderr, "");
    const output = decode(result.stdout);
    assert.match(output.error, /unknown flag --cursor/);
    assert.equal(output.code, "VALIDATION_ERROR");
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test("doctor without profiles fails operationally before any HTTP", () => {
  const home = dir();
  try {
    const result = runPlain(["doctor"], home);
    assert.equal(result.status, 1);
    assert.equal(result.stderr, "");
    const output = decode(result.stdout);
    assert.match(output.error, /No profiles are configured/);
    assert.equal(output.code, "AUTH_REQUIRED");
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test("packaged doctor checks a delegated profile with one bounded read", () => {
  const home = dir();
  try {
    createProfiles(home);
    const result = runPack(["doctor", "--profile", "soc"], home, "delegated");
    assert.equal(result.status, 0, result.stdout);
    assert.equal(result.stderr, "");
    const output = decode(result.stdout);
    assert.equal(output.count, "1 of 1 profiles ok");
    assert.equal(output.complete, true);
    assert.deepEqual(output.profiles, [{
      name: "soc", mode: "delegated",
      check: "entra user list --limit 1 (v1.0:GET:/users)",
      status: "ok", detail: "1 user row returned",
    }]);
    assert.ok(!result.stdout.includes("opaque-fixture-delegated-token"));
  } finally { rmSync(home, { recursive: true, force: true }); }
});

test("packaged doctor reports a denied profile as an operational failure", () => {
  const home = dir();
  try {
    createProfiles(home);
    const result = runPack(["doctor", "--profile", "soc"], home, "delegated", true);
    assert.equal(result.status, 1, result.stdout);
    assert.equal(result.stderr, "");
    const output = decode(result.stdout);
    assert.equal(output.count, "0 of 1 profiles ok");
    assert.equal(output.complete, false);
    assert.equal(output.profiles.length, 1);
    assert.equal(output.profiles[0].status, "failed");
    assert.equal(output.profiles[0].code, "GRAPH_ERROR");
    assert.match(output.profiles[0].error, /grant, role, licence or policy/);
    assert.ok(output.help.length);
    assert.ok(!result.stdout.includes("opaque-fixture-delegated-token"));
  } finally { rmSync(home, { recursive: true, force: true }); }
});

for (const mode of ["delegated", "application"]) {
  test(`packaged doctor reports throttling as failure for ${mode} profiles`, () => {
    const home = dir();
    try {
      createProfiles(home);
      const result = runPack(["doctor", "--profile", mode === "delegated" ? "soc" : "batch"], home, mode, false, true);
      assert.equal(result.status, 1, result.stdout);
      const output = decode(result.stdout);
      assert.equal(output.count, "0 of 1 profiles ok");
      assert.equal(output.complete, false);
      assert.equal(output.profiles[0].status, "failed");
      assert.equal(output.profiles[0].code, "GRAPH_ERROR");
      assert.match(output.profiles[0].error, /deadline exceeded; throttled \(429\)/);
    } finally { rmSync(home, { recursive: true, force: true }); }
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`packaged ${mode} critical journey reads users, groups, policies and sign-ins offline`, () => {
    const home = dir();
    try {
      createProfiles(home);
      const steps = [
        ["entra", "user", "list", "--profile", profile, "--limit", "10"],
        ["entra", "user", "show", "--profile", profile, "--id", userId],
        ["entra", "group", "list", "--profile", profile],
        ["entra", "group", "member", "list", "--profile", profile, "--group", groupId],
        ["entra", "conditional-access", "policy", "list", "--profile", profile],
        ["entra", "sign-in", "list", "--profile", profile, "--since", "2026-09-01T00:00:00Z"],
      ];
      const outputs = steps.map(args => {
        const result = runPack(args, home, mode);
        assert.equal(result.status, 0, `${args.join(" ")}:\n${result.stdout}`);
        assert.equal(result.stderr, "");
        assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
        return decode(result.stdout);
      });
      assert.equal(outputs[0].users.length, 1);
      assert.equal(outputs[1].user.id, userId);
      assert.equal(outputs[2].groups.length, 1);
      assert.equal(outputs[3].members.length, 1);
      assert.equal(outputs[4].policies.length, 1);
      assert.equal(outputs[5].signIns.length, 1);
    } finally { rmSync(home, { recursive: true, force: true }); }
  });
}

function delegatedJson(key) {
  return { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial",
    enabledPacks: ["entra"], preview: false, sensitiveAreas: [],
    credentialRef: { provider: "os-or-session", key }, allowDeviceCode: false };
}

function applicationJson(key) {
  return { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial",
    enabledPacks: ["entra"], preview: false, sensitiveAreas: [],
    credentialRef: { provider: "federated", key }, allowDeviceCode: false };
}

// Hand-written configuration without a default profile, so doctor checks
// every profile instead of selecting one.
function writeConfig(home) {
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(home, "config.json");
  writeFileSync(join(home, "config.json"), JSON.stringify({ version: 1,
    profiles: {
      soc: delegatedJson("33333333-3333-4333-8333-333333333333"),
      batch: applicationJson("44444444-4444-4333-8444-444444444444"),
    } }));
  return { store: new Profiles(), restore: () => { process.env.MG_AXI_CONFIG = previous; } };
}

test("doctorTargets selects explicit, default, sole and every profile", () => {
  const home = dir();
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(home, "config.json");
  try {
    const store = new Profiles();
    store.create("soc", tenant, client, "commercial", false);
    store.create("batch", tenant, client, "commercial", false, { federated: true });
    assert.deepEqual(doctorTargets(store, "batch"), ["batch"]);
    assert.deepEqual(doctorTargets(store), ["soc"]);
    const solo = dir();
    process.env.MG_AXI_CONFIG = join(solo, "config.json");
    try {
      const single = new Profiles();
      single.create("only", tenant, client, "commercial", false);
      assert.deepEqual(doctorTargets(single), ["only"]);
    } finally { rmSync(solo, { recursive: true, force: true }); }
    process.env.MG_AXI_CONFIG = join(home, "undefaulted.json");
    writeFileSync(join(home, "undefaulted.json"), JSON.stringify({ version: 1,
      profiles: {
        soc: delegatedJson("33333333-3333-4333-8333-333333333333"),
        batch: applicationJson("44444444-4444-4333-8444-444444444444"),
      } }));
    assert.deepEqual(doctorTargets(new Profiles()), ["soc", "batch"]);
  } finally {
    process.env.MG_AXI_CONFIG = previous;
    rmSync(home, { recursive: true, force: true });
  }
});

test("doctorTargets with no profiles selects nothing before any HTTP", () => {
  const home = dir();
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(home, "config.json");
  try {
    assert.deepEqual(doctorTargets(new Profiles()), []);
  } finally {
    process.env.MG_AXI_CONFIG = previous;
    rmSync(home, { recursive: true, force: true });
  }
});

function fixtureCredential(mode) {
  return {
    credential: async profile => ({
      token: `opaque-fixture-${mode}-token`,
      expiresAt: Date.now() + 3_600_000,
      tenantId: profile.tenantId,
      clientId: profile.clientId,
      ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
    }),
  };
}

function fixtureTransport(requests, deniedModes) {
  const send = async request => {
    requests.push(request);
    const forbidden = deniedModes.some(mode => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`);
    if (forbidden) return { status: 403, headers: {}, body: JSON.stringify({ error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }) };
    return { status: 200, headers: {}, body: JSON.stringify({ value: [{ id: userId, displayName: "Adele Vance", userPrincipalName: "AdeleV@contoso.com", mail: null }] }) };
  };
  return { requests, send };
}

for (const boundary of ["credential", "transport"]) {
  test(`doctor reports a ${boundary} deadline as failure`, async t => {
    const home = dir();
    const configured = writeConfig(home);
    let now = 0;
    const credential = fixtureCredential("delegated");
    const transport = fixtureTransport([], []).send;
    t.mock.method(systemClock, "now", () => now);
    try {
      const result = await runDoctor({ store: configured.store, names: ["soc"], session: new GraphSession({
        delegated: { credential: async profile => {
          if (boundary === "credential") now = 60_000;
          return credential.credential(profile);
        } },
        application: fixtureCredential("application"),
        transport: async request => {
          if (boundary === "transport") now = 60_000;
          return transport(request);
        },
      }) });
      assert.equal(result.failed, true);
      assert.equal(result.output.complete, false);
      assert.equal(result.output.count, "0 of 1 profiles ok");
      assert.equal(result.output.profiles[0].status, "failed");
      assert.match(result.output.profiles[0].error, /deadline exceeded/);
    } finally {
      configured.restore();
      rmSync(home, { recursive: true, force: true });
    }
  });
}

for (const [name, value, detail] of [
  ["empty result", [], "0 users matched; the absence of results is the answer"],
  ["buffered row limit", [{ id: userId }, { id: tenant }], "1 user row returned"],
  ["continuation row limit", [{ id: userId }], "1 user row returned"],
]) {
  test(`doctor accepts a successful ${name}`, async () => {
    const home = dir();
    const configured = writeConfig(home);
    const requests = [];
    try {
      const result = await runDoctor({ store: configured.store, names: ["soc"], session: new GraphSession({
        delegated: fixtureCredential("delegated"),
        application: fixtureCredential("application"),
        transport: async request => {
          requests.push(request);
          return { status: 200, headers: {}, body: JSON.stringify({ value,
            ...(name === "continuation row limit" ? { "@odata.nextLink": `${request.url}&$skiptoken=next` } : {}),
          }) };
        },
      }) });
      assert.equal(result.failed, false);
      assert.equal(result.output.complete, true);
      assert.equal(result.output.profiles[0].status, "ok");
      assert.equal(result.output.profiles[0].detail, detail);
      assert.equal(requests.length, 1);
    } finally {
      configured.restore();
      rmSync(home, { recursive: true, force: true });
    }
  });
}

async function doctorArgv(home, requests, deniedModes) {
  const previousConfig = process.env.MG_AXI_CONFIG;
  const previousExit = process.exitCode;
  process.env.MG_AXI_CONFIG = join(home, "config.json");
  try {
    return {
      output: await executeArgv(["doctor"], {
        transport: fixtureTransport(requests, deniedModes).send,
        delegated: fixtureCredential("delegated"),
        application: fixtureCredential("application"),
      }),
      exit: process.exitCode,
    };
  } finally {
    process.env.MG_AXI_CONFIG = previousConfig;
    process.exitCode = previousExit;
  }
}

test("doctor checks every profile with one bounded read each", async () => {
  const home = dir();
  const configured = writeConfig(home);
  const requests = [];
  try {
    const { output, exit } = await doctorArgv(home, requests, []);
    assert.notEqual(exit, 1);
    assert.equal(output.count, "2 of 2 profiles ok");
    assert.equal(output.complete, true);
    assert.equal(requests.length, 2);
    for (const request of requests) {
      const url = new URL(request.url);
      assert.equal(url.pathname, "/v1.0/users");
      assert.ok(url.searchParams.has("$select"));
      assert.ok(!url.searchParams.has("$skiptoken"));
    }
  } finally {
    configured.restore();
    rmSync(home, { recursive: true, force: true });
  }
});

test("doctor reports mixed success with a nonzero exit and rerun guidance", async () => {
  const home = dir();
  const configured = writeConfig(home);
  const requests = [];
  try {
    const { output, exit } = await doctorArgv(home, requests, ["application"]);
    assert.equal(exit, 1);
    assert.equal(output.count, "1 of 2 profiles ok");
    assert.equal(output.complete, false);
    assert.deepEqual(output.profiles[0].status, "ok");
    assert.deepEqual(output.profiles[1], {
      name: "batch", check: "entra user list --limit 1 (v1.0:GET:/users)",
      status: "failed", code: "GRAPH_ERROR", error: output.profiles[1].error,
    });
    assert.match(output.profiles[1].error, /grant, role, licence or policy/);
    assert.ok(output.help.some(hint => hint.includes("mg-axi doctor --profile batch")));
    assert.equal(requests.length, 2);
  } finally {
    configured.restore();
    rmSync(home, { recursive: true, force: true });
  }
});

test("doctor with no profiles throws before credentials", async () => {
  const home = dir();
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(home, "config.json");
  try {
    await assert.rejects(
      runDoctor({ store: new Profiles(), names: [], session: new GraphSession({
        delegated: fixtureCredential("delegated"),
        application: fixtureCredential("application"),
        transport: async () => { throw new Error("transport must stay untouched"); },
      }) }),
      /No profiles are configured/,
    );
  } finally {
    process.env.MG_AXI_CONFIG = previous;
    rmSync(home, { recursive: true, force: true });
  }
});

test("keeps the committed skill file generated from the catalogue", () => {
  const skill = readFileSync(new URL("../skills/mg-axi/SKILL.md", import.meta.url), "utf8").replace(/\r\n/g, "\n");
  assert.equal(skill, skillDocument());
});

test("keeps the committed skill command table generated from the catalogue", () => {
  const skill = readFileSync(new URL("../skills/mg-axi/SKILL.md", import.meta.url), "utf8").replace(/\r\n/g, "\n");
  const block = skill.split("<!-- command-registry:start -->\n")[1]?.split("\n<!-- command-registry:end -->")[0];
  assert.equal(block, skillCommandTable());
});

test("skill trigger description names the shipped writes and EXT reads", () => {
  const description = skillDescription();
  for (const write of ["account enable/disable", "session revocation", "group-membership add", "risky-user dismissal", "CA policy update"]) {
    assert.ok(description.includes(write), `description names write: ${write}`);
  }
  for (const area of ["domains and DNS records", "certificate authentication", "group lifecycle policies", "attribute sets and custom security attributes", "partner contracts", "delegated administration", "multi-tenant organization", "identity providers", "organization and branding", "access reviews"]) {
    assert.ok(description.includes(area), `description names area: ${area}`);
  }
});

test("skill next steps match the no-args home view hints", async () => {
  const home = mkdtempSync(join(tmpdir(), "mg-axi-skill-home-"));
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(home, "config.json");
  try {
    const hints = skillHomeHints();
    const skill = skillDocument();
    const output = await executeArgv([]);
    assert.ok(output && typeof output === "object" && Array.isArray(output.help));
    assert.deepEqual(hints.slice(0, 2), ["npx -y @knowttl/mg-axi profile list", "npx -y @knowttl/mg-axi login --help"]);
    assert.equal(new Set(hints).size, hints.length);
    assert.deepEqual(output.help, hints);
    const repeated = await executeArgv([]);
    assert.deepEqual(repeated.help, hints);
    assert.deepEqual(output.help, hints);
    assert.deepEqual(skillHomeHints(), hints);
    assert.equal(skillDocument(), skill);
    for (const hint of hints) assert.ok(skill.includes(hint), `skill prints home hint: ${hint}`);
  } finally {
    process.env.MG_AXI_CONFIG = previous;
    rmSync(home, { recursive: true, force: true });
  }
});

test("lists every executable leaf in the skill table exactly once", () => {
  const rows = skillCommandTable().split("\n").slice(2);
  assert.ok(rows.length > 40);
  assert.ok(rows.some(row => row.includes("`mg-axi setup` | native | local |")));
  assert.ok(rows.some(row => row.includes("`mg-axi doctor` | native | read |")));
  assert.ok(rows.some(row => row.includes("`mg-axi entra user list` | native | read |")));
  assert.ok(rows.some(row => row.includes("`mg-axi api get` | native | read |")));
});

test("keeps the committed coverage record generated from the inventory", () => {
  const coverage = readFileSync(new URL("../docs/coverage.md", import.meta.url), "utf8").replace(/\r\n/g, "\n");
  assert.equal(coverage, capabilityDocument());
});

test("write status follows the catalogue, not hard-coded prose", () => {
  const coverage = capabilityDocument();
  const shipped = coverage.split("Shipped initial writes: ")[1]?.split(".\n")[0] ?? "";
  const pending = coverage.split("Pending initial writes: ")[1]?.split(".\n")[0] ?? "";
  for (const slice of ["WRITE-01", "WRITE-02", "WRITE-03", "WRITE-04", "WRITE-05"]) assert.ok(shipped.includes(slice), `${slice} reads shipped`);
  assert.ok(pending === "none", `no initial write reads pending, got: ${pending}`);
  assert.doesNotMatch(coverage, /later writes \(WRITE-04 and beyond\)/);
});

test("calls the release the supported Entra read surface, not full coverage", () => {
  const coverage = capabilityDocument();
  assert.ok(coverage.includes("the supported Entra read surface, not full Entra coverage"));
  assert.doesNotMatch(coverage, /\d+\s*% (complete|coverage)/);
});
