import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { decode } from "@toon-format/toon";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";

function run(scenario, args = [], packs = ["entra", "mail", "security"], mode = "delegated", options = {}) {
  const home = mkdtempSync(join(tmpdir(), "mg-home-"));
  try {
    const config = join(home, options.configOverride ? "alternate config's.json" : "config.json");
    writeFileSync(config, JSON.stringify({ version: 1, profiles: { [options.profileName ?? "soc"]: {
      mode, tenantId: tenant, clientId: client, enabledPacks: packs,
      ...(mode === "delegated" ? { token: { stored: true } } : { auth: "client-secret", clientSecret: { env: "SYNTHETIC_SECRET" } }),
    } } }));
    const prefix = ["--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
      "--import", new URL("./fixtures/read-pack-cli.mjs", import.meta.url).href,
      process.env.MG_AXI_PACK_ROOT ? join(process.env.MG_AXI_PACK_ROOT, "dist/bin/mg-axi.js")
        : fileURLToPath(new URL("../dist/bin/mg-axi.js", import.meta.url))];
    const spawnOptions = {
      encoding: "utf8", timeout: 15000,
      env: { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot, HOME: home, USERPROFILE: home,
        MG_AXI_CONFIG: join(home, "config.json"), SYNTHETIC_SECRET: "synthetic-secret",
        ...(options.environment ? { MG_AXI_MODE: mode, MG_AXI_TENANT_ID: tenant, MG_AXI_CLIENT_ID: client,
          MG_AXI_TOKEN: "synthetic-cache", MG_AXI_ENABLED_PACKS: JSON.stringify(packs) } : {}),
        MG_AXI_READ_FIXTURE: JSON.stringify({ mode, home: scenario }),
        ...(process.env.MG_AXI_PACK_ROOT ? { MG_AXI_PACK_ROOT: process.env.MG_AXI_PACK_ROOT } : {}) },
    };
    const result = spawnSync(process.execPath, [...prefix, ...args, ...(options.configOverride ? ["--config", config] : [])], spawnOptions);
    assert.equal(result.stderr, "", result.stderr);
    const output = decode(result.stdout);
    if (options.checkHints) options.checkHints(output, hint => {
      const followed = spawnSync("sh", ["-c", `exec "$@" ${hint.slice("mg-axi ".length)}`, "home-hint", process.execPath, ...prefix], {
        ...spawnOptions, env: { ...spawnOptions.env, MG_AXI_READ_FIXTURE: JSON.stringify({ mode }) },
      });
      assert.equal(followed.stderr, "", followed.stderr);
      assert.equal(followed.status, 0, followed.stdout);
      return decode(followed.stdout);
    }, config);
    return { ...result, output };
  } finally { rmSync(home, { recursive: true, force: true }); }
}

test("issue 134: configured executable home renders bounded live Graph summaries", () => {
  const result = run("live");
  assert.equal(result.status, 0, result.stdout);
  assert.equal(result.output.profile, "soc");
  assert.equal(result.output.tenant, tenant);
  assert.deepEqual(result.output.domains.slice(0, 2).map(row => [row.name, row.status, row.summary]), [
    ["entra", "live", "2 users observed; total 12; sample only"],
    ["mail", "live", "1 root folders observed; total 1; complete"],
  ]);
  assert.equal(result.output.domains[2].status, "unsupported");
  assert.ok(result.stdout.length < 4000);
  assert.doesNotMatch(result.stdout, /Adele|opaque-fixture|@odata|cursor:/);
});

test("home reports successful empty Graph collections explicitly", () => {
  const result = run("empty");
  assert.equal(result.status, 0, result.stdout);
  assert.deepEqual(result.output.domains.slice(0, 2).map(row => row.summary), [
    "0 users observed; total 0; complete", "0 root folders observed; total 0; complete",
  ]);
});

for (const scenario of ["live", "missing", "denied"]) {
  for (const environment of [false, true]) {
    test(`home hints preserve alternate config and ${environment ? "environment" : "configured"} identity for ${scenario}`, () => {
      const result = run(scenario, [], undefined, "delegated", { configOverride: true, environment,
        checkHints(output, follow, config) {
          const hints = output.help.filter(hint => hint.startsWith("mg-axi "));
          for (const hint of hints) {
            assert.ok(hint.includes(`--config '${config.replaceAll("'", "'\\''")}'`), hint);
            if (environment) assert.doesNotMatch(hint, /--profile/);
            else assert.match(hint, /--profile soc/);
          }
          const offline = follow(hints[0]);
          assert.equal(offline.profile, environment ? "env" : "soc");
          assert.equal(offline.tenant, tenant);
          if (scenario === "live") {
            const details = follow(hints.find(hint => hint.includes(" entra user list ")));
            assert.equal(details.users[0].displayName, "Adele Vance");
          }
        },
      });
      assert.equal(result.status, 0, result.stdout);
    });
  }
}

test("a configured profile named env remains explicitly selected in hints", () => {
  const result = run("live", ["--profile", "env"], undefined, "delegated", { profileName: "env" });
  assert.equal(result.status, 0, result.stdout);
  for (const hint of result.output.help.filter(hint => hint.startsWith("mg-axi "))) assert.match(hint, /--profile env/);
});

for (const scenario of ["missing", "expired", "denied", "unavailable", "throttled"]) {
  test(`home reports ${scenario} data without inventing empty results`, () => {
    const result = run(scenario);
    assert.equal(result.status, 0, result.stdout);
    assert.deepEqual(result.output.domains.slice(0, 2).map(row => row.status), ["unavailable", "unavailable"]);
    assert.doesNotMatch(result.stdout, /0 (users|root folders) observed/);
    assert.match(result.stdout, /--profile soc/);
    assert.equal(result.output.tenant, tenant);
    assert.match(result.stdout, scenario === "missing" || scenario === "expired" ? /AUTH_REQUIRED|AUTH_EXPIRED/ : /GRAPH_ERROR|incomplete/);
  });
}

test("home preserves a successful domain when another domain fails", () => {
  const result = run("partial");
  assert.equal(result.output.domains[0].summary, "2 users observed; total 12; sample only");
  assert.equal(result.output.domains[1].status, "unavailable");
  assert.match(result.output.domains[1].summary, /denied/);
});

test("a sampled home links to a named read for directory details", () => {
  const home = run("live");
  const hint = home.output.help.find(value => value.includes(" entra user list "));
  const args = hint.slice(hint.indexOf(" entra user list ") + 1).split(" ");
  const result = run(undefined, args);
  assert.equal(result.status, 0, result.stdout);
  assert.equal(result.output.users[0].displayName, "Adele Vance");
});

test("disabled packs do not authenticate or contact Graph", () => {
  const result = run("offline", [], []);
  assert.deepEqual(result.output.domains.map(row => row.status), ["disabled", "disabled", "disabled"]);
});

test("unsupported application mail and security packs do not authenticate", () => {
  const result = run("offline", [], ["mail", "security"], "application");
  assert.deepEqual(result.output.domains.map(row => row.status), ["disabled", "unsupported", "unsupported"]);
});

test("application home uses existing default audience for its live sample", () => {
  const result = run("live", [], ["entra"], "application");
  assert.equal(result.output.domains[0].summary, "2 users observed; total 12; sample only");
});

test("explicit offline home remains local without credentials", () => {
  const result = run("offline", ["home", "--offline"]);
  assert.equal(result.status, 0, result.stdout);
  assert.match(result.output.domains[0].summary, /no Graph request made/);
});

test("unknown home flags fail before authentication", () => {
  const result = run("offline", ["--unknown"]);
  assert.equal(result.status, 2, result.stdout);
  assert.equal(result.output.code, "VALIDATION_ERROR");
  assert.match(result.output.error, /unknown flag --unknown/);
});
