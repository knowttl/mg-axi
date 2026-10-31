import { record, records, text, strings, type OutputRecord } from "./fixtures/typed.ts";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { decode } from "@toon-format/toon";
import { invocation } from "../dist/invocation.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";

function run(scenario: string | undefined, args: readonly string[] = [], packs: readonly string[] = ["entra", "mail", "security"], mode = "delegated", options: {
  configOverride?: boolean;
  profileName?: string;
  environment?: boolean;
  consent?: string[];
  checkHints?: (output: OutputRecord, follow: (hint: string) => OutputRecord, config: string) => void;
} = {}) {
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
      encoding: "utf8" as const, timeout: 15000,
      env: { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot, HOME: home, USERPROFILE: home,
        MG_AXI_CONFIG: join(home, "config.json"), SYNTHETIC_SECRET: "synthetic-secret",
        ...(options.environment ? { MG_AXI_MODE: mode, MG_AXI_TENANT_ID: tenant, MG_AXI_CLIENT_ID: client,
          MG_AXI_TOKEN: "synthetic-cache", MG_AXI_ENABLED_PACKS: JSON.stringify(packs) } : {}),
        MG_AXI_READ_FIXTURE: JSON.stringify({ mode, home: scenario, consent: options.consent }),
        ...(process.env.MG_AXI_PACK_ROOT ? { MG_AXI_PACK_ROOT: process.env.MG_AXI_PACK_ROOT } : {}) },
    };
    const result = spawnSync(process.execPath, [...prefix, ...args, ...(options.configOverride ? ["--config", config] : [])], spawnOptions);
    assert.equal(result.stderr, "", result.stderr);
    const output = record(decode(result.stdout));
    if (options.checkHints) options.checkHints(output, hint => {
      const { command } = invocation(prefix.at(-1), spawnOptions.env);
      assert.ok(hint.startsWith(`${command} `), hint);
      const followed = spawnSync("sh", ["-c", `exec "$@" ${hint.slice(command.length + 1)}`, "home-hint", process.execPath, ...prefix], {
        ...spawnOptions, env: { ...spawnOptions.env, MG_AXI_READ_FIXTURE: JSON.stringify({ mode, consent: options.consent }) },
      });
      assert.equal(followed.stderr, "", followed.stderr);
      assert.equal(followed.status, 0, followed.stdout);
      return record(decode(followed.stdout));
    }, config);
    return { ...result, output };
  } finally { rmSync(home, { recursive: true, force: true }); }
}

test("issue 134: configured executable home renders bounded live Graph summaries", () => {
  const result = run("live");
  assert.equal(result.status, 0, result.stdout);
  assert.equal(result.output.profile, "soc");
  assert.equal(result.output.tenant, tenant);
  assert.deepEqual(records(result.output.domains).slice(0, 2).map(row => [row.name, row.status, row.summary]), [
    ["entra", "live", "2 users observed; total 12; sample only"],
    ["mail", "live", "1 root folders observed; total 1; complete"],
  ]);
  assert.equal(records(result.output.domains)[2]!.status, "unsupported");
  assert.ok(result.stdout.length < 4000);
  assert.doesNotMatch(result.stdout, /Adele|opaque-fixture|@odata|cursor:/);
});

test("home reports successful empty Graph collections explicitly", () => {
  const result = run("empty");
  assert.equal(result.status, 0, result.stdout);
  assert.deepEqual(records(result.output.domains).slice(0, 2).map(row => row.summary), [
    "0 users observed; total 0; complete", "0 root folders observed; total 0; complete",
  ]);
});

test("home uses existing User.Read.All-only consent and preserves it in the directory hint", () => {
  const result = run("live", [], ["entra"], "delegated", { consent: ["https://graph.microsoft.com/User.Read.All"] });
  assert.equal(result.status, 0, result.stdout);
  assert.equal(records(result.output.domains)[0]!.summary, "2 users observed; total 12; sample only");
  const hint = strings(result.output.help).find(value => value.includes(" entra user list "));
  assert.match(text(hint), / entra user list --profile soc --limit 5 --scopes https:\/\/graph\.microsoft\.com\/User\.Read\.All$/);
});

test("home reports mail unavailable when only existing directory consent is present", () => {
  const result = run("live", [], ["entra", "mail"], "delegated", { consent: ["https://graph.microsoft.com/User.Read.All"] });
  assert.equal(records(result.output.domains)[0]!.status, "live");
  assert.equal(records(result.output.domains)[1]!.status, "unavailable");
  assert.match(text(records(result.output.domains)[1]!.summary), /AUTH_REQUIRED.*Mail.Read access/);
});

test("home and its folder follow-up use existing Mail.Read-only consent", () => {
  const result = run("live", [], ["mail"], "delegated", { consent: ["https://graph.microsoft.com/Mail.Read"],
    checkHints(output, follow) {
      const folders = follow(strings(output.help).find(hint => hint.includes(" mail folder list "))!);
      assert.equal(records(folders.folders)[0]!.displayName, "Inbox");
    },
  });
  assert.equal(result.status, 0, result.stdout);
  assert.equal(records(result.output.domains)[1]!.summary, "1 root folders observed; total 1; complete");
});

for (const scenario of ["live", "missing", "denied"]) {
  for (const environment of [false, true]) {
    test(`home hints preserve alternate config and ${environment ? "environment" : "configured"} identity for ${scenario}`, () => {
      const result = run(scenario, [], undefined, "delegated", { configOverride: true, environment,
        checkHints(output, follow, config) {
          const hints = strings(output.help).filter(hint => / (?:home|entra|mail|security|auth) /.test(hint));
          for (const hint of hints) {
            assert.ok(hint.includes(`--config '${config.replaceAll("'", "'\\''")}'`), hint);
            if (environment) assert.doesNotMatch(hint, /--profile/);
            else assert.match(hint, /--profile soc/);
          }
          const offline = follow(hints.find(hint => hint.includes(" home --offline "))!);
          assert.equal(offline.profile, environment ? "env" : "soc");
          assert.equal(offline.tenant, tenant);
          if (scenario === "live") {
            const details = follow(hints.find(hint => hint.includes(" entra user list "))!);
            assert.equal(records(details.users)[0]!.displayName, "Adele Vance");
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
  const hints = strings(result.output.help).filter(hint => / (?:home|entra|mail|security|auth) /.test(hint));
  assert.ok(hints.length > 0);
  for (const hint of hints) assert.match(hint, /--profile env/);
});

for (const scenario of ["missing", "expired", "denied", "unavailable", "throttled"]) {
  test(`home reports ${scenario} data without inventing empty results`, () => {
    const result = run(scenario);
    assert.equal(result.status, 0, result.stdout);
    assert.deepEqual(records(result.output.domains).slice(0, 2).map(row => row.status), ["unavailable", "unavailable"]);
    assert.doesNotMatch(result.stdout, /0 (users|root folders) observed/);
    assert.match(result.stdout, /--profile soc/);
    assert.equal(result.output.tenant, tenant);
    assert.match(result.stdout, scenario === "missing" || scenario === "expired" ? /AUTH_REQUIRED|AUTH_EXPIRED/ : /GRAPH_ERROR|incomplete/);
  });
}

test("home preserves a successful domain when another domain fails", () => {
  const result = run("partial");
  assert.equal(records(result.output.domains)[0]!.summary, "2 users observed; total 12; sample only");
  assert.equal(records(result.output.domains)[1]!.status, "unavailable");
  assert.match(text(records(result.output.domains)[1]!.summary), /denied/);
});

test("a sampled home links to a named read for directory details", () => {
  const home = run("live");
  const hint = strings(home.output.help).find(value => value.includes(" entra user list "));
  const args = hint!.slice(hint!.indexOf(" entra user list ") + 1).split(" ");
  const result = run(undefined, args);
  assert.equal(result.status, 0, result.stdout);
  assert.equal(records(result.output.users)[0]!.displayName, "Adele Vance");
});

test("disabled packs do not authenticate or contact Graph", () => {
  const result = run("offline", [], []);
  assert.deepEqual(records(result.output.domains).map(row => row.status), ["disabled", "disabled", "disabled"]);
});

test("unsupported application mail and security packs do not authenticate", () => {
  const result = run("offline", [], ["mail", "security"], "application");
  assert.deepEqual(records(result.output.domains).map(row => row.status), ["disabled", "unsupported", "unsupported"]);
});

test("application home uses existing default audience for its live sample", () => {
  const result = run("live", [], ["entra"], "application");
  assert.equal(records(result.output.domains)[0]!.summary, "2 users observed; total 12; sample only");
});

test("explicit offline home remains local without credentials", () => {
  const result = run("offline", ["home", "--offline"]);
  assert.equal(result.status, 0, result.stdout);
  assert.match(text(records(result.output.domains)[0]!.summary), /no Graph request made/);
});

test("unknown home flags fail before authentication", () => {
  const result = run("offline", ["--unknown"]);
  assert.equal(result.status, 2, result.stdout);
  assert.equal(result.output.code, "VALIDATION_ERROR");
  assert.match(text(result.output.error), /unknown flag --unknown/);
});
