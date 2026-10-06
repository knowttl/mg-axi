// Session-start hook: local-only summary, setup installer and argv strictness.
// Every journey runs the packaged binaries against disposable homes and
// configs; no test signs in or contacts Graph. The fetch poison-pill run
// proves the hook stays local even with configured profiles.
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { LEAVES } from "../dist/catalogue.js";
import { hookSummary } from "../dist/hook.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const appClient = "33333333-3333-4333-8333-333333333333";
const bin = resolve("dist/bin/mg-axi.js");
const hookBin = resolve("dist/bin/mg-axi-hook.js");
const poison = pathToFileURL(resolve("test/fixtures/hook-net-guard.mjs")).href;
const version = JSON.parse(readFileSync(resolve("package.json"), "utf8")).version;

function home() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-hook-"));
  return { dir, config: join(dir, "config.json") };
}

function baseEnv(dir, config) {
  return { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot, HOME: dir, USERPROFILE: dir, MG_AXI_CONFIG: config };
}

function createProfiles(env) {
  for (const args of [
    ["profile", "create", "--name", "soc", "--tenant", tenant, "--client", client, "--cloud", "commercial"],
    ["profile", "create", "--name", "batch", "--tenant", tenant, "--client", appClient, "--cloud", "commercial",
      "--mode", "application", "--federated"],
  ]) {
    const created = spawnSync(process.execPath, [bin, ...args], { encoding: "utf8", input: "", timeout: 30000, env });
    assert.equal(created.status, 0, created.stdout);
  }
}

function runHook(env, args = [], extraArgs = []) {
  return spawnSync(process.execPath, [...extraArgs, hookBin, ...args], { encoding: "utf8", input: "", timeout: 30000, env });
}

test("hook prints a short not-configured record with exit 0 and empty stderr when unconfigured", () => {
  const { dir, config } = home();
  try {
    const result = runHook(baseEnv(dir, config));
    assert.equal(result.status, 0);
    assert.equal(result.stderr, "");
    assert.match(result.stdout, /not configured/);
    assert.match(result.stdout, /mg-axi setup/);
    assert.ok(result.stdout.split("\n").filter(line => line.trim()).length <= 4);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("hook summarizes configured profiles with tenant, auth, writes, version and one hint", () => {
  const { dir, config } = home();
  try {
    const env = baseEnv(dir, config);
    createProfiles(env);
    const result = runHook(env);
    assert.equal(result.status, 0);
    assert.equal(result.stderr, "");
    for (const fragment of ["configured", version, "soc", "batch", tenant, "client-credentials", "disabled", "mg-axi doctor"]) {
      assert.match(result.stdout, new RegExp(fragment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")), `missing ${fragment}:\n${result.stdout}`);
    }
    assert.match(result.stdout, /cached|sign-in required|unknown/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("hook never prints credential references", () => {
  const { dir, config } = home();
  try {
    const env = baseEnv(dir, config);
    createProfiles(env);
    const stored = JSON.parse(readFileSync(config, "utf8"));
    const keys = Object.values(stored.profiles).map(profile => profile.credentialRef.key);
    assert.ok(keys.length > 0);
    const result = runHook(env);
    assert.equal(result.status, 0);
    for (const key of keys) assert.doesNotMatch(result.stdout, new RegExp(key));
    assert.doesNotMatch(result.stdout, /credentialRef/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("hook reports configuration invalid with exit 0 for a malformed config", () => {
  const { dir, config } = home();
  try {
    writeFileSync(config, "{not json");
    const result = runHook(baseEnv(dir, config));
    assert.equal(result.status, 0);
    assert.equal(result.stderr, "");
    assert.match(result.stdout, /configuration invalid/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("hook surfaces hand-enabled write operations from the profile", () => {
  const { dir, config } = home();
  try {
    const env = baseEnv(dir, config);
    createProfiles(env);
    const stored = JSON.parse(readFileSync(config, "utf8"));
    stored.profiles.soc.writes = { allowWrites: true, operations: ["ENTRA_USER_UPDATE"] };
    writeFileSync(config, JSON.stringify(stored));
    const result = runHook(env);
    assert.equal(result.status, 0);
    assert.match(result.stdout, /ENTRA_USER_UPDATE/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("hook maps injected auth presence without touching the credential store", async () => {
  const { dir, config } = home();
  try {
    const env = baseEnv(dir, config);
    createProfiles(env);
    const cached = await hookSummary({ version, configPath: config, authPresence: () => "cached" });
    assert.match(cached, /cached/);
    const failing = await hookSummary({ version, configPath: config, authPresence: () => { throw new Error("store offline"); } });
    assert.match(failing, /unknown/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("hook names only real catalogue leaves in its hints", async () => {
  for (const leaf of ["setup", "doctor"]) assert.ok(LEAVES.some(item => item.path === leaf), `catalogue lacks ${leaf}`);
  const { dir, config } = home();
  try {
    const env = baseEnv(dir, config);
    assert.match(await hookSummary({ version, configPath: config }), /mg-axi setup/);
    createProfiles(env);
    assert.match(await hookSummary({ version, configPath: config }), /mg-axi doctor/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("hook makes no network call with configured profiles", () => {
  const { dir, config } = home();
  try {
    const env = baseEnv(dir, config);
    createProfiles(env);
    const result = runHook(env, [], ["--import", poison]);
    assert.equal(result.status, 0, result.stderr);
    assert.equal(result.stderr, "");
    assert.match(result.stdout, /configured/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("hook rejects argv with a usage error and honors --help without side effects", () => {
  const { dir, config } = home();
  try {
    const env = baseEnv(dir, config);
    const rejected = runHook(env, ["extra"]);
    assert.equal(rejected.status, 2);
    assert.match(rejected.stdout, /VALIDATION_ERROR/);
    assert.match(rejected.stdout, /mg-axi-hook/);
    const help = runHook(env, ["--help"]);
    assert.equal(help.status, 0);
    assert.match(help.stdout, /mg-axi-hook/);
    assert.ok(!existsSync(join(dir, ".claude")) && !existsSync(join(dir, ".codex")) && !existsSync(join(dir, ".config")));
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("setup hooks shows help offline without installing anything", () => {
  const { dir, config } = home();
  try {
    const result = spawnSync(process.execPath, [bin, "setup", "hooks", "--help"],
      { encoding: "utf8", input: "", timeout: 30000, env: baseEnv(dir, config) });
    assert.equal(result.status, 0);
    assert.equal(result.stderr, "");
    assert.match(result.stdout, /mg-axi setup hooks/);
    assert.ok(!existsSync(join(dir, ".claude")) && !existsSync(join(dir, ".codex")) && !existsSync(join(dir, ".config")));
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("setup hooks installs SessionStart entries idempotently", () => {
  const { dir, config } = home();
  try {
    const env = baseEnv(dir, config);
    const paths = [
      join(dir, ".claude", "settings.json"),
      join(dir, ".codex", "hooks.json"),
      join(dir, ".config", "opencode", "plugins", "axi-mg-axi-hook.js"),
    ];
    const first = spawnSync(process.execPath, [bin, "setup", "hooks"],
      { encoding: "utf8", input: "", timeout: 30000, env });
    assert.equal(first.status, 0, first.stdout);
    assert.equal(first.stderr, "");
    assert.match(first.stdout, /installed/);
    for (const path of paths) assert.match(readFileSync(path, "utf8"), /mg-axi-hook/);
    const before = paths.map(path => readFileSync(path, "utf8"));
    const second = spawnSync(process.execPath, [bin, "setup", "hooks"],
      { encoding: "utf8", input: "", timeout: 30000, env });
    assert.equal(second.status, 0);
    assert.deepEqual(paths.map(path => readFileSync(path, "utf8")), before);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
