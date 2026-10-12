import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, existsSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { decode } from "@toon-format/toon";
import { toolDirectory } from "../dist/paths.js";
import { PrivateFiles } from "../dist/private-file.js";
import { readOnlyForced, resolveWriteLogPath } from "../dist/mutations.js";
import { Profiles, validateApplicationProfile } from "../dist/profiles.js";
import { Credentials } from "../dist/credentials.js";
import { authStatus } from "../dist/auth-status.js";
import { registerSecret, redactSecrets } from "../dist/redact.js";
import { resolveSecret } from "../dist/secrets.js";

function fixture() {
  const dir = mkdtempSync(join(tmpdir(), "mg-credential-files-"));
  const env = { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot, HOME: dir, USERPROFILE: dir,
    XDG_CONFIG_HOME: join(dir, "config"), XDG_STATE_HOME: join(dir, "state"), XDG_CACHE_HOME: join(dir, "cache"),
    APPDATA: join(dir, "appdata"), LOCALAPPDATA: join(dir, "localappdata") };
  const run = (args, extra = {}) => spawnSync(process.execPath, ["dist/bin/mg-axi.js", ...args], { encoding: "utf8", input: "", timeout: 10000, env: { ...env, ...extra } });
  return { dir, env, run, clean: () => rmSync(dir, { recursive: true, force: true }) };
}

for (const [platform, env, kind, expected] of [
  ["linux", { XDG_CONFIG_HOME: "/sandbox/config" }, "config", "/sandbox/config/mg-axi"],
  ["linux", { XDG_CONFIG_HOME: "relative" }, "config", "/home/test/.config/mg-axi"],
  ["darwin", {}, "state", "/home/test/.local/state/mg-axi"],
  ["win32", { XDG_CONFIG_HOME: "/ignored", APPDATA: "C:\\Roaming" }, "config", "C:\\Roaming\\mg-axi"],
  ["win32", { XDG_STATE_HOME: "/ignored", LOCALAPPDATA: "C:\\Local" }, "state", "C:\\Local\\mg-axi"],
]) test(`directory resolution uses ${platform} ${kind} conventions: ${expected}`, () => {
  assert.equal(toolDirectory(kind, env, platform, "/home/test"), expected);
});

for (const [value, expected] of [[undefined, false], ["0", false], ["FALSE", false], ["1", true], ["TrUe", true]]) test(`read-only boolean ${value} is ${expected}`, () => {
  assert.equal(readOnlyForced({ MG_AXI_READ_ONLY: value }), expected);
});

for (const value of ["", "yes", "2", " true "]) test(`invalid read-only boolean ${JSON.stringify(value)} fails closed`, () => {
  assert.throws(() => readOnlyForced({ MG_AXI_READ_ONLY: value }), { code: "CONFIG_INVALID" });
});

test("write journal defaults to the selected machine-local state directory", () => {
  const f = fixture();
  try {
    assert.equal(resolveWriteLogPath(f.env), join(process.platform === "win32" ? f.env.LOCALAPPDATA : f.env.XDG_STATE_HOME, "mg-axi", "writes.log"));
    assert.equal(resolveWriteLogPath({ ...f.env, MG_AXI_WRITE_LOG: join(f.dir, "override.log") }), join(f.dir, "override.log"));
  } finally { f.clean(); }
});

test("config flag overrides environment and config path reads no content", () => {
  const f = fixture();
  try {
    const result = f.run(["config", "path", "--config", join(f.dir, "flag.json")], { MG_AXI_CONFIG: join(f.dir, "env.json") });
    assert.equal(result.status, 0, result.stdout);
    assert.equal(decode(result.stdout).config, join(f.dir, "flag.json"));
    assert.equal(existsSync(join(f.dir, "flag.json")), false);
  } finally { f.clean(); }
});

test("legacy config is never parsed and reports both paths and the move command", () => {
  const f = fixture();
  try {
    mkdirSync(join(f.dir, ".mg-axi"));
    writeFileSync(join(f.dir, ".mg-axi/config.json"), "not JSON and not read");
    const result = f.run(["profile", "list"]);
    assert.notEqual(result.status, 0);
    const output = decode(result.stdout);
    assert.equal(output.code, "CONFIG_MOVED");
    assert.match(output.error, /\.mg-axi/);
    assert.match(output.error, /config[/\\]mg-axi/);
    assert.match(output.help.join(" "), /mv|Move-Item/);
  } finally { f.clean(); }
});

test("profile environment selects explicitly and the only-profile fallback works", () => {
  const f = fixture();
  try {
    const config = join(f.dir, "config.json");
    const create = name => f.run(["profile", "create", "--config", config, "--name", name, "--tenant", "11111111-1111-4111-8111-111111111111", "--client", "22222222-2222-4222-8222-222222222222", "--cloud", "commercial"]);
    assert.equal(create("first").status, 0);
    const only = JSON.parse(readFileSync(config, "utf8"));
    delete only.defaultProfile;
    writeFileSync(config, JSON.stringify(only));
    assert.equal(decode(f.run(["profile", "show", "--config", config]).stdout).name, "first");
    assert.equal(create("second").status, 0);
    assert.equal(decode(f.run(["profile", "show", "--config", config], { MG_AXI_PROFILE: "second" }).stdout).name, "second");
    assert.equal(decode(f.run(["profile", "show", "--config", config, "--profile", "first"], { MG_AXI_PROFILE: "second" }).stdout).name, "first");
    assert.equal(decode(f.run(["profile", "show", "--config", config], { MG_AXI_PROFILE: "missing" }).stdout).code, "PROFILE_NOT_FOUND");
    const multiple = JSON.parse(readFileSync(config, "utf8"));
    delete multiple.defaultProfile;
    writeFileSync(config, JSON.stringify(multiple));
    assert.equal(decode(f.run(["profile", "show", "--config", config]).stdout).code, "PROFILE_AMBIGUOUS");
  } finally { f.clean(); }
});

test("config read checks reject writable JSON before parsing it", { skip: process.platform === "win32" }, () => {
  const f = fixture();
  try {
    const config = join(f.dir, "config.json");
    writeFileSync(config, "unparsed fixture content");
    chmodSync(config, 0o666);
    const result = f.run(["profile", "list", "--config", config]);
    assert.equal(decode(result.stdout).code, "CONFIG_INSECURE");
    assert.doesNotMatch(result.stdout + result.stderr, /unparsed fixture content/);
  } finally { f.clean(); }
});

test("auth status and logout are offline and never create credentials for a new profile", () => {
  const f = fixture();
  try {
    const config = join(f.dir, "config.json");
    const profiles = new Profiles(config);
    profiles.create("soc", "11111111-1111-4111-8111-111111111111", "22222222-2222-4222-8222-222222222222", "commercial", false);
    const status = f.run(["auth", "status", "--config", config]);
    assert.equal(status.status, 0, status.stdout);
    const output = decode(status.stdout);
    assert.equal(output.profile, "soc");
    assert.equal(output.endpoint, "https://graph.microsoft.com");
    assert.equal(output.sources[0].available, false);
    assert.match(output.sources[0].source, /stored:.*credentials\.json$/);
    const logout = f.run(["auth", "logout", "--config", config]);
    assert.equal(logout.status, 0, logout.stdout);
    assert.equal(decode(logout.stdout).status, "signed out");
    assert.equal(existsSync(join(process.platform === "win32" ? f.env.LOCALAPPDATA : f.env.XDG_STATE_HOME, "mg-axi/credentials.json")), false);
  } finally { f.clean(); }
});

test("auth status verifies the Windows DACL and exposes only cache availability", () => {
  const f = fixture();
  try {
    const events = [];
    const files = new PrivateFiles("win32", (path, _directory, action) => { events.push({ path, action }); });
    const config = join(f.dir, "config.json");
    const profiles = new Profiles(config, files);
    profiles.create("soc", "11111111-1111-4111-8111-111111111111", "22222222-2222-4222-8222-222222222222", "commercial", false);
    const profile = profiles.resolve().profile;
    const credentials = new Credentials(join(f.dir, "credentials.json"), files);
    credentials.write(profile, JSON.stringify({ RefreshToken: { fixture: { secret: "fixture-private-refresh-token" } }, AccessToken: { fixture: { secret: "fixture-private-access-token" } } }));
    profiles.verifyPermissions();
    const status = authStatus("soc", profile, credentials);
    assert.equal(status.sources[0].available, true);
    assert.doesNotMatch(JSON.stringify(status), /fixture-private/);
    assert.doesNotMatch(readFileSync(credentials.path, "utf8"), /fixture-private-access-token/);
    assert.ok(events.some(event => event.path === config && event.action === "verify"));
    assert.ok(events.some(event => event.path === credentials.path && event.action === "verify"));
    const denied = new Credentials(credentials.path, new PrivateFiles("win32", () => { throw new Error("fixture bad DACL"); }));
    assert.throws(() => authStatus("soc", profile, denied), { code: "CONFIG_INSECURE" });
  } finally { f.clean(); }
});

test("complete env profiles precede defaults and explicit file profiles ignore env fields", () => {
  const f = fixture();
  try {
    const config = join(f.dir, "config.json");
    const profiles = new Profiles(config);
    profiles.create("soc", "11111111-1111-4111-8111-111111111111", "22222222-2222-4222-8222-222222222222", "commercial", false);
    const env = { MG_AXI_MODE: "application", MG_AXI_AUTH: "client-secret", MG_AXI_TENANT_ID: "33333333-3333-4333-8333-333333333333", MG_AXI_CLIENT_ID: "44444444-4444-4444-8444-444444444444", MG_AXI_CLIENT_SECRET: "fixture-env-secret-quote\"/" };
    const selected = f.run(["profile", "show", "--config", config], env);
    assert.equal(selected.status, 0, selected.stdout);
    assert.equal(decode(selected.stdout).name, "env");
    assert.doesNotMatch(selected.stdout + selected.stderr, /fixture-env-secret/);
    const explicit = f.run(["profile", "show", "--config", config, "--profile", "soc"], env);
    assert.equal(decode(explicit.stdout).name, "soc");
    const status = f.run(["auth", "status", "--config", join(f.dir, "absent.json")], env);
    assert.equal(status.status, 0, status.stdout);
    assert.equal(decode(status.stdout).sources[0].source, "env:MG_AXI_CLIENT_SECRET");
    assert.doesNotMatch(status.stdout + status.stderr, /fixture-env-secret/);
  } finally { f.clean(); }
});

test("partial env profiles fail before falling through to a file default", () => {
  const f = fixture();
  try {
    const config = join(f.dir, "config.json");
    const profiles = new Profiles(config);
    profiles.create("soc", "11111111-1111-4111-8111-111111111111", "22222222-2222-4222-8222-222222222222", "commercial", false);
    const result = f.run(["profile", "show", "--config", config], { MG_AXI_TENANT_ID: "11111111-1111-4111-8111-111111111111" });
    assert.equal(decode(result.stdout).code, "CONFIG_INVALID");
  } finally { f.clean(); }
});

test("secret-only environment falls through to the only file profile", () => {
  const f = fixture();
  try {
    const config = join(f.dir, "config.json");
    const created = f.run(["profile", "create", "--config", config, "--name", "soc", "--tenant", "11111111-1111-4111-8111-111111111111", "--client", "22222222-2222-4222-8222-222222222222", "--cloud", "commercial", "--mode", "application", "--federated"]);
    assert.equal(created.status, 0, created.stdout);
    const shown = f.run(["profile", "show", "--config", config], { MG_AXI_TOKEN: "fixture-headless-assertion" });
    assert.equal(shown.status, 0, shown.stdout);
    assert.equal(decode(shown.stdout).name, "soc");
    assert.doesNotMatch(shown.stdout + shown.stderr, /fixture-headless-assertion/);
  } finally { f.clean(); }
});

for (const patch of [{ token: "fixture-inline-secret" }, { token: { env: "1INVALID" } }, { token: { env: "GOOD", stored: true } }, { token: { file: "relative" } }, { preview: null }, { enabledPacks: null }, { credentialRef: { provider: "os-or-session" } }]) test(`config rejects invalid credential or policy fields on load: ${JSON.stringify(patch)}`, () => {
  const f = fixture();
  try {
    const path = join(f.dir, "config.json");
    writeFileSync(path, JSON.stringify({ version: 1, profiles: { fixture: { mode: "delegated", tenantId: "11111111-1111-4111-8111-111111111111", clientId: "22222222-2222-4222-8222-222222222222", ...patch } } }), { mode: 0o600 });
    assert.throws(() => new Profiles(path).list(), error => error.code === "CONFIG_INVALID" && !JSON.stringify(error).includes("fixture-inline-secret"));
  } finally { f.clean(); }
});

test("exact-value redaction removes raw, JSON-escaped and URL-encoded secrets", () => {
  const value = 'fixture-secret-"/\n%';
  registerSecret(value);
  assert.equal(redactSecrets([value, JSON.stringify(value).slice(1, -1), encodeURIComponent(value)].join("|")), "***redacted***|***redacted***|***redacted***");
});

test("auth login and its legacy alias reject disabled device code before network access", () => {
  const f = fixture();
  try {
    const config = join(f.dir, "config.json");
    const created = f.run(["profile", "create", "--config", config, "--name", "soc", "--tenant", "11111111-1111-4111-8111-111111111111", "--client", "22222222-2222-4222-8222-222222222222", "--cloud", "commercial"]);
    assert.equal(created.status, 0, created.stdout);
    for (const command of [["auth", "login"], ["login"]]) {
      const result = f.run([...command, "--config", config, "--method", "device-code", "--scopes", "https://graph.microsoft.com/User.Read.All"]);
      assert.notEqual(result.status, 0);
      assert.match(decode(result.stdout).error, /disabled/);
      assert.equal(result.stderr, "");
    }
  } finally { f.clean(); }
});

test("Unix reads reject exposed secrets, writable configs, links and wrong ownership", { skip: process.platform === "win32" }, () => {
  const f = fixture();
  try {
    const path = join(f.dir, "secret");
    writeFileSync(path, "\uFEFFfixture-secret \r\n", { mode: 0o600 });
    const files = new PrivateFiles();
    chmodSync(path, 0o400);
    assert.equal(resolveSecret({ file: path }, files), "fixture-secret");
    chmodSync(path, 0o644);
    assert.equal(files.read(path, false), "fixture-secret \r\n");
    assert.throws(() => files.read(path), { code: "CONFIG_INSECURE" });
    chmodSync(path, 0o666);
    assert.throws(() => files.read(path, false), { code: "CONFIG_INSECURE" });
    chmodSync(path, 0o600);
    assert.throws(() => new PrivateFiles("linux", undefined, process.getuid() + 1).read(path), { code: "CONFIG_INSECURE" });
    symlinkSync(path, join(f.dir, "link"));
    assert.throws(() => files.read(join(f.dir, "link")), { code: "CONFIG_INSECURE" });
  } finally { f.clean(); }
});

test("Unix symlink loops fail as insecure, never a raw ELOOP", { skip: process.platform === "win32" }, () => {
  const f = fixture();
  try {
    const loop = join(f.dir, "loop");
    symlinkSync(loop, loop);
    assert.throws(() => new PrivateFiles().read(join(loop, "entry")), error => error.code === "CONFIG_INSECURE" && error.message.includes(join(loop, "entry")));
  } finally { f.clean(); }
});

test("auth status resolves Windows file holders through the injected store files", () => {
  const f = fixture();
  try {
    const files = new PrivateFiles("win32", () => {});
    const profile = validateApplicationProfile({ mode: "application", tenantId: "11111111-1111-4111-8111-111111111111", clientId: "22222222-2222-4222-8222-222222222222", cloud: "commercial", auth: "federated", token: { file: "C:\\creds\\assertion" } }, files);
    const status = authStatus("win", profile, new Credentials(join(f.dir, "credentials.json"), files));
    assert.deepEqual(status.sources, [{ field: "token", source: "file:C:\\creds\\assertion", available: false }]);
  } finally { f.clean(); }
});

test("atomic writes protect existing directories and leave no temporary files", { skip: process.platform === "win32" }, () => {
  const f = fixture();
  try {
    chmodSync(f.dir, 0o755);
    const path = join(f.dir, "credentials.json");
    new PrivateFiles().write(path, "fixture");
    assert.equal(statSync(f.dir).mode & 0o777, 0o700);
    assert.equal(statSync(path).mode & 0o777, 0o600);
    assert.deepEqual(readdirSync(f.dir), ["credentials.json"]);
  } finally { f.clean(); }
});

test("Windows writes restrict directory and temp DACLs before content is written", () => {
  const f = fixture();
  try {
    const events = [];
    const files = new PrivateFiles("win32", (path, directory, action) => {
      events.push({ directory, action });
      if (!directory && action === "restrict") assert.equal(readFileSync(path, "utf8"), "");
    });
    const path = join(f.dir, "credentials.json");
    files.write(path, "fixture");
    assert.deepEqual(events, [{ directory: true, action: "restrict" }, { directory: false, action: "restrict" }]);
    chmodSync(path, 0o644);
    assert.equal(files.read(path), "fixture");
    files.check(path, true, true);
    assert.deepEqual(events.slice(2), [{ directory: false, action: "inspect" }, { directory: false, action: "verify" }]);
    assert.equal(files.absolute("C:\\secrets\\token"), true);
    for (const invalid of ["relative", "\\\\server\\secret", "\\\\?\\C:\\secret", "C:relative"]) assert.equal(files.absolute(invalid), false);
  } finally { f.clean(); }
});

test("Windows ACL or reparse-check failure fails closed and cleans temporary files", () => {
  const f = fixture();
  try {
    const path = join(f.dir, "credentials.json");
    const files = new PrivateFiles("win32", (_path, directory) => { if (!directory) throw new Error("fixture ACL failure"); });
    assert.throws(() => files.write(path, "never written"), { code: "CONFIG_INSECURE" });
    assert.deepEqual(readdirSync(f.dir), []);
    writeFileSync(path, "fixture");
    assert.throws(() => files.read(path), { code: "CONFIG_INSECURE" });
  } finally { f.clean(); }
});

for (const code of ["EPERM", "EBUSY", "EACCES"]) test(`Windows atomic rename retries ${code} with bounded backoff`, () => {
  const result = spawnSync(process.execPath, ["--disable-warning=ExperimentalWarning", "--experimental-test-module-mocks", "--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import * as fs from "node:fs";
    import { tmpdir } from "node:os";
    import { join } from "node:path";
    import { mock } from "node:test";
    const { default: unused, ...named } = fs;
    let failures = 1;
    mock.module("node:fs", { namedExports: { ...named, renameSync: (...args) => {
      if (failures-- > 0) throw Object.assign(new Error("fixture busy"), { code: ${JSON.stringify(code)} });
      return fs.renameSync(...args);
    } } });
    const { PrivateFiles } = await import("./dist/private-file.js");
    const dir = fs.mkdtempSync(join(tmpdir(), "mg-windows-rename-"));
    mock.timers.enable({ apis: ["Date"], now: 1000 });
    const waits = [];
    mock.method(Atomics, "wait", (_array, _index, _value, duration) => { waits.push(duration); mock.timers.tick(duration); return "timed-out"; });
    try {
      const files = new PrivateFiles("win32", () => {});
      const path = join(dir, "credentials.json");
      files.write(path, "first");
      assert.equal(fs.readFileSync(path, "utf8"), "first");
      assert.deepEqual(waits, [10]);
      failures = Infinity;
      assert.throws(() => files.write(path, "replacement"), { code: ${JSON.stringify(code)} });
      assert.equal(fs.readFileSync(path, "utf8"), "first");
      assert.deepEqual(fs.readdirSync(dir), ["credentials.json"]);
      assert.equal(Date.now(), 2010);
    } finally { fs.rmSync(dir, { recursive: true, force: true }); }
  `], { encoding: "utf8", timeout: 10000 });
  assert.equal(result.status, 0, result.stderr);
});
