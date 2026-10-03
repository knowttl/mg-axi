import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { spawnSync } from "node:child_process";
import { decode } from "@toon-format/toon";
import { Profiles } from "../dist/profiles.js";
import { DelegatedAuth } from "../dist/auth.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const scopes = ["https://graph.microsoft.com/User.Read"];
const profile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "os-or-session", key: "33333333-3333-4333-8333-333333333333" } };
function fixture() {
  let time = 1000;
  const calls = [];
  const credential = () => ({ token: "opaque-fixture-secret", expiresAt: time + 120_000, tenantId: tenant, clientId: client, accountId: "synthetic-account" });
  const provider = {
    storage: "session-only",
    login: async (...args) => { calls.push(["login", ...args]); return credential(); },
    silent: async (...args) => { calls.push(["silent", ...args]); return credential(); },
  };
  return { provider, calls, auth: new DelegatedAuth(provider, () => time), advance: amount => { time += amount; } };
}

test("browser is the explicit login default and credentials stay out of its result", async () => {
  const f = fixture();
  const result = await f.auth.login(profile, undefined, scopes);
  assert.equal(f.calls[0][2], "browser");
  assert.deepEqual(result, { status: "authenticated", mode: "delegated", storage: "session-only" });
});

test("device code requires profile opt-in before provider access", async () => {
  const f = fixture();
  await assert.rejects(f.auth.login(profile, "device-code", scopes), { code: "VALIDATION_ERROR" });
  assert.equal(f.calls.length, 0);
});

test("an opted-in profile still defaults to browser", async () => {
  const f = fixture();
  await f.auth.login({ ...profile, allowDeviceCode: true }, undefined, scopes);
  assert.equal(f.calls[0][2], "browser");
});

test("explicit device selection reaches only the device provider", async () => {
  const f = fixture();
  await f.auth.login({ ...profile, allowDeviceCode: true }, "device-code", scopes);
  assert.equal(f.calls[0][2], "device-code");
});

test("browser failure never falls back or exposes provider diagnostics", async () => {
  const f = fixture();
  f.provider.login = async () => { throw new Error("opaque-fixture-secret"); };
  await assert.rejects(f.auth.login(profile, "browser", scopes), error => error.code === "LOGIN_FAILED" && !JSON.stringify(error).includes("opaque-fixture-secret"));
  assert.equal(f.calls.length, 0);
});

test("ordinary credential acquisition uses silent authentication only", async () => {
  const f = fixture();
  await f.auth.credential(profile, scopes);
  assert.equal(f.calls[0][0], "silent");
});

test("cache refreshes at the expiry margin with fake time", async () => {
  const f = fixture();
  await f.auth.login(profile, "browser", scopes);
  await f.auth.credential(profile, scopes);
  assert.equal(f.calls.length, 1);
  f.advance(60_000);
  await f.auth.credential(profile, scopes);
  assert.equal(f.calls[1][0], "silent");
});

test("silent refresh failure never returns a stale credential or opens login", async () => {
  const f = fixture();
  await f.auth.login(profile, "browser", scopes);
  f.advance(120_000);
  f.provider.silent = async () => { throw new Error("opaque-fixture-secret"); };
  await assert.rejects(f.auth.credential(profile, scopes), { code: "AUTH_REQUIRED" });
  await assert.rejects(f.auth.credential(profile, scopes), { code: "AUTH_REQUIRED" });
  assert.equal(f.calls.length, 1);
});

for (const [name, patch] of [
  ["wrong tenant", { tenantId: client }], ["wrong client", { clientId: tenant }],
  ["expired", { expiresAt: 1000 }], ["unknown expiry", { expiresAt: NaN }], ["missing account", { accountId: "" }],
]) test(`${name} credential is rejected without decoding the token`, async () => {
  const f = fixture();
  f.provider.silent = async () => ({ token: "opaque-fixture-secret", expiresAt: 121_000, tenantId: tenant, clientId: client, accountId: "synthetic-account", ...patch });
  await assert.rejects(f.auth.credential(profile, scopes), { code: "AUTH_REQUIRED" });
});

test("scope and identity changes cannot reuse a cached credential", async () => {
  const f = fixture();
  await f.auth.credential(profile, scopes);
  await f.auth.credential(profile, ["https://graph.microsoft.com/Group.Read.All"]);
  await f.auth.credential({ ...profile, credentialRef: { ...profile.credentialRef, key: tenant } }, scopes);
  assert.equal(f.calls.length, 3);
});

test("explicit account login invalidates cached credentials for all previous scopes", async () => {
  const f = fixture();
  await f.auth.credential(profile, scopes);
  await f.auth.login(profile, "browser", ["https://graph.microsoft.com/Group.Read.All"]);
  await f.auth.credential(profile, scopes);
  assert.equal(f.calls[2][0], "silent");
});

for (const method of ["browser", "device-code"]) test(`${method} accepts a hyphenated delegated Graph permission`, async () => {
  const f = fixture();
  const selected = { ...profile, allowDeviceCode: true };
  const requested = ["https://graph.microsoft.com/User-LifeCycleInfo.Read.All"];
  const result = await f.auth.login(selected, method, requested);
  assert.equal(result.status, "authenticated");
  assert.deepEqual(f.calls, [["login", selected, method, requested]]);
});

test("silent acquisition accepts a hyphenated delegated Graph permission", async () => {
  const f = fixture();
  const requested = ["https://graph.microsoft.com/User-LifeCycleInfo.Read.All"];
  const result = await f.auth.credential(profile, requested);
  assert.equal(result.accountId, "synthetic-account");
  assert.deepEqual(f.calls, [["silent", profile, requested]]);
});

for (const requested of [[], ["https://graph.microsoft.com/.default"], ["https://example.invalid/User.Read"], ["https://example.invalid/User-LifeCycleInfo.Read.All"], ["https://graph.microsoft.com/User-LifeCycleInfo.Read.All", "https://graph.microsoft.com/.default"]]) test(`unsupported delegated scopes ${JSON.stringify(requested)} fail before credentials`, async () => {
  const f = fixture();
  await assert.rejects(f.auth.login(profile, "browser", requested), { code: "VALIDATION_ERROR" });
  await assert.rejects(f.auth.credential(profile, requested), { code: "VALIDATION_ERROR" });
  assert.equal(f.calls.length, 0);
});

for (const patch of [{ mode: "application" }, { cloud: "government" }, { tenantId: "common" }, { clientId: "" }, { secret: "fixture" }, { credentialRef: { provider: "file", key: "fixture" } }]) test(`invalid profile ${JSON.stringify(patch)} fails before credentials`, async () => {
  const f = fixture();
  await assert.rejects(f.auth.credential({ ...profile, ...patch }, scopes), { code: "VALIDATION_ERROR" });
  assert.equal(f.calls.length, 0);
});

test("profiles persist only validated identity, policy and separate references", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-profile-"));
  try {
    const profiles = new Profiles(join(dir, "config.json"));
    profiles.create("soc", tenant, client, "commercial", false);
    profiles.create("device", tenant, client, "commercial", true);
    assert.equal(profiles.resolve().name, "soc");
    assert.notEqual(profiles.resolve("soc").profile.credentialRef.key, profiles.resolve("device").profile.credentialRef.key);
    assert.throws(() => profiles.create("soc", client, tenant, "commercial", false), { code: "VALIDATION_ERROR" });
    assert.equal(profiles.resolve("soc").profile.tenantId, tenant);
    assert.doesNotMatch(readFileSync(profiles.path, "utf8"), /opaque-fixture-secret/);
    if (process.platform !== "win32") assert.equal(statSync(profiles.path).mode & 0o777, 0o600);
    writeFileSync(profiles.path, JSON.stringify({ version: 2, profiles: {} }));
    assert.throws(() => profiles.list(), /explicitly migrate/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("packaged profile journey and rejected device login work with closed stdin", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-profile-cli-"));
  const run = args => spawnSync(process.execPath, ["dist/bin/mg-axi.js", ...args], { encoding: "utf8", input: "", timeout: 10_000, env: { PATH: process.env.PATH, MG_AXI_CONFIG: join(dir, "config.json") } });
  try {
    assert.equal(run(["profile", "create", "--name", "soc", "--tenant", tenant, "--client", client, "--cloud", "commercial"]).status, 0);
    assert.equal(decode(run(["profile", "show"]).stdout).profile.clientId, client);
    assert.equal(decode(run([]).stdout).tenant, tenant);
    assert.equal(decode(run(["home", "--profile", "soc"]).stdout).profile, "soc");
    const rejected = run(["login", "--method", "device-code", "--scopes", scopes[0]]);
    assert.equal(rejected.status, 2);
    assert.equal(rejected.stderr, "");
    assert.match(decode(rejected.stdout).error, /disabled/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
