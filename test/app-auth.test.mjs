import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { spawnSync } from "node:child_process";
import { decode } from "@toon-format/toon";
import { Profiles } from "../dist/profiles.js";
import { ApplicationAuth } from "../dist/app-auth.js";
import { DelegatedAuth } from "../dist/auth.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const audience = "https://graph.microsoft.com/.default";
const thumbprint = "ab".repeat(20);
const cert = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "certificate", key, thumbprint } };
const federated = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "federated", key } };
function fixture() {
  let time = 1000;
  const calls = [];
  const credential = () => ({ token: "opaque-fixture-secret", expiresAt: time + 120_000, tenantId: tenant, clientId: client });
  const provider = {
    storage: "session-only",
    acquire: async (...args) => { calls.push(["acquire", ...args]); return credential(); },
  };
  return { provider, calls, auth: new ApplicationAuth(provider, () => time), advance: amount => { time += amount; } };
}

test("application credential requests the configured Graph .default audience", async () => {
  const f = fixture();
  const credential = await f.auth.credential(federated);
  assert.equal(credential.token, "opaque-fixture-secret");
  assert.deepEqual(f.calls, [["acquire", federated, [audience]]]);
});

test("application auth exposes no interactive login surface", () => {
  const f = fixture();
  assert.equal(typeof f.auth.login, "undefined");
});

for (const requested of [[], ["https://graph.microsoft.com/User.Read"], ["https://graph.microsoft.com/.default", "https://graph.microsoft.com/User.Read"], ["https://example.invalid/.default"]]) test(`non-audience scopes ${JSON.stringify(requested)} fail before provider access`, async () => {
  const f = fixture();
  await assert.rejects(f.auth.credential(cert, requested), { code: "VALIDATION_ERROR" });
  assert.equal(f.calls.length, 0);
});

test("provider failure never falls back or exposes diagnostics", async () => {
  const f = fixture();
  f.provider.acquire = async () => { throw new Error("opaque-fixture-secret"); };
  await assert.rejects(f.auth.credential(cert), error => error.code === "AUTH_REQUIRED" && error.suggestions.some(item => item.includes("application consent")) && !JSON.stringify(error).includes("opaque-fixture-secret"));
});

test("cache serves within the margin and refreshes at it with fake time", async () => {
  const f = fixture();
  await f.auth.credential(federated);
  await f.auth.credential(federated);
  assert.equal(f.calls.length, 1);
  f.advance(60_000);
  await f.auth.credential(federated);
  assert.equal(f.calls.length, 2);
});

test("failed refresh never returns a stale credential", async () => {
  const f = fixture();
  await f.auth.credential(cert);
  f.advance(120_000);
  f.provider.acquire = async () => { throw new Error("opaque-fixture-secret"); };
  await assert.rejects(f.auth.credential(cert), { code: "AUTH_REQUIRED" });
  await assert.rejects(f.auth.credential(cert), { code: "AUTH_REQUIRED" });
  assert.equal(f.calls.length, 1);
});

for (const [name, patch] of [
  ["wrong tenant", { tenantId: client }], ["wrong client", { clientId: tenant }],
  ["expired", { expiresAt: 1000 }], ["unknown expiry", { expiresAt: NaN }], ["missing token", { token: "" }],
]) test(`${name} credential is rejected without decoding the token`, async () => {
  const f = fixture();
  f.provider.acquire = async () => ({ token: "opaque-fixture-secret", expiresAt: 121_000, tenantId: tenant, clientId: client, ...patch });
  await assert.rejects(f.auth.credential(federated), { code: "AUTH_REQUIRED" });
});

test("identity changes cannot reuse a cached credential", async () => {
  const f = fixture();
  await f.auth.credential(federated);
  await f.auth.credential({ ...federated, credentialRef: { ...federated.credentialRef, key: tenant } });
  await f.auth.credential(cert);
  assert.equal(f.calls.length, 3);
});

test("delegated auth rejects application profiles before provider access", async () => {
  let time = 1000;
  const calls = [];
  const provider = { storage: "session-only", login: async (...args) => { calls.push(args); throw new Error("must not run"); }, silent: async (...args) => { calls.push(args); throw new Error("must not run"); } };
  const auth = new DelegatedAuth(provider, () => time);
  await assert.rejects(auth.login(cert, "browser", ["https://graph.microsoft.com/User.Read"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(auth.credential(federated, ["https://graph.microsoft.com/User.Read"]), { code: "VALIDATION_ERROR" });
  assert.equal(calls.length, 0);
});

for (const patch of [{ allowDeviceCode: true }, { mode: "delegated" }, { cloud: "government" }, { tenantId: "common" }, { enabledPacks: ["entra", "mail"] },
  { credentialRef: { provider: "os-or-session", key } }, { credentialRef: { provider: "certificate", key } },
  { credentialRef: { provider: "certificate", key, thumbprint: "zzzz" } }, { credentialRef: { provider: "federated", key, extra: true } },
  { credentialRef: { provider: "certificate", key, thumbprint, privateKey: "opaque-fixture-secret" } },
]) test(`invalid application profile ${JSON.stringify(patch)} fails before provider access`, async () => {
  const f = fixture();
  await assert.rejects(f.auth.credential({ ...federated, ...patch }), { code: "VALIDATION_ERROR" });
  assert.equal(f.calls.length, 0);
});

test("application creation stores references while rejecting impossible combinations", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-app-profile-"));
  try {
    const profiles = new Profiles(join(dir, "config.json"));
    profiles.create("batch", tenant, client, "commercial", false, { federated: true });
    profiles.create("daemon", tenant, client, "commercial", false, { certificateThumbprint: thumbprint.toUpperCase() });
    assert.equal(profiles.resolve("batch").profile.mode, "application");
    assert.equal(profiles.resolve("daemon").profile.credentialRef.thumbprint, thumbprint);
    assert.notEqual(profiles.resolve("batch").profile.credentialRef.key, profiles.resolve("daemon").profile.credentialRef.key);
    assert.equal(profiles.list().filter(item => item.mode === "application").length, 2);
    const saved = readFileSync(profiles.path, "utf8");
    assert.doesNotMatch(saved, /privateKey|opaque-fixture-secret/);
    for (const args of [
      ["both", false, { certificateThumbprint: thumbprint, federated: true }],
      ["neither", false, {}],
      ["device-code", true, { federated: true }],
      ["bad-thumbprint", false, { certificateThumbprint: "zzzz" }],
    ]) assert.throws(() => profiles.create(args[0], tenant, client, "commercial", args[1], args[2]), { code: "VALIDATION_ERROR" });
    writeFileSync(profiles.path, JSON.stringify({ version: 1, profiles: { bad: { ...federated, allowDeviceCode: true } } }));
    assert.throws(() => profiles.list(), { code: "VALIDATION_ERROR" });
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("packaged application journey rejects interactive login with closed stdin", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-app-cli-"));
  const run = args => spawnSync(process.execPath, ["dist/bin/mg-axi.js", ...args], { encoding: "utf8", input: "", timeout: 10_000, env: { PATH: process.env.PATH, MG_AXI_CONFIG: join(dir, "config.json") } });
  try {
    assert.equal(run(["profile", "create", "--name", "batch", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--federated"]).status, 0);
    assert.equal(run(["profile", "create", "--name", "daemon", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--certificate-thumbprint", thumbprint]).status, 0);
    assert.equal(decode(run(["profile", "show", "--profile", "batch"]).stdout).profile.mode, "application");
    assert.equal(decode(run(["profile", "show", "--profile", "daemon"]).stdout).profile.credentialRef.provider, "certificate");
    for (const name of ["batch", "daemon"]) {
      const rejected = run(["login", "--profile", name, "--scopes", "https://graph.microsoft.com/User.Read"]);
      assert.equal(rejected.status, 2);
      assert.equal(rejected.stderr, "");
      assert.match(decode(rejected.stdout).error, /client credentials/);
    }
    for (const args of [
      ["profile", "create", "--name", "x2", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application"],
      ["profile", "create", "--name", "x3", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--federated", "--certificate-thumbprint", thumbprint],
      ["profile", "create", "--name", "x4", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--certificate-thumbprint", "short"],
      ["profile", "create", "--name", "x5", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--federated", "--allow-device-code"],
      ["profile", "create", "--name", "x6", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--certificate-thumbprint", thumbprint],
      ["profile", "create", "--name", "x7", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "workload"],
    ]) {
      const result = run(args);
      assert.equal(result.status, 2, args.join(" "));
      assert.equal(result.stderr, "");
      assert.equal(decode(result.stdout).code, "VALIDATION_ERROR");
    }
    assert.equal(decode(run(["profile", "list"]).stdout).profiles.filter(item => item.mode === "application").length, 2);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
