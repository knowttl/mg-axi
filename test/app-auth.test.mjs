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
import { AxiError } from "axi-sdk-js";

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

const appProfile = ref => ({ mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: ref });
const secretEnvRef = { provider: "client-secret", key, secretEnv: "MG_AXI_CLIENT_SECRET" };
const secretFileRef = { provider: "client-secret", key, secretFile: "/run/secrets/mg-axi-secret" };
const keyFileRef = { provider: "certificate", key, thumbprint, keyFile: "/run/secrets/mg-axi-key.pem" };
const keyEnvRef = { provider: "certificate", key, thumbprint, keyEnv: "MG_AXI_PRIVATE_KEY" };
const namedFederatedRef = { provider: "federated", key, tokenFileEnv: "MG_AXI_ASSERTION_FILE" };

test("headless reference shapes validate and cache per reference", async () => {
  const f = fixture();
  for (const ref of [secretEnvRef, secretFileRef, keyFileRef, keyEnvRef, namedFederatedRef]) {
    const credential = await f.auth.credential(appProfile(ref));
    assert.equal(credential.token, "opaque-fixture-secret");
  }
  assert.equal(f.calls.length, 5);
  await f.auth.credential(appProfile(namedFederatedRef));
  assert.equal(f.calls.length, 5);
  await f.auth.credential(appProfile({ ...secretEnvRef, secretEnv: "MG_AXI_OTHER_SECRET" }));
  assert.equal(f.calls.length, 6);
});

test("reference-naming AUTH_REQUIRED survives while other failures stay generic", async () => {
  const f = fixture();
  process.env.MG_AXI_CLIENT_SECRET = "planted-fake-secret-value";
  try {
    f.provider.acquire = async () => { throw new AxiError("Application credential unavailable for the configured reference MG_AXI_CLIENT_SECRET", "AUTH_REQUIRED", ["retry"]); };
    await assert.rejects(f.auth.credential(appProfile(secretEnvRef)), error => error.code === "AUTH_REQUIRED" && error.message.includes("MG_AXI_CLIENT_SECRET") && !JSON.stringify(error).includes("planted-fake-secret-value"));
    assert.doesNotMatch(JSON.stringify(secretEnvRef), /planted-fake-secret-value/);
    f.provider.acquire = async () => { throw new Error("planted-fake-secret-value"); };
    await assert.rejects(f.auth.credential(appProfile(secretEnvRef)), error => error.code === "AUTH_REQUIRED" && !JSON.stringify(error).includes("planted-fake-secret-value"));
  } finally { delete process.env.MG_AXI_CLIENT_SECRET; }
});

test("frozen old-shape application configs load and behave unchanged", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-app-compat-"));
  try {
    writeFileSync(join(dir, "config.json"), JSON.stringify({ version: 1, profiles: {
      daemon: { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "certificate", key, thumbprint: thumbprint.toUpperCase() } },
      batch: { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, credentialRef: { provider: "federated", key: client } },
    } }));
    const profiles = new Profiles(join(dir, "config.json"));
    assert.equal(profiles.resolve("daemon").profile.credentialRef.thumbprint, thumbprint);
    assert.deepEqual(profiles.resolve("daemon").profile.credentialRef, { provider: "certificate", key, thumbprint });
    assert.deepEqual(profiles.resolve("batch").profile.credentialRef, { provider: "federated", key: client });
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

for (const patch of [{ allowDeviceCode: true }, { mode: "delegated" }, { cloud: "government" }, { tenantId: "common" }, { enabledPacks: ["entra", "mail"] },
  { credentialRef: { provider: "os-or-session", key } }, { credentialRef: { provider: "certificate", key } },
  { credentialRef: { provider: "certificate", key, thumbprint: "zzzz" } }, { credentialRef: { provider: "federated", key, extra: true } },
  { credentialRef: { provider: "certificate", key, thumbprint, privateKey: "opaque-fixture-secret" } },
  { credentialRef: { provider: "client-secret", key } },
  { credentialRef: { provider: "client-secret", key, secretEnv: "MG_AXI_CLIENT_SECRET", secretFile: "/run/secrets/x" } },
  { credentialRef: { provider: "client-secret", key, secretEnv: "1BAD-NAME" } },
  { credentialRef: { provider: "client-secret", key, secretFile: "   " } },
  { credentialRef: { provider: "client-secret", key, secretEnv: "MG_AXI_CLIENT_SECRET", clientSecret: "opaque-fixture-secret" } },
  { credentialRef: { provider: "client-secret", key, secret: "opaque-fixture-secret" } },
  { credentialRef: { provider: "certificate", key, thumbprint, keyFile: "/run/secrets/k.pem", keyEnv: "MG_AXI_PRIVATE_KEY" } },
  { credentialRef: { provider: "certificate", key, thumbprint, keyEnv: "BAD-NAME" } },
  { credentialRef: { provider: "certificate", key, thumbprint, keyFile: "" } },
  { credentialRef: { provider: "certificate", key, thumbprint, keyMaterial: "opaque-fixture-secret" } },
  { credentialRef: { provider: "federated", key, tokenFileEnv: "BAD-NAME" } },
  { credentialRef: { provider: "federated", key, tokenFileEnv: "AZURE_FEDERATED_TOKEN_FILE", extra: true } },
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

test("headless creation stores references while rejecting impossible combinations", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-app-headless-"));
  try {
    const profiles = new Profiles(join(dir, "config.json"));
    profiles.create("svc-env", tenant, client, "commercial", false, { clientSecretEnv: "MG_AXI_CLIENT_SECRET" });
    profiles.create("svc-file", tenant, client, "commercial", false, { clientSecretFile: "/run/secrets/mg-axi-secret" });
    profiles.create("key-file", tenant, client, "commercial", false, { certificateThumbprint: thumbprint, certificateKeyFile: "/run/secrets/mg-axi-key.pem" });
    profiles.create("key-env", tenant, client, "commercial", false, { certificateThumbprint: thumbprint, certificateKeyEnv: "MG_AXI_PRIVATE_KEY" });
    profiles.create("fed-named", tenant, client, "commercial", false, { federated: true, federatedTokenFileEnv: "MG_AXI_ASSERTION_FILE" });
    assert.deepEqual(profiles.resolve("svc-env").profile.credentialRef, { provider: "client-secret", key: profiles.resolve("svc-env").profile.credentialRef.key, secretEnv: "MG_AXI_CLIENT_SECRET" });
    assert.equal(profiles.resolve("key-file").profile.credentialRef.keyFile, "/run/secrets/mg-axi-key.pem");
    assert.equal(profiles.resolve("fed-named").profile.credentialRef.tokenFileEnv, "MG_AXI_ASSERTION_FILE");
    const saved = readFileSync(profiles.path, "utf8");
    assert.doesNotMatch(saved, /privateKey|clientSecret|opaque-fixture-secret/);
    for (const args of [
      ["both-secret", false, { clientSecretEnv: "A", clientSecretFile: "/run/secrets/x" }],
      ["no-secret-holder", false, { clientSecretEnv: undefined, clientSecretFile: undefined, federated: false }],
      ["bad-secret-env", false, { clientSecretEnv: "1BAD" }],
      ["empty-secret-file", false, { clientSecretFile: "  " }],
      ["both-keys", false, { certificateThumbprint: thumbprint, certificateKeyFile: "/k", certificateKeyEnv: "K" }],
      ["key-without-thumbprint", false, { certificateKeyEnv: "MG_AXI_PRIVATE_KEY" }],
      ["token-env-without-federated", false, { federatedTokenFileEnv: "MG_AXI_ASSERTION_FILE" }],
      ["two-methods", false, { certificateThumbprint: thumbprint, clientSecretEnv: "MG_AXI_CLIENT_SECRET" }],
      ["three-methods", false, { certificateThumbprint: thumbprint, federated: true }],
    ]) assert.throws(() => profiles.create(args[0], tenant, client, "commercial", args[1], args[2]), { code: "VALIDATION_ERROR" }, args[0]);
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

test("packaged headless journey records references, never values", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-app-headless-cli-"));
  const run = args => spawnSync(process.execPath, ["dist/bin/mg-axi.js", ...args], { encoding: "utf8", input: "", timeout: 10_000, env: { PATH: process.env.PATH, MG_AXI_CONFIG: join(dir, "config.json") } });
  try {
    assert.equal(run(["profile", "create", "--name", "svc", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--client-secret-env", "MG_AXI_CLIENT_SECRET"]).status, 0);
    assert.equal(run(["profile", "create", "--name", "keyed", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--certificate-thumbprint", thumbprint, "--certificate-key-file", "/run/secrets/mg-axi-key.pem"]).status, 0);
    assert.equal(run(["profile", "create", "--name", "fed", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--federated", "--federated-token-file-env", "MG_AXI_ASSERTION_FILE"]).status, 0);
    assert.deepEqual(decode(run(["profile", "show", "--profile", "svc"]).stdout).profile.credentialRef.provider, "client-secret");
    assert.equal(decode(run(["profile", "show", "--profile", "keyed"]).stdout).profile.credentialRef.keyFile, "/run/secrets/mg-axi-key.pem");
    for (const args of [
      ["profile", "create", "--name", "h2", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--client-secret-env", "A", "--client-secret-file", "/run/secrets/x"],
      ["profile", "create", "--name", "h3", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--certificate-key-env", "K"],
      ["profile", "create", "--name", "h4", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--client-secret-env", "MG_AXI_CLIENT_SECRET"],
      ["profile", "create", "--name", "h5", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--client-secret-env", "1BAD"],
    ]) {
      const result = run(args);
      assert.equal(result.status, 2, args.join(" "));
      assert.equal(result.stderr, "");
      assert.equal(decode(result.stdout).code, "VALIDATION_ERROR");
    }
    assert.doesNotMatch(readFileSync(join(dir, "config.json"), "utf8"), /privateKey|clientSecret/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
