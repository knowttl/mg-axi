import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { spawnSync } from "node:child_process";
import { decode } from "@toon-format/toon";
import { Profiles, type ApplicationProfile, type DelegatedProfile } from "../dist/profiles.js";
import { ApplicationAuth, type ApplicationCredentialProvider } from "../dist/app-auth.js";
import { DelegatedAuth, type CredentialProvider } from "../dist/auth.js";
import { AxiError } from "axi-sdk-js";
import type { SecretRef } from "../dist/secrets.js";
import { record, records, text, strings, delegatedProvider } from "./fixtures/typed.ts";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const audience = "https://graph.microsoft.com/.default";
const thumbprint = "ab".repeat(32);
const cert = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, auth: "certificate", thumbprintSha256: thumbprint, privateKey: { env: "MG_AXI_PRIVATE_KEY" } } satisfies ApplicationProfile;
const federated = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, auth: "federated", token: { env: "MG_AXI_TOKEN" } } satisfies ApplicationProfile;
function fixture() {
  let time = 1000;
  const calls: ["acquire", ...Parameters<ApplicationCredentialProvider["acquire"]>][] = [];
  const credential = () => ({ token: "opaque-fixture-secret", expiresAt: time + 120_000, tenantId: tenant, clientId: client });
  const provider: ApplicationCredentialProvider = {
    storage: "session-only",
    acquire: async (...args) => { calls.push(["acquire", ...args]); return credential(); },
  };
  return { provider, calls, auth: new ApplicationAuth(provider, () => time), advance: (amount: number) => { time += amount; } };
}

test("application credential requests the configured Graph .default audience", async () => {
  const f = fixture();
  const credential = await f.auth.credential(federated);
  assert.equal(credential.token, "opaque-fixture-secret");
  assert.deepEqual(f.calls, [["acquire", federated, [audience]]]);
});

test("application auth exposes no interactive login surface", () => {
  const f = fixture();
  assert.equal(typeof (f.auth as unknown as { login?: unknown }).login, "undefined");
});

for (const requested of [[], ["https://graph.microsoft.com/User.Read"], ["https://graph.microsoft.com/.default", "https://graph.microsoft.com/User.Read"], ["https://example.invalid/.default"]]) test(`non-audience scopes ${JSON.stringify(requested)} fail before provider access`, async () => {
  const f = fixture();
  await assert.rejects(f.auth.credential(cert, requested), { code: "VALIDATION_ERROR" });
  assert.equal(f.calls.length, 0);
});

test("provider failure never falls back or exposes diagnostics", async () => {
  const f = fixture();
  f.provider.acquire = async () => { throw new Error("opaque-fixture-secret"); };
  await assert.rejects(f.auth.credential(cert), error => record(error).code === "AUTH_REQUIRED" && strings(record(error).suggestions).some(item => item.includes("application consent")) && !JSON.stringify(error).includes("opaque-fixture-secret"));
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
] as const) test(`${name} credential is rejected without decoding the token`, async () => {
  const f = fixture();
  f.provider.acquire = async () => ({ token: "opaque-fixture-secret", expiresAt: 121_000, tenantId: tenant, clientId: client, ...patch });
  await assert.rejects(f.auth.credential(federated), { code: "AUTH_REQUIRED" });
});

test("identity changes cannot reuse a cached credential", async () => {
  const f = fixture();
  await f.auth.credential(federated);
  await f.auth.credential({ ...federated, token: { env: "OTHER_TOKEN" } });
  await f.auth.credential(cert);
  assert.equal(f.calls.length, 3);
});

test("delegated auth rejects application profiles before provider access", async () => {
  let time = 1000;
  const calls: unknown[][] = [];
  const provider: Omit<CredentialProvider, "logout"> = { storage: "session-only", login: async (...args) => { calls.push(args); throw new Error("must not run"); }, silent: async (...args) => { calls.push(args); throw new Error("must not run"); } };
  const auth = new DelegatedAuth(delegatedProvider(provider), () => time);
  await assert.rejects(auth.login(cert as unknown as DelegatedProfile, "browser", ["https://graph.microsoft.com/User.Read"]), { code: "CONFIG_INVALID" });
  await assert.rejects(auth.credential(federated as unknown as DelegatedProfile, ["https://graph.microsoft.com/User.Read"]), { code: "CONFIG_INVALID" });
  assert.equal(calls.length, 0);
});

type AppReference = { auth: "client-secret"; clientSecret: SecretRef } | { auth: "certificate"; thumbprintSha256: string; privateKey: SecretRef } | { auth: "federated"; token: SecretRef };
const appProfile = (ref: AppReference): ApplicationProfile => ({ mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false, ...ref });
const secretEnvRef = { auth: "client-secret", clientSecret: { env: "MG_AXI_CLIENT_SECRET" } } satisfies AppReference;
const secretFileRef = { auth: "client-secret", clientSecret: { file: "/run/secrets/mg-axi-secret" } } satisfies AppReference;
const keyFileRef = { auth: "certificate", thumbprintSha256: thumbprint, privateKey: { file: "/run/secrets/mg-axi-key.pem" } } satisfies AppReference;
const keyEnvRef = { auth: "certificate", thumbprintSha256: thumbprint, privateKey: { env: "MG_AXI_PRIVATE_KEY" } } satisfies AppReference;
const namedFederatedRef = { auth: "federated", token: { file: "/run/secrets/assertion" } } satisfies AppReference;

test("headless reference shapes validate and cache per reference", async () => {
  const f = fixture();
  for (const ref of [secretEnvRef, secretFileRef, keyFileRef, keyEnvRef, namedFederatedRef]) {
    const credential = await f.auth.credential(appProfile(ref));
    assert.equal(credential.token, "opaque-fixture-secret");
  }
  assert.equal(f.calls.length, 5);
  await f.auth.credential(appProfile(namedFederatedRef));
  assert.equal(f.calls.length, 5);
  await f.auth.credential(appProfile({ ...secretEnvRef, clientSecret: { env: "MG_AXI_OTHER_SECRET" } }));
  assert.equal(f.calls.length, 6);
});

test("reference-naming AUTH_REQUIRED survives while other failures stay generic", async () => {
  const f = fixture();
  process.env.MG_AXI_CLIENT_SECRET = "planted-fake-secret-value";
  try {
    f.provider.acquire = async () => { throw new AxiError("Application credential unavailable for the configured reference MG_AXI_CLIENT_SECRET", "AUTH_REQUIRED", ["retry"]); };
    await assert.rejects(f.auth.credential(appProfile(secretEnvRef)), error => record(error).code === "AUTH_REQUIRED" && text(record(error).message).includes("MG_AXI_CLIENT_SECRET") && !JSON.stringify(error).includes("planted-fake-secret-value"));
    assert.doesNotMatch(JSON.stringify(secretEnvRef), /planted-fake-secret-value/);
    f.provider.acquire = async () => { throw new Error("planted-fake-secret-value"); };
    await assert.rejects(f.auth.credential(appProfile(secretEnvRef)), error => record(error).code === "AUTH_REQUIRED" && !JSON.stringify(error).includes("planted-fake-secret-value"));
  } finally { delete process.env.MG_AXI_CLIENT_SECRET; }
});

test("flattened application configs normalize SHA-256 thumbprints and SecretRefs", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-app-compat-"));
  try {
    writeFileSync(join(dir, "config.json"), JSON.stringify({ version: 1, profiles: {
      daemon: { ...cert, thumbprintSha256: thumbprint.toUpperCase() },
      batch: federated,
    } }));
    const profiles = new Profiles(join(dir, "config.json"));
    assert.equal(record(profiles.resolve("daemon").profile).thumbprintSha256, thumbprint);
    assert.deepEqual(record(profiles.resolve("daemon").profile).privateKey, { env: "MG_AXI_PRIVATE_KEY" });
    assert.deepEqual(record(profiles.resolve("batch").profile).token, { env: "MG_AXI_TOKEN" });
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

for (const patch of [{ allowDeviceCode: true }, { mode: "delegated" }, { cloud: "government" }, { tenantId: "common" }, { enabledPacks: ["entra", "unknown"] },
  { credentialRef: { provider: "federated", key } }, { auth: "unknown" },
  { token: { env: "1BAD-NAME" } }, { token: { env: "TOKEN", extra: true } },
  { token: "opaque-fixture-secret" }, { token: { file: "relative" } },
  { token: { env: "TOKEN", file: "/run/secrets/x" } }, { token: { stored: true } },
]) test(`invalid application profile ${JSON.stringify(patch)} fails before provider access`, async () => {
  const f = fixture();
  await assert.rejects(f.auth.credential({ ...federated, ...patch } as unknown as ApplicationProfile), { code: "CONFIG_INVALID" });
  assert.equal(f.calls.length, 0);
});

test("application creation stores references while rejecting impossible combinations", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-app-profile-"));
  try {
    const profiles = new Profiles(join(dir, "config.json"));
    profiles.create("batch", tenant, client, "commercial", false, { federated: true });
    profiles.create("daemon", tenant, client, "commercial", false, { certificateThumbprint: thumbprint.toUpperCase() });
    assert.equal(profiles.resolve("batch").profile.mode, "application");
    assert.equal(record(profiles.resolve("daemon").profile).thumbprintSha256, thumbprint);
    assert.notEqual(record(profiles.resolve("batch").profile).auth, record(profiles.resolve("daemon").profile).auth);
    assert.equal(profiles.list().filter(item => item.mode === "application").length, 2);
    const saved = readFileSync(profiles.path, "utf8");
    assert.doesNotMatch(saved, /opaque-fixture-secret/);
    for (const args of [
      ["both", false, { certificateThumbprint: thumbprint, federated: true }],
      ["neither", false, {}],
      ["device-code", true, { federated: true }],
      ["bad-thumbprint", false, { certificateThumbprint: "zzzz" }],
    ] as const) assert.throws(() => profiles.create(args[0], tenant, client, "commercial", args[1], args[2]), { code: "CONFIG_INVALID" });
    writeFileSync(profiles.path, JSON.stringify({ version: 1, profiles: { bad: { ...federated, allowDeviceCode: true } } }));
    assert.throws(() => profiles.list(), { code: "CONFIG_INVALID" });
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
    profiles.create("fed-named", tenant, client, "commercial", false, { federated: true, federatedTokenFile: "/run/secrets/assertion" });
    assert.deepEqual(record(profiles.resolve("svc-env").profile).clientSecret, { env: "MG_AXI_CLIENT_SECRET" });
    assert.equal(record(record(profiles.resolve("key-file").profile).privateKey).file, "/run/secrets/mg-axi-key.pem");
    assert.equal(record(record(profiles.resolve("fed-named").profile).token).file, "/run/secrets/assertion");
    const saved = readFileSync(profiles.path, "utf8");
    assert.doesNotMatch(saved, /opaque-fixture-secret/);
    for (const args of [
      ["both-secret", false, { clientSecretEnv: "A", clientSecretFile: "/run/secrets/x" }],
      ["no-secret-holder", false, { clientSecretEnv: undefined, clientSecretFile: undefined, federated: false }],
      ["bad-secret-env", false, { clientSecretEnv: "1BAD" }],
      ["empty-secret-file", false, { clientSecretFile: "  " }],
      ["both-keys", false, { certificateThumbprint: thumbprint, certificateKeyFile: "/k", certificateKeyEnv: "K" }],
      ["key-without-thumbprint", false, { certificateKeyEnv: "MG_AXI_PRIVATE_KEY" }],
      ["token-file-without-federated", false, { federatedTokenFile: "/run/secrets/assertion" }],
      ["two-methods", false, { certificateThumbprint: thumbprint, clientSecretEnv: "MG_AXI_CLIENT_SECRET" }],
      ["three-methods", false, { certificateThumbprint: thumbprint, federated: true }],
    ] as const) assert.throws(() => profiles.create(args[0], tenant, client, "commercial", args[1], args[2]), { code: "CONFIG_INVALID" }, args[0]);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("packaged application journey rejects interactive login with closed stdin", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-app-cli-"));
  const run = (args: string[]) => spawnSync(process.execPath, ["dist/bin/mg-axi.js", ...args], { encoding: "utf8", input: "", timeout: 10_000, env: { PATH: process.env.PATH, HOME: dir, USERPROFILE: dir, XDG_CONFIG_HOME: join(dir, ".config"), MG_AXI_CONFIG: join(dir, "config.json") } });
  try {
    assert.equal(run(["profile", "create", "--name", "batch", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--federated"]).status, 0);
    assert.equal(run(["profile", "create", "--name", "daemon", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--certificate-thumbprint-sha256", thumbprint]).status, 0);
    assert.equal(record(record(decode(run(["profile", "show", "--profile", "batch"]).stdout)).profile).mode, "application");
    assert.equal(record(record(decode(run(["profile", "show", "--profile", "daemon"]).stdout)).profile).auth, "certificate");
    for (const name of ["batch", "daemon"]) {
      const rejected = run(["login", "--profile", name, "--scopes", "https://graph.microsoft.com/User.Read"]);
      assert.equal(rejected.status, 2);
      assert.equal(rejected.stderr, "");
      assert.match(text(record(decode(rejected.stdout)).error), /client credentials/);
    }
    for (const args of [
      ["profile", "create", "--name", "x2", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application"],
      ["profile", "create", "--name", "x3", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--federated", "--certificate-thumbprint-sha256", thumbprint],
      ["profile", "create", "--name", "x4", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--certificate-thumbprint-sha256", "short"],
      ["profile", "create", "--name", "x5", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--federated", "--allow-device-code"],
      ["profile", "create", "--name", "x6", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--certificate-thumbprint-sha256", thumbprint],
      ["profile", "create", "--name", "x7", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "workload"],
    ]) {
      const result = run(args);
      assert.notEqual(result.status, 0, args.join(" "));
      assert.equal(result.stderr, "");
      assert.match(text(record(decode(result.stdout)).code), /^(CONFIG_INVALID|VALIDATION_ERROR)$/);
    }
    assert.equal(records(record(decode(run(["profile", "list"]).stdout)).profiles).filter(item => item.mode === "application").length, 2);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("packaged headless journey records references, never values", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-app-headless-cli-"));
  const run = (args: string[]) => spawnSync(process.execPath, ["dist/bin/mg-axi.js", ...args], { encoding: "utf8", input: "", timeout: 10_000, env: { PATH: process.env.PATH, HOME: dir, USERPROFILE: dir, XDG_CONFIG_HOME: join(dir, ".config"), MG_AXI_CONFIG: join(dir, "config.json") } });
  try {
    assert.equal(run(["profile", "create", "--name", "svc", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--client-secret-env", "MG_AXI_CLIENT_SECRET"]).status, 0);
    assert.equal(run(["profile", "create", "--name", "keyed", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--certificate-thumbprint-sha256", thumbprint, "--certificate-key-file", "/run/secrets/mg-axi-key.pem"]).status, 0);
    assert.equal(run(["profile", "create", "--name", "fed", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--federated", "--federated-token-file", "/run/secrets/assertion"]).status, 0);
    assert.deepEqual(record(record(decode(run(["profile", "show", "--profile", "svc"]).stdout)).profile).auth, "client-secret");
    assert.equal(record(record(record(decode(run(["profile", "show", "--profile", "keyed"]).stdout)).profile).privateKey).file, "/run/secrets/mg-axi-key.pem");
    for (const args of [
      ["profile", "create", "--name", "h2", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--client-secret-env", "A", "--client-secret-file", "/run/secrets/x"],
      ["profile", "create", "--name", "h3", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--certificate-key-env", "K"],
      ["profile", "create", "--name", "h4", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--client-secret-env", "MG_AXI_CLIENT_SECRET"],
      ["profile", "create", "--name", "h5", "--tenant", tenant, "--client", client, "--cloud", "commercial", "--mode", "application", "--client-secret-env", "1BAD"],
    ]) {
      const result = run(args);
      assert.notEqual(result.status, 0, args.join(" "));
      assert.equal(result.stderr, "");
      assert.match(text(record(decode(result.stdout)).code), /^(CONFIG_INVALID|VALIDATION_ERROR)$/);
    }
    assert.doesNotMatch(readFileSync(join(dir, "config.json"), "utf8"), /BEGIN.*PRIVATE KEY|opaque-fixture-secret/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
