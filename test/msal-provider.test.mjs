import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";

const harness = `
import assert from "node:assert/strict";
import { chmodSync, existsSync, mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { mock } from "node:test";
import { Profiles, validateDelegatedProfile } from "./dist/profiles.js";
import { Credentials } from "./dist/credentials.js";
import { PrivateFiles } from "./dist/private-file.js";
const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const dir = mkdtempSync(join(tmpdir(), "mg-msal-file-"));
const profiles = new Profiles(join(dir, "config.json"));
profiles.create("first", tenant, client, "commercial", true);
profiles.create("second", tenant, client, "commercial", true);
const profile = profiles.resolve("first").profile;
const other = profiles.resolve("second").profile;
const aclEvents = [];
const files = new PrivateFiles("win32", (_path, directory, action) => { aclEvents.push({ directory, action }); });
const store = new Credentials(join(dir, "credentials.json"), files);
let account = "account-A";
const prompts = [];
const scopes = ["https://graph.microsoft.com/User.Read"];
class Application {
  constructor(config) {
    this.plugin = config.cache.cachePlugin;
    this.state = { Account: {}, AccessToken: {}, RefreshToken: {} };
    this.cache = {
      serialize: () => JSON.stringify(this.state),
      deserialize: text => { this.state = JSON.parse(text); },
      getAllAccounts: async () => { await this.plugin.beforeCacheAccess({ tokenCache: this.cache }); return Object.values(this.state.Account); },
    };
  }
  getTokenCache() { return this.cache; }
  async acquireTokenInteractive() { prompts.push("browser"); return this.acquire(account); }
  async acquireTokenByDeviceCode() { prompts.push("device-code"); return this.acquire(account); }
  async acquireTokenSilent() {
    await this.plugin.beforeCacheAccess({ tokenCache: this.cache });
    return this.acquire(Object.values(this.state.Account)[0].homeAccountId);
  }
  async acquire(id) {
    await this.plugin.beforeCacheAccess({ tokenCache: this.cache });
    this.state = { Account: { [id]: { tenantId: tenant, homeAccountId: id } },
      AccessToken: { [id]: { secret: "fixture-access-" + id } },
      RefreshToken: { [id]: { secret: "fixture-refresh-" + id + "é".repeat(2000) } } };
    await this.plugin.afterCacheAccess({ tokenCache: this.cache, cacheHasChanged: true });
    return { account: this.state.Account[id], tenantId: tenant, accessToken: "fixture-access-" + id, expiresOn: new Date(Date.now() + 120000) };
  }
}
mock.module("@azure/msal-node", { namedExports: { PublicClientApplication: Application } });
const { MsalProvider } = await import("./dist/msal-provider.js");
const { DelegatedAuth } = await import("./dist/auth.js");
try {
`;
const footer = `
} finally { rmSync(dir, { recursive: true, force: true }); }
`;
function run(script) {
  const result = spawnSync(process.execPath, ["--disable-warning=ExperimentalWarning", "--experimental-test-module-mocks", "--input-type=module", "-e", harness + script + footer], { encoding: "utf8", timeout: 10000 });
  assert.equal(result.status, 0, result.stderr);
}

for (const method of ["browser", "device-code"]) test(`${method} persists a Windows cache larger than 2560 bytes across providers`, () => {
  run(`
const auth = new DelegatedAuth(new MsalProvider(store));
assert.equal((await auth.login(profile, ${JSON.stringify(method)}, scopes)).storage, "protected-file");
assert.equal((await new MsalProvider(store).silent(profile, scopes)).accountId, "account-A");
const persisted = readFileSync(store.path, "utf8");
assert.ok(Buffer.byteLength(persisted) > 2560);
assert.doesNotMatch(persisted, /fixture-access-/);
assert.ok(aclEvents.some(entry => entry.action === "restrict" && !entry.directory));
assert.deepEqual(prompts, [${JSON.stringify(method)}]);
`);
});

test("replacement wipes the selected account while preserving another profile", () => {
  run(`
const provider = new MsalProvider(store);
await provider.login(profile, "browser", scopes);
await provider.login(other, "browser", scopes);
account = "account-B";
await provider.login(profile, "browser", scopes);
assert.equal((await new MsalProvider(store).silent(profile, scopes)).accountId, "account-B");
assert.equal((await new MsalProvider(store).silent(other, scopes)).accountId, "account-A");
assert.deepEqual(Object.keys(JSON.parse(store.read(profile)).Account), ["account-B"]);
`);
});

test("failed wipe blocks replacement before interactive authentication", () => {
  run(`
await new MsalProvider(store).login(profile, "browser", scopes);
prompts.length = 0;
mock.method(store, "remove", () => { throw new Error("fixture-private-diagnostic"); });
await assert.rejects(new DelegatedAuth(new MsalProvider(store)).login(profile, "browser", scopes), error => error.code === "LOGIN_FAILED" && !JSON.stringify(error).includes("fixture-private-diagnostic"));
assert.deepEqual(prompts, []);
assert.ok(store.read(profile));
`);
});

test("failed persistent write fails login and removes the unusable cache", () => {
  run(`
mock.method(store, "write", () => { throw new Error("fixture-private-diagnostic"); });
await assert.rejects(new DelegatedAuth(new MsalProvider(store)).login(profile, "browser", scopes), error => error.code === "LOGIN_FAILED" && !JSON.stringify(error).includes("fixture-private-diagnostic"));
assert.equal(store.read(profile), undefined);
assert.deepEqual(prompts, ["browser"]);
`);
});

test("file-holder delegated tokens resolve through the injected store files", () => {
  run(`
const fileProfile = validateDelegatedProfile({ mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", token: { file: "C:\\\\creds\\\\key.pem" } }, files);
await assert.rejects(new MsalProvider(store).silent(fileProfile, scopes), { code: "AUTH_REQUIRED" });
`);
});

for (const field of ["tenantId", "clientId"]) test(`stored credentials refuse a changed ${field} without interactive fallback`, () => {
  run(`
await new MsalProvider(store).login(profile, "browser", scopes);
const config = JSON.parse(readFileSync(profiles.path, "utf8"));
config.profiles.first.${field} = "33333333-3333-4333-8333-333333333333";
new PrivateFiles().write(profiles.path, JSON.stringify(config));
prompts.length = 0;
await assert.rejects(new DelegatedAuth(new MsalProvider(store)).credential(profiles.resolve("first").profile, scopes), { code: "AUTH_REQUIRED" });
assert.deepEqual(prompts, []);
`);
});

test("silent refresh never writes a new credential to disk", () => {
  run(`
await new MsalProvider(store).login(profile, "browser", scopes);
const before = readFileSync(store.path, "utf8");
mock.method(store, "write", () => { assert.fail("silent persistence is forbidden"); });
const provider = new MsalProvider(store);
assert.equal((await provider.silent(profile, scopes)).accountId, "account-A");
assert.equal(readFileSync(store.path, "utf8"), before);
`);
});

test("logout clears memory and the selected stored credential and is idempotent", () => {
  run(`
const auth = new DelegatedAuth(new MsalProvider(store));
await auth.login(profile, "browser", scopes);
await new MsalProvider(store).login(other, "browser", scopes);
await auth.logout(profile);
await auth.logout(profile);
await assert.rejects(auth.credential(profile, scopes), { code: "AUTH_REQUIRED" });
assert.ok(store.read(other));
assert.equal(store.read(profile), undefined);
await auth.logout(other);
assert.equal(existsSync(store.path), false);
`);
});
