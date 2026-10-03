import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";

const cases = [
  { name: "unavailable import uses session storage", failure: "import", storage: "session-only", persisted: [] },
  { name: "initial read failure blocks replacement", failure: "initial-read", error: "LOGIN_FAILED", persisted: ["account-A"] },
  { name: "removal read failure blocks replacement", failure: "removal-read", error: "LOGIN_FAILED", persisted: [] },
  { name: "invalidation failure blocks replacement", failure: "delete", error: "LOGIN_FAILED", persisted: ["account-A"] },
  { name: "removal write failure blocks replacement", failure: "removal-write", error: "LOGIN_FAILED", persisted: [] },
  { name: "ordinary token write failure blocks login", failure: "token-write", error: "LOGIN_FAILED", persisted: [] },
  { name: "non-Windows oversized write errors fail closed", platform: "linux", large: true, failure: "token-write", error: "LOGIN_FAILED", persisted: [] },
  { name: "refresh read failure requires authentication", failure: "initial-read", refresh: true, error: "AUTH_REQUIRED", persisted: ["account-A"] },
  { name: "refresh write failure requires authentication", failure: "token-write", refresh: true, error: "AUTH_REQUIRED", persisted: ["account-A"] },
  { name: "Windows oversized login uses memory after invalidation", large: true, storage: "session-only", persisted: [] },
  { name: "Windows oversized refresh invalidates the old cache", large: true, refresh: true, storage: "session-only", persisted: [] },
  { name: "Windows oversized login fails when invalidation fails", large: true, failure: "oversize-delete", error: "LOGIN_FAILED", persisted: [] },
  { name: "Windows oversized refresh fails when invalidation fails", large: true, refresh: true, failure: "oversize-delete", error: "AUTH_REQUIRED", persisted: ["account-A"] },
  { name: "Windows cache limit measures UTF-8 bytes", unicode: true, storage: "session-only", persisted: [] },
  { name: "Windows cache at the byte limit remains protected", limit: true, storage: "os-protected", persisted: ["account-B"] },
  { name: "ordinary replacement persists the new account", storage: "os-protected", persisted: ["account-B"] },
  { name: "replacement clears every old account", multiple: true, large: true, storage: "session-only", persisted: [] },
];

for (const method of ["browser", "device-code"]) for (const scenario of cases) test(`${method}: ${scenario.name}`, () => {
  const result = spawnSync(process.execPath, ["--experimental-test-module-mocks", "--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import { registerHooks } from "node:module";
    import { mock } from "node:test";
    const scenario = ${JSON.stringify(scenario)};
    const method = ${JSON.stringify(method)};
    const tenant = "11111111-1111-4111-8111-111111111111";
    const profile = { mode: "delegated", tenantId: tenant, clientId: "22222222-2222-4222-8222-222222222222", cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: true, credentialRef: { provider: "os-or-session", key: "33333333-3333-4333-8333-333333333333" } };
    const scopes = ["https://graph.microsoft.com/User.Read"];
    const saved = new Map();
    let failure;
    let reads = 0;
    let deletes = 0;
    let large = false;
    let unicode = false;
    let limit = false;
    let selectedAccount = "account-A";
    const prompts = [];
    mock.module("keytar", { defaultExport: {
      getPassword: async (_, key) => {
        reads++;
        if (failure === "initial-read" || failure === "removal-read" && reads === 2) throw new Error("Read failed: secret diagnostics");
        return saved.get(key) ?? null;
      },
      setPassword: async (_, key, value) => {
        const cache = JSON.parse(value);
        if (failure === "token-write" && cache.accounts.length || failure === "removal-write" && !cache.accounts.length) throw new Error("Write failed: secret diagnostics");
        saved.set(key, value);
      },
      deletePassword: async (_, key) => {
        deletes++;
        if (failure === "delete" || failure === "oversize-delete" && (scenario.refresh || deletes > 1)) throw new Error("Delete failed: secret diagnostics");
        return saved.delete(key);
      },
    } });
    registerHooks({ resolve(specifier, context, nextResolve) {
      if (specifier === "keytar" && failure === "import") throw new Error("Module unavailable");
      return nextResolve(specifier, context);
    } });
    class Application {
      constructor(config) {
        this.plugin = config.cache?.cachePlugin;
        this.accounts = [];
        this.token = "";
        this.cache = {
          serialize: () => {
            const cache = { accounts: this.accounts, token: this.token, padding: "" };
            const bytes = this.accounts.length ? (large ? 3000 : limit ? 2560 : 500) : 500;
            cache.padding = unicode && this.accounts.length ? "é".repeat(1300) : "x".repeat(bytes - Buffer.byteLength(JSON.stringify(cache)));
            return JSON.stringify(cache);
          },
          deserialize: value => { const cache = JSON.parse(value); this.accounts = cache.accounts; this.token = cache.token; },
          getAllAccounts: async () => { await this.access(false); return [...this.accounts]; },
          removeAccount: async account => {
            await this.access(false);
            this.accounts = this.accounts.filter(item => item.homeAccountId !== account.homeAccountId);
            await this.access(true);
          },
        };
      }
      getTokenCache() { return this.cache; }
      async access(changed) {
        const context = { tokenCache: this.cache, cacheHasChanged: changed };
        if (!changed) await this.plugin?.beforeCacheAccess(context);
        await this.plugin?.afterCacheAccess(context);
      }
      async acquireTokenInteractive() { prompts.push("browser"); return this.acquire("login-token", selectedAccount); }
      async acquireTokenByDeviceCode() { prompts.push("device-code"); return this.acquire("login-token", selectedAccount); }
      async acquireTokenSilent() { return this.acquire(this.token + "-refreshed", this.accounts[0].homeAccountId); }
      async acquire(token, accountId) {
        await this.access(false);
        this.accounts = [{ tenantId: tenant, homeAccountId: accountId }];
        this.token = token;
        await this.access(true);
        return { account: this.accounts[0], tenantId: tenant, accessToken: token, expiresOn: new Date(Date.now() + 120000) };
      }
    }
    mock.module("@azure/msal-node", { namedExports: { PublicClientApplication: Application } });
    const { MsalProvider } = await import("./dist/msal-provider.js");
    const { DelegatedAuth } = await import("./dist/auth.js");
    Object.defineProperty(process, "platform", { value: scenario.platform ?? "win32" });
    if (scenario.failure !== "import") await new MsalProvider().login(profile, "browser", scopes);
    if (scenario.multiple) {
      for (const [key, value] of saved) {
        const cache = JSON.parse(value);
        cache.accounts.push({ tenantId: tenant, homeAccountId: "account-C" });
        saved.set(key, JSON.stringify(cache));
      }
    }
    prompts.length = 0;
    reads = deletes = 0;
    failure = scenario.failure;
    large = !!scenario.large;
    unicode = !!scenario.unicode;
    limit = !!scenario.limit;
    selectedAccount = "account-B";
    const provider = new MsalProvider();
    const auth = new DelegatedAuth(provider);
    const action = () => scenario.refresh ? auth.credential(profile, scopes) : auth.login(profile, method, scopes);
    if (scenario.error) {
      await assert.rejects(action(), error => error.code === scenario.error && error.suggestions.some(item => item.includes("OS credential store")) && !JSON.stringify(error).includes("secret diagnostics"));
      assert.equal(provider.storage, "os-protected");
    } else {
      const result = await action();
      if (!scenario.refresh) assert.equal(result.storage, scenario.storage);
      assert.equal(provider.storage, scenario.storage);
      assert.equal((await new DelegatedAuth(provider).credential(profile, scopes)).accountId, scenario.refresh ? "account-A" : "account-B");
    }
    const persisted = [...saved.values()].flatMap(value => JSON.parse(value).accounts.map(account => account.homeAccountId)).sort();
    assert.deepEqual(persisted, scenario.persisted);
    if (scenario.refresh || ["initial-read", "removal-read", "delete", "removal-write"].includes(scenario.failure)) assert.deepEqual(prompts, []);
    else assert.deepEqual(prompts, [method]);
    failure = undefined;
    large = unicode = limit = false;
    if (scenario.error) {
      const recovered = new DelegatedAuth(new MsalProvider());
      assert.equal((await recovered.login(profile, method, scopes)).storage, "os-protected");
      assert.equal((await new MsalProvider().silent(profile, scopes)).accountId, "account-B");
    } else if (scenario.storage === "session-only" && scenario.failure !== "import") {
      await assert.rejects(new DelegatedAuth(new MsalProvider()).credential(profile, scopes), { code: "AUTH_REQUIRED" });
      await provider.login({ ...profile, credentialRef: { ...profile.credentialRef, key: tenant } }, method, scopes);
      assert.equal(provider.storage, "os-protected");
      assert.equal((await provider.silent(profile, scopes)).accountId, scenario.refresh ? "account-A" : "account-B");
      assert.equal(provider.storage, "session-only");
    }
  `], { encoding: "utf8", timeout: 10_000 });
  assert.equal(result.status, 0, result.stderr);
});
