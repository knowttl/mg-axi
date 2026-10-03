import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";

for (const operation of ["browser", "device-code", "silent", "read"]) test(`${operation} cache failure retains session credentials and truthful storage`, () => {
  const result = spawnSync(process.execPath, ["--experimental-test-module-mocks", "--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import { mock } from "node:test";
    const operation = ${JSON.stringify(operation)};
    const tenant = "11111111-1111-4111-8111-111111111111";
    const profile = { tenantId: tenant, clientId: "22222222-2222-4222-8222-222222222222", cloud: "commercial", credentialRef: { key: "first" } };
    const scopes = ["https://graph.microsoft.com/User.Read"];
    const saved = new Map();
    let rejectWrites = operation === "browser" || operation === "device-code";
    let rejectReads = false;
    mock.module("keytar", { defaultExport: {
      getPassword: async (_, key) => {
        if (rejectReads) throw new Error("Store unavailable");
        return saved.get(key) ?? null;
      },
      setPassword: async (_, key, value) => {
        if (rejectWrites && Buffer.byteLength(value) > 2560) throw new Error("Credential blob too large");
        saved.set(key, value);
      },
    } });
    class Application {
      constructor(config) {
        this.plugin = config.cache?.cachePlugin;
        this.accounts = [];
        this.token = "";
        this.cache = {
          serialize: () => JSON.stringify({ accounts: this.accounts, token: this.token, padding: "x".repeat(3000) }),
          deserialize: value => { const parsed = JSON.parse(value); this.accounts = parsed.accounts; this.token = parsed.token; },
          getAllAccounts: async () => { await this.access(false); return [...this.accounts]; },
          removeAccount: async () => { this.accounts = []; await this.access(true); },
        };
      }
      getTokenCache() { return this.cache; }
      async access(changed) {
        const context = { tokenCache: this.cache, cacheHasChanged: changed };
        if (!changed) await this.plugin?.beforeCacheAccess(context);
        await this.plugin?.afterCacheAccess(context);
      }
      async acquireTokenInteractive() { return this.acquire("login-token"); }
      async acquireTokenByDeviceCode() { return this.acquire("login-token"); }
      async acquireTokenSilent() { return this.acquire(this.token + "-refreshed"); }
      async acquire(token) {
        this.accounts = [{ tenantId: tenant, homeAccountId: "account" }];
        this.token = token;
        await this.access(true);
        return { account: this.accounts[0], tenantId: tenant, accessToken: token, expiresOn: new Date(Date.now() + 120000) };
      }
    }
    mock.module("@azure/msal-node", { namedExports: { PublicClientApplication: Application } });
    const { MsalProvider } = await import("./dist/msal-provider.js");
    const provider = new MsalProvider();
    assert.equal((await provider.login(profile, operation === "device-code" ? operation : "browser", scopes)).token, "login-token");
    if (operation === "silent" || operation === "read") assert.equal(provider.storage, "os-protected");
    rejectWrites = true;
    rejectReads = operation === "read";
    assert.equal((await provider.silent(profile, scopes)).token, "login-token-refreshed");
    assert.equal(provider.storage, "session-only");
    assert.equal((await provider.silent(profile, scopes)).token, "login-token-refreshed-refreshed");
    rejectWrites = rejectReads = false;
    await provider.login({ ...profile, credentialRef: { key: "second" } }, "browser", scopes);
    assert.equal(provider.storage, "os-protected");
    assert.equal((await provider.silent(profile, scopes)).token, "login-token-refreshed-refreshed-refreshed");
    assert.equal(provider.storage, "session-only");
  `], { encoding: "utf8", timeout: 10_000 });
  assert.equal(result.status, 0, result.stderr);
});
