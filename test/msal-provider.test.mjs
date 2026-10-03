import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";

for (const operation of ["browser", "device-code", "silent", "read", "replacement-browser", "replacement-device-code", "replacement-read", "replacement-session", "replacement-success-browser", "replacement-success-device-code", "replacement-success-multiple", "replacement-healthy"]) test(`${operation} preserves the credential storage contract`, () => {
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
    let rejectDeletes = false;
    let selectedAccount = "account-A";
    mock.module("keytar", { defaultExport: {
      getPassword: async (_, key) => {
        if (rejectReads) throw new Error("Store unavailable");
        return saved.get(key) ?? null;
      },
      setPassword: async (_, key, value) => {
        if (rejectWrites && Buffer.byteLength(value) > 2560) throw new Error("Credential blob too large");
        saved.set(key, value);
      },
      deletePassword: async (_, key) => {
        if (rejectDeletes) throw new Error("Invalidation failed");
        return saved.delete(key);
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
      async acquireTokenInteractive() { return this.acquire("login-token", selectedAccount); }
      async acquireTokenByDeviceCode() { return this.acquire("login-token", selectedAccount); }
      async acquireTokenSilent() { return this.acquire(this.token + "-refreshed", this.accounts[0].homeAccountId); }
      async acquire(token, accountId) {
        this.accounts = [{ tenantId: tenant, homeAccountId: accountId }];
        this.token = token;
        await this.access(true);
        return { account: this.accounts[0], tenantId: tenant, accessToken: token, expiresOn: new Date(Date.now() + 120000) };
      }
    }
    mock.module("@azure/msal-node", { namedExports: { PublicClientApplication: Application } });
    const { MsalProvider } = await import("./dist/msal-provider.js");
    const provider = new MsalProvider();
    assert.equal((await provider.login(profile, operation === "device-code" ? operation : "browser", scopes)).token, "login-token");
    if (operation.startsWith("replacement-")) {
      const method = operation.endsWith("device-code") ? "device-code" : "browser";
      if (operation === "replacement-success-multiple") {
        for (const [key, value] of saved) {
          const cache = JSON.parse(value);
          cache.accounts.push({ tenantId: tenant, homeAccountId: "account-C" });
          saved.set(key, JSON.stringify(cache));
        }
      }
      if (operation === "replacement-session") {
        rejectWrites = true;
        await provider.silent(profile, scopes);
        assert.equal(provider.storage, "session-only");
      }
      selectedAccount = "account-B";
      rejectWrites = operation !== "replacement-healthy";
      rejectReads = operation === "replacement-read";
      rejectDeletes = !operation.startsWith("replacement-success-") && operation !== "replacement-healthy";
      if (rejectDeletes) {
        await assert.rejects(provider.login(profile, method, scopes), /Invalidation failed/);
        rejectWrites = rejectReads = rejectDeletes = false;
        assert.equal((await new MsalProvider().silent(profile, scopes)).accountId, "account-A");
        assert.equal((await provider.login(profile, method, scopes)).accountId, "account-B");
        await assert.rejects(new MsalProvider().silent(profile, scopes), /Explicit login required/);
      } else if (operation === "replacement-healthy") {
        assert.equal((await provider.login(profile, method, scopes)).accountId, "account-B");
        assert.equal(provider.storage, "os-protected");
        assert.equal((await new MsalProvider().silent(profile, scopes)).accountId, "account-B");
      } else {
        assert.equal((await provider.login(profile, method, scopes)).accountId, "account-B");
        assert.equal(provider.storage, "session-only");
        rejectWrites = false;
        await assert.rejects(new MsalProvider().silent(profile, scopes), /Explicit login required/);
        assert.equal((await provider.silent(profile, scopes)).accountId, "account-B");
      }
      process.exit(0);
    }
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
