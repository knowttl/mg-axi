import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";

const scenarios = [
  { name: "protected account replacement wipes the prior cache", replace: true, storage: "os-protected" },
  { name: "session account replacement wipes the prior cache", replace: true, unavailable: true, storage: "session-only" },
  { name: "failed protected wipe prevents account replacement", replace: true, failWipe: true, error: "LOGIN_FAILED" },
  { name: "unavailable store retains session credentials", unavailable: true, storage: "session-only" },
  { name: "Windows oversized login retains session credentials", platform: "win32", large: true, storage: "session-only" },
  { name: "Windows oversized replacement retains only the new account in memory", platform: "win32", large: true, replace: true, storage: "session-only" },
  { name: "Windows oversized refresh retains session credentials", platform: "win32", large: true, refresh: true, storage: "session-only" },
];

for (const method of ["browser", "device-code"]) for (const scenario of scenarios) test(`real MSAL: ${method} ${scenario.name}`, () => {
  const result = spawnSync(process.execPath, ["--experimental-test-module-mocks", "--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import { registerHooks } from "node:module";
    import { mock } from "node:test";
    import { PublicClientApplication } from "@azure/msal-node";
    const method = ${JSON.stringify(method)};
    const scenario = ${JSON.stringify(scenario)};
    const tenant = "11111111-1111-4111-8111-111111111111";
    const client = "22222222-2222-4222-8222-222222222222";
    const profile = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: true, credentialRef: { provider: "os-or-session", key: "33333333-3333-4333-8333-333333333333" } };
    const scopes = ["https://graph.microsoft.com/User.Read"];
    const authority = "https://login.microsoftonline.com/" + tenant;
    // Keep the initial synthetic refresh cache below the Windows byte limit.
    const firstAccount = scenario.refresh ? "uid" : "44444444-4444-4444-8444-444444444444";
    const secondAccount = "55555555-5555-4555-8555-555555555555";
    mock.timers.enable({ apis: ["Date"], now: 1700000000000 });
    const saved = new Map();
    let large = false;
    let failWipe = false;
    let account = firstAccount;
    let tokenRequests = 0;
    mock.module("keytar", { defaultExport: {
      getPassword: async (_, key) => saved.get(key) ?? null,
      setPassword: async (_, key, value) => { saved.set(key, value); },
      deletePassword: async (_, key) => {
        if (failWipe) throw new Error("Fixture wipe failed");
        return saved.delete(key);
      },
    } });
    registerHooks({ resolve(specifier, context, nextResolve) {
      if (specifier === "keytar" && scenario.unavailable) throw new Error("Module unavailable");
      return nextResolve(specifier, context);
    } });
    const network = {
      sendGetRequestAsync: async () => { throw new Error("Unexpected discovery request"); },
      sendPostRequestAsync: async (url, options) => {
        if (new URL(url).pathname.endsWith("/devicecode")) return { status: 200, headers: {}, body: {
          user_code: "FIXTURE", device_code: "fixture-device-code", verification_uri: "https://microsoft.com/devicelogin", expires_in: 600, interval: 0,
        } };
        assert.equal(new URL(url).pathname, "/" + tenant + "/oauth2/v2.0/token");
        const grant = new URLSearchParams(options.body).get("grant_type");
        assert.ok(["authorization_code", "device_code", "refresh_token"].includes(grant));
        tokenRequests++;
        const claims = { aud: client, tid: tenant, oid: account, sub: account, preferred_username: "fixture@example.invalid" };
        const idToken = Buffer.from(JSON.stringify({ alg: "none" })).toString("base64url") + "." + Buffer.from(JSON.stringify(claims)).toString("base64url") + ".fixture";
        return { status: 200, headers: {}, body: {
          token_type: "Bearer", scope: scopes.join(" "), expires_in: 3600,
          access_token: large ? "large-fixture-token-" + "x".repeat(3000) : "small-fixture-token",
          refresh_token: "fixture-refresh-token", id_token: idToken,
          client_info: Buffer.from(JSON.stringify({ uid: account, utid: tenant })).toString("base64url"),
        } };
      },
    };
    class Application extends PublicClientApplication {
      constructor(config) {
        super({ ...config, auth: {
          ...config.auth, knownAuthorities: ["login.microsoftonline.com"], authorityMetadata: JSON.stringify({
            authorization_endpoint: authority + "/oauth2/v2.0/authorize",
            token_endpoint: authority + "/oauth2/v2.0/token",
            end_session_endpoint: authority + "/oauth2/v2.0/logout",
            issuer: authority + "/v2.0", jwks_uri: authority + "/discovery/v2.0/keys",
          }),
        }, system: { ...config.system, networkClient: network } });
      }
      async acquireTokenInteractive(request) {
        return this.acquireTokenByCode({ scopes: request.scopes, code: "fixture-code", redirectUri: "http://localhost" });
      }
    }
    mock.module("@azure/msal-node", { namedExports: { PublicClientApplication: Application } });
    const { MsalProvider } = await import("./dist/msal-provider.js");
    const { DelegatedAuth } = await import("./dist/auth.js");
    Object.defineProperty(process, "platform", { value: scenario.platform ?? "linux" });
    const provider = new MsalProvider();
    const auth = new DelegatedAuth(provider, () => Date.now());
    if (scenario.replace || scenario.refresh) await provider.login(profile, method, scopes);
    const requestsBeforeAction = tokenRequests;
    account = scenario.replace ? secondAccount : firstAccount;
    large = !!scenario.large;
    failWipe = !!scenario.failWipe;
    if (scenario.refresh) mock.timers.tick(3601000);
    const action = () => scenario.refresh ? auth.credential(profile, scopes) : auth.login(profile, method, scopes);
    if (scenario.error) {
      await assert.rejects(action(), { code: scenario.error });
      assert.equal(provider.storage, "os-protected");
      assert.equal(saved.size, scenario.failWipe || scenario.refresh ? 1 : 0);
      if (scenario.failWipe) {
        assert.equal(tokenRequests, requestsBeforeAction);
        failWipe = false;
        assert.equal((await new MsalProvider().silent(profile, scopes)).accountId, firstAccount + "." + tenant);
      }
    } else {
      const result = await action();
      if (!scenario.refresh) assert.equal(result.storage, scenario.storage);
      assert.equal(provider.storage, scenario.storage);
      const requestsBeforeSilent = tokenRequests;
      const cached = await provider.silent(profile, scopes);
      assert.equal(cached.accountId, account + "." + tenant);
      assert.equal(cached.token, large ? "large-fixture-token-" + "x".repeat(3000) : "small-fixture-token");
      assert.equal(tokenRequests, requestsBeforeSilent);
      assert.equal((await provider.silent(profile, scopes)).accountId, cached.accountId);
      assert.equal(tokenRequests, requestsBeforeSilent);
      if (scenario.storage === "session-only") {
        assert.equal(saved.size, 0);
        await assert.rejects(new MsalProvider().silent(profile, scopes), /Explicit login required/);
        mock.timers.tick(3601000);
        assert.equal((await provider.silent(profile, scopes)).accountId, cached.accountId);
        assert.equal(tokenRequests, requestsBeforeSilent + 1);
        assert.equal(provider.storage, "session-only");
        assert.equal(saved.size, 0);
      }
      else assert.equal((await new MsalProvider().silent(profile, scopes)).accountId, cached.accountId);
    }
  `], { encoding: "utf8", timeout: 10_000 });
  assert.equal(result.status, 0, result.stderr);
});
