import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";

const scenarios = [
  { name: "protected account replacement wipes the prior cache", replace: true, storage: "protected-file" },
  { name: "failed protected wipe prevents account replacement", replace: true, failWipe: true, error: "LOGIN_FAILED" },
  { name: "Windows oversized login persists credentials", platform: "win32", large: true, storage: "protected-file" },
  { name: "Windows oversized replacement persists only the new account", platform: "win32", large: true, replace: true, storage: "protected-file" },
  { name: "Windows oversized refresh keeps refreshed tokens in memory", platform: "win32", large: true, refresh: true, storage: "protected-file" },
];

for (const method of ["browser", "device-code"]) for (const scenario of scenarios) test(`real MSAL: ${method} ${scenario.name}`, () => {
  const result = spawnSync(process.execPath, ["--experimental-test-module-mocks", "--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import { mkdtempSync, readFileSync, rmSync } from "node:fs";
    import { tmpdir } from "node:os";
    import { join } from "node:path";
    import { mock } from "node:test";
    import { PublicClientApplication } from "@azure/msal-node";
    import { Profiles } from "./dist/profiles.js";
    import { Credentials } from "./dist/credentials.js";
    import { PrivateFiles } from "./dist/private-file.js";
    const method = ${JSON.stringify(method)};
    const scenario = ${JSON.stringify(scenario)};
    const tenant = "11111111-1111-4111-8111-111111111111";
    const client = "22222222-2222-4222-8222-222222222222";
    const dir = mkdtempSync(join(tmpdir(), "mg-msal-cache-"));
    const profiles = new Profiles(join(dir, "config.json"));
    profiles.create("fixture", tenant, client, "commercial", true);
    const profile = profiles.resolve().profile;
    const store = new Credentials(join(dir, "credentials.json"), new PrivateFiles(scenario.platform ?? process.platform, () => {}));
    const scopes = ["https://graph.microsoft.com/User.Read"];
    const authority = "https://login.microsoftonline.com/" + tenant;
    // Keep the initial synthetic refresh cache below the Windows byte limit.
    const firstAccount = scenario.refresh ? "uid" : "44444444-4444-4444-8444-444444444444";
    const secondAccount = "55555555-5555-4555-8555-555555555555";
    mock.timers.enable({ apis: ["Date"], now: 1700000000000 });
    let large = false;
    let failWipe = false;
    let account = firstAccount;
    let tokenRequests = 0;
    const remove = store.remove.bind(store);
    mock.method(store, "remove", profile => { if (failWipe) throw new Error("Fixture wipe failed"); return remove(profile); });
    const network = {
      sendGetRequestAsync: async () => { throw new Error("Unexpected discovery request"); },
      sendPostRequestAsync: async (url, options) => {
        if (new URL(url).pathname.endsWith("/devicecode")) return { status: 200, headers: {}, body: {
          user_code: "FIXTURE", device_code: "fixture-device-code", verification_uri: "https://microsoft.com/devicelogin", expires_in: 600, interval: 0,
        } };
        assert.equal(new URL(url).pathname, "/" + tenant + "/oauth2/v2.0/token");
        const params = new URLSearchParams(options.body);
        const grant = params.get("grant_type");
        assert.ok(["authorization_code", "device_code", "refresh_token"].includes(grant));
        tokenRequests++;
        const responseAccount = grant === "refresh_token" ? params.get("refresh_token").slice("fixture-refresh-token-".length).split(":")[0] : account;
        const claims = { aud: client, tid: tenant, oid: responseAccount, sub: responseAccount, preferred_username: "fixture@example.invalid" };
        const idToken = Buffer.from(JSON.stringify({ alg: "none" })).toString("base64url") + "." + Buffer.from(JSON.stringify(claims)).toString("base64url") + ".fixture";
        return { status: 200, headers: {}, body: {
          token_type: "Bearer", scope: scopes.join(" "), expires_in: 3600,
          access_token: large ? "large-fixture-token-" + "x".repeat(3000) : "small-fixture-token",
          refresh_token: "fixture-refresh-token-" + responseAccount + ":" + (large ? "x".repeat(4000) : ""), id_token: idToken,
          client_info: Buffer.from(JSON.stringify({ uid: responseAccount, utid: tenant })).toString("base64url"),
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
    const provider = new MsalProvider(store);
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
      assert.equal(provider.storage, "protected-file");
      assert.ok(store.read(profile));
      if (scenario.failWipe) {
        assert.equal(tokenRequests, requestsBeforeAction);
        failWipe = false;
        assert.equal((await new MsalProvider(store).silent(profile, scopes)).accountId, firstAccount + "." + tenant);
      }
    } else {
      const result = await action();
      if (!scenario.refresh) assert.equal(result.storage, scenario.storage);
      assert.equal(provider.storage, scenario.storage);
      const requestsBeforeSilent = tokenRequests;
      const cached = await provider.silent(profile, scopes);
      assert.equal(cached.accountId, account + "." + tenant);
      assert.equal(cached.token, large ? "large-fixture-token-" + "x".repeat(3000) : "small-fixture-token");
      assert.equal(tokenRequests, requestsBeforeSilent + (scenario.refresh ? 0 : 1));
      assert.equal((await provider.silent(profile, scopes)).accountId, cached.accountId);
      assert.equal(tokenRequests, requestsBeforeSilent + (scenario.refresh ? 0 : 1));
      assert.equal((await new MsalProvider(store).silent(profile, scopes)).accountId, cached.accountId);
      const persisted = readFileSync(store.path, "utf8");
      assert.doesNotMatch(persisted, /small-fixture-token|large-fixture-token/);
      if (scenario.large && !scenario.refresh) assert.ok(Buffer.byteLength(persisted) > 2560);
      const cache = JSON.parse(store.read(profile));
      assert.equal(Object.keys(cache.Account).length, 1);
    }
    rmSync(dir, { recursive: true, force: true });
  `], { encoding: "utf8", timeout: 10_000 });
  assert.equal(result.status, 0, result.stderr);
});
