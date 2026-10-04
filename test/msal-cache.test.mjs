import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";

for (const method of ["browser", "device-code"]) for (const transition of ["login", "refresh", "unavailable-store"]) test(`real MSAL silent acquisition retains credentials after ${method} ${transition}`, () => {
  const result = spawnSync(process.execPath, ["--experimental-test-module-mocks", "--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import { registerHooks } from "node:module";
    import { mock } from "node:test";
    import { PublicClientApplication, TokenCache } from "@azure/msal-node";
    const method = ${JSON.stringify(method)};
    const transition = ${JSON.stringify(transition)};
    const tenant = "11111111-1111-4111-8111-111111111111";
    const client = "22222222-2222-4222-8222-222222222222";
    const profile = { tenantId: tenant, clientId: client, cloud: "commercial", credentialRef: { key: "33333333-3333-4333-8333-333333333333" } };
    const scopes = ["https://graph.microsoft.com/User.Read"];
    const authority = "https://login.microsoftonline.com/" + tenant;
    mock.timers.enable({ apis: ["Date"], now: 1700000000000 });
    const clients = [];
    const saved = new Map();
    let large = transition !== "refresh";
    let tokenRequests = 0;
    mock.module("keytar", { defaultExport: {
      getPassword: async (_, key) => saved.get(key) ?? null,
      setPassword: async (_, key, value) => {
        assert.ok(Buffer.byteLength(value) <= 2560);
        saved.set(key, value);
      },
      deletePassword: async (_, key) => saved.delete(key),
    } });
    registerHooks({ resolve(specifier, context, nextResolve) {
      if (specifier === "keytar" && transition === "unavailable-store") throw new Error("Module unavailable");
      return nextResolve(specifier, context);
    } });
    const network = {
      sendGetRequestAsync: async url => { throw new Error("Unexpected discovery request: " + url); },
      sendPostRequestAsync: async (url, options) => {
        if (new URL(url).pathname.endsWith("/devicecode")) return { status: 200, headers: {}, body: {
          user_code: "FIXTURE", device_code: "fixture-device-code", verification_uri: "https://microsoft.com/devicelogin", expires_in: 600, interval: 0,
        } };
        assert.equal(new URL(url).pathname, "/" + tenant + "/oauth2/v2.0/token");
        const grant = new URLSearchParams(options.body).get("grant_type");
        assert.ok(["authorization_code", "device_code", "refresh_token"].includes(grant));
        tokenRequests++;
        const claims = { aud: client, tid: tenant, oid: "uid", sub: "uid", preferred_username: "fixture@example.invalid" };
        const idToken = Buffer.from(JSON.stringify({ alg: "none" })).toString("base64url") + "." + Buffer.from(JSON.stringify(claims)).toString("base64url") + ".fixture";
        return { status: 200, headers: {}, body: {
          token_type: "Bearer", scope: scopes.join(" "), expires_in: 3600,
          access_token: large ? "large-fixture-token-" + "x".repeat(3000) : "small-fixture-token",
          refresh_token: "fixture-refresh-token", id_token: idToken,
          client_info: Buffer.from(JSON.stringify({ uid: "uid", utid: tenant })).toString("base64url"),
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
        clients.push(this);
      }
      async acquireTokenInteractive(request) {
        return this.acquireTokenByCode({ scopes: request.scopes, code: "fixture-code", redirectUri: "http://localhost" });
      }
    }
    mock.module("@azure/msal-node", { namedExports: { PublicClientApplication: Application } });
    const { MsalProvider } = await import("./dist/msal-provider.js");
    Object.defineProperty(process, "platform", { value: "win32" });
    const provider = new MsalProvider();
    let authenticated = await provider.login(profile, method, scopes);
    if (transition === "refresh") {
      assert.equal(provider.storage, "os-protected");
      assert.equal(saved.size, 1);
      mock.timers.tick(3601000);
      large = true;
      authenticated = await provider.silent(profile, scopes);
    }
    assert.equal(provider.storage, "session-only");
    assert.equal(saved.size, 0);
    const cache = clients.at(-1).getTokenCache();
    assert.ok(cache instanceof TokenCache);
    assert.equal(cache.persistence, undefined);
    const requestsBeforeSilent = tokenRequests;
    const cached = await provider.silent(profile, scopes);
    assert.equal(cached.token, authenticated.token);
    assert.equal(cached.accountId, authenticated.accountId);
    assert.equal(cached.expiresAt, authenticated.expiresAt);
    assert.equal(tokenRequests, requestsBeforeSilent);
    assert.equal((await provider.silent(profile, scopes)).token, authenticated.token);
    assert.equal(tokenRequests, requestsBeforeSilent);
    await assert.rejects(new MsalProvider().silent(profile, scopes), /Explicit login required/);
  `], { encoding: "utf8", timeout: 10_000 });
  assert.equal(result.status, 0, result.stderr);
});
