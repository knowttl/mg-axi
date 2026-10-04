import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";

for (const scenario of [
  { name: "rejected device-code request", status: 400, body: { error: "invalid_request", error_description: "fixture-private-diagnostic" } },
  { name: "missing user challenge", status: 200, body: { device_code: "fixture-device-code", expires_in: 600 } },
  { name: "missing polling credential", status: 200, body: { user_code: "FIXTURE", expires_in: 600 } },
]) test(`${scenario.name} fails without displaying a challenge or requesting tokens`, () => {
  const result = spawnSync(process.execPath, ["--disable-warning=ExperimentalWarning", "--experimental-test-module-mocks", "--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import { registerHooks } from "node:module";
    import { mock } from "node:test";
    import { PublicClientApplication } from "@azure/msal-node";
    const scenario = ${JSON.stringify(scenario)};
    const tenant = "11111111-1111-4111-8111-111111111111";
    const profile = { mode: "delegated", tenantId: tenant, clientId: "22222222-2222-4222-8222-222222222222", cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: true, credentialRef: { provider: "os-or-session", key: "33333333-3333-4333-8333-333333333333" } };
    const authority = "https://login.microsoftonline.com/" + tenant;
    const requests = [];
    const challenges = [];
    mock.timers.enable({ apis: ["Date"], now: 1700000000000 });
    mock.method(process.stderr, "write", chunk => { challenges.push(String(chunk)); return true; });
    registerHooks({ resolve(specifier, context, nextResolve) {
      if (specifier === "keytar") throw new Error("Fixture store unavailable");
      return nextResolve(specifier, context);
    } });
    const network = {
      sendGetRequestAsync: async () => { throw new Error("Unexpected discovery request"); },
      sendPostRequestAsync: async url => {
        requests.push(new URL(url).pathname);
        assert.equal(new URL(url).pathname, "/" + tenant + "/oauth2/v2.0/devicecode");
        return { status: scenario.status, headers: {}, body: scenario.body };
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
    }
    mock.module("@azure/msal-node", { namedExports: { PublicClientApplication: Application } });
    const { MsalProvider } = await import("./dist/msal-provider.js");
    const { DelegatedAuth } = await import("./dist/auth.js");
    const auth = new DelegatedAuth(new MsalProvider(), () => Date.now());
    await assert.rejects(auth.login(profile, "device-code", ["https://graph.microsoft.com/User.Read"]), error => error.code === "LOGIN_FAILED" && !JSON.stringify(error).includes("fixture-private-diagnostic"));
    assert.deepEqual(challenges, []);
    assert.deepEqual(requests, ["/" + tenant + "/oauth2/v2.0/devicecode"]);
  `], { encoding: "utf8", timeout: 10_000 });
  assert.equal(result.status, 0, result.stderr);
});
