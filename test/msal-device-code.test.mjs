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
    import { mkdtempSync, rmSync } from "node:fs";
    import { tmpdir } from "node:os";
    import { join } from "node:path";
    import { mock } from "node:test";
    import { PublicClientApplication } from "@azure/msal-node";
    const scenario = ${JSON.stringify(scenario)};
    const tenant = "11111111-1111-4111-8111-111111111111";
    const { Profiles } = await import("./dist/profiles.js");
    const { Credentials } = await import("./dist/credentials.js");
    const dir = mkdtempSync(join(tmpdir(), "mg-device-code-"));
    const profiles = new Profiles(join(dir, "config.json"));
    profiles.create("fixture", tenant, "22222222-2222-4222-8222-222222222222", "commercial", true);
    const profile = profiles.resolve().profile;
    const authority = "https://login.microsoftonline.com/" + tenant;
    const requests = [];
    const challenges = [];
    mock.timers.enable({ apis: ["Date"], now: 1700000000000 });
    mock.method(process.stderr, "write", chunk => { challenges.push(String(chunk)); return true; });
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
    const auth = new DelegatedAuth(new MsalProvider(new Credentials(join(dir, "credentials.json"))), () => Date.now());
    await assert.rejects(auth.login(profile, "device-code", ["https://graph.microsoft.com/User.Read"]), error => error.code === "LOGIN_FAILED" && !JSON.stringify(error).includes("fixture-private-diagnostic"));
    assert.deepEqual(challenges, []);
    assert.deepEqual(requests, ["/" + tenant + "/oauth2/v2.0/devicecode"]);
    rmSync(dir, { recursive: true, force: true });
  `], { encoding: "utf8", timeout: 10_000 });
  assert.equal(result.status, 0, result.stderr);
});
