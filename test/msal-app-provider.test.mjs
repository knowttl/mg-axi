import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";

for (const provider of ["certificate", "federated"]) for (const mismatch of [false, true]) test(`${provider}: ${mismatch ? "foreign authority is rejected" : "app-only response authenticates without an ID token"}`, () => {
  const result = spawnSync(process.execPath, ["--disable-warning=ExperimentalWarning", "--experimental-test-module-mocks", "--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import { generateKeyPairSync } from "node:crypto";
    import { mkdirSync, mkdtempSync, writeFileSync, rmSync } from "node:fs";
    import { mock } from "node:test";
    import { ConfidentialClientApplication } from "@azure/msal-node";
    const provider = ${JSON.stringify(provider)};
    const mismatch = ${JSON.stringify(mismatch)};
    const tenant = "11111111-1111-4111-8111-111111111111";
    const client = "22222222-2222-4222-8222-222222222222";
    const key = "33333333-3333-4333-8333-333333333333";
    const authority = "https://login.microsoftonline.com/" + tenant;
    const audience = "https://graph.microsoft.com/.default";
    const profile = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false,
      credentialRef: provider === "certificate" ? { provider, key, thumbprint: "ab".repeat(20) } : { provider, key } };
    const { privateKey } = generateKeyPairSync("rsa", { modulusLength: 2048, privateKeyEncoding: { type: "pkcs8", format: "pem" }, publicKeyEncoding: { type: "spki", format: "pem" } });
    mock.module("keytar", { defaultExport: { getPassword: async (service, reference) => {
      assert.equal(service, "mg-axi");
      assert.equal(reference, key);
      return privateKey;
    } } });
    mkdirSync(".cache", { recursive: true });
    const dir = mkdtempSync(".cache/msal-app-");
    process.env.AZURE_FEDERATED_TOKEN_FILE = dir + "/assertion";
    writeFileSync(process.env.AZURE_FEDERATED_TOKEN_FILE, "fixture-federated-assertion");
    const requests = [];
    const network = {
      sendGetRequestAsync: async () => { throw new Error("Unexpected discovery request"); },
      sendPostRequestAsync: async (url, options) => {
        const endpoint = new URL(url);
        assert.equal(endpoint.origin + endpoint.pathname, authority + "/oauth2/v2.0/token");
        const body = new URLSearchParams(options.body);
        assert.equal(body.get("grant_type"), "client_credentials");
        assert.equal(body.get("client_id"), client);
        assert.equal(body.get("scope"), audience);
        assert.equal(body.get("client_assertion_type"), "urn:ietf:params:oauth:client-assertion-type:jwt-bearer");
        assert.ok(body.get("client_assertion"));
        if (provider === "federated") assert.equal(body.get("client_assertion"), "fixture-federated-assertion");
        requests.push(endpoint.origin + endpoint.pathname);
        return { status: 200, headers: {}, body: { token_type: "Bearer", access_token: "opaque-app-token", expires_in: 3600 } };
      },
    };
    class Application extends ConfidentialClientApplication {
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
      async acquireTokenByClientCredential(request) {
        const response = await super.acquireTokenByClientCredential(request);
        assert.equal(response.tenantId, "");
        assert.equal(response.idToken, "");
        assert.equal(response.authority, authority + "/");
        return mismatch ? { ...response, authority: "https://login.microsoftonline.com/" + client + "/" } : response;
      }
    }
    mock.module("@azure/msal-node", { namedExports: { ConfidentialClientApplication: Application } });
    const { MsalApplicationProvider } = await import("./dist/msal-app-provider.js");
    const { ApplicationAuth } = await import("./dist/app-auth.js");
    try {
      const auth = new ApplicationAuth(new MsalApplicationProvider());
      if (mismatch) await assert.rejects(auth.credential(profile), { code: "AUTH_REQUIRED" });
      else {
        const credential = await auth.credential(profile);
        assert.equal(credential.token, "opaque-app-token");
        assert.equal(credential.tenantId, tenant);
        assert.equal(credential.clientId, client);
        assert.ok(credential.expiresAt > Date.now() + 60_000);
        assert.deepEqual(await auth.credential(profile), credential);
      }
      assert.deepEqual(requests, [authority + "/oauth2/v2.0/token"]);
    } finally { rmSync(dir, { recursive: true, force: true }); }
  `], { encoding: "utf8", timeout: 10_000 });
  assert.equal(result.status, 0, result.stderr);
});
