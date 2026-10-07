import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { generateKeyPairSync } from "node:crypto";
import { test } from "node:test";

const keyEncoding = { type: "pkcs8", format: "pem" };
const { privateKey } = generateKeyPairSync("rsa", { modulusLength: 2048, privateKeyEncoding: keyEncoding, publicKeyEncoding: { type: "spki", format: "pem" } });
const escapedKey = privateKey.replaceAll("\n", "\\n");
const encryptedKey = generateKeyPairSync("rsa", { modulusLength: 2048, privateKeyEncoding: { ...keyEncoding, cipher: "aes-256-cbc", passphrase: "fixture-passphrase" }, publicKeyEncoding: { type: "spki", format: "pem" } }).privateKey;
assert.match(encryptedKey, /ENCRYPTED PRIVATE KEY/);

// Shared offline MSAL harness: mocked network records token-request bodies.
const harness = `
import assert from "node:assert/strict";
import { mkdirSync, mkdtempSync, writeFileSync, rmSync, chmodSync, symlinkSync } from "node:fs";
import { mock } from "node:test";
import { ConfidentialClientApplication } from "@azure/msal-node";
const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const key = "33333333-3333-4333-8333-333333333333";
const authority = "https://login.microsoftonline.com/" + tenant;
const audience = "https://graph.microsoft.com/.default";
const base = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false };
const seen = [];
const network = {
  sendGetRequestAsync: async () => { throw new Error("Unexpected discovery request"); },
  sendPostRequestAsync: async (url, options) => {
    const endpoint = new URL(url);
    assert.equal(endpoint.origin + endpoint.pathname, authority + "/oauth2/v2.0/token");
    const body = new URLSearchParams(options.body);
    assert.equal(body.get("grant_type"), "client_credentials");
    assert.equal(body.get("client_id"), client);
    assert.equal(body.get("scope"), audience);
    seen.push(Object.fromEntries(body.entries()));
    return { status: 200, headers: {}, body: { token_type: "Bearer", access_token: "opaque-app-token", expires_in: 3600 } };
  },
};
class Application extends ConfidentialClientApplication {
  constructor(config) {
    super({ ...config, auth: { ...config.auth, knownAuthorities: ["login.microsoftonline.com"], authorityMetadata: JSON.stringify({
      authorization_endpoint: authority + "/oauth2/v2.0/authorize",
      token_endpoint: authority + "/oauth2/v2.0/token",
      end_session_endpoint: authority + "/oauth2/v2.0/logout",
      issuer: authority + "/v2.0", jwks_uri: authority + "/discovery/v2.0/keys" }) },
      system: { ...config.system, networkClient: network } });
  }
}
mock.module("@azure/msal-node", { namedExports: { ConfidentialClientApplication: Application } });
const { MsalApplicationProvider } = await import("./dist/msal-app-provider.js");
const { ApplicationAuth } = await import("./dist/app-auth.js");
mkdirSync(".cache", { recursive: true });
const dir = mkdtempSync(".cache/msal-headless-");
async function failsClosed(profile, reference, secrets) {
  try {
    await new ApplicationAuth(new MsalApplicationProvider()).credential(profile);
    assert.fail("expected AUTH_REQUIRED for " + reference);
  } catch (error) {
    assert.equal(error.code, "AUTH_REQUIRED");
    assert.ok(error.message.includes(reference), error.message);
    const text = error.message + JSON.stringify(error);
    for (const secret of secrets) assert.ok(!text.includes(secret), "leaked holder content for " + reference);
  }
}
try {
`;
const footer = `
} finally { rmSync(dir, { recursive: true, force: true }); }
`;
function runChild(script) {
  const result = spawnSync(process.execPath, ["--disable-warning=ExperimentalWarning", "--experimental-test-module-mocks", "--input-type=module", "-e", script], { encoding: "utf8", timeout: 20_000 });
  assert.equal(result.status, 0, result.stderr);
}

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

test("client-secret via env authenticates and rereads at acquire time", () => {
  runChild(harness + `
process.env.MG_AXI_CLIENT_SECRET = "fixture-secret-first";
const profile = { ...base, credentialRef: { provider: "client-secret", key, secretEnv: "MG_AXI_CLIENT_SECRET" } };
const provider = new MsalApplicationProvider();
const first = await new ApplicationAuth(provider).credential(profile);
assert.equal(first.token, "opaque-app-token");
assert.equal(provider.storage, "session-only");
process.env.MG_AXI_CLIENT_SECRET = "fixture-secret-second";
const second = await new ApplicationAuth(new MsalApplicationProvider()).credential(profile);
assert.equal(second.token, "opaque-app-token");
assert.deepEqual(seen.map(entry => entry.client_secret), ["fixture-secret-first", "fixture-secret-second"]);
assert.ok(seen.every(entry => entry.client_assertion_type === undefined));
delete process.env.MG_AXI_CLIENT_SECRET;
` + footer);
});

test("client-secret via file authenticates; missing and exposed holders fail closed", () => {
  runChild(harness + `
const posix = process.platform !== "win32";
const secretPath = dir + "/secret";
writeFileSync(secretPath, "fixture-file-secret\\n", { mode: 0o600 });
chmodSync(secretPath, 0o600);
const fileProfile = { ...base, credentialRef: { provider: "client-secret", key, secretFile: secretPath } };
const credential = await new ApplicationAuth(new MsalApplicationProvider()).credential(fileProfile);
assert.equal(credential.token, "opaque-app-token");
assert.equal(seen[0].client_secret, "fixture-file-secret");
delete process.env.MG_AXI_CLIENT_SECRET;
await failsClosed({ ...base, credentialRef: { provider: "client-secret", key, secretEnv: "MG_AXI_CLIENT_SECRET" } }, "MG_AXI_CLIENT_SECRET", ["fixture-file-secret"]);
process.env.MG_AXI_CLIENT_SECRET = "";
await failsClosed({ ...base, credentialRef: { provider: "client-secret", key, secretEnv: "MG_AXI_CLIENT_SECRET" } }, "MG_AXI_CLIENT_SECRET", ["fixture-file-secret"]);
delete process.env.MG_AXI_CLIENT_SECRET;
await failsClosed({ ...base, credentialRef: { provider: "client-secret", key, secretFile: dir + "/missing" } }, dir + "/missing", ["fixture-file-secret"]);
if (posix) {
  const openPath = dir + "/open-secret";
  writeFileSync(openPath, "fixture-file-secret", { mode: 0o600 });
  chmodSync(openPath, 0o644);
  await failsClosed({ ...base, credentialRef: { provider: "client-secret", key, secretFile: openPath } }, openPath, ["fixture-file-secret"]);
  const linkPath = dir + "/linked-secret";
  symlinkSync(secretPath, linkPath);
  await failsClosed({ ...base, credentialRef: { provider: "client-secret", key, secretFile: linkPath } }, linkPath, ["fixture-file-secret"]);
}
` + footer);
});

test("certificate via keyFile authenticates without keytar", () => {
  runChild(harness + `
mock.module("keytar", { defaultExport: { getPassword: async () => { throw new Error("keytar unavailable"); }, setPassword: async () => {}, deletePassword: async () => {} } });
const keyPath = dir + "/key.pem";
writeFileSync(keyPath, ${JSON.stringify(privateKey)}, { mode: 0o600 });
const profile = { ...base, credentialRef: { provider: "certificate", key, thumbprint: "ab".repeat(20), keyFile: keyPath } };
const provider = new MsalApplicationProvider();
const credential = await new ApplicationAuth(provider).credential(profile);
assert.equal(credential.token, "opaque-app-token");
assert.equal(provider.storage, "session-only");
assert.equal(seen[0].client_assertion_type, "urn:ietf:params:oauth:client-assertion-type:jwt-bearer");
assert.ok(seen[0].client_assertion);
` + footer);
});

test("certificate via keyEnv handles escaped newlines and rejects passphrases", () => {
  runChild(harness + `
process.env.MG_AXI_PRIVATE_KEY = ${JSON.stringify(escapedKey)};
const profile = { ...base, credentialRef: { provider: "certificate", key, thumbprint: "ab".repeat(20), keyEnv: "MG_AXI_PRIVATE_KEY" } };
const credential = await new ApplicationAuth(new MsalApplicationProvider()).credential(profile);
assert.equal(credential.token, "opaque-app-token");
assert.equal(seen[0].client_assertion_type, "urn:ietf:params:oauth:client-assertion-type:jwt-bearer");
const lockedPath = dir + "/locked.pem";
writeFileSync(lockedPath, ${JSON.stringify(encryptedKey)}, { mode: 0o600 });
try {
  await new ApplicationAuth(new MsalApplicationProvider()).credential({ ...base, credentialRef: { provider: "certificate", key, thumbprint: "ab".repeat(20), keyFile: lockedPath } });
  assert.fail("expected AUTH_REQUIRED for passphrase-protected key");
} catch (error) {
  assert.equal(error.code, "AUTH_REQUIRED");
  assert.match(error.message, /passphrase/);
  assert.ok(!JSON.stringify(error).includes("ENCRYPTED PRIVATE KEY"));
}
delete process.env.MG_AXI_PRIVATE_KEY;
` + footer);
});

test("certificate key holders fail closed without leaking key text", () => {
  runChild(harness + `
const posix = process.platform !== "win32";
const keyPath = dir + "/key.pem";
writeFileSync(keyPath, ${JSON.stringify(privateKey)}, { mode: 0o600 });
const secrets = ["PRIVATE KEY"];
delete process.env.MG_AXI_PRIVATE_KEY;
await failsClosed({ ...base, credentialRef: { provider: "certificate", key, thumbprint: "ab".repeat(20), keyEnv: "MG_AXI_PRIVATE_KEY" } }, "MG_AXI_PRIVATE_KEY", secrets);
process.env.MG_AXI_PRIVATE_KEY = "";
await failsClosed({ ...base, credentialRef: { provider: "certificate", key, thumbprint: "ab".repeat(20), keyEnv: "MG_AXI_PRIVATE_KEY" } }, "MG_AXI_PRIVATE_KEY", secrets);
delete process.env.MG_AXI_PRIVATE_KEY;
await failsClosed({ ...base, credentialRef: { provider: "certificate", key, thumbprint: "ab".repeat(20), keyFile: dir + "/missing.pem" } }, dir + "/missing.pem", secrets);
if (posix) {
  const openPath = dir + "/open.pem";
  writeFileSync(openPath, ${JSON.stringify(privateKey)}, { mode: 0o600 });
  chmodSync(openPath, 0o644);
  await failsClosed({ ...base, credentialRef: { provider: "certificate", key, thumbprint: "ab".repeat(20), keyFile: openPath } }, openPath, secrets);
  const linkPath = dir + "/linked.pem";
  symlinkSync(keyPath, linkPath);
  await failsClosed({ ...base, credentialRef: { provider: "certificate", key, thumbprint: "ab".repeat(20), keyFile: linkPath } }, linkPath, secrets);
}
` + footer);
});

test("federated custom token-file env authenticates and fails closed", () => {
  runChild(harness + `
const assertionPath = dir + "/assertion";
writeFileSync(assertionPath, "fixture-federated-assertion");
process.env.MG_AXI_ASSERTION_FILE = assertionPath;
const profile = { ...base, credentialRef: { provider: "federated", key, tokenFileEnv: "MG_AXI_ASSERTION_FILE" } };
const credential = await new ApplicationAuth(new MsalApplicationProvider()).credential(profile);
assert.equal(credential.token, "opaque-app-token");
assert.equal(seen[0].client_assertion, "fixture-federated-assertion");
delete process.env.MG_AXI_ASSERTION_FILE;
await failsClosed(profile, "MG_AXI_ASSERTION_FILE", ["fixture-federated-assertion"]);
process.env.MG_AXI_ASSERTION_FILE = "";
await failsClosed(profile, "MG_AXI_ASSERTION_FILE", ["fixture-federated-assertion"]);
delete process.env.MG_AXI_ASSERTION_FILE;
` + footer);
});
