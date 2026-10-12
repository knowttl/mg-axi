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
import { resolve } from "node:path";
import { mock } from "node:test";
import { ConfidentialClientApplication } from "@azure/msal-node";
const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
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
const dir = resolve(mkdtempSync(".cache/msal-headless-"));
async function failsClosed(profile, reference, secrets, code = "AUTH_REQUIRED") {
  try {
    await new ApplicationAuth(new MsalApplicationProvider()).credential(profile);
    assert.fail("expected AUTH_REQUIRED for " + reference);
  } catch (error) {
    assert.equal(error.code, code);
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
    import { resolve } from "node:path";
    import { mock } from "node:test";
    import { ConfidentialClientApplication } from "@azure/msal-node";
    const provider = ${JSON.stringify(provider)};
    const mismatch = ${JSON.stringify(mismatch)};
    const tenant = "11111111-1111-4111-8111-111111111111";
    const client = "22222222-2222-4222-8222-222222222222";
    const authority = "https://login.microsoftonline.com/" + tenant;
    const audience = "https://graph.microsoft.com/.default";
    const profile = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", enabledPacks: ["entra"], preview: false, sensitiveAreas: [], allowDeviceCode: false,
      auth: provider, ...(provider === "certificate" ? { thumbprintSha256: "ab".repeat(32), privateKey: { env: "MG_AXI_PRIVATE_KEY" } } : { token: { env: "MG_AXI_TOKEN" } }) };
    const { privateKey } = generateKeyPairSync("rsa", { modulusLength: 2048, privateKeyEncoding: { type: "pkcs8", format: "pem" }, publicKeyEncoding: { type: "spki", format: "pem" } });
    process.env.MG_AXI_PRIVATE_KEY = privateKey;
    mkdirSync(".cache", { recursive: true });
    const dir = resolve(mkdtempSync(".cache/msal-app-"));
    process.env.MG_AXI_TOKEN = "fixture-federated-assertion";
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
const profile = { ...base, auth: "client-secret", clientSecret: { env: "MG_AXI_CLIENT_SECRET" } };
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
const fileProfile = { ...base, auth: "client-secret", clientSecret: { file: secretPath } };
const credential = await new ApplicationAuth(new MsalApplicationProvider()).credential(fileProfile);
assert.equal(credential.token, "opaque-app-token");
assert.equal(seen[0].client_secret, "fixture-file-secret");
delete process.env.MG_AXI_CLIENT_SECRET;
await failsClosed({ ...base, auth: "client-secret", clientSecret: { env: "MG_AXI_CLIENT_SECRET" } }, "MG_AXI_CLIENT_SECRET", ["fixture-file-secret"]);
process.env.MG_AXI_CLIENT_SECRET = "";
await failsClosed({ ...base, auth: "client-secret", clientSecret: { env: "MG_AXI_CLIENT_SECRET" } }, "MG_AXI_CLIENT_SECRET", ["fixture-file-secret"]);
delete process.env.MG_AXI_CLIENT_SECRET;
await failsClosed({ ...base, auth: "client-secret", clientSecret: { file: dir + "/missing" } }, dir + "/missing", ["fixture-file-secret"]);
if (posix) {
  const openPath = dir + "/open-secret";
  writeFileSync(openPath, "fixture-file-secret", { mode: 0o600 });
  chmodSync(openPath, 0o644);
  await failsClosed({ ...base, auth: "client-secret", clientSecret: { file: openPath } }, openPath, ["fixture-file-secret"], "CONFIG_INSECURE");
  const linkPath = dir + "/linked-secret";
  symlinkSync(secretPath, linkPath);
  await failsClosed({ ...base, auth: "client-secret", clientSecret: { file: linkPath } }, linkPath, ["fixture-file-secret"], "CONFIG_INSECURE");
}
` + footer);
});

test("certificate via privateKey file authenticates with SHA-256", () => {
  runChild(harness + `
const keyPath = dir + "/key.pem";
writeFileSync(keyPath, ${JSON.stringify(privateKey)}, { mode: 0o600 });
const profile = { ...base, auth: "certificate", thumbprintSha256: "ab".repeat(32), privateKey: { file: keyPath } };
const provider = new MsalApplicationProvider();
const credential = await new ApplicationAuth(provider).credential(profile);
assert.equal(credential.token, "opaque-app-token");
assert.equal(provider.storage, "session-only");
assert.equal(seen[0].client_assertion_type, "urn:ietf:params:oauth:client-assertion-type:jwt-bearer");
assert.ok(seen[0].client_assertion);
const header = JSON.parse(Buffer.from(seen[0].client_assertion.split(".")[0], "base64url"));
assert.equal(header["x5t#S256"], Buffer.from("ab".repeat(32), "hex").toString("base64url"));
` + footer);
});

test("certificate via keyEnv handles escaped newlines and rejects passphrases", () => {
  runChild(harness + `
process.env.MG_AXI_PRIVATE_KEY = ${JSON.stringify(escapedKey)};
const profile = { ...base, auth: "certificate", thumbprintSha256: "ab".repeat(32), privateKey: { env: "MG_AXI_PRIVATE_KEY" } };
const credential = await new ApplicationAuth(new MsalApplicationProvider()).credential(profile);
assert.equal(credential.token, "opaque-app-token");
assert.equal(seen[0].client_assertion_type, "urn:ietf:params:oauth:client-assertion-type:jwt-bearer");
const lockedPath = dir + "/locked.pem";
writeFileSync(lockedPath, ${JSON.stringify(encryptedKey)}, { mode: 0o600 });
try {
  await new ApplicationAuth(new MsalApplicationProvider()).credential({ ...base, auth: "certificate", thumbprintSha256: "ab".repeat(32), privateKey: { file: lockedPath } });
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
await failsClosed({ ...base, auth: "certificate", thumbprintSha256: "ab".repeat(32), privateKey: { env: "MG_AXI_PRIVATE_KEY" } }, "MG_AXI_PRIVATE_KEY", secrets);
process.env.MG_AXI_PRIVATE_KEY = "";
await failsClosed({ ...base, auth: "certificate", thumbprintSha256: "ab".repeat(32), privateKey: { env: "MG_AXI_PRIVATE_KEY" } }, "MG_AXI_PRIVATE_KEY", secrets);
delete process.env.MG_AXI_PRIVATE_KEY;
await failsClosed({ ...base, auth: "certificate", thumbprintSha256: "ab".repeat(32), privateKey: { file: dir + "/missing.pem" } }, dir + "/missing.pem", secrets);
if (posix) {
  const openPath = dir + "/open.pem";
  writeFileSync(openPath, ${JSON.stringify(privateKey)}, { mode: 0o600 });
  chmodSync(openPath, 0o644);
  await failsClosed({ ...base, auth: "certificate", thumbprintSha256: "ab".repeat(32), privateKey: { file: openPath } }, openPath, secrets, "CONFIG_INSECURE");
  const linkPath = dir + "/linked.pem";
  symlinkSync(keyPath, linkPath);
  await failsClosed({ ...base, auth: "certificate", thumbprintSha256: "ab".repeat(32), privateKey: { file: linkPath } }, linkPath, secrets, "CONFIG_INSECURE");
}
` + footer);
});

test("federated SecretRef file authenticates and rejects exposed or missing holders", () => {
  runChild(harness + `
const assertionPath = dir + "/assertion";
writeFileSync(assertionPath, "fixture-federated-assertion", { mode: 0o600 });
const profile = { ...base, auth: "federated", token: { file: assertionPath } };
const credential = await new ApplicationAuth(new MsalApplicationProvider()).credential(profile);
assert.equal(credential.token, "opaque-app-token");
assert.equal(seen[0].client_assertion, "fixture-federated-assertion");
if (process.platform !== "win32") {
  chmodSync(assertionPath, 0o400);
  assert.equal((await new ApplicationAuth(new MsalApplicationProvider()).credential(profile)).token, "opaque-app-token");
  chmodSync(assertionPath, 0o644);
  await failsClosed(profile, assertionPath, ["fixture-federated-assertion"], "CONFIG_INSECURE");
  chmodSync(assertionPath, 0o600);
  const linked = dir + "/linked-assertion";
  symlinkSync(assertionPath, linked);
  await failsClosed({ ...profile, token: { file: linked } }, linked, ["fixture-federated-assertion"], "CONFIG_INSECURE");
}
rmSync(assertionPath);
await failsClosed(profile, assertionPath, ["fixture-federated-assertion"]);
` + footer);
});

test("deprecated AZURE_FEDERATED_TOKEN_FILE alias resolves with a stderr notice; MG_AXI_TOKEN wins", () => {
  runChild(harness + `
const warnings = [];
const originalWrite = process.stderr.write.bind(process.stderr);
process.stderr.write = (chunk, ...rest) => { warnings.push(String(chunk)); return originalWrite(chunk, ...rest); };
const legacyPath = dir + "/legacy-assertion";
writeFileSync(legacyPath, "fixture-legacy-assertion\\n", { mode: 0o600 });
delete process.env.MG_AXI_TOKEN;
process.env.AZURE_FEDERATED_TOKEN_FILE = legacyPath;
const profile = { ...base, auth: "federated", token: { env: "MG_AXI_TOKEN" } };
const credential = await new ApplicationAuth(new MsalApplicationProvider()).credential(profile);
assert.equal(credential.token, "opaque-app-token");
assert.equal(seen[0].client_assertion, "fixture-legacy-assertion");
assert.ok(warnings.some(line => line.includes("AZURE_FEDERATED_TOKEN_FILE") && line.includes("MG_AXI_TOKEN")), warnings.join(""));
process.env.MG_AXI_TOKEN = "fixture-canonical-assertion";
warnings.length = 0;
await new ApplicationAuth(new MsalApplicationProvider()).credential(profile);
assert.equal(seen[1].client_assertion, "fixture-canonical-assertion");
assert.equal(warnings.length, 0);
delete process.env.MG_AXI_TOKEN;
delete process.env.AZURE_FEDERATED_TOKEN_FILE;
` + footer);
});
