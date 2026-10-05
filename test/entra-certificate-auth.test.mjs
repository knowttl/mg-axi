import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const certAuthScopes = ["https://graph.microsoft.com/PublicKeyInfrastructure.Read.All"];
const base = "/v1.0/directory/publicKeyInfrastructure/certificateBasedAuthConfigurations";

const longStatus = `upload failed: ${"retry the PKI file upload. ".repeat(25)}`;
const pki1 = {
  id: "pki-contoso-1",
  deletedDateTime: null,
  displayName: "Contoso PKI",
  status: "succeeded",
  statusDetails: longStatus,
  lastModifiedDateTime: "2024-10-16T18:09:56Z",
};
const pki2 = { id: "pki-bare-2" };
const pkis = [pki1, pki2];
const certBlob = `MIIFfixture-${"A".repeat(600)}`;
const ca1 = {
  id: "ca-root-1",
  deletedDateTime: null,
  certificateAuthorityType: "root",
  certificate: certBlob,
  displayName: "Contoso Root CA",
  issuer: "Contoso",
  issuerSubjectKeyIdentifier: "C0E9",
  createdDateTime: "2024-10-25T18:05:28Z",
  expirationDateTime: "2027-08-29T02:05:57Z",
  thumbprint: "C6FA",
  certificateRevocationListUrl: null,
  deltacertificateRevocationListUrl: null,
  isIssuerHintEnabled: true,
};
const ca2 = { id: "ca-bare-2" };
const authorities = [ca1, ca2];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-cert-auth-"));
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(dir, "config.json");
  const profiles = new Profiles();
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  return { dir, previous };
}

function teardownProfiles(state) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function credentialService(mode, calls) {
  return {
    credential: async (...args) => {
      calls.push([mode, ...args]);
      return {
        token: `opaque-fixture-${mode}-token`,
        expiresAt: Date.now() + 3_600_000,
        tenantId: tenant,
        clientId: client,
        ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
      };
    },
  };
}

function transport(handler) {
  const requests = [];
  const send = async request => {
    requests.push(request);
    const response = await handler(request);
    return { headers: {}, body: "", ...response };
  };
  return { requests, send };
}

function json(status, body, headers = {}) {
  return { status, headers, body: JSON.stringify(body) };
}

function certAuthTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === `${base}/$count`) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(pkis.length) };
    }
    if (path === base) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [pki2] });
      return json(200, {
        value: [pki1],
        "@odata.nextLink": `https://graph.microsoft.com${base}?%24skiptoken=page2`,
      });
    }
    const caCount = new RegExp(`^${base}/([^/]+)/certificateAuthorities/\\$count$`).exec(path);
    if (caCount) {
      return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(authorities.length) };
    }
    const caList = new RegExp(`^${base}/([^/]+)/certificateAuthorities$`).exec(path);
    if (caList) return json(200, { value: authorities });
    const caSingle = new RegExp(`^${base}/([^/]+)/certificateAuthorities/([^/]+)$`).exec(path);
    if (caSingle) {
      const found = authorities.find(row => row.id === decodeURIComponent(caSingle[2]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such authority" } });
    }
    const single = new RegExp(`^${base}/([^/]+)$`).exec(path);
    if (single) {
      const found = pkis.find(row => row.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such PKI" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? certAuthTransport();
  const credential = credentialService(mode, calls);
  return {
    requests: fixture.requests,
    calls,
    overrides: {
      transport: fixture.send,
      delegated: mode === "delegated" ? credential : credentialService("delegated", []),
      application: mode === "application" ? credential : credentialService("application", []),
    },
  };
}

function runCertAuthCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-certificate-auth-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, pkis: [pki1, pki2], authorities: [ca1, ca2], denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists PKIs with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "certificate-auth-pki", "list", "--profile", profile,
        "--select", "id,displayName,status,statusDetails"], overrides);
      assert.deepEqual(result.certificateAuthPkis, [
        { id: pki1.id, displayName: "Contoso PKI", status: "succeeded",
          statusDetails: `${longStatus.slice(0, 500)}... (truncated, ${longStatus.length} chars total)` },
        { id: pki2.id },
      ]);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra certificate-auth-pki show --id <pki-id>")));
      assert.ok(result.help.some(hint => hint.includes("not configured")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com${base}?`));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(certAuthScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists PKIs with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "certificate-auth-pki", "list", "--profile", profile,
        "--filter", "displayName eq 'Contoso PKI'"], overrides);
      assert.equal(result.count.returned, 2);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "displayName eq 'Contoso PKI'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped PKI list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "certificate-auth-pki", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.certificateAuthPkis.map(row => row.id), [pki1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "certificate-auth-pki", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.certificateAuthPkis.map(row => row.id), [pki2.id]);
      assert.deepEqual(second.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one PKI with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "certificate-auth-pki", "show", "--id", pki1.id, "--profile", profile,
        "--select", "id,deletedDateTime,displayName,status,lastModifiedDateTime"], overrides);
      assert.deepEqual(result.certificateAuthPki, {
        id: pki1.id, deletedDateTime: null, displayName: "Contoso PKI", status: "succeeded",
        lastModifiedDateTime: "2024-10-16T18:09:56Z",
      });
      assert.ok(result.help.some(hint => hint.includes("entra certificate-authority list --pki")));
      const missing = await executeArgv(["entra", "certificate-auth-pki", "show", "--id", pki2.id, "--profile", profile], overrides);
      assert.deepEqual(missing.certificateAuthPki, { id: pki2.id });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts PKIs as one scalar with optional server-side filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "certificate-auth-pki", "count", "--profile", profile], overrides);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, `https://graph.microsoft.com${base}/$count`);
      const filtered = await executeArgv(["entra", "certificate-auth-pki", "count", "--profile", profile,
        "--filter", "displayName eq 'Contoso PKI'"], overrides);
      assert.deepEqual(filtered.count, { returned: 2, complete: true });
      assert.ok(filtered.help.some(hint => hint.includes("--filter")));
      assert.equal(new URL(requests[1].url).searchParams.get("$filter"), "displayName eq 'Contoso PKI'");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists authorities without the certificate blob by default`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "certificate-authority", "list", "--pki", pki1.id, "--profile", profile], overrides);
      assert.deepEqual(result.certificateAuthorities, [
        { id: ca1.id, displayName: "Contoso Root CA", certificateAuthorityType: "root", expirationDateTime: "2027-08-29T02:05:57Z" },
        { id: ca2.id },
      ]);
      assert.ok(!("certificate" in result.certificateAuthorities[0]));
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.ok(result.help.some(hint => hint.includes(`entra certificate-authority show --pki ${pki1.id} --id <authority-id>`)));
      assert.ok(result.help.some(hint => hint.includes("--select certificate")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com${base}/${pki1.id}/certificateAuthorities?`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one authority with the blob only on explicit select`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "certificate-authority", "show",
        "--pki", pki1.id, "--id", ca1.id, "--profile", profile], overrides);
      assert.ok(!("certificate" in result.certificateAuthority));
      assert.equal(result.certificateAuthority.thumbprint, "C6FA");
      const blob = await executeArgv(["entra", "certificate-authority", "show",
        "--pki", pki1.id, "--id", ca1.id, "--profile", profile, "--select", "id,certificate"], overrides);
      assert.equal(blob.certificateAuthority.id, ca1.id);
      assert.match(blob.certificateAuthority.certificate, /truncated, \d+ chars total/);
      assert.ok(blob.help.some(hint => hint.includes("--full")));
      const full = await executeArgv(["entra", "certificate-authority", "show",
        "--pki", pki1.id, "--id", ca1.id, "--profile", profile, "--select", "id,certificate", "--full"], overrides);
      assert.equal(full.certificateAuthority.certificate, certBlob);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts authorities of one PKI as one scalar`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "certificate-authority", "count", "--pki", pki1.id, "--profile", profile], overrides);
      assert.deepEqual(result.count, { returned: 2, complete: true });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, `https://graph.microsoft.com${base}/${pki1.id}/certificateAuthorities/$count`);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown PKI and authority ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "show", "--id", "pki-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "certificate-authority", "show", "--pki", pki1.id, "--id", "ca-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied certificate-auth reads surface scope, roles and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("PublicKeyInfrastructure.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Privileged Authentication Administrator or Authentication Administrator")));
        assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "certificate-authority", "list", "--pki", pki1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("PublicKeyInfrastructure.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "certificate-authority", "count", "--pki", pki1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("PublicKeyInfrastructure.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts certificate-auth reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runCertAuthCli(["entra", "certificate-auth-pki", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.certificateAuthPkis.map(row => row.id), [pki1.id, pki2.id]);
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runCertAuthCli(["entra", "certificate-auth-pki", "show", "--id", pki1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.equal(decode(shown.stdout).certificateAuthPki.displayName, "Contoso PKI");

      const counted = runCertAuthCli(["entra", "certificate-auth-pki", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).count, { returned: 2, complete: true });

      const authoritiesListed = runCertAuthCli(["entra", "certificate-authority", "list", "--pki", pki1.id, "--profile", profile], state, mode);
      assert.equal(authoritiesListed.status, 0, authoritiesListed.stdout);
      const authoritiesOut = decode(authoritiesListed.stdout);
      assert.deepEqual(authoritiesOut.certificateAuthorities.map(row => row.id), [ca1.id, ca2.id]);
      assert.ok(!("certificate" in authoritiesOut.certificateAuthorities[0]));

      const authorityShown = runCertAuthCli(["entra", "certificate-authority", "show", "--pki", pki1.id, "--id", ca1.id, "--profile", profile], state, mode);
      assert.equal(authorityShown.status, 0, authorityShown.stdout);
      assert.equal(decode(authorityShown.stdout).certificateAuthority.thumbprint, "C6FA");

      const authoritiesCounted = runCertAuthCli(["entra", "certificate-authority", "count", "--pki", pki1.id, "--profile", profile], state, mode);
      assert.equal(authoritiesCounted.status, 0, authoritiesCounted.stdout);
      assert.deepEqual(decode(authoritiesCounted.stdout).count, { returned: 2, complete: true });
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied certificate-auth reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runCertAuthCli(["entra", "certificate-auth-pki", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.certificateAuthPkis, undefined);
      assert.ok(output.help.some(hint => hint.includes("PublicKeyInfrastructure.Read.All")));
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["certificate-auth-pki", "list"], ["--limit", "1"]],
    [["certificate-auth-pki", "show"], ["--id", pki1.id]],
    [["certificate-auth-pki", "count"], []],
    [["certificate-authority", "list"], ["--pki", pki1.id]],
    [["certificate-authority", "show"], ["--pki", pki1.id, "--id", ca1.id]],
    [["certificate-authority", "count"], ["--pki", pki1.id]],
  ]) {
    for (const preview of [false, true]) {
      test(`${mode} ${command.join(" ")} refuses beta before credentials with preview=${preview}`, async () => {
        const state = setupProfiles();
        try {
          const path = join(state.dir, "config.json");
          const config = JSON.parse(readFileSync(path, "utf8"));
          config.profiles[profile].preview = preview;
          writeFileSync(path, JSON.stringify(config));
          const { calls, requests, overrides } = overridesFor(mode);
          await assert.rejects(
            executeArgv(["entra", ...command, ...args, "--profile", profile, "--api-version", "beta"], overrides),
            { code: "VALIDATION_ERROR", message: "Certificate-auth reads support v1.0 only; beta needs its own review" },
          );
          assert.equal(calls.length, 0);
          assert.equal(requests.length, 0);
        } finally {
          teardownProfiles(state);
        }
      });
    }
  }

  for (const id of ["delta()", "DELTA()"]) {
    test(`${mode} certificate-auth reads reject function-style identifier ${id} before credentials`, async () => {
      const state = setupProfiles();
      try {
        const { requests, calls, overrides } = overridesFor(mode);
        await assert.rejects(
          executeArgv(["entra", "certificate-auth-pki", "show", "--id", id, "--profile", profile], overrides),
          { code: "VALIDATION_ERROR" },
        );
        await assert.rejects(
          executeArgv(["entra", "certificate-authority", "list", "--pki", id, "--profile", profile], overrides),
          { code: "VALIDATION_ERROR" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }
}

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "certificate-auth-pki", "list", "--profile", "batch", "--scopes", certAuthScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "certificate-authority", "count", "--pki", pki1.id, "--profile", "batch", "--scopes", certAuthScopes[0]], overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("unknown properties and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "list", "--profile", "soc", "--select", "id,owner"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "certificate-authority", "list", "--pki", pki1.id, "--profile", "soc", "--fields", "thumbprint"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "certificate-authority", "show", "--pki", pki1.id, "--id", ca1.id, "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("malformed certificate-auth count bodies fail as unknown, not zero", async () => {
  const state = setupProfiles();
  try {
    for (const body of ["{}", "-1", "2.5", "\"3\""]) {
      const malformed = transport(() => ({ status: 200, headers: {}, body }));
      const scoped = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "count", "--profile", "soc"], scoped.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /malformed certificate-auth PKI count|non-numeric success body|non-JSON success body/.test(error.message);
      });
    }
  } finally {
    teardownProfiles(state);
  }
});

test("certificate-auth read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "certificate-authority", "list"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "certificate-authority", "list", "--pki="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "certificate-authority", "show", "--pki", pki1.id]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "count", "--limit", "5"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "count", "--select", "id"]), /unknown flag --select/);
  await assert.rejects(executeArgv(["entra", "certificate-authority", "count", "--cursor", "x"]), /unknown flag --cursor/);
  await assert.rejects(executeArgv(["entra", "certificate-auth-pki", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
});
