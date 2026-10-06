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
const domainScopes = ["https://graph.microsoft.com/Domain.Read.All"];

const dom1 = {
  id: "contoso.com",
  authenticationType: "Managed",
  availabilityStatus: null,
  isAdminManaged: true,
  isDefault: true,
  isInitial: true,
  isRoot: true,
  isVerified: true,
  supportedServices: ["Email", "OfficeCommunicationsOnline"],
  passwordValidityPeriodInDays: 90,
  passwordNotificationWindowInDays: 14,
  state: null,
};
const dom2 = {
  id: "sub.contoso.com",
  authenticationType: "Managed",
  availabilityStatus: null,
  isAdminManaged: true,
  isDefault: false,
  isInitial: false,
  isRoot: false,
  isVerified: false,
  supportedServices: ["Email"],
  passwordValidityPeriodInDays: null,
  passwordNotificationWindowInDays: null,
  state: null,
};
const dom3 = { id: "fabrikam.com" };
const domains = [dom1, dom2, dom3];

const v1 = {
  "@odata.type": "#microsoft.graph.domainDnsTxtRecord",
  id: "rec-verify-1",
  isOptional: false,
  label: "contoso.com",
  recordType: "Txt",
  supportedService: "Email",
  ttl: 3600,
  text: "MS=ms12345678",
};
const verification = [v1];

const longSpf = `v=spf1 ${"include:spf.example.com ".repeat(40)}~all`;
const s1 = {
  "@odata.type": "#microsoft.graph.domainDnsMxRecord",
  id: "rec-svc-1",
  isOptional: false,
  label: "contoso.com",
  recordType: "Mx",
  supportedService: "Email",
  ttl: 3600,
  mailExchange: "contoso-com.mail.protection.outlook.com",
  preference: 0,
};
const s2 = {
  "@odata.type": "#microsoft.graph.domainDnsTxtRecord",
  id: "rec-svc-2",
  isOptional: false,
  label: "contoso.com",
  recordType: "Txt",
  supportedService: "Email",
  ttl: 3600,
  text: longSpf,
};
const service = [s1, s2];

const t1 = {
  "@odata.type": "#microsoft.graph.domainDnsCnameRecord",
  id: "rec-top-1",
  isOptional: true,
  label: "autodiscover.contoso.com",
  recordType: "CName",
  supportedService: "Email",
  ttl: 3600,
  canonicalName: "autodiscover.outlook.com",
};
const topRecords = [t1];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-domains-"));
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
      calls.push(args);
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

function domainTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/domains") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [dom3] });
      return json(200, {
        value: [dom1, dom2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/domains?%24skiptoken=page2",
      });
    }
    if (path === "/v1.0/domains/contoso.com/verificationDnsRecords") return json(200, { value: verification });
    if (path === "/v1.0/domains/contoso.com/serviceConfigurationRecords") return json(200, { value: service });
    if (path === "/v1.0/domainDnsRecords") return json(200, { value: topRecords });
    const singleDns = /^\/v1\.0\/(?:domains\/[^/]+\/(?:verificationDnsRecords|serviceConfigurationRecords)|domainDnsRecords)\/([^/]+)$/.exec(path);
    if (singleDns) {
      const found = [...verification, ...service, ...topRecords].find(row => row.id === decodeURIComponent(singleDns[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such DNS record" } });
    }
    const singleDomain = /^\/v1\.0\/domains\/([^/]+)$/.exec(path);
    if (singleDomain) {
      const found = domains.find(row => row.id === decodeURIComponent(singleDomain[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such domain" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? domainTransport();
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

function runDomainCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-domains-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, domains, verification, service, topRecords, denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists domains with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "domain", "list", "--profile", profile], overrides);
      assert.deepEqual(result.domains, [
        { id: dom1.id, authenticationType: "Managed", isVerified: true, isDefault: true },
        { id: dom2.id, authenticationType: "Managed", isVerified: false, isDefault: false },
        { id: dom3.id },
      ]);
      assert.deepEqual(result.count, "3 domains");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("no --filter")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/domains?"));
      assert.ok(!new URL(requests[0].url).searchParams.has("$filter"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} domain list has no --filter flag to misuse`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      await assert.rejects(
        executeArgv(["entra", "domain", "list", "--profile", profile, "--filter", "isVerified eq true"], overrides),
        error => {
          assert.equal(error.code, "VALIDATION_ERROR");
          return /unknown flag --filter/.test(error.message);
        },
      );
      assert.equal(requests.length, 0);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped domain list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "domain", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.domains.map(row => row.id), [dom1.id]);
      assert.equal(first.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "domain", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.domains.map(row => row.id), [dom2.id, dom3.id]);
      assert.deepEqual(second.count, "2 domains");
      assert.equal(second.total, null);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a verified domain bare and an unverified one with its verification hint`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const verified = await executeArgv(["entra", "domain", "show", "--id", dom1.id, "--profile", profile], overrides);
      assert.deepEqual(verified.domain, dom1);
      assert.equal(verified.help, undefined);
      const unverified = await executeArgv(["entra", "domain", "show", "--id", dom2.id, "--profile", profile,
        "--select", "id,isVerified"], overrides);
      assert.deepEqual(unverified.domain, { id: dom2.id, isVerified: false });
      assert.ok(unverified.help.some(hint => hint.includes("entra domain verification-dns-record list --domain")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists verification records with kinds and a show hint`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "domain", "verification-dns-record", "list",
        "--domain", dom1.id, "--profile", profile], overrides);
      assert.deepEqual(result.verificationDnsRecords, [
        { "@odata.type": "#microsoft.graph.domainDnsTxtRecord", id: v1.id, label: "contoso.com", recordType: "Txt", supportedService: "Email" },
      ]);
      assert.deepEqual(result.count, "1 DNS records");
      assert.equal(result.total, null);
      assert.ok(result.help.some(hint => hint.includes("@odata.type")));
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/domains/${dom1.id}/verificationDnsRecords`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one verification record with base properties`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "domain", "verification-dns-record", "show",
        "--domain", dom1.id, "--id", v1.id, "--profile", profile, "--select", "id,text"], overrides);
      assert.deepEqual(result.verificationDnsRecord, {
        "@odata.type": "#microsoft.graph.domainDnsTxtRecord",
        id: v1.id,
        text: v1.text,
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists service records with derived-type detail on explicit select`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "domain", "service-configuration-record", "list",
        "--domain", dom1.id, "--profile", profile, "--select", "id,label,recordType,mailExchange,preference"], overrides);
      const mx = result.serviceConfigurationRecords.find(row => row.id === s1.id);
      assert.equal(mx.mailExchange, s1.mailExchange);
      assert.equal(mx.preference, 0);
      assert.equal(mx["@odata.type"], s1["@odata.type"]);
      assert.deepEqual(result.count, "2 DNS records");
      assert.equal(result.total, null);
      assert.ok(new URL(requests[0].url).searchParams.get("$select").includes("mailExchange"));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one service record and lists top-level DNS records`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const shown = await executeArgv(["entra", "domain", "service-configuration-record", "show",
        "--domain", dom1.id, "--id", s1.id, "--profile", profile, "--select", "id,mailExchange"], overrides);
      assert.deepEqual(shown.serviceConfigurationRecord, {
        "@odata.type": "#microsoft.graph.domainDnsMxRecord",
        id: s1.id,
        mailExchange: s1.mailExchange,
      });
      const listed = await executeArgv(["entra", "domain-dns-record", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.domainDnsRecords, [
        { "@odata.type": "#microsoft.graph.domainDnsCnameRecord", id: t1.id, label: t1.label, recordType: "CName", supportedService: "Email" },
      ]);
      assert.deepEqual(listed.count, "1 DNS records");
      assert.equal(listed.total, null);
      const top = await executeArgv(["entra", "domain-dns-record", "show", "--id", t1.id, "--profile", profile,
        "--select", "id,canonicalName"], overrides);
      assert.deepEqual(top.domainDnsRecord, {
        "@odata.type": "#microsoft.graph.domainDnsCnameRecord",
        id: t1.id,
        canonicalName: t1.canonicalName,
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} states an empty DNS list as the answer`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const result = await executeArgv(["entra", "domain", "service-configuration-record", "list",
        "--domain", dom1.id, "--profile", profile], overrides);
      assert.deepEqual(result.serviceConfigurationRecords, []);
      assert.deepEqual(result.count, "0 DNS records");
      assert.equal(result.total, null);
      assert.equal(result.complete, true);
      assert.ok(result.help.some(hint => hint.includes("0 DNS records matched; the absence of results is the answer, not an error")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied domain reads surface scope, roles and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const missing = transport(() => json(404, { error: { code: "Request_ResourceNotFound", message: "gone" } }));
      const overrides = {
        transport: denied.send,
        delegated: credentialService("delegated", []),
        application: credentialService("application", []),
      };
      await assert.rejects(executeArgv(["entra", "domain", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Domain.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Domain Name Administrator or Global Reader")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "domain", "verification-dns-record", "show",
        "--domain", dom1.id, "--id", v1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("Domain.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "domain", "show", "--id", dom1.id, "--profile", profile],
        { ...overrides, transport: missing.send }), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists domains and records, shows one of each`, () => {
    const state = setupProfiles();
    try {
      const listed = runDomainCli(["entra", "domain", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const domainsOut = decode(listed.stdout);
      assert.deepEqual(domainsOut.domains.map(row => row.id), [dom1.id, dom2.id, dom3.id]);
      assert.deepEqual(domainsOut.count, "3 domains");
      assert.equal(domainsOut.total, null);
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runDomainCli(["entra", "domain", "show", "--id", dom1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).domain, dom1);

      const unverified = runDomainCli(["entra", "domain", "show", "--id", dom2.id, "--profile", profile], state, mode);
      assert.equal(unverified.status, 0, unverified.stdout);
      assert.ok(decode(unverified.stdout).help.some(hint => hint.includes("verification-dns-record list")));

      const verificationOut = runDomainCli(["entra", "domain", "verification-dns-record", "list",
        "--domain", dom1.id, "--profile", profile], state, mode);
      assert.equal(verificationOut.status, 0, verificationOut.stdout);
      assert.equal(decode(verificationOut.stdout).verificationDnsRecords.length, 1);

      const serviceOut = runDomainCli(["entra", "domain", "service-configuration-record", "list",
        "--domain", dom1.id, "--profile", profile], state, mode);
      assert.equal(serviceOut.status, 0, serviceOut.stdout);
      assert.equal(decode(serviceOut.stdout).serviceConfigurationRecords.length, 2);

      const topOut = runDomainCli(["entra", "domain-dns-record", "list", "--profile", profile], state, mode);
      assert.equal(topOut.status, 0, topOut.stdout);
      assert.equal(decode(topOut.stdout).domainDnsRecords.length, 1);
      assert.ok(!topOut.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied domain reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runDomainCli(["entra", "domain", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.domains, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });
}

test("truncated DNS text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const partial = await executeArgv(["entra", "domain", "service-configuration-record", "list",
      "--domain", dom1.id, "--profile", "soc", "--select", "id,label,recordType,text"], overrides);
    const txt = partial.serviceConfigurationRecords.find(row => row.id === s2.id);
    assert.match(txt.text, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "domain", "service-configuration-record", "list",
      "--domain", dom1.id, "--profile", "soc", "--full", "--select", "id,label,recordType,text"], overrides);
    assert.equal(full.serviceConfigurationRecords.find(row => row.id === s2.id).text, longSpf);
    assert.ok(!full.help.some(hint => hint.includes("--full")));
  } finally {
    teardownProfiles(state);
  }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "domain", "list", "--profile", "batch", "--scopes", domainScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "domain", "verification-dns-record", "list",
        "--domain", dom1.id, "--profile", "batch", "--scopes", domainScopes[0]], overrides),
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
    await assert.rejects(executeArgv(["entra", "domain", "list", "--profile", "soc", "--select", "id,owner"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "domain", "list", "--profile", "soc", "--fields", "isRoot"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "domain", "verification-dns-record", "list",
      "--domain", dom1.id, "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "domain-dns-record", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("domain read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "domain", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "domain", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "domain", "list", "--filter", "isVerified eq true"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "domain", "verification-dns-record", "list"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "domain", "verification-dns-record", "list", "--domain="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "domain", "service-configuration-record", "show",
    "--domain", dom1.id]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "domain-dns-record", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "domain", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
});

for (const [command, args] of [
  [["domain", "show"], []],
  [["domain", "verification-dns-record", "show"], ["--domain", dom1.id]],
  [["domain", "service-configuration-record", "show"], ["--domain", dom1.id]],
  [["domain-dns-record", "show"], []],
]) {
  for (const id of ["$count", "$value", "$ref"]) {
    test(`${command.join(" ")} rejects reserved resource binding ${id} before credentials`, async () => {
      const state = setupProfiles();
      try {
        const { calls, requests, overrides } = overridesFor("delegated");
        await assert.rejects(
          executeArgv(["entra", ...command, ...args, "--id", id, "--profile", "soc"], overrides),
          { code: "VALIDATION_ERROR", message: "OData reserved segments cannot be resource identifiers" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  for (const preview of [false, true]) {
    for (const [command, args] of [
      [["domain", "list"], []],
      [["domain", "show"], ["--id", dom1.id]],
      [["domain", "verification-dns-record", "list"], ["--domain", dom1.id]],
      [["domain", "verification-dns-record", "show"], ["--domain", dom1.id, "--id", v1.id]],
      [["domain", "service-configuration-record", "list"], ["--domain", dom1.id]],
      [["domain", "service-configuration-record", "show"], ["--domain", dom1.id, "--id", s1.id]],
      [["domain-dns-record", "list"], []],
      [["domain-dns-record", "show"], ["--id", t1.id]],
    ]) {
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
            { code: "VALIDATION_ERROR", message: "Domain reads support v1.0 only; beta needs its own review" },
          );
          assert.equal(calls.length, 0);
          assert.equal(requests.length, 0);
        } finally {
          teardownProfiles(state);
        }
      });
    }
  }
}
