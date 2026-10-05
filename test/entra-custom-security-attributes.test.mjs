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
const customSecurityScopes = ["https://graph.microsoft.com/CustomSecAttributeDefinition.Read.All"];
const base = "/v1.0/directory";

const longSetDescription = `Attributes for the engineering organization${" with considerable detail".repeat(30)}`;
const set1 = { id: "Engineering", description: "Attributes for engineering team", maxAttributesPerSet: 25 };
const set2 = { id: "Marketing", description: longSetDescription, maxAttributesPerSet: 10 };
const set3 = { id: "Bare" };
const attributeSets = [set1, set2, set3];

const longDefinitionDescription = `Active projects for the user${" with considerable detail".repeat(30)}`;
const def1 = {
  attributeSet: "Engineering",
  description: "Active projects for user",
  id: "Engineering_Project",
  isCollection: true,
  isSearchable: true,
  name: "Project",
  status: "Available",
  type: "String",
  usePreDefinedValuesOnly: true,
};
const def2 = {
  attributeSet: "Engineering",
  description: longDefinitionDescription,
  id: "Engineering_ProjectDate",
  isCollection: false,
  isSearchable: true,
  name: "ProjectDate",
  status: "Available",
  type: "String",
  usePreDefinedValuesOnly: false,
};
const def3 = { id: "Engineering_Bare" };
const definitions = [def1, def2, def3];

const val1 = { id: "Alpine", isActive: true };
const val2 = { id: "Baker", isActive: false };
const val3 = { id: "Cascade" };
const allowedValues = [val1, val2, val3];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-custom-security-"));
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

function text(value) {
  return { status: 200, headers: { "Content-Type": "text/plain" }, body: String(value) };
}

function paged(first, route) {
  return json(200, {
    value: first,
    "@odata.nextLink": `https://graph.microsoft.com${route}?%24skiptoken=page2`,
  });
}

function customSecurityTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === `${base}/attributeSets/$count`) return text(attributeSets.length);
    if (path === `${base}/customSecurityAttributeDefinitions/$count`) return text(definitions.length);
    if (path.endsWith("/allowedValues/$count")) return text(allowedValues.length);
    if (path === `${base}/attributeSets`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [set3] });
      return paged([set1, set2], `${base}/attributeSets`);
    }
    if (path === `${base}/customSecurityAttributeDefinitions`) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [def3] });
      return paged([def1, def2], `${base}/customSecurityAttributeDefinitions`);
    }
    const valuesList = new RegExp(`^${base}/customSecurityAttributeDefinitions/([^/]+)/allowedValues$`).exec(path);
    if (valuesList) {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [val3] });
      return paged([val1, val2], `${base}/customSecurityAttributeDefinitions/${valuesList[1]}/allowedValues`);
    }
    const valueSingle = new RegExp(`^${base}/customSecurityAttributeDefinitions/([^/]+)/allowedValues/([^/]+)$`).exec(path);
    if (valueSingle) {
      const found = allowedValues.find(row => row.id === decodeURIComponent(valueSingle[2]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such value" } });
    }
    const setSingle = new RegExp(`^${base}/attributeSets/([^/]+)$`).exec(path);
    if (setSingle) {
      const found = attributeSets.find(row => row.id === decodeURIComponent(setSingle[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such set" } });
    }
    const defSingle = new RegExp(`^${base}/customSecurityAttributeDefinitions/([^/]+)$`).exec(path);
    if (defSingle) {
      const found = definitions.find(row => row.id === decodeURIComponent(defSingle[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such definition" } });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? customSecurityTransport();
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

function runCustomSecurityCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-custom-security-attributes-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, attributeSets: [set1, set2], definitions: [def1, def2], allowedValues: [val1, val2], denied }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists attribute sets with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "attribute-set", "list", "--profile", profile,
        "--select", "id,description,maxAttributesPerSet"], overrides);
      assert.deepEqual(result.attributeSets, [
        set1,
        { id: set2.id, description: `${longSetDescription.slice(0, 500)}... (truncated, ${longSetDescription.length} chars total)`, maxAttributesPerSet: 10 },
        { id: set3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra attribute-set show --id <set-id>")));
      assert.ok(result.help.some(hint => hint.includes("offer no --filter")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com${base}/attributeSets?`));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(customSecurityScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped attribute-set list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "attribute-set", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.attributeSets.map(row => row.id), [set1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "attribute-set", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.attributeSets.map(row => row.id), [set2.id, set3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one attribute set with the full reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "attribute-set", "show", "--id", set1.id, "--profile", profile,
        "--select", "id,description,maxAttributesPerSet"], overrides);
      assert.deepEqual(result.attributeSet, set1);
      assert.equal(result.help, undefined);
      const missing = await executeArgv(["entra", "attribute-set", "show", "--id", set3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.attributeSet, { id: set3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts attribute sets as one scalar without a collection query`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "attribute-set", "count", "--profile", profile], overrides);
      assert.deepEqual(result, { count: { returned: 3, complete: true } });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, `https://graph.microsoft.com${base}/attributeSets/$count`);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists definitions with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "custom-security-attribute-definition", "list", "--profile", profile,
        "--select", "id,attributeSet,name,status,type,description"], overrides);
      assert.deepEqual(result.customSecurityAttributeDefinitions, [
        { id: def1.id, attributeSet: "Engineering", name: "Project", status: "Available", type: "String", description: "Active projects for user" },
        { id: def2.id, attributeSet: "Engineering", name: "ProjectDate", status: "Available", type: "String",
          description: `${longDefinitionDescription.slice(0, 500)}... (truncated, ${longDefinitionDescription.length} chars total)` },
        { id: def3.id },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra custom-security-attribute-definition show --id <definition-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com${base}/customSecurityAttributeDefinitions?`));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(customSecurityScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists definitions with a plain documented $filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "custom-security-attribute-definition", "list", "--profile", profile,
        "--filter", "status eq 'Available'"], overrides);
      assert.equal(result.count.returned, 3);
      const sent = new URL(requests[0].url).searchParams;
      assert.equal(sent.get("$filter"), "status eq 'Available'");
      assert.equal(requests[0].headers.ConsistencyLevel, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped definition list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "custom-security-attribute-definition", "list", "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.customSecurityAttributeDefinitions.map(row => row.id), [def1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "custom-security-attribute-definition", "list", "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.customSecurityAttributeDefinitions.map(row => row.id), [def2.id, def3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one definition with the full reviewed set and an allowed-values hint`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "custom-security-attribute-definition", "show", "--id", def1.id, "--profile", profile,
        "--select", "attributeSet,description,id,isCollection,isSearchable,name,status,type,usePreDefinedValuesOnly"], overrides);
      assert.deepEqual(result.customSecurityAttributeDefinition, def1);
      assert.ok(!("allowedValues" in result.customSecurityAttributeDefinition));
      assert.ok(result.help.some(hint => hint.includes(`entra allowed-value list --definition ${def1.id}`)));
      const missing = await executeArgv(["entra", "custom-security-attribute-definition", "show", "--id", def3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.customSecurityAttributeDefinition, { id: def3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts definitions as one scalar with optional server-side filter`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "custom-security-attribute-definition", "count", "--profile", profile], overrides);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, `https://graph.microsoft.com${base}/customSecurityAttributeDefinitions/$count`);
      const filtered = await executeArgv(["entra", "custom-security-attribute-definition", "count", "--profile", profile,
        "--filter", "status eq 'Available'"], overrides);
      assert.deepEqual(filtered.count, { returned: 3, complete: true });
      assert.ok(filtered.help.some(hint => hint.includes("--filter")));
      assert.equal(new URL(requests[1].url).searchParams.get("$filter"), "status eq 'Available'");
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists allowed values of one definition preserving active and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, calls, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "allowed-value", "list", "--definition", def1.id, "--profile", profile], overrides);
      assert.deepEqual(result.allowedValues, [val1, val2, { id: val3.id }]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes(`entra allowed-value show --definition ${def1.id} --id <value-id>`)));
      assert.ok(result.help.some(hint => hint.includes("offer no --filter")));
      assert.ok(requests[0].url.startsWith(`https://graph.microsoft.com${base}/customSecurityAttributeDefinitions/${def1.id}/allowedValues?`));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
      if (mode === "delegated") assert.ok(calls.some(([, , scopes]) => JSON.stringify(scopes) === JSON.stringify(customSecurityScopes)));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped allowed-value list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "allowed-value", "list", "--definition", def1.id, "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.allowedValues.map(row => row.id), [val1.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "allowed-value", "list", "--definition", def1.id, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.allowedValues.map(row => row.id), [val2.id, val3.id]);
      assert.deepEqual(second.count, { returned: 2, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows one allowed value with the reviewed set`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "allowed-value", "show",
        "--definition", def1.id, "--id", val1.id, "--profile", profile], overrides);
      assert.deepEqual(result.allowedValue, val1);
      assert.equal(result.help, undefined);
      const missing = await executeArgv(["entra", "allowed-value", "show",
        "--definition", def1.id, "--id", val3.id, "--profile", profile], overrides);
      assert.deepEqual(missing.allowedValue, { id: val3.id });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} counts allowed values of one definition as one scalar`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "allowed-value", "count", "--definition", def1.id, "--profile", profile], overrides);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.equal(requests.length, 1);
      assert.equal(requests[0].url, `https://graph.microsoft.com${base}/customSecurityAttributeDefinitions/${def1.id}/allowedValues/$count`);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} empty custom-security lists stay definitive`, async () => {
    const state = setupProfiles();
    try {
      const empty = transport(() => json(200, { value: [] }));
      const { overrides } = overridesFor(mode, empty);
      const listed = await executeArgv(["entra", "attribute-set", "list", "--profile", profile], overrides);
      assert.deepEqual(listed.attributeSets, []);
      assert.deepEqual(listed.count, { returned: 0, complete: true });
      assert.ok(listed.help.some(hint => hint.includes("0 attribute sets matched")));
      const defined = await executeArgv(["entra", "custom-security-attribute-definition", "list", "--profile", profile], overrides);
      assert.deepEqual(defined.customSecurityAttributeDefinitions, []);
      assert.deepEqual(defined.count, { returned: 0, complete: true });
      assert.ok(defined.help.some(hint => hint.includes("0 custom security attribute definitions matched")));
      const valued = await executeArgv(["entra", "allowed-value", "list", "--definition", def1.id, "--profile", profile], overrides);
      assert.deepEqual(valued.allowedValues, []);
      assert.deepEqual(valued.count, { returned: 0, complete: true });
      assert.ok(valued.help.some(hint => hint.includes("0 allowed values matched")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} unknown set, definition and value ids report absence, not emptiness`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      await assert.rejects(executeArgv(["entra", "attribute-set", "show", "--id", "set-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.deepEqual(error.suggestions, ["Verify the bound identifier; absence is not proof of nonexistence"]);
        return /not found or inaccessible \(404\)/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "custom-security-attribute-definition", "show", "--id", "definition-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "allowed-value", "show", "--definition", def1.id, "--id", "value-missing", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible \(404\)/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied custom-security reads surface scope, attribute roles and licensing`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const { overrides } = overridesFor(mode, denied);
      await assert.rejects(executeArgv(["entra", "attribute-set", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("CustomSecAttributeDefinition.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("Attribute Definition Reader")));
        assert.ok(error.suggestions.some(hint => hint.includes("even for Global Administrators")));
        assert.ok(error.suggestions.some(hint => hint.includes("Personal Microsoft accounts are not supported")));
        assert.ok(error.suggestions.some(hint => hint.includes("never diagnose licence solely from HTTP 403")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "custom-security-attribute-definition", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("CustomSecAttributeDefinition.Read.All")));
        assert.ok(error.suggestions.some(hint => hint.includes("even for Global Administrators")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "allowed-value", "list", "--definition", def1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("CustomSecAttributeDefinition.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "allowed-value", "count", "--definition", def1.id, "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        assert.ok(error.suggestions.some(hint => hint.includes("CustomSecAttributeDefinition.Read.All")));
        return /grant, role, licence/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists, shows and counts custom-security reads`, () => {
    const state = setupProfiles();
    try {
      const listed = runCustomSecurityCli(["entra", "attribute-set", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const listOut = decode(listed.stdout);
      assert.deepEqual(listOut.attributeSets.map(row => row.id), [set1.id, set2.id]);
      assert.deepEqual(listOut.count, { returned: 2, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runCustomSecurityCli(["entra", "attribute-set", "show", "--id", set1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).attributeSet, set1);

      const counted = runCustomSecurityCli(["entra", "attribute-set", "count", "--profile", profile], state, mode);
      assert.equal(counted.status, 0, counted.stdout);
      assert.deepEqual(decode(counted.stdout).count, { returned: 2, complete: true });

      const defined = runCustomSecurityCli(["entra", "custom-security-attribute-definition", "list", "--profile", profile], state, mode);
      assert.equal(defined.status, 0, defined.stdout);
      const definedOut = decode(defined.stdout);
      assert.deepEqual(definedOut.customSecurityAttributeDefinitions.map(row => row.id), [def1.id, def2.id]);
      assert.deepEqual(definedOut.count, { returned: 2, complete: true });

      const definitionShown = runCustomSecurityCli(["entra", "custom-security-attribute-definition", "show", "--id", def1.id, "--profile", profile], state, mode);
      assert.equal(definitionShown.status, 0, definitionShown.stdout);
      assert.deepEqual(decode(definitionShown.stdout).customSecurityAttributeDefinition, def1);

      const definitionCounted = runCustomSecurityCli(["entra", "custom-security-attribute-definition", "count", "--profile", profile], state, mode);
      assert.equal(definitionCounted.status, 0, definitionCounted.stdout);
      assert.deepEqual(decode(definitionCounted.stdout).count, { returned: 2, complete: true });

      const valued = runCustomSecurityCli(["entra", "allowed-value", "list", "--definition", def1.id, "--profile", profile], state, mode);
      assert.equal(valued.status, 0, valued.stdout);
      const valuedOut = decode(valued.stdout);
      assert.deepEqual(valuedOut.allowedValues.map(row => row.id), [val1.id, val2.id]);

      const valueShown = runCustomSecurityCli(["entra", "allowed-value", "show", "--definition", def1.id, "--id", val1.id, "--profile", profile], state, mode);
      assert.equal(valueShown.status, 0, valueShown.stdout);
      assert.deepEqual(decode(valueShown.stdout).allowedValue, val1);

      const valueCounted = runCustomSecurityCli(["entra", "allowed-value", "count", "--definition", def1.id, "--profile", profile], state, mode);
      assert.equal(valueCounted.status, 0, valueCounted.stdout);
      assert.deepEqual(decode(valueCounted.stdout).count, { returned: 2, complete: true });
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied custom-security reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runCustomSecurityCli(["entra", "attribute-set", "list", "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.attributeSets, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  for (const [command, args] of [
    [["attribute-set", "list"], ["--limit", "1"]],
    [["attribute-set", "show"], ["--id", set1.id]],
    [["attribute-set", "count"], []],
    [["custom-security-attribute-definition", "list"], ["--limit", "1"]],
    [["custom-security-attribute-definition", "show"], ["--id", def1.id]],
    [["custom-security-attribute-definition", "count"], []],
    [["allowed-value", "list"], ["--definition", def1.id]],
    [["allowed-value", "show"], ["--definition", def1.id, "--id", val1.id]],
    [["allowed-value", "count"], ["--definition", def1.id]],
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
            { code: "VALIDATION_ERROR", message: "Custom-security-attribute reads support v1.0 only; beta needs its own review" },
          );
          assert.equal(calls.length, 0);
          assert.equal(requests.length, 0);
        } finally {
          teardownProfiles(state);
        }
      });
    }
  }

  for (const id of ["delta()", "DELTA()", "delta(", "delta)", "delta%28%29"]) {
    test(`${mode} attribute-set show rejects function-style identifier ${id} before credentials`, async () => {
      const state = setupProfiles();
      try {
        const { requests, calls, overrides } = overridesFor(mode);
        await assert.rejects(
          executeArgv(["entra", "attribute-set", "show", "--id", id, "--profile", profile], overrides),
          { code: "VALIDATION_ERROR" },
        );
        assert.equal(calls.length, 0);
        assert.equal(requests.length, 0);
      } finally {
        teardownProfiles(state);
      }
    });
  }

  for (const id of ["delta()", "DELTA()"]) {
    test(`${mode} definition and allowed-value reads reject function-style identifier ${id} before credentials`, async () => {
      const state = setupProfiles();
      try {
        const { requests, calls, overrides } = overridesFor(mode);
        await assert.rejects(
          executeArgv(["entra", "custom-security-attribute-definition", "show", "--id", id, "--profile", profile], overrides),
          { code: "VALIDATION_ERROR" },
        );
        await assert.rejects(
          executeArgv(["entra", "allowed-value", "list", "--definition", id, "--profile", profile], overrides),
          { code: "VALIDATION_ERROR" },
        );
        await assert.rejects(
          executeArgv(["entra", "allowed-value", "show", "--definition", def1.id, "--id", id, "--profile", profile], overrides),
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

test("truncated definition text carries a --full hint without lifting caps", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const partial = await executeArgv(["entra", "custom-security-attribute-definition", "list", "--profile", "soc",
      "--select", "id,description"], overrides);
    const truncated = partial.customSecurityAttributeDefinitions.find(row => row.id === def2.id);
    assert.match(truncated.description, /truncated, \d+ chars total/);
    assert.ok(partial.help.some(hint => hint.includes("--full")));
    const full = await executeArgv(["entra", "custom-security-attribute-definition", "list", "--profile", "soc", "--full",
      "--select", "id,description"], overrides);
    assert.equal(full.customSecurityAttributeDefinitions.find(row => row.id === def2.id).description, longDefinitionDescription);
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
      executeArgv(["entra", "attribute-set", "list", "--profile", "batch", "--scopes", customSecurityScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "allowed-value", "count", "--definition", def1.id, "--profile", "batch", "--scopes", customSecurityScopes[0]], overrides),
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
    await assert.rejects(executeArgv(["entra", "attribute-set", "list", "--profile", "soc", "--select", "id,owner"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "custom-security-attribute-definition", "list", "--profile", "soc", "--fields", "usePreDefinedValuesOnly"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "allowed-value", "show", "--definition", def1.id, "--id", val1.id, "--profile", "soc", "--select", "id,zone"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "attribute-set", "list", "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("malformed custom-security count bodies fail as unknown, not zero", async () => {
  const state = setupProfiles();
  try {
    for (const body of ["{}", "-1", "2.5", "\"3\""]) {
      const malformed = transport(() => ({ status: 200, headers: {}, body }));
      const scoped = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "attribute-set", "count", "--profile", "soc"], scoped.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /non-numeric success body/.test(error.message);
      });
      const defined = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "custom-security-attribute-definition", "count", "--profile", "soc"], defined.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /non-numeric success body/.test(error.message);
      });
      const valued = overridesFor("delegated", malformed);
      await assert.rejects(executeArgv(["entra", "allowed-value", "count", "--definition", def1.id, "--profile", "soc"], valued.overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /non-numeric success body/.test(error.message);
      });
    }
  } finally {
    teardownProfiles(state);
  }
});

test("custom-security read flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "attribute-set", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "attribute-set", "show", "--id="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "custom-security-attribute-definition", "show"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "allowed-value", "list"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "allowed-value", "list", "--definition="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "allowed-value", "show", "--definition", def1.id]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "attribute-set", "count", "--limit", "5"]), /unknown flag --limit/);
  await assert.rejects(executeArgv(["entra", "attribute-set", "count", "--filter", "status eq 'Available'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "attribute-set", "list", "--filter", "id eq 'Engineering'"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "custom-security-attribute-definition", "count", "--select", "id"]), /unknown flag --select/);
  await assert.rejects(executeArgv(["entra", "custom-security-attribute-definition", "count", "--cursor", "x"]), /unknown flag --cursor/);
  await assert.rejects(executeArgv(["entra", "allowed-value", "list", "--definition", def1.id, "--filter", "isActive eq true"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "allowed-value", "count", "--definition", def1.id, "--filter", "isActive eq true"]), /unknown flag --filter/);
  await assert.rejects(executeArgv(["entra", "attribute-set", "list", "--bogus"]), { code: "VALIDATION_ERROR" });
});
