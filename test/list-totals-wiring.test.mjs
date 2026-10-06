import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

// Uniform "N of M" wiring: every remaining core-directory list renders its
// server-supplied total through the shared list-totals helper. One table row
// per wired list command covers the three honest shapes (complete unknown
// total, complete known total, partial unknown total) through the real
// executeArgv path with fake transports only.

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";

const ID_A = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const ID_B = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const CONTACT_ID = "25caf6a2-d5cb-470d-8940-20ba795ef62d";
const AU_ID = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";
const ORG_ID = "84841066-274d-4ec0-a5c1-276be684bdd3";

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-list-totals-"));
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(dir, "config.json");
  const profiles = new Profiles();
  profiles.create("soc", tenant, client, "commercial", false);
  return { dir, previous };
}

function teardownProfiles(state) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function credentialService() {
  return {
    credential: async () => ({
      token: "opaque-fixture-delegated-token",
      expiresAt: Date.now() + 3_600_000,
      tenantId: tenant,
      clientId: client,
      accountId: "synthetic-account",
    }),
  };
}

function transport(handler) {
  const send = async request => handler(request);
  return { send };
}

function json(status, body, headers = {}) {
  return { status, headers, body: JSON.stringify(body) };
}

function overridesFor(handler) {
  const credential = credentialService();
  return {
    transport: transport(handler).send,
    delegated: credential,
    application: credentialService(),
  };
}

const userRows = () => [
  { id: ID_A, displayName: "Adele Vance", userPrincipalName: "AdeleV@contoso.com", mail: "AdeleV@contoso.com" },
  { id: ID_B, displayName: "Alex Wilber", userPrincipalName: "AlexW@contoso.com", mail: null },
];

const appRows = () => [
  { id: ID_A, appId: "99999999-9999-4999-8999-999999999999", displayName: "Contoso Web" },
  { id: ID_B, appId: "88888888-8888-4888-8888-888888888888", displayName: "Daemon Batch" },
];

// Every core-directory list command wired through src/list-totals.ts: the
// argv to run, the output collection key, and the plural noun in its count
// line. Fixture rows stay minimal; they only need to survive projection.
const wired = [
  { name: "user list", argv: ["entra", "user", "list", "--profile", "soc"], key: "users", noun: "users", rows: userRows },
  { name: "application list", argv: ["entra", "application", "list", "--profile", "soc"], key: "applications", noun: "applications", rows: appRows },
  { name: "service-principal list", argv: ["entra", "service-principal", "list", "--profile", "soc"], key: "servicePrincipals", noun: "service principals", rows: appRows },
  {
    name: "application owner list",
    argv: ["entra", "application", "owner", "list", "--application", ID_A, "--profile", "soc"],
    key: "owners",
    noun: "owners",
    rows: () => [
      { "@odata.type": "#microsoft.graph.user", id: ID_A, displayName: "Adele Vance" },
      { "@odata.type": "#microsoft.graph.user", id: ID_B, displayName: "Alex Wilber" },
    ],
  },
  {
    name: "service-principal owner list",
    argv: ["entra", "service-principal", "owner", "list", "--service-principal", ID_A, "--profile", "soc"],
    key: "owners",
    noun: "owners",
    rows: () => [
      { "@odata.type": "#microsoft.graph.user", id: ID_A, displayName: "Adele Vance" },
      { "@odata.type": "#microsoft.graph.user", id: ID_B, displayName: "Alex Wilber" },
    ],
  },
  {
    name: "device list",
    argv: ["entra", "device", "list", "--profile", "soc"],
    key: "devices",
    noun: "devices",
    rows: () => [
      { id: ID_A, displayName: "CONTOSO-WIN10" },
      { id: ID_B, displayName: "CONTOSO-ANDROID" },
    ],
  },
  {
    name: "administrative-unit list",
    argv: ["entra", "administrative-unit", "list", "--profile", "soc"],
    key: "administrativeUnits",
    noun: "administrative units",
    rows: () => [
      { id: ID_A, displayName: "Seattle Schools" },
      { id: ID_B, displayName: "US Sales" },
    ],
  },
  {
    name: "administrative-unit member list",
    argv: ["entra", "administrative-unit", "member", "list", "--administrative-unit", AU_ID, "--profile", "soc"],
    key: "members",
    noun: "members",
    rows: () => [
      { "@odata.type": "#microsoft.graph.user", id: ID_A, displayName: "Adele Vance" },
      { "@odata.type": "#microsoft.graph.user", id: ID_B, displayName: "Alex Wilber" },
    ],
  },
  {
    name: "contact list",
    argv: ["entra", "contact", "list", "--profile", "soc"],
    key: "contacts",
    noun: "contacts",
    rows: () => [
      { id: ID_A, displayName: "Adele Vance", mail: "AdeleVance@adatum.com", companyName: "Adatum" },
      { id: ID_B, displayName: "Alex Wilber", mail: null, companyName: "Adatum" },
    ],
  },
  {
    name: "contact list-direct-reports",
    argv: ["entra", "contact", "list-direct-reports", "--id", CONTACT_ID, "--profile", "soc"],
    key: "directReports",
    noun: "direct reports",
    rows: () => [
      { "@odata.type": "#microsoft.graph.user", id: ID_A, displayName: "Adele Vance" },
      { "@odata.type": "#microsoft.graph.orgContact", id: ID_B, displayName: "Alex Wilber" },
    ],
  },
  {
    name: "contact list-member-of",
    argv: ["entra", "contact", "list-member-of", "--id", CONTACT_ID, "--profile", "soc"],
    key: "memberOf",
    noun: "memberships",
    rows: () => [
      { "@odata.type": "#microsoft.graph.group", id: ID_A, displayName: "All Staff" },
      { "@odata.type": "#microsoft.graph.administrativeUnit", id: ID_B, displayName: "US Sales" },
    ],
  },
  {
    name: "directory-role list",
    argv: ["entra", "directory-role", "list", "--profile", "soc"],
    key: "directoryRoles",
    noun: "directory roles",
    rows: () => [
      { id: ID_A, displayName: "Global Administrator" },
      { id: ID_B, displayName: "Directory Readers" },
    ],
  },
  {
    name: "role-assignment list",
    argv: ["entra", "role-assignment", "list", "--profile", "soc"],
    key: "roleAssignments",
    noun: "role assignments",
    rows: () => [
      { id: ID_A, principalId: ID_A, roleDefinitionId: ID_B },
      { id: ID_B, principalId: ID_B, roleDefinitionId: ID_A },
    ],
  },
  {
    name: "pim eligible list",
    argv: ["entra", "pim", "eligible", "list", "--profile", "soc"],
    key: "eligibleAssignments",
    noun: "eligible assignments",
    rows: () => [
      { id: ID_A, principalId: ID_A, roleDefinitionId: ID_B, memberType: "Direct" },
      { id: ID_B, principalId: ID_B, roleDefinitionId: ID_A, memberType: "Direct" },
    ],
  },
  {
    name: "pim active list",
    argv: ["entra", "pim", "active", "list", "--profile", "soc"],
    key: "activeAssignments",
    noun: "active assignments",
    rows: () => [
      { id: ID_A, principalId: ID_A, roleDefinitionId: ID_B, assignmentType: "Assigned", memberType: "Direct" },
      { id: ID_B, principalId: ID_B, roleDefinitionId: ID_A, assignmentType: "Activated", memberType: "Direct" },
    ],
  },
  {
    name: "directory-object list",
    argv: ["entra", "directory-object", "list", "--profile", "soc"],
    key: "directoryObjects",
    noun: "directory objects",
    rows: () => [
      { "@odata.type": "#microsoft.graph.user", id: ID_A, displayName: "Adele Vance" },
      { "@odata.type": "#microsoft.graph.group", id: ID_B, displayName: "All Staff" },
    ],
  },
  {
    name: "deleted-user list",
    argv: ["entra", "deleted-user", "list", "--profile", "soc"],
    key: "deletedItems",
    noun: "deleted items",
    rows: () => [
      { "@odata.type": "#microsoft.graph.user", id: ID_A, displayName: "Adele Vance" },
      { "@odata.type": "#microsoft.graph.user", id: ID_B, displayName: "Alex Wilber" },
    ],
  },
  {
    name: "domain list",
    argv: ["entra", "domain", "list", "--profile", "soc"],
    key: "domains",
    noun: "domains",
    rows: () => [
      { id: "a.example", authenticationType: "Managed", isVerified: true },
      { id: "b.example", authenticationType: "Managed", isVerified: false },
    ],
  },
  {
    name: "domain verification-dns-record list",
    argv: ["entra", "domain", "verification-dns-record", "list", "--domain", "contoso.com", "--profile", "soc"],
    key: "verificationDnsRecords",
    noun: "DNS records",
    rows: () => [
      { "@odata.type": "#microsoft.graph.domainDnsTxtRecord", id: "rec1" },
      { "@odata.type": "#microsoft.graph.domainDnsTxtRecord", id: "rec2" },
    ],
  },
  {
    name: "contract list",
    argv: ["entra", "contract", "list", "--profile", "soc"],
    key: "contracts",
    noun: "contracts",
    rows: () => [
      { id: ID_A, displayName: "Contoso", contractType: "ResellerPartner" },
      { id: ID_B, displayName: "Fabrikam", contractType: "BreadthPartner" },
    ],
  },
  {
    name: "organization list",
    argv: ["entra", "organization", "list", "--profile", "soc"],
    key: "organizations",
    noun: "organizations",
    rows: () => [
      { id: ORG_ID, displayName: "Contoso" },
      { id: ID_B, displayName: "Fabrikam" },
    ],
  },
  {
    name: "organization branding-localization list",
    argv: ["entra", "organization", "branding-localization", "list", "--organization", ORG_ID, "--profile", "soc"],
    key: "brandingLocalizations",
    noun: "branding localizations",
    rows: () => [
      { id: "0", signInPageText: "Contoso" },
      { id: "fr-FR", signInPageText: "Contoso FR" },
    ],
  },
  {
    name: "subscription list",
    argv: ["entra", "subscription", "list", "--profile", "soc"],
    key: "subscriptions",
    noun: "subscriptions",
    rows: () => [
      { id: ID_A, skuPartNumber: "ENTERPRISEPACK", status: "Enabled" },
      { id: ID_B, skuPartNumber: "FLOW_FREE", status: "Suspended" },
    ],
  },
];

for (const entry of wired) {
  test(`${entry.name} reports uniform totals`, async () => {
    const state = setupProfiles();
    try {
      const rows = entry.rows();
      // Complete read, pages carry no usable total: names its rows.
      const complete = await executeArgv(entry.argv, overridesFor(() => json(200, { value: rows })));
      assert.equal(complete[entry.key].length, 2);
      assert.deepEqual(complete.count, `2 ${entry.noun}`);
      assert.equal(complete.total, null);
      assert.equal(complete.complete, true);
      // Complete read, pages carry the server total: names rows of total.
      const counted = await executeArgv(entry.argv, overridesFor(() => json(200, { value: rows, "@odata.count": 9 })));
      assert.deepEqual(counted.count, `2 of 9 ${entry.noun}`);
      assert.equal(counted.total, 9);
      assert.equal(counted.complete, true);
      // Partial read, total unknown: says more is available, never invents one.
      const partial = await executeArgv([...entry.argv, "--limit", "1"], overridesFor(() => json(200, { value: rows })));
      assert.equal(partial[entry.key].length, 1);
      assert.deepEqual(partial.count, `1 ${entry.noun} shown, more available`);
      assert.equal(partial.total, null);
      assert.equal(partial.complete, false);
      assert.equal(typeof partial.cursor, "string");
      assert.equal(typeof partial.reason, "string");
    } finally {
      teardownProfiles(state);
    }
  });
}
