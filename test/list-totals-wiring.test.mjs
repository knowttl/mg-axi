import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { LEAVES } from "../dist/catalogue.js";
import { Profiles } from "../dist/profiles.js";

// Uniform "N of M" wiring: every list command renders its server-supplied
// total through the shared list-totals helper. One table row per wired list
// command covers the three honest shapes (complete unknown total, complete
// known total, partial unknown total) through the real executeArgv path with
// fake transports only. The catalogue-enumeration test at the end fails on
// any list leaf that is neither wired above nor explicitly pending.

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";

const ID_A = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const ID_B = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const CONTACT_ID = "25caf6a2-d5cb-470d-8940-20ba795ef62d";
const AU_ID = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";
const ORG_ID = "84841066-274d-4ec0-a5c1-276be684bdd3";
const GROUP_PIM_FILTER = `groupId eq '${ID_B}'`;
const REVIEW_DEF_ID = "98dcebed-c7f6-46f4-bcf3-4a3fccdb3e2a";
const REVIEW_INST_ID = "7bc18cf4-3d70-4009-bc8e-a7c5adb30849";
const AUTH_USER_ID = "aaaaaaaa-1111-4111-8111-111111111111";
const PKI_ID = "bbbbbbbb-2222-4222-8222-222222222222";
const REL_ID = "cccccccc-3333-4333-8333-333333333333";
const CUST_ID = "dddddddd-4444-4434-8434-444444444444";
const PKG_ID = "eeeeeeee-5555-4555-8555-555555555555";
const CSA_DEF_ID = "ffffffff-6666-4666-8666-666666666666";
const SINCE = "2026-09-01T00:00:00Z";

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
  {
    name: "pim group-assignment-schedule list",
    argv: ["entra", "pim", "group-assignment-schedule", "list", "--profile", "soc", "--filter", GROUP_PIM_FILTER],
    key: "assignmentSchedules",
    noun: "assignment schedules",
    rows: () => [
      { id: ID_A, principalId: ID_A, groupId: ID_B, accessId: "member", assignmentType: "Assigned" },
      { id: ID_B, principalId: ID_B, groupId: ID_B, accessId: "owner", assignmentType: "Activated" },
    ],
  },
  {
    name: "pim group-assignment-instance list",
    argv: ["entra", "pim", "group-assignment-instance", "list", "--profile", "soc", "--filter", GROUP_PIM_FILTER],
    key: "assignmentScheduleInstances",
    noun: "assignment instances",
    rows: () => [
      { id: ID_A, principalId: ID_A, groupId: ID_B, accessId: "member", assignmentType: "Activated" },
      { id: ID_B, principalId: ID_B, groupId: ID_B, accessId: "owner", assignmentType: "Activated" },
    ],
  },
  {
    name: "pim group-eligibility-schedule list",
    argv: ["entra", "pim", "group-eligibility-schedule", "list", "--profile", "soc", "--filter", GROUP_PIM_FILTER],
    key: "eligibilitySchedules",
    noun: "eligibility schedules",
    rows: () => [
      { id: ID_A, principalId: ID_A, groupId: ID_B, accessId: "member", memberType: "Direct" },
      { id: ID_B, principalId: ID_B, groupId: ID_B, accessId: "owner", memberType: "Direct" },
    ],
  },
  {
    name: "pim group-eligibility-instance list",
    argv: ["entra", "pim", "group-eligibility-instance", "list", "--profile", "soc", "--filter", GROUP_PIM_FILTER],
    key: "eligibilityScheduleInstances",
    noun: "eligibility instances",
    rows: () => [
      { id: ID_A, principalId: ID_A, groupId: ID_B, accessId: "member", memberType: "Direct" },
      { id: ID_B, principalId: ID_B, groupId: ID_B, accessId: "owner", memberType: "Direct" },
    ],
  },
  {
    name: "pim group-eligibility-request list",
    argv: ["entra", "pim", "group-eligibility-request", "list", "--profile", "soc", "--filter", GROUP_PIM_FILTER],
    key: "eligibilityScheduleRequests",
    noun: "eligibility requests",
    rows: () => [
      { id: ID_A, action: "adminAssign", status: "Provisioned", principalId: ID_A, groupId: ID_B, accessId: "member" },
      { id: ID_B, action: "adminAssign", status: "Provisioned", principalId: ID_B, groupId: ID_B, accessId: "member" },
    ],
  },
  {
    name: "federation-configuration list",
    argv: ["entra", "federation-configuration", "list", "--profile", "soc"],
    key: "federationConfigurations",
    noun: "federation configurations",
    rows: () => [
      { id: ID_A, displayName: "Contoso" },
      { id: ID_B, displayName: "Fabrikam" },
    ],
  },
  {
    name: "service-principal oauth2-grant list",
    argv: ["entra", "service-principal", "oauth2-grant", "list", "--service-principal", ID_A, "--profile", "soc"],
    key: "oauth2PermissionGrants",
    noun: "oauth2 grants",
    rows: () => [
      { id: ID_A, consentType: "AllPrincipals", principalId: null, resourceId: ID_B, scope: "User.Read" },
      { id: ID_B, consentType: "Principal", principalId: ID_A, resourceId: ID_B, scope: "Mail.Read" },
    ],
  },
  {
    name: "service-principal app-role-assignment list",
    argv: ["entra", "service-principal", "app-role-assignment", "list", "--service-principal", ID_A, "--profile", "soc"],
    key: "appRoleAssignments",
    noun: "app role assignments",
    rows: () => [
      { id: ID_A, appRoleId: ID_B, resourceDisplayName: "Graph", resourceId: ID_B },
      { id: ID_B, appRoleId: ID_A, resourceDisplayName: "Exchange", resourceId: ID_A },
    ],
  },
  {
    name: "group-lifecycle-policy list",
    argv: ["entra", "group-lifecycle-policy", "list", "--profile", "soc"],
    key: "groupLifecyclePolicies",
    noun: "lifecycle policies",
    rows: () => [
      { id: ID_A, groupLifetimeInDays: 365, managedGroupTypes: "All" },
      { id: ID_B, groupLifetimeInDays: 180, managedGroupTypes: "Selected" },
    ],
  },
  {
    name: "group-setting-template list",
    argv: ["entra", "group-setting-template", "list", "--profile", "soc"],
    key: "groupSettingTemplates",
    noun: "setting templates",
    rows: () => [
      { id: ID_A, displayName: "Group.Unified", description: "Unified" },
      { id: ID_B, displayName: "Group.Unified.Guest", description: "Guest" },
    ],
  },
  {
    name: "identity-provider list",
    argv: ["entra", "identity-provider", "list", "--profile", "soc"],
    key: "identityProviders",
    noun: "identity providers",
    rows: () => [
      { id: ID_A, displayName: "Google" },
      { id: ID_B, displayName: "Facebook" },
    ],
  },
  {
    name: "lifecycle workflow list",
    argv: ["entra", "lifecycle", "workflow", "list", "--profile", "soc"],
    key: "workflows",
    noun: "workflows",
    rows: () => [
      { id: ID_A, displayName: "Joiner", category: "joiner", isEnabled: true, isSchedulingEnabled: true },
      { id: ID_B, displayName: "Leaver", category: "leaver", isEnabled: false, isSchedulingEnabled: false },
    ],
  },
  {
    name: "lifecycle workflow-template list",
    argv: ["entra", "lifecycle", "workflow-template", "list", "--profile", "soc"],
    key: "workflowTemplates",
    noun: "workflow templates",
    rows: () => [
      { id: ID_A, displayName: "Joiner", category: "joiner" },
      { id: ID_B, displayName: "Leaver", category: "leaver" },
    ],
  },
  {
    name: "lifecycle task-definition list",
    argv: ["entra", "lifecycle", "task-definition", "list", "--profile", "soc"],
    key: "taskDefinitions",
    noun: "task definitions",
    rows: () => [
      { id: ID_A, displayName: "Enable account", category: "joiner", version: 1 },
      { id: ID_B, displayName: "Disable account", category: "leaver", version: 2 },
    ],
  },
  {
    name: "lifecycle run list",
    argv: ["entra", "lifecycle", "run", "list", "--workflow", ID_A, "--profile", "soc"],
    key: "runs",
    noun: "runs",
    rows: () => [
      { id: ID_A, processingStatus: "completed", totalUsersCount: 2, failedUsersCount: 0, successfulUsersCount: 2 },
      { id: ID_B, processingStatus: "inProgress", totalUsersCount: 1, failedUsersCount: 1, successfulUsersCount: 0 },
    ],
  },
  {
    name: "lifecycle user-processing-result list",
    argv: ["entra", "lifecycle", "user-processing-result", "list", "--workflow", ID_A, "--profile", "soc"],
    key: "userProcessingResults",
    noun: "user processing results",
    rows: () => [
      { id: ID_A, processingStatus: "completed", failedTasksCount: 0, totalTasksCount: 3, totalUnprocessedTasksCount: 0 },
      { id: ID_B, processingStatus: "failed", failedTasksCount: 1, totalTasksCount: 3, totalUnprocessedTasksCount: 0 },
    ],
  },
  {
    name: "lifecycle subject-processing-result list",
    argv: ["entra", "lifecycle", "subject-processing-result", "list", "--workflow", ID_A, "--profile", "soc"],
    key: "subjectProcessingResults",
    noun: "subject processing results",
    rows: () => [
      { id: ID_A, subjectType: "user", processingStatus: "completed", failedTasksCount: 0, totalTasksCount: 3, totalUnprocessedTasksCount: 0 },
      { id: ID_B, subjectType: "provisioningObject", processingStatus: "failed", failedTasksCount: 1, totalTasksCount: 2, totalUnprocessedTasksCount: 0 },
    ],
  },
  {
    name: "lifecycle task-report list",
    argv: ["entra", "lifecycle", "task-report", "list", "--workflow", ID_A, "--profile", "soc"],
    key: "taskReports",
    noun: "task reports",
    rows: () => [
      { id: ID_A, runId: ID_B, processingStatus: "completed", totalUsersCount: 2, failedUsersCount: 0, successfulUsersCount: 2 },
      { id: ID_B, runId: ID_A, processingStatus: "inProgress", totalUsersCount: 1, failedUsersCount: 1, successfulUsersCount: 0 },
    ],
  },
  {
    name: "lifecycle run user-processing-result list",
    argv: ["entra", "lifecycle", "run", "user-processing-result", "list", "--workflow", ID_A, "--run", ID_B, "--profile", "soc"],
    key: "userProcessingResults",
    noun: "user processing results",
    rows: () => [
      { id: ID_A, processingStatus: "completed", failedTasksCount: 0, totalTasksCount: 3, totalUnprocessedTasksCount: 0 },
      { id: ID_B, processingStatus: "failed", failedTasksCount: 1, totalTasksCount: 3, totalUnprocessedTasksCount: 0 },
    ],
  },
  {
    name: "lifecycle run subject-processing-result list",
    argv: ["entra", "lifecycle", "run", "subject-processing-result", "list", "--workflow", ID_A, "--run", ID_B, "--profile", "soc"],
    key: "subjectProcessingResults",
    noun: "subject processing results",
    rows: () => [
      { id: ID_A, subjectType: "user", processingStatus: "completed", failedTasksCount: 0, totalTasksCount: 3, totalUnprocessedTasksCount: 0 },
      { id: ID_B, subjectType: "provisioningObject", processingStatus: "failed", failedTasksCount: 1, totalTasksCount: 2, totalUnprocessedTasksCount: 0 },
    ],
  },
  {
    name: "lifecycle run task-processing-result list",
    argv: ["entra", "lifecycle", "run", "task-processing-result", "list", "--workflow", ID_A, "--run", ID_B, "--profile", "soc"],
    key: "taskProcessingResults",
    noun: "task processing results",
    rows: () => [
      { id: ID_A, processingStatus: "completed", failureReason: null },
      { id: ID_B, processingStatus: "failed", failureReason: "Timeout" },
    ],
  },
  {
    name: "multi-tenant-organization tenant list",
    argv: ["entra", "multi-tenant-organization", "tenant", "list", "--profile", "soc"],
    key: "multiTenantOrganizationTenants",
    noun: "tenants",
    rows: () => [
      { tenantId: tenant, displayName: "Contoso", role: "owner", state: "active" },
      { tenantId: client, displayName: "Fabrikam", role: "member", state: "active" },
    ],
  },
  {
    name: "on-premises-synchronization list",
    argv: ["entra", "on-premises-synchronization", "list", "--profile", "soc"],
    key: "synchronizations",
    noun: "synchronizations",
    rows: () => [
      { id: tenant, configuration: { synchronizationInterval: "PT30M" }, features: { passwordHashSyncEnabled: true } },
      { id: client, configuration: { synchronizationInterval: "PT2H" }, features: { passwordHashSyncEnabled: false } },
    ],
  },
  {
    name: "risky-user list",
    argv: ["entra", "risky-user", "list", "--profile", "soc"],
    key: "riskyUsers",
    noun: "risky users",
    rows: () => [
      { id: ID_A, userPrincipalName: "AdeleV@contoso.com", riskLevel: "high", riskState: "atRisk" },
      { id: ID_B, userPrincipalName: "AlexW@contoso.com", riskLevel: "medium", riskState: "confirmedCompromised" },
    ],
  },
  {
    name: "risk-detection list",
    argv: ["entra", "risk-detection", "list", "--since", "2026-09-01T00:00:00Z", "--profile", "soc"],
    key: "riskDetections",
    noun: "risk detections",
    rows: () => [
      { id: ID_A, detectedDateTime: "2026-09-02T00:00:00Z", userPrincipalName: "AdeleV@contoso.com", riskLevel: "high" },
      { id: ID_B, detectedDateTime: "2026-09-03T00:00:00Z", userPrincipalName: "AlexW@contoso.com", riskLevel: "medium" },
    ],
  },
  {
    name: "risky-service-principal list",
    argv: ["entra", "risky-service-principal", "list", "--profile", "soc"],
    key: "riskyServicePrincipals",
    noun: "risky service principals",
    rows: () => [
      { id: ID_A, displayName: "Daemon Batch", riskLevel: "high", riskState: "atRisk" },
      { id: ID_B, displayName: "Contoso Web", riskLevel: "medium", riskState: "confirmedCompromised" },
    ],
  },
  {
    name: "risky-service-principal history list",
    argv: ["entra", "risky-service-principal", "history", "list", "--service-principal", ID_A, "--profile", "soc"],
    key: "riskyServicePrincipalHistory",
    noun: "history items",
    rows: () => [
      { id: ID_A, displayName: "Daemon Batch", riskLevel: "high", riskState: "atRisk" },
      { id: ID_B, displayName: "Daemon Batch", riskLevel: "medium", riskState: "dismissed" },
    ],
  },
  {
    name: "service-principal-risk-detection list",
    argv: ["entra", "service-principal-risk-detection", "list", "--since", "2026-09-01T00:00:00Z", "--profile", "soc"],
    key: "servicePrincipalRiskDetections",
    noun: "service principal risk detections",
    rows: () => [
      { id: ID_A, detectedDateTime: "2026-09-02T00:00:00Z", servicePrincipalDisplayName: "Daemon Batch", riskLevel: "high" },
      { id: ID_B, detectedDateTime: "2026-09-03T00:00:00Z", servicePrincipalDisplayName: "Contoso Web", riskLevel: "medium" },
    ],
  },
  {
    name: "fraud-protection-provider list",
    argv: ["entra", "fraud-protection-provider", "list", "--profile", "soc"],
    key: "fraudProtectionProviders",
    noun: "fraud protection providers",
    rows: () => [
      { id: ID_A, displayName: "Contoso Fraud" },
      { id: ID_B, displayName: "Fabrikam Fraud" },
    ],
  },
  {
    name: "web-application-firewall-provider list",
    argv: ["entra", "web-application-firewall-provider", "list", "--profile", "soc"],
    key: "webApplicationFirewallProviders",
    noun: "web application firewall providers",
    rows: () => [
      { id: ID_A, displayName: "Contoso WAF" },
      { id: ID_B, displayName: "Fabrikam WAF" },
    ],
  },
  {
    name: "web-application-firewall-verification list",
    argv: ["entra", "web-application-firewall-verification", "list", "--profile", "soc"],
    key: "webApplicationFirewallVerifications",
    noun: "web application firewall verifications",
    rows: () => [
      { id: ID_A, verifiedHost: "a.example", providerType: "waf" },
      { id: ID_B, verifiedHost: "b.example", providerType: "waf" },
    ],
  },
  {
    name: "agreement list",
    argv: ["entra", "agreement", "list", "--profile", "soc"],
    key: "agreements",
    noun: "agreements",
    rows: () => [
      { id: ID_A, displayName: "Contoso ToU" },
      { id: ID_B, displayName: "Fabrikam ToU" },
    ],
  },
  {
    name: "agreement acceptance list",
    argv: ["entra", "agreement", "acceptance", "list", "--agreement", ID_A, "--profile", "soc"],
    key: "agreementAcceptances",
    noun: "agreement acceptances",
    rows: () => [
      { id: ID_A, agreementId: ID_A, state: "accepted", recordedDateTime: "2026-09-02T00:00:00Z" },
      { id: ID_B, agreementId: ID_A, state: "declined", recordedDateTime: "2026-09-03T00:00:00Z" },
    ],
  },
  {
    name: "agreement-acceptance list",
    argv: ["entra", "agreement-acceptance", "list", "--profile", "soc"],
    key: "agreementAcceptances",
    noun: "agreement acceptances",
    rows: () => [
      { id: ID_A, agreementId: ID_A, state: "accepted", recordedDateTime: "2026-09-02T00:00:00Z" },
      { id: ID_B, agreementId: ID_B, state: "declined", recordedDateTime: "2026-09-03T00:00:00Z" },
    ],
  },
  // Already-wired lists with no table row yet: group/member/memberOf (PR 64),
  // the deleted-* siblings of deleted-user, and the DNS-record siblings of
  // verification-dns-record. Same helper, same three shapes.
  {
    name: "group list",
    argv: ["entra", "group", "list", "--profile", "soc"],
    key: "groups",
    noun: "groups",
    rows: () => [
      { id: ID_A, displayName: "All Staff", mail: "allstaff@contoso.com" },
      { id: ID_B, displayName: "Security Enabled", mail: null },
    ],
  },
  {
    name: "group member list",
    argv: ["entra", "group", "member", "list", "--group", ID_A, "--profile", "soc"],
    key: "members",
    noun: "members",
    rows: () => [
      { "@odata.type": "#microsoft.graph.user", id: ID_A, displayName: "Adele Vance" },
      { "@odata.type": "#microsoft.graph.user", id: ID_B, displayName: "Alex Wilber" },
    ],
  },
  {
    name: "group member-of list",
    argv: ["entra", "group", "member-of", "list", "--group", ID_A, "--profile", "soc"],
    key: "memberOf",
    noun: "memberships",
    rows: () => [
      { "@odata.type": "#microsoft.graph.group", id: ID_A, displayName: "All Staff" },
      { "@odata.type": "#microsoft.graph.group", id: ID_B, displayName: "Security Enabled" },
    ],
  },
  {
    name: "deleted-group list",
    argv: ["entra", "deleted-group", "list", "--profile", "soc"],
    key: "deletedItems",
    noun: "deleted items",
    rows: () => [
      { "@odata.type": "#microsoft.graph.group", id: ID_A, displayName: "All Staff" },
      { "@odata.type": "#microsoft.graph.group", id: ID_B, displayName: "Security Enabled" },
    ],
  },
  {
    name: "deleted-application list",
    argv: ["entra", "deleted-application", "list", "--profile", "soc"],
    key: "deletedItems",
    noun: "deleted items",
    rows: () => [
      { "@odata.type": "#microsoft.graph.application", id: ID_A, displayName: "Contoso Web" },
      { "@odata.type": "#microsoft.graph.application", id: ID_B, displayName: "Daemon Batch" },
    ],
  },
  {
    name: "deleted-service-principal list",
    argv: ["entra", "deleted-service-principal", "list", "--profile", "soc"],
    key: "deletedItems",
    noun: "deleted items",
    rows: () => [
      { "@odata.type": "#microsoft.graph.servicePrincipal", id: ID_A, displayName: "Contoso Web" },
      { "@odata.type": "#microsoft.graph.servicePrincipal", id: ID_B, displayName: "Daemon Batch" },
    ],
  },
  {
    name: "deleted-administrative-unit list",
    argv: ["entra", "deleted-administrative-unit", "list", "--profile", "soc"],
    key: "deletedItems",
    noun: "deleted items",
    rows: () => [
      { "@odata.type": "#microsoft.graph.administrativeUnit", id: ID_A, displayName: "Seattle Schools" },
      { "@odata.type": "#microsoft.graph.administrativeUnit", id: ID_B, displayName: "US Sales" },
    ],
  },
  {
    name: "domain service-configuration-record list",
    argv: ["entra", "domain", "service-configuration-record", "list", "--domain", "contoso.com", "--profile", "soc"],
    key: "serviceConfigurationRecords",
    noun: "DNS records",
    rows: () => [
      { "@odata.type": "#microsoft.graph.domainDnsTxtRecord", id: "rec1" },
      { "@odata.type": "#microsoft.graph.domainDnsTxtRecord", id: "rec2" },
    ],
  },
  {
    name: "domain-dns-record list",
    argv: ["entra", "domain-dns-record", "list", "--profile", "soc"],
    key: "domainDnsRecords",
    noun: "DNS records",
    rows: () => [
      { "@odata.type": "#microsoft.graph.domainDnsTxtRecord", id: "rec1" },
      { "@odata.type": "#microsoft.graph.domainDnsTxtRecord", id: "rec2" },
    ],
  },
  // Governance slice: access reviews, audit logs, authentication methods,
  // certificate auth, conditional access, custom security attributes, data
  // policy operations, delegated admin, entitlement management.
  {
    name: "access-review definition list",
    argv: ["entra", "access-review", "definition", "list", "--profile", "soc"],
    key: "definitions",
    noun: "access-review definitions",
    rows: () => [
      { id: ID_A, displayName: "Q1 access review", status: "InProgress" },
      { id: ID_B, displayName: "Monthly role review", status: "Completed" },
    ],
  },
  {
    name: "access-review instance list",
    argv: ["entra", "access-review", "instance", "list", "--definition", REVIEW_DEF_ID, "--profile", "soc"],
    key: "instances",
    noun: "access-review instances",
    rows: () => [
      { id: ID_A, status: "InProgress", startDateTime: "2026-09-01T00:00:00Z", endDateTime: "2026-09-30T00:00:00Z" },
      { id: ID_B, status: "Completed", startDateTime: "2026-08-01T00:00:00Z", endDateTime: "2026-08-31T00:00:00Z" },
    ],
  },
  {
    name: "access-review decision list",
    argv: ["entra", "access-review", "decision", "list", "--definition", REVIEW_DEF_ID, "--instance", REVIEW_INST_ID, "--profile", "soc"],
    key: "decisions",
    noun: "access-review decisions",
    rows: () => [
      { id: ID_A, accessReviewId: REVIEW_INST_ID, decision: "NotReviewed", recommendation: "Deny" },
      { id: ID_B, accessReviewId: REVIEW_INST_ID, decision: "Approve", recommendation: "Approve" },
    ],
  },
  {
    name: "access-review contacted-reviewer list",
    argv: ["entra", "access-review", "contacted-reviewer", "list", "--definition", REVIEW_DEF_ID, "--instance", REVIEW_INST_ID, "--profile", "soc"],
    key: "contactedReviewers",
    noun: "contacted reviewers",
    rows: () => [
      { id: ID_A, displayName: "Adele Vance", userPrincipalName: "AdeleV@contoso.com" },
      { id: ID_B, displayName: "Diego Siciliani", userPrincipalName: "DiegoS@contoso.com" },
    ],
  },
  {
    name: "access-review stage list",
    argv: ["entra", "access-review", "stage", "list", "--definition", REVIEW_DEF_ID, "--instance", REVIEW_INST_ID, "--profile", "soc"],
    key: "stages",
    noun: "access-review stages",
    rows: () => [
      { id: ID_A, status: "InProgress", startDateTime: "2026-09-01T00:00:00Z", endDateTime: "2026-09-10T00:00:00Z" },
      { id: ID_B, status: "NotStarted", startDateTime: "2026-09-11T00:00:00Z", endDateTime: "2026-09-20T00:00:00Z" },
    ],
  },
  {
    name: "sign-in list",
    argv: ["entra", "sign-in", "list", "--profile", "soc", "--since", SINCE],
    key: "signIns",
    noun: "sign-ins",
    rows: () => [
      { id: ID_A, createdDateTime: "2026-09-10T12:00:00Z", userPrincipalName: "AdeleV@contoso.com", appDisplayName: "Azure Portal" },
      { id: ID_B, createdDateTime: "2026-09-10T12:05:00Z", userPrincipalName: "AlexW@contoso.com", appDisplayName: "Azure Portal" },
    ],
  },
  {
    name: "directory-audit list",
    argv: ["entra", "directory-audit", "list", "--profile", "soc", "--since", SINCE],
    key: "directoryAudits",
    noun: "directory audits",
    rows: () => [
      { id: ID_A, activityDateTime: "2026-09-10T12:00:00Z", activityDisplayName: "Add member to group", result: "success" },
      { id: ID_B, activityDateTime: "2026-09-10T12:05:00Z", activityDisplayName: "Remove member from group", result: "success" },
    ],
  },
  {
    name: "user authentication-method list",
    argv: ["entra", "user", "authentication-method", "list", "--user", AUTH_USER_ID, "--profile", "soc"],
    key: "authenticationMethods",
    noun: "authentication methods",
    rows: () => [
      { id: ID_A, displayName: "Mobile phone", createdDateTime: "2024-01-01T00:00:00Z", "@odata.type": "#microsoft.graph.phoneAuthenticationMethod" },
      { id: ID_B, displayName: "Authenticator app", createdDateTime: null, "@odata.type": "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" },
    ],
  },
  {
    name: "registration list",
    argv: ["entra", "registration", "list", "--profile", "soc"],
    key: "registrationDetails",
    noun: "registration rows",
    rows: () => [
      { id: ID_A, userPrincipalName: "AdeleV@contoso.com", userDisplayName: "Adele Vance", isMfaRegistered: true },
      { id: ID_B, userPrincipalName: "AlexW@contoso.com", userDisplayName: "Alex Wilber" },
    ],
  },
  {
    name: "certificate-auth-pki list",
    argv: ["entra", "certificate-auth-pki", "list", "--profile", "soc"],
    key: "certificateAuthPkis",
    noun: "PKI configurations",
    rows: () => [
      { id: ID_A, displayName: "Contoso PKI", status: "succeeded" },
      { id: ID_B, displayName: "Fabrikam PKI" },
    ],
  },
  {
    name: "certificate-authority list",
    argv: ["entra", "certificate-authority", "list", "--pki", PKI_ID, "--profile", "soc"],
    key: "certificateAuthorities",
    noun: "certificate authorities",
    rows: () => [
      { id: ID_A, displayName: "Contoso Root CA", certificateAuthorityType: "root", expirationDateTime: "2027-08-29T02:05:57Z" },
      { id: ID_B, displayName: "Contoso Issuing CA", certificateAuthorityType: "unknown" },
    ],
  },
  {
    name: "conditional-access policy list",
    argv: ["entra", "conditional-access", "policy", "list", "--profile", "soc"],
    key: "policies",
    noun: "policies",
    rows: () => [
      { id: ID_A, displayName: "Require MFA for admins", state: "enabled" },
      { id: ID_B, displayName: "Block legacy auth", state: "disabled" },
    ],
  },
  {
    name: "conditional-access named-location list",
    argv: ["entra", "conditional-access", "named-location", "list", "--profile", "soc"],
    key: "namedLocations",
    noun: "named locations",
    rows: () => [
      { id: ID_A, displayName: "Corporate HQ", "@odata.type": "#microsoft.graph.ipNamedLocation" },
      { id: ID_B, displayName: "Blocked countries", "@odata.type": "#microsoft.graph.countryNamedLocation" },
    ],
  },
  {
    name: "attribute-set list",
    argv: ["entra", "attribute-set", "list", "--profile", "soc"],
    key: "attributeSets",
    noun: "attribute sets",
    rows: () => [
      { id: "Engineering", description: "Attributes for engineering team", maxAttributesPerSet: 25 },
      { id: "Marketing", description: "Attributes for marketing team", maxAttributesPerSet: 10 },
    ],
  },
  {
    name: "custom-security-attribute-definition list",
    argv: ["entra", "custom-security-attribute-definition", "list", "--profile", "soc"],
    key: "customSecurityAttributeDefinitions",
    noun: "custom security attribute definitions",
    rows: () => [
      { id: ID_A, attributeSet: "Engineering", name: "Project", status: "Available", type: "String" },
      { id: ID_B, attributeSet: "Engineering", name: "CostCenter", status: "Available", type: "Integer" },
    ],
  },
  {
    name: "allowed-value list",
    argv: ["entra", "allowed-value", "list", "--definition", CSA_DEF_ID, "--profile", "soc"],
    key: "allowedValues",
    noun: "allowed values",
    rows: () => [
      { id: ID_A, isActive: true },
      { id: ID_B, isActive: false },
    ],
  },
  {
    name: "data-policy-operation list",
    argv: ["entra", "data-policy-operation", "list", "--profile", "soc"],
    key: "dataPolicyOperations",
    noun: "data-policy operations",
    rows: () => [
      { id: ID_A, status: "complete", userId: "user-1", submittedDateTime: "2026-09-30T00:00:00Z" },
      { id: ID_B, status: "running", userId: "user-2", submittedDateTime: "2026-10-02T00:00:00Z" },
    ],
  },
  {
    name: "delegated-admin-customer list",
    argv: ["entra", "delegated-admin-customer", "list", "--profile", "soc"],
    key: "delegatedAdminCustomers",
    noun: "customers",
    rows: () => [
      { id: ID_A, displayName: "Contoso Inc", tenantId: ID_B },
      { id: ID_B, displayName: "Fabrikam Inc", tenantId: ID_A },
    ],
  },
  {
    name: "delegated-admin-relationship list",
    argv: ["entra", "delegated-admin-relationship", "list", "--profile", "soc"],
    key: "delegatedAdminRelationships",
    noun: "relationships",
    rows: () => [
      { id: ID_A, displayName: "Contoso admin relationship", status: "active" },
      { id: ID_B, displayName: "Fabrikam admin relationship", status: "approvalPending" },
    ],
  },
  {
    name: "delegated-admin-relationship list-access-assignments",
    argv: ["entra", "delegated-admin-relationship", "list-access-assignments", "--id", REL_ID, "--profile", "soc"],
    key: "delegatedAdminAccessAssignments",
    noun: "access assignments",
    rows: () => [
      { id: ID_A, status: "active" },
      { id: ID_B, status: "pending" },
    ],
  },
  {
    name: "delegated-admin-relationship list-operations",
    argv: ["entra", "delegated-admin-relationship", "list-operations", "--id", REL_ID, "--profile", "soc"],
    key: "delegatedAdminRelationshipOperations",
    noun: "relationship operations",
    rows: () => [
      { id: ID_A, operationType: "delegatedAdminAccessAssignmentUpdate", status: "succeeded", lastModifiedDateTime: "2022-02-09T22:17:43Z" },
      { id: ID_B, operationType: "delegatedAdminAccessAssignmentUpdate", status: "running", lastModifiedDateTime: "2022-02-10T22:17:43Z" },
    ],
  },
  {
    name: "delegated-admin-relationship list-requests",
    argv: ["entra", "delegated-admin-relationship", "list-requests", "--id", REL_ID, "--profile", "soc"],
    key: "delegatedAdminRelationshipRequests",
    noun: "relationship requests",
    rows: () => [
      { id: ID_A, action: "lockForApproval", status: "succeeded", lastModifiedDateTime: "2022-02-09T22:17:43Z" },
      { id: ID_B, action: "terminate", status: "running", lastModifiedDateTime: "2022-02-10T22:17:43Z" },
    ],
  },
  {
    name: "delegated-admin-customer list-service-management-details",
    argv: ["entra", "delegated-admin-customer", "list-service-management-details", "--id", CUST_ID, "--profile", "soc"],
    key: "delegatedAdminServiceManagementDetails",
    noun: "service-management details",
    rows: () => [
      { id: ID_A, serviceManagementUrl: "https://lighthouse.microsoft.com", serviceName: "Microsoft 365 Lighthouse" },
      { id: ID_B, serviceName: "Teams" },
    ],
  },
  {
    name: "entitlement catalog list",
    argv: ["entra", "entitlement", "catalog", "list", "--profile", "soc"],
    key: "catalogs",
    noun: "catalogs",
    rows: () => [
      { id: ID_A, displayName: "General", state: "published", catalogType: "userManaged" },
      { id: ID_B, displayName: "Service default", state: "published", catalogType: "serviceDefault" },
    ],
  },
  {
    name: "entitlement access-package list",
    argv: ["entra", "entitlement", "access-package", "list", "--profile", "soc"],
    key: "accessPackages",
    noun: "access packages",
    rows: () => [
      { id: ID_A, displayName: "Engineering bundle", isHidden: false },
      { id: ID_B, displayName: "Quiet bundle", isHidden: true },
    ],
  },
  {
    name: "entitlement assignment-policy list",
    argv: ["entra", "entitlement", "assignment-policy", "list", "--access-package", PKG_ID, "--profile", "soc"],
    key: "assignmentPolicies",
    noun: "assignment policies",
    rows: () => [
      { id: ID_A, displayName: "Engineering requests", allowedTargetScope: "specificDirectoryUsers" },
      { id: ID_B, displayName: "Open requests", allowedTargetScope: "allMemberUsers" },
    ],
  },
  {
    name: "entitlement resource-role-scope list",
    argv: ["entra", "entitlement", "resource-role-scope", "list", "--access-package", PKG_ID, "--profile", "soc"],
    key: "resourceRoleScopes",
    noun: "resource-role scopes",
    rows: () => [
      { id: ID_A },
      { id: ID_B },
    ],
  },
  {
    name: "entitlement assignment list",
    argv: ["entra", "entitlement", "assignment", "list", "--profile", "soc"],
    key: "assignments",
    noun: "assignments",
    rows: () => [
      { id: ID_A, state: "delivered" },
      { id: ID_B, state: "expired" },
    ],
  },
  {
    name: "entitlement assignment-request list",
    argv: ["entra", "entitlement", "assignment-request", "list", "--profile", "soc"],
    key: "assignmentRequests",
    noun: "assignment requests",
    rows: () => [
      { id: ID_A, requestType: "userAdd", state: "delivered" },
      { id: ID_B, requestType: "adminAdd", state: "pendingApproval" },
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

// Catalogue enumeration: every entra list leaf renders uniform totals. A
// leaf counts as wired when its flag-free argv path appears in the table
// above; the pending set names unwired leaves owned by concurrent pieces.
// Any other unwired leaf fails this test.
const pendingTotals = new Set([]);

test("catalogue list leaves are all wired or explicitly pending", () => {
  // No list argv carries positionals: every value rides behind a flag, so
  // the catalogue path is the flag-free argv prefix.
  const wiredPaths = new Set(wired.map(entry => {
    const flagAt = entry.argv.findIndex(token => token.startsWith("--"));
    return entry.argv.slice(0, flagAt === -1 ? entry.argv.length : flagAt).join(" ");
  }));
  const listPaths = LEAVES.map(leaf => leaf.path)
    .filter(path => path.startsWith("entra ") && path.split(" ").at(-1).startsWith("list"));
  const covered = new Set([...wiredPaths, ...pendingTotals]);
  const missing = listPaths.filter(path => !covered.has(path));
  const extra = [...covered].filter(path => !listPaths.includes(path));
  assert.deepEqual(missing, [], `unwired list leaves: ${missing.join(", ")}`);
  assert.deepEqual(extra, [], `stale wiring-test paths: ${extra.join(", ")}`);
});
