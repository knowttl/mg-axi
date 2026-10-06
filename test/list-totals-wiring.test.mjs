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
const GROUP_PIM_FILTER = `groupId eq '${ID_B}'`;

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
