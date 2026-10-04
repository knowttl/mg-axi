import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { resolve, sep } from "node:path";
import { AxiError } from "axi-sdk-js";

// Use INV-01 identities and evidence directly; discovery never grants access.
const schema = JSON.parse(readFileSync(new URL("../inventory/schema.json", import.meta.url), "utf8"));
export const API_VERSIONS: string[] = schema.$defs.version.enum;
export const DESCRIPTION = "Inspect Microsoft Graph capabilities, read-only by default";

type Flag = { description: string; value?: string; default?: string; required?: true };
type Leaf = { path: string; description: string; flags: Record<string, Flag>; examples: string[]; operation?: string; positional?: { name: string; description: string } };
const common = {
  profile: { value: "name", description: "Select a configured profile" },
  "api-version": { value: API_VERSIONS.join("|"), default: "v1.0", description: "Explicit API version; no fallback" },
};
const PARSE_MODES = ["delegated", "application"];
const logRead = {
  since: { value: "iso-time", description: "Required for a new query: earliest instant bounding the server time range; resume with --cursor instead of repeating it" },
  until: { value: "iso-time", description: "Latest instant bounding the server time range; must be after --since" },
  filter: { value: "odata-filter", description: "OData $filter combined with the time bounds; unsupported combinations fail before credentials" },
  select: { value: "comma-separated-properties", description: "Request server properties; CA policy detail additionally needs a policy permission in both modes and a supported CA-data role for delegated access; request delegated policy access with --scopes https://graph.microsoft.com/AuditLog.Read.All,https://graph.microsoft.com/Policy.Read.All" },
  fields: { value: "comma-separated-properties", description: "Project returned rows locally; every field must be fetched via the default or --select set" },
  full: { description: "Show complete text values without truncation; never lifts redaction, row caps or time bounds" },
  cursor: { value: "opaque-cursor|-", description: "Resume a capped collection losslessly; - reads the token from stdin (16 MB ceiling for either input); repeat the original query flags or omit them" },
  scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to AuditLog.Read.All" },
};
const userRead = {
  filter: { value: "odata-filter", description: "OData $filter passed to Graph; unsupported combinations fail before credentials" },
  select: { value: "comma-separated-properties", description: "Request server properties; richer properties need User.Read.All, accountEnabled needs User.EnableDisableAccount.All" },
  fields: { value: "comma-separated-properties", description: "Project returned rows locally; every field must be fetched via the default or --select set" },
  full: { description: "Show complete text values without truncation; never lifts redaction or row caps" },
  cursor: { value: "opaque-cursor", description: "Resume a capped collection losslessly; repeat the original query flags or omit them" },
  scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to User.Read.All" },
};
const groupRead = {
  filter: { value: "odata-filter", description: "OData $filter passed to Graph with ConsistencyLevel eventual; unsupported combinations fail before credentials" },
  select: { value: "comma-separated-properties", description: "Request server properties; richer group properties need Group.Read.All" },
  fields: { value: "comma-separated-properties", description: "Project returned rows locally; every field must be fetched via the default or --select set" },
  full: { description: "Show complete text values without truncation; never lifts redaction or row caps" },
  cursor: { value: "opaque-cursor", description: "Resume a capped collection losslessly; repeat the original query flags or omit them" },
  scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to GroupMember.Read.All, hidden members need Member.Read.Hidden" },
};
const appRead = {
  filter: { value: "odata-filter", description: "OData $filter passed to Graph with ConsistencyLevel eventual; unsupported combinations fail before credentials" },
  select: { value: "comma-separated-properties", description: "Request server properties; credential fields carry expiry metadata only, never secret values" },
  fields: { value: "comma-separated-properties", description: "Project returned rows locally; every field must be fetched via the default or --select set" },
  full: { description: "Show complete text values without truncation; never lifts redaction or row caps" },
  cursor: { value: "opaque-cursor", description: "Resume a capped collection losslessly; repeat the original query flags or omit them" },
  scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to Application.Read.All" },
};
const ownerRead = {
  filter: { value: "odata-filter", description: "OData $filter passed to Graph with ConsistencyLevel eventual; unsupported combinations fail before credentials" },
  select: { value: "comma-separated-properties", description: "Request server properties from id, displayName, mail; richer owner fields need single-object reads" },
  fields: { value: "comma-separated-properties", description: "Project returned rows locally; every field must be fetched via the default or --select set" },
  full: { description: "Show complete text values without truncation; never lifts redaction or row caps" },
  cursor: { value: "opaque-cursor", description: "Resume a capped collection losslessly; repeat the original query flags or omit them" },
  scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to Application.Read.All" },
};
const roleRead = {
  filter: { value: "odata-filter", description: "OData $filter passed to Graph as plain $filter; unsupported combinations fail before credentials" },
  select: { value: "comma-separated-properties", description: "Request server properties from the reviewed role property set" },
  fields: { value: "comma-separated-properties", description: "Project returned rows locally; every field must be fetched via the default or --select set" },
  full: { description: "Show complete text values without truncation; never lifts redaction or row caps" },
  cursor: { value: "opaque-cursor", description: "Resume a capped collection losslessly; repeat the original query flags or omit them" },
  scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to the operation's least-privileged read scope" },
};
export const LEAVES: Leaf[] = [
  { path: "home", description: "Show local profile status without authenticating", flags: { profile: common.profile }, examples: ["mg-axi", "mg-axi home --profile soc"] },
  { path: "profile create", description: "Create a dedicated-app profile without signing in", flags: {
    name: { value: "name", required: true, description: "New profile name; existing identities cannot be overwritten" },
    tenant: { value: "tenant-id", required: true, description: "Explicit workforce tenant UUID" },
    client: { value: "client-id", required: true, description: "Organization-owned application UUID" },
    cloud: { value: "commercial", required: true, description: "Explicit cloud; only commercial is supported" },
    mode: { value: "delegated|application", default: "delegated", description: "Delegated analyst sign-in or application client credentials" },
    "allow-device-code": { description: "Delegated only: opt in only when permitted by organization policy; never enables fallback" },
    "certificate-thumbprint": { value: "40-hex-digits", description: "Application only: certificate thumbprint; the private key stays in protected storage" },
    federated: { description: "Application only: workload federation via the AZURE_FEDERATED_TOKEN_FILE assertion source" },
  }, examples: ["mg-axi profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial", "mg-axi profile create --name batch --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --federated", "mg-axi profile create --help"] },
  { path: "profile list", description: "List configured profiles without accessing credentials", flags: {}, examples: ["mg-axi profile list", "mg-axi profile list --help"] },
  { path: "profile show", description: "Show profile policy and credential reference, never credentials", flags: {
    profile: common.profile,
  }, examples: ["mg-axi profile show --profile soc", "mg-axi profile show"] },
  { path: "login", description: "Explicit delegated login; browser by default, no automatic device-code fallback", flags: {
    profile: common.profile,
    method: { value: "browser|device-code", default: "browser", description: "Device code additionally requires profile opt-in" },
    scopes: { value: "comma-separated-Graph-scopes", required: true, description: "Explicit delegated permissions using full https://graph.microsoft.com/ scope names" },
  }, examples: ["mg-axi login --profile soc --scopes https://graph.microsoft.com/User.Read", "mg-axi login --profile soc --method device-code --scopes https://graph.microsoft.com/User.Read"] },
  { path: "entra user list", description: "List users with basic properties (id, displayName, userPrincipalName, mail)", operation: "GET:/users", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: userRead.filter,
    select: userRead.select,
    fields: userRead.fields,
    full: userRead.full,
    cursor: userRead.cursor,
    scopes: userRead.scopes,
  }, examples: ["mg-axi entra user list --profile soc", "mg-axi entra user list --profile soc --limit 10", "mg-axi entra user list --profile soc --filter \"accountEnabled eq true\" --select id,displayName,accountEnabled"] },
  { path: "entra user show", description: "Show one user with richer default properties", operation: "GET:/users/{user-id}", flags: {
    ...common, id: { value: "user-id-or-upn", required: true, description: "User object ID or UPN" },
    select: userRead.select,
    fields: userRead.fields,
    full: userRead.full,
    scopes: userRead.scopes,
  }, examples: ["mg-axi entra user show --id <user-id-or-upn> --profile soc", "mg-axi entra user show --id <user-id-or-upn> --profile soc --full"] },
  { path: "entra group list", description: "List groups with compact properties (id, displayName, mail, groupTypes)", operation: "GET:/groups", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: groupRead.filter,
    select: groupRead.select,
    fields: groupRead.fields,
    full: groupRead.full,
    cursor: groupRead.cursor,
    scopes: groupRead.scopes,
  }, examples: ["mg-axi entra group list --profile soc", "mg-axi entra group list --profile soc --limit 10", "mg-axi entra group list --profile soc --filter \"securityEnabled eq true\" --select id,displayName,isAssignableToRole"] },
  { path: "entra group show", description: "Show one group with richer default properties including isAssignableToRole", operation: "GET:/groups/{group-id}", flags: {
    ...common, id: { value: "group-id", required: true, description: "Group object ID" },
    select: groupRead.select,
    fields: groupRead.fields,
    full: groupRead.full,
    scopes: groupRead.scopes,
  }, examples: ["mg-axi entra group show --id <group-id> --profile soc", "mg-axi entra group show --id <group-id> --profile soc --full"] },
  { path: "entra group member list", description: "List direct group members; --transitive flattens nested membership. Hidden members are omitted without Member.Read.Hidden; v1.0 may omit service principals", operation: "GET:/groups/{group-id}/members", flags: {
    ...common, group: { value: "group-id", required: true, description: "Group object ID whose members are listed" },
    transitive: { description: "List the flat transitive closure instead of direct members" },
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: groupRead.filter,
    select: { value: "comma-separated-properties", description: "Request server properties from id, displayName, mail; richer member fields need single-object reads" },
    fields: { value: "comma-separated-properties", description: "Project returned rows locally; every field must be fetched via the default or --select set" },
    full: groupRead.full,
    cursor: groupRead.cursor,
    scopes: groupRead.scopes,
  }, examples: ["mg-axi entra group member list --group <group-id> --profile soc", "mg-axi entra group member list --group <group-id> --transitive --profile soc"] },
  { path: "entra group member-of list", description: "List groups the group is a member of; --transitive flattens nested membership. Hidden memberships are omitted without Member.Read.Hidden", operation: "GET:/groups/{group-id}/memberOf", flags: {
    ...common, group: { value: "group-id", required: true, description: "Group object ID whose memberships are listed" },
    transitive: { description: "List the flat transitive closure instead of direct memberships" },
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: groupRead.filter,
    select: { value: "comma-separated-properties", description: "Request server properties from id, displayName, mail; richer member fields need single-object reads" },
    fields: { value: "comma-separated-properties", description: "Project returned rows locally; every field must be fetched via the default or --select set" },
    full: groupRead.full,
    cursor: groupRead.cursor,
    scopes: groupRead.scopes,
  }, examples: ["mg-axi entra group member-of list --group <group-id> --profile soc", "mg-axi entra group member-of list --group <group-id> --transitive --profile soc"] },
  { path: "entra directory-role list", description: "List activated directory roles (id, displayName, description, roleTemplateId); roles appear only after activation", operation: "GET:/directoryRoles", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: roleRead.filter,
    select: roleRead.select,
    fields: roleRead.fields,
    full: roleRead.full,
    cursor: roleRead.cursor,
    scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to RoleManagement.Read.Directory; delegated callers also need a supported directory role" },
  }, examples: ["mg-axi entra directory-role list --profile soc", "mg-axi entra directory-role list --profile soc --filter \"displayName eq 'Global Administrator'\""] },
  { path: "entra directory-role show", description: "Show one activated directory role with richer default properties", operation: "GET:/directoryRoles/{directoryRole-id}", flags: {
    ...common, id: { value: "role-id", required: true, description: "Activated directory-role object ID" },
    select: roleRead.select,
    fields: roleRead.fields,
    full: roleRead.full,
    scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to RoleManagement.Read.Directory; delegated callers also need a supported directory role" },
  }, examples: ["mg-axi entra directory-role show --id <role-id> --profile soc", "mg-axi entra directory-role show --id <role-id> --profile soc --full"] },
  { path: "entra role-assignment list", description: "List direct role assignments (id, principalId, roleDefinitionId, directoryScopeId); activated PIM eligibility is not listed here", operation: "GET:/roleManagement/directory/roleAssignments", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: roleRead.filter,
    select: roleRead.select,
    fields: roleRead.fields,
    full: roleRead.full,
    cursor: roleRead.cursor,
    scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to RoleManagement.Read.Directory; delegated callers also need Directory Readers, Global Reader or Privileged Role Administrator" },
  }, examples: ["mg-axi entra role-assignment list --profile soc", "mg-axi entra role-assignment list --profile soc --filter \"principalId eq '<principal-id>'\""] },
  { path: "entra pim eligible list", description: "List PIM-eligible assignments, which are not active; activation is a PIM workflow outside these reads", operation: "GET:/roleManagement/directory/roleEligibilityScheduleInstances", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: roleRead.filter,
    select: roleRead.select,
    fields: roleRead.fields,
    full: roleRead.full,
    cursor: roleRead.cursor,
    scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to RoleEligibilitySchedule.Read.Directory; delegated callers also need a supported PIM read role; PIM needs P2 or ID Governance" },
  }, examples: ["mg-axi entra pim eligible list --profile soc", "mg-axi entra pim eligible list --profile soc --filter \"roleDefinitionId eq '<role-definition-id>'\""] },
  { path: "entra pim active list", description: "List active assignments: directly assigned plus activated eligible (see assignmentType); the direct-only view is role-assignment list", operation: "GET:/roleManagement/directory/roleAssignmentScheduleInstances", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: roleRead.filter,
    select: roleRead.select,
    fields: roleRead.fields,
    full: roleRead.full,
    cursor: roleRead.cursor,
    scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit Graph scopes using full https://graph.microsoft.com/ names; defaults to RoleAssignmentSchedule.Read.Directory; delegated callers also need a supported PIM read role; PIM needs P2 or ID Governance" },
  }, examples: ["mg-axi entra pim active list --profile soc", "mg-axi entra pim active list --profile soc --filter \"assignmentType eq 'Activated'\""] },
  { path: "entra sign-in list", description: "List sign-ins in a bounded time window (AuditLog.Read.All; delegated callers also need Global Reader, Reports Reader, Security Administrator, Security Operator or Security Reader; conservative P1/P2 deployment prerequisite)", operation: "GET:/auditLogs/signIns", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    since: { ...logRead.since, description: "Required for a new query (resume with --cursor instead): earliest createdDateTime bounding the server time range" },
    until: logRead.until,
    filter: logRead.filter,
    select: logRead.select,
    fields: logRead.fields,
    full: logRead.full,
    cursor: logRead.cursor,
    scopes: logRead.scopes,
  }, examples: ["mg-axi entra sign-in list --profile soc --since 2026-09-01T00:00:00Z", "mg-axi entra sign-in list --profile soc --since 2026-09-01T00:00:00Z --filter \"status/errorCode ne 0\" --limit 10", "mg-axi entra sign-in list --profile soc --since 2026-09-01T00:00:00Z --until 2026-09-08T00:00:00Z --all"] },
  { path: "entra sign-in show", description: "Show one sign-in with the full reviewed property set; absent CA policy detail reports its required policy permission and delegated role instead of an empty value", operation: "GET:/auditLogs/signIns/{signIn-id}", flags: {
    ...common, id: { value: "sign-in-id", required: true, description: "Sign-in object ID" },
    select: logRead.select,
    fields: logRead.fields,
    full: logRead.full,
    scopes: logRead.scopes,
  }, examples: ["mg-axi entra sign-in show --id <sign-in-id> --profile soc", "mg-axi entra sign-in show --id <sign-in-id> --profile soc --full"] },
  { path: "entra directory-audit list", description: "List directory audits in a bounded time window (AuditLog.Read.All; delegated callers also need Reports Reader, Security Administrator or Security Reader; conservative P1/P2 deployment prerequisite)", operation: "GET:/auditLogs/directoryAudits", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    since: { ...logRead.since, description: "Required for a new query (resume with --cursor instead): earliest activityDateTime bounding the server time range" },
    until: logRead.until,
    filter: logRead.filter,
    select: logRead.select,
    fields: logRead.fields,
    full: logRead.full,
    cursor: logRead.cursor,
    scopes: logRead.scopes,
  }, examples: ["mg-axi entra directory-audit list --profile soc --since 2026-09-01T00:00:00Z", "mg-axi entra directory-audit list --profile soc --since 2026-09-01T00:00:00Z --filter \"category eq 'UserManagement'\"", "mg-axi entra directory-audit list --profile soc --since 2026-09-01T00:00:00Z --all"] },
  { path: "entra directory-audit show", description: "Show one directory audit with the full reviewed property set", operation: "GET:/auditLogs/directoryAudits/{directoryAudit-id}", flags: {
    ...common, id: { value: "directory-audit-id", required: true, description: "Directory audit object ID" },
    select: logRead.select,
    fields: logRead.fields,
    full: logRead.full,
    scopes: logRead.scopes,
  }, examples: ["mg-axi entra directory-audit show --id <directory-audit-id> --profile soc", "mg-axi entra directory-audit show --id <directory-audit-id> --profile soc --full"] },
  { path: "entra application list", description: "List applications with compact properties (id, appId, displayName); appId is the client ID, distinct from the object id", operation: "GET:/applications", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: appRead.filter,
    select: appRead.select,
    fields: appRead.fields,
    full: appRead.full,
    cursor: appRead.cursor,
    scopes: appRead.scopes,
  }, examples: ["mg-axi entra application list --profile soc", "mg-axi entra application list --profile soc --limit 10", "mg-axi entra application list --profile soc --filter \"startswith(displayName,'A')\""] },
  { path: "entra application show", description: "Show one application with credential expiry metadata only; never any secret value, hint or key material", operation: "GET:/applications/{application-id}", flags: {
    ...common, id: { value: "application-object-id", required: true, description: "Application object ID; appId is the client ID, not the object ID" },
    select: appRead.select,
    fields: appRead.fields,
    full: appRead.full,
    scopes: appRead.scopes,
  }, examples: ["mg-axi entra application show --id <application-object-id> --profile soc", "mg-axi entra application show --id <application-object-id> --profile soc --full"] },
  { path: "entra service-principal list", description: "List service principals with compact properties (id, appId, displayName); appId is the client ID, distinct from the object id", operation: "GET:/servicePrincipals", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: appRead.filter,
    select: appRead.select,
    fields: appRead.fields,
    full: appRead.full,
    cursor: appRead.cursor,
    scopes: appRead.scopes,
  }, examples: ["mg-axi entra service-principal list --profile soc", "mg-axi entra service-principal list --profile soc --limit 10"] },
  { path: "entra service-principal show", description: "Show one service principal with richer default properties including servicePrincipalType", operation: "GET:/servicePrincipals/{servicePrincipal-id}", flags: {
    ...common, id: { value: "service-principal-object-id", required: true, description: "Service principal object ID; appId is the client ID, not the object ID" },
    select: appRead.select,
    fields: appRead.fields,
    full: appRead.full,
    scopes: appRead.scopes,
  }, examples: ["mg-axi entra service-principal show --id <service-principal-object-id> --profile soc"] },
  { path: "entra application owner list", description: "List owners of one application; rows carry @odata.type naming the owner kind", operation: "GET:/applications/{application-id}/owners", flags: {
    ...common, application: { value: "application-object-id", required: true, description: "Application object ID whose owners are listed" },
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: ownerRead.filter,
    select: ownerRead.select,
    fields: ownerRead.fields,
    full: ownerRead.full,
    cursor: ownerRead.cursor,
    scopes: ownerRead.scopes,
  }, examples: ["mg-axi entra application owner list --application <application-object-id> --profile soc"] },
  { path: "entra service-principal owner list", description: "List owners of one service principal; rows carry @odata.type naming the owner kind", operation: "GET:/servicePrincipals/{servicePrincipal-id}/owners", flags: {
    ...common, "service-principal": { value: "service-principal-object-id", required: true, description: "Service principal object ID whose owners are listed" },
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; the remainder is buffered into an opaque cursor, never discarded; incompatible with --all" },
    all: { description: "Follow pages within request, byte and deadline budgets" },
    filter: ownerRead.filter,
    select: ownerRead.select,
    fields: ownerRead.fields,
    full: ownerRead.full,
    cursor: ownerRead.cursor,
    scopes: ownerRead.scopes,
  }, examples: ["mg-axi entra service-principal owner list --service-principal <service-principal-object-id> --profile soc"] },
  { path: "api get", description: "Reviewed read-only raw Graph GET (API-01, v1.0 only): users, groups, conditional access, authentication methods, audit/sign-in, risk, apps, roles/PIM, devices and administrative units; unreviewed, secret-value, mail/file-content, beta and write routes are refused before credentials", positional: { name: "path", description: "Server-relative Graph path, e.g. /users" }, flags: {
    ...common,
    odata: { value: "k=v&k2=v2", description: "OData query reviewed per route ($select/$filter/$top/$orderby on collections; $select on singles); defaults to reviewed fields" },
    cursor: { value: "token|-", description: "Resume a partial collection; - reads the token from stdin (16 MB ceiling for either input); use the same path, profile and scopes, and omit --odata to reuse its query" },
    scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit full https://graph.microsoft.com/ scope names; application profiles use the .default audience" },
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; use --all to follow pages within budget" },
    all: { description: "Follow @odata.nextLink pages within the request budget" },
    full: { description: "Disable 4000-character string truncation; never disables redaction or row caps" },
  }, examples: ["mg-axi api get /users --scopes https://graph.microsoft.com/User.Read.All", "mg-axi api get /groups --odata '$filter=securityEnabled eq true&$top=5' --scopes https://graph.microsoft.com/GroupMember.Read.All", "mg-axi api get /users --cursor - --all --scopes https://graph.microsoft.com/User.Read.All < cursor.txt"] },
];

export function leafHelp(leaf: Leaf): string {
  return [
    `mg-axi ${leaf.path}${leaf.positional ? ` <${leaf.positional.name}>` : ""}`, leaf.description,
    ...(leaf.positional ? [`<${leaf.positional.name}>: ${leaf.positional.description}`] : []),
    ...Object.entries(leaf.flags).map(([name, flag]) => `--${name}${flag.value ? ` <${flag.value}>` : ""}: ${flag.description}${flag.default ? ` (default: ${flag.default})` : ""}${flag.required ? " (required)" : ""}`),
    "--help: Show this reference",
    ...leaf.examples,
  ].join("\n");
}

export const TOP_LEVEL_HELP = [DESCRIPTION, ...LEAVES.map(leaf => leafHelp(leaf)), "-v, -V, --version: Print the bare version"].join("\n\n");

export function resolveCommand(argv: string[]): { leaf: Leaf; flags: Record<string, string | boolean>; positional?: string } {
  const words = argv.slice(0, argv.findIndex(arg => arg.startsWith("-")) < 0 ? argv.length : argv.findIndex(arg => arg.startsWith("-")));
  const path = words.join(" ") || "home";
  let leaf = LEAVES.find(item => item.path === path);
  let positional: string | undefined;
  let flagStart = words.length;
  if (!leaf) {
    if (words[0] !== "api") throw new AxiError("unknown or incomplete command", "VALIDATION_ERROR", [TOP_LEVEL_HELP]);
    leaf = LEAVES.find(item => item.path === "api get")!;
    const help = leafHelp(leaf);
    const verb = words[1];
    if (verb === undefined) throw new AxiError("unknown or incomplete command", "VALIDATION_ERROR", [TOP_LEVEL_HELP]);
    if (verb.toLowerCase() !== "get") throw new AxiError(`mg-axi api serves reviewed GET reads only; got api ${verb}`, "VALIDATION_ERROR", [help]);
    if (words.length < 3 || words[2]!.startsWith("-")) throw new AxiError("missing path for `mg-axi api get`", "VALIDATION_ERROR", ["Example: mg-axi api get /users --scopes https://graph.microsoft.com/User.Read.All", help]);
    if (words.length > 3) throw new AxiError(`unexpected argument \`${words[3]}\` for \`mg-axi api get\``, "VALIDATION_ERROR", ["Pass one path before flags: mg-axi api get <path> [--odata 'k=v']", help]);
    positional = words[2];
    flagStart = 3;
  }
  const help = leafHelp(leaf);
  const fail = (message: string): never => { throw new AxiError(message, "VALIDATION_ERROR", [help]); };
  const flags: Record<string, string | boolean> = Object.create(null);
  for (let i = flagStart; i < argv.length; i++) {
    const arg = argv[i]!;
    const match = /^--([a-z-]+)(?:=(.*))?$/.exec(arg);
    if (!match) fail("unexpected argument or short flag");
    const name = match![1]!;
    const flag = Object.hasOwn(leaf.flags, name) ? leaf.flags[name] : name === "help" ? { description: "Help" } : undefined;
    if (!flag) fail(leaf.path === "api get" && name === "query" ? "--query is reserved for output queries; use --odata for server OData parameters" : `unknown flag --${name}`);
    if (Object.hasOwn(flags, name)) fail(`duplicate flag --${name}`);
    if (flag!.value) {
      const value = match![2] ?? argv[++i];
      if (!value?.trim() || (value.startsWith("-") && !(name === "cursor" && value === "-"
        && (leaf.path === "api get" || leaf.path === "entra sign-in list" || leaf.path === "entra directory-audit list")))) fail(`--${name} requires a non-empty value`);
      flags[name] = value!;
    } else {
      if (match![2] !== undefined) fail(`--${name} does not take a value`);
      flags[name] = true;
    }
  }
  if (flags["api-version"] && !API_VERSIONS.includes(String(flags["api-version"]))) fail(`--api-version must be ${API_VERSIONS.join(" or ")}`);
  if (flags.mode && !PARSE_MODES.includes(String(flags.mode))) fail("--mode must be delegated or application");
  if (flags.limit && (!/^[1-9]\d*$/.test(String(flags.limit)) || !Number.isSafeInteger(Number(flags.limit)))) fail("--limit must be a positive safe integer");
  if (flags.limit && flags.all) fail("--limit and --all cannot be combined");
  if (!flags.help) for (const [name, flag] of Object.entries(leaf.flags)) if (flag.required && !flags[name]) fail(`--${name} is required`);
  if (!flags.help && leaf.positional && positional === undefined) fail(`missing ${leaf.positional.name} for \`mg-axi ${leaf.path}\``);
  return { leaf, flags, positional };
}

export function operationFor(leaf: Leaf, version: string) {
  const inventory = JSON.parse(readFileSync(new URL("../inventory/operations.json", import.meta.url), "utf8"));
  return inventory.operations.find((row: { id: string }) => row.id === `${version}:${leaf.operation}`);
}

// A useful local status, without treating unavailable tenant summaries as zeros.
export function home() {
  const bin = resolve(process.argv[1]!);
  return {
    bin: bin.startsWith(`${homedir()}${sep}`) ? `~${bin.slice(homedir().length)}` : bin,
    description: DESCRIPTION,
    profile: "unavailable: no profile configured",
    tenant: "unavailable: no tenant selected",
    domains: [{ name: "entra", status: "scheduled", summary: "Tenant summaries await Graph execution" }],
    help: ["mg-axi entra user list --help", "mg-axi entra user show --help", "mg-axi entra group list --help", "mg-axi entra group member list --help", "mg-axi entra application list --help", "mg-axi entra service-principal list --help", "mg-axi api get --help"],
  };
}
