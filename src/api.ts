import { AxiError } from "axi-sdk-js";
import { ApplicationAuth } from "./app-auth.js";
import { DelegatedAuth } from "./auth.js";
import { KNOWN_DECISION_FIELDS, KNOWN_DEFINITION_FIELDS, KNOWN_INSTANCE_FIELDS } from "./entra-access-reviews.js";
import { KNOWN_BRANDING_FIELDS, KNOWN_ORGANIZATION_FIELDS } from "./entra-organization.js";
import { KNOWN_CONTRACT_FIELDS } from "./entra-contracts.js";
import { encodeGraphPathSegment, GraphSession, resolveSessionOperation, type GraphTransport } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// API-01: the reviewed read-only raw Graph surface.
//
// The discovery inventory (inventory/operations.json) stays discovery-only: it
// carries no reviewed-raw-read rows and must be regenerated, never edited. The
// review therefore lives here, in the implementation catalogue: exact v1.0
// inventory row ids, collection/singleton kind, per-route query keys, reviewed
// $select fields, and sourced access constraints. Access choices come from the
// representative SOC contracts in docs/graph-coverage.md; each entry cites its
// upstream operation documentation (retained verbatim from its inventory row,
// except authentication methods where upstream supplies no externalDocs and the
// cited method documentation is used instead). Group permission choices were
// rechecked against current upstream documentation on REVIEWED_ON: group-list
// now marks a write permission least-privileged while listing
// GroupMember.Read.All and Group.Read.All as supported read alternatives, so
// the review records the read choices rather than the first table cell.
//
// What this module never reviews: secret-value and credential-minting routes,
// mail/file content, beta routes, and anything but GET. Those fail closed in
// matchReviewed before any profile, credential or transport is touched. Pack,
// preview, /me and sensitive-area policy still runs inside the shared session,
// so raw reads cannot bypass named-command gating.

export const REVIEWED_ON = "2026-10-04";
const TRUNCATE_AT = 4000;

export interface ReviewedRawRoute {
  /** Inventory row id: version:GET:template, the exact binding resolveSessionOperation needs. */
  readonly id: string;
  readonly kind: "collection" | "single";
  /** Allowed query keys: a subset of the session allowlist, reviewed per route shape. */
  readonly query: readonly string[];
  /** Allowed $select fields: reviewed, non-secret server properties. */
  readonly fields: readonly string[];
  /** Supported read permission choices with delegated/application distinction. */
  readonly access: string;
  readonly note?: string;
  readonly warning?: string;
  /** Primary-source operation documentation, rechecked on REVIEWED_ON. */
  readonly sources: readonly string[];
}

// Collections support server filtering, ordering and paging hints. Advanced
// queries ($search, $count=true) need per-endpoint ConsistencyLevel review and
// stay unsupported here; $skiptoken is server paging state, paged via --all.
const COLLECTION_QUERY = ["$select", "$filter", "$top", "$orderby"] as const;
const SINGLE_QUERY = ["$select"] as const;

const USER_FIELDS = ["id", "displayName", "userPrincipalName", "mail", "accountEnabled", "userType", "jobTitle", "department", "officeLocation", "businessPhones", "mobilePhone", "createdDateTime"];
const GROUP_FIELDS = ["id", "displayName", "description", "mail", "mailEnabled", "mailNickname", "securityEnabled", "groupTypes", "visibility", "classification", "isAssignableToRole", "createdDateTime", "expirationDateTime", "renewedDateTime", "membershipRule", "membershipRuleProcessingState"];
const MEMBER_FIELDS = ["id", "displayName"];
const CA_POLICY_FIELDS = ["id", "displayName", "state", "createdDateTime", "modifiedDateTime", "conditions", "grantControls", "sessionControls"];
const NAMED_LOCATION_FIELDS = ["id", "displayName", "createdDateTime", "modifiedDateTime", "isTrusted", "ipRanges", "countriesAndRegions", "includeUnknownCountriesAndRegions"];
const AUTH_METHOD_FIELDS = ["id", "createdDateTime", "displayName", "phoneNumber", "phoneType", "smsSignInState", "emailAddress"];
const REGISTRATION_FIELDS = ["id", "userPrincipalName", "userDisplayName", "userType", "isAdmin", "isMfaRegistered", "isMfaCapable", "isPasswordlessCapable", "isSsprRegistered", "isSsprEnabled", "isSsprCapable", "userPreferredMethodForSecondaryAuthentication", "lastUpdatedDateTime"];
const SIGNIN_FIELDS = ["id", "createdDateTime", "userId", "userPrincipalName", "userDisplayName", "appDisplayName", "appId", "ipAddress", "location", "status", "conditionalAccessStatus", "riskDetail", "riskLevelAggregated", "riskLevelDuringSignIn", "riskState", "resourceDisplayName", "resourceId", "clientAppUsed"];
const AUDIT_FIELDS = ["id", "activityDateTime", "activityDisplayName", "category", "loggedByService", "operationType", "result", "resultReason", "correlationId", "initiatedBy", "targetResources"];
const RISKY_USER_FIELDS = ["id", "userPrincipalName", "userDisplayName", "riskDetail", "riskLastUpdatedDateTime", "riskLevel", "riskState"];
const RISK_DETECTION_FIELDS = ["id", "detectedDateTime", "activityDateTime", "userId", "userPrincipalName", "userDisplayName", "ipAddress", "location", "riskDetail", "riskLevel", "riskState", "riskEventType", "detectionTimingType", "lastUpdatedDateTime", "source"];
const APP_FIELDS = ["id", "appId", "displayName", "createdDateTime", "signInAudience", "publisherDomain", "keyCredentials", "passwordCredentials", "requiredResourceAccess", "web", "spa", "publicClient"];
const SP_FIELDS = ["id", "appId", "displayName", "servicePrincipalType", "accountEnabled", "appOwnerOrganizationId", "appRoleAssignmentRequired", "preferredSingleSignOnMode", "loginUrl"];
const GRANT_FIELDS = ["id", "clientId", "consentType", "principalId", "resourceId", "scope", "startTime", "expiryTime"];
const APP_ROLE_FIELDS = ["id", "appRoleId", "principalId", "principalDisplayName", "principalType", "resourceId", "resourceDisplayName", "createdDateTime"];
const ROLE_ASSIGNMENT_FIELDS = ["id", "principalId", "roleDefinitionId", "directoryScopeId", "appScopeId", "createdDateTime"];
const PIM_INSTANCE_FIELDS = ["id", "roleDefinitionId", "principalId", "assignmentType", "memberType", "startDateTime", "endDateTime", "activatedUsing"];
const DEVICE_FIELDS = ["id", "deviceId", "displayName", "operatingSystem", "operatingSystemVersion", "trustType", "isCompliant", "isManaged", "accountEnabled", "createdDateTime", "approximateLastSignInDateTime", "manufacturer", "model"];
const AU_FIELDS = ["id", "displayName", "description", "visibility", "membershipType", "membershipRule"];
const DOMAIN_FIELDS = ["id", "authenticationType", "availabilityStatus", "isAdminManaged", "isDefault", "isInitial", "isRoot", "isVerified", "supportedServices", "passwordValidityPeriodInDays", "passwordNotificationWindowInDays", "state"];
const DNS_RECORD_FIELDS = ["id", "isOptional", "label", "recordType", "supportedService", "ttl", "mailExchange", "preference", "canonicalName", "nameTarget", "port", "priority", "protocol", "service", "weight", "text"];

// The reviewed surface, exported for capability reporting (PACK-01) and tests.
export const REVIEWED_ROUTES: readonly ReviewedRawRoute[] = [
  { id: "v1.0:GET:/users", kind: "collection", query: COLLECTION_QUERY, fields: USER_FIELDS,
    access: "D User.ReadBasic.All for basics, User.Read.All for richer SOC data; A User.Read.All. Delegated callers pass one as --scopes.",
    sources: ["https://learn.microsoft.com/graph/api/user-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/users/{user-id}", kind: "single", query: SINGLE_QUERY, fields: USER_FIELDS,
    access: "D User.ReadBasic.All for basics, User.Read.All for richer SOC data; A User.Read.All. Delegated callers pass one as --scopes.",
    sources: ["https://learn.microsoft.com/graph/api/user-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/groups", kind: "collection", query: COLLECTION_QUERY, fields: GROUP_FIELDS,
    access: "D/A GroupMember.Read.All is a supported read choice; Group.Read.All for further properties. Delegated callers pass one as --scopes.",
    note: "Upstream marks a write permission least-privileged here; the review records the supported read alternatives instead.",
    sources: ["https://learn.microsoft.com/graph/api/group-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/groups/{group-id}", kind: "single", query: SINGLE_QUERY, fields: GROUP_FIELDS,
    access: "D/A GroupMember.Read.All is a supported read choice; Group.Read.All for further properties. Delegated callers pass one as --scopes.",
    sources: ["https://learn.microsoft.com/graph/api/group-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/groups/{group-id}/members", kind: "collection", query: COLLECTION_QUERY, fields: MEMBER_FIELDS,
    access: "D/A GroupMember.ReadBasic.All minimum; richer access GroupMember.Read.All. Delegated callers pass one as --scopes.",
    note: "Hidden membership needs Member.Read.Hidden; the server omits what the caller cannot see rather than failing. Richer member fields need single-object reads.",
    warning: "Microsoft Graph v1.0 may omit service principals from group members; completed pagination does not establish complete membership.",
    sources: ["https://learn.microsoft.com/graph/api/group-list-members?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identity/conditionalAccess/policies", kind: "collection", query: COLLECTION_QUERY, fields: CA_POLICY_FIELDS,
    access: "D/A Policy.Read.All. Delegated callers pass it as --scopes; a supported administrator role (for example Security Reader) is also required.",
    note: "CA needs P1; risk-based CA needs P2.",
    sources: ["https://learn.microsoft.com/graph/api/conditionalaccessroot-list-policies?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identity/conditionalAccess/policies/{conditionalAccessPolicy-id}", kind: "single", query: SINGLE_QUERY, fields: CA_POLICY_FIELDS,
    access: "D/A Policy.Read.All. Delegated callers pass it as --scopes; a supported administrator role (for example Security Reader) is also required.",
    note: "CA needs P1; risk-based CA needs P2.",
    sources: ["https://learn.microsoft.com/graph/api/conditionalaccesspolicy-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identity/conditionalAccess/namedLocations", kind: "collection", query: COLLECTION_QUERY, fields: NAMED_LOCATION_FIELDS,
    access: "D/A Policy.Read.All. Delegated callers pass it as --scopes; a supported administrator role (for example Security Reader) is also required.",
    sources: ["https://learn.microsoft.com/graph/api/conditionalaccessroot-list-namedlocations?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identity/conditionalAccess/namedLocations/{namedLocation-id}", kind: "single", query: SINGLE_QUERY, fields: NAMED_LOCATION_FIELDS,
    access: "D/A Policy.Read.All. Delegated callers pass it as --scopes; a supported administrator role (for example Security Reader) is also required.",
    sources: ["https://learn.microsoft.com/graph/api/countrynamedlocation-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/users/{user-id}/authentication/methods", kind: "collection", query: ["$select"], fields: AUTH_METHOD_FIELDS,
    access: "D/A UserAuthenticationMethod.Read.All for other users; delegated self UserAuthenticationMethod.Read. Dedicated administrator roles also apply.",
    note: "Targeted per-user inspection only; aggregate coverage belongs to the registration report, never to a scan across users.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/authentication-list-methods?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/users/{user-id}/authentication/methods/{authenticationMethod-id}", kind: "single", query: SINGLE_QUERY, fields: AUTH_METHOD_FIELDS,
    access: "D/A UserAuthenticationMethod.Read.All for other users; delegated self UserAuthenticationMethod.Read. Dedicated administrator roles also apply.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/authentication-list-methods?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/reports/authenticationMethods/userRegistrationDetails", kind: "collection", query: COLLECTION_QUERY, fields: REGISTRATION_FIELDS,
    access: "D/A AuditLog.Read.All. Delegated callers pass it as --scopes.",
    note: "The report does not cover disabled users; absence is not proof of no MFA.",
    sources: ["https://learn.microsoft.com/graph/api/authenticationmethodsroot-list-userregistrationdetails?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/auditLogs/signIns", kind: "collection", query: COLLECTION_QUERY, fields: SIGNIN_FIELDS,
    access: "D/A AuditLog.Read.All; CA details carry additional policy permission and role requirements. Delegated callers pass the read scope as --scopes.",
    note: "Conservative P1/P2 deployment prerequisite; retention, premium fields and the generating feature licence stay separate constraints.",
    sources: ["https://learn.microsoft.com/graph/api/signin-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/auditLogs/directoryAudits", kind: "collection", query: COLLECTION_QUERY, fields: AUDIT_FIELDS,
    access: "D/A AuditLog.Read.All. Delegated callers pass it as --scopes.",
    sources: ["https://learn.microsoft.com/graph/api/directoryaudit-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityProtection/riskyUsers", kind: "collection", query: COLLECTION_QUERY, fields: RISKY_USER_FIELDS,
    access: "D/A IdentityRiskyUser.Read.All. Delegated callers pass it as --scopes.",
    note: "Full investigation needs P2/Suite; limited results stay limited, never reinterpreted as empty.",
    sources: ["https://learn.microsoft.com/graph/api/riskyuser-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityProtection/riskyUsers/{riskyUser-id}", kind: "single", query: SINGLE_QUERY, fields: RISKY_USER_FIELDS,
    access: "D/A IdentityRiskyUser.Read.All. Delegated callers pass it as --scopes.",
    sources: ["https://learn.microsoft.com/graph/api/riskyuser-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityProtection/riskDetections", kind: "collection", query: COLLECTION_QUERY, fields: RISK_DETECTION_FIELDS,
    access: "D/A IdentityRiskEvent.Read.All. Delegated callers pass it as --scopes.",
    note: "The API permits P1 or P2; full details need P2.",
    sources: ["https://learn.microsoft.com/graph/api/riskdetection-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityProtection/riskDetections/{riskDetection-id}", kind: "single", query: SINGLE_QUERY, fields: RISK_DETECTION_FIELDS,
    access: "D/A IdentityRiskEvent.Read.All. Delegated callers pass it as --scopes.",
    sources: ["https://learn.microsoft.com/graph/api/riskdetection-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/applications", kind: "collection", query: COLLECTION_QUERY, fields: APP_FIELDS,
    access: "D/A Application.Read.All. Delegated callers pass it as --scopes.",
    note: "appId (client ID) is distinct from the object id. Credential fields carry expiry metadata only; GET never returns secret values and success redaction still applies.",
    sources: ["https://learn.microsoft.com/graph/api/application-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/applications/{application-id}", kind: "single", query: SINGLE_QUERY, fields: APP_FIELDS,
    access: "D/A Application.Read.All. Delegated callers pass it as --scopes.",
    note: "appId (client ID) is distinct from the object id. Credential fields carry expiry metadata only; GET never returns secret values and success redaction still applies.",
    sources: ["https://learn.microsoft.com/graph/api/application-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/servicePrincipals", kind: "collection", query: COLLECTION_QUERY, fields: SP_FIELDS,
    access: "D/A Application.Read.All. Delegated callers pass it as --scopes.",
    note: "appId (client ID) is distinct from the object id.",
    sources: ["https://learn.microsoft.com/graph/api/serviceprincipal-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/servicePrincipals/{servicePrincipal-id}", kind: "single", query: SINGLE_QUERY, fields: SP_FIELDS,
    access: "D/A Application.Read.All. Delegated callers pass it as --scopes.",
    note: "appId (client ID) is distinct from the object id.",
    sources: ["https://learn.microsoft.com/graph/api/serviceprincipal-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/servicePrincipals/{servicePrincipal-id}/appRoleAssignments", kind: "collection", query: COLLECTION_QUERY, fields: APP_ROLE_FIELDS,
    access: "D/A Application.Read.All. Delegated callers pass it as --scopes.",
    note: "Actual granted assignments; requested permissions are a separate contract.",
    sources: ["https://learn.microsoft.com/graph/api/serviceprincipal-list-approleassignments?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/oauth2PermissionGrants", kind: "collection", query: COLLECTION_QUERY, fields: GRANT_FIELDS,
    access: "D/A Directory.Read.All as the read choice. Delegated callers pass it as --scopes.",
    note: "Actual granted consent; requested permissions are a separate contract.",
    sources: ["https://learn.microsoft.com/graph/api/oauth2permissiongrant-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/oauth2PermissionGrants/{oAuth2PermissionGrant-id}", kind: "single", query: SINGLE_QUERY, fields: GRANT_FIELDS,
    access: "D/A Directory.Read.All as the read choice. Delegated callers pass it as --scopes.",
    sources: ["https://learn.microsoft.com/graph/api/oauth2permissiongrant-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/roleManagement/directory/roleAssignments", kind: "collection", query: COLLECTION_QUERY, fields: ROLE_ASSIGNMENT_FIELDS,
    access: "D/A RoleManagement.Read.Directory. Delegated callers pass it as --scopes.",
    note: "Built-in roles are base; custom role assignments need P1.",
    sources: ["https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignments?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/roleManagement/directory/roleEligibilityScheduleInstances", kind: "collection", query: COLLECTION_QUERY, fields: PIM_INSTANCE_FIELDS,
    access: "D/A RoleEligibilitySchedule.Read.Directory. Delegated callers pass it as --scopes.",
    note: "Eligible (not active) assignments; PIM needs P2 or ID Governance.",
    sources: ["https://learn.microsoft.com/graph/api/rbacapplication-list-roleeligibilityscheduleinstances?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/roleManagement/directory/roleAssignmentScheduleInstances", kind: "collection", query: COLLECTION_QUERY, fields: PIM_INSTANCE_FIELDS,
    access: "D/A RoleAssignmentSchedule.Read.Directory. Delegated callers pass it as --scopes.",
    note: "Active (including activated eligible) assignments; PIM needs P2 or ID Governance.",
    sources: ["https://learn.microsoft.com/graph/api/rbacapplication-list-roleassignmentscheduleinstances?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/devices", kind: "collection", query: COLLECTION_QUERY, fields: DEVICE_FIELDS,
    access: "D/A Device.Read.All. Delegated callers pass it as --scopes.",
    note: "Directory-device scope only; Intune device management is a separately authorized surface.",
    sources: ["https://learn.microsoft.com/graph/api/device-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/devices/{device-id}", kind: "single", query: SINGLE_QUERY, fields: DEVICE_FIELDS,
    access: "D/A Device.Read.All. Delegated callers pass it as --scopes.",
    sources: ["https://learn.microsoft.com/graph/api/device-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/administrativeUnits", kind: "collection", query: COLLECTION_QUERY, fields: AU_FIELDS,
    access: "D/A AdministrativeUnit.Read.All. Delegated callers pass it as --scopes.",
    note: "P1 for scoped administrators, Free for members; dynamic membership needs additional P1 licensing.",
    sources: ["https://learn.microsoft.com/graph/api/directory-list-administrativeunits?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/administrativeUnits/{administrativeUnit-id}", kind: "single", query: SINGLE_QUERY, fields: AU_FIELDS,
    access: "D/A AdministrativeUnit.Read.All. Delegated callers pass it as --scopes.",
    sources: ["https://learn.microsoft.com/graph/api/administrativeunit-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/organization", kind: "collection", query: ["$select"], fields: KNOWN_ORGANIZATION_FIELDS,
    access: "D User.Read for restricted basics (id, displayName and verifiedDomains only; other properties return null) or Organization.Read.All for full metadata; A Organization.Read.All. Delegated callers pass it as --scopes and additionally need a supported Entra role (Directory Readers and Global Reader are among the supported least-privilege roles). No P1/P2 prerequisite is stated for this list; contact fields are personal data.",
    note: "Graph documents $select only on the organization list; $filter/$top are not reviewed here. Exactly one organization exists per tenant.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/organization-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/organization/{organization-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_ORGANIZATION_FIELDS,
    access: "D User.Read for restricted basics (id, displayName and verifiedDomains only; other properties return null) or Organization.Read.All for full metadata; A Organization.Read.All. Delegated callers pass it as --scopes and additionally need a supported Entra role (Directory Readers and Global Reader are among the supported least-privilege roles). No P1/P2 prerequisite is stated for this read; contact fields are personal data.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/organization-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/organization/{organization-id}/branding", kind: "single", query: SINGLE_QUERY, fields: KNOWN_BRANDING_FIELDS,
    access: "D User.Read least-privileged or the purpose-built OrganizationalBranding.Read.All (Organization.Read.All also works). Delegated callers pass it as --scopes and additionally need Global Reader or Organizational Branding Administrator; A OrganizationalBranding.Read.All. No P1/P2 prerequisite is stated for this read, but configuring custom branding needs P1/P2. A 404 may indicate unconfigured branding or a missing or inaccessible organization.",
    note: "The session sends Accept-Language: 0 to read the default branding; only non-Stream properties are reviewed here and Stream image bytes are refused.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/organizationalbranding-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/organization/{organization-id}/branding/localizations", kind: "collection", query: ["$select"], fields: KNOWN_BRANDING_FIELDS,
    access: "D User.Read least-privileged or the purpose-built OrganizationalBranding.Read.All (Organization.Read.All also works). Delegated callers pass it as --scopes and additionally need Global Reader or Organizational Branding Administrator; A OrganizationalBranding.Read.All. No P1/P2 prerequisite is stated for this list.",
    note: "Graph documents $select only on the localization list; $filter/$top are not reviewed here.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/organizationalbranding-list-localizations?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/organization/{organization-id}/branding/localizations/{organizationalBrandingLocalization-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_BRANDING_FIELDS,
    access: "D User.Read least-privileged or the purpose-built OrganizationalBranding.Read.All (Organization.Read.All also works). Delegated callers pass it as --scopes and additionally need Global Reader or Organizational Branding Administrator; A OrganizationalBranding.Read.All. No P1/P2 prerequisite is stated for this read.",
    note: "Only non-Stream properties are reviewed here; Stream image bytes are refused.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/organizationalbrandinglocalization-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/domains", kind: "collection", query: ["$select"], fields: DOMAIN_FIELDS,
    access: "D/A Domain.Read.All. Delegated callers pass it as --scopes; a supported Entra role is also required (Domain Name Administrator or Global Reader are least-privileged). No P1/P2 prerequisite is stated for domain reads.",
    note: "Graph documents a known issue with $search, $top and $filter on domain lists; only $select is reviewed here.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/domain-list?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/domains/{domain-id}", kind: "single", query: SINGLE_QUERY, fields: DOMAIN_FIELDS,
    access: "D/A Domain.Read.All. Delegated callers pass it as --scopes; a supported Entra role is also required (Domain Name Administrator or Global Reader are least-privileged). No P1/P2 prerequisite is stated for domain reads.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/domain-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/domains/{domain-id}/verificationDnsRecords", kind: "collection", query: COLLECTION_QUERY, fields: DNS_RECORD_FIELDS,
    access: "D/A Domain.Read.All. Delegated callers pass it as --scopes; Domain Name Administrator or Global Reader are the least-privileged delegated roles. No P1/P2 prerequisite is stated for DNS record reads.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/domain-list-verificationdnsrecords?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/domains/{domain-id}/verificationDnsRecords/{domainDnsRecord-id}", kind: "single", query: SINGLE_QUERY, fields: DNS_RECORD_FIELDS,
    access: "D/A Domain.Read.All. Delegated callers pass it as --scopes; Domain Name Administrator or Global Reader are the least-privileged delegated roles. No P1/P2 prerequisite is stated for DNS record reads.",
    note: "No operation-level documentation page; access follows the parent verification-records contract and the domainDnsRecord resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/domain-list-verificationdnsrecords?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/domaindnsrecord?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/domains/{domain-id}/serviceConfigurationRecords", kind: "collection", query: COLLECTION_QUERY, fields: DNS_RECORD_FIELDS,
    access: "D/A Domain.Read.All. Delegated callers pass it as --scopes; Domain Name Administrator or Global Reader are the least-privileged delegated roles. No P1/P2 prerequisite is stated for DNS record reads.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/domain-list-serviceconfigurationrecords?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/domains/{domain-id}/serviceConfigurationRecords/{domainDnsRecord-id}", kind: "single", query: SINGLE_QUERY, fields: DNS_RECORD_FIELDS,
    access: "D/A Domain.Read.All. Delegated callers pass it as --scopes; Domain Name Administrator or Global Reader are the least-privileged delegated roles. No P1/P2 prerequisite is stated for DNS record reads.",
    note: "No operation-level documentation page; access follows the parent service-configuration contract and the domainDnsRecord resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/domain-list-serviceconfigurationrecords?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/domaindnsrecord?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/domainDnsRecords", kind: "collection", query: COLLECTION_QUERY, fields: DNS_RECORD_FIELDS,
    access: "D/A Domain.Read.All. Delegated callers pass it as --scopes; Domain Name Administrator or Global Reader are the least-privileged delegated roles. No P1/P2 prerequisite is stated for DNS record reads.",
    note: "No operation-level documentation page; access follows the documented domain/DNS-read contract and the domainDnsRecord resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/domain-list?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/domaindnsrecord?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/domainDnsRecords/{domainDnsRecord-id}", kind: "single", query: SINGLE_QUERY, fields: DNS_RECORD_FIELDS,
    access: "D/A Domain.Read.All. Delegated callers pass it as --scopes; Domain Name Administrator or Global Reader are the least-privileged delegated roles. No P1/P2 prerequisite is stated for DNS record reads.",
    note: "No operation-level documentation page; access follows the documented domain/DNS-read contract and the domainDnsRecord resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/domain-list?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/domaindnsrecord?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contracts", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_CONTRACT_FIELDS,
    access: "D/A Directory.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers is the least-privileged role); personal Microsoft accounts are not supported. Contracts exist in partner tenants only.",
    note: "Filtering is documented for customerId, defaultDomainName and displayName; no P1/P2 prerequisite is stated for contract reads.",
    sources: ["https://learn.microsoft.com/graph/api/contract-list?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/contract?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contracts/{contract-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_CONTRACT_FIELDS,
    access: "D/A Directory.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers is the least-privileged role); personal Microsoft accounts are not supported. Contracts exist in partner tenants only.",
    note: "No P1/P2 prerequisite is stated for contract reads.",
    sources: ["https://learn.microsoft.com/graph/api/contract-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/contract?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_DEFINITION_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    note: "Definitions are review schedules (a series), never their occurrences; access reviews need P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewset-list-definitions?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions/{accessReviewScheduleDefinition-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DEFINITION_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    note: "A definition never carries its instances; access reviews need P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewscheduledefinition-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions/{accessReviewScheduleDefinition-id}/instances", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_INSTANCE_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    note: "Instances are occurrences of one definition schedule, never schedules themselves; access reviews need P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewscheduledefinition-list-instances?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions/{accessReviewScheduleDefinition-id}/instances/{accessReviewInstance-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_INSTANCE_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewinstance-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions/{accessReviewScheduleDefinition-id}/instances/{accessReviewInstance-id}/decisions", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_DECISION_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    note: "Decisions are read-only here: listing never approves, denies or applies anything; access reviews need P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewinstance-list-decisions?view=graph-rest-1.0"] },
];

function splitPath(path: string): string[] {
  return path.split("/").filter(segment => segment.length > 0);
}

// Structural match of a server-relative path against a reviewed template:
// literals compare case-insensitively like Graph routing; whole-segment
// placeholders capture one non-empty segment each for the session to bind and
// validate. Returns the bindings, or null when the route is not this one.
function matchTemplate(template: string, pathname: string): Record<string, string> | null {
  const expected = splitPath(template);
  const actual = splitPath(pathname);
  if (expected.length !== actual.length) return null;
  const params: Record<string, string> = Object.create(null);
  for (let i = 0; i < expected.length; i++) {
    const slot = expected[i]!;
    const segment = actual[i]!;
    const name = /^\{([^{}]+)\}$/.exec(slot)?.[1];
    if (name) {
      if (!segment.length) return null;
      encodeGraphPathSegment(segment);
      params[name] = segment;
    } else if (slot.toLowerCase() !== segment.toLowerCase()) return null;
  }
  return params;
}

export function matchReviewed(pathname: string): { route: ReviewedRawRoute; params: Record<string, string> } | null {
  for (const route of REVIEWED_ROUTES) {
    const params = matchTemplate(route.id.slice("v1.0:GET:".length), pathname);
    if (params) return { route, params };
  }
  return null;
}

function parseQuery(raw: string | undefined): Record<string, string> {
  if (!raw) return {};
  const text = raw.startsWith("?") ? raw.slice(1) : raw;
  const params = new URLSearchParams(text);
  const out: Record<string, string> = Object.create(null);
  for (const [key, value] of params) {
    if (Object.hasOwn(out, key)) throw new AxiError(`Duplicate query key ${key}`, "VALIDATION_ERROR", ["Pass each query key once: --odata 'k=v&k2=v2'"]);
    out[key] = value;
  }
  if (!Object.keys(out).length && text.trim() !== "") throw new AxiError(`Invalid --odata '${raw}'`, "VALIDATION_ERROR", ["Example: --odata '$top=5&$select=id,displayName'"]);
  return out;
}

function checkQueryKeys(route: ReviewedRawRoute, query: Record<string, string>): void {
  for (const key of Object.keys(query)) {
    if (route.query.includes(key)) continue;
    if (key === "$search" || key === "$count") {
      throw new AxiError(`Advanced query ${key} is not reviewed for raw reads`, "VALIDATION_ERROR", [
        "Advanced queries need per-endpoint ConsistencyLevel review; named slices declare supported combinations",
      ]);
    }
    if (key === "$skiptoken") {
      throw new AxiError("Server paging state cannot be supplied; page with --all", "VALIDATION_ERROR", ["mg-axi api get --help"]);
    }
    throw new AxiError(`Unsupported query key ${key} for ${route.id}`, "VALIDATION_ERROR", [
      `Reviewed query keys: ${route.query.join(", ")}`,
    ]);
  }
}

function checkQuery(route: ReviewedRawRoute, query: Record<string, string>): void {
  checkQueryKeys(route, query);
  const select = query["$select"] ??= route.fields.join(",");
  const fields = select.split(",").map(field => field.trim()).filter(field => field.length > 0);
  if (!fields.length) throw new AxiError("Empty $select names no fields", "VALIDATION_ERROR", [`Reviewed fields: ${route.fields.join(", ")}`]);
  const unknown = fields.filter(field => !route.fields.includes(field));
  if (unknown.length) {
    throw new AxiError(`Unreviewed $select field${unknown.length === 1 ? "" : "s"} ${unknown.join(", ")} for ${route.id}`, "VALIDATION_ERROR", [
      `Reviewed fields: ${route.fields.join(", ")}`,
    ]);
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && Object.getPrototypeOf(value) === Object.prototype;
}

function reviewedFields(route: ReviewedRawRoute, value: unknown): unknown {
  if (!isRecord(value)) return null;
  return Object.fromEntries(Object.entries(value).filter(([key]) => route.fields.includes(key)));
}

// String truncation at the output boundary, mirroring az-axi's 4000-char
// convention. --full disables truncation; redaction already ran in the
// session and row caps are unaffected. Plain objects and arrays only.
function truncateForOutput(value: unknown, full: boolean): { value: unknown; truncated: boolean } {
  if (typeof value === "string") {
    if (full || value.length <= TRUNCATE_AT) return { value, truncated: false };
    return { value: `${value.slice(0, TRUNCATE_AT)}... (truncated, ${value.length} chars total)`, truncated: true };
  }
  if (Array.isArray(value)) {
    let truncated = false;
    const shaped = value.map(item => {
      const child = truncateForOutput(item, full);
      truncated = truncated || child.truncated;
      return child.value;
    });
    return { value: shaped, truncated };
  }
  if (isRecord(value)) {
    let truncated = false;
    const shaped: Record<string, unknown> = {};
    for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
      const next = truncateForOutput(child, full);
      truncated = truncated || next.truncated;
      shaped[key] = next.value;
    }
    return { value: shaped, truncated };
  }
  return { value, truncated: false };
}

export interface ApiGetArgs {
  path: string;
  apiVersion: string;
  odata?: string;
  cursor?: string;
  scopes?: string;
  /** Row cap; undefined with --all follows pages within the session budget. */
  limit?: number;
  full?: boolean;
  profile: AnyProfile;
}

export interface ApiDeps {
  delegated: DelegatedAuth;
  application: ApplicationAuth;
  transport: GraphTransport;
}

export async function runApiGet(args: ApiGetArgs, deps: ApiDeps): Promise<Record<string, unknown>> {
  const version = args.apiVersion || "v1.0";
  const pathname = args.path ?? "";
  if (!pathname.startsWith("/") || pathname.includes("://") || pathname.includes("?") || pathname.includes("#") || pathname.includes("\\")) {
    throw new AxiError(`Refused raw path '${pathname || "(empty)"}'`, "VALIDATION_ERROR", [
      "Pass a server-relative Graph path such as /users; query belongs in --odata 'k=v&k2=v2'",
      "The command never takes a host, scheme or headers; the session owns the Graph destination",
    ]);
  }
  const matched = matchReviewed(pathname);
  if (!matched || version !== "v1.0") {
    if (version !== "v1.0") {
      throw new AxiError(`No reviewed ${version} raw reads (API-01 reviews v1.0 only): GET ${pathname}`, "VALIDATION_ERROR", [
        "mg-axi api get --help",
        "Beta reads need an explicitly reviewed contract; v1.0 stays the default with no fallback",
      ]);
    }
    throw new AxiError(`GET ${pathname} is not in the reviewed raw inventory (API-01, reviewed ${REVIEWED_ON})`, "VALIDATION_ERROR", [
      "mg-axi api get --help",
      "Named commands, secret-value, mail/file-content and write routes are never served raw",
    ]);
  }
  const { route, params } = matched;
  if (route.kind === "single" && args.cursor !== undefined) {
    throw new AxiError("--cursor is available for collection reads only", "VALIDATION_ERROR", ["Resume with the same collection path that returned the cursor"]);
  }
  if (args.profile.mode === "application" && args.scopes !== undefined) {
    throw new AxiError("Application profiles use the configured Graph .default audience; delegated scopes are unavailable", "VALIDATION_ERROR", [
      "mg-axi profile show --profile <name>",
    ]);
  }
  const query = args.cursor !== undefined && args.odata === undefined ? undefined : parseQuery(args.odata);
  // The inventory row is authoritative for destination, method and policy; the
  // review above only selects which row may run. Binding, query-shape and
  // policy failures below still throw before the session acquires credentials.
  const operation = resolveSessionOperation(version, "GET", route.id.slice(`${version}:GET:`.length));
  const localSelection = operation.path === "/users/{user-id}/authentication/methods";
  const scopes = args.scopes === undefined ? undefined : args.scopes.split(",").map(scope => scope.trim()).filter(scope => scope.length > 0);
  const session = new GraphSession({ ...deps, transport: async request => {
    const url = new URL(request.url);
    const query = Object.fromEntries(url.searchParams);
    delete query["$skiptoken"];
    checkQuery(route, query);
    if (localSelection) url.searchParams.delete("$select");
    else url.searchParams.set("$select", query["$select"]!);
    const response = await deps.transport({ ...request, url: url.toString() });
    if (response.status < 200 || response.status >= 300 || !response.body) return response;
    let body: unknown;
    try { body = JSON.parse(response.body); } catch { return response; }
    if (route.kind === "single") body = reviewedFields(route, body);
    else if (isRecord(body) && Array.isArray(body.value)) body = { ...body, value: body.value.map(row => reviewedFields(outputRoute, row)) };
    return { ...response, body: JSON.stringify(body), receivedBodyBytes: response.receivedBodyBytes ?? Buffer.byteLength(response.body, "utf8") };
  } });
  const selection = query ?? { ...session.cursorQuery(operation, args.cursor!) };
  checkQuery(route, selection);
  const outputRoute = localSelection ? { ...route, fields: selection["$select"]!.split(",").map(field => field.trim()).filter(Boolean) } : route;
  const full = !!args.full;
  if (route.kind === "single") {
    const body = await session.execute({ profile: args.profile, operation, params, query, scopes });
    const shaped = truncateForOutput(reviewedFields(route, body), full);
    const record = isRecord(shaped.value) ? (shaped.value as Record<string, unknown>) : { value: shaped.value };
    return shaped.truncated ? { ...record, help: ["Strings truncated at 4000 chars; re-run with --full"] } : record;
  }
  const collected = await session.collect({ profile: args.profile, operation, params, query, scopes, limit: args.limit, cursor: args.cursor });
  const shaped = truncateForOutput(collected.value.map(row => reviewedFields(outputRoute, row)), full);
  const value = shaped.value as unknown[];
  const warnings = route.warning ? { warnings: [route.warning] } : {};
  if (collected.complete) {
    return shaped.truncated
      ? { returned: value.length, complete: true, value, ...warnings, help: ["Strings truncated at 4000 chars; re-run with --full"] }
      : { returned: value.length, complete: true, value, ...warnings };
  }
  const hint = "Resume the same path, profile and scopes with --cursor - and supply the cursor token on stdin; use --limit <rows> or --all";
  const help = shaped.truncated ? [hint, "Strings truncated at 4000 chars; re-run with --full"] : [hint];
  return { returned: value.length, complete: false, reason: collected.reason, value, ...warnings, cursor: collected.cursor, help };
}

// Minimal HTTPS transport: sends only what the session authorized (GET on the
// commercial Graph host with session-owned headers) and returns status,
// headers and text. Timeouts and cancellation arrive through the session's
// AbortSignal; credential attachment happened before this call.
export const fetchTransport: GraphTransport = async request => {
  const response = await fetch(request.url, { method: "GET", headers: request.headers, signal: request.signal, redirect: "manual" });
  const headers: Record<string, string> = {};
  response.headers.forEach((value, key) => {
    headers[key] = value;
  });
  return { status: response.status, headers, body: await response.text() };
};
