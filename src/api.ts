import { AxiError } from "axi-sdk-js";
import { ApplicationAuth } from "./app-auth.js";
import { DelegatedAuth } from "./auth.js";
import { KNOWN_CONTACTED_REVIEWER_FIELDS, KNOWN_DECISION_FIELDS, KNOWN_DEFINITION_FIELDS, KNOWN_INSTANCE_FIELDS, KNOWN_STAGE_FIELDS } from "./entra-access-reviews.js";
import { KNOWN_ACCESS_PACKAGE_FIELDS, KNOWN_ASSIGNMENT_POLICY_FIELDS, KNOWN_CATALOG_FIELDS, KNOWN_RESOURCE_ROLE_SCOPE_FIELDS } from "./entra-entitlement-management.js";
import { KNOWN_LIFECYCLE_SETTINGS_FIELDS, KNOWN_TASK_DEFINITION_FIELDS, KNOWN_WORKFLOW_FIELDS, KNOWN_WORKFLOW_TEMPLATE_FIELDS } from "./entra-lifecycle-workflows.js";
import { KNOWN_BRANDING_FIELDS, KNOWN_ORGANIZATION_FIELDS } from "./entra-organization.js";
import { KNOWN_CONTRACT_FIELDS } from "./entra-contracts.js";
import { KNOWN_DELEGATED_ADMIN_CUSTOMER_FIELDS, KNOWN_DELEGATED_ADMIN_RELATIONSHIP_FIELDS, KNOWN_DELEGATED_ADMIN_ACCESS_ASSIGNMENT_FIELDS, KNOWN_DELEGATED_ADMIN_OPERATION_FIELDS, KNOWN_DELEGATED_ADMIN_REQUEST_FIELDS } from "./entra-delegated-admin.js";
import { KNOWN_TENANT_INFORMATION_FIELDS } from "./entra-tenant-information.js";
import { KNOWN_MTO_FIELDS, KNOWN_MTO_JOIN_REQUEST_FIELDS, KNOWN_MTO_TENANT_FIELDS } from "./entra-multi-tenant-organization.js";
import { KNOWN_DATA_POLICY_FIELDS } from "./entra-data-policy-operations.js";
import { KNOWN_FEDERATION_FIELDS } from "./entra-federation-configurations.js";
import { KNOWN_CA_FIELDS, KNOWN_PKI_FIELDS } from "./entra-certificate-auth.js";
import { KNOWN_SUBSCRIPTION_FIELDS } from "./entra-subscriptions.js";
import { KNOWN_SYNC_FIELDS } from "./entra-on-premises-synchronization.js";
import { KNOWN_ACCEPTANCE_FIELDS, KNOWN_AGREEMENT_FIELDS } from "./entra-terms-of-use.js";
import { KNOWN_DIRECTORY_OBJECT_FIELDS } from "./entra-directory-objects.js";
import { KNOWN_DELETED_ADMINISTRATIVE_UNIT_FIELDS, KNOWN_DELETED_APPLICATION_FIELDS, KNOWN_DELETED_GROUP_FIELDS, KNOWN_DELETED_SERVICE_PRINCIPAL_FIELDS, KNOWN_DELETED_SHOW_FIELDS, KNOWN_DELETED_USER_FIELDS } from "./entra-deleted-items.js";
import { KNOWN_CONTACT_FIELDS, KNOWN_NAV_FIELDS } from "./entra-contacts.js";
import { KNOWN_LIFECYCLE_FIELDS, KNOWN_TEMPLATE_FIELDS } from "./entra-group-lifecycle.js";
import { KNOWN_ALLOWED_VALUE_FIELDS, KNOWN_ATTRIBUTE_SET_FIELDS, KNOWN_CUSTOM_SECURITY_DEFINITION_FIELDS } from "./entra-custom-security-attributes.js";
import { encodeGraphPathSegment, functionBindingFor, GraphSession, resolveSessionOperation, type GraphTransport } from "./graph-session.js";
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
  readonly defaultFields?: readonly string[];
  /** Supported read permission choices with delegated/application distinction. */
  readonly access: string;
  /** When true, application profiles are refused before credentials: Graph documents no supported application permission. */
  readonly delegatedOnly?: boolean;
  readonly note?: string;
  readonly warning?: string;
  /** When true, the @odata.type subtype discriminator rides along in output without being a selectable field. */
  readonly keepODataType?: boolean;
  /** Primary-source operation documentation, rechecked on REVIEWED_ON. */
  readonly sources: readonly string[];
}

// Collections support server filtering, ordering and paging hints. Advanced
// queries ($search, $count=true) need per-endpoint ConsistencyLevel review and
// stay unsupported here; $skiptoken is server paging state, paged via --all.
const COLLECTION_QUERY = ["$select", "$filter", "$top", "$orderby"] as const;
const SINGLE_QUERY = ["$select"] as const;
const DEFAULT_CA_FIELDS = KNOWN_CA_FIELDS.filter(field => field !== "certificate");

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
const PROVIDER_FIELDS = ["id", "displayName", "identityProviderType", "clientId"];

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
  { id: "v1.0:GET:/directory/publicKeyInfrastructure/certificateBasedAuthConfigurations", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_PKI_FIELDS,
    access: "D/A PublicKeyInfrastructure.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Privileged Authentication Administrator or Authentication Administrator; personal Microsoft accounts are not supported. No P1/P2 prerequisite is stated for PKI reads; an empty list may mean certificate-based authentication is not configured.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/publickeyinfrastructureroot-list-certificatebasedauthconfigurations?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/publicKeyInfrastructure/certificateBasedAuthConfigurations/{certificateBasedAuthPki-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_PKI_FIELDS,
    access: "D/A PublicKeyInfrastructure.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Privileged Authentication Administrator or Authentication Administrator; personal Microsoft accounts are not supported. No P1/P2 prerequisite is stated for this read.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/certificatebasedauthpki-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/publicKeyInfrastructure/certificateBasedAuthConfigurations/{certificateBasedAuthPki-id}/certificateAuthorities", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_CA_FIELDS, defaultFields: DEFAULT_CA_FIELDS,
    access: "D/A PublicKeyInfrastructure.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Privileged Authentication Administrator or Authentication Administrator; personal Microsoft accounts are not supported. No P1/P2 prerequisite is stated for authority reads.",
    note: "The certificate field carries the public CA key only; large blobs stay truncated at the output boundary unless --full is passed.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/certificatebasedauthpki-list-certificateauthorities?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/certificateauthoritydetail?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/publicKeyInfrastructure/certificateBasedAuthConfigurations/{certificateBasedAuthPki-id}/certificateAuthorities/{certificateAuthorityDetail-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_CA_FIELDS, defaultFields: DEFAULT_CA_FIELDS,
    access: "D/A PublicKeyInfrastructure.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Privileged Authentication Administrator or Authentication Administrator; personal Microsoft accounts are not supported. No P1/P2 prerequisite is stated for this read.",
    note: "No operation-level query documentation beyond $select; access follows the parent authority-list contract and the certificateAuthorityDetail resource reference. The certificate field carries the public CA key only.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/certificatebasedauthpki-list-certificateauthorities?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/certificateauthoritydetail-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/certificateauthoritydetail?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/subscriptions", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_SUBSCRIPTION_FIELDS,
    access: "D/A Organization.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Global Reader, Directory Readers, or Dynamics 365 Business Central Administrator for read-only standard properties); personal Microsoft accounts are not supported. No P1/P2 prerequisite is stated for subscription reads.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directory-list-subscriptions?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/companysubscription?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/subscriptions/{companySubscription-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_SUBSCRIPTION_FIELDS,
    access: "D/A Organization.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Global Reader, Directory Readers, or Dynamics 365 Business Central Administrator for read-only standard properties); personal Microsoft accounts are not supported. No P1/P2 prerequisite is stated for this read.",
    note: "The commerceSubscriptionId alternate-key lookup ships as its own reviewed route below; object-ID shows never accept function-style identifiers.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/companysubscription-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/companysubscription?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/subscriptions(commerceSubscriptionId='{commerceSubscriptionId}')", kind: "single", query: SINGLE_QUERY, fields: KNOWN_SUBSCRIPTION_FIELDS,
    access: "D/A Organization.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Global Reader, Directory Readers, or Dynamics 365 Business Central Administrator for read-only standard properties); personal Microsoft accounts are not supported. No P1/P2 prerequisite is stated for this read.",
    note: "Alternate-key lookup by commerceSubscriptionId: the raw path carries the key as an OData-quoted function argument and the shared session re-validates, re-quotes and encodes it; only this allowlisted shape binds.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/companysubscription-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/companysubscription?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/onPremisesSynchronization", kind: "collection", query: ["$select"], fields: KNOWN_SYNC_FIELDS,
    access: "D OnPremDirectorySynchronization.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Global Administrator, the only supported Entra role for this operation; personal Microsoft accounts are not supported and Graph documents no supported application permission, so application profiles are refused before credentials. No P1/P2 prerequisite is stated for on-premises-synchronization reads.",
    delegatedOnly: true,
    note: "Graph documents $select only on the on-premises-synchronization list; $filter is not reviewed here. The $count scalar stays scheduled for a later counts subfamily.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/onpremisesdirectorysynchronization-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/onpremisesdirectorysynchronization?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/onPremisesSynchronization/{onPremisesDirectorySynchronization-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_SYNC_FIELDS,
    access: "D OnPremDirectorySynchronization.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Global Administrator, the only supported Entra role for this operation; personal Microsoft accounts are not supported and Graph documents no supported application permission, so application profiles are refused before credentials. No P1/P2 prerequisite is stated for this read.",
    delegatedOnly: true,
    note: "Secret-shaped values inside configuration and features stay redacted by the shared session; no credential fields are projected.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/onpremisesdirectorysynchronization-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/onpremisesdirectorysynchronization?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/agreements", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_AGREEMENT_FIELDS, defaultFields: ["id", "displayName"],
    access: "D Agreement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Security Reader, the least-privileged supported Entra role for this operation (Global Reader, Conditional Access Administrator or Security Administrator also work); personal Microsoft accounts are not supported and Graph documents no supported application permission, so application profiles are refused before credentials. Terms of use needs Microsoft Entra ID P1.",
    delegatedOnly: true,
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. The legacy /agreements root shares the termsOfUse operation contract documented on /identityGovernance/termsOfUse/agreements. Agreement file and files navigation properties are never requested: file bytes are never downloaded or printed.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/termsofusecontainer-list-agreements?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/agreement?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/agreements/{agreement-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_AGREEMENT_FIELDS,
    access: "D Agreement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Security Reader, the least-privileged supported Entra role for this operation (Global Reader, Conditional Access Administrator or Security Administrator also work); personal Microsoft accounts are not supported and Graph documents no supported application permission, so application profiles are refused before credentials. Terms of use needs Microsoft Entra ID P1.",
    delegatedOnly: true,
    note: "Graph documents $select only on the agreement get; $filter is not reviewed here. Agreement file and files navigation properties are never requested: file bytes are never downloaded or printed.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/agreement-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/agreement?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/agreements/{agreement-id}/acceptances", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_ACCEPTANCE_FIELDS, defaultFields: ["id", "agreementId", "state", "recordedDateTime"],
    access: "D AgreementAcceptance.Read least-privileged (AgreementAcceptance.Read.All is the documented higher-privileged alternative). Delegated callers pass one as --scopes; delegated access additionally needs Security Reader, the least-privileged supported Entra role for this operation (Global Reader, Conditional Access Administrator or Security Administrator also work); personal Microsoft accounts are not supported and Graph documents no supported application permission, so application profiles are refused before credentials. Terms of use needs Microsoft Entra ID P1.",
    delegatedOnly: true,
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. Acceptance records are personal data: default rows carry id, agreementId, state and recordedDateTime only, and identifying fields need an explicit $select.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/agreement-list-acceptances?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/agreementacceptance?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/agreements/{agreement-id}/acceptances/{agreementAcceptance-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_ACCEPTANCE_FIELDS, defaultFields: ["id", "agreementId", "state", "recordedDateTime"],
    access: "D AgreementAcceptance.Read least-privileged (AgreementAcceptance.Read.All is the documented higher-privileged alternative). Delegated callers pass one as --scopes; delegated access additionally needs Security Reader, the least-privileged supported Entra role for this operation (Global Reader, Conditional Access Administrator or Security Administrator also work); personal Microsoft accounts are not supported and Graph documents no supported application permission, so application profiles are refused before credentials. Terms of use needs Microsoft Entra ID P1.",
    delegatedOnly: true,
    note: "No single-acceptance operation page; access follows the parent acceptance-list contract and the agreementAcceptance resource reference. Acceptance records are personal data: default rows carry id, agreementId, state and recordedDateTime only.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/agreement-list-acceptances?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/agreementacceptance?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/agreementAcceptances", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_ACCEPTANCE_FIELDS, defaultFields: ["id", "agreementId", "state", "recordedDateTime"],
    access: "D AgreementAcceptance.Read least-privileged (AgreementAcceptance.Read.All is the documented higher-privileged alternative). Delegated callers pass one as --scopes; delegated access additionally needs Security Reader, the least-privileged supported Entra role for this operation (Global Reader, Conditional Access Administrator or Security Administrator also work); personal Microsoft accounts are not supported and Graph documents no supported application permission, so application profiles are refused before credentials. Terms of use needs Microsoft Entra ID P1.",
    delegatedOnly: true,
    note: "The root acceptance list has no dedicated operation page; access follows the reviewed acceptance-list contract and the agreementAcceptance resource reference. Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. Acceptance records are personal data: default rows carry id, agreementId, state and recordedDateTime only.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/agreement-list-acceptances?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/agreementacceptance?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/agreementAcceptances/{agreementAcceptance-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_ACCEPTANCE_FIELDS, defaultFields: ["id", "agreementId", "state", "recordedDateTime"],
    access: "D AgreementAcceptance.Read least-privileged (AgreementAcceptance.Read.All is the documented higher-privileged alternative). Delegated callers pass one as --scopes; delegated access additionally needs Security Reader, the least-privileged supported Entra role for this operation (Global Reader, Conditional Access Administrator or Security Administrator also work); personal Microsoft accounts are not supported and Graph documents no supported application permission, so application profiles are refused before credentials. Terms of use needs Microsoft Entra ID P1.",
    delegatedOnly: true,
    note: "No single-acceptance operation page; access follows the parent acceptance-list contract and the agreementAcceptance resource reference. Acceptance records are personal data: default rows carry id, agreementId, state and recordedDateTime only.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/agreement-list-acceptances?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/agreementacceptance?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directoryObjects", kind: "collection", query: ["$select"], fields: KNOWN_DIRECTORY_OBJECT_FIELDS, defaultFields: ["id"], keepODataType: true,
    access: "D/A Directory.Read.All. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for directory-object reads. Personal Microsoft accounts are not supported.",
    note: "No List operation page exists; access follows the sibling directoryobject-get contract and the directoryObject resource reference. Graph documents no collection filter contract here, so only $select is reviewed. Rows are polymorphic: the @odata.type discriminator rides along automatically and subtype properties need the subtype's named reads, never raw $select of unreviewed fields.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directoryobject-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/directoryobject?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directoryObjects/{directoryObject-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DIRECTORY_OBJECT_FIELDS, defaultFields: ["id", "deletedDateTime"], keepODataType: true,
    access: "D/A Directory.Read.All. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for this read. Personal Microsoft accounts are not supported.",
    note: "Single reads return the base-type properties plus the @odata.type discriminator; subtype detail needs the subtype's named reads. Only base properties are projected, so subtype secrets or credentials can never appear.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directoryobject-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/directoryobject?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/deletedItems/graph.user", kind: "collection", query: ["$select"], fields: KNOWN_DELETED_USER_FIELDS, defaultFields: ["id", "displayName", "deletedDateTime"], keepODataType: true,
    access: "D/A User.Read.All. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for deleted-user reads. Personal Microsoft accounts are not supported.",
    note: "Upstream requires the user cast: untyped GET /directory/deletedItems is not supported. Rows are polymorphic: the @odata.type discriminator rides along automatically. Deleted users are personal data and default rows stay minimal.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directory-deleteditems-list?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/deletedItems/graph.group", kind: "collection", query: ["$select"], fields: KNOWN_DELETED_GROUP_FIELDS, defaultFields: ["id", "displayName", "deletedDateTime"], keepODataType: true,
    access: "D/A Group.Read.All. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for deleted-group reads. Personal Microsoft accounts are not supported.",
    note: "Upstream requires the group cast: untyped GET /directory/deletedItems is not supported. Soft-deleted security groups report securityEnabled false through a known upstream limitation; read groupTypes to name the real kind.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directory-deleteditems-list?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/deletedItems/graph.application", kind: "collection", query: ["$select"], fields: KNOWN_DELETED_APPLICATION_FIELDS, defaultFields: ["id", "appId", "displayName", "deletedDateTime"], keepODataType: true,
    access: "D/A Application.Read.All. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for deleted-application reads. Personal Microsoft accounts are not supported.",
    note: "Upstream requires the application cast: untyped GET /directory/deletedItems is not supported. appId (client ID) is distinct from the object id; credential collections are never projected and GET never returns secret values.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directory-deleteditems-list?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/application?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/deletedItems/graph.servicePrincipal", kind: "collection", query: ["$select"], fields: KNOWN_DELETED_SERVICE_PRINCIPAL_FIELDS, defaultFields: ["id", "appId", "displayName", "deletedDateTime"], keepODataType: true,
    access: "D/A Application.Read.All. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for deleted-service-principal reads. Personal Microsoft accounts are not supported.",
    note: "Upstream requires the servicePrincipal cast: untyped GET /directory/deletedItems is not supported. appId (client ID) is distinct from the object id; secret-bearing fields are never projected.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directory-deleteditems-list?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/serviceprincipal?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/deletedItems/graph.administrativeUnit", kind: "collection", query: ["$select"], fields: KNOWN_DELETED_ADMINISTRATIVE_UNIT_FIELDS, defaultFields: ["id", "displayName", "deletedDateTime"], keepODataType: true,
    access: "D/A AdministrativeUnit.Read.All. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for deleted-administrative-unit reads. Personal Microsoft accounts are not supported.",
    note: "Upstream requires the administrativeUnit cast: untyped GET /directory/deletedItems is not supported. Rows are polymorphic: the @odata.type discriminator rides along automatically.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directory-deleteditems-list?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/administrativeunit?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/deletedItems/{directoryObject-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DELETED_SHOW_FIELDS, keepODataType: true,
    access: "D/A User.Read.All, Group.Read.All, Application.Read.All or AdministrativeUnit.Read.All matching the object's type. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for this read. Personal Microsoft accounts are not supported.",
    note: "The untyped get can return any deletable type, so only the cross-type safe properties are reviewed here; per-type detail needs the typed deleted-item lists.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directory-deleteditems-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/directoryobject?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_CONTACT_FIELDS, defaultFields: ["id", "displayName", "mail", "companyName"],
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All is the documented higher-privileged alternative). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "Filtering passes through as plain $filter with $count=true and ConsistencyLevel eventual; $search and $orderby stay unreviewed. Contacts are personal data: default rows carry id, displayName, mail and companyName only, and identifying fields need an explicit $select. Only flat scalar properties are reviewed; $expand is never offered and the nested phones/addresses/error collections need their own projection review.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_CONTACT_FIELDS, defaultFields: ["id", "displayName", "mail", "companyName"],
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All is the documented higher-privileged alternative). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "Single reads return the reviewed flat scalar set with minimal personal-data defaults; navigation objects are never expanded and subtype detail is not applicable. Only $select is reviewed.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/manager", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "Single directoryObject: the manager user or contact. Rows carry the @odata.type discriminator plus minimal personal-data defaults; subtype detail needs the subtype's named reads. Only $select is reviewed.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-get-manager?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/directReports", kind: "collection", query: ["$select"], fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "Graph documents $select only on directReports; $filter/$search/$top stay unreviewed. Rows are polymorphic users/contacts: the @odata.type discriminator rides along automatically and subtype properties need the subtype's named reads, never raw $select of unreviewed fields.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-directreports?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/directReports/graph.orgContact", kind: "collection", query: ["$select"], fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "OData cast selects only contact direct reports; access follows the parent directReports contract and the orgContact resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-directreports?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/directReports/graph.user", kind: "collection", query: ["$select"], fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "OData cast selects only user direct reports; access follows the parent directReports contract.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-directreports?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/directReports/{directoryObject-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; access follows the parent directReports contract. Single directoryObject with the @odata.type discriminator plus minimal personal-data defaults.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-directreports?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/directReports/{directoryObject-id}/graph.orgContact", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; OData cast selects the contact subtype and access follows the parent directReports contract and the orgContact resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-directreports?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/directReports/{directoryObject-id}/graph.user", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; OData cast selects the user subtype and access follows the parent directReports contract.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-directreports?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/memberOf", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "Filtering passes through as plain $filter with $count=true and ConsistencyLevel eventual; $search and $orderby stay unreviewed. Rows are groups and administrative units: the @odata.type discriminator rides along automatically and subtype properties need the subtype's named reads, never raw $select of unreviewed fields.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-memberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/memberOf/graph.group", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "OData cast selects only group memberships; access follows the parent memberOf contract and the group resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-memberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/memberOf/graph.administrativeUnit", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "OData cast selects only administrative-unit memberships; access follows the parent memberOf contract and the administrativeUnit resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-memberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/administrativeunit?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; access follows the parent memberOf contract. Single directoryObject with the @odata.type discriminator plus minimal defaults.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-memberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}/graph.group", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; OData cast selects the group subtype and access follows the parent memberOf contract and the group resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-memberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}/graph.administrativeUnit", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All least-privileged (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are the documented higher-privileged alternatives). Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; OData cast selects the administrativeUnit subtype and access follows the parent memberOf contract and the administrativeUnit resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-memberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/administrativeunit?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/transitiveMemberOf", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All and Group.Read.All together least-privileged (Directory.Read.All is the documented higher-privileged alternative). Delegated callers pass both as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "Filtering passes through as plain $filter with $count=true and ConsistencyLevel eventual; $search and $orderby stay unreviewed. Rows are groups and administrative units: the @odata.type discriminator rides along automatically and subtype properties need the subtype's named reads, never raw $select of unreviewed fields.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-transitivememberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.group", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All and Group.Read.All together least-privileged (Directory.Read.All is the documented higher-privileged alternative). Delegated callers pass both as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "OData cast selects only group memberships; access follows the parent transitiveMemberOf contract and the group resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-transitivememberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.administrativeUnit", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All and Group.Read.All together least-privileged (Directory.Read.All is the documented higher-privileged alternative). Delegated callers pass both as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "OData cast selects only administrative-unit memberships; access follows the parent transitiveMemberOf contract and the administrativeUnit resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-transitivememberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/administrativeunit?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All and Group.Read.All together least-privileged (Directory.Read.All is the documented higher-privileged alternative). Delegated callers pass both as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; access follows the parent transitiveMemberOf contract. Single directoryObject with the @odata.type discriminator plus minimal defaults.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-transitivememberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}/graph.group", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All and Group.Read.All together least-privileged (Directory.Read.All is the documented higher-privileged alternative). Delegated callers pass both as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; OData cast selects the group subtype and access follows the parent transitiveMemberOf contract and the group resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-transitivememberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}/graph.administrativeUnit", kind: "single", query: SINGLE_QUERY, fields: KNOWN_NAV_FIELDS, defaultFields: ["id", "displayName"], keepODataType: true,
    access: "D/A OrgContact.Read.All and Group.Read.All together least-privileged (Directory.Read.All is the documented higher-privileged alternative). Delegated callers pass both as --scopes; delegated access additionally needs a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work). Personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; OData cast selects the administrativeUnit subtype and access follows the parent transitiveMemberOf contract and the administrativeUnit resource reference.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/orgcontact-list-transitivememberof?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/administrativeunit?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/groupLifecyclePolicies", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_LIFECYCLE_FIELDS,
    access: "D/A Directory.Read.All. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for lifecycle-policy reads. Personal Microsoft accounts are not supported.",
    note: "Filtering uses plain $filter with no $count or ConsistencyLevel contract; no P1/P2 prerequisite is stated for lifecycle-policy reads.",
    sources: ["https://learn.microsoft.com/graph/api/grouplifecyclepolicy-list?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/grouplifecyclepolicy?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/groupLifecyclePolicies/{groupLifecyclePolicy-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_LIFECYCLE_FIELDS,
    access: "D/A Directory.Read.All. Delegated callers pass it as --scopes; no delegated role prerequisite is stated for lifecycle-policy reads. Personal Microsoft accounts are not supported.",
    note: "No P1/P2 prerequisite is stated for lifecycle-policy reads.",
    sources: ["https://learn.microsoft.com/graph/api/grouplifecyclepolicy-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/grouplifecyclepolicy?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/groupSettingTemplates", kind: "collection", query: ["$select"], fields: KNOWN_TEMPLATE_FIELDS,
    access: "D/A GroupSettings.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers or Global Reader are the least-privileged roles). Personal Microsoft accounts are not supported.",
    note: "Graph documents $select only for template lists; $filter is not supported. No P1/P2 prerequisite is stated for setting-template reads.",
    sources: ["https://learn.microsoft.com/graph/api/groupsettingtemplate-list?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/groupsettingtemplate?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/groupSettingTemplates/{groupSettingTemplate-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_TEMPLATE_FIELDS,
    access: "D/A GroupSettings.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers or Global Reader are the least-privileged roles). Personal Microsoft accounts are not supported.",
    note: "No P1/P2 prerequisite is stated for setting-template reads.",
    sources: ["https://learn.microsoft.com/graph/api/groupsettingtemplate-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/groupsettingtemplate?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/attributeSets", kind: "collection", query: ["$select", "$top", "$orderby"], fields: KNOWN_ATTRIBUTE_SET_FIELDS,
    access: "D/A CustomSecAttributeDefinition.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a custom-security-attribute role (Attribute Assignment Reader, Attribute Definition Reader, Attribute Assignment Administrator or Attribute Definition Administrator); by default Global Administrator and other administrator roles have no custom-security-attribute access. Personal Microsoft accounts are not supported.",
    note: "Graph documents $select, $top and $orderby only for attribute-set lists; $filter is not supported. No P1/P2 prerequisite is stated for attribute-set reads.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directory-list-attributesets?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/attributeset?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/attributeSets/{attributeSet-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_ATTRIBUTE_SET_FIELDS,
    access: "D/A CustomSecAttributeDefinition.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a custom-security-attribute role (Attribute Assignment Reader, Attribute Definition Reader, Attribute Assignment Administrator or Attribute Definition Administrator); by default Global Administrator and other administrator roles have no custom-security-attribute access. Personal Microsoft accounts are not supported.",
    note: "No P1/P2 prerequisite is stated for attribute-set reads.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/attributeset-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/attributeset?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/customSecurityAttributeDefinitions", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_CUSTOM_SECURITY_DEFINITION_FIELDS,
    access: "D/A CustomSecAttributeDefinition.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a custom-security-attribute role (Attribute Definition Reader, Attribute Assignment Administrator or Attribute Definition Administrator); by default Global Administrator and other administrator roles have no custom-security-attribute access. Personal Microsoft accounts are not supported.",
    note: "Filtering uses plain $filter (eq) with no $count or ConsistencyLevel contract; the allowedValues navigation property is not returned by default and needs its own list/show reads ($expand is not reviewed here). No P1/P2 prerequisite is stated for definition reads.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/directory-list-customsecurityattributedefinitions?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/customsecurityattributedefinition?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/customSecurityAttributeDefinitions/{customSecurityAttributeDefinition-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_CUSTOM_SECURITY_DEFINITION_FIELDS,
    access: "D/A CustomSecAttributeDefinition.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a custom-security-attribute role (Attribute Assignment Reader, Attribute Definition Reader, Attribute Assignment Administrator or Attribute Definition Administrator); by default Global Administrator and other administrator roles have no custom-security-attribute access. Personal Microsoft accounts are not supported.",
    note: "No P1/P2 prerequisite is stated for definition reads.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/customsecurityattributedefinition-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/customsecurityattributedefinition?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/customSecurityAttributeDefinitions/{customSecurityAttributeDefinition-id}/allowedValues", kind: "collection", query: ["$select"], fields: KNOWN_ALLOWED_VALUE_FIELDS,
    access: "D/A CustomSecAttributeDefinition.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a custom-security-attribute role (Attribute Assignment Reader, Attribute Definition Reader, Attribute Assignment Administrator or Attribute Definition Administrator); by default Global Administrator and other administrator roles have no custom-security-attribute access. Personal Microsoft accounts are not supported.",
    note: "Graph documents $select only for allowed-value lists; $filter is not supported. Definitions that allow free-form values carry no predefined values. No P1/P2 prerequisite is stated for allowed-value reads.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/customsecurityattributedefinition-list-allowedvalues?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/allowedvalue?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/customSecurityAttributeDefinitions/{customSecurityAttributeDefinition-id}/allowedValues/{allowedValue-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_ALLOWED_VALUE_FIELDS,
    access: "D/A CustomSecAttributeDefinition.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Attribute Definition Reader or Attribute Definition Administrator; by default Global Administrator and other administrator roles have no custom-security-attribute access. Personal Microsoft accounts are not supported.",
    note: "No P1/P2 prerequisite is stated for allowed-value reads.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/allowedvalue-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/allowedvalue?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contracts", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_CONTRACT_FIELDS,
    access: "D/A Directory.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers is the least-privileged role); personal Microsoft accounts are not supported. Contracts exist in partner tenants only.",
    note: "Filtering is documented for customerId, defaultDomainName and displayName; no P1/P2 prerequisite is stated for contract reads.",
    sources: ["https://learn.microsoft.com/graph/api/contract-list?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/contract?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminCustomers", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_DELEGATED_ADMIN_CUSTOMER_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. Customer objects are created by the system when a relationship exists and deleted when none remain, so a non-partner tenant lists zero customers. No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/tenantrelationship-list-delegatedadmincustomers?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadmincustomer?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminCustomers/{delegatedAdminCustomer-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DELEGATED_ADMIN_CUSTOMER_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/delegatedadmincustomer-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadmincustomer?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminRelationships", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_DELEGATED_ADMIN_RELATIONSHIP_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. Rows may carry the resellerDelegatedAdminRelationship subtype; its unreviewed extras are never projected. No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/tenantrelationship-list-delegatedadminrelationships?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadminrelationship?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DELEGATED_ADMIN_RELATIONSHIP_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "Single reads may return the resellerDelegatedAdminRelationship subtype; its unreviewed extras are never projected. No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/delegatedadminrelationship-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadminrelationship?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/accessAssignments", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_DELEGATED_ADMIN_ACCESS_ASSIGNMENT_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract; $top supports up to 300 objects. No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/delegatedadminrelationship-list-accessassignments?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadminaccessassignment?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/accessAssignments/{delegatedAdminAccessAssignment-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DELEGATED_ADMIN_ACCESS_ASSIGNMENT_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/delegatedadminaccessassignment-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadminaccessassignment?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/operations", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_DELEGATED_ADMIN_OPERATION_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract; $top supports up to 300 objects. The data payload is a JSON-encoded string. No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/delegatedadminrelationship-list-operations?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadminrelationshipoperation?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/operations/{delegatedAdminRelationshipOperation-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DELEGATED_ADMIN_OPERATION_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "The data payload is a JSON-encoded string. No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/delegatedadminrelationshipoperation-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadminrelationshipoperation?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/requests", kind: "collection", query: ["$select", "$filter", "$top"], fields: KNOWN_DELEGATED_ADMIN_REQUEST_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/delegatedadminrelationship-list-requests?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadminrelationshiprequest?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/requests/{delegatedAdminRelationshipRequest-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DELEGATED_ADMIN_REQUEST_FIELDS,
    access: "D/A DelegatedAdminRelationship.Read.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. Delegated-admin reads run in the partner tenant.",
    note: "No P1/P2 prerequisite is stated for delegated-admin reads.",
    sources: ["https://learn.microsoft.com/graph/api/delegatedadminrelationshiprequest-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/delegatedadminrelationshiprequest?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/multiTenantOrganization", kind: "single", query: SINGLE_QUERY, fields: KNOWN_MTO_FIELDS,
    access: "D/A MultiTenantOrganization.Read.All. Delegated callers pass it as --scopes (the lower-privileged delegated MultiTenantOrganization.ReadBasic.All is also accepted); delegated access additionally needs Security Reader or Global Reader, the least-privileged supported Entra roles; personal Microsoft accounts are not supported. Multi-tenant-organization reads run in the commercial Global service.",
    note: "At most one multitenant organization exists per tenant; a tenant outside any multitenant organization reads state inactive with null properties. Multi-tenant-organization participation needs Entra ID P1.",
    sources: ["https://learn.microsoft.com/graph/api/multitenantorganization-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/multitenantorganization?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/multiTenantOrganization/joinRequest", kind: "single", query: SINGLE_QUERY, fields: KNOWN_MTO_JOIN_REQUEST_FIELDS,
    access: "D/A MultiTenantOrganization.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Security Reader or Global Reader, the least-privileged supported Entra roles; personal Microsoft accounts are not supported. Multi-tenant-organization reads run in the commercial Global service.",
    note: "Join acceptance is a write and stays out of raw reads. Multi-tenant-organization participation needs Entra ID P1.",
    sources: ["https://learn.microsoft.com/graph/api/multitenantorganizationjoinrequestrecord-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/multitenantorganizationjoinrequestrecord?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/multiTenantOrganization/tenants", kind: "collection", query: ["$select", "$filter"], fields: KNOWN_MTO_TENANT_FIELDS,
    access: "D/A MultiTenantOrganization.Read.All. Delegated callers pass it as --scopes (the lower-privileged delegated MultiTenantOrganization.ReadBasic.All returns displayName and tenantId of active tenants only); delegated access additionally needs Security Reader or Global Reader, the least-privileged supported Entra roles; personal Microsoft accounts are not supported. Multi-tenant-organization reads run in the commercial Global service.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. Multi-tenant-organization participation needs Entra ID P1.",
    sources: ["https://learn.microsoft.com/graph/api/multitenantorganization-list-tenants?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/multitenantorganizationmember?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/findTenantInformationByDomainName(domainName='{domainName}')", kind: "single", query: [], fields: KNOWN_TENANT_INFORMATION_FIELDS,
    access: "D/A CrossTenantInformation.ReadBasic.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. No Entra role is required and no P1/P2 prerequisite is stated for this read.",
    note: "Tenant lookup by domain name: the raw path carries the domain as an OData-quoted function argument and the shared session re-validates, re-quotes and encodes it; only this allowlisted shape binds. The operation documents no query parameters, so the full reviewed tenantInformation set is always returned.",
    sources: ["https://learn.microsoft.com/graph/api/tenantrelationship-findtenantinformationbydomainname?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/tenantinformation?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/tenantRelationships/findTenantInformationByTenantId(tenantId='{tenantId}')", kind: "single", query: [], fields: KNOWN_TENANT_INFORMATION_FIELDS,
    access: "D/A CrossTenantInformation.ReadBasic.All. Delegated callers pass it as --scopes; application callers need it admin-consented; personal Microsoft accounts are not supported. No Entra role is required and no P1/P2 prerequisite is stated for this read.",
    note: "Tenant lookup by tenant ID: the raw path carries the GUID as an OData-quoted function argument and the shared session re-validates, re-quotes and encodes it; only this allowlisted shape binds. The operation documents no query parameters, so the full reviewed tenantInformation set is always returned.",
    sources: ["https://learn.microsoft.com/graph/api/tenantrelationship-findtenantinformationbytenantid?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/tenantinformation?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/contracts/{contract-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_CONTRACT_FIELDS,
    access: "D/A Directory.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (Directory Readers is the least-privileged role); personal Microsoft accounts are not supported. Contracts exist in partner tenants only.",
    note: "No P1/P2 prerequisite is stated for contract reads.",
    sources: ["https://learn.microsoft.com/graph/api/contract-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/contract?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identity/identityProviders", kind: "collection", query: COLLECTION_QUERY, fields: PROVIDER_FIELDS,
    access: "D/A IdentityProvider.Read.All. Delegated callers pass it as --scopes; delegated reads additionally need a directory role that can read federation configuration (Global Reader is the least-privileged read-only directory role). Personal Microsoft accounts are not supported; no per-operation licence prerequisite is stated.",
    note: "Workforce context only; external-customer user-flow provider bindings are separate scheduled operations. clientSecret and certificateData are never projected: the former is write-only and the latter is key material.",
    sources: ["https://learn.microsoft.com/graph/api/identitycontainer-list-identityproviders?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identity/identityProviders/{identityProviderBase-id}", kind: "single", query: SINGLE_QUERY, fields: PROVIDER_FIELDS,
    access: "D/A IdentityProvider.Read.All. Delegated callers pass it as --scopes; delegated reads additionally need a directory role that can read federation configuration (Global Reader is the least-privileged read-only directory role). Personal Microsoft accounts are not supported; no per-operation licence prerequisite is stated.",
    note: "Workforce context only. clientSecret and certificateData are never projected: the former is write-only and the latter is key material.",
    sources: ["https://learn.microsoft.com/graph/api/identityproviderbase-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/federationConfigurations", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_FEDERATION_FIELDS,
    access: "D/A Domain.Read.All least-privileged (Domain.ReadWrite.All is the documented higher-privileged alternative). Delegated callers pass one as --scopes; delegated access additionally needs External Identity Provider Administrator, the least-privileged supported Entra role for this operation; personal Microsoft accounts are not supported. No per-operation licence prerequisite is stated for these reads.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. The domains navigation property needs its own review ($expand is not reviewed here). signingCertificate carries the public token-signing key only and stays truncated at the output boundary unless --full is passed.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/samlorwsfedexternaldomainfederation-list?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/samlorwsfedexternaldomainfederation?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/directory/federationConfigurations/{identityProviderBase-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_FEDERATION_FIELDS,
    access: "D/A Domain.Read.All least-privileged (Domain.ReadWrite.All is the documented higher-privileged alternative). Delegated callers pass one as --scopes; delegated access additionally needs External Identity Provider Administrator, the least-privileged supported Entra role for this operation; personal Microsoft accounts are not supported. No per-operation licence prerequisite is stated for this read.",
    note: "The domains navigation property needs its own review ($expand is not reviewed here). signingCertificate carries the public token-signing key only; no private key material exists on this resource.",
    sources: ["https://learn.microsoft.com/en-us/graph/api/samlorwsfedexternaldomainfederation-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/samlorwsfedexternaldomainfederation?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/dataPolicyOperations", kind: "collection", query: ["$select"], fields: KNOWN_DATA_POLICY_FIELDS,
    access: "D/A User.Export.All plus User.Read.All. Delegated callers pass both as --scopes; delegated access additionally needs Company Administrator, the privileged role documented for export reads; personal Microsoft accounts are not supported. No P1/P2 prerequisite is stated for these reads.",
    note: "No list operation documentation page exists; the list shares the single-get permission contract and resource shape. Graph documents $select only here, so $filter is not reviewed. storageLocation always renders as the redaction marker: export blob URLs and signed links never reach output. Workforce context only; export submission belongs to no read slice.",
    sources: ["https://learn.microsoft.com/graph/api/datapolicyoperation-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/datapolicyoperation?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/dataPolicyOperations/{dataPolicyOperation-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DATA_POLICY_FIELDS,
    access: "D/A User.Export.All plus User.Read.All. Delegated callers pass both as --scopes; delegated access additionally needs Company Administrator, the privileged role documented for export reads; personal Microsoft accounts are not supported. No P1/P2 prerequisite is stated for this read.",
    note: "storageLocation always renders as the redaction marker: export blob URLs and signed links never reach output. Workforce context only; export submission belongs to no read slice.",
    sources: ["https://learn.microsoft.com/graph/api/datapolicyoperation-get?view=graph-rest-1.0", "https://learn.microsoft.com/graph/api/resources/datapolicyoperation?view=graph-rest-1.0"] },
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
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions/{accessReviewScheduleDefinition-id}/instances/{accessReviewInstance-id}/decisions/{accessReviewInstanceDecisionItem-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_DECISION_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    note: "Showing a decision never approves, denies or applies anything; access reviews need P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewinstancedecisionitem-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions/{accessReviewScheduleDefinition-id}/instances/{accessReviewInstance-id}/contactedReviewers", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_CONTACTED_REVIEWER_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    note: "Contacted reviewers are reviewer identities recorded on one instance, whether or not notified; access reviews need P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewinstance-list-contactedreviewers?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions/{accessReviewScheduleDefinition-id}/instances/{accessReviewInstance-id}/contactedReviewers/{accessReviewReviewer-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_CONTACTED_REVIEWER_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    note: "The single shares its list's reviewed resource contract: accessReviewReviewer carries no relationships; access reviews need P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewinstance-list-contactedreviewers?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/accessreviewreviewer?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions/{accessReviewScheduleDefinition-id}/instances/{accessReviewInstance-id}/stages", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_STAGE_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    note: "Stages are sequential phases of one instance, present only when stageSettings is defined; filters document eq only; access reviews need P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewinstance-list-stages?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/accessReviews/definitions/{accessReviewScheduleDefinition-id}/instances/{accessReviewInstance-id}/stages/{accessReviewStage-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_STAGE_FIELDS,
    access: "D/A AccessReview.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role (group/app reviews: review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator); personal Microsoft accounts are not supported.",
    note: "durationInDays is not a stage property (it folds into endDateTime); per-stage decisions belong to a later slice; access reviews need P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accessreviewstage-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/entitlementManagement/catalogs", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_CATALOG_FIELDS,
    access: "D/A EntitlementManagement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role with catalog visibility (Global Reader and Identity Governance Administrator are among the supported roles); personal Microsoft accounts are not supported.",
    note: "Catalogs are package containers: a catalog never carries its access packages; filtering passes through as plain $filter with no $count or ConsistencyLevel contract. Entitlement management needs P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/entitlementmanagement-list-catalogs?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/entitlementManagement/catalogs/{accessPackageCatalog-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_CATALOG_FIELDS,
    access: "D/A EntitlementManagement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role with catalog visibility (Global Reader and Identity Governance Administrator are among the supported roles); personal Microsoft accounts are not supported.",
    note: "Only scalar catalog properties are reviewed; accessPackages, resources, resourceScopes, resourceRoles and customWorkflowExtensions are navigation properties needing $expand and stay out. Entitlement management needs P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accesspackagecatalog-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/entitlementManagement/accessPackages", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_ACCESS_PACKAGE_FIELDS,
    access: "D/A EntitlementManagement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role with catalog visibility (Global Reader and Identity Governance Administrator are among the supported roles); personal Microsoft accounts are not supported.",
    note: "Access packages are the assignable bundles; assignment and request detail lives on the assignment reads, while approvals and subjects carry personal data and belong to a later part. Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. Entitlement management needs P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/entitlementmanagement-list-accesspackages?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_ACCESS_PACKAGE_FIELDS,
    access: "D/A EntitlementManagement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role with catalog visibility (Global Reader and Identity Governance Administrator are among the supported roles); personal Microsoft accounts are not supported.",
    note: "Only scalar package properties are reviewed; the catalog link and incompatible sets are navigation properties needing $expand and stay out. Entitlement management needs P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accesspackage-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/assignmentPolicies", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_ASSIGNMENT_POLICY_FIELDS,
    access: "D/A EntitlementManagement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role with catalog visibility (Global Reader and Identity Governance Administrator are among the supported roles); personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; access follows the top-level assignment-policy list contract and the accessPackageAssignmentPolicy resource reference. Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. Questions and custom-extension stages are navigation properties and stay out. Entitlement management needs P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/entitlementmanagement-list-assignmentpolicies?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/accesspackageassignmentpolicy?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/assignmentPolicies/{accessPackageAssignmentPolicy-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_ASSIGNMENT_POLICY_FIELDS,
    access: "D/A EntitlementManagement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role with catalog visibility (Global Reader and Identity Governance Administrator are among the supported roles); personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; access follows the top-level assignment-policy list contract and the accessPackageAssignmentPolicy resource reference. Entitlement management needs P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/entitlementmanagement-list-assignmentpolicies?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/accesspackageassignmentpolicy?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/resourceRoleScopes", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_RESOURCE_ROLE_SCOPE_FIELDS,
    access: "D/A EntitlementManagement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role with catalog visibility (Global Reader and Identity Governance Administrator are among the supported roles); personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; access follows the parent access-package contract and the accessPackageResourceRoleScope resource reference. Only the pairing identity is reviewed; the role and scope links need $expand and stay out. Filtering passes through as plain $filter with no $count or ConsistencyLevel contract. Entitlement management needs P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accesspackage-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/accesspackageresourcerolescope?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/resourceRoleScopes/{accessPackageResourceRoleScope-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_RESOURCE_ROLE_SCOPE_FIELDS,
    access: "D/A EntitlementManagement.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs a supported Entra role with catalog visibility (Global Reader and Identity Governance Administrator are among the supported roles); personal Microsoft accounts are not supported.",
    note: "No operation-level documentation page; access follows the parent access-package contract and the accessPackageResourceRoleScope resource reference. Entitlement management needs P2 or ID Governance depending on capability.",
    sources: ["https://learn.microsoft.com/graph/api/accesspackage-get?view=graph-rest-1.0", "https://learn.microsoft.com/en-us/graph/api/resources/accesspackageresourcerolescope?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/lifecycleWorkflows/workflows", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_WORKFLOW_FIELDS,
    access: "D/A LifecycleWorkflows-Workflow.ReadBasic.All (LifecycleWorkflows-Workflow.Read.All or LifecycleWorkflows.Read.All for richer detail). Delegated callers pass one as --scopes; delegated access additionally needs Global Reader or Lifecycle Workflows Administrator; personal Microsoft accounts are not supported.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract ($search/$orderby/$expand stay unreviewed). Only scalar workflow properties are reviewed; tasks, runs, processing results, versions, scopes and createdBy/lastModifiedBy links stay out. Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite.",
    sources: ["https://learn.microsoft.com/graph/api/identitygovernance-lifecycleworkflowscontainer-list-workflows?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/lifecycleWorkflows/workflows/{workflow-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_WORKFLOW_FIELDS,
    access: "D/A LifecycleWorkflows-Workflow.ReadBasic.All (LifecycleWorkflows-Workflow.Read.All or LifecycleWorkflows.Read.All for richer detail). Delegated callers pass one as --scopes; delegated access additionally needs Global Reader or Lifecycle Workflows Administrator; personal Microsoft accounts are not supported.",
    note: "Tasks ride expanded by default on the wire and are dropped in local projection; createdBy, lastModifiedBy and previewScope need $expand and stay out. Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite.",
    sources: ["https://learn.microsoft.com/graph/api/identitygovernance-workflow-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/lifecycleWorkflows/workflowTemplates", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_WORKFLOW_TEMPLATE_FIELDS,
    access: "D/A LifecycleWorkflows.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Global Reader or Lifecycle Workflows Administrator; personal Microsoft accounts are not supported.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract ($orderby stays unreviewed). Only scalar template properties are reviewed; template tasks belong to a later slice. Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite.",
    sources: ["https://learn.microsoft.com/graph/api/identitygovernance-lifecycleworkflowscontainer-list-workflowtemplates?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/lifecycleWorkflows/workflowTemplates/{workflowTemplate-id}", kind: "single", query: [], fields: KNOWN_WORKFLOW_TEMPLATE_FIELDS,
    access: "D/A LifecycleWorkflows.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Global Reader or Lifecycle Workflows Administrator; personal Microsoft accounts are not supported.",
    note: "Graph documents no query parameters on the template get, so only whole rows are reviewed here; template tasks belong to a later slice. Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite.",
    sources: ["https://learn.microsoft.com/graph/api/identitygovernance-workflowtemplate-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/lifecycleWorkflows/taskDefinitions", kind: "collection", query: COLLECTION_QUERY, fields: KNOWN_TASK_DEFINITION_FIELDS,
    access: "D/A LifecycleWorkflows.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Global Reader or Lifecycle Workflows Administrator; personal Microsoft accounts are not supported.",
    note: "Filtering passes through as plain $filter with no $count or ConsistencyLevel contract ($orderby stays unreviewed). Task definitions carry no relationships. Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite.",
    sources: ["https://learn.microsoft.com/graph/api/identitygovernance-lifecycleworkflowscontainer-list-taskdefinitions?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/lifecycleWorkflows/taskDefinitions/{taskDefinition-id}", kind: "single", query: SINGLE_QUERY, fields: KNOWN_TASK_DEFINITION_FIELDS,
    access: "D/A LifecycleWorkflows.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Global Reader or Lifecycle Workflows Administrator; personal Microsoft accounts are not supported.",
    note: "Only scalar definition properties are reviewed. Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite.",
    sources: ["https://learn.microsoft.com/graph/api/identitygovernance-taskdefinition-get?view=graph-rest-1.0"] },
  { id: "v1.0:GET:/identityGovernance/lifecycleWorkflows/settings", kind: "single", query: SINGLE_QUERY, fields: KNOWN_LIFECYCLE_SETTINGS_FIELDS,
    access: "D/A LifecycleWorkflows.Read.All. Delegated callers pass it as --scopes; delegated access additionally needs Global Reader or Lifecycle Workflows Administrator; personal Microsoft accounts are not supported.",
    note: "The tenant settings singleton carries no relationships; the update action belongs to no read slice. Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite.",
    sources: ["https://learn.microsoft.com/graph/api/identitygovernance-lifecyclemanagementsettings-get?view=graph-rest-1.0"] },
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
      if (!segment.length || /[()]/.test(segment)) return null;
      encodeGraphPathSegment(segment);
      params[name] = segment;
    } else if (/[{}]/.test(slot)) {
      const bound = matchFunctionArgument(template, slot, segment);
      if (!bound) return null;
      Object.assign(params, bound);
    } else if (slot.toLowerCase() !== segment.toLowerCase()) return null;
  }
  return params;
}

// Structural match of one allowlisted function-argument segment from a raw
// path: the template slot must be the exact allowlisted shape and the
// actual segment must carry the value as one OData-quoted literal (single
// quotes doubled). The unquoted value is validated here and bound for the
// session, which re-validates, re-quotes and encodes it canonically before
// credentials. Anything else is not this route.
function matchFunctionArgument(template: string, slot: string, segment: string): Record<string, string> | null {
  const binding = functionBindingFor(template);
  if (!binding || slot !== `${binding.functionName}(${binding.param}='{${binding.placeholder}}')`) return null;
  const head = `${binding.functionName}(${binding.param}=`;
  if (!segment.toLowerCase().startsWith(head.toLowerCase()) || !segment.endsWith(")")) return null;
  const literal = /^'((?:[^']|'')*)'$/.exec(segment.slice(head.length, -1))?.[1];
  if (literal === undefined) return null;
  const value = literal.replaceAll("''", "'");
  binding.validate(value);
  return { [binding.placeholder]: value };
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
  // Routes without a documented $select contract (the tenant-information
  // function lookups) never carry server field selection; every other
  // reviewed route defaults to its reviewed set.
  if (!route.query.includes("$select")) return;
  const select = query["$select"] ??= (route.defaultFields ?? route.fields).join(",");
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
  const kept = Object.fromEntries(Object.entries(value).filter(([key]) => route.fields.includes(key)));
  if (route.keepODataType === true && typeof value["@odata.type"] === "string") kept["@odata.type"] = value["@odata.type"];
  return kept;
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
  if (route.delegatedOnly === true && args.profile.mode === "application") {
    throw new AxiError(`GET ${pathname} needs a delegated profile; Graph documents no supported application permission for ${route.id}`, "VALIDATION_ERROR", [
      "mg-axi profile create --name <name> --tenant <tenant-id> --client <client-id> --cloud commercial",
    ]);
  }
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
    if (route.query.includes("$select")) query["$select"] ??= selection["$select"]!;
    checkQuery(route, query);
    if (localSelection) url.searchParams.delete("$select");
    else if (query["$select"] !== undefined) url.searchParams.set("$select", query["$select"]!);
    const response = await deps.transport({ ...request, url: url.toString() });
    if (response.status < 200 || response.status >= 300 || !response.body) return response;
    let body: unknown;
    try { body = JSON.parse(response.body); } catch { return response; }
    if (route.kind === "single") body = reviewedFields(outputRoute, body);
    else if (isRecord(body) && Array.isArray(body.value)) body = { ...body, value: body.value.map(row => reviewedFields(outputRoute, row)) };
    return { ...response, body: JSON.stringify(body), receivedBodyBytes: response.receivedBodyBytes ?? Buffer.byteLength(response.body, "utf8") };
  } });
  const selection = query ?? { ...session.cursorQuery(operation, args.cursor!) };
  checkQuery(route, selection);
  // Select-less routes project the full reviewed set; every other route
  // projects exactly the requested selection.
  const outputRoute = { ...route, fields: selection["$select"] !== undefined ? selection["$select"].split(",").map(field => field.trim()).filter(Boolean) : [...route.fields] };
  const full = !!args.full;
  if (route.kind === "single") {
    const body = await session.execute({ profile: args.profile, operation, params, query, scopes });
    const shaped = truncateForOutput(reviewedFields(outputRoute, body), full);
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
