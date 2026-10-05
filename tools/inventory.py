"""Offline operation discovery. This is build tooling, not the runtime catalogue."""

import argparse
import hashlib
import json
from pathlib import Path

import yaml
from jsonschema import Draft202012Validator

ROOT = Path(__file__).resolve().parents[1]
REVISION = "7b2914c8ad1340129f52aa785f13c074cb46fd7c"
METHODS = {"get", "post", "put", "patch", "delete", "head", "options"}
MAP_SOURCE = "https://learn.microsoft.com/en-us/graph/api/resources/identity-network-access-overview?view=graph-rest-1.0"
LICENCE_SOURCE = "https://learn.microsoft.com/en-us/entra/fundamentals/licensing"

# Only identity navigation properties belong to the Entra boundary.
USER_NAV = set("agreementAcceptances appConsentRequestsForApproval appRoleAssignedResources appRoleAssignments authentication createdObjects directReports extensions identities invitedBy joinedGroups licenseDetails manager memberOf oauth2PermissionGrants ownedDevices ownedObjects pendingAccessReviewInstances permissionGrants registeredDevices scopedRoleMemberOf scopedRoleMemberships serviceProvisioningErrors sponsorOf transitiveMemberOf transitiveReports onPremisesSyncBehavior sponsors authorizationInfo signInIdentifiers usageRights".split())
GROUP_NAV = set("appRoleAssignments createdOnBehalfOf endpoints extensions groupLifecyclePolicies members membersWithLicenseErrors memberOf owners permissionGrants serviceProvisioningErrors transitiveMembers transitiveMemberOf settings onPremisesSyncBehavior".split())
ROOT_SLICES = {
    "applications": "READ-07", "servicePrincipals": "READ-07",
    "directoryRoles": "READ-09", "directoryRoleTemplates": "READ-09",
    "devices": "READ-10", "directory": "EXT-01", "directoryObjects": "EXT-01",
    "directorySettings": "EXT-01", "directorySettingTemplates": "EXT-01",
    "settings": "EXT-01", "filteringPolicies": "EXT-04",
    "filterOperators": "EXT-03", "functions": "EXT-03",
    "groupSettings": "EXT-01", "groupSettingTemplates": "EXT-01",
    "organization": "EXT-01", "organizationSettings": "EXT-01",
    "domains": "EXT-01", "subscribedSkus": "EXT-01", "companySubscriptions": "EXT-01",
    "contacts": "EXT-01", "onPremisesDirectorySynchronization": "EXT-01",
    "policies": "EXT-01", "identityProviders": "EXT-03", "invitations": "EXT-03",
    "identity": "EXT-03", "identityGovernance": "EXT-02",
    "identityProtection": "READ-06", "auditLogs": "READ-05",
    "tenantRelationships": "EXT-04", "contracts": "EXT-04",
    "networkAccess": "EXT-04", "agentRegistry": "EXT-04",
    "externalUserProfiles": "EXT-03", "pendingExternalUserProfiles": "EXT-03",
    "dataPolicyOperations": "EXT-03", "certificateBasedAuthConfiguration": "EXT-01",
    "trustFramework": "EXT-03",
    "administrativeUnits": "READ-10", "applicationTemplates": "READ-07",
    "agreements": "EXT-01", "agreementAcceptances": "EXT-01",
    "schemaExtensions": "EXT-01", "domainDnsRecords": "EXT-01",
    "allowedDataLocations": "EXT-01", "groupLifecyclePolicies": "EXT-01",
    "mobilityManagementPolicies": "EXT-01", "onPremisesPublishingProfiles": "EXT-04",
    "permissionGrants": "READ-08", "scopedRoleMemberships": "READ-09",
    "authenticationMethodsPolicy": "READ-04", "authenticationMethodConfigurations": "READ-04",
    "authenticationMethodDevices": "READ-04", "certificateAuthorities": "EXT-01",
    "riskyUsers": "READ-06", "riskDetections": "READ-06",
    "accessReviews": "EXT-02", "accessReviewDecisions": "EXT-02",
    "programs": "EXT-02", "programControls": "EXT-02", "programControlTypes": "EXT-02",
    "businessFlowTemplates": "EXT-02", "governanceResources": "EXT-02",
    "governanceRoleAssignments": "EXT-02", "governanceRoleAssignmentRequests": "EXT-02",
    "governanceRoleDefinitions": "EXT-02", "governanceRoleSettings": "EXT-02",
    "governanceSubjects": "EXT-02", "privilegedRoles": "READ-09",
    "privilegedRoleAssignments": "READ-09", "privilegedRoleAssignmentRequests": "READ-09",
    "privilegedApproval": "READ-09", "privilegedOperationEvents": "READ-09",
    "privilegedSignupStatus": "READ-09",
}
USER_ACTIONS = set("assignLicense changePassword checkMemberGroups checkMemberObjects convertExternalToInternalMemberUser deletePasswordSingleSignOnCredentials exportPersonalData getMemberGroups getMemberObjects getPasswordSingleSignOnCredentials invalidateAllRefreshTokens reprocessLicenseAssignment restore retryServiceProvisioning revokeSignInSessions validatePassword".split())
GROUP_ACTIONS = set("assignLicense checkGrantedPermissionsForApp checkMemberGroups checkMemberObjects deletePasswordSingleSignOnCredentials evaluateDynamicMembership getMemberGroups getMemberObjects getPasswordSingleSignOnCredentials getByIds renew restore retryServiceProvisioning validateProperties".split())
DIRECTORY_READ_ACTIONS = set("checkMemberGroups checkMemberObjects getMemberGroups getMemberObjects getByIds getAvailableExtensionProperties getUserOwnedObjects validateProperties".split())
USER_ACTIONS |= DIRECTORY_READ_ACTIONS
GROUP_ACTIONS |= DIRECTORY_READ_ACTIONS
READ_ACTION_SOURCES = {
    **{action: f"directoryobject-{action.lower()}" for action in DIRECTORY_READ_ACTIONS},
    "getUserOwnedObjects": "directory-deleteditems-getuserownedobjects",
    "evaluateDynamicMembership": "group-evaluatedynamicmembership",
    "checkGrantedPermissionsForApp": "group-checkgrantedpermissionsforapp",
    "evaluate": "conditionalaccessroot-evaluate",
    "getPasswordSingleSignOnCredentials": "serviceprincipal-getpasswordsinglesignoncredentials",
    "getApplicablePolicyRequirements": "accesspackage-getapplicablepolicyrequirements",
    "validateAuthenticationConfiguration": "customauthenticationextension-validateauthenticationconfiguration",
    "validateCredentials": "synchronization-synchronizationjob-validatecredentials",
    "validatePassword": "user-validatepassword",
    "parseExpression": "synchronization-synchronizationschema-parseexpression",
    "previewTaskFailures": "identitygovernance-workflow-previewtaskfailures",
    "tenantSearch": "managedtenants-tenantgroup-tenantsearch",
}
# The five approved EXT-01b organization/branding reads, shared by both API versions.
APPROVED_ORGANIZATION_READS = {
    "/organization",
    "/organization/{organization-id}",
    "/organization/{organization-id}/branding",
    "/organization/{organization-id}/branding/localizations",
    "/organization/{organization-id}/branding/localizations/{organizationalBrandingLocalization-id}",
}
# The three approved EXT-04a v1.0 partner-contract reads: list, show and count.
APPROVED_CONTRACT_READS = {
    "/contracts",
    "/contracts/{contract-id}",
    "/contracts/$count",
}
# The twelve approved EXT-04b/EXT-04d v1.0 delegated-admin reads: customer
# and relationship list/show plus the access-assignment, operation, request
# and service-management-detail navigation list/show reads. Counts, the
# container root and the multi-tenant-organization/lookup functions stay
# scheduled for later EXT-04 subfamilies.
APPROVED_DELEGATED_ADMIN_READS = {
    "/tenantRelationships/delegatedAdminCustomers",
    "/tenantRelationships/delegatedAdminCustomers/{delegatedAdminCustomer-id}",
    "/tenantRelationships/delegatedAdminCustomers/{delegatedAdminCustomer-id}/serviceManagementDetails",
    "/tenantRelationships/delegatedAdminCustomers/{delegatedAdminCustomer-id}/serviceManagementDetails/{delegatedAdminServiceManagementDetail-id}",
    "/tenantRelationships/delegatedAdminRelationships",
    "/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}",
    "/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/accessAssignments",
    "/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/accessAssignments/{delegatedAdminAccessAssignment-id}",
    "/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/operations",
    "/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/operations/{delegatedAdminRelationshipOperation-id}",
    "/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/requests",
    "/tenantRelationships/delegatedAdminRelationships/{delegatedAdminRelationship-id}/requests/{delegatedAdminRelationshipRequest-id}",
}
# The four approved EXT-04c v1.0 multi-tenant-organization reads:
# organization show, join-request show, tenant list and tenant count. The
# single-member read stays scheduled: its documented least privilege is the
# write scope.
APPROVED_MTO_READS = {
    "/tenantRelationships/multiTenantOrganization",
    "/tenantRelationships/multiTenantOrganization/joinRequest",
    "/tenantRelationships/multiTenantOrganization/tenants",
    "/tenantRelationships/multiTenantOrganization/tenants/$count",
}
# The two approved mg-ext-04e v1.0 tenant-lookup function reads: tenant
# information by domain name and by tenant ID. Both bind only through the
# allowlisted session function-argument contract.
APPROVED_TENANT_LOOKUP_READS = {
    "/tenantRelationships/findTenantInformationByDomainName(domainName='{domainName}')",
    "/tenantRelationships/findTenantInformationByTenantId(tenantId='{tenantId}')",
}
# The ten approved EXT-01m v1.0 contact manager/directReports navigation
# reads: the manager single, the directReports collection and scalar plus
# their user/contact typed casts, and the per-report singles plus casts.
APPROVED_CONTACT_NAV_READS = {
    "/contacts/{orgContact-id}/manager",
    "/contacts/{orgContact-id}/directReports",
    "/contacts/{orgContact-id}/directReports/$count",
    "/contacts/{orgContact-id}/directReports/graph.orgContact",
    "/contacts/{orgContact-id}/directReports/graph.orgContact/$count",
    "/contacts/{orgContact-id}/directReports/graph.user",
    "/contacts/{orgContact-id}/directReports/graph.user/$count",
    "/contacts/{orgContact-id}/directReports/{directoryObject-id}",
    "/contacts/{orgContact-id}/directReports/{directoryObject-id}/graph.orgContact",
    "/contacts/{orgContact-id}/directReports/{directoryObject-id}/graph.user",
}
# The six approved EXT-01c v1.0 group lifecycle reads: list, show and count
# per family (group lifecycle policies and group setting templates).
APPROVED_GROUP_LIFECYCLE_READS = {
    "/groupLifecyclePolicies",
    "/groupLifecyclePolicies/{groupLifecyclePolicy-id}",
    "/groupLifecyclePolicies/$count",
    "/groupSettingTemplates",
    "/groupSettingTemplates/{groupSettingTemplate-id}",
    "/groupSettingTemplates/$count",
}
# Beta mirrors of the twelve approved EXT-02d v1.0 entitlement-management
# reads (catalog, access-package, assignment-policy and resource-role-scope
# list/show/count; beta names catalogs and assignment policies
# accessPackageCatalogs/accessPackageAssignmentPolicies). The v1.0 reads
# themselves need no rule: they keep their scheduled discovery rows while
# the reviewed command and raw contracts live beside each command.
BETA_ENTITLEMENT_MIRROR_READS = {
    "/identityGovernance/entitlementManagement/accessPackageCatalogs",
    "/identityGovernance/entitlementManagement/accessPackageCatalogs/{accessPackageCatalog-id}",
    "/identityGovernance/entitlementManagement/accessPackageCatalogs/$count",
    "/identityGovernance/entitlementManagement/accessPackages",
    "/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}",
    "/identityGovernance/entitlementManagement/accessPackages/$count",
    "/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/accessPackageAssignmentPolicies",
    "/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/accessPackageAssignmentPolicies/{accessPackageAssignmentPolicy-id}",
    "/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/accessPackageAssignmentPolicies/$count",
    "/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/accessPackageResourceRoleScopes",
    "/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/accessPackageResourceRoleScopes/{accessPackageResourceRoleScope-id}",
    "/identityGovernance/entitlementManagement/accessPackages/{accessPackage-id}/accessPackageResourceRoleScopes/$count",
}
MANAGED_TENANT_NAV = set("auditEvents conditionalAccessPolicyCoverages credentialUserRegistrationsSummaries myRoles tenantGroups tenantTags tenants tenantsCustomizedInformation tenantsDetailedInformation".split())
EXCLUDED_NAV = {
    "Mail": set("mailboxSettings messages mailFolders calendars calendar contactFolders outlook".split()),
    "Files": {"drive", "drives"},
    "Teams": set("joinedTeams chats teamwork".split()),
    "Intune": set("managedDevices managedAppRegistrations deviceManagement setMobileDeviceManagementAuthority cloudPcConnections cloudPcDevices cloudPcsOverview aggregatedPolicyCompliances appPerformances deviceAppPerformances deviceCompliancePolicySettingStateSummaries deviceHealthStatuses managedDeviceComplianceTrends managedDeviceCompliances".split()),
    "Security": {"security", "windowsDeviceMalwareStates", "windowsProtectionStates"},
    "M365": set("contactInsights itemInsights peopleInsights microsoftApplicationDataAccess managedTenantAlertLogs managedTenantAlertRuleDefinitions managedTenantAlertRules managedTenantAlerts managedTenantApiNotifications managedTenantEmailNotifications managedTenantTicketingEndpoints managementActionTenantDeploymentStatuses managementActions managementIntents managementTemplateCollectionTenantSummaries managementTemplateCollections managementTemplateStepTenantSummaries managementTemplateStepVersions managementTemplateSteps managementTemplates".split()),
}


def exclusion_reason(path):
    segments = {part.removeprefix("microsoft.graph.").split("(")[0] for part in path.strip("/").split("/")}
    for domain, navigations in EXCLUDED_NAV.items():
        if segments & navigations:
            return f"Out-of-pack {domain} navigation requires separate authorization."
    return None


def scoped_slice(path):
    """Return the owning family, or None for a separately authorized domain."""
    parts = path.strip("/").split("/")
    if exclusion_reason(path):
        return None
    root = parts[0].split("(")[0]
    if root == "tenantRelationships" and len(parts) > 1 and parts[1] == "managedTenants":
        return "EXT-04" if len(parts) > 2 and parts[2].split("(")[0] in MANAGED_TENANT_NAV else None
    if root == "templates":
        return "READ-10" if len(parts) > 1 and parts[1] == "deviceTemplates" else None
    if root == "admin":
        return "EXT-01" if parts[1:3] == ["entra", "uxSetting"] else None
    if root == "me":
        root = "users"
        parts = ["users", "{user-id}", *parts[1:]]
    if root in {"users", "groups"}:
        owner = "READ-01" if root == "users" else "READ-02"
        nav = USER_NAV if root == "users" else GROUP_NAV
        actions = USER_ACTIONS if root == "users" else GROUP_ACTIONS
        tail = parts[1:]
        if tail and (tail[0].startswith("{") or tail[0] == "microsoft.graph.user" or tail[0] == "microsoft.graph.group"):
            tail = tail[1:]
        if tail:
            first = tail[0].removeprefix("microsoft.graph.").split("(")[0]
            if first not in nav | actions | {"$count", "delta"}:
                return None
            if first == "authentication":
                return "READ-04"
            if first in {"appRoleAssignments", "oauth2PermissionGrants"}:
                return "READ-08"
        return owner
    if root == "roleManagement":
        if len(parts) > 1 and parts[1] == "entitlementManagement":
            return "EXT-02"
        return "READ-09" if len(parts) == 1 or parts[1] == "directory" else None
    if root == "privilegedAccess":
        return "READ-09" if len(parts) == 1 or parts[1].startswith(("aad", "microsoft.graph.aad")) else None
    if root == "oauth2PermissionGrants":
        return "READ-08"
    if root == "reports":
        if len(parts) > 1 and parts[1].startswith("authenticationMethods"):
            return "READ-04"
        if len(parts) > 1 and (parts[1] in {"azureADPremiumLicenseInsight", "correlations", "healthMonitoring", "identityAnalytics", "sla"} or any(word in parts[1].lower() for word in ("credential", "registration", "application", "serviceprincipal", "agent", "directory", "authentication", "signin", "tenant", "permission", "relyingparty"))):
            return "EXT-04"
        return None
    if root == "informationProtection":
        return "READ-10" if len(parts) > 1 and parts[1] == "bitlocker" else None
    if root == "identity" and len(parts) > 1 and parts[1] == "conditionalAccess":
        return "READ-03"
    if root == "directory" and len(parts) > 1 and parts[1] == "administrativeUnits":
        return "READ-10"
    if root in {"applications", "servicePrincipals"} and "synchronization" in parts:
        return "EXT-03"
    if root == "servicePrincipals" and "appRoleAssignments" in parts:
        return "READ-08"
    return ROOT_SLICES.get(root)


def discover(source):
    """Parse one path at a time rather than retaining Graph's response schemas."""
    block = []
    in_paths = False
    with source.open() as stream:
        for line in stream:
            if line == "paths:\n":
                in_paths = True
                continue
            if not in_paths:
                continue
            if not line.startswith(" ") and line.strip():
                break
            if line.startswith("  ") and len(line) > 2 and line[2] != " " and block:
                yield from parse_path(block)
                block = []
            block.append(line)
    if block:
        yield from parse_path(block)


def parse_path(block):
    path, item = next(iter(yaml.load("".join(block), Loader=yaml.CSafeLoader).items()))
    for method, operation in item.items():
        if method in METHODS:
            yield path, method.upper(), operation


def make_row(version, path, method, operation):
    owner = scoped_slice(path)
    disposition, reason = "scheduled", "No implemented command or reviewed raw contract yet."
    action = path.rsplit("/", 1)[-1].rsplit(".", 1)[-1]
    read_source = READ_ACTION_SOURCES.get(action) if method == "POST" else None
    mutates = method not in {"GET", "HEAD", "OPTIONS"} and read_source is None
    lower = path.lower()
    secret = any(word in lower for word in ("addpassword", "passwordprofile", "getpassword", "resetpassword", "generatepassword", "uploadsecret", "getsecret", "/secrets", "devicelocalcredentials/", "bitlocker/recoverykeys/")) and not lower.endswith("/$count")
    secret = secret or (method == "POST" and "temporaryaccesspassmethods" in lower)
    if owner is None:
        disposition, reason = "excluded", exclusion_reason(path) or "Unscoped operation outside the accepted Entra boundary."
    elif operation.get("deprecated") or "permissionsmanagement" in lower:
        disposition, reason = "deprecated", "Deprecated metadata or multicloud permissions management."
    elif secret or "trustframework/keysets" in lower:
        disposition, reason = "intentionally-blocked", "Credential values, recovery keys, LAPS passwords or secret minting."
    elif version == "beta" and mutates:
        disposition, reason = "intentionally-blocked", "Beta writes are denied by the approved plan."
    elif any(word in lower for word in ("b2c", "authenticationeventsflows", "trustframework/policies")):
        disposition, reason = "intentionally-blocked", "External-customer launch support requires separate authorization."
    if owner == "EXT-01" and not mutates and disposition == "scheduled" and path.split("/")[1] == "organization" and path not in APPROVED_ORGANIZATION_READS:
        if "/certificateBasedAuthConfiguration" in path:
            reason = "Deferred by firstmate organization scope to a later EXT-01 organization certificate-auth subfamily: trusted-CA certificate material needs a separate redaction and output review."
        elif "/extensions" in path:
            reason = "Deferred by firstmate organization scope to a later EXT-01 organization extensions subfamily: open-extension payloads need their own query and projection review."
        elif path.rsplit("/", 1)[-1] in {"backgroundImage", "bannerLogo", "customCSS", "favicon", "headerLogo", "squareLogo", "squareLogoDark"}:
            reason = "Deferred by firstmate organization scope to a later EXT-01 organization branding-stream subfamily: binary image and CSS bytes need a separate output contract from the approved metadata reads."
        elif path.endswith(("/$count", "/delta()")):
            reason = "Deferred by firstmate organization scope to a later EXT-01 organization counts subfamily: scalar counts and deltas need a separate query and response contract from the approved list/show reads."
        elif method == "POST":
            reason = "Deferred by firstmate organization scope to a later EXT-01 organization lookup-actions subfamily: membership-check and lookup POST actions need their own request and projection review."
        elif version == "beta":
            reason = "Deferred by firstmate organization scope to a later EXT-01 beta organization subfamily: the approved reads cover the shared routes on both versions; beta-only settings, partner and theme contracts need separate review."
    if owner == "EXT-04" and not mutates and disposition == "scheduled" and path.split("/")[1] == "contracts" and (path not in APPROVED_CONTRACT_READS or version == "beta"):
        if path.endswith("/delta()"):
            reason = "Deferred by firstmate contracts scope to a later EXT-04 contracts delta subfamily: delta-token sync needs its own paging and change-tracking contract beyond the approved list/show/count reads."
        elif method == "POST":
            reason = "Deferred by firstmate contracts scope to a later EXT-04 contracts lookup-actions subfamily: membership-check and lookup POST actions need their own request and projection review."
        elif version == "beta":
            reason = "Deferred by firstmate contracts scope to a later EXT-04 beta contracts subfamily: the three approved reads cover v1.0 only; beta contracts need separate review."
    if owner == "EXT-01" and not mutates and disposition == "scheduled" and path.split("/")[1] in {"groupLifecyclePolicies", "groupSettingTemplates"} and (path not in APPROVED_GROUP_LIFECYCLE_READS or version == "beta"):
        if path.endswith("/delta()"):
            reason = "Deferred by firstmate group-lifecycle scope to a later EXT-01 group-lifecycle delta subfamily: delta-token sync needs its own paging and change-tracking contract beyond the approved list/show/count reads."
        elif method == "POST":
            reason = "Deferred by firstmate group-lifecycle scope to a later EXT-01 group-lifecycle lookup-actions subfamily: membership-check and lookup POST actions need their own request and projection review."
        elif version == "beta":
            reason = "Deferred by firstmate group-lifecycle scope to a later EXT-01 group-lifecycle beta subfamily: the six approved reads cover v1.0 only; beta policies and templates need separate review."
    if owner == "EXT-01" and not mutates and disposition == "scheduled" and path.split("/")[1] == "directory" and ("/attributeSets" in path or "/customSecurityAttributeDefinitions" in path) and version == "beta":
        reason = "Deferred by firstmate custom-security-attributes scope to a later EXT-01 custom-security-attributes beta subfamily: the nine approved reads cover v1.0 only; beta attribute sets, definitions and allowed values need separate review."
    if owner == "EXT-02" and method == "GET" and version == "beta" and disposition == "scheduled" and path in BETA_ENTITLEMENT_MIRROR_READS:
        reason = "Deferred by firstmate entitlement-management scope to a later EXT-02 entitlement-management beta subfamily: the twelve approved reads cover v1.0 only; beta catalogs, access packages, assignment policies and resource-role scopes need separate review."
    if owner == "EXT-04" and method == "GET" and version == "v1.0" and disposition == "scheduled" and path.split("/")[1] == "tenantRelationships" and path not in APPROVED_DELEGATED_ADMIN_READS and path not in APPROVED_MTO_READS and path not in APPROVED_TENANT_LOOKUP_READS:
        if path == "/tenantRelationships/multiTenantOrganization/tenants/{multiTenantOrganizationMember-id}":
            reason = "Deferred by firstmate multi-tenant-organization scope to a later EXT-04 multi-tenant-organization subfamily: documented least privilege is the write scope MultiTenantOrganization.ReadWrite.All in both delegated and application modes, so no named command or raw access is approved for the single-member read."
        elif path.startswith("/tenantRelationships/multiTenantOrganization"):
            reason = "Deferred by firstmate delegated-admin scope to a later EXT-04 multi-tenant-organization subfamily: multi-tenant-organization reads need their own MultiTenantOrganization access, Entra P1 licensing and projection review beyond the approved delegated-admin list/show reads."
        elif "findTenantInformation" in path:
            reason = "Deferred by firstmate delegated-admin scope to a later EXT-04 tenant-lookup subfamily: tenant-information lookup functions need their own CrossTenantInformation access and explicit domain/tenant-argument review beyond the approved delegated-admin list/show reads."
        elif path.endswith("/$count"):
            reason = "Deferred by firstmate delegated-admin scope to a later EXT-04 delegated-admin counts subfamily: scalar counts need a separate query and response contract from the approved list/show reads."
        elif path == "/tenantRelationships":
            reason = "Deferred by firstmate delegated-admin scope to a later EXT-04 delegated-admin root subfamily: the tenantRelationship container singleton needs its own projection review beyond the approved customer and relationship list/show reads."
        elif "/serviceManagementDetails" in path:
            reason = "Deferred by firstmate delegated-admin scope to a later EXT-04 delegated-admin customer-details subfamily: service-management detail reads need their own query, access and projection review beyond the approved customer list/show reads."
        elif "/accessAssignments" in path or "/operations" in path or "/requests" in path:
            reason = "Deferred by firstmate delegated-admin scope to a later EXT-04 delegated-admin relationship-details subfamily: access-assignment, operation and request navigation reads need their own query, access and projection review beyond the approved relationship list/show reads."
    if owner == "EXT-02" and method == "GET" and disposition == "scheduled" and "/accessReviews/historyDefinitions" in path:
        disposition, reason = "intentionally-blocked", "Blocked: documented least privilege is the write scope AccessReview.ReadWrite.All (no read scope), and history instances return SAS download URLs in downloadUri; recording or emitting that URL needs its own redaction and output review."
        owner = "EXT-02c"
    if owner == "EXT-02" and method == "GET" and disposition == "scheduled" and (path.split("/")[1] in {"accessReviews", "accessReviewDecisions"} or "/accessReviews/unified" in path):
        owner = "EXT-02c"
        reason = "Split into EXT-02c to limit this piece's size: legacy accessReviews and unified alias reads need their own query, access and projection review; beta contracts need separate review."
    if owner == "EXT-01" and method == "GET" and disposition == "scheduled" and path.split("/")[1] in {"domains", "domainDnsRecords"}:
        if "/federationConfiguration" in path:
            reason = "Deferred by firstmate R1 to a later EXT-01 domain federation subfamily: federation configuration can carry signing-certificate material and needs a separate output review."
        elif "/domainNameReferences" in path or path.endswith("/rootDomain"):
            reason = "Deferred by firstmate R1 to a later EXT-01 domain relationships subfamily: directory-object references and root-domain navigation need their own query and projection review."
        elif "/sharedEmailDomainInvitations" in path:
            reason = "Deferred by firstmate R1 to a later EXT-01 shared-email domains subfamily: beta invitation relationships need a separate access and output review."
        elif path.endswith("/$count"):
            reason = "Deferred by firstmate R1 to a later EXT-01 domain counts subfamily: scalar counts need a separate query and response contract from the eight approved list/show reads."
        elif version == "beta":
            reason = "Deferred by firstmate R1 to a later EXT-01 beta domains subfamily: the eight approved reads cover v1.0 only; beta contracts need separate review."
    if owner == "EXT-03" and method == "GET" and disposition == "scheduled" and (path == "/identity/identityProviders" or path.startswith("/identity/identityProviders/")):
        if version == "beta":
            reason = "Deferred to a later EXT-03 beta identity-providers subfamily: Identity-provider reads support v1.0 only; beta needs its own review."
    if owner == "EXT-01" and method == "GET" and version == "beta" and disposition == "scheduled" and (path == "/directory/federationConfigurations" or path.startswith("/directory/federationConfigurations")):
        reason = "Deferred by firstmate federation-configuration scope to a later EXT-01 federation-configuration beta subfamily: the four approved reads cover v1.0 only; beta federation configurations need separate review."
    if owner == "EXT-03" and method == "GET" and version == "beta" and disposition == "scheduled" and (path == "/dataPolicyOperations" or path.startswith("/dataPolicyOperations")):
        reason = "Deferred by firstmate data-policy-operations scope to a later EXT-03 data-policy-operations beta subfamily: the three approved reads cover v1.0 only; beta data-policy operations need separate review."
    if owner == "EXT-03" and method == "GET" and version == "v1.0" and disposition == "scheduled" and (path == "/invitations" or path.startswith("/invitations/")):
        disposition, reason = "unavailable", "Marked unavailable by firstmate EXT-03b decision: The v1.0 invitation resource Methods table documents Create only, with no documented GET contract for these invitation reads (https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/v1.0/resources/invitation.md)."
    if owner == "EXT-01" and method == "GET" and version == "v1.0" and disposition == "scheduled" and (path == "/certificateBasedAuthConfiguration" or path.startswith("/certificateBasedAuthConfiguration/")):
        disposition, reason = "unavailable", "Marked unavailable by firstmate mg-ext-01d decision: the v1.0 certificateBasedAuthConfiguration resource Methods table documents List/Create/Get/Delete only on the org-scoped /organization/{organization-id}/certificateBasedAuthConfiguration routes, with no documented GET contract for these root reads (https://learn.microsoft.com/en-us/graph/api/resources/certificatebasedauthconfiguration?view=graph-rest-1.0)."
    if owner == "EXT-01" and method == "GET" and version == "v1.0" and disposition == "scheduled" and path == "/directory/onPremisesSynchronization/$count":
        reason = "Deferred by firstmate on-premises-synchronization scope to a later EXT-01 on-premises-synchronization counts subfamily: scalar counts need a separate query and response contract from the approved list/show reads."
    if owner == "EXT-01" and method == "GET" and version == "beta" and disposition == "scheduled" and path.split("/")[1] == "directory" and "/onPremisesSynchronization" in path:
        reason = "Deferred by firstmate on-premises-synchronization scope to a later EXT-01 on-premises-synchronization beta subfamily: the two approved reads cover v1.0 only; beta on-premises synchronization needs separate review."
    if owner == "EXT-01" and method == "GET" and version == "v1.0" and disposition == "scheduled" and path == "/directoryObjects/delta()":
        reason = "Deferred by firstmate directory-objects scope to a later EXT-01 directory-objects delta subfamily: delta-token sync needs its own change-tracking contract beyond the approved list/show/count reads."
    if owner == "EXT-01" and method == "GET" and version == "beta" and disposition == "scheduled" and path.split("/")[1] == "directoryObjects":
        reason = "Deferred by firstmate directory-objects scope to a later EXT-01 directory-objects beta subfamily: the three approved reads cover v1.0 only; beta directory objects need separate review."
    if owner == "EXT-01" and method == "POST" and disposition == "scheduled" and "/directoryObjects" in path and path != "/directoryObjects" and not path.endswith("/restore"):
        reason = "Deferred by firstmate directory-objects scope to a later EXT-01 directory-objects lookup-actions subfamily: the POST lookup and validation actions need their own request and projection review."
    if owner == "EXT-01" and method == "GET" and version == "v1.0" and disposition == "scheduled" and path.split("/")[1] == "directory" and "/deletedItems" in path:
        if path in ("/directory/deletedItems", "/directory/deletedItems/$count"):
            reason = "Deferred by firstmate deleted-items scope: upstream requires the OData cast as part of the list URI and states calling GET /directory/deletedItems without a type is not supported, so no untyped list or count command is approved (https://learn.microsoft.com/en-us/graph/api/directory-deleteditems-list?view=graph-rest-1.0)."
        elif "/graph.device" in path:
            reason = "Deferred by firstmate deleted-items scope: the v1.0 deletedItems list/get pages name the supported types and document per-type least-privilege permissions, and device is not among them, so no device cast command is approved (https://learn.microsoft.com/en-us/graph/api/directory-deleteditems-list?view=graph-rest-1.0)."
        elif path.startswith("/directory/deletedItems/{directoryObject-id}/graph."):
            reason = "Deferred by firstmate deleted-items scope: covered by the approved deleted-item show, which reads the untyped get returning the same object for every type; the cast suffix only asserts the expected type and no separate reviewed route is approved."
    if owner == "EXT-01" and method == "GET" and version == "beta" and disposition == "scheduled" and path.split("/")[1] == "directory" and "/deletedItems" in path:
        reason = "Deferred by firstmate deleted-items scope to a later EXT-01 deleted-items beta subfamily: the eleven approved reads cover v1.0 only; beta deleted items need separate review."
    if owner == "EXT-01" and method == "POST" and not mutates and disposition == "scheduled" and path.split("/")[1] == "directory" and "/deletedItems" in path and not path.endswith("/restore"):
        reason = "Deferred by firstmate deleted-items scope to a later EXT-01 deleted-items lookup-actions subfamily: the POST lookup and validation actions need their own request and projection review."
    if owner == "EXT-01" and method == "GET" and version == "v1.0" and disposition == "scheduled" and path == "/contacts/delta()":
        reason = "Deferred by firstmate contacts scope to a later EXT-01 contacts delta subfamily: delta-token sync needs its own change-tracking contract beyond the approved list/show/count reads."
    if owner == "EXT-01" and method == "GET" and version == "v1.0" and disposition == "scheduled" and path in {
            "/contacts/{orgContact-id}/serviceProvisioningErrors",
            "/contacts/{orgContact-id}/serviceProvisioningErrors/$count",
            "/contacts/{orgContact-id}/onPremisesSyncBehavior"}:
        disposition, reason = "unavailable", "Marked unavailable by firstmate mg-ext-01m decision: no operation-level permission documentation exists for these reads (the guessed operation pages return 404 and onPremisesSyncBehavior is absent from the orgContact resource page), so no read-scope contract can be established (https://learn.microsoft.com/en-us/graph/api/resources/orgcontact?view=graph-rest-1.0)."
    if owner == "EXT-01" and method == "GET" and version == "v1.0" and disposition == "scheduled" and path.startswith("/contacts/{orgContact-id}/") and path not in APPROVED_CONTACT_NAV_READS:
        reason = "Split into EXT-01n (mg-ext-01n) to limit this piece's size: memberOf and transitiveMemberOf reads need their own multi-scope and advanced-query review; the ten approved manager/directReports navigation reads ship in mg-ext-01m."
    if owner == "EXT-01" and method == "GET" and version == "beta" and disposition == "scheduled" and path.split("/")[1] == "contacts":
        reason = "Deferred by firstmate contacts scope to a later EXT-01 contacts beta subfamily: the three approved reads cover v1.0 only; beta contacts need separate review."
    if owner == "EXT-01" and method == "POST" and not mutates and disposition == "scheduled" and "/contacts" in path and path != "/contacts" and not path.endswith("/restore"):
        reason = "Deferred by firstmate contacts scope to a later EXT-01 contacts lookup-actions subfamily: the POST lookup and validation actions need their own request and projection review."
    if owner == "EXT-01" and method == "GET" and disposition == "scheduled" and path.split("/")[1] in {"agreements", "agreementAcceptances"}:
        if path.endswith("/$count"):
            reason = "Deferred by firstmate terms-of-use scope to a later EXT-01 terms-of-use counts subfamily: scalar counts need a separate query and response contract from the approved list/show reads."
        elif "/file" in path or "/files" in path or "/localizations" in path or "/versions" in path:
            reason = "Deferred by firstmate terms-of-use scope to a later EXT-01 terms-of-use file-content subfamily: agreement file contents and localizations need a separate binary-output and projection contract, and file bytes are never downloaded or printed."
        elif version == "beta":
            reason = "Deferred by firstmate terms-of-use scope to a later EXT-01 terms-of-use beta subfamily: the six approved reads cover v1.0 only; beta agreements and acceptances need separate review."
    if mutates and owner is not None:
        owner = "WRITE-N"
        if disposition == "scheduled":
            reason = "No implemented command or reviewed raw contract yet."
        if method == "PATCH" and path == "/users/{user-id}":
            owner = "WRITE-02"
        elif method == "POST" and path == "/groups/{group-id}/members/$ref":
            owner = "WRITE-01"
        elif path.endswith(("/microsoft.graph.revokeSignInSessions", "/revokeSignInSessions")):
            owner = "WRITE-03"
        elif method == "PATCH" and path == "/identity/conditionalAccess/policies/{conditionalAccessPolicy-id}":
            owner = "WRITE-04"
        elif path.endswith(("/riskyUsers/microsoft.graph.dismiss", "/riskyUsers/dismiss")):
            owner = "WRITE-05"
    doc = operation.get("externalDocs", {}).get("url")
    source = doc or (f"https://learn.microsoft.com/en-us/graph/api/{read_source}?view=graph-rest-{'beta' if version == 'beta' else '1.0'}" if read_source else MAP_SOURCE)
    return {
        "id": f"{version}:{method}:{path}", "operationId": operation["operationId"],
        "version": version, "method": method, "path": path, "cloud": "commercial",
        "tenantType": "workforce", "documentation": doc,
        "authModes": {mode: {"status": "not-reviewed", "sources": [source]} for mode in ("delegated", "application")},
        "permissions": {"status": "not-reviewed", "sources": [source]},
        "roles": {"status": "not-reviewed", "sources": [source]},
        "licences": {"status": "not-reviewed", "sources": [source, LICENCE_SOURCE]},
        "cloudAvailability": {"status": "not-reviewed", "sources": [source]},
        "disposition": disposition, "reason": reason, "owningSlice": owner,
    }


def build(sources):
    rows, pins = [], []
    for version, source in sources.items():
        raw = source.read_bytes()
        pins.append({"version": version, "path": f"openapi/{version}/openapi.yaml", "sha256": hashlib.sha256(raw).hexdigest(), "bytes": len(raw)})
        discovered, excluded = 0, {}
        for path, method, operation in discover(source):
            discovered += 1
            row = make_row(version, path, method, operation)
            rows.append(row)
            if row["disposition"] == "excluded":
                root = path.split("/")[1].split("(")[0]
                excluded[root] = excluded.get(root, 0) + 1
        pins[-1]["discoveredOperations"] = discovered
        pins[-1]["excludedByRoot"] = dict(sorted(excluded.items()))
    rows.sort(key=lambda row: row["id"])
    return {"schemaVersion": 2, "repository": "microsoftgraph/msgraph-metadata", "revision": REVISION, "checkedOn": "2026-10-03", "sources": pins, "operations": rows}


def validate(inventory):
    schema = json.loads((ROOT / "inventory/schema.json").read_text())
    Draft202012Validator(schema).validate(inventory)
    ids = set()
    for row in inventory["operations"]:
        if row["id"] in ids or row["id"] != f'{row["version"]}:{row["method"]}:{row["path"]}':
            raise ValueError("Duplicate or inconsistent operation identity")
        ids.add(row["id"])
        if row["disposition"] in {"named-command", "reviewed-raw-read"}:
            raise ValueError("Discovery alone cannot claim implementation or reviewed raw access")
        if (scoped_slice(row["path"]) is None) != (row["disposition"] == "excluded"):
            raise ValueError("Boundary mismatch: out-of-scope operations must be excluded")
    if {source["version"] for source in inventory["sources"]} != {"v1.0", "beta"}:
        raise ValueError("Both metadata versions must be pinned independently")
    for source in inventory["sources"]:
        rows = [row for row in inventory["operations"] if row["version"] == source["version"]]
        if len(rows) != source["discoveredOperations"]:
            raise ValueError("Missing operation: every discovered operation must have a row")
        excluded = {}
        for row in rows:
            if row["disposition"] == "excluded":
                root = row["path"].split("/")[1].split("(")[0]
                excluded[root] = excluded.get(root, 0) + 1
        if excluded != source["excludedByRoot"]:
            raise ValueError("Excluded operation counts must reconcile with rows")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("generate", "check", "validate"))
    parser.add_argument("--source-dir", type=Path, default=ROOT / ".cache/inventory")
    args = parser.parse_args()
    output = ROOT / "inventory/operations.json"
    if args.action == "validate":
        inventory = json.loads(output.read_text())
        validate(inventory)
        print(f'Validated {len(inventory["operations"])} discovered operation rows; no completeness percentage claimed.')
        return
    inventory = build({version: args.source_dir / f"{version}.yaml" for version in ("v1.0", "beta")})
    validate(inventory)
    # One operation per line keeps generated diffs localized to changed routes.
    header = {key: value for key, value in inventory.items() if key != "operations"}
    encoded = json.dumps(header, indent=2)[:-2] + ',\n  "operations": [\n'
    encoded += ",\n".join("    " + json.dumps(row) for row in inventory["operations"])
    encoded += "\n  ]\n}\n"
    if args.action == "generate":
        output.parent.mkdir(exist_ok=True)
        output.write_text(encoded)
    elif output.read_text() != encoded:
        raise ValueError("Inventory drift: regenerate and review the metadata/boundary diff")
    print(f'Validated {len(inventory["operations"])} discovered operation rows; no completeness percentage claimed.')


if __name__ == "__main__":
    main()
