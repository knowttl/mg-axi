# Microsoft Graph capability and access contract

Research baseline: 2026-10-03.
Reverify each operation's linked documentation when implementing it.

## Representative SOC read coverage

All following operations are documented in v1.0.
D means delegated work/school; A means application.
These are supported read permission choices for representative operations, not a blanket permission bundle or claims about all related actions.
Detail, relationship, premium property, and write contracts require their own inventory rows.

| Phase and capability | Representative endpoint / documentation | D / A read permission |
|---|---|---|
| 1 Users | [/users](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0) | D User.ReadBasic.All for basics, User.Read.All for richer SOC data; A User.Read.All. |
| 1 Groups | [/groups](https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0) | D/A GroupMember.Read.All is a supported read choice; Group.Read.All for further properties. |
| 1 Membership | [/groups/{id}/members](https://learn.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0) | D/A GroupMember.ReadBasic.All minimum in current docs; richer access GroupMember.Read.All; hidden membership adds Member.Read.Hidden. |
| 1 CA policies | [/identity/conditionalAccess/policies](https://learn.microsoft.com/en-us/graph/api/conditionalaccessroot-list-policies?view=graph-rest-1.0) | D/A Policy.Read.All. |
| 1 Named locations | [/identity/conditionalAccess/namedLocations](https://learn.microsoft.com/en-us/graph/api/conditionalaccessroot-list-namedlocations?view=graph-rest-1.0) | D/A Policy.Read.All. |
| 1 Authentication methods | [/users/{id}/authentication/methods](https://learn.microsoft.com/en-us/graph/api/authentication-list-methods?view=graph-rest-1.0) | D/A for others UserAuthenticationMethod.Read.All; delegated self UserAuthenticationMethod.Read. |
| 1 Registration report | [/reports/authenticationMethods/userRegistrationDetails](https://learn.microsoft.com/en-us/graph/api/authenticationmethodsroot-list-userregistrationdetails?view=graph-rest-1.0) | D/A AuditLog.Read.All. |
| 1 Sign-ins | [/auditLogs/signIns](https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0) | D/A AuditLog.Read.All; CA details have additional policy permission/role requirements. |
| 1 Directory audits | [/auditLogs/directoryAudits](https://learn.microsoft.com/en-us/graph/api/directoryaudit-list?view=graph-rest-1.0) | D/A AuditLog.Read.All. |
| 1 Provisioning logs | [/auditLogs/provisioning](https://learn.microsoft.com/en-us/graph/api/provisioningobjectsummary-list?view=graph-rest-1.0) | D/A AuditLog.Read.All and Directory.Read.All; delegated role and P1/P2 licence apply. |
| 1 Risky users | [/identityProtection/riskyUsers](https://learn.microsoft.com/en-us/graph/api/riskyuser-list?view=graph-rest-1.0) | D/A IdentityRiskyUser.Read.All. |
| 1 Risk detections | [/identityProtection/riskDetections](https://learn.microsoft.com/en-us/graph/api/riskdetection-list?view=graph-rest-1.0) | D/A IdentityRiskEvent.Read.All. |
| 1 Risky service principals | [/identityProtection/riskyServicePrincipals](https://learn.microsoft.com/en-us/graph/api/identityprotectionroot-list-riskyserviceprincipals?view=graph-rest-1.0) | D/A IdentityRiskyServicePrincipal.Read.All. |
| 1 Service-principal risk detections | [/identityProtection/servicePrincipalRiskDetections](https://learn.microsoft.com/en-us/graph/api/identityprotectionroot-list-serviceprincipalriskdetections?view=graph-rest-1.0) | D/A IdentityRiskEvent.Read.All. |
| 2 Applications | [/applications](https://learn.microsoft.com/en-us/graph/api/application-list?view=graph-rest-1.0) | D/A Application.Read.All. |
| 2 Service principals | [/servicePrincipals](https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list?view=graph-rest-1.0) | D/A Application.Read.All. |
| 2 Directory roles | [/roleManagement/directory/roleAssignments](https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleassignments?view=graph-rest-1.0) | D/A RoleManagement.Read.Directory. |
| 2 PIM eligible | [/roleManagement/directory/roleEligibilityScheduleInstances](https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleeligibilityscheduleinstances?view=graph-rest-1.0) | D/A RoleEligibilitySchedule.Read.Directory. |
| 2 PIM active | [/roleManagement/directory/roleAssignmentScheduleInstances](https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleassignmentscheduleinstances?view=graph-rest-1.0) | D/A RoleAssignmentSchedule.Read.Directory. |
| 2 Devices | [/devices](https://learn.microsoft.com/en-us/graph/api/device-list?view=graph-rest-1.0) | D/A Device.Read.All. |
| 2 Administrative units | [/directory/administrativeUnits](https://learn.microsoft.com/en-us/graph/api/directory-list-administrativeunits?view=graph-rest-1.0) | D/A AdministrativeUnit.Read.All. |
| 2 Delegated consent | [/oauth2PermissionGrants](https://learn.microsoft.com/en-us/graph/api/oauth2permissiongrant-list?view=graph-rest-1.0) | D/A Directory.Read.All as read choice. |
| 2 Application consent | [/servicePrincipals/{id}/appRoleAssignments](https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignments?view=graph-rest-1.0) | D/A Application.Read.All. |

Supported delegated roles must also be satisfied, as specified by each linked operation.
For example, Security Reader is supported for CA reads, Reports Reader/Security Reader for audit reporting, and authentication-method inspection has dedicated supported administrator roles.
App-only access uses application consent rather than a signed-in user's delegated role, except where the particular API specifies additional requirements.

## Representative extended reads

All rows below are v1.0; D/A distinctions and field/role constraints still belong to the individual operation contract.

| Family | Endpoint and source | Supported permission choice | Licence or access caveat |
|---|---|---|---|
| Organization | [/organization](https://learn.microsoft.com/en-us/graph/api/organization-list?view=graph-rest-1.0) | D User.Read for restricted basics, Organization.Read.All for full metadata; A Organization.Read.All | Base inventory; contact fields are personal data. |
| Licences | [/subscribedSkus](https://learn.microsoft.com/en-us/graph/api/subscribedsku-list?view=graph-rest-1.0) | D/A LicenseAssignment.Read.All | SKU presence does not prove individual entitlement. |
| Domains | [/domains](https://learn.microsoft.com/en-us/graph/api/domain-list?view=graph-rest-1.0) | D/A Domain.Read.All | No P1/P2 prerequisite stated for this list. |
| Authentication policy | [/policies/authenticationMethodsPolicy](https://learn.microsoft.com/en-us/graph/api/authenticationmethodspolicy-get?view=graph-rest-1.0) | D/A Policy.Read.AuthenticationMethod | Features/methods have separate requirements. |
| Authorization policy | [/policies/authorizationPolicy](https://learn.microsoft.com/en-us/graph/api/authorizationpolicy-get?view=graph-rest-1.0) | D/A Policy.Read.All | Individual controlled features may be licensed. |
| Cross-tenant policy | [/policies/crossTenantAccessPolicy](https://learn.microsoft.com/en-us/graph/api/crosstenantaccesspolicy-get?view=graph-rest-1.0) | D/A Policy.Read.All | Default B2B and premium targeted/sync scenarios differ. |
| Provisioning | [/servicePrincipals/{id}/synchronization/jobs](https://learn.microsoft.com/en-us/graph/api/synchronization-synchronization-list-jobs?view=graph-rest-1.0) | D/A Synchronization.Read.All is a supported read choice | Provisioning and cross-tenant sync feature licensing applies. |
| Access reviews | [/identityGovernance/accessReviews/definitions](https://learn.microsoft.com/en-us/graph/api/accessreviewset-list-definitions?view=graph-rest-1.0) | D/A AccessReview.Read.All | P2 or ID Governance depending on capability. |
| Access packages | [/identityGovernance/entitlementManagement/accessPackages](https://learn.microsoft.com/en-us/graph/api/entitlementmanagement-list-accesspackages?view=graph-rest-1.0) | D/A EntitlementManagement.Read.All | Legacy P2 capabilities versus Governance features. |
| PIM groups | [/identityGovernance/privilegedAccess/group/eligibilityScheduleInstances](https://learn.microsoft.com/en-us/graph/api/privilegedaccessgroup-list-eligibilityscheduleinstances?view=graph-rest-1.0) | D/A PrivilegedEligibilitySchedule.Read.AzureADGroup | P2 or ID Governance; required group/principal filter. |
| Security attributes | [/directory/customSecurityAttributeDefinitions](https://learn.microsoft.com/en-us/graph/api/directory-list-customsecurityattributedefinitions?view=graph-rest-1.0) | D/A CustomSecAttributeDefinition.Read.All | Dedicated delegated attribute role; assignments use separate permissions; no P1/P2 prerequisite stated for this list. |
| Partner GDAP | [/tenantRelationships/delegatedAdminRelationships](https://learn.microsoft.com/en-us/graph/api/tenantrelationship-list-delegatedadminrelationships?view=graph-rest-1.0) | D/A DelegatedAdminRelationship.Read.All | Partner eligibility; no P1/P2 prerequisite stated for this list. |

## Licence matrix by area

| Area | Deployment/licence contract | Source |
|---|---|---|
| Users, groups, apps, service principals, devices, consent | Base inventory docs do not state P1/P2; dynamic groups and other premium features differ. | Linked operation docs above and [licensing](https://learn.microsoft.com/en-us/entra/fundamentals/licensing). |
| CA/named locations | P1 for CA, P2 for risk-based CA. | [Entra licensing](https://learn.microsoft.com/en-us/entra/fundamentals/licensing). |
| Authentication methods/reporting | Targeted method availability differs by method; registration reporting premium requirements must be recorded per feature. | [Method API](https://learn.microsoft.com/en-us/graph/api/authentication-list-methods?view=graph-rest-1.0), [registration API](https://learn.microsoft.com/en-us/graph/api/authenticationmethodsroot-list-userregistrationdetails?view=graph-rest-1.0). |
| Sign-in/audit Graph reporting | Conservative P1/P2 deployment prerequisite until the operation's documented access distinction is resolved; base logs and Graph access guidance differ. | [Activity log access](https://learn.microsoft.com/en-us/entra/identity/monitoring-health/howto-access-activity-logs). |
| Risky users / full protection | The riskyUsers API requires P2; full Identity Protection investigation requires P2/Suite. | [Risky-user API requirement](https://learn.microsoft.com/en-us/graph/api/riskyuser-get?view=graph-rest-1.0), [Identity Protection](https://learn.microsoft.com/en-us/entra/id-protection/overview-identity-protection). |
| Risk detections | API permits P1 or P2; full details need P2. | [Risk detections](https://learn.microsoft.com/en-us/graph/api/riskdetection-list?view=graph-rest-1.0). |
| Workload risk | The riskyServicePrincipal and servicePrincipalRiskDetection APIs require Microsoft Entra Workload Identities Premium; service-principal detections report riskDetail and riskLevel hidden without it. | [Risky-service-principal list](https://learn.microsoft.com/en-us/graph/api/identityprotectionroot-list-riskyserviceprincipals?view=graph-rest-1.0), [service-principal detection list](https://learn.microsoft.com/en-us/graph/api/identityprotectionroot-list-serviceprincipalriskdetections?view=graph-rest-1.0). |
| Directory roles / PIM | Built-in roles base, custom role assignments P1, PIM P2 or ID Governance. | [Licensing](https://learn.microsoft.com/en-us/entra/fundamentals/licensing), [Governance licensing](https://learn.microsoft.com/en-us/entra/id-governance/licensing-fundamentals). |
| Administrative units | P1 scoped administrators, Free members; dynamic membership P1. | [Administrative units](https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/administrative-units). |
| Reviews, entitlements, lifecycle | P2 legacy capabilities versus ID Governance advanced features; not one uniform licence. | [Governance licensing](https://learn.microsoft.com/en-us/entra/id-governance/licensing-fundamentals). |
| Workload/agent identity and network access | Do not label these uniformly P2; apply current Workload ID, agent and network-product requirements during inventory. | [Current Entra Graph map](https://learn.microsoft.com/en-us/graph/api/resources/identity-network-access-overview?view=graph-rest-1.0). |

## Licensing and completeness findings

Core users/groups/app/device inventory operation docs generally state no P1/P2 prerequisite; this is not a claim that every feature in these families is free.
Conditional Access uses P1; risk-based CA uses P2.
Built-in roles are base; custom role assignments use P1.
PIM requires P2 or ID Governance, not only P2.
Administrative units use P1 for scoped administrators and Free for members; dynamic membership requires additional P1 licensing.
These distinctions are supported by [Entra licensing](https://learn.microsoft.com/en-us/entra/fundamentals/licensing), [administrative-unit requirements](https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/administrative-units), and [Governance licensing](https://learn.microsoft.com/en-us/entra/id-governance/licensing-fundamentals).

The licence matrix above distinguishes risky-user API access from risk-detection detail access.
Workload-identity reads follow the same pattern one level up: the risky-service-principal and service-principal-detection APIs require Workload Identities Premium, and hidden riskDetail/riskLevel values report that licence boundary rather than no risk.
See [README.md](../README.md) for the implemented risk reads, limited/hidden-value handling and sign-in correlation workflow.

The licensing overview makes base sign-in/audit logs available on Free, while [Graph access guidance](https://learn.microsoft.com/en-us/entra/identity/monitoring-health/howto-access-activity-logs) describes P1/P2 tenants.
This contract exposes this distinction and adopts a conservative P1/P2 Graph deployment prerequisite rather than flattening the documentation into an unconditional licence claim.
Retention, premium fields, and licences for the generating feature remain separate constraints.
Never diagnose licence solely from HTTP 403.

Microsoft [authentication-method list guidance](https://learn.microsoft.com/en-us/graph/api/authentication-list-methods?view=graph-rest-1.0) discourages tenant-wide enumeration through individual users' methods.
Use the registration report for aggregate coverage and per-user calls for targeted inspection.
The [registration report](https://learn.microsoft.com/en-us/graph/api/authenticationmethodsroot-list-userregistrationdetails?view=graph-rest-1.0) does not work for disabled users, so absence is not proof of no MFA.

The v1.0 group-members route documents omitted service principals and suggests beta or expansion as workarounds.
See [README.md](../README.md) for the current raw and named member-read warning and fallback behavior.
Current group-list documentation places a write permission in its least-privileged column while listing read alternatives.
The plan therefore selects a supported read permission for required fields rather than mechanically copying the first cell.

## Full-family inventory and truthful coverage

INV-01's pinned machine-readable discovery inventory, boundary, evidence states and offline checks are documented in [the inventory contract](inventory.md).
It records discovered operations and later-slice ownership without claiming implemented or reviewed raw coverage.

The [current Entra Graph map](https://learn.microsoft.com/en-us/graph/api/resources/identity-network-access-overview?view=graph-rest-1.0) is broader than the initial SOC list.
The inventory must cover directory/tenant administration; users/groups/relationships/extensions/deleted objects; apps/service principals/consent/credential metadata; authentication and sign-in policy; directory/group PIM; access reviews, entitlements and lifecycle workflows; external identities and user flows; cross-tenant policies and provisioning; identity/workload risk and reporting; agent identity/governance; network access; partner GDAP and contracts.
The map identifies beta-only areas including network access and selected backup, user-flow, policy and tenant-setting surfaces.
Deprecated multicloud permissions management is recorded as deprecated rather than implemented.
Credential values, device recovery keys, LAPS passwords and secret minting are explicit blocked dispositions for the initial supported surface, separate from useful metadata.
Mail/files and other non-Entra content are disabled unless explicitly enabled for raw read access; named future packs remain separately authorized work.

The metadata inventory is not a substitute for permissions/licensing documentation.
Report named coverage and raw reachability separately.
No coverage percentage is claimed before a complete denominator exists.
The final full-Entra milestone cannot close while agreed operations remain merely scheduled.
