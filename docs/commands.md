# Command reference

Contents: [Executable leaves](#executable-leaves) • [Help and permissions](#help-and-permissions)

Generated from the executable catalogue; do not edit by hand.

## Executable leaves

| Command | Capability | Effect |
|---|---|---|
| `mg-axi home` | native | local |
| `mg-axi profile create` | native | local |
| `mg-axi profile list` | native | local |
| `mg-axi profile show` | native | local |
| `mg-axi auth login` | native | local |
| `mg-axi setup` | native | local |
| `mg-axi setup hooks` | native | local |
| `mg-axi doctor` | native | read |
| `mg-axi entra user list` | native | read |
| `mg-axi entra user show` | native | read |
| `mg-axi entra user update` | native | write |
| `mg-axi entra user revoke-sessions` | native | write |
| `mg-axi entra user authentication-method list` | native | read |
| `mg-axi entra registration list` | native | read |
| `mg-axi entra group list` | native | read |
| `mg-axi entra group show` | native | read |
| `mg-axi entra group member list` | native | read |
| `mg-axi entra group member-of list` | native | read |
| `mg-axi entra group member add` | native | write |
| `mg-axi entra directory-role list` | native | read |
| `mg-axi entra directory-role show` | native | read |
| `mg-axi entra directory-role-template list` | native | read |
| `mg-axi entra directory-role-template show` | native | read |
| `mg-axi entra directory-role member list` | native | read |
| `mg-axi entra directory-role scoped-member list` | native | read |
| `mg-axi entra directory-role scoped-member show` | native | read |
| `mg-axi entra scoped-role-membership list` | native | read |
| `mg-axi entra scoped-role-membership show` | native | read |
| `mg-axi entra role-assignment list` | native | read |
| `mg-axi entra pim eligible list` | native | read |
| `mg-axi entra pim active list` | native | read |
| `mg-axi entra pim group-assignment-schedule list` | native | read |
| `mg-axi entra pim group-assignment-schedule show` | native | read |
| `mg-axi entra pim group-assignment-instance list` | native | read |
| `mg-axi entra pim group-assignment-instance show` | native | read |
| `mg-axi entra pim group-eligibility-schedule list` | native | read |
| `mg-axi entra pim group-eligibility-schedule show` | native | read |
| `mg-axi entra pim group-eligibility-instance list` | native | read |
| `mg-axi entra pim group-eligibility-instance show` | native | read |
| `mg-axi entra pim group-eligibility-request list` | native | read |
| `mg-axi entra pim group-eligibility-request show` | native | read |
| `mg-axi entra device list` | native | read |
| `mg-axi entra device show` | native | read |
| `mg-axi entra device owners` | native | read |
| `mg-axi entra device users` | native | read |
| `mg-axi entra administrative-unit list` | native | read |
| `mg-axi entra administrative-unit show` | native | read |
| `mg-axi entra administrative-unit member list` | native | read |
| `mg-axi entra organization list` | native | read |
| `mg-axi entra organization show` | native | read |
| `mg-axi entra organization branding show` | native | read |
| `mg-axi entra organization branding-localization list` | native | read |
| `mg-axi entra organization branding-localization show` | native | read |
| `mg-axi entra domain list` | native | read |
| `mg-axi entra domain show` | native | read |
| `mg-axi entra domain verification-dns-record list` | native | read |
| `mg-axi entra domain verification-dns-record show` | native | read |
| `mg-axi entra domain service-configuration-record list` | native | read |
| `mg-axi entra domain service-configuration-record show` | native | read |
| `mg-axi entra domain-dns-record list` | native | read |
| `mg-axi entra domain-dns-record show` | native | read |
| `mg-axi entra certificate-auth-pki list` | native | read |
| `mg-axi entra certificate-auth-pki show` | native | read |
| `mg-axi entra certificate-auth-pki count` | native | read |
| `mg-axi entra certificate-authority list` | native | read |
| `mg-axi entra certificate-authority show` | native | read |
| `mg-axi entra certificate-authority count` | native | read |
| `mg-axi entra subscription list` | native | read |
| `mg-axi entra subscription show` | native | read |
| `mg-axi entra subscription count` | native | read |
| `mg-axi entra on-premises-synchronization list` | native | read |
| `mg-axi entra on-premises-synchronization show` | native | read |
| `mg-axi entra agreement list` | native | read |
| `mg-axi entra agreement show` | native | read |
| `mg-axi entra agreement acceptance list` | native | read |
| `mg-axi entra agreement acceptance show` | native | read |
| `mg-axi entra agreement-acceptance list` | native | read |
| `mg-axi entra agreement-acceptance show` | native | read |
| `mg-axi entra directory-object list` | native | read |
| `mg-axi entra directory-object show` | native | read |
| `mg-axi entra directory-object count` | native | read |
| `mg-axi entra contact list` | native | read |
| `mg-axi entra contact show` | native | read |
| `mg-axi entra contact count` | native | read |
| `mg-axi entra contact show-manager` | native | read |
| `mg-axi entra contact list-direct-reports` | native | read |
| `mg-axi entra contact show-direct-report` | native | read |
| `mg-axi entra contact count-direct-reports` | native | read |
| `mg-axi entra contact list-member-of` | native | read |
| `mg-axi entra contact show-member-of` | native | read |
| `mg-axi entra contact count-member-of` | native | read |
| `mg-axi entra group-lifecycle-policy list` | native | read |
| `mg-axi entra group-lifecycle-policy show` | native | read |
| `mg-axi entra group-lifecycle-policy count` | native | read |
| `mg-axi entra group-setting-template list` | native | read |
| `mg-axi entra group-setting-template show` | native | read |
| `mg-axi entra group-setting-template count` | native | read |
| `mg-axi entra attribute-set list` | native | read |
| `mg-axi entra attribute-set show` | native | read |
| `mg-axi entra attribute-set count` | native | read |
| `mg-axi entra custom-security-attribute-definition list` | native | read |
| `mg-axi entra custom-security-attribute-definition show` | native | read |
| `mg-axi entra custom-security-attribute-definition count` | native | read |
| `mg-axi entra allowed-value list` | native | read |
| `mg-axi entra allowed-value show` | native | read |
| `mg-axi entra allowed-value count` | native | read |
| `mg-axi entra contract list` | native | read |
| `mg-axi entra contract show` | native | read |
| `mg-axi entra contract count` | native | read |
| `mg-axi entra delegated-admin-customer list` | native | read |
| `mg-axi entra delegated-admin-customer show` | native | read |
| `mg-axi entra delegated-admin-relationship list` | native | read |
| `mg-axi entra delegated-admin-relationship show` | native | read |
| `mg-axi entra delegated-admin-relationship list-access-assignments` | native | read |
| `mg-axi entra delegated-admin-relationship show-access-assignment` | native | read |
| `mg-axi entra delegated-admin-relationship list-operations` | native | read |
| `mg-axi entra delegated-admin-relationship show-operation` | native | read |
| `mg-axi entra delegated-admin-relationship list-requests` | native | read |
| `mg-axi entra delegated-admin-relationship show-request` | native | read |
| `mg-axi entra delegated-admin-customer list-service-management-details` | native | read |
| `mg-axi entra delegated-admin-customer show-service-management-detail` | native | read |
| `mg-axi entra delegated-admin-customer count` | native | read |
| `mg-axi entra delegated-admin-customer count-service-management-details` | native | read |
| `mg-axi entra delegated-admin-relationship count` | native | read |
| `mg-axi entra delegated-admin-relationship count-access-assignments` | native | read |
| `mg-axi entra delegated-admin-relationship count-operations` | native | read |
| `mg-axi entra delegated-admin-relationship count-requests` | native | read |
| `mg-axi entra multi-tenant-organization show` | native | read |
| `mg-axi entra multi-tenant-organization join-request show` | native | read |
| `mg-axi entra multi-tenant-organization tenant list` | native | read |
| `mg-axi entra multi-tenant-organization tenant count` | native | read |
| `mg-axi entra tenant-information show` | native | read |
| `mg-axi entra sign-in list` | native | read |
| `mg-axi entra sign-in show` | native | read |
| `mg-axi entra directory-audit list` | native | read |
| `mg-axi entra directory-audit show` | native | read |
| `mg-axi entra provisioning list` | native | read |
| `mg-axi entra provisioning show` | native | read |
| `mg-axi entra application list` | native | read |
| `mg-axi entra application show` | native | read |
| `mg-axi entra service-principal list` | native | read |
| `mg-axi entra service-principal show` | native | read |
| `mg-axi entra application owner list` | native | read |
| `mg-axi entra service-principal owner list` | native | read |
| `mg-axi entra service-principal oauth2-grant list` | native | read |
| `mg-axi entra service-principal app-role-assignment list` | native | read |
| `mg-axi entra service-principal app-role-assignment show` | native | read |
| `mg-axi entra service-principal app-role-assigned-to list` | native | read |
| `mg-axi entra service-principal app-role-assigned-to show` | native | read |
| `mg-axi entra application federated-credential list` | native | read |
| `mg-axi entra application federated-credential show` | native | read |
| `mg-axi entra service-principal federated-credential list` | native | read |
| `mg-axi entra service-principal federated-credential show` | native | read |
| `mg-axi entra risky-user list` | native | read |
| `mg-axi entra risky-user show` | native | read |
| `mg-axi entra risky-user dismiss` | native | write |
| `mg-axi entra risk-detection list` | native | read |
| `mg-axi entra risk-detection show` | native | read |
| `mg-axi entra risky-service-principal list` | native | read |
| `mg-axi entra risky-service-principal show` | native | read |
| `mg-axi entra risky-service-principal history list` | native | read |
| `mg-axi entra risky-service-principal history show` | native | read |
| `mg-axi entra service-principal-risk-detection list` | native | read |
| `mg-axi entra service-principal-risk-detection show` | native | read |
| `mg-axi entra conditional-access policy list` | native | read |
| `mg-axi entra conditional-access policy show` | native | read |
| `mg-axi entra conditional-access policy update` | native | write |
| `mg-axi entra conditional-access named-location list` | native | read |
| `mg-axi entra conditional-access named-location show` | native | read |
| `mg-axi entra conditional-access auth-strength-policy list` | native | read |
| `mg-axi entra conditional-access auth-strength-policy show` | native | read |
| `mg-axi entra conditional-access combination-configuration list` | native | read |
| `mg-axi entra conditional-access combination-configuration show` | native | read |
| `mg-axi entra conditional-access auth-method-mode list` | native | read |
| `mg-axi entra conditional-access auth-method-mode show` | native | read |
| `mg-axi entra conditional-access template list` | native | read |
| `mg-axi entra conditional-access template show` | native | read |
| `mg-axi entra conditional-access auth-context list` | native | read |
| `mg-axi entra conditional-access auth-context show` | native | read |
| `mg-axi entra conditional-access deleted-policy list` | native | read |
| `mg-axi entra conditional-access deleted-policy show` | native | read |
| `mg-axi entra conditional-access deleted-named-location list` | native | read |
| `mg-axi entra conditional-access deleted-named-location show` | native | read |
| `mg-axi entra identity-provider list` | native | read |
| `mg-axi entra identity-provider show` | native | read |
| `mg-axi entra identity-provider count` | native | read |
| `mg-axi entra identity-provider available-types` | native | read |
| `mg-axi entra federation-configuration list` | native | read |
| `mg-axi entra federation-configuration show` | native | read |
| `mg-axi entra federation-configuration count` | native | read |
| `mg-axi entra federation-configuration available-types` | native | read |
| `mg-axi entra deleted-user list` | native | read |
| `mg-axi entra deleted-user count` | native | read |
| `mg-axi entra deleted-group list` | native | read |
| `mg-axi entra deleted-group count` | native | read |
| `mg-axi entra deleted-application list` | native | read |
| `mg-axi entra deleted-application count` | native | read |
| `mg-axi entra deleted-service-principal list` | native | read |
| `mg-axi entra deleted-service-principal count` | native | read |
| `mg-axi entra deleted-administrative-unit list` | native | read |
| `mg-axi entra deleted-administrative-unit count` | native | read |
| `mg-axi entra deleted-item show` | native | read |
| `mg-axi entra data-policy-operation list` | native | read |
| `mg-axi entra data-policy-operation show` | native | read |
| `mg-axi entra data-policy-operation count` | native | read |
| `mg-axi entra fraud-protection-provider list` | native | read |
| `mg-axi entra fraud-protection-provider show` | native | read |
| `mg-axi entra fraud-protection-provider count` | native | read |
| `mg-axi entra web-application-firewall-provider list` | native | read |
| `mg-axi entra web-application-firewall-provider show` | native | read |
| `mg-axi entra web-application-firewall-provider count` | native | read |
| `mg-axi entra web-application-firewall-verification list` | native | read |
| `mg-axi entra web-application-firewall-verification show` | native | read |
| `mg-axi entra web-application-firewall-verification count` | native | read |
| `mg-axi entra web-application-firewall-verification provider show` | native | read |
| `mg-axi entra access-review definition list` | native | read |
| `mg-axi entra access-review definition show` | native | read |
| `mg-axi entra access-review instance list` | native | read |
| `mg-axi entra access-review instance show` | native | read |
| `mg-axi entra access-review decision list` | native | read |
| `mg-axi entra access-review decision show` | native | read |
| `mg-axi entra access-review contacted-reviewer list` | native | read |
| `mg-axi entra access-review contacted-reviewer show` | native | read |
| `mg-axi entra access-review stage list` | native | read |
| `mg-axi entra access-review stage show` | native | read |
| `mg-axi entra entitlement catalog list` | native | read |
| `mg-axi entra entitlement catalog show` | native | read |
| `mg-axi entra entitlement catalog count` | native | read |
| `mg-axi entra entitlement access-package list` | native | read |
| `mg-axi entra entitlement access-package show` | native | read |
| `mg-axi entra entitlement access-package count` | native | read |
| `mg-axi entra entitlement assignment-policy list` | native | read |
| `mg-axi entra entitlement assignment-policy show` | native | read |
| `mg-axi entra entitlement assignment-policy count` | native | read |
| `mg-axi entra entitlement resource-role-scope list` | native | read |
| `mg-axi entra entitlement resource-role-scope show` | native | read |
| `mg-axi entra entitlement resource-role-scope count` | native | read |
| `mg-axi entra entitlement assignment list` | native | read |
| `mg-axi entra entitlement assignment show` | native | read |
| `mg-axi entra entitlement assignment count` | native | read |
| `mg-axi entra entitlement assignment-request list` | native | read |
| `mg-axi entra entitlement assignment-request show` | native | read |
| `mg-axi entra entitlement assignment-request count` | native | read |
| `mg-axi entra lifecycle workflow list` | native | read |
| `mg-axi entra lifecycle workflow show` | native | read |
| `mg-axi entra lifecycle workflow count` | native | read |
| `mg-axi entra lifecycle workflow-template list` | native | read |
| `mg-axi entra lifecycle workflow-template show` | native | read |
| `mg-axi entra lifecycle workflow-template count` | native | read |
| `mg-axi entra lifecycle task-definition list` | native | read |
| `mg-axi entra lifecycle task-definition show` | native | read |
| `mg-axi entra lifecycle task-definition count` | native | read |
| `mg-axi entra lifecycle settings show` | native | read |
| `mg-axi entra lifecycle run list` | native | read |
| `mg-axi entra lifecycle run show` | native | read |
| `mg-axi entra lifecycle run count` | native | read |
| `mg-axi entra lifecycle user-processing-result list` | native | read |
| `mg-axi entra lifecycle user-processing-result show` | native | read |
| `mg-axi entra lifecycle user-processing-result count` | native | read |
| `mg-axi entra lifecycle subject-processing-result list` | native | read |
| `mg-axi entra lifecycle subject-processing-result show` | native | read |
| `mg-axi entra lifecycle subject-processing-result count` | native | read |
| `mg-axi entra lifecycle task-report list` | native | read |
| `mg-axi entra lifecycle task-report show` | native | read |
| `mg-axi entra lifecycle task-report count` | native | read |
| `mg-axi entra lifecycle task-report task show` | native | read |
| `mg-axi entra lifecycle task-report task-definition show` | native | read |
| `mg-axi entra lifecycle run user-processing-result list` | native | read |
| `mg-axi entra lifecycle run user-processing-result show` | native | read |
| `mg-axi entra lifecycle run user-processing-result count` | native | read |
| `mg-axi entra lifecycle run subject-processing-result list` | native | read |
| `mg-axi entra lifecycle run subject-processing-result show` | native | read |
| `mg-axi entra lifecycle run subject-processing-result count` | native | read |
| `mg-axi entra lifecycle run task-processing-result list` | native | read |
| `mg-axi entra lifecycle run task-processing-result show` | native | read |
| `mg-axi entra lifecycle run task-processing-result count` | native | read |
| `mg-axi api get` | native | read |
| `mg-axi security hunting query` | native | read |
| `mg-axi mail folder create` | native | write |
| `mg-axi mail folder delete` | native | write |
| `mg-axi mail category create` | native | write |
| `mg-axi mail rule create` | native | write |
| `mg-axi mail rule update` | native | write |
| `mg-axi mail rule delete` | native | write |
| `mg-axi mail message tag` | native | write |
| `mg-axi mail message flag` | native | write |
| `mg-axi mail message mark` | native | write |
| `mg-axi mail message move` | native | write |
| `mg-axi mail message copy` | native | write |
| `mg-axi mail message trash` | native | write |
| `mg-axi mail attachment save` | native | read |
| `mg-axi mail message export` | native | read |
| `mg-axi mail attachment list` | native | read |
| `mg-axi mail settings show` | native | read |
| `mg-axi mail rule list` | native | read |
| `mg-axi mail rule show` | native | read |
| `mg-axi mail category list` | native | read |
| `mg-axi mail message show` | native | read |
| `mg-axi mail message list` | native | read |
| `mg-axi mail folder list` | native | read |
| `mg-axi mail folder show` | native | read |
| `mg-axi login` | native | local |
| `mg-axi config path` | native | local |
| `mg-axi auth status` | native | local |
| `mg-axi auth logout` | native | local |

## Help and permissions

Run `mg-axi <leaf-path> --help` for accepted flags and defaults without signing in.
See [write gates](writes.md), [read permissions](graph-coverage.md) and [mail authorization](mail.md).
