# mg-axi

Agent-facing Microsoft Graph CLI with one shared core and domain packs, read-only by default.
The Entra pack comes first, with phased full coverage and gated named writes.

CLI-01 provides a local TypeScript/AXI shell, strict command catalogue, leaf help and fast version probes.
AUTH-01 adds versioned dedicated-app delegated profiles and explicit login.
AUTH-02 adds certificate and workload-federated application profiles and a client-credentials service.
CORE-01 adds the shared policy-enforced Graph read session, exercised through an injected fixture HTTP transport.
CORE-02 adds session collections, query validation, bounded retries and cancellation under the [read execution contract](docs/execution.md#read-mechanics-and-source-contracts).
API-01 executes `mg-axi api get <path>`, serving the reviewed v1.0 raw surface through that session.
READ-01 executes Entra user list/show through that session in both delegated and application modes; current usage follows below.
READ-02 adds group list/show and direct or transitive member and parent-membership reads through the same session.
READ-04 adds targeted per-user authentication-method reads and the tenant registration report through the same session, with phone numbers redacted.
READ-09 adds directory-role list/show, current role-assignment inventory and active/eligible PIM reads through the same session.
READ-03 adds Conditional Access policy, named-location, authentication-strength-policy, combination-configuration, authentication-method-mode, template, auth-context, deleted-policy and deleted-named-location list/show as separate grammar through the same session; Conditional Access usage follows the device and administrative-unit usage below.
READ-05 executes Entra sign-in, directory-audit and provisioning-log list/show through that session; log usage follows the Conditional Access usage below.
READ-07 adds application and service-principal list/show with credential expiry metadata and owner reads through the same session.
READ-10 adds directory-device and administrative-unit list/show and unit-member reads through the same session.
READ-08 adds service-principal delegated-grant and app-role-assignment consent reads for a named client through the same session.
EXT-01 (domains) adds tenant-domain list/show, per-domain verification and service-configuration DNS record reads, and top-level domain DNS record reads through the same session.
EXT-01 (certificate auth) adds PKI configuration and certificate-authority list/show/count reads with default certificate-blob omission through the same session.
EXT-01 (directory subscriptions) adds commercial-subscription list/show/count reads with compact licence rows through the same session.
EXT-01 (on-premises synchronization) adds on-premises directory-synchronization list/show reads with Global Administrator role guidance through the same session.
EXT-01 (terms of use) adds terms-of-use agreement list/show, per-agreement acceptance list/show and tenant-wide agreement-acceptance list/show reads with Security Reader role and Entra ID P1 guidance through the same session.
EXT-01 (directory objects) adds directory-object list/show/count reads with @odata.type subtype rows and Directory.Read.All guidance through the same session.
EXT-01 (deleted items) adds soft-deleted user, group, application, service-principal and administrative-unit list/count reads plus deleted-item show with per-type least-privilege scope guidance through the same session.
EXT-01 (contacts) adds organizational-contact list/show/count reads with minimal personal-data rows and OrgContact.Read.All guidance through the same session.
EXT-01m (contact navigation) adds per-contact manager and direct-report reads (show-manager, list/show/count-direct-reports with user/contact casts) as directory objects with type plus minimal rows through the same session.
EXT-01n (contact membership) adds per-contact memberOf and transitiveMemberOf reads (list/show/count-member-of with transitive and group/administrativeUnit casts) as directory objects with type plus minimal rows through the same session.
EXT-03 (identity providers) adds workforce identity-provider list/show/count/available-types reads with secret scrubbing through the same session.
EXT-01 (federation configurations) adds workforce directory federation-configuration list/show/count/available-types reads with default signing-certificate omission through the same session.
EXT-03 (data policy operations) adds workforce data-policy-operation list/show/count reads with storage-location redaction through the same session.
EXT-03 (risk prevention) adds workforce fraud-protection-provider, web-application-firewall-provider and web-application-firewall-verification list/show/count reads plus verification provider show with key-material scrubbing through the same delegated session.
READ-06 executes Entra risky-user, risk-detection, risky-service-principal (with per-principal history) and service-principal-risk-detection list/show through that session; risk usage follows the log usage below.
EXT-01 (organization) adds tenant-organization list/show, default sign-in branding metadata and locale branding reads through the same session.
EXT-04 (contracts) adds partner-tenant customer-contract list/show/count through the same session.
EXT-04 (delegated-admin) adds partner-tenant delegated-admin customer and relationship list/show through the same session.
EXT-04 (multi-tenant-organization) adds multitenant-organization show, join-request show and member-tenant list/count through the same session.
EXT-04 (tenant-information) adds tenant-information show by domain name or tenant ID through the allowlisted session function-argument contract.
EXT-01 (group lifecycle) adds group lifecycle-policy and group setting-template list/show/count reads through the same session.
EXT-01 (custom security attributes) adds attribute-set, custom-security-attribute-definition and allowed-value list/show/count reads with attribute-role denial guidance through the same session.
Tests use fixture credential and transport providers; no tenant, real credentials or network access are required for help or an unconfigured home view.

Run commands non-interactively as `npx -y @knowttl/mg-axi ...`: no global install needed and no interactive prompts.
Version pinning is the installer's choice: use `npx -y @knowttl/mg-axi@<version> ...` to select a specific release.

```sh
npx -y @knowttl/mg-axi --version
npx -y @knowttl/mg-axi --help
npx -y @knowttl/mg-axi entra user list --help
```

For a local build, use the Node requirement and pinned pnpm version declared in [package.json](package.json):

```sh
corepack pnpm install --frozen-lockfile --ignore-scripts --config.confirm-modules-purge=false
corepack pnpm build
node dist/bin/mg-axi.js
node dist/bin/mg-axi.js --version
node dist/bin/mg-axi.js entra user list --help
node dist/bin/mg-axi.js entra user show --help
```

Run checkout commands from the repository root with `node dist/bin/mg-axi.js`.
To install the built checkout binary on PATH, run this explicitly from the repository root:

```sh
npm link --ignore-scripts
mg-axi --version
mg-axi --help
```

Without linking, substitute `node dist/bin/mg-axi.js` from the repository root wherever examples or CLI output use `mg-axi`.
The no-args home view and `--help` hints print the `npx -y @knowttl/mg-axi ...` form, which works with or without a linked checkout.

The home view reports unavailable tenant summaries explicitly.
`entra user list` defaults to basic properties (`id`, `displayName`, `userPrincipalName`, `mail`); `entra user show --id <user-id-or-upn>` defaults to the richer server property set.
`--select` requests properties from the [supported user property set](src/entra-users.ts); `--fields` projects locally and must be a subset of the fetched selection.
Text values longer than 500 characters are truncated, including strings in `businessPhones`; `--full` removes text truncation without lifting redaction or row caps.
Explicit null values stay null, missing properties stay absent, and denied reads return structured errors rather than empty results.
User lists accept `--filter` for an OData filter and default to a 100-row cap; `--limit` changes the cap, while the incompatible `--all` follows pages within request, byte and deadline budgets.
Partial lists report `complete: false` with a uniform `count` line (`1 users shown, more available`), a reason and an opaque `cursor` preserving unreturned rows.
Resume with `--cursor <cursor-from-output>` using the same profile, authentication scopes and API version; original `--select` and `--filter` values may be repeated or omitted, and conflicting values fail validation.
Repeat `--fields` and `--full` when the same local view is wanted; these are not saved in the cursor.
Delegated user reads default to `https://graph.microsoft.com/User.Read.All` with `--scopes` available for least-privilege basics; application profiles use the configured Graph `.default` audience and reject delegated scopes.
All delegated reads, including raw reads and cursor resumes, reject scopes outside `READ_SCOPES` in the [shared session](src/graph-session.ts) before credential acquisition; [Graph coverage](docs/graph-coverage.md) explains the operation-specific read permission choices.
Write scopes are refused with `VALIDATION_ERROR` and a list of supported read scopes.
User reads acquire credentials silently; a resume containing only buffered rows can finish without another Graph request.
WRITE-02 adds a named write: `mg-axi entra user update --user <user-id-or-upn> --account-enabled true|false` sets one user's `accountEnabled` through `PATCH /users/{id}` with only that property sent.
`--account-enabled` accepts exactly `true` or `false`; uppercase and whitespace-padded values are rejected.
The command supports only `--api-version v1.0`; beta writes are rejected before credentials or HTTP.
Without `--execute` the command previews the desired-state diff read through the user show route and journals nothing; an already-desired value is a no-op with exit 0.
Preview also requires the write-enabled profile and operation allowlist described below.
User IDs and UPNs are accepted, including guest UPNs containing `#EXT#`; pass the literal identifier, quoted for the shell, rather than percent-encoding it.
The first lookup pins the Graph object ID for the PATCH and subsequent reads; successful execution and no-op results return that object ID in `user.id`.
Enabling and disabling are disruptive: every `--execute` run needs `--confirm '<user-id-or-upn>'` repeating the target exactly, including already-desired states.
After a successful PATCH the command rereads the user and reports a `WRITE_CONFLICT` when the value is not what was sent; user-update answers 204 with an empty body, so the reread is the only proof.
A failed verification read reports `OUTCOME_UNKNOWN`; read back the pinned object ID before proceeding and never replay the intent.
Fresh reads do not make this write atomic: no ETag condition is sent, so another actor can change the account between the read, PATCH and verification.
The least-privileged permission pair is `User.EnableDisableAccount.All` plus `User.Read.All` in both modes.
Delegated PATCH credentials request both scopes together; reads request `User.Read.All`, and credentials are acquired silently.
Use explicit `mg-axi login --profile soc --scopes https://graph.microsoft.com/User.EnableDisableAccount.All,https://graph.microsoft.com/User.Read.All` to sign in for the write.
Application profiles use the configured Graph `.default` audience with the pair admin-consented on the app registration; the command does not request a per-operation scope subset.
The update command rejects caller-supplied `--scopes`.
Delegated callers need `Privileged Authentication Administrator` for admin targets and must generally outrank the target; app-only callers need the pair plus a higher-privileged admin role assignment, and 403 denials surface both rules because a 403 never says which prerequisite is missing.
WRITE-03 adds a named action write: `mg-axi entra user revoke-sessions --user <user-id>` revokes one user's sign-in sessions through `POST /users/{id}/revokeSignInSessions` with no request body.
The command treats an accepted 2xx status as success without validating the response body or performing a verification reread; it returns `user.id`, `user.sessionsRevoked: true` and `auditId`.
The command supports only `--api-version v1.0`; beta writes are rejected before credentials or HTTP.
Without `--execute` the command previews the action and journals nothing: the preview states that Graph resets `signInSessionsValidFromDateTime`, invalidating issued refresh tokens and browser session cookies so the user must sign in again, and that the action cannot be undone - there is no rollback.
The preview also carries the two Microsoft-stated limits: token revocation can lag a few minutes after the call returns, and external users are unaffected because they sign in through their home tenant.
`--user` takes the user object ID only; UPNs are not resolved.
The target is verified as a user through the user show route before preview and again before sending; a failed or malformed user read blocks the operation.
Revocation is always disruptive: every `--execute` run needs `--confirm '<user-id>'` repeating the target exactly.
The revocation action's least-privileged permission is `User.RevokeSessions.All` in both modes.
Delegated revocation credentials request that scope; reads request `User.ReadBasic.All`, and credentials are acquired silently.
Use explicit `mg-axi login --profile soc --scopes https://graph.microsoft.com/User.RevokeSessions.All,https://graph.microsoft.com/User.ReadBasic.All` to sign in for the write.
Application profiles use the configured Graph `.default` audience with `User.RevokeSessions.All` and the prerequisite [user-read permission](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0), `User.Read.All`, admin-consented on the app registration; the command does not request a per-operation scope subset.
The revoke command rejects caller-supplied `--scopes`.
A 403 denial surfaces the `User.RevokeSessions.All` requirement without inventing a role verdict; a timeout or 5xx after sending reports `OUTCOME_UNKNOWN`, and neither outcome is ever replayed.
Unknown-outcome guidance includes a user read with `--select id`; that read checks target accessibility, not whether revocation took effect.
Unknown flags, unexpected arguments, missing required values and unsupported combinations exit 2 before credential acquisition or HTTP.
Resource identifiers in named commands and raw paths cannot begin with `$` or contain parentheses, except the three allowlisted function-argument routes (the two tenant-information lookups and the subscription commerce-key lookup), whose parenthesised values arrive only as validated CLI flags and are OData-quoted and encoded by the shared session; OData route segments such as `$count`, `$value` and `$ref`, and other function-style segments such as `delta()`, cannot be used as IDs and fail validation before credentials.
Help and successful views, including partial lists, exit 0; authentication, policy and Graph failures exit 1.
Data and structured errors use TOON on stdout; diagnostics belong on stderr.
Bare `-v`, `-V` and `--version` print only the package version without importing the catalogue.

Create a delegated profile with your organization-owned public client registration, explicit workforce tenant UUID and explicit commercial cloud:

```sh
mg-axi profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial
mg-axi profile list
mg-axi profile show --profile soc
mg-axi login --profile soc --scopes https://graph.microsoft.com/User.Read.All,https://graph.microsoft.com/GroupMember.Read.All,https://graph.microsoft.com/Group.Read.All,https://graph.microsoft.com/Application.Read.All
```

After login, use:

```sh
mg-axi entra user list --profile soc --limit 10
mg-axi entra user list --profile soc --select id,displayName,department --fields id,department
mg-axi entra user show --profile soc --id <user-id-or-upn> --full
mg-axi entra group list --profile soc --limit 10
mg-axi entra group show --profile soc --id <group-id> --scopes https://graph.microsoft.com/Group.Read.All
mg-axi entra group member list --profile soc --group <group-id>
mg-axi entra group member list --profile soc --group <group-id> --transitive
mg-axi entra group member-of list --profile soc --group <group-id> --transitive
mg-axi entra application list --profile soc --limit 10
mg-axi entra application show --profile soc --id <application-object-id>
mg-axi entra application owner list --profile soc --application <application-object-id>
mg-axi entra service-principal list --profile soc --limit 10
mg-axi entra service-principal show --profile soc --id <service-principal-object-id>
mg-axi entra service-principal owner list --profile soc --service-principal <service-principal-object-id>
```

Read granted consent for one client service principal; delegated profiles first need explicit login with the read scopes:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Directory.Read.All,https://graph.microsoft.com/Application.Read.All
mg-axi entra service-principal oauth2-grant list --profile soc --service-principal <service-principal-object-id>
mg-axi entra service-principal app-role-assignment list --profile soc --service-principal <service-principal-object-id> --filter "resourceId eq '<resource-id>'"
```

`entra service-principal oauth2-grant list` defaults to `id`, `consentType`, `principalId`, `resourceId` and `scope`: the delegated scopes granted to the client, with `consentType` AllPrincipals covering every user (and an explicit null `principalId`) versus Principal covering the one user named by `principalId`.
`entra service-principal app-role-assignment list` defaults to `id`, `appRoleId`, `resourceDisplayName` and `resourceId`: the app-only roles granted to the client on each resource API.
Both lists show actual granted consent records; the application's requested permissions (`requiredResourceAccess`) are declared on the application object and are never shown here, and grant creation, revocation and consent belong to later write slices, never to these reads.
`--select` requests properties from the [reviewed grant property sets](src/entra-grants.ts); `--fields` must be a subset of the fetched selection.
`--filter` passes through as plain `$filter` with no `$count` or `ConsistencyLevel` contract; the named-list caps, uniform `count` totals, cursor resume rules and 500-character text truncation described above also apply.
Resume either list with the same `--service-principal` object ID.
Delegated oauth2-grant reads default to `https://graph.microsoft.com/Directory.Read.All` and app-role reads to `https://graph.microsoft.com/Application.Read.All`, while application profiles use the configured `.default` audience; reads never request a write-consent scope such as `DelegatedPermissionGrant.ReadWrite.All`, `Application.ReadWrite.All` or `Directory.ReadWrite.All`.
Delegated callers additionally need a supported directory role per operation (for example Directory Readers, Global Reader or Application Administrator).
Denied reads name that role and read-scope requirement instead of only the generic grant/role/licence cause.

`entra group list` defaults to compact properties (`id`, `displayName`, `mail`, `groupTypes`); `entra group show --id <group-id>` defaults to the richer reviewed group set including `isAssignableToRole`, which marks groups eligible for role assignment.
Group `--select` accepts the [reviewed group property set](src/entra-groups.ts); `--fields` must be a subset of the fetched selection.
Role-assignable membership changes need role-management permission and are refused by the membership write; that grant belongs to a later slice, never to this command.
`entra group member add --group <group-id> --user <user-id>` previews adding one user to one non-role-assignable security or Microsoft 365 group through a directoryObjects reference.
Both identifiers must be GUID object IDs; UPNs are not resolved, and only v1.0 is supported.
Role-assignable, dynamic-membership and distribution groups are refused.
The command verifies the user through `/users/<user-id>` before preview and again before sending; this read needs delegated User.ReadBasic.All or application User.Read.All, and the group and membership reads need D/A GroupMember.Read.All.
The command needs a hand-enabled profile whose writes allow `mg.entra.group.member.add`; the mutation needs D/A GroupMember.ReadWrite.All, and delegated callers additionally need a groups role such as Groups Administrator.
Delegated `--scopes` overrides only the mutation scope, which defaults to `https://graph.microsoft.com/GroupMember.ReadWrite.All`; application profiles reject `--scopes` and use their configured Graph `.default` audience.
Without `--execute`, the command previews without journaling; `--execute --confirm <group-id>` sends once after a fresh read.
An already-member user is a no-op without journaling when found by the initial read; a duplicate-reference 400 after sending is a journaled no-op.
Execution reserves a journaled intent before the fresh read and records its outcome; uncertain outcomes require reading back the membership before any further action, with no automatic retry or replay.
`entra group member list --group <group-id>` lists direct members and `entra group member-of list --group <group-id>` lists direct memberships; `--transitive` selects the flat nested closure instead.
Relationship rows default to `id` and `displayName`; `--select` accepts only `id`, `displayName` and `mail`, and `--fields` must be a subset of that selection.
Returned `@odata.type` stays visible alongside any `--fields` projection.
Group lists return `groups`, member lists return `members`, parent-membership lists return `memberOf`, and single-group reads return `group`.
Group and relationship lists report uniform totals through the shared list-totals helper: `total` holds the server-supplied `@odata.count` when the query carries `$count` (filtered reads) and `null` otherwise, beside a `count` line such as `3 of 10 groups`, `3 groups` or, when the total is unknown and more pages remain, `3 groups shown, more available`.
The same helper covers the other wired core-directory lists: users, applications, service principals, owners, devices, administrative units and unit members, contacts, direct reports and memberships, directory roles, role assignments and PIM eligible/active assignments, group PIM assignment/eligibility schedules, instances and eligibility requests, directory objects, deleted items, domains and DNS records, contracts, organizations and branding localizations, subscriptions, access-review definitions, instances, decisions, contacted reviewers and stages, sign-ins and directory audits, authentication methods and registration reports, PKI configurations and certificate authorities, Conditional Access policies, named locations, authentication-strength policies, combination configurations, authentication method modes, templates, authentication contexts, deleted policies and deleted named locations, attribute sets, custom-security-attribute definitions and allowed values, data-policy operations, delegated-admin customers, relationships and navigation lists, entitlement catalogs, access packages, assignment policies, resource-role scopes, assignments and assignment requests, federation configurations, oauth2 grants and app-role assignments, group lifecycle policies and setting templates, identity providers, lifecycle workflows, workflow templates, task definitions, runs, processing results and task reports, multi-tenant-organization tenants, on-premises synchronizations, risky users, risk detections, risky service principals and their history and detections, fraud-protection, web-application-firewall and verification providers, and agreements and acceptances. Every entra list command now reports through the same helper; only the dedicated `count` and `available-types` commands keep their `count` aggregates.
`complete` names pagination state, with `reason` and an opaque `cursor` on partial reads.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to group and relationship reads.
Resume relationships with the same `--group` and direct or `--transitive` command, profile, scopes and API version; omit or repeat the original server query flags, and repeat local `--fields` and `--full` when wanted.
Help warns that hidden members are omitted without `Member.Read.Hidden`; completion describes pagination, not visibility.
Rows without non-null selected descriptive properties are preserved and reported as possibly limited by consent or unset properties; null properties stay null.
Direct member results always carry the [v1.0 service-principal limitation](docs/graph-coverage.md#licensing-and-completeness-findings) warning; there is no silent beta or expansion fallback.
`--filter` on group collections is sent with `$count=true` and `ConsistencyLevel: eventual`, which the relationship endpoints require.
Delegated group reads default to `https://graph.microsoft.com/GroupMember.Read.All`; hidden members need `Member.Read.Hidden` and richer group properties may need `Group.Read.All`, while application profiles use the configured `.default` audience.
When richer group access is needed, pass `--scopes https://graph.microsoft.com/Group.Read.All`; for hidden-member access, explicitly log in and read with `--scopes https://graph.microsoft.com/GroupMember.Read.All,https://graph.microsoft.com/Member.Read.Hidden` and satisfy the operation's delegated role requirements.

Read directory roles, current assignments and PIM activity through four views; delegated profiles first need explicit login with the read scopes:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/RoleManagement.Read.Directory,https://graph.microsoft.com/RoleEligibilitySchedule.Read.Directory,https://graph.microsoft.com/RoleAssignmentSchedule.Read.Directory
mg-axi entra directory-role list --profile soc --limit 10
mg-axi entra directory-role show --profile soc --id <role-id>
mg-axi entra role-assignment list --profile soc --filter "principalId eq '<principal-id>'"
mg-axi entra pim eligible list --profile soc
mg-axi entra pim active list --profile soc --filter "assignmentType eq 'Activated'"
```

`entra directory-role list` defaults to `id`, `displayName`, `description` and `roleTemplateId`; `entra directory-role show --id <role-id>` defaults to the full reviewed role set.
Directory roles are activated instances only: a role appears after activation, never before, so an empty result never proves the role does not exist.
`entra role-assignment list` defaults to `id`, `principalId`, `roleDefinitionId` and `directoryScopeId` and returns the current assignment inventory, including direct and PIM-activated assignments.
Use `entra pim active list` to classify those assignments as described below.
`entra pim eligible list` defaults to `id`, `principalId`, `roleDefinitionId` and `memberType`; it covers PIM-eligible assignments, which are not active, and eligible instances carry no `assignmentType`.
`entra pim active list` defaults to those properties plus `assignmentType` and covers both directly assigned (`assignmentType` Assigned) and activated eligible (`assignmentType` Activated) assignments; `memberType` names how the instance reaches the principal.
For built-in roles the unified `roleDefinitionId` matches the directory-role `roleTemplateId`.
`--select` requests properties from the [reviewed role property sets](src/entra-roles.ts); `--fields` must be a subset of the fetched selection.
`--filter` passes through as plain `$filter` with no `$count` or `ConsistencyLevel` contract; the named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply.
Delegated directory-role and role-assignment reads default to `https://graph.microsoft.com/RoleManagement.Read.Directory`; eligible PIM reads default to `https://graph.microsoft.com/RoleEligibilitySchedule.Read.Directory` and active PIM reads to `https://graph.microsoft.com/RoleAssignmentSchedule.Read.Directory`, while application profiles use the configured `.default` audience.
Delegated callers additionally need a supported directory role per operation (for example Privileged Role Administrator, Global Reader or Security Reader for role reads, Directory Readers for assignments, Security Operator for PIM reads).
Denied reads name that role requirement instead of only the generic grant/role/licence cause.
Built-in roles are base inventory and custom role assignments need P1; PIM reads need P2 or ID Governance.
Role assignment, activation and every other PIM mutation belongs to later write slices, never to these reads.

Read PIM-for-Groups assignment and eligibility through ten views; delegated profiles first need explicit login with the read scopes:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/PrivilegedAssignmentSchedule.Read.AzureADGroup,https://graph.microsoft.com/PrivilegedEligibilitySchedule.Read.AzureADGroup
mg-axi entra pim group-assignment-schedule list --profile soc --filter "groupId eq '<group-id>'"
mg-axi entra pim group-assignment-schedule show --profile soc --id <assignment-schedule-id>
mg-axi entra pim group-assignment-instance list --profile soc --filter "principalId eq '<principal-id>'"
mg-axi entra pim group-assignment-instance show --profile soc --id <assignment-instance-id>
mg-axi entra pim group-eligibility-schedule list --profile soc --filter "groupId eq '<group-id>'"
mg-axi entra pim group-eligibility-schedule show --profile soc --id <eligibility-schedule-id>
mg-axi entra pim group-eligibility-instance list --profile soc --filter "groupId eq '<group-id>'"
mg-axi entra pim group-eligibility-instance show --profile soc --id <eligibility-instance-id>
mg-axi entra pim group-eligibility-request list --profile soc --filter "groupId eq '<group-id>'"
mg-axi entra pim group-eligibility-request show --profile soc --id <eligibility-request-id>
```

Assignment schedules govern group membership or ownership over time and instances are the provisioned windows, never the schedule itself.
Eligibility schedules govern who may activate and are not active grants, while eligibility instances are the provisioned eligibility windows and requests carry the ask plus its outcome.
Every list requires `--filter` with an `eq` clause on `groupId` or `principalId` scoping the list to one group or principal, and the requirement fails validation before credentials.
Assignment schedule and instance lists default to `id`, `principalId`, `groupId`, `accessId` and `assignmentType`; eligibility schedule and instance lists default to `id`, `principalId`, `groupId`, `accessId` and `memberType`; eligibility-request lists default to `id`, `action`, `status`, `principalId`, `groupId` and `accessId`.
Every show defaults to its full reviewed property set, except eligibility-request show omits `justification` unless an explicit `--select` names it.
`--select` requests properties from the [reviewed group PIM property sets](src/entra-group-pim.ts); `--fields` must be a subset of the fetched selection.
`--filter` passes through as plain `$filter` with no `$count` or `ConsistencyLevel` contract; these collections document no `$count`, so the uniform `count` totals name the returned rows with a `null` total unless a page carries a server total.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply; resume a list with the same `--filter`, profile, scopes and API version.
Delegated assignment reads default to `https://graph.microsoft.com/PrivilegedAssignmentSchedule.Read.AzureADGroup` and eligibility reads to `https://graph.microsoft.com/PrivilegedEligibilitySchedule.Read.AzureADGroup`, while application profiles use the configured `.default` audience and reject `--scopes`.
Delegated callers additionally need owner or member of the group or a supported directory role scoped at directory level (Global Reader or Privileged Role Administrator for role-assignable groups; otherwise Global Reader, Directory Writer, Groups Administrator, Identity Governance Administrator or User Administrator).
Denied reads name that scope, role and licensing requirement instead of only the generic grant/role/licence cause.
PIM for Groups needs P2 or ID Governance, and delegated personal Microsoft accounts are not supported.
Assignment requests, approvals, `filterByCurrentUser` functions, group/principal navigation, resources, every count/ref/cast tail, beta and all mutations stay out; assignment requests stay out because their least privilege is the write scope `PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup`.

Read access reviews through ten views; delegated profiles first need explicit login with the read scope:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/AccessReview.Read.All
mg-axi entra access-review definition list --profile soc --limit 10
mg-axi entra access-review definition list --profile soc --filter "status eq 'InProgress'"
mg-axi entra access-review definition show --profile soc --id <definition-id>
mg-axi entra access-review instance list --profile soc --definition <definition-id>
mg-axi entra access-review instance show --profile soc --definition <definition-id> --id <instance-id>
mg-axi entra access-review decision list --profile soc --definition <definition-id> --instance <instance-id>
mg-axi entra access-review decision show --profile soc --definition <definition-id> --instance <instance-id> --id <decision-id>
mg-axi entra access-review contacted-reviewer list --profile soc --definition <definition-id> --instance <instance-id>
mg-axi entra access-review contacted-reviewer show --profile soc --definition <definition-id> --instance <instance-id> --id <reviewer-id>
mg-axi entra access-review stage list --profile soc --definition <definition-id> --instance <instance-id>
mg-axi entra access-review stage show --profile soc --definition <definition-id> --instance <instance-id> --id <stage-id>
```

Definitions are review schedules (a series) and never carry their occurrences: each recurrence creates one instance per reviewed resource, and a one-time review creates one instance per resource.
Instances are occurrences of one definition schedule, and each reviewed principal or resource in an instance carries one decision item.
`entra access-review definition list` defaults to `id`, `displayName` and `status`; `entra access-review definition show --id <definition-id>` defaults to the full reviewed schedule set.
`entra access-review instance list --definition <definition-id>` defaults to `id`, `status`, `startDateTime` and `endDateTime`; `entra access-review instance show` takes both `--definition` and `--id` and defaults to the full reviewed occurrence set.
`entra access-review decision list --definition <definition-id> --instance <instance-id>` defaults to `id`, `accessReviewId`, `decision` and `recommendation`, where `accessReviewId` names the parent instance; `entra access-review decision show` takes `--definition`, `--instance` and `--id` and defaults to the full reviewed outcome set.
`entra access-review contacted-reviewer list --definition <definition-id> --instance <instance-id>` defaults to `id`, `displayName` and `userPrincipalName`; contacted reviewers are reviewer identities recorded on one instance, whether or not notified, never review outcomes.
`entra access-review stage list --definition <definition-id> --instance <instance-id>` defaults to `id`, `status`, `startDateTime` and `endDateTime`; stages are sequential phases of one instance (up to three when the definition sets `stageSettings`), and an instance without `stageSettings` has no stages.
`durationInDays` is not a stage property; per-stage decisions belong to a later slice.
Contacted-reviewer and stage `show` commands take `--definition`, `--instance` and `--id` and default to their full [reviewed property sets](src/entra-access-reviews.ts).
Decision reads are read-only: listing never approves, denies or applies anything, and submitting or stopping a review belongs to a later slice, never to these reads.
`--select` requests properties from the [reviewed access-review property sets](src/entra-access-reviews.ts); `--fields` must be a subset of the fetched selection.
`--filter` passes through as plain `$filter` with no `$count` or `ConsistencyLevel` contract; on definitions only `contains()` over the scope query and `eq` on status are documented, on stages only `eq` is documented.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to access-review reads.
Resume any access-review list with `--cursor -` and supply the returned cursor on stdin, for example `mg-axi entra access-review definition list --profile soc --cursor - < cursor.txt`; small cursors can also use `--cursor <token>`.
Resume instance, decision, contacted-reviewer and stage lists with the same `--definition` (and `--instance`), profile, scopes and API version; a cursor from another definition or instance fails validation instead of returning foreign rows.
To replay a resumed result with `--full`, supply the original input cursor on stdin; the returned cursor continues after that result.
Delegated reads default to `https://graph.microsoft.com/AccessReview.Read.All`, while application profiles use the configured `.default` audience.
See the [access-review scope decisions](docs/coverage.md#ext-02-access-review-scope-decisions) for deferred legacy/unified reads and blocked history reads.
Delegated callers additionally need a supported Entra role per review scope: group or app reviews need the review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator, while Entra-role reviews need Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator.
Denied reads name that role requirement instead of only the generic grant/role/licence cause.
Access reviews need P2 or ID Governance depending on capability, not one uniform licence, and delegated personal Microsoft accounts are not supported.

Read entitlement catalogs and access packages through twelve views; delegated profiles first need explicit login with the read scope:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/EntitlementManagement.Read.All
mg-axi entra entitlement catalog list --profile soc --limit 10
mg-axi entra entitlement catalog list --profile soc --filter "state eq 'Published'"
mg-axi entra entitlement catalog show --profile soc --id <catalog-id>
mg-axi entra entitlement catalog count --profile soc
mg-axi entra entitlement access-package list --profile soc --limit 10
mg-axi entra entitlement access-package show --profile soc --id <access-package-id>
mg-axi entra entitlement access-package count --profile soc
mg-axi entra entitlement assignment-policy list --profile soc --access-package <access-package-id>
mg-axi entra entitlement assignment-policy show --profile soc --access-package <access-package-id> --id <policy-id>
mg-axi entra entitlement assignment-policy count --profile soc --access-package <access-package-id>
mg-axi entra entitlement resource-role-scope list --profile soc --access-package <access-package-id>
mg-axi entra entitlement resource-role-scope show --profile soc --access-package <access-package-id> --id <scope-id>
mg-axi entra entitlement resource-role-scope count --profile soc --access-package <access-package-id>
```

Catalogs are package containers and never carry their access packages: `entra entitlement catalog list` defaults to `id`, `displayName`, `state` and `catalogType`, and `entra entitlement catalog show --id <catalog-id>` defaults to the full reviewed container set.
Access packages are the assignable bundles: `entra entitlement access-package list` defaults to `id`, `displayName` and `isHidden`, and `entra entitlement access-package show --id <access-package-id>` defaults to the full reviewed bundle set.
`entra entitlement assignment-policy list --access-package <access-package-id>` defaults to `id`, `displayName` and `allowedTargetScope`: policies name who may request and how approval and review run, and `entra entitlement assignment-policy show` takes `--access-package` and `--id` and defaults to the full reviewed request, approval and review set.
`entra entitlement resource-role-scope list --access-package <access-package-id>` defaults to `id` and `createdDateTime`: role-plus-scope pairs a package grants, and `entra entitlement resource-role-scope show` takes `--access-package` and `--id`.
Approval and subject reads carry personal data and belong to later parts, never to these reads; questions, custom-extension stages and the linked role/scope detail need `$expand` and belong to later slices.
`--select` requests properties from the [reviewed entitlement property sets](src/entra-entitlement-management.ts); `--fields` must be a subset of the fetched selection.
`--filter` passes through as plain `$filter` with no `$count` or `ConsistencyLevel` contract.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to entitlement reads.
Resume any entitlement list with `--cursor -` and supply the returned cursor on stdin; resume policy and role-scope lists with the same `--access-package`, profile, scopes and API version.
Delegated reads default to `https://graph.microsoft.com/EntitlementManagement.Read.All`, while application profiles use the configured `.default` audience.
See the [entitlement-management scope decisions](docs/coverage.md#ext-02-entitlement-management-scope-decisions) for the scheduled personal-data part and deferred beta reads.
Delegated callers additionally need a supported Entra role with catalog visibility (Global Reader and Identity Governance Administrator are among the supported roles).
Denied reads name that role requirement instead of only the generic grant/role/licence cause.
Entitlement management needs P2 or ID Governance depending on capability, not one uniform licence, and delegated personal Microsoft accounts are not supported.

Read assignments and assignment requests through six views; they use the same scope default and login as the twelve reads above:

```sh
mg-axi entra entitlement assignment list --profile soc --limit 10
mg-axi entra entitlement assignment list --profile soc --filter "state eq 'Delivered'"
mg-axi entra entitlement assignment show --profile soc --id <assignment-id>
mg-axi entra entitlement assignment count --profile soc
mg-axi entra entitlement assignment-request list --profile soc --limit 10
mg-axi entra entitlement assignment-request show --profile soc --id <request-id>
mg-axi entra entitlement assignment-request count --profile soc
```

The wire resources carry no target or access-package linkage as scalars, so each list and show sends one fixed documented `$expand` set in code (target and accessPackage for assignments; accessPackage and assignment for requests) and projects only flattened linkage identifiers: `entra entitlement assignment list` defaults to `id`, `state`, `targetId`, `targetDisplayName`, `accessPackageId`, `expiredDateTime` and `schedule`, and `entra entitlement assignment-request list` defaults to `id`, `requestType`, `state`, `accessPackageId`, `assignmentId`, `createdDateTime`, `completedDateTime` and `schedule`.
There is no user-supplied `$expand` anywhere: `--select` names only reviewed scalar or flattened fields, and navigation names fail as unknown properties.
Request justification text and answers are personal-data payloads and ride only behind an explicit `--select`, even on show; the requestor subject, the assignment policy link and the `filterByCurrentUser`/`additionalAccess` functions belong to later slices.
Callers with only catalog-scoped roles must filter to one access package, for example `--filter "accessPackage/id eq '<access-package-id>'"`.
Assignment and request reads never mutate: no request creation, cancellation, approval or reprocessing, and no beta.
Denied assignment and request reads name the Catalog reader least-privileged role and the supported Entra roles instead of only the generic grant/role/licence cause.

Read lifecycle workflows, workflow templates, task definitions and tenant settings through ten views; workflow reads default to the ReadBasic scope while template, task-definition and settings reads default to the broader read scope:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/LifecycleWorkflows-Workflow.ReadBasic.All
mg-axi entra lifecycle workflow list --profile soc --limit 10
mg-axi entra lifecycle workflow list --profile soc --filter "category eq 'leaver'"
mg-axi entra lifecycle workflow show --profile soc --id <workflow-id>
mg-axi entra lifecycle workflow count --profile soc
mg-axi login --profile soc --scopes https://graph.microsoft.com/LifecycleWorkflows.Read.All
mg-axi entra lifecycle workflow-template list --profile soc --limit 10
mg-axi entra lifecycle workflow-template show --profile soc --id <template-id>
mg-axi entra lifecycle workflow-template count --profile soc
mg-axi entra lifecycle task-definition list --profile soc --limit 10
mg-axi entra lifecycle task-definition show --profile soc --id <task-definition-id>
mg-axi entra lifecycle task-definition count --profile soc
mg-axi entra lifecycle settings show --profile soc
```

Read one workflow's runs and user/subject processing results through nine views; reporting reads default to the reports scope:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/LifecycleWorkflows-Reports.Read.All
mg-axi entra lifecycle run list --workflow <workflow-id> --profile soc --limit 10
mg-axi entra lifecycle run list --workflow <workflow-id> --profile soc --filter "processingStatus eq 'completed'"
mg-axi entra lifecycle run show --workflow <workflow-id> --profile soc --id <run-id>
mg-axi entra lifecycle run count --workflow <workflow-id> --profile soc
mg-axi entra lifecycle user-processing-result list --workflow <workflow-id> --profile soc --limit 10
mg-axi entra lifecycle user-processing-result show --workflow <workflow-id> --profile soc --id <result-id>
mg-axi entra lifecycle user-processing-result count --workflow <workflow-id> --profile soc
mg-axi entra lifecycle subject-processing-result list --workflow <workflow-id> --profile soc --limit 10
mg-axi entra lifecycle subject-processing-result show --workflow <workflow-id> --profile soc --id <result-id>
mg-axi entra lifecycle subject-processing-result count --workflow <workflow-id> --profile soc
```

Read one workflow's task reports through five views; task-report reads share the reports scope:

```sh
mg-axi entra lifecycle task-report list --workflow <workflow-id> --profile soc --limit 10
mg-axi entra lifecycle task-report list --workflow <workflow-id> --profile soc --filter "processingStatus eq 'completed'"
mg-axi entra lifecycle task-report show --workflow <workflow-id> --profile soc --id <report-id>
mg-axi entra lifecycle task-report count --workflow <workflow-id> --profile soc
mg-axi entra lifecycle task-report task show --workflow <workflow-id> --report <report-id> --profile soc
mg-axi entra lifecycle task-report task-definition show --workflow <workflow-id> --report <report-id> --profile soc
```

Read one run's user, subject and task processing results through nine views; run-nested reads share the reports scope:

```sh
mg-axi entra lifecycle run user-processing-result list --workflow <workflow-id> --run <run-id> --profile soc --limit 10
mg-axi entra lifecycle run user-processing-result list --workflow <workflow-id> --run <run-id> --profile soc --filter "processingStatus eq 'completed'"
mg-axi entra lifecycle run user-processing-result show --workflow <workflow-id> --run <run-id> --profile soc --id <result-id>
mg-axi entra lifecycle run user-processing-result count --workflow <workflow-id> --run <run-id> --profile soc
mg-axi entra lifecycle run subject-processing-result list --workflow <workflow-id> --run <run-id> --profile soc --limit 10
mg-axi entra lifecycle run subject-processing-result show --workflow <workflow-id> --run <run-id> --profile soc --id <result-id>
mg-axi entra lifecycle run subject-processing-result count --workflow <workflow-id> --run <run-id> --profile soc
mg-axi entra lifecycle run task-processing-result list --workflow <workflow-id> --run <run-id> --profile soc --limit 10
mg-axi entra lifecycle run task-processing-result show --workflow <workflow-id> --run <run-id> --profile soc --id <result-id>
mg-axi entra lifecycle run task-processing-result count --workflow <workflow-id> --run <run-id> --profile soc
```

`entra lifecycle run task-processing-result list` defaults to `id`, `processingStatus` and `failureReason`, and `entra lifecycle run task-processing-result show` defaults to the full reviewed scalar set including failure detail; task rows never carry subject, task or workflowSubject detail.
Run user results share the top-level user query contract (`--select` reviewed, `--filter` plain), while run subject rows always arrive whole like their top-level counterparts (Graph documents no `$select` on the list and `$expand` only on the get) and the run task single has no operation-level page, so both arrive whole and `--fields` projects them locally; `--select` is refused before credentials.
User, subject and task processing results are personal data: user and subject rows default to status and counts only (subject rows add `subjectType`), and the subject link is never requested.
Third-level processing results, reprocessed runs, subject and task sub-reads and the time-windowed summary functions belong to later slices.

`entra lifecycle task-report list` defaults to `id`, `runId`, `processingStatus`, `totalUsersCount`, `failedUsersCount` and `successfulUsersCount`, and `entra lifecycle task-report show` defaults to the full reviewed scalar set; report rows never carry task, task-definition or processing-result detail.
The task-report get and the taskReport-nested task and taskDefinition singles document no query parameters, so those rows always arrive whole and `--fields` projects them locally; `--select` is refused before credentials.
Third-level processing results, reprocessed runs, subject sub-reads and the time-windowed summary functions belong to later slices.

`entra lifecycle run list` defaults to `id`, `processingStatus`, `totalUsersCount`, `failedUsersCount` and `successfulUsersCount`, and `entra lifecycle run show` defaults to the full reviewed scalar set; run rows never carry processing results.
User and subject processing results are personal data: lists and singles default to status and counts only (subject lists add `subjectType`), and the subject link is never requested.
The subject list documents `$filter` but not `$select` and the subject get documents `$expand` only, so subject rows always arrive whole and `--fields` projects them locally; `--select` is refused before credentials.
Third-level processing results, taskProcessingResults sub-reads under a user or subject result, reprocessed runs, subject sub-reads and the time-windowed summary functions belong to later slices.

`entra lifecycle workflow list` defaults to `id`, `displayName`, `category`, `isEnabled` and `isSchedulingEnabled`, and `entra lifecycle workflow show --id <workflow-id>` defaults to the full reviewed scalar set.
Workflow rows never carry tasks, runs, processing results or task reports: tasks ride expanded by default on the get and are dropped in local projection, and third-level processing results and summaries belong to later parts, never to these reads.
`entra lifecycle workflow-template list` defaults to `id`, `displayName` and `category`; the template get documents no query parameters, so `entra lifecycle workflow-template show --id <template-id>` takes no `--select` and always returns whole rows.
`entra lifecycle task-definition list` defaults to `id`, `displayName`, `category` and `version`, and `entra lifecycle task-definition show --id <task-definition-id>` defaults to the full reviewed parameter set.
`entra lifecycle settings show` returns the tenant schedule interval with the email and quarantine configuration; updating settings belongs to no read slice.
`--select` requests properties from the [reviewed lifecycle property sets](src/entra-lifecycle-workflows.ts); `--fields` must be a subset of the fetched selection, and navigation names (`tasks`, `createdBy`, `lastModifiedBy`, `previewScope`) fail as unknown properties.
`--filter` passes through as plain `$filter` with no `$count` or `ConsistencyLevel` contract; `$search`, `$orderby` and `$expand` stay unreviewed.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to lifecycle reads.
Resume any lifecycle list with `--cursor -` and supply the returned cursor on stdin.
Delegated workflow reads default to `https://graph.microsoft.com/LifecycleWorkflows-Workflow.ReadBasic.All` (richer detail may need `LifecycleWorkflows-Workflow.Read.All` or `LifecycleWorkflows.Read.All`), template, task-definition and settings reads to `https://graph.microsoft.com/LifecycleWorkflows.Read.All`, and run, processing-result and task-report reads to `https://graph.microsoft.com/LifecycleWorkflows-Reports.Read.All` (richer detail may need `LifecycleWorkflows.Read.All` or `LifecycleWorkflows.ReadWrite.All`), while application profiles use the configured `.default` audience.
See the [lifecycle-workflows scope decisions](docs/coverage.md#ext-02-lifecycle-workflows-scope-decisions) for the scheduled nested-result and summary parts.
Delegated callers additionally need Global Reader or Lifecycle Workflows Administrator.
Denied reads name that scope, role and licensing requirement instead of only the generic grant/role/licence cause.
Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite (every governed user, not only administrators), and delegated personal Microsoft accounts are not supported.

Log in with `https://graph.microsoft.com/Device.Read.All` or `https://graph.microsoft.com/AdministrativeUnit.Read.All`, then inspect directory devices and administrative units:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Device.Read.All
mg-axi entra device list --profile soc --limit 10
mg-axi entra device show --profile soc --id <device-id>
mg-axi login --profile soc --scopes https://graph.microsoft.com/AdministrativeUnit.Read.All
mg-axi entra administrative-unit list --profile soc
mg-axi entra administrative-unit show --profile soc --id <administrative-unit-id>
mg-axi entra administrative-unit member list --profile soc --administrative-unit <administrative-unit-id>
```

`entra device list` defaults to compact properties (`id`, `displayName`, `operatingSystem`, `accountEnabled`); `entra device show --id <device-id>` takes the object `id`, not `deviceId`, and defaults to the full reviewed device set.
Directory devices are Entra directory objects; Intune managed devices and device actions are a separately authorized surface, never these commands.
`entra administrative-unit list` defaults to `id`, `displayName`, `visibility` and `membershipType`; `entra administrative-unit show --id <administrative-unit-id>` defaults to the full reviewed unit set including the membership rule.
Device and unit `--select` accept their [reviewed property sets](src/entra-directory.ts); `--fields` must be a subset of the fetched selection.
Unit show adds a licensing hint when the projected `membershipType` is `Dynamic`; custom `--select` or `--fields` that omit it also omit the hint.
See the [licence matrix](docs/graph-coverage.md#licence-matrix-by-area) for device and administrative-unit licensing requirements.
`entra administrative-unit member list --administrative-unit <administrative-unit-id>` lists member users, groups and devices with the same `id`/`displayName`/`mail` selection, `@odata.type` preservation, hidden-membership and limited-information behavior as group relationships.
Device and unit lists return `devices` and `administrativeUnits`, member lists return `members`, and single-object reads return `device` and `administrativeUnit`.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to device and administrative-unit reads.
Resume unit-member lists with the same `--administrative-unit`, profile, scopes and API version; omit or repeat the original server query flags, and repeat local `--fields` and `--full` when wanted.
Delegated device reads default to `https://graph.microsoft.com/Device.Read.All` and unit reads to `https://graph.microsoft.com/AdministrativeUnit.Read.All`, while application profiles use the configured `.default` audience; hidden unit memberships need `Member.Read.Hidden`.
Denied directory reads return operation-specific permission, delegated-role and licensing guidance rather than empty results; HTTP 403 alone does not identify which prerequisite is missing.
`--filter` on these collections is sent with `$count=true` and `ConsistencyLevel: eventual`.

Log in with `https://graph.microsoft.com/Organization.Read.All` for organization reads and `https://graph.microsoft.com/User.Read` for branding reads, then inspect the tenant and its sign-in branding:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Organization.Read.All,https://graph.microsoft.com/User.Read
mg-axi entra organization list --profile soc
mg-axi entra organization show --profile soc --id <organization-id>
mg-axi entra organization branding show --profile soc --organization <organization-id>
mg-axi entra organization branding-localization list --profile soc --organization <organization-id>
mg-axi entra organization branding-localization show --profile soc --organization <organization-id> --id fr-FR
```

`entra organization list` defaults to compact properties (`id`, `displayName`, `tenantType`, `verifiedDomains`); exactly one organization exists per tenant.
`entra organization show --id <organization-id>` defaults to the full reviewed organization set including technical notification mails and the privacy profile.
`--select` requests properties from the [reviewed organization and branding field sets](src/entra-organization.ts); `--fields` projects locally and must be a subset of the fetched selection.
Delegated organization reads default to `https://graph.microsoft.com/Organization.Read.All` for full metadata; delegated `User.Read` returns only `id`, `displayName` and `verifiedDomains` with every other property null.
`entra organization branding show --organization <organization-id>` reads the default branding metadata (non-Stream text and URLs); the session sends the documented `Accept-Language: 0` header and locale variants come from the localizations collection.
Branding leaves default to delegated `https://graph.microsoft.com/User.Read`, the documented least-privileged scope; `OrganizationalBranding.Read.All` is the purpose-built alternative and `Organization.Read.All` also works when passed as `--scopes`.
Application profiles need admin-consented `Organization.Read.All` for organization metadata or `OrganizationalBranding.Read.All` for branding, using the configured Graph `.default` audience.
Stream image properties (`bannerLogo`, `backgroundImage` and friends) are refused before credentials: they need a later piece with its own binary-output contract.
A branding 404 may indicate unconfigured branding, a missing locale, or a missing or inaccessible organization; configuring custom branding needs P1/P2, and contact fields on the organization are personal data.
Delegated callers additionally need a supported Entra role (Directory Readers or Global Reader for organizations; Global Reader or Organizational Branding Administrator for branding); personal Microsoft accounts are not supported.
Organization and localization lists return `organizations` and `brandingLocalizations`, single-object reads return `organization`, `branding` and `brandingLocalization`.
Localization lists default to `id`, `signInPageText`, `usernameHintText` and `backgroundColor`; branding and localization show commands default to the full reviewed non-Stream field set.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to organization and branding reads.
All five named reads default to `--api-version v1.0`; explicit `--api-version beta` requires a preview-enabled profile, with no automatic fallback.
Organization and localization lists offer no `--filter`: Graph documents `$select` only on these routes, so the flag is refused before credentials.
Denied organization and branding reads name the scope, role and licensing guidance instead of only the generic cause.
No organization mutation lives here; organization-scoped certificate-based-auth configuration, extensions, beta-only settings and the POST lookup actions belong to later pieces; see the [organization scope decisions](docs/coverage.md#ext-01-organization-scope-decisions) for deferred reads and later subfamilies.

Log in with `https://graph.microsoft.com/Directory.Read.All`, then inspect group lifecycle policies:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Directory.Read.All
mg-axi entra group-lifecycle-policy list --profile soc
mg-axi entra group-lifecycle-policy show --profile soc --id <policy-id>
mg-axi entra group-lifecycle-policy count --profile soc
```

Log in with `https://graph.microsoft.com/GroupSettings.Read.All`, then inspect group setting templates:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/GroupSettings.Read.All
mg-axi entra group-setting-template list --profile soc
mg-axi entra group-setting-template show --profile soc --id <template-id>
mg-axi entra group-setting-template count --profile soc
```

`entra group-lifecycle-policy list` defaults to compact properties (`id`, `groupLifetimeInDays`, `managedGroupTypes`); `entra group-lifecycle-policy show --id <policy-id>` defaults to the full reviewed lifecycle set (`id`, `alternateNotificationEmails`, `groupLifetimeInDays`, `managedGroupTypes`).
`entra group-setting-template list` defaults to compact properties (`id`, `displayName`, `description`); `entra group-setting-template show --id <template-id>` defaults to the full reviewed template set (`id`, `deletedDateTime`, `displayName`, `description`, `values`).
`--select` requests properties from the [reviewed group lifecycle field sets](src/entra-group-lifecycle.ts); `--fields` projects locally and must be a subset of the fetched selection.
Lifecycle-policy lists accept `--filter` as plain `$filter` without adding `$count=true` or `ConsistencyLevel`; template lists offer no `--filter` because Graph documents `$select` only for `/groupSettingTemplates`.
Each `count` returns one scalar (`groupLifecyclePolicyCount`, `groupSettingTemplateCount`) from its text/plain `$count` route and takes no `--filter`, `--select`, `--limit` or `--cursor`.
Raw `api get` supports `/groupLifecyclePolicies`, `/groupLifecyclePolicies/<policy-id>`, `/groupSettingTemplates` and `/groupSettingTemplates/<template-id>`; the `$count` routes are available only through the named count commands.
Lifecycle-policy lists return `groupLifecyclePolicies`, single-policy reads return `groupLifecyclePolicy`, template lists return `groupSettingTemplates` and single-template reads return `groupSettingTemplate`.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to these reads.
Named reads truncate top-level strings and strings in string arrays at 500 characters with a `--full` recovery hint; nested objects in template `values` pass through without text truncation.
All six named reads support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Delegated lifecycle-policy reads default to `https://graph.microsoft.com/Directory.Read.All` and delegated template reads default to `https://graph.microsoft.com/GroupSettings.Read.All`, while application profiles use the configured `.default` audience.
Delegated template callers additionally need a supported Entra role (Directory Readers or Global Reader are the least-privileged roles); no delegated role prerequisite is stated for lifecycle-policy reads; personal Microsoft accounts are not supported on either family.
No P1/P2 prerequisite is stated for these reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
No lifecycle or template mutation lives here; `groupSettingTemplates/delta()`, beta policies and templates and the POST lookup actions belong to later pieces; see the [group lifecycle scope decisions](docs/coverage.md#ext-01-group-lifecycle-scope-decisions) for deferred reads and later subfamilies.

Log in with `https://graph.microsoft.com/CustomSecAttributeDefinition.Read.All`, then inspect attribute sets, custom security attribute definitions and their allowed values:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/CustomSecAttributeDefinition.Read.All
mg-axi entra attribute-set list --profile soc
mg-axi entra attribute-set show --profile soc --id <set-id>
mg-axi entra attribute-set count --profile soc
mg-axi entra custom-security-attribute-definition list --profile soc
mg-axi entra custom-security-attribute-definition show --profile soc --id <definition-id>
mg-axi entra custom-security-attribute-definition count --profile soc
mg-axi entra allowed-value list --profile soc --definition <definition-id>
mg-axi entra allowed-value show --profile soc --definition <definition-id> --id <value-id>
mg-axi entra allowed-value count --profile soc --definition <definition-id>
```

`entra attribute-set list` defaults to the reviewed set (`id`, `description`, `maxAttributesPerSet`); `entra custom-security-attribute-definition list` defaults to compact properties (`id`, `attributeSet`, `name`, `status`, `type`) while show defaults to the full reviewed definition set; `entra allowed-value list` defaults to the reviewed set (`id`, `isActive`).
`--select` requests properties from the [reviewed custom-security field sets](src/entra-custom-security-attributes.ts); `--fields` projects locally and must be a subset of the fetched selection.
Definition lists accept `--filter` as plain `$filter` (eq) without adding `$count=true` or `ConsistencyLevel`; attribute-set and allowed-value lists offer no `--filter` because Graph documents no `$filter` for those routes.
Allowed-value lists and shows take `--definition <definition-id>`; definition lists never inline `allowedValues` (`$expand` is not reviewed here) and show points at the allowed-value list instead.
Each `count` returns `count: { returned: <total>, complete: true }` from its text/plain `$count` route; only the definition count accepts `--filter` to narrow the total server-side, and no count takes `--select`, `--limit` or `--cursor`.
Raw `api get` supports `/directory/attributeSets`, `/directory/attributeSets/<set-id>`, `/directory/customSecurityAttributeDefinitions`, `/directory/customSecurityAttributeDefinitions/<definition-id>`, `/directory/customSecurityAttributeDefinitions/<definition-id>/allowedValues` and the single allowed-value route; the `$count` routes are available only through the named count commands.
Attribute-set lists return `attributeSets` and single-set reads return `attributeSet`; definition lists return `customSecurityAttributeDefinitions` and single-definition reads return `customSecurityAttributeDefinition`; allowed-value lists return `allowedValues` and single-value reads return `allowedValue`.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to these reads.
All nine named reads support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Delegated reads default to `https://graph.microsoft.com/CustomSecAttributeDefinition.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers additionally need a custom-security-attribute role even for Global Administrators (Attribute Definition Reader or Attribute Definition Administrator work for every read; some reads also accept Attribute Assignment Reader or Attribute Assignment Administrator); personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for these reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
No custom-security-attribute mutation lives here; beta attribute sets, definitions and allowed values belong to a later piece; see the [custom-security-attributes scope decisions](docs/coverage.md#ext-01-custom-security-attributes-scope-decisions) for deferred reads and later subfamilies.

Log in with `https://graph.microsoft.com/Directory.Read.All`, then inspect partner-tenant customer contracts:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Directory.Read.All
mg-axi entra contract list --profile soc
mg-axi entra contract show --profile soc --id <contract-id>
mg-axi entra contract count --profile soc
```

Contracts exist in partner tenants only (Cloud Solution Provider, Office 365 Syndication or Advisor programs); a non-partner tenant lists zero contracts, which is an answer rather than an error.
`entra contract list` defaults to compact properties (`id`, `displayName`, `contractType`, `defaultDomainName`); `entra contract show --id <contract-id>` defaults to the full reviewed contract set (`contractType`, `customerId`, `defaultDomainName`, `displayName`, `id`).
`--select` requests properties from the [reviewed contract field set](src/entra-contracts.ts); `--fields` projects locally and must be a subset of the fetched selection.
Contract lists accept `--filter` as plain `$filter` without adding `$count=true` or `ConsistencyLevel`; filtering is documented for `customerId`, `defaultDomainName` and `displayName`.
`entra contract count` returns one scalar (`contractCount`) from the text/plain `$count` route and takes no `--filter`, `--select`, `--limit` or `--cursor`.
Raw `api get` supports `/contracts` and `/contracts/<contract-id>`; `/contracts/$count` is available only through the named count command.
Contract lists return `contracts` and single-contract reads return `contract`.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to contract reads.
All three named reads support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Delegated contract reads default to `https://graph.microsoft.com/Directory.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers additionally need a supported Entra role (Directory Readers is the least-privileged role); personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for contract reads; denied reads name the scope, role, partner-tenant and licensing guidance instead of only the generic cause.
No contract mutation lives here; `contracts/delta()`, beta contracts and the POST lookup actions belong to later pieces; see the [partner contracts scope decisions](docs/coverage.md#ext-04-partner-contracts-scope-decisions) for deferred reads and later subfamilies.

Log in with `https://graph.microsoft.com/DelegatedAdminRelationship.Read.All`, then inspect partner-tenant delegated administration:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/DelegatedAdminRelationship.Read.All
mg-axi entra delegated-admin-customer list --profile soc
mg-axi entra delegated-admin-customer show --profile soc --id <customer-id>
mg-axi entra delegated-admin-relationship list --profile soc
mg-axi entra delegated-admin-relationship show --profile soc --id <relationship-id>
mg-axi entra delegated-admin-relationship list-access-assignments --profile soc --id <relationship-id>
mg-axi entra delegated-admin-relationship show-access-assignment --profile soc --id <relationship-id> --assignment-id <assignment-id>
mg-axi entra delegated-admin-relationship list-operations --profile soc --id <relationship-id>
mg-axi entra delegated-admin-relationship show-operation --profile soc --id <relationship-id> --operation-id <operation-id>
mg-axi entra delegated-admin-relationship list-requests --profile soc --id <relationship-id>
mg-axi entra delegated-admin-relationship show-request --profile soc --id <relationship-id> --request-id <request-id>
mg-axi entra delegated-admin-customer list-service-management-details --profile soc --id <customer-id>
mg-axi entra delegated-admin-customer show-service-management-detail --profile soc --id <customer-id> --detail-id <detail-id>
mg-axi entra delegated-admin-customer count --profile soc
mg-axi entra delegated-admin-customer count-service-management-details --profile soc --id <customer-id>
mg-axi entra delegated-admin-relationship count --profile soc
mg-axi entra delegated-admin-relationship count-access-assignments --profile soc --id <relationship-id>
mg-axi entra delegated-admin-relationship count-operations --profile soc --id <relationship-id>
mg-axi entra delegated-admin-relationship count-requests --profile soc --id <relationship-id>
```

Delegated-admin reads run in the partner tenant; customer objects are created by the system when a relationship exists and deleted when none remain, so a non-partner tenant lists zero customers, which is an answer rather than an error.
`entra delegated-admin-customer list` defaults to compact properties (`id`, `displayName`, `tenantId`); `entra delegated-admin-relationship list` defaults to (`id`, `displayName`, `status`, `customer`, `endDateTime`); both top-level show commands default to the full reviewed property set.
`--select` requests properties from the reviewed delegated-admin field sets; `--fields` projects locally and must be a subset of the fetched selection.
Both top-level lists accept `--filter` as plain `$filter` without adding `$count=true` or `ConsistencyLevel`.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to delegated-admin reads.
The access-assignment, operation and request navigation lists bind their parent relationship through `--id` and take `--assignment-id`, `--operation-id` and `--request-id` on their shows; the service-management-detail reads bind their parent customer through `--id` and take `--detail-id` on the show.
The navigation lists default to compact properties (`id`, `status`, `accessContainer`, `accessDetails` for assignments; `id`, `operationType`, `status`, `lastModifiedDateTime` for operations; `id`, `action`, `status`, `lastModifiedDateTime` for requests) while the service-management-detail reads always carry (`id`, `serviceName`, `serviceManagementUrl`); every show defaults to its full reviewed property set.
The access-assignment, operation and request lists accept `--filter` as plain `$filter` without adding `$count=true` or `ConsistencyLevel`; Graph documents no query parameters for service-management details, so those two reads take no `--select` or `--filter` and project `--fields` locally.
Each `count` returns one scalar (`delegatedAdminCustomerCount`, `delegatedAdminServiceManagementDetailCount`, `delegatedAdminRelationshipCount`, `delegatedAdminAccessAssignmentCount`, `delegatedAdminRelationshipOperationCount`, `delegatedAdminRelationshipRequestCount`) from its text/plain `$count` route and takes no `--filter`, `--select`, `--limit` or `--cursor`.
All eighteen named reads support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Delegated delegated-admin reads default to `https://graph.microsoft.com/DelegatedAdminRelationship.Read.All`, while application profiles use the configured `.default` audience.
Personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for delegated-admin reads; denied reads name the scope, partner-tenant and licensing guidance instead of only the generic cause.
No delegated-admin mutation lives here; the container root belongs to a later piece, while the multi-tenant-organization reads shipped in the next section and tenant-lookup functions ship as `entra tenant-information show` in the tenant-information usage below; see the [delegated-admin scope decisions](docs/coverage.md#ext-04-delegated-admin-scope-decisions) for deferred reads and later subfamilies.

Log in with `https://graph.microsoft.com/MultiTenantOrganization.Read.All`, then inspect the tenant multitenant organization:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/MultiTenantOrganization.Read.All
mg-axi entra multi-tenant-organization show --profile soc
mg-axi entra multi-tenant-organization join-request show --profile soc
mg-axi entra multi-tenant-organization tenant list --profile soc
mg-axi entra multi-tenant-organization tenant count --profile soc
```

At most one multitenant organization exists per tenant; a tenant outside any multitenant organization reads the container as state inactive with null properties and lists zero tenants, which is an answer rather than an error.
`entra multi-tenant-organization tenant list` defaults to compact properties (`tenantId`, `displayName`, `role`, `state`); both show commands default to the full reviewed property set.
`--select` requests properties from the reviewed multi-tenant-organization field sets; `--fields` projects locally and must be a subset of the fetched selection.
The tenant list accepts `--filter` as plain `$filter` without adding `$count=true` or `ConsistencyLevel`.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to multi-tenant-organization reads.
All four named reads support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Delegated multi-tenant-organization reads default to `https://graph.microsoft.com/MultiTenantOrganization.Read.All` (the lower-privileged delegated `MultiTenantOrganization.ReadBasic.All` returns displayName and tenantId only) and additionally need Security Reader or Global Reader, while application profiles use the configured `.default` audience.
Personal Microsoft accounts are not supported, and these reads run in the commercial Global service only.
Multi-tenant-organization participation needs Entra ID P1; denied reads name the scope, roles and licensing guidance instead of only the generic cause.
No multi-tenant-organization mutation lives here; the single-member read stays scheduled because its documented least privilege is the write scope `MultiTenantOrganization.ReadWrite.All`, and tenant-lookup functions ship as `entra tenant-information show` in the tenant-information usage below; see the [multi-tenant-organization scope decisions](docs/coverage.md#ext-04-multi-tenant-organization-scope-decisions) for the deferred read.

Log in with `https://graph.microsoft.com/CrossTenantInformation.ReadBasic.All`, then look up tenant information by domain name or tenant ID:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/CrossTenantInformation.ReadBasic.All
mg-axi entra tenant-information show --domain-name example.invalid --profile soc
mg-axi entra tenant-information show --tenant-id <tenant-id> --profile soc
```

`entra tenant-information show` needs exactly one of `--domain-name` or `--tenant-id`; combining them or omitting both fails validation before credentials.
Both lookups return the full reviewed tenantInformation set (`defaultDomainName`, `displayName`, `federationBrandName`, `tenantId`); Graph documents no query parameters here, so the lookup takes no `--select`, `--filter`, `--limit` or `--cursor`.
Both lookups support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Delegated tenant-information reads default to `https://graph.microsoft.com/CrossTenantInformation.ReadBasic.All`, while application profiles use the configured `.default` audience.
No Entra role is required for tenant-information lookups; personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for tenant-information lookups; denied reads name the scope and licensing guidance instead of only the generic cause.
Null/missing preservation and 500-character text truncation described above also apply to tenant-information reads; `--full` removes text truncation.
Lookup values bind only through the allowlisted session function-argument contract (validated, OData-quoted and encoded); every other parenthesised path keeps the refusal.
No tenant-information mutation lives here; see the [tenant-lookup scope decisions](docs/coverage.md#ext-04-tenant-lookup-scope-decisions) for the three reads this contract unblocks.

Log in with `https://graph.microsoft.com/Domain.Read.All`, then inspect tenant domains and their DNS records:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Domain.Read.All
mg-axi entra domain list --profile soc
mg-axi entra domain show --profile soc --id contoso.com
mg-axi entra domain verification-dns-record list --profile soc --domain contoso.com
mg-axi entra domain service-configuration-record list --profile soc --domain contoso.com
mg-axi entra domain-dns-record list --profile soc
mg-axi entra domain-dns-record show --profile soc --id <record-id>
```

`entra domain list` defaults to compact properties (`id`, `authenticationType`, `isVerified`, `isDefault`); domain ids are fully qualified names, not object UUIDs.
All eight domain and DNS record commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Domain lists offer no `--filter`: Graph documents a known issue with `$search`, `$top` and `$filter` on domain lists, so the flag is refused before credentials.
`entra domain show --id <domain-name>` defaults to the full reviewed domain set; an unverified domain points at its verification DNS records.
Verification and service-configuration record lists take `--domain <domain-name>` and default to `id`, `label`, `recordType` and `supportedService`.
Their show commands also take `--domain <domain-name>` and require `--id <record-id>`; all DNS record show commands default to `id`, `isOptional`, `label`, `recordType`, `supportedService` and `ttl`.
DNS record lists accept `--filter` as plain `$filter`, without adding `$count=true` or `ConsistencyLevel`.
Record rows carry `@odata.type` naming the derived record kind; derived-type detail (`mailExchange`, `preference`, `canonicalName`, SRV fields, `text`) needs an explicit `--select` naming the derived property.
Domain and record `--select` accept their [reviewed property sets](src/entra-domains.ts); `--fields` must be a subset of the fetched selection.
Domain lists return `domains` and single-domain reads return `domain`; record lists return `verificationDnsRecords`, `serviceConfigurationRecords` and `domainDnsRecords`, with single-record reads returning the singular key.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to domain and DNS record reads.
Resume record lists with the same `--domain` where applicable, profile, scopes and API version; omit or repeat the original server query flags, and repeat local `--fields` and `--full` when wanted.
Delegated domain and DNS record reads default to `https://graph.microsoft.com/Domain.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers additionally need a supported Entra role (Domain Name Administrator or Global Reader are least-privileged); personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for domain reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
No domain mutation lives here; see the [domain scope decisions](docs/coverage.md#ext-01-domain-scope-decisions) for deferred reads and later subfamilies.

Log in with `https://graph.microsoft.com/PublicKeyInfrastructure.Read.All`, then inspect certificate-based-auth PKI configurations and their certificate authorities:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/PublicKeyInfrastructure.Read.All
mg-axi entra certificate-auth-pki list --profile soc
mg-axi entra certificate-auth-pki show --profile soc --id <pki-id>
mg-axi entra certificate-auth-pki count --profile soc
mg-axi entra certificate-authority list --profile soc --pki <pki-id>
mg-axi entra certificate-authority show --profile soc --pki <pki-id> --id <authority-id>
mg-axi entra certificate-authority count --profile soc --pki <pki-id>
```

`entra certificate-auth-pki list` defaults to compact properties (`id`, `displayName`, `status`); an empty list may mean certificate-based authentication is not configured.
All six certificate-auth commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
PKI and authority lists accept `--filter` as plain `$filter`, without adding `$count=true` or `ConsistencyLevel`; counts accept `--filter` to narrow the total server-side.
`entra certificate-auth-pki show --id <pki-id>` defaults to the full reviewed PKI set and points at its authorities; authority lists take `--pki <pki-id>` and default to `id`, `displayName`, `certificateAuthorityType` and `expirationDateTime`.
Authority show commands also take `--pki <pki-id>` and require `--id <authority-id>`; the show default is the full reviewed set except the public-certificate blob.
Authority entries carry public certificates only, but the base64 `certificate` blob (up to 8 KB per CA file) is omitted from every default select and needs an explicit `--select certificate`; explicitly selected blobs still truncate at 500 characters unless `--full` is passed.
PKI lists return `certificateAuthPkis` and single-PKI reads return `certificateAuthPki`; authority lists return `certificateAuthorities` with single-authority reads returning `certificateAuthority`.
Count commands return `count: { returned: <total>, complete: true }` and accept no `--select`, `--fields`, `--limit` or `--cursor`.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to certificate-auth reads.
Delegated certificate-auth reads default to `https://graph.microsoft.com/PublicKeyInfrastructure.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers additionally need Privileged Authentication Administrator or Authentication Administrator; personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for certificate-auth reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
No certificate-auth mutation lives here; the root `/certificateBasedAuthConfiguration` reads carry no documented v1.0 operation contract and the org-scoped certificate-auth reads belong to a later piece; see the [certificate-auth scope decisions](docs/coverage.md#ext-01-certificate-auth-scope-decisions).

Log in with `https://graph.microsoft.com/Organization.Read.All`, then inspect commercial subscriptions:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Organization.Read.All
mg-axi entra subscription list --profile soc
mg-axi entra subscription show --profile soc --id <subscription-id>
mg-axi entra subscription show --profile soc --commerce-subscription-id <commerce-subscription-id>
mg-axi entra subscription count --profile soc
```

`entra subscription list` defaults to compact properties (`id`, `skuPartNumber`, `status`, `totalLicenses`); an empty list may mean the tenant holds no commercial subscriptions.
All three subscription commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Subscription lists accept `--filter` as plain `$filter`, without adding `$count=true` or `ConsistencyLevel`; the show command documents `$select` only and the count command takes no `--filter`, `--select`, `--limit` or `--cursor`.
`entra subscription show --id <subscription-id>` defaults to the full reviewed companySubscription set; `entra subscription show --commerce-subscription-id <commerce-subscription-id>` serves the same object through the commerce-system alternate key, binding it only through the allowlisted session function-argument contract (validated, OData-quoted, encoded).
Subscription lists return `subscriptions` and single-subscription reads return `subscription`; count commands return `count: { returned: <total>, complete: true }`.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to subscription reads.
Delegated subscription reads default to `https://graph.microsoft.com/Organization.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers additionally need Global Reader, Directory Readers, or Dynamics 365 Business Central Administrator for read-only standard properties; personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for subscription reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
No subscription mutation lives here; see the [directory-subscriptions scope decisions](docs/coverage.md#ext-01-directory-subscriptions-scope-decisions).

Log in with `https://graph.microsoft.com/OnPremDirectorySynchronization.Read.All`, then inspect on-premises directory synchronization (delegated profiles only):

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/OnPremDirectorySynchronization.Read.All
mg-axi entra on-premises-synchronization list --profile soc
mg-axi entra on-premises-synchronization show --profile soc --id <synchronization-id>
```

`entra on-premises-synchronization list` and `show` default to the reviewed property set (`id`, `configuration`, `features`); an empty list may mean on-premises directory sync is not configured for the tenant.
Both commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Neither leaf offers `--filter`: Graph documents `$select` only here, so strict input validation refuses the flag before credentials.
Synchronization lists return `synchronizations` and single reads return `synchronization`; secret-shaped values inside `configuration` and `features` stay redacted by the shared session.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to on-premises-synchronization reads.
Delegated reads default to `https://graph.microsoft.com/OnPremDirectorySynchronization.Read.All`; application profiles are refused before credentials because Graph documents no supported application permission for this operation.
Delegated callers additionally need Global Administrator, the only supported Entra role for this operation; personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for on-premises-synchronization reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
No on-premises-synchronization mutation lives here; the `$count` scalar and beta operations stay scheduled; see the [on-premises-synchronization scope decisions](docs/coverage.md#ext-01-on-premises-synchronization-scope-decisions).

Log in with `https://graph.microsoft.com/Agreement.Read.All`, then inspect terms-of-use agreements (delegated profiles only):

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Agreement.Read.All
mg-axi entra agreement list --profile soc
mg-axi entra agreement show --profile soc --id <agreement-id>
```

`entra agreement list` defaults to compact rows (`id`, `displayName`) and `show` defaults to the full reviewed metadata set; an empty list means no terms-of-use agreements are configured for the tenant.
Agreement lists pass `--filter` through as plain `$filter`; agreement shows support `$select` only, so strict input validation refuses `--filter` before credentials.
Log in with `https://graph.microsoft.com/AgreementAcceptance.Read`, then inspect acceptance records (delegated profiles only):

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/AgreementAcceptance.Read
mg-axi entra agreement acceptance list --agreement <agreement-id> --profile soc
mg-axi entra agreement acceptance show --agreement <agreement-id> --id <acceptance-id> --profile soc
mg-axi entra agreement-acceptance list --profile soc
mg-axi entra agreement-acceptance show --id <acceptance-id> --profile soc
```

Acceptance lists return `agreementAcceptances` and single reads return `agreementAcceptance`; default rows stay minimal (`id`, `agreementId`, `state`, `recordedDateTime`) because acceptance records are personal data, and identifying fields need an explicit `--select` naming them.
All six commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to terms-of-use reads.
Delegated agreement reads default to `https://graph.microsoft.com/Agreement.Read.All` and delegated acceptance reads default to `https://graph.microsoft.com/AgreementAcceptance.Read` (`AgreementAcceptance.Read.All` is the documented higher-privileged alternative); application profiles are refused before credentials because Graph documents no supported application permission for these operations.
Delegated callers additionally need Security Reader, the least-privileged supported Entra role for these operations; personal Microsoft accounts are not supported.
Terms of use needs Microsoft Entra ID P1; denied reads name the scope, role and licensing guidance instead of only the generic cause.
Agreement file contents are never downloaded or printed; only agreement metadata is projected.
No terms-of-use mutation lives here; the acceptances `$count` scalar, the agreement file/localization sub-reads and beta operations stay scheduled; see the [terms-of-use scope decisions](docs/coverage.md#ext-01-terms-of-use-scope-decisions).

Log in with `https://graph.microsoft.com/Directory.Read.All`, then inspect directory objects:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Directory.Read.All
mg-axi entra directory-object list --profile soc
mg-axi entra directory-object show --profile soc --id <object-id>
mg-axi entra directory-object count --profile soc
```

`entra directory-object list` defaults to compact rows (`id` plus the `@odata.type` subtype discriminator) and `show` defaults to the reviewed base set (`id`, `deletedDateTime`) plus the discriminator; an empty list is unexpected because every tenant directory carries objects, so verify the profile tenant before treating it as empty.
All three commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
Neither list nor show offers `--filter`: Graph documents no List operation page for the collection, so strict input validation refuses the flag before credentials.
Lists return `directoryObjects`, single reads return `directoryObject`, and counts return `count` with the scalar total.
Rows are polymorphic: only the base-type properties are ever requested or projected (`$expand` is not offered), the discriminator rides along automatically without being selectable, and subtype secrets or credentials can never appear; subtype detail needs the subtype's named reads.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to directory-object reads.
Reads default to `https://graph.microsoft.com/Directory.Read.All` for delegated access, while application profiles use the configured `.default` audience.
No delegated role or P1/P2 prerequisite is stated for directory-object reads; personal Microsoft accounts are not supported, and denied reads name the scope, role and licensing guidance instead of only the generic cause.
No directory-object mutation lives here; the delta sync, the POST lookup/validation actions and beta operations stay scheduled; see the [directory-objects scope decisions](docs/coverage.md#ext-01-directory-objects-scope-decisions).

Log in with the scope matching the deleted type (`User.Read.All`, `Group.Read.All`, `Application.Read.All` or `AdministrativeUnit.Read.All`), then inspect the directory recycle bin:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Group.Read.All
mg-axi entra deleted-user list --profile soc
mg-axi entra deleted-group list --profile soc
mg-axi entra deleted-group count --profile soc
mg-axi entra deleted-item show --profile soc --id <object-id> --scopes https://graph.microsoft.com/Group.Read.All
```

Each deleted-type list defaults to compact rows (`id`, `displayName`, `deletedDateTime`, plus `appId` for applications and service principals) and carries the `@odata.type` discriminator naming the kind; deleted users are personal data, so user defaults stay minimal and identifying fields need an explicit `--select`.
`deleted-item show --id <object-id>` reads the untyped get returning the same object for every type, so it takes no default scope: delegated callers pass the scope matching the object's type as `--scopes` (list the recycle bin first and read each row's `@odata.type` kind).
All eleven deleted-item commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
No typed list offers `--filter`: only `$select` is reviewed here, so strict input validation refuses the flag before credentials.
Type lists return `deletedItems`, single reads return `deletedItem`, and counts return `count` with the scalar total.
Upstream requires the OData cast as part of the list URI, so untyped list/count commands do not exist; device casts carry no documented v1.0 permission contract and stay out with an explicit deferred disposition.
Rows are polymorphic: only the reviewed per-type properties are ever requested or projected, the discriminator rides along automatically without being selectable, and credential collections (`keyCredentials`, `passwordCredentials`) are never selectable, so secret values can never appear.
`appId` (client ID) is distinct from the object `id` on applications and service principals.
Soft-deleted security groups report `securityEnabled` false through a known upstream limitation; read `groupTypes` to name the real kind.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to deleted-item reads.
List and count reads default to the type's least-privileged scope for delegated access (`User.Read.All`, `Group.Read.All`, `Application.Read.All` or `AdministrativeUnit.Read.All`, all read scopes already allowlisted), while application profiles use the configured `.default` audience.
No delegated role or P1/P2 prerequisite is stated for deleted-item reads; personal Microsoft accounts are not supported, and denied reads name the scope, role and licensing guidance instead of only the generic cause.
No deleted-item mutation lives here; restore and permanent delete stay out entirely, and the POST lookup/validation actions and beta operations stay scheduled; see the [deleted-items scope decisions](docs/coverage.md#ext-01-deleted-items-scope-decisions).
Log in with `https://graph.microsoft.com/OrgContact.Read.All`, then inspect organizational contacts:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/OrgContact.Read.All
mg-axi entra contact list --profile soc
mg-axi entra contact show --profile soc --id <contact-id>
mg-axi entra contact count --profile soc
mg-axi entra contact show-manager --profile soc --id <contact-id>
mg-axi entra contact list-direct-reports --profile soc --id <contact-id>
mg-axi entra contact show-direct-report --profile soc --id <contact-id> --report-id <report-id>
mg-axi entra contact count-direct-reports --profile soc --id <contact-id>
mg-axi entra contact list-member-of --profile soc --id <contact-id>
mg-axi entra contact show-member-of --profile soc --id <contact-id> --member-id <membership-id>
mg-axi entra contact count-member-of --profile soc --id <contact-id>
```

Contacts are personal data: `entra contact list` and `show` default to minimal rows (`id`, `displayName`, `mail`, `companyName`), and identifying fields beyond that need an explicit `--select`.
Navigation results are directory objects: `show-manager`, `list-direct-reports`, `show-direct-report`, `list-member-of` and `show-member-of` default to the `@odata.type` discriminator plus `id` and `displayName` only, with `mail` one explicit `--select` away; `--as user|orgContact` selects the typed cast route on direct-report reads and `--as group|administrativeUnit` on member-of reads, `--transitive` reads the transitiveMemberOf closure on member-of reads, and application callers with narrow consent receive limited-information rows carrying only type and id.
All ten contact commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
`entra contact list` and `list-member-of` offer `--filter`, passed through as plain `$filter` with `$count=true` and `ConsistencyLevel eventual`; `$search` and `$orderby` stay unreviewed.
Direct-report reads take `$select` only; `$filter`, `$search` and `$top` stay unreviewed there.
Top-level `list` returns `contacts` and `show` returns `contact`; `show-manager` returns `manager`, `list-direct-reports` returns `directReports` and `show-direct-report` returns `directReport`, `list-member-of` returns `memberOf` and `show-member-of` returns `memberOf`, and counts return `count` with the scalar total; each count sends `ConsistencyLevel eventual` like the documented `$count` example.
Top-level list and show request and project only flat scalar properties (`$expand` is not offered, so navigation objects never appear there); navigation reads return directory objects through their own routes, and the nested phones/addresses collections need their own projection review.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to contact reads.
Reads default to `https://graph.microsoft.com/OrgContact.Read.All` for delegated access, while application profiles use the configured `.default` audience; transitive member-of reads additionally need `https://graph.microsoft.com/Group.Read.All`.
Delegated callers additionally need a supported Entra role (Directory Readers reads basic properties; Global Reader, Directory Writers, Intune Administrator or User Administrator also work); personal Microsoft accounts are not supported, and denied reads name the scope, role and licensing guidance instead of only the generic cause.
No contact mutation lives here; the delta sync, the POST lookup actions and beta operations stay scheduled, and the error/sync navigation reads stay unavailable with no documented permission contract; see the [contacts scope decisions](docs/coverage.md#ext-01-contacts-scope-decisions).

Log in with `https://graph.microsoft.com/IdentityProvider.Read.All`, then read workforce identity providers:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/IdentityProvider.Read.All
mg-axi entra identity-provider list --profile soc
mg-axi entra identity-provider show --profile soc --id <provider-id>
mg-axi entra identity-provider count --profile soc
mg-axi entra identity-provider available-types --profile soc
```

`entra identity-provider list` defaults to compact properties (`id`, `displayName`); `show --id <provider-id>` defaults to the full reviewed provider set.
All four identity-provider commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
List and show accept `--select` from the [reviewed provider property set](src/entra-identity-providers.ts); `--fields` must be a subset of the fetched selection.
Provider lists return `identityProviders`, single-provider reads return `identityProvider`, and counts return `count` with the scalar total.
`available-types` returns `availableProviderTypes`, an array of type names available for the tenant, with a `count` aggregate.
Available types depend on tenant configuration and licensing; availability does not mean a provider is configured.
Provider lists accept `--filter` as plain `$filter`; counts accept `--filter` to narrow the total server-side.
Rows carry `@odata.type` naming the provider kind (social or built-in).
The reviewed workforce fields are `id`, `displayName`, `identityProviderType`, and `clientId`.
`clientSecret` and `certificateData` are never selectable and any row carrying them is scrubbed before output, so key material can never reach stdout, errors or logs.
Provider lists use the named-list caps, uniform `count` totals and cursors described above.
List and show preserve null/missing properties and truncate text longer than 500 characters; `--full` restores complete text without lifting redaction or row caps.
Delegated provider reads default to `https://graph.microsoft.com/IdentityProvider.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers additionally need a directory role that can read federation configuration (Global Reader is the least-privileged read-only directory role); personal Microsoft accounts are not supported.
No per-operation licence prerequisite is stated for these reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
Workforce tenants only; external-customer (B2C/External ID) user flows, cross-tenant access and provisioning are separate later pieces (see the [identity-provider scope decisions](docs/coverage.md#ext-03-identity-provider-scope-decisions)).

Log in with `https://graph.microsoft.com/Domain.Read.All`, then read workforce directory federation configurations:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Domain.Read.All
mg-axi entra federation-configuration list --profile soc
mg-axi entra federation-configuration show --profile soc --id <configuration-id>
mg-axi entra federation-configuration count --profile soc
mg-axi entra federation-configuration available-types --profile soc
```

`entra federation-configuration list` defaults to compact properties (`id`, `displayName`); `show --id <configuration-id>` defaults to the full reviewed set except the signing-certificate blob, which needs an explicit `--select signingCertificate`.
All four federation-configuration commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
List and show accept `--select` from the [reviewed configuration property set](src/entra-federation-configurations.ts); `--fields` must be a subset of the fetched selection.
Configuration lists return `federationConfigurations`, single-configuration reads return `federationConfiguration`, and counts return `count` with the scalar total.
`available-types` returns `availableProviderTypes`, an array of type names available for the tenant, with a `count` aggregate; it defaults to `https://graph.microsoft.com/IdentityProvider.Read.All` while list, show and count default to `https://graph.microsoft.com/Domain.Read.All`, and application profiles use the configured `.default` audience.
`signingCertificate` carries the public token-signing key only and is omitted from every default select; request it explicitly when rotation evidence is needed. No private key material exists on these resources.
Configuration lists accept `--filter` as plain `$filter`; counts accept `--filter` to narrow the total server-side.
Rows carry `@odata.type` naming the configuration kind.
Configuration lists use the named-list caps, uniform `count` totals and cursors described above.
List and show preserve null/missing properties and truncate text longer than 500 characters; `--full` restores complete text without lifting row caps.
Delegated callers additionally need External Identity Provider Administrator, the least-privileged supported Entra role for these reads; personal Microsoft accounts are not supported.
No per-operation licence prerequisite is stated for these reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
No domain federationConfiguration sub-read, no domains navigation expansion and no mutation live here; beta operations stay scheduled (see the [federation-configuration scope decisions](docs/coverage.md#ext-01-federation-configuration-scope-decisions)).

Log in with `https://graph.microsoft.com/User.Export.All,https://graph.microsoft.com/User.Read.All`, then read workforce data-policy operations:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/User.Export.All,https://graph.microsoft.com/User.Read.All
mg-axi entra data-policy-operation list --profile soc
mg-axi entra data-policy-operation show --profile soc --id <operation-id>
mg-axi entra data-policy-operation count --profile soc
```

`entra data-policy-operation list` defaults to compact properties (`id`, `status`, `userId`, `submittedDateTime`); `show --id <operation-id>` defaults to the full reviewed operation set.
All three data-policy-operation commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
List and show accept `--select` from the [reviewed operation property set](src/entra-data-policy-operations.ts); `--fields` must be a subset of the fetched selection.
The list documents `$select` only, so no `--filter`; the `$count` route takes no `--filter`, `--select`, `--limit` or `--cursor`.
Operation lists return `dataPolicyOperations`, single-operation reads return `dataPolicyOperation`, and counts return `count` with the scalar total.
`storageLocation` always renders as `***redacted***`: export blob URLs and signed links never reach output, errors or logs.
Operation lists use the named-list caps, uniform `count` totals and cursors described above.
List and show preserve null/missing properties and truncate text longer than 500 characters; `--full` restores complete text without lifting redaction or row caps.
Delegated operation reads default to `https://graph.microsoft.com/User.Export.All,https://graph.microsoft.com/User.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers additionally need Company Administrator, the privileged role documented for export reads; personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for these reads; denied reads name the scopes, role and licensing guidance instead of only the generic cause.
No export submission lives here; the `$count` scalar aside, beta operations stay scheduled (see the [data-policy-operations scope decisions](docs/coverage.md#ext-03-data-policy-operations-scope-decisions)).

Log in with `https://graph.microsoft.com/RiskPreventionProviders.Read.All`, then read workforce risk-prevention providers and verifications (delegated profiles only):

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/RiskPreventionProviders.Read.All
mg-axi entra fraud-protection-provider list --profile soc
mg-axi entra fraud-protection-provider show --profile soc --id <provider-id>
mg-axi entra fraud-protection-provider count --profile soc
mg-axi entra web-application-firewall-provider list --profile soc
mg-axi entra web-application-firewall-provider show --profile soc --id <provider-id>
mg-axi entra web-application-firewall-provider count --profile soc
mg-axi entra web-application-firewall-verification list --profile soc
mg-axi entra web-application-firewall-verification show --profile soc --id <verification-id>
mg-axi entra web-application-firewall-verification count --profile soc
mg-axi entra web-application-firewall-verification provider show --profile soc --id <verification-id>
```

`entra fraud-protection-provider list` and `entra web-application-firewall-provider list` default to compact properties (`id`, `displayName`); `entra web-application-firewall-verification list` defaults to (`id`, `verifiedHost`, `providerType`); every `show` defaults to its full reviewed set.
All ten risk-prevention commands support only `--api-version v1.0`; `--api-version beta` fails validation before credentials, including on preview-enabled profiles.
List and show accept `--select` from the [reviewed risk-prevention property sets](src/entra-risk-prevention.ts); `--fields` must be a subset of the fetched selection.
The fraud list documents `$select` only, so no `--filter`; the WAF and verification lists accept `--filter` as plain `$filter`; every `$count` route takes no `--filter`, `--select`, `--limit` or `--cursor`.
Fraud lists return `fraudProtectionProviders`, WAF lists return `webApplicationFirewallProviders`, verification lists return `webApplicationFirewallVerifications`; single reads return `fraudProtectionProvider`, `webApplicationFirewallProvider` or `webApplicationFirewallVerification`; counts return `count` with the scalar total.
Verification rows never carry provider detail; `provider show --id <verification-id>` returns the owning `webApplicationFirewallProvider`.
Fraud and WAF rows carry `@odata.type` naming the provider kind (Arkose or HUMAN; Akamai or Cloudflare).
The reviewed fraud fields are `id`, `displayName`, `clientSubDomain`, `verifySubDomain` and `appId`; the reviewed WAF fields are `id`, `displayName`, `hostPrefix` and `zoneId`; the reviewed verification fields are `id`, `verifiedHost`, `providerType`, `verificationResult` and `verifiedDetails`.
Arkose `privateKey`/`publicKey`, HUMAN `serverToken`, Akamai `clientSecret`/`clientToken`/`accessToken` and Cloudflare `apiToken` are never selectable and any row carrying them is scrubbed before output, so key material can never reach stdout, errors or logs.
Risk-prevention lists use the named-list caps, uniform `count` totals and cursors described above.
List and show preserve null/missing properties and truncate text longer than 500 characters; `--full` restores complete text without lifting redaction or row caps.
Delegated risk-prevention reads default to `https://graph.microsoft.com/RiskPreventionProviders.Read.All`; Graph documents no supported application permission, so application profiles are refused before credentials.
Delegated callers additionally need a directory role that can read risk-prevention configuration (Security Reader is the least-privileged read-only role); personal Microsoft accounts are not supported.
No per-operation licence prerequisite is stated for these reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
No provider create/update/delete and no verification submission live here; beta operations stay scheduled (see the [risk-prevention scope decisions](docs/coverage.md#ext-03-risk-prevention-scope-decisions)).

Log in with `https://graph.microsoft.com/Policy.Read.All`, then read Conditional Access policies, named locations and templates as separate grammar:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Policy.Read.All
mg-axi entra conditional-access policy list --profile soc --limit 10
mg-axi entra conditional-access policy list --profile soc --filter "state eq 'enabled'"
mg-axi entra conditional-access policy show --profile soc --id <policy-id>
mg-axi entra conditional-access named-location list --profile soc
mg-axi entra conditional-access named-location show --profile soc --id <named-location-id>
mg-axi entra conditional-access template list --profile soc --limit 10
mg-axi entra conditional-access template show --profile soc --id <template-id>
mg-axi entra conditional-access deleted-policy list --profile soc --limit 10
mg-axi entra conditional-access deleted-policy show --profile soc --id <deleted-policy-id>
mg-axi entra conditional-access deleted-named-location list --profile soc
mg-axi entra conditional-access deleted-named-location show --profile soc --id <deleted-named-location-id>
```

Log in with `https://graph.microsoft.com/AuthenticationContext.Read.All`, then read authentication context class references through the same session:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/AuthenticationContext.Read.All
mg-axi entra conditional-access auth-context list --profile soc --limit 10
mg-axi entra conditional-access auth-context show --profile soc --id <auth-context-id>
```

Log in with `https://graph.microsoft.com/Policy.Read.AuthenticationMethod`, then read the authentication-strength family through the same session:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Policy.Read.AuthenticationMethod
mg-axi entra conditional-access auth-strength-policy list --profile soc --limit 10
mg-axi entra conditional-access auth-strength-policy show --profile soc --id <auth-strength-policy-id>
mg-axi entra conditional-access combination-configuration list --policy <auth-strength-policy-id> --profile soc
mg-axi entra conditional-access combination-configuration show --policy <auth-strength-policy-id> --id <combination-configuration-id> --profile soc
mg-axi entra conditional-access auth-method-mode list --profile soc --limit 10
mg-axi entra conditional-access auth-method-mode show --profile soc --id <auth-method-mode-id>
```

`entra conditional-access policy list` defaults to compact properties (`id`, `displayName`, `state`); `policy show --id <policy-id>` defaults to the full reviewed condition and control property set.
`entra conditional-access named-location list` defaults to `id` and `displayName`; `named-location show --id <named-location-id>` defaults to the full reviewed location set.
`entra conditional-access auth-strength-policy list` defaults to compact properties (`id`, `displayName`, `policyType`); `auth-strength-policy show --id <auth-strength-policy-id>` defaults to the full reviewed strength set.
`entra conditional-access combination-configuration list --policy <auth-strength-policy-id>` defaults to `id`; `combination-configuration show` takes `--policy` plus `--id` and defaults to the full reviewed base set.
`entra conditional-access auth-method-mode list` defaults to `id` and `displayName`; `auth-method-mode show --id <auth-method-mode-id>` defaults to the full reviewed mode set.
`entra conditional-access template list` defaults to `id` and `name`; `template show --id <template-id>` defaults to the full reviewed template set.
`entra conditional-access auth-context list` defaults to `id` and `displayName`; `auth-context show --id <auth-context-id>` defaults to the full reviewed authentication-context set.
`entra conditional-access deleted-policy list` defaults to compact properties (`id`, `displayName`, `state`, `deletedDateTime`); `deleted-policy show --id <deleted-policy-id>` defaults to the full reviewed deleted-policy set. Deleted-policy reads never restore; restore is a beta-only mutation outside mg-axi scope.
`entra conditional-access deleted-named-location list` defaults to `id`, `displayName` and `deletedDateTime`; `deleted-named-location show --id <deleted-named-location-id>` defaults to the full reviewed deleted-location set. Deleted named-location reads never restore; restore is a beta-only mutation outside mg-axi scope.
Policy `--select` accepts the [reviewed policy property set](src/entra-conditional-access.ts) and location `--select` accepts the [reviewed location property set](src/entra-conditional-access.ts); strength, combination-configuration, method-mode, template, authentication-context, deleted-policy and deleted-named-location `--select` accept their own [reviewed property sets](src/entra-conditional-access.ts); `--fields` must be a subset of the fetched selection in each family.
Returned `@odata.type` stays visible on named-location rows so IP and country locations stay distinguishable alongside any `--fields` projection.
Returned `@odata.type` also stays visible on combination-configuration rows so configuration kinds stay distinguishable alongside any `--fields` projection.
Returned `@odata.type` also stays visible on deleted-named-location rows so IP and country locations stay distinguishable alongside any `--fields` projection.
Policy lists return `policies`, location lists return `namedLocations`, single-policy reads return `policy` and single-location reads return `namedLocation`.
Strength lists return `authStrengthPolicies`, combination lists return `combinationConfigurations`, mode lists return `authMethodModes`, template lists return `templates`, authentication-context lists return `authContexts`, deleted-policy lists return `deletedPolicies` and deleted-named-location lists return `deletedNamedLocations`; single reads return `authStrengthPolicy`, `combinationConfiguration`, `authMethodMode`, `template`, `authContext`, `deletedPolicy` and `deletedNamedLocation`.
Successful empty collections return an empty list with a uniform `count` line (`0 policies` or `0 named locations`, `total: null`) and absence guidance.
The authentication-strength, combination-configuration, authentication-method-mode, template, authentication-context, deleted-policy and deleted-named-location lists report uniform totals through the shared list-totals helper beside a `count` line such as `2 auth-strength policies` or `2 templates` (`0 ...` with absence guidance on empty results).
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply to Conditional Access reads, including nested condition values; every truncated value carries a `--full` hint.
Resume any Conditional Access list with `--cursor -` and supply the returned cursor on stdin, for example `mg-axi entra conditional-access policy list --profile soc --cursor - < cursor.txt`.
Resume combination-configuration lists with the same `--policy`, profile, scopes and API version.
Small cursors can also use `--cursor <token>`; both forms enforce a 16 MB size ceiling.
Use `--full` when reasoning from complete condition or control text; `--select` and `--fields` still determine which properties are visible, and missing properties remain unknown.
To replay a resumed result with `--full`, reuse the original input cursor; the returned cursor continues after that result.
Delegated policy, named-location, template, deleted-policy and deleted-named-location reads default to `https://graph.microsoft.com/Policy.Read.All`; delegated authentication-strength-policy, combination-configuration and authentication-method-mode reads default to `https://graph.microsoft.com/Policy.Read.AuthenticationMethod`; delegated authentication-context reads default to `https://graph.microsoft.com/AuthenticationContext.Read.All`; application profiles require the admin-consented scope for the family, use the configured Graph `.default` audience and reject `--scopes`.
HTTP 403 policy, location, template, deleted-policy and deleted-named-location errors name the operation's supported directory roles (Conditional Access Administrator, Global Reader, Global Secure Access Administrator, Security Administrator or Security Reader for delegated access); HTTP 403 authentication-context errors name the same broad role set with `AuthenticationContext.Read.All`; HTTP 403 authentication-strength errors name the narrower role set (Conditional Access Administrator, Security Administrator or Security Reader) with `Policy.Read.AuthenticationMethod`; all include guidance from the [licensing contract](docs/graph-coverage.md#licence-matrix-by-area), and Conditional Access reads need P1; the response alone does not identify the missing prerequisite.
WRITE-04 adds a named write: `mg-axi entra conditional-access policy update --id <policy-id>` PATCHes one policy with only the reviewed fields `--display-name`, `--state`, `--conditions`, `--grant-controls` and `--session-controls` (JSON objects for the last three); any other property has no flag and is refused before credentials.
At least one reviewed field flag is required.
The command supports only `--api-version v1.0`; `--state` accepts `enabled`, `enabledForReportingButNotEnforced` or `disabled`, and each JSON flag supplies the full corresponding block.
Without `--execute` the command previews the redacted current-versus-proposed diff read through the policy show route and journals nothing; an already-desired value set is a no-op with exit 0 only when all requested values can be verified equal, with redacted current values treated as unknown.
Every `--execute` run is disruptive and needs `--confirm '<policy-id>'` repeating the target exactly.
The preview carries a lockout analysis over the proposed effective policy; changes to `state`, `conditions`, `grantControls` or `sessionControls` enforce its gates before sending and on fresh reads.
An enabled all-users policy under block controls is refused outright, even with `--acknowledge-lockout-risk`, when it covers browser and modern clients on all cloud apps or `MicrosoftAdminPortals` without exclusions of admin portals or other statically decidable narrowing conditions.
Application filters cannot be evaluated statically and never narrow coverage; user, group and role exclusions also cannot establish emergency-access protection without verified principal coverage, so they do not exempt an otherwise full lockout from refusal.
Unrelated application exclusions preserve administrative coverage; universal platform (`all`) and location (`All`) includes with empty exclusions also preserve full coverage and hard refusal.
Only explicitly recognized conditions reduce coverage; protocol annotations, unknown properties and null values do not narrow it.
Non-empty `signInRiskLevels` or `userRiskLevels` containing only `low`, `medium` or `high` narrow coverage to risky sign-ins or users.
Policies whose client types do not cover both browser and modern clients, whose application inclusions cover neither `All` nor `MicrosoftAdminPortals`, or whose application exclusions contain `All` or `MicrosoftAdminPortals` show a scoped warning and require acknowledgement rather than being classified as full lockouts.
Non-universal platform or location includes, non-empty platform or location exclusions, and the risk-level conditions described above also produce scoped warnings.
Every other enabled policy needs `--acknowledge-lockout-risk` after confirming emergency-access coverage: explicit user, group or role targets and exclusions alone cannot establish that admins and break-glass accounts are protected.
Enforcement changes stay disabled when the effective state, conditions, user scope, exclusions, client types, application inclusions or exclusions, platform or location inclusions or exclusions, or grant controls are unreadable; omitted `excludeUsers`, `excludeGroups`, `clientAppTypes`, `applications.includeApplications` or `applications.excludeApplications` makes the analysis unavailable.
Disabled and report-only effective policies do not enforce access; display-name-only changes need no lockout acknowledgement or available analysis.
Graph documents no ETag or If-Match precondition for this endpoint, so the command sends none and promises no concurrency protection.
After execution, verify the policy with `mg-axi entra conditional-access policy show --id <policy-id> --profile soc`; the successful PATCH returns no policy body, and an unknown outcome requires readback before further action, never a blind replay.
The mutation needs D/A `Policy.Read.All` plus `Policy.ReadWrite.ConditionalAccess`; delegated callers additionally need Conditional Access Administrator or Security Administrator, and CA needs P1 (P2 for risk-based features).

Inspect one user's authentication methods and the tenant registration report; delegated profiles first need explicit login with the read scopes:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/UserAuthenticationMethod.Read.All,https://graph.microsoft.com/AuditLog.Read.All
mg-axi entra user authentication-method list --profile soc --user <user-id>
mg-axi entra user authentication-method list --profile soc --user <user-id> --select id,displayName,phoneType
mg-axi entra registration list --profile soc
mg-axi entra registration list --profile soc --filter "isMfaRegistered eq false"
```

`entra user authentication-method list` targets one named user (`--user` takes the user object ID or UPN) and defaults to `id`, `displayName` and `createdDateTime`; rows carry `@odata.type` naming the method kind.
There is no tenant scan through per-user methods: aggregate MFA coverage belongs to `entra registration list`, and the method output points there.
`entra registration list` defaults to `id`, `userPrincipalName`, `userDisplayName` and `isMfaRegistered` and returns the tenant MFA/SSPR posture.
Method lists return `authenticationMethods`; registration reports return `registrationDetails`.
The report does not cover disabled users, so absence from it is never proof of no MFA; that gap rides in the leaf help and every report output.
Phone numbers are protected values: the shared session replaces them with the redaction marker before output or cursor buffering (including resume cursors), so neither output nor cursors ever carry one; `--full` never lifts that redaction.
Method registration and deletion belong to no read slice and are never constructed.
For methods, `--select` selects output properties locally from the [reviewed authentication property sets](src/entra-auth-methods.ts); no `$select` is sent to Graph.
For the registration report, `--select` requests server properties; `--fields` must be a subset of the default or explicit selection for either command.
Only the registration report supports `--filter`, passed through as plain `$filter` with no `$count` or `ConsistencyLevel` contract.
The named-list caps, uniform `count` totals, cursors, null/missing preservation and 500-character text truncation described above also apply.
Resume method lists with the same `--user` ID or UPN, profile, scopes and API version; omit or repeat the original `--select`, and repeat local `--fields` and `--full` when wanted.
Delegated method reads default to `https://graph.microsoft.com/UserAuthenticationMethod.Read.All` (delegated self-reads may use `UserAuthenticationMethod.Read`) and registration reads default to `https://graph.microsoft.com/AuditLog.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers acting on another user additionally need Global Reader, Authentication Administrator or Privileged Authentication Administrator for methods, and Reports Reader, Security Reader, Security Administrator or Global Reader for the report.
Denied reads name that role requirement instead of only the generic grant/role/licence cause.

Log in with `https://graph.microsoft.com/AuditLog.Read.All`, then query sign-ins, directory audits and provisioning logs in bounded time windows (provisioning needs the extra scope below):

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/AuditLog.Read.All
mg-axi entra sign-in list --profile soc --since 2026-09-01T00:00:00Z --limit 10
mg-axi entra sign-in list --profile soc --since 2026-09-01T00:00:00Z --filter "status/errorCode ne 0" --all
mg-axi entra sign-in show --profile soc --id <sign-in-id>
mg-axi entra directory-audit list --profile soc --since 2026-09-01T00:00:00Z
mg-axi entra directory-audit show --profile soc --id <directory-audit-id>
mg-axi entra provisioning list --profile soc --since 2026-09-01T00:00:00Z
mg-axi entra provisioning show --profile soc --id <provisioning-id>
```

Provisioning logs additionally need `https://graph.microsoft.com/Directory.Read.All` alongside `AuditLog.Read.All`, plus a Microsoft Entra ID P1 or P2 licence on the tenant:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/AuditLog.Read.All,https://graph.microsoft.com/Directory.Read.All
```

Log lists always carry an explicit time bound: `--since` is required for a new query (with optional `--until` and `--filter` refinements), and resume reuses `--cursor` instead.
Resume validates the saved time bounds; a cursor from an unbounded raw query is rejected, so start a new query with `--since`.
Resume sign-in, directory-audit and provisioning lists with `--cursor -` and supply the returned cursor on stdin, for example `mg-axi entra sign-in list --profile soc --cursor - < cursor.txt`.
Small cursors can also use `--cursor <token>`; both forms enforce a 16 MB size ceiling.
`entra sign-in list` defaults to `id`, `createdDateTime`, `userPrincipalName` and `appDisplayName`; `entra directory-audit list` defaults to `id`, `activityDateTime`, `activityDisplayName` and `result`; `entra provisioning list` defaults to `id`, `activityDateTime`, `action` and `provisioningStatusInfo`.
All three log show commands default to the full reviewed property set.
`--select` requests properties from the [supported log property sets](src/entra-audit-logs.ts); `--fields` projects locally and must be a subset of the fetched selection.
Log reads truncate text longer than 500 characters, including nested values; `--full` restores complete text without lifting redaction, row caps or time bounds.
To replay a resumed result with `--full`, supply the original input cursor on stdin; the returned cursor continues after that result.
Graph omits CA policy detail without CA-data access, so an absent `appliedConditionalAccessPolicies` value reports its required policy permission and delegated role as unavailable rather than empty.
For CA detail, use Policy.Read.All or Policy.Read.ConditionalAccess in addition to AuditLog.Read.All; delegated callers also need Conditional Access Administrator, Global Reader, Security Administrator or Security Reader.
For delegated CA detail, log in and repeat the read with `--scopes https://graph.microsoft.com/AuditLog.Read.All,https://graph.microsoft.com/Policy.Read.All`.
Denied log reads name the operation's supported directory roles and the conservative P1/P2 deployment prerequisite instead of only the generic grant/role/licence cause.

`entra application list` and `entra service-principal list` default to compact properties (`id`, `appId`, `displayName`); `appId` is the client ID, distinct from the object `id`, and both are included by default, while custom `--select` or `--fields` can omit either.
`entra application show --id <application-object-id>` defaults to the richer reviewed application set including `keyCredentials` and `passwordCredentials`.
`entra service-principal show --id <service-principal-object-id>` defaults to the richer reviewed service-principal set including `servicePrincipalType` and both credential collections.
Credential collections expose only expiry metadata (`keyId`, `displayName`, `startDateTime`, `endDateTime`); the shared session drops every other credential subfield before output or cursor buffering and applies the same filtering to buffered rows on resume, including older cursors.
Non-array credential collections become empty arrays, and non-object entries are dropped; other properties retain the null/missing behavior described above.
Secret-minting routes are never constructed.
`entra application owner list --application <application-object-id>` and `entra service-principal owner list --service-principal <service-principal-object-id>` list owners; rows carry `@odata.type` naming the owner kind.
For consent grants, use the service-principal grant commands described above.
Application lists return `applications`, service-principal lists return `servicePrincipals`, owner lists return `owners`, and single-object reads return `application` or `servicePrincipal`.
App and service-principal `--select` accepts the [reviewed property sets](src/entra-apps.ts); `--fields` must be a subset of the fetched selection.
Owner rows default to `id`, `displayName` and `mail`, the only selectable owner properties; rows without non-null descriptive properties are preserved with a hint about limited consent or unset properties.
The named-list caps, uniform `count` totals, cursor resume rules and 500-character text truncation described above also apply to these reads; resume owner lists with the same `--application` or `--service-principal` object ID.
Delegated application reads default to `https://graph.microsoft.com/Application.Read.All`, while application profiles use the configured `.default` audience.

Log in with `https://graph.microsoft.com/IdentityRiskyUser.Read.All`, `https://graph.microsoft.com/IdentityRiskEvent.Read.All` and `https://graph.microsoft.com/IdentityRiskyServicePrincipal.Read.All`, then triage risky users, risk detections, risky service principals and service-principal risk detections:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/IdentityRiskyUser.Read.All,https://graph.microsoft.com/IdentityRiskEvent.Read.All,https://graph.microsoft.com/IdentityRiskyServicePrincipal.Read.All
mg-axi entra risky-user list --profile soc --limit 10
mg-axi entra risky-user show --profile soc --id <risky-user-id>
mg-axi entra risk-detection list --profile soc --since 2026-09-01T00:00:00Z --limit 10
mg-axi entra risk-detection show --profile soc --id <risk-detection-id>
mg-axi entra risky-service-principal list --profile soc --limit 10
mg-axi entra risky-service-principal show --profile soc --id <risky-service-principal-id>
mg-axi entra risky-service-principal history list --profile soc --service-principal <risky-service-principal-id>
mg-axi entra risky-service-principal history show --profile soc --service-principal <risky-service-principal-id> --id <history-item-id>
mg-axi entra service-principal-risk-detection list --profile soc --since 2026-09-01T00:00:00Z --limit 10
mg-axi entra service-principal-risk-detection show --profile soc --id <service-principal-risk-detection-id>
```

Delegated risky-user reads default to `https://graph.microsoft.com/IdentityRiskyUser.Read.All`; user and service-principal risk-detection reads default to `https://graph.microsoft.com/IdentityRiskEvent.Read.All`; risky-service-principal and history reads default to `https://graph.microsoft.com/IdentityRiskyServicePrincipal.Read.All`.
`--scopes` overrides those defaults; application profiles use the configured Graph `.default` audience and reject delegated scopes.
`entra risky-user list` and `entra risky-service-principal list` are state collections with an optional `--filter`; `entra risk-detection list` and `entra service-principal-risk-detection list` always carry an explicit time bound, so `--since` is required for a new query (with optional `--until` and `--filter` refinements) and resume reuses `--cursor` instead.
`entra risky-service-principal history list` and `history show` always carry `--service-principal` naming the owning risky-service-principal object ID.
Resume validates the saved detectedDateTime bounds; a cursor from an unbounded raw query is rejected, so start a new query with `--since`.
Resume risk lists with `--cursor -` and supply the returned cursor on stdin, for example `mg-axi entra risk-detection list --profile soc --cursor - < cursor.txt`.
`entra risky-user list` defaults to `id`, `userPrincipalName`, `riskLevel` and `riskState`; `entra risk-detection list` defaults to `id`, `detectedDateTime`, `userPrincipalName` and `riskLevel`.
`entra risky-service-principal list` and `history list` default to `id`, `displayName`, `riskLevel` and `riskState`; `entra service-principal-risk-detection list` defaults to `id`, `detectedDateTime`, `servicePrincipalDisplayName` and `riskLevel`.
All risk show commands default to the full reviewed property set.
`--select` requests properties from the [supported risk property sets](src/entra-risk.ts); `--fields` projects locally and must be a subset of the fetched selection.
Risk reads truncate text longer than 500 characters, including nested values such as `location` and the `additionalInfo` JSON string; `--full` restores complete text without lifting redaction, row caps or time bounds.
See the [access and licence contract](docs/graph-coverage.md#licence-matrix-by-area) for the riskyUsers P2 requirement, risk-detection P1/P2 detail boundaries and the Workload Identities Premium requirement behind the workload-identity reads.
Limited views stay limited: a premium detection without P2 detail reports `riskEventType` generic, hidden risk levels report the licence boundary instead of the level, service-principal detections report `riskDetail` and `riskLevel` hidden without Workload Identities Premium, and a null detection `correlationId` means no sign-in is associated.
To correlate a detection to sign-ins, use `risk-detection show` or select `activityDateTime`, then filter the sign-in list above on the detection's `userPrincipalName` in that activity window; sign-in reads require the separate `AuditLog.Read.All` login above, and there is no riskySignIns endpoint.
To correlate a workload detection, filter `service-principal-risk-detection list` on the principal's `servicePrincipalId` in the detection window, or list that principal's risk history with `risky-service-principal history list --service-principal <risky-service-principal-id>`; workload reads never confirm, dismiss or remediate risk.
WRITE-05 adds a named write: `mg-axi entra risky-user dismiss --user <risky-user-id>` dismisses one user's risk through `POST /identityProtection/riskyUsers/dismiss` with a single-element `{ "userIds": [...] }` body.
There is no bulk form: `--user` takes exactly one ID and a comma-separated list is a usage error.
The command supports only `--api-version v1.0`; beta writes are rejected before credentials or HTTP.
Without `--execute` the command previews the user's current risk state read through the risky-user show route and journals nothing; an already-dismissed user is a no-op with exit 0.
Preview also requires the write-enabled profile and operation allowlist described below.
Dismissal is disruptive: every `--execute` run needs `--confirm '<risky-user-id>'` repeating the target exactly, including already-dismissed states.
Dismissal answers 204 with an empty body, so after a send the command rereads the user and reports a `WRITE_CONFLICT` when the state is not dismissed; a failed verification read reports `OUTCOME_UNKNOWN`.
A timeout or 5xx after send is `OUTCOME_UNKNOWN` with no replay: read back the user before doing anything else.
Fresh reads do not make dismissal atomic: no ETag condition is sent, so another actor can change the risk state between the read, POST and verification.
Dismissal is not remediation: it records the risk as dismissed without resetting credentials or revoking sessions.
Delegated dismissal requests `IdentityRiskyUser.ReadWrite.All` while preview, pre-send and verification reads request `IdentityRiskyUser.Read.All`; credentials are acquired silently.
Use explicit `mg-axi login --profile soc --scopes https://graph.microsoft.com/IdentityRiskyUser.ReadWrite.All,https://graph.microsoft.com/IdentityRiskyUser.Read.All` to sign in for the write and its prerequisite reads.
Application profiles use the configured Graph `.default` audience with `IdentityRiskyUser.ReadWrite.All` admin-consented on the app registration; the prerequisite reads also require a [supported risky-user read permission](https://learn.microsoft.com/en-us/graph/api/riskyuser-get?view=graph-rest-1.0).
The dismissal command rejects caller-supplied `--scopes`.
Delegated callers additionally need `Security Administrator`, and the riskyUsers API requires a Microsoft Entra ID P2 licence; 403 denials surface the permission, role and licence rules because a 403 never says which prerequisite is missing.
Risk reads never confirm, dismiss or remediate risk.

Configure the registration's Mobile and desktop applications redirect URI as `http://localhost` for browser login.
The first created profile is the default; `--profile` selects another identity explicitly.
Configuration defaults to `~/.mg-axi/config.json`; `MG_AXI_CONFIG` selects a separate configuration file.
Version 1 stores tenant/client/cloud, delegated or application mode, enabled packs, preview/sensitive policy, device-code opt-in, a unique credential reference and an optional write policy.
Unknown fields, unsupported versions/clouds and inlined credential material are rejected with recovery guidance.
Creation never overwrites an existing profile.
Preview is disabled and sensitive areas are empty in newly created profiles.

Writes stay disabled unless a human hand-edits a `writes` object into the profile file: `{ "allowWrites": true, "operations": ["<operation-name>"] }`.
The object accepts only `allowWrites` (boolean) and `operations` (1 to 64 nonempty operation names, each at most 256 characters), including when `allowWrites` is false.
No command writes that object, and `MG_AXI_READ_ONLY=1` overrides any opt-in.
For the supported membership write, see the group usage above.
Named writes run through the shared coordinator under the [named-write execution contract](docs/execution.md#named-writes).
WRITE-02 binds the `entra.user.update` operation name: hand-enable account writes with `{ "allowWrites": true, "operations": ["entra.user.update"] }`.
WRITE-03 binds the `entra.user.revokeSessions` operation name: hand-enable session revocation with `{ "allowWrites": true, "operations": ["entra.user.revokeSessions"] }`.
WRITE-04 binds the `entra.conditional-access.policy.update` operation name: hand-enable policy writes with `{ "allowWrites": true, "operations": ["entra.conditional-access.policy.update"] }`.
WRITE-05 binds the `entra.risky-user.dismiss` operation name: hand-enable risk dismissals with `{ "allowWrites": true, "operations": ["entra.risky-user.dismiss"] }`.
The journal defaults to `~/.mg-axi/writes.log`; a nonblank `MG_AXI_WRITE_LOG` overrides that path.

Browser login uses Microsoft's [MSAL interactive API](https://learn.microsoft.com/en-us/entra/msal/javascript/node/acquire-token-requests) and PKCE.
Request delegated Graph scope names explicitly with `--scopes`, separated by commas.
Use full `https://graph.microsoft.com/` scope names; delegated login rejects `.default`.
The dedicated registration needs corresponding delegated consent, and operations can additionally require user roles.
Login never changes app registration or requests blanket directory write consent automatically.
Device code requires creating the profile with `--allow-device-code` and selecting `--method device-code` during login, only when organization policy permits it.
Browser failure never falls back to device code.
Only explicit login can open a browser or display the device challenge on stderr.
Ordinary credential acquisition uses silent refresh and returns actionable errors when login, consent or policy intervention is required.

Read reviewed raw Graph data without waiting for a named command:

```sh
mg-axi api get /users --scopes https://graph.microsoft.com/User.Read.All
mg-axi api get /groups --odata '$filter=securityEnabled eq true&$top=5' --scopes https://graph.microsoft.com/GroupMember.Read.All
mg-axi api get /identity/conditionalAccess/policies --scopes https://graph.microsoft.com/Policy.Read.All
```

`api get` accepts only GET routes in the [reviewed route catalogue](src/api.ts), which owns route-specific query keys, `$select` fields and access constraints.
OData parameters use `--odata`; `--query` is reserved for output queries and is not implemented here.
Omitting `$select` selects the route's default reviewed fields, and every response is filtered to the selected subset before output.
Raw PKI and certificate-authority list/show reads use the certificate-auth paths in the [reviewed route catalogue](src/api.ts); their `$count` routes are available only through the named count commands above.
Raw certificate-authority defaults omit `certificate`; request it explicitly with `--odata '$select=certificate'`, subject to the raw text truncation described below.
For `/users/<user-id>/authentication/methods`, `$select` is the only supported OData parameter and selects output properties locally; an explicit selection restricts output to that subset.
Method requests omit `$select` on initial requests, continuations, redirects and retries, while cursors preserve the local selection.
Other routes send `$select` to Graph.
Relationship expansion (`$expand`) is unavailable.
Unreviewed, secret-value, mail/file-content, beta and write routes fail before credentials, and pack, preview and sensitive-area policy still runs in the shared session.
Delegated raw reads require explicit `--scopes` from the supported read choices described above; application profiles use the configured `.default` audience and reject `--scopes`.
Collections return `returned`, `complete` and `value`, default to 100 rows, and follow pages within budget under `--all`.
`--limit` and `--all` cannot be combined.
Completion describes pagination, not visibility of every directory object; group-member results include a warning for the [v1.0 service-principal limitation](docs/graph-coverage.md#licensing-and-completeness-findings), even when `complete` is true.
Strings longer than 500 characters are truncated; `--full` removes string truncation without disabling redaction, reviewed-field filtering or row caps.
Partial results include a `cursor` preserving buffered rows and the next page.
Resume with `--cursor -` and supply the cursor token on stdin under the same collection path, profile and scopes, optionally with `--all` or a new `--limit`.
For example, `mg-axi api get /users --cursor - --all --scopes https://graph.microsoft.com/User.Read.All < cursor.txt` reads a saved token through stdin.
Small tokens can also use `--cursor <token>`; both input forms use the same validation and a 16 MB size ceiling.
Omit `--odata` on resume to reuse the original query; supplying a different query is refused.

Delegated MSAL caches use the OS credential store through optional `keytar`, with login reporting `storage: os-protected`.
If `keytar` cannot load, caches remain in process memory and login reports `storage: session-only`; authentication then lasts only for that process.
Installing with `--ignore-scripts` can leave the native module unavailable.
Session-only mode uses an MSAL client without a persistence plugin and is available only when no usable protected store exists.
On Windows, a serialized cache exceeding 2,560 UTF-8 bytes cannot be persisted; after verifying a full protected-cache wipe, login or silent refresh succeeds with the authenticated cache in session-only memory.
If that wipe fails or cannot be verified, authentication fails without switching storage modes.
Store read, write or invalidation failures fail authentication with `LOGIN_FAILED` during login or `AUTH_REQUIRED` during silent acquisition, rather than switching storage modes.
Each profile holds one account.
Every explicit login wipes that profile's entire in-memory and protected MSAL cache before acquiring credentials, and fails if the wipe cannot be verified.
Switching accounts within a profile requires a fresh login; use separate profiles for separate identities.
Restore OS credential store access before retrying a failed wipe.
An inaccessible credential service on a headless system can therefore block authentication even when the native module loads.
There is no plaintext credential-cache fallback.
Tokens remain opaque and never appear in profile views, stdout or authentication diagnostics.
The delegated credential service binds account context to the configured tenant/client and refreshes at a 60-second expiry margin.
Tests use fake credential providers, fake time and real MSAL cache handling with fake network and storage boundaries; they never sign in or contact a real token endpoint.

Create an application profile with an organization-owned app registration and exactly one credential provider:

```sh
mg-axi profile create --name daemon --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --certificate-thumbprint <40-hex-digit-thumbprint>
mg-axi profile create --name batch --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --federated
mg-axi profile show --profile daemon
```

Omitting `--mode` selects delegated mode; certificate and federation flags require application mode.
Application profiles reject `--allow-device-code`, and `login` with either login method exits 2 before authentication.
For certificates, register the matching public certificate on the app registration.
The provider expects a PEM private key in the OS credential store under service `mg-axi` and account equal to the profile's `credentialRef.key`, shown by `profile show`.
Profile creation records the thumbprint and reference; it does not import the private key, and there is no CLI key-import command.
Provision that key through protected storage tooling; never put private keys in argv or profile JSON.
Certificate acquisition fails if `keytar` or the referenced key is unavailable; there is no plaintext or session-only key fallback.
For workload federation, configure the app registration's federated credential mapping for the workload and set `AZURE_FEDERATED_TOKEN_FILE` to its assertion file.
The provider reads the file on each assertion request so projected tokens can rotate; missing or empty assertions fail authentication.
Application credentials remain in process memory and are reacquired at a 60-second expiry margin.
The service requests only `https://graph.microsoft.com/.default`, representing the app registration's admin-consented Graph application permissions, with no signed-in user.
Ask an administrator to grant those permissions on the configured app registration; per-command delegated scopes cannot narrow the application token.
Acquisition failures return `AUTH_REQUIRED` with consent and certificate/federation guidance, without user or device-code fallback.

Run `corepack pnpm build`, `corepack pnpm test` (test files run at most 4 at a time via `--test-concurrency=4` in the `test` script) and `corepack pnpm lint` for shell validation.
The [CI workflow](.github/workflows/ci.yml) defines the platform/runtime matrix for shell build, test and lint checks, validates the Python inventory tooling separately, and rejects stale generated docs.
The [implementation plan](PLAN.md) remains the design authority.

The pinned [Entra operation inventory](docs/inventory.md) defines the INV-01 discovery boundary and schema for later build slices.

## Release

This is the supported Entra read and gated write surface, not full Entra coverage.
See the generated [capability report](docs/coverage.md) for implemented reads, writes and discovery dispositions, and the [skill command table](skills/mg-axi/SKILL.md#orientation) for all executable leaves, including local commands.
The package is published to npm as `@knowttl/mg-axi` with public access.
Install it globally with `npm install --global @knowttl/mg-axi`, or run without a global install via `npx -y @knowttl/mg-axi --help`.
The `.github/workflows/release-please.yml` workflow manages release PRs on `main` and builds, checks and publishes newly created releases with provenance using npm trusted publishing (OIDC) from the public repository.
Publication is skipped when the release version already exists on npm.
Version `0.1.0` was bootstrap-published on npm, and release-please continues from that version.
The packed files are `dist`, the discovery inventory, `skills/mg-axi`, `docs/coverage.md` and this README.
Follow the [checkout instructions](#mg-axi) to install dependencies, build, link the binary and run the version probe.

After linking the checkout binary, `mg-axi setup` shows the build steps, the selected configuration path, the configured profiles and the capability summary.
It writes nothing, signs in nowhere and installs no hooks; ordinary commands never gain installation side effects.
Create profiles with `mg-axi profile create`, sign delegated profiles in with `mg-axi login`, and check access with `mg-axi doctor`.
Doctor checks `--profile <name>` when supplied, otherwise the configured default, or all configured profiles when no default exists.
Doctor performs one bounded `entra user list --limit 1` read per profile with silent credential acquisition only: it never opens a browser, never shows a device-code challenge, never auto-installs and never enables writes.
Missing profiles, an unknown selected profile or invalid configuration fail before Graph reads.
Read failures report per profile with rerun guidance and a nonzero exit.

The installable skill lives at [skills/mg-axi/SKILL.md](skills/mg-axi/SKILL.md).
Agent discovery has two install paths, and either suffices alone.
Install the skill explicitly with `npx skills add knowttl/mg-axi --skill mg-axi`; the setup command only shows guidance and does not install skills.
For ambient context at every agent session start instead, install the CLI globally and opt into the session hook:
```sh
npm install --global @knowttl/mg-axi
mg-axi setup hooks
```
This installs a `SessionStart` hook for Claude Code, Codex and OpenCode that prints a short local-only summary (configured profile names with tenant labels, auth-cache presence, write posture and version).
The hook makes no network, Graph or sign-in call, and prints a short `not configured` record with exit 0 when unconfigured.
Restart the agent session after running the installer.
No other ordinary command installs hooks, plugins or configuration.
The skill file is generated in full from the template in [src/docs.ts](src/docs.ts) and the command catalogue; [docs/coverage.md](docs/coverage.md) also uses the discovery inventory.
After building, regenerate both with `corepack pnpm run docs:generate` and verify freshness with `corepack pnpm run docs:check`; CI runs the freshness check.
Critical journeys stay packaged offline: `test/pack.test.mjs` drives setup, doctor and the user, group, Conditional Access and sign-in reads through the packaged executable with fixture credentials and blocked networking.
