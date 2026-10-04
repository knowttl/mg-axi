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
READ-03 adds Conditional Access policy and named-location list/show as separate grammar through the same session; Conditional Access usage follows the device and administrative-unit usage below.
READ-05 executes Entra sign-in and directory-audit list/show through that session; log usage follows the Conditional Access usage below.
READ-07 adds application and service-principal list/show with credential expiry metadata and owner reads through the same session.
READ-10 adds directory-device and administrative-unit list/show and unit-member reads through the same session.
READ-08 adds service-principal delegated-grant and app-role-assignment consent reads for a named client through the same session.
EXT-01 (domains) adds tenant-domain list/show, per-domain verification and service-configuration DNS record reads, and top-level domain DNS record reads through the same session.
READ-06 executes Entra risky-user and risk-detection list/show through that session; risk usage follows the log usage below.
EXT-01 (organization) adds tenant-organization list/show, default sign-in branding metadata and locale branding reads through the same session.
Tests use fixture credential and transport providers; no tenant, real credentials or network access are required for help or an unconfigured home view.

Use the Node requirement and pinned pnpm version declared in [package.json](package.json):

```sh
corepack pnpm install --frozen-lockfile --ignore-scripts --config.confirm-modules-purge=false
corepack pnpm build
node dist/bin/mg-axi.js
node dist/bin/mg-axi.js --version
node dist/bin/mg-axi.js entra user list --help
node dist/bin/mg-axi.js entra user show --help
```

Run checkout commands from the repository root with `node dist/bin/mg-axi.js`.
Where examples or CLI output use `mg-axi`, substitute `node dist/bin/mg-axi.js` unless you have separately installed the executable on PATH.

The home view reports unavailable tenant summaries explicitly.
`entra user list` defaults to basic properties (`id`, `displayName`, `userPrincipalName`, `mail`); `entra user show --id <user-id-or-upn>` defaults to the richer server property set.
`--select` requests properties from the [supported user property set](src/entra-users.ts); `--fields` projects locally and must be a subset of the fetched selection.
Text values longer than 500 characters are truncated, including strings in `businessPhones`; `--full` removes text truncation without lifting redaction or row caps.
Explicit null values stay null, missing properties stay absent, and denied reads return structured errors rather than empty results.
User lists accept `--filter` for an OData filter and default to a 100-row cap; `--limit` changes the cap, while the incompatible `--all` follows pages within request, byte and deadline budgets.
Partial lists report `count.complete: false`, a reason and an opaque `cursor` preserving unreturned rows.
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
Resource identifiers in named commands and raw paths cannot begin with `$`; OData route segments such as `$count`, `$value` and `$ref` cannot be used as IDs and fail validation before credentials.
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
`--filter` passes through as plain `$filter` with no `$count` or `ConsistencyLevel` contract; the named-list caps, `count`, cursor resume rules and 500-character text truncation described above also apply.
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
The named-list caps, `count`, cursors, null/missing preservation and 500-character text truncation described above also apply to group and relationship reads.
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
`--filter` passes through as plain `$filter` with no `$count` or `ConsistencyLevel` contract; the named-list caps, `count`, cursors, null/missing preservation and 500-character text truncation described above also apply.
Delegated directory-role and role-assignment reads default to `https://graph.microsoft.com/RoleManagement.Read.Directory`; eligible PIM reads default to `https://graph.microsoft.com/RoleEligibilitySchedule.Read.Directory` and active PIM reads to `https://graph.microsoft.com/RoleAssignmentSchedule.Read.Directory`, while application profiles use the configured `.default` audience.
Delegated callers additionally need a supported directory role per operation (for example Privileged Role Administrator, Global Reader or Security Reader for role reads, Directory Readers for assignments, Security Operator for PIM reads).
Denied reads name that role requirement instead of only the generic grant/role/licence cause.
Built-in roles are base inventory and custom role assignments need P1; PIM reads need P2 or ID Governance.
Role assignment, activation and every other PIM mutation belongs to later write slices, never to these reads.

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
The named-list caps, `count`, cursors, null/missing preservation and 500-character text truncation described above also apply to device and administrative-unit reads.
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
Delegated organization reads default to `https://graph.microsoft.com/Organization.Read.All` for full metadata; delegated `User.Read` returns only `id`, `displayName` and `verifiedDomains` with every other property null.
`entra organization branding show --organization <organization-id>` reads the default branding metadata (non-Stream text and URLs); the session sends the documented `Accept-Language: 0` header and locale variants come from the localizations collection.
Branding leaves default to delegated `https://graph.microsoft.com/User.Read`, the documented least-privileged scope; `OrganizationalBranding.Read.All` is the purpose-built alternative and `Organization.Read.All` also works when passed as `--scopes`.
Stream image properties (`bannerLogo`, `backgroundImage` and friends) are refused before credentials: they need a later piece with its own binary-output contract.
A branding 404 may indicate unconfigured branding, a missing locale, or a missing or inaccessible organization; configuring custom branding needs P1/P2, and contact fields on the organization are personal data.
Delegated callers additionally need a supported Entra role (Directory Readers or Global Reader for organizations; Global Reader or Organizational Branding Administrator for branding); personal Microsoft accounts are not supported.
Organization and localization lists return `organizations` and `brandingLocalizations`, single-object reads return `organization`, `branding` and `brandingLocalization`.
The named-list caps, `count`, cursors, null/missing preservation and 500-character text truncation described above also apply to organization and branding reads.
Organization and localization lists offer no `--filter`: Graph documents `$select` only on these routes, so the flag is refused before credentials.
Denied organization and branding reads name the scope, role and licensing guidance instead of only the generic cause.
No organization mutation lives here; certificate-based-auth configuration, extensions, beta-only settings and the POST lookup actions belong to later pieces; see the [organization scope decisions](docs/coverage.md#ext-01-organization-scope-decisions) for deferred reads and later subfamilies.

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
The named-list caps, `count`, cursors, null/missing preservation and 500-character text truncation described above also apply to domain and DNS record reads.
Resume record lists with the same `--domain` where applicable, profile, scopes and API version; omit or repeat the original server query flags, and repeat local `--fields` and `--full` when wanted.
Delegated domain and DNS record reads default to `https://graph.microsoft.com/Domain.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers additionally need a supported Entra role (Domain Name Administrator or Global Reader are least-privileged); personal Microsoft accounts are not supported.
No P1/P2 prerequisite is stated for domain reads; denied reads name the scope, role and licensing guidance instead of only the generic cause.
No domain mutation lives here; see the [domain scope decisions](docs/coverage.md#ext-01-domain-scope-decisions) for deferred reads and later subfamilies.

Log in with `https://graph.microsoft.com/Policy.Read.All`, then read Conditional Access policies and named locations as separate grammar:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/Policy.Read.All
mg-axi entra conditional-access policy list --profile soc --limit 10
mg-axi entra conditional-access policy list --profile soc --filter "state eq 'enabled'"
mg-axi entra conditional-access policy show --profile soc --id <policy-id>
mg-axi entra conditional-access named-location list --profile soc
mg-axi entra conditional-access named-location show --profile soc --id <named-location-id>
```

`entra conditional-access policy list` defaults to compact properties (`id`, `displayName`, `state`); `policy show --id <policy-id>` defaults to the full reviewed condition and control property set.
`entra conditional-access named-location list` defaults to `id` and `displayName`; `named-location show --id <named-location-id>` defaults to the full reviewed location set.
Policy `--select` accepts the [reviewed policy property set](src/entra-conditional-access.ts) and location `--select` accepts the [reviewed location property set](src/entra-conditional-access.ts); `--fields` must be a subset of the fetched selection in each family.
Returned `@odata.type` stays visible on named-location rows so IP and country locations stay distinguishable alongside any `--fields` projection.
Policy lists return `policies`, location lists return `namedLocations`, single-policy reads return `policy` and single-location reads return `namedLocation`.
Successful empty collections return an empty list with `count.returned: 0`, `count.complete: true` and absence guidance.
The named-list caps, `count`, cursors, null/missing preservation and 500-character text truncation described above also apply to Conditional Access reads, including nested condition values; every truncated value carries a `--full` hint.
Resume either Conditional Access list with `--cursor -` and supply the returned cursor on stdin, for example `mg-axi entra conditional-access policy list --profile soc --cursor - < cursor.txt`.
Small cursors can also use `--cursor <token>`; both forms enforce a 16 MB size ceiling.
Use `--full` when reasoning from complete condition or control text; `--select` and `--fields` still determine which properties are visible, and missing properties remain unknown.
To replay a resumed result with `--full`, reuse the original input cursor; the returned cursor continues after that result.
Delegated policy and location reads default to `https://graph.microsoft.com/Policy.Read.All`; application profiles require admin-consented `Policy.Read.All`, use the configured Graph `.default` audience and reject `--scopes`.
HTTP 403 policy and location errors name the operation's supported directory roles (Conditional Access Administrator, Global Reader, Global Secure Access Administrator, Security Administrator or Security Reader for delegated access) and include guidance from the [licensing contract](docs/graph-coverage.md#licence-matrix-by-area); the response alone does not identify the missing prerequisite.
No policy mutation lives here; policy updates belong to a later write slice.

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
The named-list caps, `count`, cursors, null/missing preservation and 500-character text truncation described above also apply.
Resume method lists with the same `--user` ID or UPN, profile, scopes and API version; omit or repeat the original `--select`, and repeat local `--fields` and `--full` when wanted.
Delegated method reads default to `https://graph.microsoft.com/UserAuthenticationMethod.Read.All` (delegated self-reads may use `UserAuthenticationMethod.Read`) and registration reads default to `https://graph.microsoft.com/AuditLog.Read.All`, while application profiles use the configured `.default` audience.
Delegated callers acting on another user additionally need Global Reader, Authentication Administrator or Privileged Authentication Administrator for methods, and Reports Reader, Security Reader, Security Administrator or Global Reader for the report.
Denied reads name that role requirement instead of only the generic grant/role/licence cause.

Log in with `https://graph.microsoft.com/AuditLog.Read.All`, then query sign-ins and directory audits in bounded time windows:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/AuditLog.Read.All
mg-axi entra sign-in list --profile soc --since 2026-09-01T00:00:00Z --limit 10
mg-axi entra sign-in list --profile soc --since 2026-09-01T00:00:00Z --filter "status/errorCode ne 0" --all
mg-axi entra sign-in show --profile soc --id <sign-in-id>
mg-axi entra directory-audit list --profile soc --since 2026-09-01T00:00:00Z
mg-axi entra directory-audit show --profile soc --id <directory-audit-id>
```

Log lists always carry an explicit time bound: `--since` is required for a new query (with optional `--until` and `--filter` refinements), and resume reuses `--cursor` instead.
Resume validates the saved time bounds; a cursor from an unbounded raw query is rejected, so start a new query with `--since`.
Resume sign-in and directory-audit lists with `--cursor -` and supply the returned cursor on stdin, for example `mg-axi entra sign-in list --profile soc --cursor - < cursor.txt`.
Small cursors can also use `--cursor <token>`; both forms enforce a 16 MB size ceiling.
`entra sign-in list` defaults to `id`, `createdDateTime`, `userPrincipalName` and `appDisplayName`; `entra directory-audit list` defaults to `id`, `activityDateTime`, `activityDisplayName` and `result`.
Both log show commands default to the full reviewed property set.
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
The named-list caps, `count`, cursor resume rules and 500-character text truncation described above also apply to these reads; resume owner lists with the same `--application` or `--service-principal` object ID.
Delegated application reads default to `https://graph.microsoft.com/Application.Read.All`, while application profiles use the configured `.default` audience.

Log in with `https://graph.microsoft.com/IdentityRiskyUser.Read.All` and `https://graph.microsoft.com/IdentityRiskEvent.Read.All`, then triage risky users and risk detections:

```sh
mg-axi login --profile soc --scopes https://graph.microsoft.com/IdentityRiskyUser.Read.All,https://graph.microsoft.com/IdentityRiskEvent.Read.All
mg-axi entra risky-user list --profile soc --limit 10
mg-axi entra risky-user show --profile soc --id <risky-user-id>
mg-axi entra risk-detection list --profile soc --since 2026-09-01T00:00:00Z --limit 10
mg-axi entra risk-detection show --profile soc --id <risk-detection-id>
```

Delegated risky-user reads default to `https://graph.microsoft.com/IdentityRiskyUser.Read.All`; risk-detection reads default to `https://graph.microsoft.com/IdentityRiskEvent.Read.All`.
`--scopes` overrides those defaults; application profiles use the configured Graph `.default` audience and reject delegated scopes.
`entra risky-user list` is a state collection with an optional `--filter`; `entra risk-detection list` always carries an explicit time bound, so `--since` is required for a new query (with optional `--until` and `--filter` refinements) and resume reuses `--cursor` instead.
Resume validates the saved detectedDateTime bounds; a cursor from an unbounded raw query is rejected, so start a new query with `--since`.
Resume risk lists with `--cursor -` and supply the returned cursor on stdin, for example `mg-axi entra risk-detection list --profile soc --cursor - < cursor.txt`.
`entra risky-user list` defaults to `id`, `userPrincipalName`, `riskLevel` and `riskState`; `entra risk-detection list` defaults to `id`, `detectedDateTime`, `userPrincipalName` and `riskLevel`.
Both risk show commands default to the full reviewed property set.
`--select` requests properties from the [supported risk property sets](src/entra-risk.ts); `--fields` projects locally and must be a subset of the fetched selection.
Risk reads truncate text longer than 500 characters, including nested values such as `location` and the `additionalInfo` JSON string; `--full` restores complete text without lifting redaction, row caps or time bounds.
See the [access and licence contract](docs/graph-coverage.md#licence-matrix-by-area) for the riskyUsers P2 requirement and risk-detection P1/P2 detail boundaries.
Limited views stay limited: a premium detection without P2 detail reports `riskEventType` generic, hidden risk levels report the licence boundary instead of the level, and a null detection `correlationId` means no sign-in is associated.
To correlate a detection to sign-ins, use `risk-detection show` or select `activityDateTime`, then filter the sign-in list above on the detection's `userPrincipalName` in that activity window; sign-in reads require the separate `AuditLog.Read.All` login above, and there is no riskySignIns endpoint.
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
Omitting `$select` selects the route's reviewed fields, and every response is filtered to reviewed fields before output.
For `/users/<user-id>/authentication/methods`, `$select` is the only supported OData parameter and selects output properties locally; an explicit selection restricts output to that subset.
Method requests omit `$select` on initial requests, continuations, redirects and retries, while cursors preserve the local selection.
Other routes send `$select` to Graph.
Relationship expansion (`$expand`) is unavailable.
Unreviewed, secret-value, mail/file-content, beta and write routes fail before credentials, and pack, preview and sensitive-area policy still runs in the shared session.
Delegated raw reads require explicit `--scopes` from the supported read choices described above; application profiles use the configured `.default` audience and reject `--scopes`.
Collections return `returned`, `complete` and `value`, default to 100 rows, and follow pages within budget under `--all`.
`--limit` and `--all` cannot be combined.
Completion describes pagination, not visibility of every directory object; group-member results include a warning for the [v1.0 service-principal limitation](docs/graph-coverage.md#licensing-and-completeness-findings), even when `complete` is true.
Strings longer than 4000 characters are truncated; `--full` removes string truncation without disabling redaction, reviewed-field filtering or row caps.
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

Run `corepack pnpm build`, `corepack pnpm test` and `corepack pnpm lint` for shell validation.
The [CI workflow](.github/workflows/ci.yml) defines the platform/runtime matrix for shell build, test and lint checks, validates the Python inventory tooling separately, and rejects stale generated docs.
The [implementation plan](PLAN.md) remains the design authority.

The pinned [Entra operation inventory](docs/inventory.md) defines the INV-01 discovery boundary and schema for later build slices.

## Release

This is the supported Entra read and gated write surface, not full Entra coverage.
See the generated [capability report](docs/coverage.md) for implemented reads, writes and discovery dispositions, and the [skill command table](skills/mg-axi/SKILL.md#orientation) for all executable leaves, including local commands.
The package is marked private and ships no publish workflow: preparing this release never publishes it.
The packed files are `dist`, the discovery inventory, `skills/mg-axi`, `docs/coverage.md` and this README.
Follow the [checkout instructions](#mg-axi) to install dependencies, build and run the version probe.

From the repository root, `node dist/bin/mg-axi.js setup` shows the build steps, the selected configuration path, the configured profiles and the capability summary.
It writes nothing, signs in nowhere and installs no hooks; ordinary commands never gain installation side effects.
Create profiles with `node dist/bin/mg-axi.js profile create`, sign delegated profiles in with `node dist/bin/mg-axi.js login`, and check access with `node dist/bin/mg-axi.js doctor`.
Doctor checks `--profile <name>` when supplied, otherwise the configured default, or all configured profiles when no default exists.
Doctor performs one bounded `entra user list --limit 1` read per profile with silent credential acquisition only: it never opens a browser, never shows a device-code challenge, never auto-installs and never enables writes.
Missing profiles, an unknown selected profile or invalid configuration fail before Graph reads.
Read failures report per profile with rerun guidance and a nonzero exit.

The installable skill lives at [skills/mg-axi/SKILL.md](skills/mg-axi/SKILL.md).
Install it explicitly through your agent's skill installation mechanism; the setup command only shows guidance and does not install skills.
Its command table is generated from the command catalogue; [docs/coverage.md](docs/coverage.md) also uses the discovery inventory.
After building, regenerate both with `corepack pnpm run docs:generate` and verify freshness with `corepack pnpm run docs:check`; CI runs the freshness check.
Critical journeys stay packaged offline: `test/pack.test.mjs` drives setup, doctor and the user, group, Conditional Access and sign-in reads through the packaged executable with fixture credentials and blocked networking.
