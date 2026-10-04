# mg-axi

Agent-facing Microsoft Graph CLI with one shared core and domain packs, read-only by default.
The Entra pack comes first, with phased full coverage and later gated named writes.

CLI-01 provides a local TypeScript/AXI shell, strict command catalogue, leaf help and fast version probes.
AUTH-01 adds versioned dedicated-app delegated profiles and explicit login.
AUTH-02 adds certificate and workload-federated application profiles and a client-credentials service.
CORE-01 adds the shared policy-enforced Graph read session, exercised through an injected fixture HTTP transport.
CORE-02 adds session collections, query validation, bounded retries and cancellation under the [read execution contract](docs/execution.md#read-mechanics-and-source-contracts).
API-01 executes `mg-axi api get <path>`, serving the reviewed v1.0 raw surface through that session.
READ-01 executes Entra user list/show through that session in both delegated and application modes; current usage follows below.
READ-02 adds group list/show and direct or transitive member and parent-membership reads through the same session.
READ-05 executes Entra sign-in and directory-audit list/show through that session; log usage follows the user usage below.
READ-07 adds application and service-principal list/show with credential expiry metadata and owner reads through the same session.
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

The home view reports unavailable tenant summaries explicitly.
`entra user list` defaults to basic properties (`id`, `displayName`, `userPrincipalName`, `mail`); `entra user show --id <user-id-or-upn>` defaults to the richer server property set.
`--select` requests properties from the [supported user property set](src/entra-users.ts); `--fields` projects locally and must be a subset of the fetched selection.
Text values longer than 500 characters are truncated, including strings in `businessPhones`; `--full` removes text truncation without lifting redaction or row caps.
Explicit null values stay null, missing properties stay absent, and denied reads return structured errors rather than empty results.
Lists accept `--filter` for an OData filter and default to a 100-row cap; `--limit` changes the cap, while the incompatible `--all` follows pages within request, byte and deadline budgets.
Partial lists report `count.complete: false`, a reason and an opaque `cursor` preserving unreturned rows.
Resume with `--cursor <cursor-from-output>` using the same profile, authentication scopes and API version; original `--select` and `--filter` values may be repeated or omitted, and conflicting values fail validation.
Repeat `--fields` and `--full` when the same local view is wanted; these are not saved in the cursor.
Delegated user reads default to `https://graph.microsoft.com/User.Read.All` with `--scopes` available for least-privilege basics; application profiles use the configured Graph `.default` audience and reject delegated scopes.
User reads acquire credentials silently; a resume containing only buffered rows can finish without another Graph request.
Unknown flags, unexpected arguments, missing required values and unsupported combinations exit 2 before credential acquisition or HTTP.
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

`entra group list` defaults to compact properties (`id`, `displayName`, `mail`, `groupTypes`); `entra group show --id <group-id>` defaults to the richer reviewed group set including `isAssignableToRole`, which marks groups eligible for role assignment.
Group `--select` accepts the [reviewed group property set](src/entra-groups.ts); `--fields` must be a subset of the fetched selection.
Role-assignable membership changes need role-management permission and belong to a later write slice, never to these reads.
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
Both modes need Policy.Read.All, Policy.Read.ConditionalAccess or Policy.ReadWrite.ConditionalAccess in addition to AuditLog.Read.All; delegated callers also need Conditional Access Administrator, Global Reader, Security Administrator or Security Reader.
For delegated CA detail, log in and repeat the read with `--scopes https://graph.microsoft.com/AuditLog.Read.All,https://graph.microsoft.com/Policy.Read.All`.
Denied log reads name the operation's supported directory roles and the conservative P1/P2 deployment prerequisite instead of only the generic grant/role/licence cause.

`entra application list` and `entra service-principal list` default to compact properties (`id`, `appId`, `displayName`); `appId` is the client ID, distinct from the object `id`, and both are included by default, while custom `--select` or `--fields` can omit either.
`entra application show --id <application-object-id>` defaults to the richer reviewed application set including `keyCredentials` and `passwordCredentials`.
`entra service-principal show --id <service-principal-object-id>` defaults to the richer reviewed service-principal set including `servicePrincipalType` and both credential collections.
Credential collections expose only expiry metadata (`keyId`, `displayName`, `startDateTime`, `endDateTime`); the shared session drops every other credential subfield before output or cursor buffering and applies the same filtering to buffered rows on resume, including older cursors.
Non-array credential collections become empty arrays, and non-object entries are dropped; other properties retain the null/missing behavior described above.
Secret-minting routes are never constructed.
`entra application owner list --application <application-object-id>` and `entra service-principal owner list --service-principal <service-principal-object-id>` list owners; rows carry `@odata.type` naming the owner kind, and consent grants stay out - they belong to READ-08.
Application lists return `applications`, service-principal lists return `servicePrincipals`, owner lists return `owners`, and single-object reads return `application` or `servicePrincipal`.
App and service-principal `--select` accepts the [reviewed property sets](src/entra-apps.ts); `--fields` must be a subset of the fetched selection.
Owner rows default to `id`, `displayName` and `mail`, the only selectable owner properties; rows without non-null descriptive properties are preserved with a hint about limited consent or unset properties.
The named-list caps, `count`, cursor resume rules and 500-character text truncation described above also apply to these reads; resume owner lists with the same `--application` or `--service-principal` object ID.
Delegated application reads default to `https://graph.microsoft.com/Application.Read.All`, while application profiles use the configured `.default` audience.

Configure the registration's Mobile and desktop applications redirect URI as `http://localhost` for browser login.
The first created profile is the default; `--profile` selects another identity explicitly.
Configuration defaults to `~/.mg-axi/config.json`; `MG_AXI_CONFIG` selects a separate configuration file.
Version 1 stores tenant/client/cloud, delegated or application mode, enabled packs, preview/sensitive policy, device-code opt-in and a unique credential reference only.
Unknown fields, unsupported versions/clouds and inlined credential material are rejected with recovery guidance.
Creation never overwrites an existing profile.
Preview is disabled and sensitive areas are empty in newly created profiles.

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
Server OData parameters use `--odata`; `--query` is reserved for output queries and is not implemented here.
Omitting `$select` requests the route's reviewed fields, and every response is filtered to reviewed fields before output.
Relationship expansion (`$expand`) is unavailable.
Unreviewed, secret-value, mail/file-content, beta and write routes fail before credentials, and pack, preview and sensitive-area policy still runs in the shared session.
Delegated reads take explicit `--scopes` like login; application profiles use the configured `.default` audience and reject `--scopes`.
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
The [CI workflow](.github/workflows/ci.yml) defines the platform/runtime matrix for shell build, test and lint checks, and validates the Python inventory tooling separately.
The [implementation plan](PLAN.md) remains the design authority.
Generated skill, setup and capability reporting ship in PACK-01; no session hooks are installed by ordinary commands.

The pinned [Entra operation inventory](docs/inventory.md) defines the INV-01 discovery boundary and schema for later build slices.
