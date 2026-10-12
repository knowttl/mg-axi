<h1 align="center">mg-axi</h1>

<p align="center">
  <a href="https://www.npmjs.com/package/@knowttl/mg-axi"><img alt="npm" src="https://img.shields.io/npm/v/@knowttl/mg-axi?style=flat-square" /></a>
  <a href="https://github.com/kunchenguid/axi"><img alt="AXI" src="https://img.shields.io/badge/AXI-Agent%20eXperience%20Interface-green?style=flat-square" /></a>
</p>

Microsoft Graph CLI for agents, designed with AXI (Agent eXperience Interface).

mg-axi provides reviewed Entra reads and explicitly gated named writes for workforce tenants in the commercial cloud.
It emits compact TOON output, includes contextual next-step suggestions and keeps errors machine-readable on stdout.
It is read-only by default; the shipped surface does not cover all of Microsoft Graph or Entra.

## Agent Ergonomics

- Compact list schemas keep the first response small.
- Lists distinguish complete, empty and partial results, with server totals when available.
- Opaque cursors preserve remaining rows for bounded reads.
- `--full` removes text truncation while preserving redaction and row limits.
- Leaf-level `--help` describes required arguments, flags and examples.
- Version probes, help and the local home view require no tenant sign-in.

## Quick Start

Requires Node.js **22.12.0 or later**.
Graph access also requires an organization-owned app registration, the appropriate consent and a configured profile.

### Skill installation

Install the agent skill:

```sh
npx skills add knowttl/mg-axi --skill mg-axi
```

The [skill](skills/mg-axi/SKILL.md) guides agents to use `npx -y @knowttl/mg-axi` and contains the full executable command table.
The skill and session hooks below are alternative discovery paths; either can be used alone.
Neither grants Graph access.

### Zero setup

Inspect local status and help without installing a global binary:

```sh
npx -y @knowttl/mg-axi
npx -y @knowttl/mg-axi --version
npx -y @knowttl/mg-axi entra user list --help
```

Use `npx -y @knowttl/mg-axi` in place of `mg-axi` in the examples below when running without a global install.
To pin a release, use `npx -y @knowttl/mg-axi@<version>`.

## Other ways to install

### Global install and session hooks

Install optional ambient context for Claude Code, Codex and OpenCode:

```sh
npm install --global @knowttl/mg-axi
mg-axi setup hooks
```

Restart the agent session after installation.
The session hook reports local profile names, tenant labels, authentication-cache presence, write policy and version.
It makes no network, Graph or sign-in calls and exits successfully when unconfigured.
Only the explicit hook installer changes agent integrations.

### Direct install

For direct command entry, install globally and inspect setup guidance:

```sh
npm install --global @knowttl/mg-axi
mg-axi setup
```

`setup` displays installation, configuration and capability guidance.
It writes nothing, signs in nowhere and installs no hooks or skills.

## Usage

Create a delegated profile using the tenant and app registration UUIDs, then explicitly sign in for user reads:

```sh
mg-axi profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial
mg-axi login --profile soc --scopes https://graph.microsoft.com/User.Read.All
mg-axi doctor --profile soc
mg-axi entra user list --profile soc --limit 10
mg-axi entra user show --profile soc --id <user-id-or-upn>
```

The first profile created becomes the default; `--profile` selects another identity.
Profile creation does not overwrite existing identities or sign in.
`doctor` performs one bounded user-list read per selected profile with silent credential acquisition.
Without `--profile`, it checks the default profile, or all profiles if no default is set.
Missing configuration and failed reads return errors rather than suggesting that access succeeded.

Browser login is the default and uses MSAL with PKCE.
Use full Graph scope names separated by commas; delegated login rejects `.default`.
Device code requires both profile creation with `--allow-device-code` and login with `--method device-code`, subject to organization policy.
Browser failure never falls back to device code.
Ordinary reads acquire credentials silently and report when login, consent or policy intervention is required.
Login does not modify app registration or automatically request blanket write consent.

### Command surface and subcommands

The [capability report](docs/coverage.md) distinguishes implemented operations from deferred or excluded ones.
The [command table](skills/mg-axi/SKILL.md#orientation) lists every executable leaf; use its leaf-level help for exact flags and permission defaults.

| Area | Shipped surface |
| --- | --- |
| Local configuration | `profile create`, `profile list`, `profile show`, `login`, `home`, `setup`, `setup hooks`, `doctor` |
| Directory | Users, groups and memberships, devices, administrative units, directory objects, deleted items, contacts, organizations and branding |
| Applications and consent | Applications, service principals, owners, federated credentials, delegated grants and app-role assignments |
| Roles and governance | Directory roles and assignments, PIM reads, access reviews, entitlement management and lifecycle workflows |
| Security | Conditional Access, authentication methods and registration reports, audit/sign-in/provisioning logs, user and workload risk |
| Tenant configuration | Domains and DNS records, subscriptions, certificate authentication, synchronization, terms of use, custom security attributes and group lifecycle policies |
| Identity and partner configuration | Identity providers, federation configurations, risk-prevention providers, data policy operations, contracts, delegated administration, multitenant organization and tenant information |
| Reviewed raw reads | `api get <path>` for allowlisted v1.0 routes |
| Named writes | Group member add, account enable/disable, session revocation, Conditional Access policy update and risky-user dismissal |

Delegated read commands default to the scope for their operation, which may differ from the user-read scope in the quick start.
Sign in explicitly with the appropriate scopes before switching areas.
Read commands reject unapproved or write scopes before acquiring credentials.
Permissions, directory roles and licences vary by operation; see the [access and licence guidance](docs/graph-coverage.md).
A 403 response alone does not identify which prerequisite is missing.

```sh
# Group membership reads
mg-axi login --profile soc --scopes https://graph.microsoft.com/GroupMember.Read.All
mg-axi entra group list --profile soc --limit 10
mg-axi entra group member list --profile soc --group <group-id>

# Conditional Access inspection
mg-axi login --profile soc --scopes https://graph.microsoft.com/Policy.Read.All
mg-axi entra conditional-access policy list --profile soc
mg-axi entra conditional-access policy show --profile soc --id <policy-id>

# Reviewed raw reads
mg-axi login --profile soc --scopes https://graph.microsoft.com/User.Read.All,https://graph.microsoft.com/GroupMember.Read.All
mg-axi api get /users --profile soc --scopes https://graph.microsoft.com/User.Read.All
mg-axi api get /groups --profile soc --odata '$filter=securityEnabled eq true&$top=5' --scopes https://graph.microsoft.com/GroupMember.Read.All
```

Raw access is not an unrestricted Graph client.
Unreviewed routes, secret-returning routes, mail/file content, beta raw reads and raw writes are refused before credentials.
Returned data is restricted to reviewed fields; completeness describes the requested collection, not full Graph property coverage.
Secret material, authentication-method phone numbers and data-policy storage locations are redacted, including in full output, previews and audit metadata.

### TOON output behavior

| Control | Behavior |
| --- | --- |
| `--select` | Request reviewed properties where supported; some routes select locally instead |
| `--fields` | Project locally from properties already fetched |
| `--full` | Remove 500-character text truncation without lifting redaction or row caps |
| `--limit` | Cap collection rows, normally 100 by default |
| `--all` | Follow pages within request, byte and deadline budgets; incompatible with `--limit` |
| `--cursor` | Resume a partial collection under the same profile, scopes and API version |

Availability varies by leaf; consult its `--help`.
List output includes `count`, `total` and `complete` information.
For example, a partial user list without a server total reports `count: 10 users shown, more available` and `total: null`.
A known server total uses a count such as `10 of 143 users`; totals are never invented or fetched with an extra request.

Partial results preserve buffered rows and the next page in an opaque cursor.
For named reads, repeat the original query flags or omit them; conflicting queries are rejected.
Repeat `--fields` and `--full` to retain the same local view.
Raw reads reuse their original query when `--odata` is omitted on resume.
For large raw cursors, save the token to a file and supply it through stdin:

```sh
mg-axi api get /users --profile soc --cursor - --all --scopes https://graph.microsoft.com/User.Read.All < cursor.txt
```

Both raw cursor input forms have a 16 MB ceiling.
Treat cursor tokens and output as tenant data.
Denied reads return structured errors, not empty success results.

### Named writes and safety

**Writes require manual profile opt-in, an operation allowlist, `--execute` and exact target confirmation.**
Without `--execute`, a named write previews the change and journals nothing, but preview still requires the write-enabled profile and prerequisite read access.
`MG_AXI_READ_ONLY=1` overrides profile opt-in.
Raw writes and beta writes remain blocked.

Manually add a `writes` object to the intended profile in the configuration file:

```json
"writes": {
  "allowWrites": true,
  "operations": ["mg.entra.group.member.add"]
}
```

No CLI command enables this policy.
The operation names are exact and scoped to the configured tenant:

| Command | Allowlist operation | `--confirm` repeats |
| --- | --- | --- |
| `entra group member add` | `mg.entra.group.member.add` | `--group` |
| `entra user update` | `entra.user.update` | `--user` |
| `entra user revoke-sessions` | `entra.user.revokeSessions` | `--user` |
| `entra conditional-access policy update` | `entra.conditional-access.policy.update` | `--id` |
| `entra risky-user dismiss` | `entra.risky-user.dismiss` | `--user` |

These examples preview only after the corresponding policy and permissions are configured:

```sh
mg-axi entra group member add --profile soc --group <group-id> --user <user-id>
mg-axi entra user update --profile soc --user <user-id-or-upn> --account-enabled false
mg-axi entra user revoke-sessions --profile soc --user <user-id>
mg-axi entra conditional-access policy update --profile soc --id <policy-id> --state disabled
mg-axi entra risky-user dismiss --profile soc --user <risky-user-id>
```

Adding `--execute --confirm <exact-target>` authorizes sending.
Review the preview and operation-specific permissions first:

| Write | Permission and safety boundary |
| --- | --- |
| Group member add | `GroupMember.ReadWrite.All`, plus user-read permission; delegated callers need a supported groups role. Only user object IDs are accepted. Role-assignable, dynamic and distribution groups are refused. Existing membership is a no-op. |
| Account enable/disable | `User.EnableDisableAccount.All` plus `User.Read.All`. Delegated callers need Privileged Authentication Administrator for admin targets and must generally outrank the target; application callers also need a higher-privileged admin role assignment. Both enabling and disabling are disruptive. `--account-enabled` accepts exactly `true` or `false`. IDs and literal UPNs are accepted; quote guest UPNs containing `#EXT#`. |
| Session revocation | `User.RevokeSessions.All` plus prerequisite user-read access. Only user object IDs are accepted. The action cannot be undone, may take minutes to affect sessions and does not revoke external users' home-tenant sessions. Accepted 2xx means the action was accepted, without verification readback. |
| Conditional Access update | `Policy.Read.All` plus `Policy.ReadWrite.ConditionalAccess`; delegated callers need Conditional Access Administrator or Security Administrator. P1 is required, with P2 for risk-based features. Reviewed fields only; JSON flags replace their entire corresponding block. |
| Risk dismissal | `IdentityRiskyUser.ReadWrite.All` plus `IdentityRiskyUser.Read.All`; delegated callers need Security Administrator, and risky-user access requires P2. Dismissal is not remediation: it resets no credential and revokes no session. |

Conditional Access enforcement changes require a readable lockout analysis.
Enabled policies require `--acknowledge-lockout-risk` after verifying emergency-access coverage.
An all-users block policy covering browser and modern clients on all apps or admin portals without recognized narrowing conditions is refused even with acknowledgement.
User, group and role exclusions or application filters alone cannot establish emergency-access protection.
Disabled and report-only policies do not enforce access; display-name-only changes do not require lockout acknowledgement.
See the [policy implementation](src/entra-ca-policy-update.ts) for the precise static-analysis boundary.

Execution journals intent durably before sending and records its outcome.
A failed intent record blocks the send; journal locks and incomplete records require manual reconciliation.
Writes are not automatically retried or replayed.
After `OUTCOME_UNKNOWN` or an outcome-recording failure, read back the target before further action and never replay that intent.
Account updates and risk dismissal verify the resulting state; a mismatch reports `WRITE_CONFLICT` and a failed verification read reports `OUTCOME_UNKNOWN`.
After a Conditional Access update, explicitly read the policy back.
Fresh reads do not make writes atomic, and account and policy updates send no ETag precondition or promise of concurrency protection.
Do not assume rollback.
The [execution contract](docs/execution.md#named-writes) describes the shared guarantees.

## Configuration

Profiles store tenant/client/cloud identity, authentication mode, policy and credential references, never secret values.
Only commercial-cloud workforce profiles are supported.
Unknown configuration fields, unsupported versions/clouds and inline credentials are rejected.

| Variable | Meaning |
| --- | --- |
| `MG_AXI_CONFIG` | Configuration file override; default `~/.mg-axi/config.json` |
| `MG_AXI_READ_ONLY` | Set exactly `1` to refuse named writes regardless of profile policy |
| `MG_AXI_WRITE_LOG` | Nonblank mutation journal path override; default `~/.mg-axi/writes.log` |
| `AZURE_FEDERATED_TOKEN_FILE` | Default application workload-assertion file holder |

### Delegated credential storage

Delegated MSAL caches use the OS credential store through optional `keytar`, reported as `storage: os-protected`.
If the native module cannot load, authentication is session-only and lasts only for that process; there is no plaintext cache fallback.
Installing with `--ignore-scripts` can leave the native module unavailable.
If the credential service loads but cannot read, write or invalidate the cache, authentication fails rather than silently changing storage modes.
An inaccessible headless credential service can therefore block authentication.

Each profile holds one account.
Every explicit login wipes that profile's protected and in-memory cache before acquiring replacement credentials and fails if the wipe cannot be verified.
Use separate profiles for separate identities; restore credential-store access before retrying a failed wipe.
On Windows, caches larger than 2,560 UTF-8 bytes use session-only memory only after a verified protected-cache wipe.
Tokens remain opaque and never appear in profile views, stdout or authentication diagnostics.
Credential acquisition is bound to the configured tenant/client and refreshes at a 60-second expiry margin.

### Application credentials and headless use

Application profiles use client credentials without a signed-in user and request the configured Graph `.default` audience.
An administrator must grant Graph application permissions on the app registration; delegated `--scopes` cannot narrow an application token.
Application profiles reject device-code opt-in and `login`, and acquisition failures never fall back to user authentication.

Choose exactly one credential provider:

```sh
# Certificate: register the matching public certificate on the app registration
mg-axi profile create --name daemon --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --certificate-thumbprint <40-hex-digit-thumbprint>

# Workload federation: configure the app registration's federated credential mapping
mg-axi profile create --name batch --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --federated

# Headless client secret: supply the value through the named environment variable
mg-axi profile create --name svc --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --client-secret-env MG_AXI_CLIENT_SECRET
mg-axi profile show --profile svc
```

By default, certificate keys come from OS credential-store service `mg-axi`, with the account given by `credentialRef.key` in `profile show`.
Profile creation records the reference and thumbprint; it does not import a key and there is no CLI key-import command.
Provision keys through protected storage tooling, never argv or profile JSON.
Certificate acquisition fails if the referenced key or credential store is unavailable.

For headless certificate holders, add `--certificate-key-file <path>` or `--certificate-key-env <variable-name>`.
Client secrets similarly accept `--client-secret-file <path>` instead of `--client-secret-env`.
Key and secret files must be regular files, never symlinks, owned by the current user with mode 0600; Windows skips ownership and mode checks.
Relative paths resolve at acquisition time.
Missing, empty or exposed holders fail closed without disclosing values.
PEM environment values may contain escaped `\n`; passphrase-protected keys are unsupported.
Prefer certificates or federation over bearer client secrets, and rotate secrets on a schedule.

Federation reads the assertion file on every acquisition so projected tokens can rotate.
Use `--federated-token-file-env <variable-name>` to select another holder variable.
References, rather than values, are stored in configuration; rotation updates the environment or file without changing the profile.
Application credentials remain in process memory and are reacquired at a 60-second expiry margin.
See [headless credential guidance](docs/headless-credentials.md) for holder examples.

## Development & Testing

Use Node.js 22.12.0 or later and the pinned pnpm 10.34.6 from [package.json](package.json).
From a checkout root:

```sh
corepack pnpm install --frozen-lockfile --ignore-scripts --config.confirm-modules-purge=false
corepack pnpm build
node dist/bin/mg-axi.js --version
node dist/bin/mg-axi.js entra user list --help
corepack pnpm test
corepack pnpm lint
corepack pnpm run docs:check
```

Tests use fixture credentials and transports without signing in or contacting real tenants.
The test script limits concurrency to four files.
To use the built checkout binary on PATH, run `npm link --ignore-scripts` from the checkout root.
Otherwise substitute `node dist/bin/mg-axi.js` for `mg-axi` in examples.

The agent skill and capability report are generated from the command catalogue, documentation template and discovery inventory.
After changing their sources, run `corepack pnpm run docs:generate`, then `corepack pnpm run docs:check`.
The [implementation plan](PLAN.md), [operation inventory](docs/inventory.md) and [execution contract](docs/execution.md) describe design and coverage boundaries.
