# mg-axi

Microsoft Graph CLI for agents and the humans running them, with reviewed Entra and Defender XDR hunting reads, compact output and seventeen explicitly gated writes.

[![npm](https://img.shields.io/npm/v/@knowttl/mg-axi)](https://www.npmjs.com/package/@knowttl/mg-axi)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D22.12.0-green)](package.json)
[![Provenance](https://img.shields.io/badge/npm-provenance-blue)](https://www.npmjs.com/package/@knowttl/mg-axi?activeTab=provenance)
[![AXI](https://img.shields.io/badge/AXI-Agent%20eXperience%20Interface-green)](https://github.com/kunchenguid/axi)

```console
$ npx -y @knowttl/mg-axi@0.1.10 --version
0.1.10
```

[Quickstart](#quickstart) • [Use with your agent](#use-with-your-agent) • [Features](#features) • [Safety](#safety-and-permissions) • [Configuration](#configuration) • [Commands](#commands)

## Why this exists

Working directly with Graph REST requires agents to manage route selection, permissions, paging and response size.
mg-axi exposes reviewed commands with bounded lists, TOON output and contextual next steps, without requiring an MCP server.
It targets workforce Entra tenants in the commercial cloud; sovereign clouds, personal accounts and general Microsoft 365 file access are outside its supported surface.
The opt-in [mail pack](#mail-pack) provides nine read commands, two guarded downloads and twelve gated mailbox writes.
It covers message search, bodies, headers, attachment metadata, folders, settings, inbox rules, categories and message triage.
Sending, reply, forward and permanent deletion are unsupported.

## Features

- **Entra reads** cover directory, applications, governance, security and tenant configuration.
- **Bounded output** distinguishes empty, complete and partial lists, with opaque continuation cursors.
- **Credential isolation** binds authentication to an explicit tenant and organization-owned app.
- **Write gates** require manual profile enablement, an operation allowlist and exact target confirmation.
- **Offline discovery** provides local status, group help and leaf references without sign-in.
- **Redaction** protects sensitive fields in output, previews, errors and audit metadata, including with `--full`.

## Defender XDR advanced hunting

`mg-axi security hunting query --file hunt.kql --profile soc --timespan P7D --limit 1000` runs a reviewed read-only Graph POST to `/security/runHuntingQuery`.
Save this example as `hunt.kql`:

```kql
UrlClickEvents
| where ActionType == "ClickAllowed"
| project Timestamp, AccountUpn, Url, ActionType
| order by Timestamp desc
```

Before use, an administrator must grant Microsoft Graph **ThreatHunting.Read.All**, with admin consent, for the profile's delegated or application authentication mode.
Existing Entra permissions do not grant hunting access.
Add `"security"` to that profile's `enabledPacks` array in the configuration file (for example `["entra", "security"]`); existing profiles default to Entra only.
For delegated authentication, sign in explicitly with `mg-axi auth login --profile soc --scopes https://graph.microsoft.com/ThreatHunting.Read.All` after consent.
Application profiles use their configured Graph `.default` audience.
Defender data availability, licensing and the signed-in user's data access still apply; HTTP 403 alone does not identify the missing prerequisite.

The default `--timespan P1D` looks back from now; use whole-day durations such as `P7D`, or time durations such as `PT6H` or `PT30M`, from one second to 30 days.
Graph also applies any narrower time filter in KQL.
The default `--limit 100` accepts 1-10000 rows and appends a server-side `take` for one extra row to detect overflow.
Overflow returns `complete: false` and `reason: row-limit`; it never invents a total or cursor.
Microsoft documents no paging or cursor for this action; narrow the time window or query filters to retrieve more results.
Results use TOON and existing secret redaction, with recursive 500-character string previews; `--full` removes text truncation while retaining redaction and query bounds.
The client uses a 30-second deadline and a 5 MB response budget, refuses redirects and does not automatically replay hunting requests.
Microsoft's service limits are 100000 rows, 50 MB and three minutes per request, with tenant CPU and call quotas.
See [runHuntingQuery](https://learn.microsoft.com/en-us/graph/api/security-security-runhuntingquery?view=graph-rest-1.0), [security quotas](https://learn.microsoft.com/en-us/graph/api/resources/security-api-overview?view=graph-rest-1.0#quotas-and-resource-allocation), and [permission consent](https://learn.microsoft.com/en-us/graph/permissions-reference#threathuntingreadall).

KQL input is file-only and is never echoed in results or server errors.
The local validator accepts one table pipeline, ordinary quoted strings, `//` comments, balanced parentheses/brackets, and the following operators: `where`, `project`, `project-away`, `project-rename`, `project-reorder`, `extend`, `summarize`, `distinct`, `count`, `take`, `limit`, `top`, `sort`, `order`.
Only reviewed scalar and aggregate function calls are accepted, including `ago`, `datetime`, `bin`, `count`, `dcount`, `sum`, `arg_max`, `tostring`, `parse_json` and `parse_url`; the exact allowlist lives in [security-hunting.ts](src/security-hunting.ts).
The table allowlist covers `DeviceEvents`, `DeviceProcessEvents`, `DeviceNetworkEvents`, `DeviceFileEvents`, `DeviceFileCertificateInfo`, `DeviceRegistryEvents`, `DeviceLogonEvents`, `DeviceImageLoadEvents`, `DeviceInfo`, `DeviceNetworkInfo`, `DeviceTvmSoftwareInventory`, `DeviceTvmSoftwareVulnerabilities`, `DeviceTvmSoftwareVulnerabilitiesKB`, `EmailEvents`, `EmailPostDeliveryEvents`, `EmailUrlInfo`, `EmailAttachmentInfo`, `UrlClickEvents`, `IdentityLogonEvents`, `IdentityQueryEvents`, `IdentityDirectoryEvents`, `CloudAppEvents`, `AlertInfo`, `AlertEvidence`, `BehaviorInfo`, `BehaviorEntities`, `ExposureGraphNodes` and `ExposureGraphEdges`.
Everything outside this restricted grammar is refused before credentials: semicolon statements (including trailing semicolons), `let`/`set`, dot-prefixed control commands, stored/source functions, qualified functions, external data, plugins (`evaluate`/`invoke`), joins/unions, nested query pipelines and block comments.
Microsoft's reference Kusto parser checks syntax locally before the allowlist and before credentials or network access.
Syntax failures keep `VALIDATION_ERROR` (exit 2) and include up to five diagnostics with line, column, code, token and message; allowlist failures also identify the offending token and position.
Queries and quoted literals are never printed as error excerpts.
Successful parsing prints no validation claim: table/column availability, types and permissions are still checked by Graph.
Recognized hunting 400 errors expose only the semantic/syntax error class and line/column, with an opaque fallback for unknown shapes.
The undocumented Graph error shape is covered by synthetic offline fixtures and has not been verified against a live tenant.
Raw `api get` keeps its reviewed GET-only route list and does not expose `/security`.

## Quickstart

Requires Node.js **22.12.0 or later**.
Inspect local configuration without credentials:

```sh
npx -y @knowttl/mg-axi profile list
npx -y @knowttl/mg-axi entra user list --help
```

An unconfigured installation reports `profiles: 0 profiles configured` and a next-step help command.
Use `npx -y @knowttl/mg-axi` in place of `mg-axi` throughout these examples, or install globally:

```sh
npm install --global @knowttl/mg-axi
mg-axi setup
```

For Graph reads, an administrator must first provision an organization-owned app registration and consent to the appropriate permissions.
Replace the tenant and client placeholders with that registration's UUIDs:

```sh
mg-axi profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/User.Read.All
mg-axi doctor --profile soc
mg-axi entra user list --profile soc --limit 10 --fields id,displayName
```

Browser login uses MSAL with PKCE.
Authenticate in the browser; never paste credentials into commands or chat.
Profile creation stores identity and credential references, never signs in and never overwrites an existing profile.
See [configuration](docs/configuration.md) for application authentication and device-code opt-in.

The data portion below was captured from the last command using a synthetic directory fixture; your tenant rows will differ.
The CLI also emits a next-step `help` entry.

```toon
users[1]{id,displayName}:
  aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa,Example User
total: null
count: 1 users
complete: true
```

## Use with your agent

Install the [generated skill](skills/mg-axi/SKILL.md), which includes the complete executable command table:

```sh
npx skills add knowttl/mg-axi --skill mg-axi
```

For optional session context in Claude Code, Codex and OpenCode, use the global installation:

```sh
mg-axi setup hooks
```

Restart the agent session.
Hooks inject local profile names, tenant labels, authentication-cache presence, write policy and version; they make no Graph or sign-in calls.
Skills and hooks are independent discovery options and grant no permissions.
`setup` only displays guidance; only `setup hooks` installs integrations.
The hook installer rejects temporary, linked-worktree and validation-build targets.

There is no CLI hook-uninstall command.
To remove the integration, delete only the `mg-axi-hook` SessionStart entries from `~/.claude/settings.json` and `~/.codex/hooks.json`, and `~/.config/opencode/plugins/axi-mg-axi-hook.js`.
Preserve other hooks and shared Codex hook settings.

## Usage

`doctor` performs one bounded user-list read per selected profile with silent credential acquisition.
Without `--profile`, it checks the default profile, or all profiles if no default is set.
Missing configuration and failed reads return errors.
Ordinary commands never initiate interactive login.

## Commands

The [capability report](docs/coverage.md) distinguishes implemented operations from deferred or excluded ones.
The [command table](skills/mg-axi/SKILL.md#orientation) lists every executable leaf; use its leaf-level help for exact flags and permission defaults.
Top-level `--help` lists immediate commands; `entra --help`, `entra user --help` and deeper group help list only their immediate children.
Unknown or incomplete commands return usage exit code 2 with guidance for the nearest valid group.
Group and leaf `--help` return exit code 0 without signing in.

| Area | Shipped surface |
| --- | --- |
| Local configuration | `profile create`, `profile list`, `profile show`, `config path`, `auth login`, `auth logout`, `auth status`, `home`, `setup`, `setup hooks`, `doctor` |
| Directory | Users, groups and memberships, devices, administrative units, directory objects, deleted items, contacts, organizations and branding |
| Applications and consent | Applications, service principals, owners, federated credentials, delegated grants and app-role assignments |
| Roles and governance | Directory roles and assignments, PIM reads, access reviews, entitlement management and lifecycle workflows |
| Security | Conditional Access, authentication methods and registration reports, audit/sign-in/provisioning logs, user and workload risk, Defender XDR advanced hunting (`security hunting query`) |
| Mail (opt-in) | Folder list/show, message list/show/search, attachment list/save, MIME export, settings show, inbox rule list/show, master category list and the named mail writes below |
| Tenant configuration | Domains and DNS records, subscriptions, certificate authentication, synchronization, terms of use, custom security attributes and group lifecycle policies |
| Identity and partner configuration | Identity providers, federation configurations, risk-prevention providers, data policy operations, contracts, delegated administration, multitenant organization and tenant information |
| Reviewed raw reads | `api get <path>` for allowlisted v1.0 routes |
| Named writes | Group member add, account enable/disable, session revocation, Conditional Access policy update, risky-user dismissal, mail folder create, folder delete, category create, rule create, rule update, rule delete, and mail message tag, flag, mark, move, copy and trash |

Delegated read commands default to the scope for their operation, which may differ from the user-read scope in the quick start.
Sign in explicitly with the appropriate scopes before switching areas.
Read commands reject unapproved or write scopes before acquiring credentials.
Permissions, directory roles and licences vary by operation; see the [access and licence guidance](docs/graph-coverage.md).
A 403 response alone does not identify which prerequisite is missing.

```sh
# Group membership reads
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/GroupMember.Read.All
mg-axi entra group list --profile soc --limit 10
mg-axi entra group member list --profile soc --group <group-id>

# Conditional Access inspection
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/Policy.Read.All
mg-axi entra conditional-access policy list --profile soc
mg-axi entra conditional-access policy show --profile soc --id <policy-id>

# Reviewed raw reads
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/User.Read.All,https://graph.microsoft.com/GroupMember.Read.All
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

## Safety and permissions

Read-only by default: seventeen named mutations and two guarded mail downloads can be enabled.
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
| `mail folder create` | `mail.folder.create` | `--name` |
| `mail folder delete` | `mail.folder.delete` | `--ids` in identical order |
| `mail category create` | `mail.category.create` | `--name` |
| `mail rule create` | `mail.rule.create` | file `displayName` |
| `mail rule update` | `mail.rule.update` | `--ids` in identical order |
| `mail rule delete` | `mail.rule.delete` | `--ids` in identical order |
| `mail message tag` | `mail.message.tag` | `--ids` in identical order |
| `mail message flag` | `mail.message.flag` | `--ids` in identical order |
| `mail message mark` | `mail.message.mark` | `--ids` in identical order |
| `mail message move` | `mail.message.move` | `--ids` in identical order |
| `mail message copy` | `mail.message.copy` | `--ids` in identical order |
| `mail message trash` | `mail.message.trash` | `--ids` in identical order |
| `mail attachment save` | `mail.attachment.save` | `--attachment` |
| `mail message export` | `mail.message.export` | `--id` |

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
| Folder create and delete | Delegated preview and fresh reads request `Mail.ReadBasic`; execution requests `Mail.ReadWrite`. Application mode uses scoped Exchange Application `Mail.ReadWrite` with no `Mail.*` application grant in Entra. Creation makes a visible folder at the mailbox root from `--name`; deletion takes 1 to 20 distinct `--ids` for empty custom folders only and refuses well-known folders. See [the generated skill](skills/mg-axi/SKILL.md) for confirm, empty-folder checks and stop semantics. |
| Category create | Delegated preview and fresh reads request `MailboxSettings.Read`; creation requests `MailboxSettings.ReadWrite`. Application mode uses scoped Exchange Application `MailboxSettings.ReadWrite` with no `MailboxSettings.*` application grant in Entra. Creation takes `--name` and `--color`; an existing name with the same color is a no-op and a conflicting color is refused. See [the generated skill](skills/mg-axi/SKILL.md) for bounds and stop semantics. |
| Rule create, update and delete | Delegated preview and fresh reads request `MailboxSettings.Read`; execution requests `MailboxSettings.ReadWrite`. Application mode uses scoped Exchange Application `MailboxSettings.ReadWrite` with no `MailboxSettings.*` application grant in Entra. Create and update take `--file` rule JSON (UTF-8 regular file, at most 64 KiB); forwarding, redirect and permanentDelete actions are refused and read-only rules cannot be changed. `--ids` takes 1 to 20 distinct IDs and `--confirm` repeats the file `displayName` for create or every ID in identical order for update and delete. Inbox only. See [the generated skill](skills/mg-axi/SKILL.md) for reviewed fields and stop semantics. |
| Message tag, flag, mark, move, copy and trash | Delegated preview and fresh reads request `Mail.Read`; execution requests `Mail.ReadWrite`. Application mode uses Exchange Application `Mail.ReadWrite` scoped through the owner-chosen mail-enabled group, with no `Mail.*` application grant in Entra. `--ids` takes 1 to 20 distinct IDs and `--confirm` repeats every ID in identical order. Move and copy take `--destination` in the same mailbox; trash always moves to Deleted Items and never uses DELETE or permanentDelete. Results carry only the original id and the returned newId; a missing newId stops the batch as unknown. See [the generated skill](skills/mg-axi/SKILL.md) for per-command flags and multi-target stop semantics. |

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

| Read journey | Delegated Graph permission |
| --- | --- |
| Users and `doctor` | `User.Read.All` |
| Group membership | `GroupMember.Read.All` |
| Conditional Access policies | `Policy.Read.All` |
| Defender XDR hunting | `ThreatHunting.Read.All` (admin-consented; requires the `security` pack) |
| Reviewed raw reads | Explicit approved `--scopes` for the requested route |

Use full `https://graph.microsoft.com/` scope names with `auth login`.
Application profiles request Graph `.default`; delegated scopes cannot narrow application authorization.
For mail, grant scoped Exchange RBAC roles as described below, with no Entra mail application grants.
The v1.0 group-members route omits service principals.

Credentials live in protected environment/file holders or the delegated MSAL cache in the state directory.
Secret values are never printed; diagnostics may show holder names and paths.
Explicit login contacts Microsoft identity endpoints, and Graph commands send the requested tenant operation to Microsoft Graph.
Local home, setup, help and profile inspection make no tenant calls.
Treat returned directory data, cursors and mutation journals as sensitive tenant data.

## Mail pack

Manually add `"mail"` to the intended profile's `enabledPacks`, for example `"enabledPacks": ["entra", "mail"]`.
Entra remains the default; enabling mail does not enable writes or downloads.
Every mail command requires `--mailbox <user-id|upn|me>`.
Standalone mg-axi supports delegated access to the signed-in user's own mailbox and scoped application access; hosted mail through artium uses application sign-in only.
`me` is delegated only; explicit delegated IDs or UPNs must match authenticated MSAL account metadata.
Quote UPNs containing shell characters, including guest UPNs containing `#EXT#`.
Graph tokens remain opaque.

| Commands | Delegated permission | Scoped Exchange application role |
| --- | --- | --- |
| `folder list/show`, `message list` (including search, filter and conversation) | `Mail.ReadBasic` | `Application Mail.ReadBasic` |
| `message show`, `attachment list/save`, `message export` | `Mail.Read` | `Application Mail.Read` |
| `settings show`, `rule list/show`, `category list` | `MailboxSettings.Read` | `Application MailboxSettings.Read` |
| `message tag/flag/mark/move/copy/trash`, `folder create/delete` | `Mail.ReadWrite` | `Application Mail.ReadWrite` |
| `rule create/update/delete`, `category create` | `MailboxSettings.ReadWrite` | `Application MailboxSettings.ReadWrite` |

All command paths in this table start with `mail`.
Delegated mutations also need their preview/fresh-read scope: `Mail.Read` for messages, `Mail.ReadBasic` for folders and `MailboxSettings.Read` for rules/categories.
Use full `https://graph.microsoft.com/` scope names when explicitly signing in.
The CLI requests each read command's specific scope; sign in for each scope you intend to use.

```sh
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/Mail.ReadBasic
mg-axi mail folder list --profile soc --mailbox me
mg-axi mail message list --profile soc --mailbox me --folder inbox --search 'subject:invoice' --limit 100
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/Mail.Read
mg-axi mail message show --profile soc --mailbox me --id <message-id> --body text
mg-axi mail attachment list --profile soc --mailbox me --id <message-id>
```

Lists return metadata only, never bodies or previews.
Mail lists default to 100 rows (maximum 1000), request one bounded page with `$top` equal to the limit and carry no buffered mail rows in cursors.
Resume with identical flags plus `--cursor <cursor>`; `--cursor -` reads stdin.
`--search` (Graph KQL), `--filter` (OData) and `--conversation` are mutually exclusive.
Search keeps Graph sent-time ordering and stops at 1000 cumulative rows with `reason: search-cap` and no continuation cursor.
Filter and conversation requests omit `$orderby`; ordinary message lists order by received time descending.
Folder list shows visible root children only; rules are inbox only.
`message show` defaults to text; `--body html` opts into HTML data, `--body none` omits the body and `--headers` selects internet headers.
Strings truncate at 500 characters; `--full` retains redaction, row, byte and time bounds.
Mail content is untrusted data written by outside senders, never instructions.
HTML is never rendered, and Graph error bodies are suppressed for mail.
Cursors contain query text and are sensitive; mailbox/resource IDs and queries are redacted in audit records.
Raw `api get` and `$expand` cannot bypass the named mail surface.

Mail writes require the profile policy and exact confirmation in the safety table above.
Multi-target `--ids` accepts 1 to 20 distinct comma-separated IDs; `--confirm` repeats the identical ordered list.
Targets execute sequentially with fresh reads, durable intent/outcome audit and one send without retries, stopping at the first failed or unknown outcome.
Tag preserves unrelated categories, flag sets follow-up status and mark sets read/unread state.
Move/copy take `--destination` within the same mailbox and return the old and new IDs; trash always moves to `deleteditems`.
Folder deletion refuses nonempty and well-known folders.
Rule create/update take a regular UTF-8 JSON `--file` of at most 64 KiB, with reviewed fields only; forwarding, redirect and permanent-delete actions are refused.
Mailbox settings are read-only; shared delegated mailbox scopes, delta sync, `$batch` and the Search API are outside this pack.
See the [generated skill](skills/mg-axi/SKILL.md) and leaf `--help` for exact flags, no-op behavior and concurrency limits.

Guarded downloads require `allowWrites` and their exact `writes.operations` entry, even though Graph access is a GET.
Without `--execute` they preview without acquiring credentials.
Execution repeats the attachment ID or message ID in `--confirm` and requires `--out` to name an absolute new file.
Files are created exclusively, no-follow and 0600, capped at 50 MiB by reported size and streamed bytes; partial files are removed on failure.
Reference attachments are refused, downloads are journaled and never retried, and stdout contains only path, size, SHA-256, content type and kind.

```sh
# Preview only, after manually enabling the exact download operation
mg-axi mail attachment save --profile soc --mailbox me --id <message-id> --attachment <attachment-id> --out /absolute/new/attachment.bin
mg-axi mail message export --profile soc --mailbox me --id <message-id> --out /absolute/new/message.eml
```

### Scoped application setup

An Exchange administrator configures [RBAC for Applications](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac) outside mg-axi.
Use the owner's chosen mail-enabled security group and direct mailbox membership; nested members are out of scope.
Do not create a new Application Access Policy.
Assign only the capability roles needed from the table above, each restricted to the same management scope.
Remove all `Mail.*` and `MailboxSettings.*` application grants in Entra: Entra grants and Exchange RBAC grants are additive, so unscoped Entra grants bypass this scope.
Keep Graph `User.Read.All` application permission for `doctor` and the hosted credential health check, which test user-list access and cannot validate mail authorization.

In Exchange Online PowerShell, use an administrator with Organization Management or the required delegating assignments (Microsoft also documents the Entra Exchange Administrator requirement).
Substitute your app's client ID and the **Enterprise applications service principal object ID**, not the app registration object ID.
The following example grants metadata reads only; choose a different role from the table for other capabilities.

```powershell
Connect-ExchangeOnline
$appId = '<application-client-id>'
$servicePrincipalId = '<enterprise-application-object-id>'
$group = Get-Group -Identity '<owner-chosen-mail-enabled-security-group>'
New-ServicePrincipal -AppId $appId -ObjectId $servicePrincipalId -DisplayName 'Mailbox tool'
New-ManagementScope -Name 'Mailbox tool scope' -RecipientRestrictionFilter "MemberOfGroup -eq '$($group.DistinguishedName)'"
New-ManagementRoleAssignment -Name 'Mailbox tool metadata read' -App $servicePrincipalId -Role 'Application Mail.ReadBasic' -CustomResourceScope 'Mailbox tool scope'
Test-ServicePrincipalAuthorization -Identity $servicePrincipalId -Resource 'allowed@contoso.example'
Test-ServicePrincipalAuthorization -Identity $servicePrincipalId -Resource 'outside@contoso.example'
```

Reuse an existing Exchange service-principal pointer if already registered.
Check `InScope` is true for the intended mailbox and false for an excluded mailbox for every assigned role.
This test excludes Entra grants and bypasses the permission cache; independently verify the absence of Entra mail application grants.
Permission changes may take 30 minutes to 2 hours to reach normal API calls.
Create an application profile using the [protected credential holders](docs/configuration.md#application-credentials-and-headless-use), enable its mail pack and use explicit mailbox IDs or UPNs.
The CLI cannot independently prove the Exchange scope or absence of unscoped Entra grants.
These are administrator setup instructions; repository tests use offline fixture mailboxes and never execute this recipe against a tenant.

## Configuration

The full [configuration reference](docs/configuration.md) covers paths, profile selection, environment profiles, credential storage and headless authentication.
The [AXI credential standard](https://github.com/knowttl/artium-axi/blob/main/docs/credential-standard.md) defines SecretRefs and file protection.

| Flag, environment variable or key | Purpose |
| --- | --- |
| `--config` / `MG_AXI_CONFIG` | Config path; flag wins |
| `--profile` / `MG_AXI_PROFILE` | Profile selection; flag wins |
| `MG_AXI_READ_ONLY=1` | Force read-only regardless of profile policy |
| `MG_AXI_WRITE_LOG` | Mutation journal path override |
| `defaultProfile` | Default file profile when no explicit or env profile is selected |
| `writes.allowWrites`, `writes.operations` | Manual per-profile opt-in and exact operation allowlist |

Run `mg-axi config path` to see the effective config location.
The default is `~/.config/mg-axi/config.json` on Linux/macOS or `%APPDATA%\mg-axi\config.json` on Windows.
State defaults to `~/.local/state/mg-axi/` or `%LOCALAPPDATA%\mg-axi\`.
No secrets belong in profile JSON or argv.

## Development and contributing

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
Pull requests rely on local validation; there is no GitHub test CI.
See [CONTRIBUTING.md](CONTRIBUTING.md) for isolated dependency installation, focused tests and contribution checks.
To use the built checkout binary on PATH, run `npm link --ignore-scripts` from the checkout root.
Otherwise substitute `node dist/bin/mg-axi.js` for `mg-axi` in examples.

The agent skill and capability report are generated from the command catalogue, documentation template and discovery inventory.
The package includes a compact runtime index for shipped named and reviewed raw operations, cached once per process.
The full discovery inventory stays in the repository; `pnpm build` regenerates the runtime index before compilation and `pnpm docs:check` checks it for drift.
After changing their sources, run `corepack pnpm run docs:generate`, then `corepack pnpm run docs:check`.
See the [operation inventory](docs/inventory.md) and [execution contract](docs/execution.md) for coverage and safety boundaries.

## Versioning and releases

During 0.x development, feature changes produce patch releases under the current [release configuration](release-please-config.json).
Minor versions are cut intentionally.
The deprecated top-level `login` and `AZURE_FEDERATED_TOKEN_FILE` aliases remain supported through the next intentionally cut minor release and become eligible for removal after that release.
Patch releases do not expire either alias.

[GitHub releases](https://github.com/knowttl/mg-axi/releases) and [npm versions](https://www.npmjs.com/package/@knowttl/mg-axi?activeTab=versions) have two historical differences:
npm has the initial `0.1.0` but GitHub has no `v0.1.0` release; GitHub has `v0.1.5` but its publish failed, so npm jumps from `0.1.4` to `0.1.6`.
The `0.1.5` npm gap remains unfilled.
Pin an existing npm version when reproducibility matters.

The [release workflow](.github/workflows/release-please.yml) builds, lints, tests and checks generated docs before publishing with npm provenance.
There is no GitHub test CI for pull requests; contributors run the local checks above.

## Security reporting

See [SECURITY.md](SECURITY.md) for reporting instructions.
Never include credentials, tokens, cursors or tenant data in a public issue.

## License

[MIT](LICENSE).

If this project saved you time or helped your workflow, please consider giving it a ⭐ on GitHub to help others find it!
