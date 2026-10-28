# Write gates and permissions

Contents: [Gates](#write-gates-and-permissions) • [Read permissions](#read-permissions)

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
| Folder create and delete | Delegated preview and fresh reads request `Mail.ReadBasic`; execution requests `Mail.ReadWrite`. Application mode uses scoped Exchange Application `Mail.ReadWrite` with no `Mail.*` application grant in Entra. Creation makes a visible folder at the mailbox root from `--name`; deletion takes 1 to 20 distinct `--ids` for empty custom folders only and refuses well-known folders. See [the generated skill](../skills/mg-axi/SKILL.md) for confirm, empty-folder checks and stop semantics. |
| Category create | Delegated preview and fresh reads request `MailboxSettings.Read`; creation requests `MailboxSettings.ReadWrite`. Application mode uses scoped Exchange Application `MailboxSettings.ReadWrite` with no `MailboxSettings.*` application grant in Entra. Creation takes `--name` and `--color`; an existing name with the same color is a no-op and a conflicting color is refused. See [the generated skill](../skills/mg-axi/SKILL.md) for bounds and stop semantics. |
| Rule create, update and delete | Delegated preview and fresh reads request `MailboxSettings.Read`; execution requests `MailboxSettings.ReadWrite`. Application mode uses scoped Exchange Application `MailboxSettings.ReadWrite` with no `MailboxSettings.*` application grant in Entra. Create and update take `--file` rule JSON (UTF-8 regular file, at most 64 KiB); forwarding, redirect and permanentDelete actions are refused and read-only rules cannot be changed. `--ids` takes 1 to 20 distinct IDs and `--confirm` repeats the file `displayName` for create or every ID in identical order for update and delete. Inbox only. See [the generated skill](../skills/mg-axi/SKILL.md) for reviewed fields and stop semantics. |
| Message tag, flag, mark, move, copy and trash | Delegated preview and fresh reads request `Mail.Read`; execution requests `Mail.ReadWrite`. Application mode uses Exchange Application `Mail.ReadWrite` scoped through the owner-chosen mail-enabled group, with no `Mail.*` application grant in Entra. `--ids` takes 1 to 20 distinct IDs and `--confirm` repeats every ID in identical order. Move and copy take `--destination` in the same mailbox; trash always moves to Deleted Items and never uses DELETE or permanentDelete. Results carry only the original id and the returned newId; a missing newId stops the batch as unknown. See [the generated skill](../skills/mg-axi/SKILL.md) for per-command flags and multi-target stop semantics. |

Conditional Access enforcement changes require a readable lockout analysis.
Execution requires `--acknowledge-lockout-risk` only when the preview reports elevated lockout risk, after verifying emergency-access coverage.
An all-users block policy covering browser and modern clients on all apps or admin portals without recognized narrowing conditions is refused even with acknowledgement.
User, group and role exclusions or application filters alone cannot establish emergency-access protection.
Disabled and report-only policies do not enforce access; display-name-only changes do not require lockout acknowledgement.
See the [policy implementation](../src/entra-ca-policy-update.ts) for the precise static-analysis boundary.

Execution journals intent durably before sending and records its outcome.
A failed intent record blocks the send; journal locks and incomplete records require manual reconciliation.
Writes are not automatically retried or replayed.
After `OUTCOME_UNKNOWN` or an outcome-recording failure, read back the target before further action and never replay that intent.
Account updates and risk dismissal verify the resulting state; a mismatch reports `WRITE_CONFLICT` and a failed verification read reports `OUTCOME_UNKNOWN`.
After a Conditional Access update, explicitly read the policy back.
Fresh reads do not make writes atomic, and account and policy updates send no ETag precondition or promise of concurrency protection.
Do not assume rollback.
The [execution contract](execution.md#named-writes) describes the shared guarantees.

## Read permissions

| Read journey | Delegated Graph permission |
| --- | --- |
| Users and `doctor` | `User.Read.All` |
| Group membership | `GroupMember.Read.All` |
| Conditional Access policies | `Policy.Read.All` |
| Defender XDR hunting | `ThreatHunting.Read.All` (admin-consented; requires the `security` pack) |
| Reviewed raw reads | Explicit approved `--scopes` for the requested route |

Use full `https://graph.microsoft.com/` scope names with `auth login`.
Application profiles request Graph `.default`; delegated scopes cannot narrow application authorization.
For mail, grant [scoped Exchange RBAC roles](mail.md#scoped-application-setup), with no Entra mail application grants.
The v1.0 group-members route omits service principals.

Credentials live in protected environment/file holders or the delegated MSAL cache in the state directory.
Secret values are never printed; diagnostics may show holder names and paths.
Explicit login contacts Microsoft identity endpoints, and Graph commands send the requested tenant operation to Microsoft Graph.
Local home, setup, help and profile inspection make no tenant calls.
Treat returned directory data, cursors and mutation journals as sensitive tenant data.
