# mg-axi

Agent-facing Microsoft Graph CLI, read-only by default

[![npm](https://img.shields.io/npm/v/@knowttl/mg-axi)](https://www.npmjs.com/package/@knowttl/mg-axi) [![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE) [![Node.js](https://img.shields.io/badge/node-%3E%3D22.12.0-green)](package.json) [![Provenance](https://img.shields.io/badge/npm-provenance-blue)](https://www.npmjs.com/package/@knowttl/mg-axi?activeTab=provenance) [![AXI](https://img.shields.io/badge/AXI-Agent%20eXperience%20Interface-green)](https://github.com/kunchenguid/axi)

[Why](#why-this-exists) • [Quickstart](#quickstart) • [Agent setup](#use-with-your-agent) • [Coverage](#what-it-covers) • [Safety](#safety) • [Configuration](#configuration) • [Output](#output-and-exit-codes) • [Troubleshooting](#troubleshooting) • [Contributing](#contributing) • [Security](#security) • [License](#license)

## Why this exists

Raw Graph APIs leave agents guessing routes, permissions, paging and response size.
mg-axi offers reviewed Entra commands, bounded TOON output and contextual next steps.
Reads distinguish empty results from incomplete or denied requests; named writes require explicit gates and target confirmation.
It serves workforce Entra tenants in the commercial cloud, with opt-in Defender XDR hunting and mailbox packs.

## Quickstart

Requires Node.js **22.12.0 or later**.
Install globally, then inspect the home view:

```sh
npm install --global @knowttl/mg-axi
mg-axi
# Or without installing: npx -y @knowttl/mg-axi
```

Configured home returns bounded live Graph samples for the selected profile: up to five users and five root mail folders, one request and a five-second deadline per supported enabled pack, without paging or retries.
Delegated reads use existing silent `User.ReadBasic.All` and `Mail.ReadBasic` access; application Entra reads use the existing `.default` audience.
No permissions are added and no interactive sign-in occurs; unavailable consent, credentials or Graph access are reported per pack rather than as zero results.
Security requires an explicit hunting query, and application mail requires an explicit mailbox, so home reports these as unsupported automatic summaries with discovery hints.
Use `mg-axi home --offline` for local configuration without credentials or Graph calls.
Session hooks remain offline and installation stays explicit.

An unconfigured home view includes this synthetic excerpt:

```toon
profile: "unavailable: no profile configured"
```

An administrator must provision an organization-owned app registration and consent to Graph `User.Read.All`.
Replace the placeholders with that registration's UUIDs, sign in through the browser, then read users:

```sh
mg-axi profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/User.Read.All
mg-axi entra user list --profile soc --limit 10 --fields id,displayName
```

See [configuration](docs/configuration.md) for application authentication and device-code opt-in.

## Use with your agent

Install session context from a global CLI installation, or install the on-demand skill:

```sh
mg-axi setup hooks
npx skills add knowttl/mg-axi --skill mg-axi
```

Restart the agent session after hook installation.
Hooks inject local profile, authentication-cache and write-policy context without Graph calls; the skill supplies command guidance.
Either option works independently and grants no permissions.
Start with the [compact investigation guide](docs/agent-guide.md); the [generated skill](skills/mg-axi/SKILL.md) is the full command reference.
There is no CLI uninstall command: remove only `mg-axi-hook` SessionStart entries from `~/.claude/settings.json` and `~/.codex/hooks.json`, and delete `~/.config/opencode/plugins/axi-mg-axi-hook.js`.
Preserve other hooks and shared settings; see [hook details](docs/agent-setup.md).

<a id="features"></a>
<a id="commands"></a>

## What it covers

| Group or pack | Purpose | Access | Reference |
| --- | --- | --- | --- |
| Setup and home | Profiles, authentication, home, setup and health checks | Configured home and doctor read Graph; `home --offline` stays local | [Commands](docs/commands.md) |
| Entra directory | Users, groups, memberships, devices and deleted objects | Read; gated member/account writes | [Coverage](docs/graph-coverage.md) |
| Applications and consent | Apps, service principals, owners and grants | Read | [Commands](docs/commands.md) |
| Roles and governance | Roles, PIM, access reviews, entitlement and lifecycle | Read | [Coverage](docs/coverage.md) |
| Entra security | Conditional Access, methods, logs and risk | Read; gated policy/risk/session writes | [Permissions](docs/writes.md) |
| Defender XDR | Restricted file-based KQL hunting | Read; opt-in security pack | [Hunting](docs/security-hunting.md) |
| Mail | Messages, folders, attachments, settings, rules and categories | Read; gated writes/downloads; opt-in | [Mail](docs/mail.md) |
| Tenant and partner configuration | Domains, policies, federation and partner relationships | Read | [Commands](docs/commands.md) |
| Reviewed raw reads | Allowlisted v1.0 `api get <path>` routes | Read | [Output and paging](docs/output-and-paging.md) |

Use `mg-axi <group> --help` or leaf `--help` for exact flags without signing in.
The full generated leaf table lives in [docs/commands.md](docs/commands.md) and [SKILL.md](skills/mg-axi/SKILL.md#orientation).
Permissions, roles and licensing vary by operation; a 403 alone does not identify the missing prerequisite.

<a id="safety-and-permissions"></a>

## Safety

Read-only by default.
Seventeen named mutations and two guarded mail downloads require manual profile opt-in and an exact `writes.operations` allowlist.
Writes preview without sending; execution requires `--execute` and `--confirm` repeating the target.
`MG_AXI_READ_ONLY=1` overrides profile opt-in; there is no allow-write environment gate.
Raw writes and beta writes are blocked.
Credentials use protected environment/file holders or the delegated MSAL cache; never put secret values in argv or profile JSON.
Treat output, cursors and journals as sensitive tenant data, and mail content as untrusted data.
After an unknown outcome, read back the target before further action and never replay the intent.
See the [write gates and permission matrix](docs/writes.md), [execution contract](docs/execution.md) and [read permissions and licences](docs/graph-coverage.md).

## Configuration

Run `mg-axi config path` to see the effective path; `--config` wins over `MG_AXI_CONFIG`.
Profiles select explicit tenant/client identities; `--profile` wins over `MG_AXI_PROFILE`, then env profile, default file profile or sole file profile.
See [configuration](docs/configuration.md) for schema, credential references and precedence.

| Location | Linux/macOS default | Windows default |
| --- | --- | --- |
| Config | `~/.config/mg-axi/config.json` | `%APPDATA%\mg-axi\config.json` |
| State | `~/.local/state/mg-axi/` | `%LOCALAPPDATA%\mg-axi\` |

| Environment variable | Purpose | Default |
| --- | --- | --- |
| `MG_AXI_CONFIG`, `MG_AXI_PROFILE` | Config override; profile selector | Paths/selection above |
| `MG_AXI_READ_ONLY` | Force read-only (`1` or `true`) | `false` |
| `MG_AXI_WRITE_LOG` | Journal path | State directory's `writes.log` |
| `MG_AXI_MODE`, `MG_AXI_TENANT_ID`, `MG_AXI_CLIENT_ID` | Required env-profile identity | Unset |
| `MG_AXI_CLOUD` | Cloud | `commercial` |
| `MG_AXI_AUTH` | Application credential provider | Required for application env profiles |
| `MG_AXI_CLIENT_SECRET`, `MG_AXI_PRIVATE_KEY`, `MG_AXI_TOKEN` | Credential holders | Unset |
| `MG_AXI_THUMBPRINT_SHA256` | Certificate thumbprint | Required for certificate env profiles |
| `MG_AXI_ENABLED_PACKS` | JSON pack array | `["entra"]` |
| `MG_AXI_SENSITIVE_AREAS` | JSON array; only empty is supported | `[]` |
| `MG_AXI_WRITES` | JSON write policy | Disabled |
| `MG_AXI_PREVIEW`, `MG_AXI_ALLOW_DEVICE_CODE` | Profile booleans | `false` |
| `AZURE_FEDERATED_TOKEN_FILE` | Deprecated assertion-file fallback | Unset |
| `XDG_CONFIG_HOME`, `XDG_STATE_HOME` | Absolute Unix base directories | `~/.config`, `~/.local/state` |
| `APPDATA`, `LOCALAPPDATA` | Windows base directories | Home's AppData Roaming/Local |
| `HOME`, `USERPROFILE`, `PATH` | Skill-home discovery and executable lookup | Process environment |
| `CLAUDE_CONFIG_DIR`, `CODEX_HOME`, `OPENCODE_CONFIG_DIR` | Skill-discovery locations | [Discovery defaults](docs/agent-setup.md) |
| `CLAUDECODE`, `CODEX_THREAD_ID`, `CODEX_SESSION_ID`, `PI_CODING_AGENT`, `OPENCODE` | Agent detection by presence | Unset |

<a id="toon-output-behavior"></a>

## Output and exit codes

Structured output, errors and response diagnostics use TOON on stdout.
See [response diagnostics](docs/output-and-paging.md#response-diagnostics) for malformed-response evidence and handling.
Strings truncate at 500 characters; `--full` removes text truncation while retaining redaction and bounds.
Lists report completeness; use the returned cursor to resume partial results where supported.
`--all` follows pages within budgets rather than promising an unlimited result.
See [output and paging](docs/output-and-paging.md) for projection, limits and cursor handling.

| Code | Meaning |
| --- | --- |
| 0 | Success, including no-ops and help |
| 1 | Operational failure or incomplete hunting result |
| 2 | Usage or validation error |

<a id="usage"></a>

## Troubleshooting

| Symptom | Next step |
| --- | --- |
| `PROFILE_REQUIRED` | Create a profile with the Quickstart commands |
| HTTP 403 | Check operation scopes, roles, licensing and consent in [permissions](docs/graph-coverage.md) |
| Failed health check | `mg-axi doctor --profile soc` silently tests one bounded user-list read; see [doctor](docs/output-and-paging.md#doctor) |
| `OUTCOME_UNKNOWN` | Read back the target and reconcile the journal; never replay |
| Hook installation refused | Install globally outside temporary directories and linked worktrees |

<a id="mail-pack"></a>
<a id="scoped-application-setup"></a>

The mail guide contains [pack usage](docs/mail.md) and [scoped application setup](docs/mail.md#scoped-application-setup).

<a id="development-and-contributing"></a>
<a id="versioning-and-releases"></a>

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for offline checks, generated documentation and [release mechanics](CONTRIBUTING.md#versioning-and-releases).

## Security

Report vulnerabilities using [SECURITY.md](SECURITY.md).

## License

[MIT](LICENSE).
