---
name: mg-axi
description: Use mg-axi for Microsoft Entra inspection and gated named writes - users, groups, roles, devices, sign-ins, applications, consent grants, risk and Conditional Access through agent-ergonomic TOON output.
user-invocable: false
---

# mg-axi

Agent-ergonomic CLI for Microsoft Graph, read-only by default.
Entra SOC reads through token-efficient TOON output.
The generated [command table](#orientation) lists supported named writes through the gated mutation coordinator; raw API remains read-only and every other mutation is refused.
See [README.md](../../README.md) for write usage, execution gates, identity pinning, permissions, target-role hierarchy and write configuration.

From a built checkout, run commands from the repository root with `node dist/bin/mg-axi.js`.
Throughout this skill and CLI output, substitute that invocation for `mg-axi` unless you have separately installed the executable on PATH.
Run `node dist/bin/mg-axi.js doctor` first.
See [README.md](../../README.md#release) for doctor profile selection, checks and failure behavior.

## Orientation

The exact current leaf registry is `src/catalogue.ts`.
Its capability label is `native` (implemented by an mg-axi handler).
Its effect is `read` for Graph read leaves and the doctor health check, `write` for the gated named writes, or `local` for home, profile, login and setup views.
Doctor acquires credentials silently and contacts Graph; only its help view stays offline.
The list below records current executable leaves; it makes no coverage claim for other Graph operations.
See `docs/coverage.md` for the per-operation disposition records.
`docs/coverage.md` and this table are generated from the same catalogue and inventory sources; `node tools/generate-docs.mjs --check` fails when they are stale.

<!-- command-registry:start -->
| Command | Capability | Effect |
|---|---|---|
| `mg-axi home` | native | local |
| `mg-axi profile create` | native | local |
| `mg-axi profile list` | native | local |
| `mg-axi profile show` | native | local |
| `mg-axi login` | native | local |
| `mg-axi setup` | native | local |
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
| `mg-axi entra role-assignment list` | native | read |
| `mg-axi entra pim eligible list` | native | read |
| `mg-axi entra pim active list` | native | read |
| `mg-axi entra device list` | native | read |
| `mg-axi entra device show` | native | read |
| `mg-axi entra administrative-unit list` | native | read |
| `mg-axi entra administrative-unit show` | native | read |
| `mg-axi entra administrative-unit member list` | native | read |
| `mg-axi entra domain list` | native | read |
| `mg-axi entra domain show` | native | read |
| `mg-axi entra domain verification-dns-record list` | native | read |
| `mg-axi entra domain verification-dns-record show` | native | read |
| `mg-axi entra domain service-configuration-record list` | native | read |
| `mg-axi entra domain service-configuration-record show` | native | read |
| `mg-axi entra domain-dns-record list` | native | read |
| `mg-axi entra domain-dns-record show` | native | read |
| `mg-axi entra sign-in list` | native | read |
| `mg-axi entra sign-in show` | native | read |
| `mg-axi entra directory-audit list` | native | read |
| `mg-axi entra directory-audit show` | native | read |
| `mg-axi entra application list` | native | read |
| `mg-axi entra application show` | native | read |
| `mg-axi entra service-principal list` | native | read |
| `mg-axi entra service-principal show` | native | read |
| `mg-axi entra application owner list` | native | read |
| `mg-axi entra service-principal owner list` | native | read |
| `mg-axi entra service-principal oauth2-grant list` | native | read |
| `mg-axi entra service-principal app-role-assignment list` | native | read |
| `mg-axi entra risky-user list` | native | read |
| `mg-axi entra risky-user show` | native | read |
| `mg-axi entra risky-user dismiss` | native | write |
| `mg-axi entra risk-detection list` | native | read |
| `mg-axi entra risk-detection show` | native | read |
| `mg-axi entra conditional-access policy list` | native | read |
| `mg-axi entra conditional-access policy show` | native | read |
| `mg-axi entra conditional-access policy update` | native | write |
| `mg-axi entra conditional-access named-location list` | native | read |
| `mg-axi entra conditional-access named-location show` | native | read |
| `mg-axi entra access-review definition list` | native | read |
| `mg-axi entra access-review definition show` | native | read |
| `mg-axi entra access-review instance list` | native | read |
| `mg-axi entra access-review instance show` | native | read |
| `mg-axi entra access-review decision list` | native | read |
| `mg-axi api get` | native | read |
<!-- command-registry:end -->

Run `mg-axi <leaf-path> --help` for that leaf's accepted flags and reference.
Unknown flags fail before any credential or HTTP work.

## Setup (explicit only)

Follow [README.md](../../README.md#release) for checkout builds and explicit skill installation.
Create profiles explicitly and keep secrets out of argv and config files:

```sh
node dist/bin/mg-axi.js setup                    # build steps, config path and capabilities; writes nothing
node dist/bin/mg-axi.js profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial
node dist/bin/mg-axi.js login --profile soc --scopes https://graph.microsoft.com/User.Read.All
node dist/bin/mg-axi.js doctor                   # one bounded read per selected profile
node dist/bin/mg-axi.js entra user list --profile soc --limit 10
```

Configuration defaults to `~/.mg-axi/config.json`; `MG_AXI_CONFIG` selects a separate file.
`node dist/bin/mg-axi.js setup` shows the selected path and writes nothing.

## Selecting a profile

Read leaves accept `--profile <name>`.
Selection uses `--profile` or the configured default.
Without either, read leaves report `AUTH_REQUIRED`, even if only one profile exists.
Doctor selection is documented in [README.md](../../README.md#release).
Without profiles, home, setup and profile list show unconfigured state; Graph read leaves report `AUTH_REQUIRED`.

## Safety

Reads run through a session that authorizes only catalogued operations before silent credential acquisition; only explicit `mg-axi login` opens a browser or prints a device-code challenge.
Output is TOON on stdout; diagnostics use stderr.
Text truncates at 500 characters (`--full` restores text, never redaction); row caps resume through opaque cursors.
