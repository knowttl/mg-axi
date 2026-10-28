# Configuration

Contents: [Delegated credential storage](#delegated-credential-storage) • [Application credentials and headless use](#application-credentials-and-headless-use)

Profiles store tenant/client/cloud identity, authentication mode, policy and credential references, never secret values.
Only commercial-cloud workforce profiles are supported.
Unknown configuration fields, unsupported versions/clouds and inline credentials are rejected.

| Variable | Meaning |
| --- | --- |
| `MG_AXI_CONFIG` | Configuration file override; `--config` takes precedence |
| `MG_AXI_PROFILE` | Profile selection override; `--profile` takes precedence |
| `MG_AXI_READ_ONLY` | `1` or `true` refuses named writes; `0`, `false` or unset allows profile policy; other values fail |
| `MG_AXI_WRITE_LOG` | Nonblank mutation journal path override; defaults to the state directory's `writes.log` |

On Linux and macOS, configuration lives at `$XDG_CONFIG_HOME/mg-axi/config.json` (default `~/.config/mg-axi/config.json`) and state at `$XDG_STATE_HOME/mg-axi/` (default `~/.local/state/mg-axi/`).
Only absolute XDG directory overrides are used.
On Windows, XDG variables are ignored; configuration uses `%APPDATA%\mg-axi\config.json` and state uses `%LOCALAPPDATA%\mg-axi\`.
Run `mg-axi config path` to see the selected configuration path.
Legacy `~/.mg-axi/config.json` is never read; when it exists and the selected new config does not, `CONFIG_MOVED` reports both paths and the exact move command.

Profile selection is `--profile`, `MG_AXI_PROFILE`, a complete env profile, `defaultProfile`, then the only configured profile.
The first profile created becomes the default.
An env profile requires `MG_AXI_MODE`, `MG_AXI_TENANT_ID`, `MG_AXI_CLIENT_ID` and its credential fields, plus `MG_AXI_AUTH` for application mode.
For example, client-secret auth requires `MG_AXI_CLIENT_SECRET`; certificate auth requires `MG_AXI_PRIVATE_KEY` and `MG_AXI_THUMBPRINT_SHA256`; federation requires `MG_AXI_TOKEN` containing the assertion.
Delegated env profiles require `MG_AXI_TOKEN` containing a serialized MSAL cache.
Optional `MG_AXI_ENABLED_PACKS`, `MG_AXI_SENSITIVE_AREAS` and `MG_AXI_WRITES` use JSON values; `MG_AXI_PREVIEW` and `MG_AXI_ALLOW_DEVICE_CODE` use the same boolean forms as `MG_AXI_READ_ONLY`.
A partial env profile fails with `CONFIG_INVALID` instead of falling through to a file profile.
Explicitly selected file profiles ignore the env profile fields and resolve only their own SecretRefs.

The profile schema is shipped as `axi-credential.schema.json`.
mg-axi retains `mode`; application profiles add `auth: "certificate" | "client-secret" | "federated"` and the corresponding `privateKey`, `clientSecret` or `token` SecretRef.
Delegated profiles created by the CLI use `token: {"stored": true}` for their MSAL cache.
Other SecretRefs are `{"env": "VARIABLE_NAME"}` and `{"file": "/absolute/local/path"}`.
Omitted secret fields default to their `MG_AXI_<FIELD>` environment reference.
Old `credentialRef` objects and SHA-1 `thumbprint` fields are rejected; update them to the flattened fields and 64-hex-digit `thumbprintSha256` after moving a legacy file.

### Delegated credential storage

Browser login is the default and uses MSAL with PKCE.
Use full Graph scope names separated by commas; delegated login rejects `.default`.
Device code requires profile creation with `--allow-device-code` and login with `--method device-code`, subject to organization policy.
Browser failure never falls back to device code.
Ordinary reads acquire credentials silently and report when login, consent or policy intervention is required.
Login does not modify the app registration or automatically request blanket write consent.

Explicit `auth login` stores a protected MSAL cache in the state directory's `credentials.json`, reported as `storage: protected-file`.
Windows and Linux use the same backend, with no keyring dependency or 2,560-byte cache limit.
Each entry is keyed by profile name and bound to its tenant and client IDs.
Changing that binding requires a new explicit login.

Each profile holds one account.
Every explicit login wipes that profile's protected and in-memory cache before acquiring replacement credentials and fails if the wipe cannot be verified.
Use separate profiles for separate identities; restore file access before retrying a failed wipe.
Only explicit login persists credentials; silent refresh keeps refreshed credentials in memory.
Access tokens are excluded from the persisted MSAL cache.
`auth logout` deletes the selected profile's cache and is idempotent; MSAL exposes no upstream token revocation API, so organizational revocation remains separate.
`auth status` reports the selected identity, Graph endpoint, secret source names or paths, and local availability without signing in or making a live request.
The old top-level `login` alias remains supported through the next intentionally cut minor release and becomes eligible for removal after that release; patch releases do not expire it.
See [Versioning and releases](../CONTRIBUTING.md#versioning-and-releases).
Tokens remain opaque and never appear in profile views, stdout or authentication diagnostics.
Credential acquisition is bound to the configured tenant/client and refreshes at a 60-second expiry margin.

### Application credentials and headless use

<details>
<summary>Application providers, protected files and rotation</summary>

Application profiles use client credentials without a signed-in user and request the configured Graph `.default` audience.
An administrator grants the required application authorization; delegated `--scopes` cannot narrow an application token.
Entra operations use Graph application permissions on the app registration.
Mail uses scoped Exchange RBAC for Applications with no `Mail.*` or `MailboxSettings.*` application grants in Entra; see the [mail setup recipe](mail.md#scoped-application-setup).
Application profiles reject device-code opt-in and `login`, and acquisition failures never fall back to user authentication.

Choose exactly one credential provider:

```sh
# Certificate: register the matching public certificate on the app registration
mg-axi profile create --name daemon --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --certificate-thumbprint-sha256 <64-hex-digit-thumbprint>

# Workload federation: configure the app registration's federated credential mapping
mg-axi profile create --name batch --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --federated

# Headless client secret: supply the value through the named environment variable
mg-axi profile create --name svc --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --client-secret-env MG_AXI_CLIENT_SECRET
mg-axi profile show --profile svc
```

By default, certificate keys come from `privateKey: {"env": "MG_AXI_PRIVATE_KEY"}`.
Profile creation records the reference and thumbprint; it does not import a key and there is no CLI key-import command.
Provision keys through protected storage tooling, never argv or profile JSON.
Certificate acquisition fails if the referenced key is unavailable.

For headless certificate holders, add `--certificate-key-file <path>` or `--certificate-key-env <variable-name>`.
Client secrets similarly accept `--client-secret-file <path>` instead of `--client-secret-env`.
All credential files, including federated assertions, must use absolute local paths.
On Unix, they must be regular files owned by the current user with no group or other permission bits; 0400 and 0600 are accepted.
Configuration must not be writable by group or other.
On Windows, writes apply a protected DACL for the current user, SYSTEM and Administrators; reads reject reparse points, and `auth status` and `doctor` verify the DACL.
UNC and device paths are rejected.
Missing, empty or exposed holders fail closed without disclosing values.
PEM environment values may contain escaped `\n`; passphrase-protected keys are unsupported.
Prefer certificates or federation over bearer client secrets, and rotate secrets on a schedule.

Federation reads the assertion file on every acquisition so projected tokens can rotate.
Use `--federated-token-file <absolute-path>` for a projected assertion file, or set the default `MG_AXI_TOKEN` assertion value through the environment.
When the token reference is the default `MG_AXI_TOKEN` env holder and that variable is unset, `AZURE_FEDERATED_TOKEN_FILE` stays accepted as a deprecated alias with a stderr notice.
It remains supported through the next intentionally cut minor release and becomes eligible for removal after that release; patch releases do not expire it.
See [Versioning and releases](../CONTRIBUTING.md#versioning-and-releases).
`MG_AXI_TOKEN` takes precedence and the alias must name an absolute file path.
References, rather than values, are stored in configuration; rotation updates the environment or file without changing the profile.
Application credentials remain in process memory and are reacquired at a 60-second expiry margin.
See [headless credential guidance](headless-credentials.md) for holder examples.

</details>

[Back to the README](../README.md#configuration).
