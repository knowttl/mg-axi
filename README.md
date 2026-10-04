# mg-axi

Agent-facing Microsoft Graph CLI with one shared core and domain packs, read-only by default.
The Entra pack comes first, with phased full coverage and later gated named writes.

CLI-01 provides a local TypeScript/AXI shell, strict command catalogue, leaf help and fast version probes.
AUTH-01 adds versioned dedicated-app delegated profiles and explicit login.
AUTH-02 adds certificate and workload-federated application profiles and a client-credentials service.
Graph execution remains scheduled for later [build slices](docs/build-plan.md); the CLI does not yet acquire application tokens.
No tenant, credentials or network access are required for help or an unconfigured home view.

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
The initial user leaves expose scheduled grammar only; valid invocations exit 1 with their inventory disposition and owning slice.
They do not authenticate or send requests.
Unknown flags, unexpected arguments, missing required values and unsupported combinations exit 2 before execution.
Help and successful local views exit 0.
Data and structured errors use TOON on stdout; diagnostics belong on stderr.
Bare `-v`, `-V` and `--version` print only the package version without importing the catalogue.

Create a delegated profile with your organization-owned public client registration, explicit workforce tenant UUID and explicit commercial cloud:

```sh
mg-axi profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial
mg-axi profile list
mg-axi profile show --profile soc
mg-axi login --profile soc --scopes https://graph.microsoft.com/User.Read
```

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
