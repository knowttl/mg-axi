# mg-axi

Agent-facing Microsoft Graph CLI with one shared core and domain packs, read-only by default.
The Entra pack comes first, with phased full coverage and later gated named writes.

CLI-01 provides a local TypeScript/AXI shell, strict command catalogue, leaf help and fast version probes.
AUTH-01 adds versioned dedicated-app delegated profiles and explicit login.
Graph execution and application authentication remain scheduled for later [build slices](docs/build-plan.md).
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

Create a profile with your organization-owned public client registration, explicit workforce tenant UUID and explicit commercial cloud:

```sh
mg-axi profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial
mg-axi profile list
mg-axi profile show --profile soc
mg-axi login --profile soc --scopes https://graph.microsoft.com/User.Read
```

Configure the registration's Mobile and desktop applications redirect URI as `http://localhost` for browser login.
The first created profile is the default; `--profile` selects another identity explicitly.
Configuration defaults to `~/.mg-axi/config.json`; `MG_AXI_CONFIG` selects a separate configuration file.
Version 1 stores tenant/client/cloud, delegated mode, enabled packs, preview/sensitive policy and a unique credential reference only.
Unknown fields, unsupported versions/clouds and application credentials are rejected with recovery guidance.
Creation never overwrites an existing profile.
Preview is disabled and sensitive areas are empty in newly created profiles.

Browser login uses Microsoft's [MSAL interactive API](https://learn.microsoft.com/en-us/entra/msal/javascript/node/acquire-token-requests) and PKCE.
Request delegated Graph scope names explicitly with `--scopes`, separated by commas.
The dedicated registration needs corresponding delegated consent, and operations can additionally require user roles.
Login never changes app registration or requests blanket directory write consent automatically.
Device code requires creating the profile with `--allow-device-code` and selecting `--method device-code` during login, only when organization policy permits it.
Browser failure never falls back to device code.
Only explicit login can open a browser or display the device challenge on stderr.
Ordinary credential acquisition uses silent refresh and returns actionable errors when login, consent or policy intervention is required.

MSAL caches are held by the OS credential store through optional `keytar` when its native module and store are available.
Otherwise they remain in process memory, and login reports `storage: session-only`; authentication then lasts only for that process.
There is no plaintext credential-cache fallback.
Installing with `--ignore-scripts` can leave the native store unavailable, as can a headless system without a credential service.
Tokens remain opaque and never appear in profile views, stdout or authentication diagnostics.
The credential service binds account context to the configured tenant/client and refreshes at a 60-second expiry margin.
Tests use fake credential providers and fake time; they never sign in or contact a token endpoint.

Run `corepack pnpm build`, `corepack pnpm test` and `corepack pnpm lint` for shell validation.
The [CI workflow](.github/workflows/ci.yml) defines the platform/runtime matrix for shell build, test and lint checks, and validates the Python inventory tooling separately.
The [implementation plan](PLAN.md) remains the design authority.
Generated skill, setup and capability reporting ship in PACK-01; no session hooks are installed by ordinary commands.

The pinned [Entra operation inventory](docs/inventory.md) defines the INV-01 discovery boundary and schema for later build slices.
