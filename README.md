# mg-axi

Agent-facing Microsoft Graph CLI with one shared core and domain packs, read-only by default.
The Entra pack comes first, with phased full coverage and later gated named writes.

CLI-01 provides a local TypeScript/AXI shell, strict command catalogue, leaf help and fast version probes.
Authentication and Graph execution are scheduled for later [build slices](docs/build-plan.md).
No tenant, credentials or network access are required to use this shell.

Use Node 22.12 or newer and the pinned pnpm version:

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

Run `corepack pnpm build`, `corepack pnpm test` and `corepack pnpm lint` for shell validation.
CI builds and tests on Ubuntu and Windows with Node 22 and 24, and validates the existing Python inventory tooling separately.
The [implementation plan](PLAN.md) remains the design authority.
Generated skill, setup and capability reporting ship in PACK-01; no session hooks are installed by ordinary commands.

The pinned [Entra operation inventory](docs/inventory.md) defines the INV-01 discovery boundary and schema for later build slices.
