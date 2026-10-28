# Contributing

Use Node.js 22.12.0 or later and the pinned pnpm 10.34.6.
From the checkout root, install dependencies and run the local checks:

```sh
node tools/isolated-home.mjs corepack pnpm install --frozen-lockfile --ignore-scripts --config.confirm-modules-purge=false
corepack pnpm build
corepack pnpm lint
corepack pnpm test
corepack pnpm run docs:check
```

Tests use synthetic credentials and fake transports, without signing in or contacting real tenants.
The test command runs in a temporary home, strips credential and agent-session environment markers and limits concurrency to four files.
For a focused test, build first, then run `node tools/isolated-home.mjs node --test test/isolated-home.test.mjs` (substitute the relevant test file).
Never use real credentials, services or tenant data for validation.

Before opening a pull request, run all four checks above and describe the behavior changed and the validation performed.
Pull requests rely on local validation; there is no GitHub test CI.
The release workflow is separate and does not replace these checks.

The agent skill, command reference, capability report and investigation guide are generated.
After changing the command catalogue, documentation template (including `src/docs.ts` for the investigation guide) or discovery inventory, run `corepack pnpm run docs:generate`, then `corepack pnpm run docs:check`.
Do not edit generated files or `CHANGELOG.md` by hand.
Keep changes focused and add regression tests for observable behavior.

## Documentation ownership

README.md owns the overview, installation and first actions.
The command catalogue generates skills/mg-axi/SKILL.md and docs/commands.md as the complete leaf references, and `src/docs.ts` generates docs/agent-guide.md as the compact investigation entry point.
Detailed authorization, paging and pack behavior live in docs/.

## Versioning and releases

During 0.x development, feature changes produce patch releases under the current [release configuration](release-please-config.json).
Minor versions are cut intentionally.
The deprecated top-level `login` and `AZURE_FEDERATED_TOKEN_FILE` aliases remain supported through the next intentionally cut minor release and become eligible for removal after that release.
Patch releases do not expire either alias.
The [release workflow](.github/workflows/release-please.yml) builds, lints, tests and checks generated docs before publishing with npm provenance.
