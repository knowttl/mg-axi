# Verification plan for the v1.0 read fixes

Run all checks without Microsoft credentials or live Graph requests.
The regression cases use recorded request shapes with the existing fake transport and credential services.
Run executable tests through `tools/isolated-home.mjs` so HOME and XDG directories are temporary.

1. Run `pnpm build` and `pnpm lint` to verify compilation and types.
2. Run `node tools/isolated-home.mjs node --test --test-concurrency=4 test/graph-session.test.mjs test/graph-collections.test.mjs test/entra-roles.test.mjs test/entra-conditional-access.test.mjs test/entra-group-lifecycle.test.mjs test/entra-audit-logs.test.mjs test/entra-provisioning.test.mjs test/entra-domains.test.mjs test/entra-directory-objects.test.mjs test/api.test.mjs`.
3. Run `pnpm docs:check` to verify generated command help and coverage documentation.

The six regression cases in `test/graph-session.test.mjs` verify:

- A: both single and collection PIM requests send a supported culture and expose the real licence error.
- B: scoped membership reads request `roleMemberInfo` and `administrativeUnitId`, and direct root membership queries are refused.
- C: authentication-context lists, authentication-strength detail, and lifecycle-policy lists omit `$select` while preserving compact output.
- D: unsupported root DNS-record and directory-object collections fail before credentials and offer supported alternatives.
- E: provisioning lists preserve their time filter, omit `$select`, and return `provisioningAction`.
- F: a Graph query error with an empty code retains its message.

For D, keep the legacy named list commands as explicit unsupported errors and remove the root routes from the raw allowlist.
There is no equivalent global list to map to: DNS records require a parent domain, and directory objects have documented detail and typed list operations.
The alternatives already exist as named commands.
This follows the [domainDnsRecord reference](https://learn.microsoft.com/en-us/graph/api/resources/domaindnsrecord?view=graph-rest-1.0) and [directoryObject methods](https://learn.microsoft.com/en-us/graph/api/resources/directoryobject?view=graph-rest-1.0).

Delivery checks:

- Use gh-axi for GitHub operations, and include `--repo knowttl/mg-axi -B main` on every PR creation.
- The PR body must state the problem, consequence, fix and scope, including the choice for D described above.
- Keep all validation offline with the existing fakes; do not change GitHub Actions.
- Confirm the PR is not a draft and is mergeable when checks pass; merging and the release PR belong to the supervising delivery process.
