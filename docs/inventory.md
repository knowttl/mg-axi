# INV-01 discovery contract

The machine-readable [inventory](../inventory/operations.json) records the pinned Microsoft Graph OpenAPI revision.
Both `openapi/v1.0/openapi.yaml` and `openapi/beta/openapi.yaml` are inputs.
The inventory records their exact byte sizes and SHA-256 hashes, along with the upstream repository, revision and check date.
These are discovery inputs, not runtime compatibility guarantees or permission grants.
The [JSON Schema](../inventory/schema.json) defines the public discovery format.
The discovery schema records every discovered operation, including explicit excluded rows.
The [offline discovery tool](../tools/inventory.py) owns boundary classification and deterministic generation.
Generated operation rows must be regenerated rather than edited manually.

## Boundary

The initial context is commercial-cloud workforce tenants.
Sovereign-cloud and external-customer support require separate authorization.
Each row's `cloud` and `tenantType` identify that selected context, while `cloudAvailability` records the separate unresolved operation-level availability check.
A commercial context is not a claim that every discovered route is available there.

The boundary follows the approved [Entra map](https://learn.microsoft.com/en-us/graph/api/resources/identity-network-access-overview?view=graph-rest-1.0), checked on 2026-10-03.
It includes directory objects, users, groups and identity relationships; applications and service principals; requested and granted permissions; policies and authentication; tenant configuration, domains and licences; directory and group PIM; governance; external identities and provisioning; identity risk and reporting; agent identity; network access; and partner tenant administration.
Legacy root aliases, alternate-key paths, casts, actions, count routes and `/me` identity routes remain separate discovered operations.
`ROOT_SLICES`, `USER_NAV`, `GROUP_NAV`, `MANAGED_TENANT_NAV`, `EXCLUDED_NAV`, action sets and the small exceptional branches in `scoped_slice` are the exact executable boundary.
New upstream roots or navigation properties require an explicit boundary review.

Mail, calendar, files, Teams, contacts belonging to a user's mailbox, Intune, Security-pack endpoints and other separately authorized domains are excluded from the scoped Entra surface.
Organizational `/contacts` are directory contacts and remain included.
Users and groups have explicit identity navigation/action allowlists because their nested routes also expose other products.
Mailbox settings are excluded at every descendant depth, as are M365 insight settings and Intune authority changes.
`EXCLUDED_NAV` stops cross-pack descendant acceptance before family ownership is assigned and supplies the out-of-pack domain reason.
Managed-tenant navigation is limited to identity coverage, audit and partner tenant administration; Cloud PC, device compliance, Windows protection and general management templates are excluded.
Directory roles under `/roleManagement/directory` and entitlement roles under `/roleManagement/entitlementManagement` are included; Intune, Exchange and Cloud PC role-management branches are excluded.
Entra registration/authentication/identity reports are included; M365 usage and Intune reports are excluded.
Provisioning schema aliases `/filterOperators` and `/functions` are included alongside application and service-principal synchronization routes.
Each input records the total number of discovered operations and excluded counts by root, making the boundary's exclusions visible without calling them implemented capabilities.
Each discovered version/method/path has exactly one row, including excluded operations with explicit out-of-pack or unscoped reasons.
The validator checks that row totals match discovery totals, exclusion counts match excluded rows, and cross-pack operations cannot have scoped dispositions.
The inventory describes the pinned metadata and this explicit boundary, rather than promising that metadata contains every product feature.

## Row identity and evidence

`id` is the unique `version:METHOD:path` key.
`operationId` is preserved from upstream, but is not used as a cross-version unique key.
`version`, `method` and `path` preserve the exact discovered route.
There is no beta fallback and no inferred stable counterpart for a beta-only route.
Absence from one metadata file is not sufficient evidence of runtime unavailability.

Every row records independent delegated and application `authModes`, plus `permissions`, `roles`, `licences` and `cloudAvailability` evidence records.
An evidence record has `status: not-reviewed` and one or more primary-source URLs.
When upstream supplies `externalDocs`, its exact operation-documentation URL is retained as `documentation` and used as the access-review source.
When it does not, `documentation` is null and a documented read action uses its `READ_ACTION_SOURCES` reference; other operations use the Entra overview as an explicit discovery fallback, not operation-level access evidence.
Licence review also points at the Entra licensing overview.
No permission names, administrator roles, licence entitlement or auth-mode support are inferred from OpenAPI discovery.
For scoped operations, these references identify where the owning slice must verify access, including premium properties and national-cloud tables, before implementing an operation.
The existing [coverage research](graph-coverage.md) provides useful representative contracts, but cannot safely be applied to every route in a family.
An unreviewed value is never interpreted as unsupported, unrestricted, Free or supported.

This discovery schema deliberately permits only unresolved evidence.
API-01's [implementation catalogue](../src/api.ts) owns reviewed raw routes, query keys, fields and sourced access constraints.
Later slices must extend the implementation catalogue's reviewed access contract explicitly, with sourced supported/unsupported auth modes, exact permission choices and role/licence constraints.
The discovery inventory remains reproducible from the pinned metadata; shipped-operation dispositions are joined from the implementation catalogue at generation time.

## Dispositions and ownership

The disposition names are `named-command`, `reviewed-raw-read`, `scheduled`, `deferred`, `intentionally-blocked`, `deprecated`, `unavailable` and `excluded`.
Every row has a nonempty `reason`; scoped rows have an `owningSlice` from the dispatch plan.
An `excluded` row has `owningSlice: null`, carries no dispatch authorization, and remains visible solely for discovery accounting.
`scheduled` means the operation is assigned to a later implementation slice, not that the operation works or its access is verified.
Only beta operations may remain scheduled: beta was never agreed for v1, so `scheduled` no longer blocks the v1 milestone.
`deferred` means the operation is explicitly out of v1 by a reviewed cutoff decision; the reason names that decision.
Extended slice IDs are family dispatch templates and must be split into the approved small subfamily changes.
`WRITE-N` is likewise a dispatch template rather than authorization to implement all mutations.
The five initial write routes have their specific write owner where identified.
Operation-level ownership does not authorize all writable properties on a PATCH route.

Credential values, secret minting, LAPS credential detail and BitLocker recovery-key detail are intentionally blocked.
Metadata collection/count routes remain scheduled for a future reviewed safe projection.
Trust-framework key-set surfaces remain blocked because keys and secret-bearing operations require explicit safety contracts.
Beta writes and external-customer-only user-flow surfaces remain intentionally blocked.
Documented lookup, membership-check, evaluation and validation POST actions keep their family owner and are not denied as beta writes.
`READ_ACTION_SOURCES` records their operation-effect evidence, shared across route aliases and both disposition and ownership decisions.
Only explicitly listed POST actions are classified as reads; unknown POST actions remain writes, regardless of name prefixes.
Multicloud permissions management and upstream-deprecated operations are recorded as deprecated.
These route dispositions do not replace later field/query redaction or the Graph session's policy checks.
For example, a scheduled application or user GET is not permission to project credential values from it.

Rows claim `named-command` only when the row matches a shipped operation in the named-command catalogue (`src/catalogue.ts`) or an alternate route a shipped leaf reaches through cli flag routing (`--transitive`/`--as`, from the `TRANSITIVE_OPERATION`, `NAV_CASTS`, `MEMBERSHIP_TRANSITIVE` and `MEMBERSHIP_CASTS` tables, attributed to the owning leaf), and `reviewed-raw-read` only when the row matches a reviewed raw route in the raw-route catalogue (`src/api.ts`); a shipped operation or table alternate that also matches a reviewed raw route is `named-command`; the generator derives those dispositions from the catalogues and routing tables and never from a hand list. Every table alternate must resolve through table keys to a catalogue operation; one with no shipped command behind it fails generation.
The validator rejects unbacked claims, rejects any shipped catalogue operation or table-backed alternate whose disposition is not `named-command`, rejects any reviewed raw route outside backed whose disposition is not `reviewed-raw-read`, and rejects a missing inventory row for any shipped operation, reviewed raw route or table-backed alternate.
`unavailable` records an explicitly sourced version/cloud limitation or a reviewed decision that the selected version has no documented operation contract.
It is never inferred merely from a missing metadata route.
See the generated [capability report](coverage.md) for operation-level dispositions and their sourced reasons.
There is no completeness percentage.
Future capability reports must count named, raw, deferred, scheduled, blocked, deprecated and unavailable scoped operations separately and cannot count deferred or scheduled work as complete.
Excluded rows must be counted separately from scoped capabilities.

## Offline verification and refresh

Build tooling uses Python 3, PyYAML and jsonschema, independently of the TypeScript runtime shell.
Install the pinned tooling dependencies with `python3 -m pip install -r tools/requirements.txt` in your development environment.
The repository's validation commands need no sign-in, credentials or tenant traffic:

```sh
python3 tools/inventory.py validate
python3 -m unittest discover -s tools -p 'test_*.py'
```

For reproduction, retrieve the two pinned upstream metadata files into `.cache/inventory/v1.0.yaml` and `.cache/inventory/beta.yaml`.
Their source URLs are `https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/<revision>/openapi/<version>/openapi.yaml`, substituting the recorded revision and version.
Then run:

```sh
python3 tools/inventory.py check
```

`check` compares the entire regenerated artifact, including source hashes, operation identities, exclusions and dispositions.
It fails on missing operations, changed source bytes, boundary drift or a changed generated row.
`generate` writes the same artifact for an explicitly reviewed refresh.
Generation streams individual path definitions and does not retain Graph's full response schemas.
Rows are sorted by identity and stored one per line to make operation-level diffs local.
Tests use synthetic metadata files and the real discovery/validation interface, including quoted paths, both versions, cross-pack exclusions, blocked routes and rejected coverage claims.

A metadata refresh changes the pinned revision deliberately, retrieves both corresponding files, regenerates the inventory, and reviews all added/removed routes and boundary exclusions.
Refreshing metadata does not enable commands, request consent or enable writes.
Before dispatching an implementation slice, select exact row identities and reverify their current operation documentation, access and licensing contracts.

The az-axi reference was rechecked at `origin/main` revision `0c84d29fa2aac7f7aed7ee5e19b7d4ef5bb415fc` on 2026-10-03.
Its [registry](https://github.com/knowttl/az-axi/blob/0c84d29fa2aac7f7aed7ee5e19b7d4ef5bb415fc/src/lib/registry.ts) distinguishes implemented leaves, reviewed raw API coverage and blocked capabilities.
INV-01 preserves that operation-level distinction without introducing CLI commands or copying ARM policy.
