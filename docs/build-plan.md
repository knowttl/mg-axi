# Dispatch plan

Status: approved planning baseline, 2026-10-03.
Repository: `knowttl/mg-axi`; base: `main`.
Implementation is authorized for a fresh GPT-6.1 Sol crew at medium effort.
This table defines sequencing and acceptance; it does not claim that a slice has shipped.
Use small feature branches and reviewable changes.
Every slice uses offline fixtures; no real tenant, real Graph call, or interactive sign-in is a validation requirement.

A dependency means its public contract is merged and usable.
A phase-3 family is a dispatch template that must be instantiated from the completed inventory before work starts, not an unbounded umbrella assignment.
No implementation crew should be asked to deliver all of phase 3 in one change.

| ID | Deliverable | Depends on | Independent acceptance |
|---|---|---|---|
| INV-01 | Pin Graph metadata, define Entra boundary and machine-readable operation dispositions | This plan | Every discovered scoped operation has version, cloud, auth-mode, permission/role/licence sources, disposition and owning slice; no unsupported completeness percentage. |
| CLI-01 | TypeScript/AXI package shell, strict catalogue, leaf help and fast version | INV-01 schema | Invalid flags/combinations fail before adapters; home/help/TOON/exit contracts; no tenant required. |
| AUTH-01 | Profiles and dedicated-app delegated authentication per the [authentication contract](../PLAN.md#authentication-app-registration-and-profiles) | CLI-01 | Explicit tenant/client/cloud; separated credential references; approved browser/device-code selection contract; no prompt outside login; safe cache and expiry/error behavior through credential seam. |
| AUTH-02 | Certificate and workload-federated application profiles | AUTH-01 | Client credentials and configured Graph .default audience; no user/device-code fallback; consent guidance; mode-incompatible operations rejected. |
| CORE-01 | Policy-enforced Graph session and fixture HTTP adapter | AUTH-02 | Validate host/version/path/redirects before credentials; pack and sensitive-area denial; translate errors and redact sentinel values. |
| CORE-02 | Collections, query options, bounded retries and cancellation | CORE-01 | Exact nextLink/header handling; cap preserves unreturned rows; truthful partial results; Retry-After and deadline through fake clock. |
| READ-01 | Users list/show and initial vertical packaged CLI journey | CORE-02 | Basic/richer properties, null/missing/denied distinction, both authentication modes, opaque tokens, stdout/stderr/exit contract. |
| READ-02 | Groups and direct/transitive relationships | READ-01 | Hidden and limited-information members; role-assignable distinctions; documented service-principal v1.0 limitation; no silent beta. |
| READ-03 | CA policies and named locations | READ-01 | Separate policy/location grammar, role and consent errors, P1/P2 guidance, safe policy projection. |
| READ-04 | Targeted authentication methods and registration reporting | READ-01 | No tenant scan through per-user methods; disabled-user report gap explicit; protected fields redacted. |
| READ-05 | Sign-ins and directory audit logs | READ-01 | Bounded time/filter queries, paging, unavailable premium/CA fields, per-operation roles and licensing guidance. |
| READ-06 | Risky users and risk detections | READ-05 | P1/P2 detail boundaries, sign-in correlation, no invented riskySignIns endpoint. |
| READ-07 | Applications/service principals and credential metadata | READ-01 | appId distinct from object ID, safe owners/relationships, expiry metadata, no secrets. |
| READ-08 | Delegated/application consent grants | READ-07 | Actual grants distinct from requested permissions; supported read scopes, no write-consent requests for reads. |
| READ-09 | Directory roles and active/eligible PIM | READ-02 | Active/eligible and direct/activated states distinct; role and filter requirements; P2/Governance documentation. |
| READ-10 | Devices and administrative units | READ-01 | Directory-device scope distinct from Intune; AU scope/roles/licensing and paging. |
| API-01 | Reviewed read-only raw Graph surface | CORE-02 | GET-only, reviewed route/query/field access, no secret endpoints, no pack bypass; mail/files disabled unless explicitly enabled. |
| PACK-01 | Read release, setup/doctor, generated docs and installable skill | READ-02 through READ-10, API-01 | Packaged critical journeys offline; no noisy auto-login/install; capability report generated from catalogue; no claim of full Entra yet. |
| EXT-01 | Tenant/directory policies, domains, licences, attributes and remaining directory relationships | PACK-01 | One reviewed subfamily per change from INV-01; per-operation query/access fixtures and sourced coverage updates. |
| EXT-02 | Governance: access reviews, entitlements, lifecycle and group PIM | PACK-01 | Split by inventory subfamily; permissions, roles and P2/Governance feature distinctions; instance versus schedule semantics. |
| EXT-02c | Access-review history reads (declined, out of v1) | EXT-02 | Declined by the captain to keep mg-axi strictly read-only: the history reads carry the ReadWrite least-privilege scope AccessReview.ReadWrite.All (no read scope) and history instances return SAS download URLs. No commands or raw access. |
| EXT-03 | External identities, cross-tenant controls and provisioning | PACK-01 | Split by subfamily; workforce context only initially; no silent external-customer support claim. |
| EXT-04 | Workload/agent identity, remaining risk/reporting, network access and partner operations | PACK-01 | Split by subfamily and version; explicit beta isolation, product-specific licensing and partner prerequisites; unsupported operations visible. |
| FULL-01 | Full Entra coverage audit | EXT-01 through EXT-04 and all inventory-created read slices | Every agreed read has a named command or an explicit reviewed blocked/unavailable/deprecated/deferred disposition; the deferred tail is declared out of v1, so scheduled no longer blocks the milestone. |
| WRITE-00 | Mutation coordinator and execution gates, fixture-only enablement | PACK-01 | Forced read-only, hand-enabled profile, immutable tenant/operation scope, preview, execute, confirmation, audit intent/outcome, no automatic replay. |
| WRITE-01 | Add user group membership | WRITE-00, READ-02 | Initially supported non-role-assignable groups; correct $ref target; no-op/duplicate behavior and permission/confirmation checks. |
| WRITE-02 | Enable/disable account | WRITE-00, READ-01 | Documented permission pair and target-role hierarchy; desired-state preview; conflict/reread behavior. |
| WRITE-03 | Revoke sessions | WRITE-00, READ-01 | Action preview, delayed/external-user limitations, target confirmation, timeout outcome unknown, no invented rollback. |
| WRITE-04 | Update CA policy | WRITE-00, READ-03 | Explicit reviewed fields; lockout analysis; no unsupported concurrency promise; operation stays disabled if required protection is unavailable. |
| WRITE-05 | Dismiss risky user | WRITE-00, READ-06 | Single explicit user initially; P2 and permission/role guidance; dismissal distinct from remediation. |
| WRITE-N | Remaining agreed Entra mutation families | WRITE-00, corresponding named reads, inventory contract | One operation family per change; sourced permissions, preview, target, concurrency, unknown-outcome and audit tests before enablement. |
| COMPLETE-01 | Full agreed Entra capability audit including later writes | FULL-01, all agreed WRITE slices | Five shipped writes plus deferred WRITE-N, with no merely scheduled agreed operation left unlabelled; publish separate named/raw/blocked/deprecated/unavailable/deferred counts and outstanding upstream limitations. |

PACK-01 is a useful SOC-focused read release, not the full-Entra completion milestone.
Later writes may proceed after PACK-01 alongside extended reads without waiting for every long-tail family.
Existing scope covers the full Entra map; the five concrete writes are the first slices, not the entire promised write inventory.
The inventory may reveal unsupported or dangerous operations that need an explicit disposition rather than an invented command.
A new genuine product choice discovered during implementation must be raised before deciding it.
Sovereign clouds, external-customer launch support and named Security/Intune/M365 packs remain separately authorized extensions.

## FULL-01 audit record

FULL-01 audit snapshot (2026-10-07, historical point-in-time, latest origin/main at that time): v1.0 named-command 289, reviewed-raw-read 5, scheduled 0, deferred 4429 (out of v1), intentionally-blocked 57, deprecated 16, unavailable 14; beta scheduled 3717 (never agreed for v1). Every agreed read had a named command or a reviewed disposition, so the milestone was met at that time. See docs/coverage.md for current counts.

## COMPLETE-01 audit record

COMPLETE-01 audit snapshot (2026-10-07, historical point-in-time, latest origin/main at that time): five writes WRITE-01..05 ship as named, gated commands through the WRITE-00 mutation coordinator (v1.0 named-command 289: 284 reads plus 5 writes; reviewed-raw-read 5; scheduled 0); WRITE-N carries deferred 1878, intentionally-blocked 25 v1.0 plus 2937 beta, deprecated 6 v1.0 plus 398 beta, scheduled 0; v1.0 deferred 4429 (out of v1), intentionally-blocked 57, deprecated 16, unavailable 14; beta scheduled 3717 (never agreed for v1). No agreed v1 operation, read or write, is left merely scheduled, and the declined EXT-02c access-review history reads stay out. The milestone was met at that time. See docs/coverage.md for current counts and outstanding upstream limitations.

## Per-slice handoff

Each implementation handoff contains its slice ID, prerequisite commit(s), exact operation inventory rows, in-scope commands, accepted API versions/auth modes/clouds, documented access/licence constraints, and acceptance behavior.
Use existing repository conventions.
Do not add a separate fixture system, generic endpoint wrapper, dynamic plugin framework, or abstraction merely to facilitate a test.
Revalidate cited upstream contracts at implementation time and record the check date.
Use the repository's validation/shipping gate before publishing each change.
