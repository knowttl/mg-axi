# Design evidence and provenance

Research inspected the actual az-axi source, rather than treating its public command shape as proof of implementation.
`git -C az-axi rev-parse HEAD origin/main` returned `cefed201ad46e343025138859821e3ec368301f9` for both refs.
The parity scout independently verified GitHub main at the same commit.
The parity review recommends az-like noun/subnoun/verb grammar and strict leaf flag schemas.
The links below preserve the independently reviewable code evidence.

| Actual code | Finding and implication |
|---|---|
| [auth.ts:34](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/auth.ts#L34), [auth.ts:74](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/auth.ts#L74) | Resolves az or environment-token credentials, with resource-specific audience and cache; Graph uses `az account get-access-token --resource-type ms-graph`. |
| [client.ts:72](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/client.ts#L72) | Exact HTTPS host checks and policy/gates precede token acquisition and fetch. |
| [client.ts:161](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/client.ts#L161) | One retry for 429/503 with a short supported Retry-After; Graph needs its own bounded read-retry contract. |
| [client.ts:204](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/client.ts#L204) | Generic list paging follows ARM `nextLink`, not Graph `@odata.nextLink`; not reusable unchanged. |
| [policy.ts:126](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/policy.ts#L126) | Classifies method/path; exact query POST allowlist; credential action list. |
| [policy.ts:148](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/policy.ts#L148) | Graph writes are categorically blocked. Entra writes are new policy, not an existing az-axi feature. |
| [gates.ts:68](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/gates.ts#L68), [config.ts:153](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/config.ts#L153) | Environment override, hand-enabled profile, original configured subscription allowlist, explicit execute, destructive confirmation; read overrides cannot broaden write scope. |
| [dryRun.ts:127](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/dryRun.ts#L127) | Current-state probe, diff, ARM locks and deployment what-if; Graph needs endpoint-specific preview rules. |
| [execute.ts:19](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/execute.ts#L19) | Reread, no-op, If-Match, ARM async handling, request IDs, outcome audit. |
| [registry.ts:22](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/registry.ts#L22), [redact.ts:27](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/lib/redact.ts#L27) | Lazy command modules, effect enforcement, central output redaction. |
| [help.ts:4](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/src/help.ts#L4), [PLAN.md:117](https://github.com/knowttl/az-axi/blob/cefed201ad46e343025138859821e3ec368301f9/PLAN.md#L117) | Curated current surface differs from proposed az-like grammar; the plan explicitly excludes Graph writes from az-axi. |

The shared-core recommendation concerns cross-repository extraction only.
The shared Graph core across mg-axi packs is already required.
The deletion test favors a deep Graph session and mutation coordinator because removing them spreads complex ordering and error handling across callers.
A pass-through wrapper around every endpoint would be shallow and is not proposed.
Existing az-axi's body-only request convenience and independent effect/global state patterns are not automatically copied; their relevance must be assessed at the chosen seam.
No refactoring of az-axi is authorized by this planning task.

CLI-01 rechecked az-axi `origin/main` at `17125324b00cc97e0445d1a7eb170553f707ce9c` on 2026-10-03.
Its package, leaf registry/parser, executable, no-mistakes configuration and CI informed the shell conventions.
mg-axi uses the AXI SDK fast-path export and retains its own scheduled INV-01 operation identities without copying ARM authentication or endpoint policy.
