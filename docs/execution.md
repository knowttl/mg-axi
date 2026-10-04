# Execution contract

## Named writes

Only named, approved pack commands can mutate; raw API remains read-only.
The proposed coordinator checks forced read-only environment, hand-enabled profile, original tenant and approved-operation scope, preview versus explicit execution, and target confirmation.
Tenant/object confirmation applies to security-impacting mutations and deletion, including access changes through POST/PATCH and membership references.
Every execute rechecks state and endpoint-supported preconditions.
There is no assumption of universal Graph dry-run, what-if, ETag or rollback.
If server concurrency protection is unavailable, the race is stated and high-impact operations stay disabled until their individual contract is reviewed.
Audit intent is persisted before sending; failure prevents send.
Outcome audit failure after sending preserves the possibility that the mutation occurred.
An ambiguous timeout is `outcome: unknown` with a verification command, never automatic mutation replay.
Unknown actions, raw writes, secret-returning endpoints, and unsupported beta writes fail closed.

## Read mechanics and source contracts

[Paging](https://learn.microsoft.com/en-us/graph/paging) follows exact `@odata.nextLink` values and preserves required headers.
Validate HTTPS origin and version on initial requests, redirects and every continuation before adding credentials.
An output cap cannot discard the remainder of a fetched page; continuation state preserves query/context and buffered rows if necessary.
CORE-02 exposes these mechanics through `GraphSession.collect`, with `limit`, `budget` and an opaque `cursor`.
API-01 uses that interface for raw collection reads, READ-01 maps user-list CLI flags onto it, READ-02 maps group, member and memberOf list flags onto it, READ-03 maps Conditional Access policy and named-location list flags onto it, READ-05 maps bounded sign-in and directory-audit list flags onto it, READ-06 maps risky-user list flags and bounded risk-detection list flags onto it, READ-07 maps application, service-principal and owner list flags onto it, READ-08 maps service-principal delegated-grant and app-role-assignment list flags onto it, READ-09 maps directory-role, role-assignment and PIM active/eligible list flags onto it, and READ-10 maps device, administrative-unit and AU-member list flags onto it; [README.md](../README.md) owns CLI options, output and resume usage.
Results contain `value`, the effective `query`, `complete`, `requests` and `bytes`; a row limit that leaves buffered rows or another page, request/byte/deadline exhaustion, oversized throttle waits and continuation cycles return `complete: false` with a `reason` and cursor.
Resume with the same operation, resource bindings, authentication mode, tenant/client/cloud, credential reference and scope set.
The cursor restores omitted query entries and the consistency level; conflicting explicit arguments fail validation.
Before returning buffered rows or fetching more pages, the session reacquires credentials and checks the original delegated account or application identity; a refreshed token for the same identity is accepted.
[Advanced query](https://learn.microsoft.com/en-us/graph/aad-advanced-queries) support differs per endpoint; declare supported combinations and required eventual-consistency headers.
The session requires `consistencyLevel: "eventual"` for `$search` or `$count=true`, preserving `ConsistencyLevel` across pages, redirects and cursor resumes.
`$expand` accepts explicit relationship paths without nested query options, rejects sensitive relationships, and cannot accompany `$search` or `$count=true`.
The shared query allowlist is defined in [the session implementation](../src/graph-session.ts); it does not establish endpoint-specific support.
[Throttling guidance](https://learn.microsoft.com/en-us/graph/throttling) requires respecting Retry-After and bounded backoff where absent.
Safe reads can retry within deadline; exhausted budget reports the cause and completeness.
Both `execute` and `collect` retry 429/503 responses using Retry-After seconds or HTTP dates, or bounded exponential backoff when the header is absent or invalid.
Their request, wait and deadline ceilings are defined in [the session implementation](../src/graph-session.ts); `collect` accepts positive integer request, response-body byte and deadline budgets per invocation.
The deadline includes credential acquisition and in-flight requests, and `signal` cancellation rejects the call rather than returning a partial result.
The injected `clock` controls deadlines and waits for offline verification in [the collection tests](../test/graph-collections.test.mjs).
Use request IDs and structured errors, never raw dependency noise or secrets.
[v1.0](https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0) is GA; [beta](https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-beta) can break and is not recommended as a production dependency.


## Initial later-write operations

Slice IDs and dependencies are defined in the [dispatch plan](build-plan.md).

| Slice | Operation and permission | Material constraints |
|---|---|---|
| WRITE-01 Membership | [POST /groups/{id}/members/$ref](https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0), D/A GroupMember.ReadWrite.All for user members | Start with supported non-role-assignable groups; other member types have additional permissions; role-assignable groups add RoleManagement.ReadWrite.Directory and applicable roles. |
| WRITE-02 Account state | [PATCH /users/{id}](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0), D/A User.EnableDisableAccount.All + User.Read.All for accountEnabled | Sensitive-target role hierarchy also applies, including documented app-role requirements for app-only sensitive updates. |
| WRITE-03 Revoke sessions | [POST /users/{id}/revokeSignInSessions](https://learn.microsoft.com/en-us/graph/api/user-revokesigninsessions?view=graph-rest-1.0), D/A User.RevokeSessions.All | Preview the action, do not invent a state diff or promise immediate universal session termination; no automatic replay of ambiguous requests. |
| WRITE-04 CA policy update | [PATCH /identity/conditionalAccess/policies/{id}](https://learn.microsoft.com/en-us/graph/api/conditionalaccesspolicy-update?view=graph-rest-1.0), D/A Policy.Read.All + Policy.ReadWrite.ConditionalAccess | P1, P2 for risk-based features; delegated administrator role; review lockout risk and concurrency limitations. |
| WRITE-05 Risk dismissal | [POST /identityProtection/riskyUsers/dismiss](https://learn.microsoft.com/en-us/graph/api/riskyuser-dismiss?view=graph-rest-1.0), D/A IdentityRiskyUser.ReadWrite.All | P2; initially one explicit user; confirmation targets the user, not the collection action; dismissal is not remediation. |


Each write ships independently after the shared coordinator and its own contract pass offline checks.
The table describes permission choices, not automatic consent or authority.
Raw writes remain denied even after named writes ship.
