# Execution contract

## Named writes

Only named, approved pack commands can mutate; raw API remains read-only.
WRITE-00 implements the [mutation coordinator](../src/mutations.ts) for fixture-driven execution; named mutation commands and production mutation transport remain unavailable.
Profile enablement and environment configuration are documented in [README.md](../README.md).
The coordinator snapshots the validated profile's tenant, identity and approved-operation scope at creation; later profile edits cannot widen that scope.
It checks forced read-only and profile enablement before preview, and the sender checks forced read-only again before credential acquisition and immediately before transport handoff.
Preview redacts current state and proposed payload using the shared Graph redaction boundary.
Omitting explicit execution returns a preview without sending or journaling; verified already-desired state is a no-op.
Tenant/object confirmation applies to security-impacting mutations and deletion, including access changes through POST/PATCH and membership references.
The coordinator requires exact target confirmation for definitions classified as disruptive.
Every execute rechecks state and endpoint-supported preconditions.
There is no assumption of universal Graph dry-run, what-if, ETag or rollback.
If server concurrency protection is unavailable, the race is stated and high-impact operations stay disabled until their individual contract is reviewed.
Audit intent is persisted before sending; failure prevents send.
The journal contains redacted metadata, never payloads, headers or response bodies, and each record is fsynced before reporting.
The journal directory is created before intent reservation; directory fsync is skipped on native Windows, where the journal file is pre-created and fsynced instead.
After intent reservation, a failed fresh read, a newly satisfied desired state or a sender gate failure records `NOT_SENT`.
Outcome audit failure after sending preserves the possibility that the mutation occurred.
HTTP 4xx responses other than 408 record `FAILED`; transport failures, HTTP 408 and server errors return `kind: unknown` and record `OUTCOME_UNKNOWN`, with guidance to read back the target before doing anything else.
An accepted 2xx response records `SUCCESS` even if its response body cannot be decoded.
No mutation is automatically retried or replayed; an already reserved intent ID is refused even after an outcome was recorded.
Unknown actions, raw writes, secret-returning endpoints, and unsupported beta writes fail closed.

## Read mechanics and source contracts

[Paging](https://learn.microsoft.com/en-us/graph/paging) follows exact `@odata.nextLink` values and preserves required headers, except for the authentication-method query handling documented in [README.md](../README.md).
Validate HTTPS origin and version on initial requests, redirects and every continuation before adding credentials.
An output cap cannot discard the remainder of a fetched page; continuation state preserves query/context and buffered rows if necessary.
CORE-02 exposes these mechanics through `GraphSession.collect`, with `limit`, `budget` and an opaque `cursor`.
API-01 uses that interface for raw collection reads, READ-01 maps user-list CLI flags onto it, READ-02 maps group, member and memberOf list flags onto it, READ-03 maps Conditional Access policy and named-location list flags onto it, READ-04 maps per-user authentication-method and registration-report list flags onto it, READ-05 maps bounded sign-in and directory-audit list flags onto it, READ-06 maps risky-user list flags and bounded risk-detection list flags onto it, READ-07 maps application, service-principal and owner list flags onto it, READ-08 maps service-principal delegated-grant and app-role-assignment list flags onto it, READ-09 maps directory-role, role-assignment and PIM active/eligible list flags onto it, and READ-10 maps device, administrative-unit and AU-member list flags onto it; [README.md](../README.md) owns CLI options, output and resume usage.
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
