# Execution contract

## Named writes

Only named, approved pack commands can mutate; raw API remains read-only.
WRITE-00 provides the shared [mutation coordinator](../src/mutations.ts).
WRITE-01 ships the first named command (`entra group member add`); WRITE-02 adds the named `entra user update` account-state command; WRITE-03 adds the named `entra user revoke-sessions` session-revocation action; WRITE-04 adds the named `entra conditional-access policy update` command; WRITE-05 adds the named `entra risky-user dismiss` single-user action.
All five commands use the shared production [mutation transport](../src/mutations.ts); later families still ship independently with their own contracts.
Profile enablement, environment configuration and named-write usage, confirmation, identity pinning, permissions and limitations are documented in [README.md](../README.md).
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
HTTP 4xx responses other than 408 record `FAILED`, except a WRITE-01 duplicate-reference 400 (Graph reporting the membership already exists), which records `NOOP` and returns a no-op instead of a failure; transport failures, HTTP 408 and server errors return `kind: unknown` and record `OUTCOME_UNKNOWN`, with guidance to read back the target before doing anything else.
An accepted 2xx response records `SUCCESS` even if its response body cannot be decoded.
No mutation is automatically retried or replayed; an already reserved intent ID is refused even after an outcome was recorded.
Unknown actions, raw writes, secret-returning endpoints, and unsupported beta writes fail closed.

## Read mechanics and source contracts

[Paging](https://learn.microsoft.com/en-us/graph/paging) follows exact `@odata.nextLink` values and preserves required headers, except for the authentication-method query handling documented in [README.md](../README.md).
Validate HTTPS origin and version on initial requests, redirects and every continuation before adding credentials.
An output cap cannot discard the remainder of a fetched page; continuation state preserves query/context and buffered rows if necessary.
CORE-02 exposes these mechanics through `GraphSession.collect`, with `limit`, `budget` and an opaque `cursor`.
EXT-01 (domains) maps domain, verification-record, service-record and top-level DNS-record list flags onto it, EXT-01 (certificate auth) maps PKI-configuration and certificate-authority list flags onto it, EXT-01 (organization) maps organization and branding-localization list flags onto it, EXT-01 (directory subscriptions) maps subscription-list flags onto it, EXT-01 (on-premises synchronization) maps on-premises-synchronization-list flags onto it, EXT-01 (terms of use) maps agreement and agreement-acceptance list flags onto it, EXT-01 (directory objects) maps directory-object-list flags onto it, EXT-01 (contacts) maps contact-list flags onto it, EXT-01 (group lifecycle) maps lifecycle-policy and setting-template list flags onto it, EXT-01 (custom security attributes) maps attribute-set, definition and allowed-value list flags onto it,
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
| WRITE-01 Membership | [POST /groups/{id}/members/$ref](https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0) | Shipped; supported targets and permission requirements are documented in [README.md](../README.md). |
| WRITE-02 Account state | [PATCH /users/{id}](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0) | Shipped; account-state usage, permissions and sensitive-target role hierarchy are documented in [README.md](../README.md). |
| WRITE-03 Revoke sessions | [POST /users/{id}/revokeSignInSessions](https://learn.microsoft.com/en-us/graph/api/user-revokesigninsessions?view=graph-rest-1.0) | Shipped; usage, permissions, limitations and unknown-outcome handling are documented in [README.md](../README.md); implementation: [entra-user-revoke-sessions](../src/entra-user-revoke-sessions.ts). |
| WRITE-04 CA policy update | [PATCH /identity/conditionalAccess/policies/{id}](https://learn.microsoft.com/en-us/graph/api/conditionalaccesspolicy-update?view=graph-rest-1.0) | Shipped; usage, permissions, lockout gates and concurrency limitations are documented in [README.md](../README.md). |
| WRITE-05 Risk dismissal | [POST /identityProtection/riskyUsers/dismiss](https://learn.microsoft.com/en-us/graph/api/riskyuser-dismiss?view=graph-rest-1.0) | Shipped; single-user dismissal usage, permissions and P2/role guidance are documented in [README.md](../README.md). |


Each write ships independently after the shared coordinator and its own contract pass offline checks.
The table describes permission choices, not automatic consent or authority.
Raw writes remain denied even after named writes ship.

The WRITE-01 implementation in [entra-group-member-add](../src/entra-group-member-add.ts) classifies membership adds as disruptive; command usage, enablement and permission requirements are owned by [README.md](../README.md).
The $ref body carries exactly `{"@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/<user-id>"}`; both identifiers must be object IDs.
The user is verified through `GET:/users/{user-id}` with `$select=id` before preview or no-op detection and again before sending; failed or malformed user reads block the operation.
Group `isAssignableToRole` must be explicitly false or null; true, missing and malformed values are refused.
Desired state is read through `GET:/groups/{group-id}/members` in the preview and rechecked before the single send; an incomplete member window proceeds to the POST where a duplicate 400 lands as a no-op.

The WRITE-04 implementation in [entra-ca-policy-update](../src/entra-ca-policy-update.ts) supplies a READ-03 policy-show reread callback to the coordinator, reapplying lockout gates on each fresh effective policy before the single send.
Its reviewed fields, diff and no-op behavior, confirmation, lockout requirements, permissions and concurrency limitations are owned by [README.md](../README.md).
