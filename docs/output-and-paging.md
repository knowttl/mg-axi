# Output and paging

## Doctor

`doctor` performs one bounded user-list read per selected profile with silent credential acquisition.
Without `--profile`, it checks the default profile, or all profiles if no default is set.
Missing configuration and failed reads return errors.
Ordinary commands never initiate interactive login.

Delegated read commands default to the scope for their operation, which may differ from the user-read scope in the quick start.
Sign in explicitly with the appropriate scopes before switching areas.
Read commands reject unapproved or write scopes before acquiring credentials.
Permissions, directory roles and licences vary by operation; see the [access and licence guidance](graph-coverage.md).
A 403 response alone does not identify which prerequisite is missing.

```sh
# Group membership reads
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/GroupMember.Read.All
mg-axi entra group list --profile soc --limit 10
mg-axi entra group member list --profile soc --group <group-id>

# Conditional Access inspection
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/Policy.Read.All
mg-axi entra conditional-access policy list --profile soc
mg-axi entra conditional-access policy show --profile soc --id <policy-id>

# Reviewed raw reads
mg-axi auth login --profile soc --scopes https://graph.microsoft.com/User.Read.All,https://graph.microsoft.com/GroupMember.Read.All
mg-axi api get /users --profile soc --scopes https://graph.microsoft.com/User.Read.All
mg-axi api get /groups --profile soc --odata '$filter=securityEnabled eq true&$top=5' --scopes https://graph.microsoft.com/GroupMember.Read.All
```

Raw access is not an unrestricted Graph client.
Unreviewed routes, secret-returning routes, mail/file content, beta raw reads and raw writes are refused before credentials.
Returned data is restricted to reviewed fields; completeness describes the requested collection, not full Graph property coverage.
Secret material, authentication-method phone numbers and data-policy storage locations are redacted, including in full output, previews and audit metadata.

## Response diagnostics

Structured output, errors and response diagnostics use TOON on stdout; progress and deprecation notices use stderr.
Empty or non-JSON read bodies, malformed collection envelopes and non-numeric scalar count bodies fail with `code: RESPONSE_INVALID` (exit 1).
Detection correlation instead reports sign-in-stage failures as incomplete results under its [correlation contract](#risk-detection-correlation).
Read responses with status 204 or 205 return no body without this error.
Malformed command-specific data can still fail as `GRAPH_ERROR`.
Treat malformed responses as unknown, never as empty results.

Response-invalid errors include `status`, `contentType`, `shape` and `shapeTruncated`.
Status is the HTTP status or `null` when invalid; content type is the bounded media type without parameters, or `null` when missing or invalid.
Shape entries report paths and received types, plus array lengths; arrays are sampled at their first element.
Traversal is bounded, and `shapeTruncated` indicates omitted structure.
Collection-envelope failures also include bounded `issues` and `issuePaths`, with received and expected types and an `issuePathsTruncated` indicator.
Response values, raw bodies and validator messages are never printed in these diagnostics.
Only conservative identifier keys can appear in paths; other keys become `<key>` and secret redaction still applies.
Mail, hunting and custom-security-attribute routes replace every key name with `<key>`, including issue paths.
Use this evidence to check API compatibility or proxy behavior.

Accepted mutation responses retain their separate [responseUnreadable warning contract](execution.md#named-writes); do not replay a mutation because its response could not be decoded.

## TOON output behavior

| Control | Behavior |
| --- | --- |
| `--select` | Request reviewed properties where supported; some routes select locally instead |
| `--fields` | Project locally from properties already fetched |
| `--full` | Remove 500-character text truncation without lifting redaction or row caps |
| `--limit` | Cap collection rows, normally 100 by default |
| `--all` | Follow pages within request, byte and deadline budgets; incompatible with `--limit` |
| `--cursor` | Resume a partial collection under the same profile, scopes and API version |

Availability varies by leaf; consult its `--help`.
List output includes `count`, `total` and `complete` information.
For example, a partial user list without a server total reports `count: 10 users shown, more available` and `total: null`.
A known server total uses a count such as `10 of 143 users`; totals are never invented or fetched with an extra request.

Partial collection results preserve buffered rows and the next page in an opaque cursor; detection correlation uses the separate contract below.
For named reads, repeat the original query flags or omit them; conflicting queries are rejected.
Repeat `--fields` and `--full` to retain the same local view.
Raw reads reuse their original query when `--odata` is omitted on resume.
For large raw cursors, save the token to a file and supply it through stdin:

```sh
mg-axi api get /users --profile soc --cursor - --all --scopes https://graph.microsoft.com/User.Read.All < cursor.txt
```

Both raw cursor input forms have a 16 MB ceiling.
Treat cursor tokens and output as tenant data.
Denied reads return structured errors, except sign-in-stage failures in detection correlation, which return explicit incomplete results as described below.

## Risk detection correlation

`mg-axi entra risk-detection correlate --profile soc --id <risk-detection-id> --limit 20` reads a detection, then queries sign-ins for its `userId` from five minutes before through five minutes after `activityDateTime`.
The combined read stays within 20 requests, 5 MB and 30 seconds; `--limit` caps inspected sign-ins and defaults to 100.
It uses the existing [access and licence prerequisites](graph-coverage.md#licensing-and-completeness-findings).
The form accepts neither custom window flags nor `--all`, `--cursor`, `--select` or `--fields`.
For a different window, use the bounded `entra sign-in list` command suggested in output help.
Sign-in list/show expose the reviewed `correlationId` property through `--select` and `--fields`.

Output includes `riskDetection`, `signIns`, `status` and `complete`, plus `window` when a sign-in query can be attempted.
Each sign-in includes `correlationIdMatches`, `requestIdMatches` and `match`.
An `exact` row requires both matching non-empty `correlationId` values and `detection.requestId = signIn.id`; user/time matching alone produces a `candidate`, even when only one row is returned.
All inspected rows remain in output, including candidates beside an exact row.

| Status | Meaning |
| --- | --- |
| `not-associated` | Detection `correlationId` is null, meaning no sign-in is associated; no sign-in query is sent and `complete` is true. |
| `unavailable` | Detection lacks a usable `userId` or `activityDateTime`; no sign-in query is sent and `complete` is false. |
| `incomplete` | A limit, budget or sign-in-stage failure prevents completion; `reason` explains why and `complete` is false. |
| `exact` | The bounded read completed with exactly one row carrying exact identifier evidence. |
| `ambiguous` | The bounded read completed with multiple rows without exactly one exact row. |
| `candidates` | The bounded read completed with one candidate and no exact row. |
| `no-match` | The bounded read completed without rows; this says nothing about sign-ins outside the window or visibility allowed by retention and licensing. |

Incomplete results retain the detection and any sign-ins collected before a later failure, including row-level identifier evidence, but cannot establish a unique match or absence.
Correlation returns these results with exit 0, so inspect `complete` and `reason` rather than relying on the exit code.
Initial detection-read failures remain structured errors.
There is no correlation cursor; rerun with a larger `--limit` or use the suggested bounded sign-in list to continue investigating.
