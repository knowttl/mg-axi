# Defender XDR advanced hunting

`mg-axi security hunting query --file hunt.kql --profile soc --timespan P7D --limit 1000` runs a reviewed read-only Graph POST to `/security/runHuntingQuery`.
Save this example as `hunt.kql`:

```kql
UrlClickEvents
| where ActionType == "ClickAllowed"
| project Timestamp, AccountUpn, Url, ActionType
| order by Timestamp desc
```

Before use, an administrator must grant Microsoft Graph **ThreatHunting.Read.All**, with admin consent, for the profile's delegated or application authentication mode.
Existing Entra permissions do not grant hunting access.
Add `"security"` to that profile's `enabledPacks` array in the configuration file (for example `["entra", "security"]`); existing profiles default to Entra only.
For delegated authentication, sign in explicitly with `mg-axi auth login --profile soc --scopes https://graph.microsoft.com/ThreatHunting.Read.All` after consent.
Application profiles use their configured Graph `.default` audience.
Defender data availability, licensing and the signed-in user's data access still apply; HTTP 403 alone does not identify the missing prerequisite.

The default `--timespan P1D` looks back from now; use whole-day durations such as `P7D`, or time durations such as `PT6H` or `PT30M`, from one second to 30 days.
Graph also applies any narrower time filter in KQL.
The default `--limit 100` accepts 1-10000 rows and appends a server-side `take` for one extra row to detect overflow.
Overflow returns `complete: false` and `reason: row-limit`; it never invents a total or cursor.
Microsoft documents no paging or cursor for this action; narrow the time window or query filters to retrieve more results.
Results use TOON and existing secret redaction, with recursive 500-character string previews; `--full` removes text truncation while retaining redaction and query bounds.
The client uses a 30-second deadline and a 5 MB response budget, refuses redirects and does not automatically replay hunting requests.
Microsoft's service limits are 100000 rows, 50 MB and three minutes per request, with tenant CPU and call quotas.
See [runHuntingQuery](https://learn.microsoft.com/en-us/graph/api/security-security-runhuntingquery?view=graph-rest-1.0), [security quotas](https://learn.microsoft.com/en-us/graph/api/resources/security-api-overview?view=graph-rest-1.0#quotas-and-resource-allocation), and [permission consent](https://learn.microsoft.com/en-us/graph/permissions-reference#threathuntingreadall).

KQL input is file-only and is never echoed in results or server errors.
The local validator accepts one table pipeline, ordinary quoted strings, `//` comments, balanced parentheses/brackets, and the following operators: `where`, `project`, `project-away`, `project-rename`, `project-reorder`, `extend`, `summarize`, `distinct`, `count`, `take`, `limit`, `top`, `sort`, `order`.
Only reviewed scalar and aggregate function calls are accepted, including `ago`, `datetime`, `bin`, `count`, `dcount`, `sum`, `arg_max`, `tostring`, `parse_json` and `parse_url`; the exact allowlist lives in [security-hunting.ts](../src/security-hunting.ts).
The table allowlist covers `DeviceEvents`, `DeviceProcessEvents`, `DeviceNetworkEvents`, `DeviceFileEvents`, `DeviceFileCertificateInfo`, `DeviceRegistryEvents`, `DeviceLogonEvents`, `DeviceImageLoadEvents`, `DeviceInfo`, `DeviceNetworkInfo`, `DeviceTvmSoftwareInventory`, `DeviceTvmSoftwareVulnerabilities`, `DeviceTvmSoftwareVulnerabilitiesKB`, `EmailEvents`, `EmailPostDeliveryEvents`, `EmailUrlInfo`, `EmailAttachmentInfo`, `UrlClickEvents`, `IdentityLogonEvents`, `IdentityQueryEvents`, `IdentityDirectoryEvents`, `CloudAppEvents`, `AlertInfo`, `AlertEvidence`, `BehaviorInfo`, `BehaviorEntities`, `ExposureGraphNodes` and `ExposureGraphEdges`.
Everything outside this restricted grammar is refused before credentials: semicolon statements (including trailing semicolons), `let`/`set`, dot-prefixed control commands, stored/source functions, qualified functions, external data, plugins (`evaluate`/`invoke`), joins/unions, nested query pipelines and block comments.
Microsoft's reference Kusto parser checks syntax locally before the allowlist and before credentials or network access.
Syntax failures keep `VALIDATION_ERROR` (exit 2) and include up to five diagnostics with line, column, code, token and message; allowlist failures also identify the offending token and position.
Queries and quoted literals are never printed as error excerpts.
Successful parsing prints no validation claim: table/column availability, types and permissions are still checked by Graph.
Recognized hunting 400 errors expose only the semantic/syntax error class and line/column, with an opaque fallback for unknown shapes.
The undocumented Graph error shape is covered by synthetic offline fixtures and has not been verified against a live tenant.
Raw `api get` keeps its reviewed GET-only route list and does not expose `/security`.
