import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// READ-06: the Entra risky-user and risk-detection read mapping behind
// `mg-axi entra risky-user list/show` and `mg-axi entra risk-detection
// list/show`, extended with the workload-identity surface behind
// `mg-axi entra risky-service-principal list/show`,
// `mg-axi entra risky-service-principal history list/show` and
// `mg-axi entra service-principal-risk-detection list/show`. Operation construction stays beside its command; the shared
// session owns URLs, credentials, paging, retries and error translation, and
// the SDK owns TOON rendering. This module only maps flags to session calls,
// bounds detection queries in time, and projects rows for compact output.
//
// Reviewed against the v1.0 riskyuser-list/get and riskdetection-list/get
// operation documentation on 2026-10-04, and against the v1.0
// riskyserviceprincipal-list/get, history-list and
// serviceprincipalriskdetection-list/get operation documentation on
// 2026-10-06. Risky-user reads take D/A
// IdentityRiskyUser.Read.All and risk-detection reads take D/A
// IdentityRiskEvent.Read.All; risky-service-principal reads (including
// history) take D/A IdentityRiskyServicePrincipal.Read.All while
// service-principal risk detections reuse D/A IdentityRiskEvent.Read.All.
// Delegated callers additionally need Global
// Reader, Security Operator, Security Reader or Security Administrator. The
// riskyUsers API requires a P2 licence, while risk detection permits P1 or
// P2; the workload-identity APIs require a Microsoft Entra Workload
// Identities Premium licence, and service-principal detections report
// riskDetail and riskLevel hidden without it. Limited views stay limited: a premium detection without P2 detail
// reports riskEventType generic, hidden risk levels report the licence
// boundary instead of the level, and a null detection correlationId means no
// sign-in is associated. None of these is ever reinterpreted as empty or
// invented into a fuller value, and risk dismissal stays WRITE-05.

// Every property this slice may request or display. Anything else fails
// before credentials so typos never become misleading server queries.
// tokenIssuerType is deliberately absent: the shared session redacts every
// *token* key, so the enum could only ever render as ***redacted*** and is
// left out rather than displayed as hidden material.
export const KNOWN_RISKY_USER_FIELDS: readonly string[] = [
  "id",
  "isDeleted",
  "isProcessing",
  "riskDetail",
  "riskLastUpdatedDateTime",
  "riskLevel",
  "riskState",
  "userDisplayName",
  "userPrincipalName",
];
export const KNOWN_RISK_DETECTION_FIELDS: readonly string[] = [
  "id",
  "requestId",
  "correlationId",
  "riskEventType",
  "riskState",
  "riskLevel",
  "riskDetail",
  "source",
  "detectionTimingType",
  "activity",
  "ipAddress",
  "location",
  "activityDateTime",
  "detectedDateTime",
  "lastUpdatedDateTime",
  "userId",
  "userDisplayName",
  "userPrincipalName",
  "additionalInfo",
];
// Workload-identity surface. History items inherit every risky-service-
// principal property and add the service-principal identifier plus the
// actor/activity of the risk change. Service-principal detections carry the
// service-principal join keys (servicePrincipalId, servicePrincipalDisplayName,
// appId, keyIds) instead of user keys; keyIds are key-credential identifiers,
// never secret material. tokenIssuerType stays absent for the same redaction
// reason as above.
export const KNOWN_RISKY_SP_FIELDS: readonly string[] = [
  "id",
  "isEnabled",
  "isProcessing",
  "riskDetail",
  "riskLastUpdatedDateTime",
  "riskLevel",
  "riskState",
  "displayName",
  "appId",
  "servicePrincipalType",
];
export const KNOWN_RISKY_SP_HISTORY_FIELDS: readonly string[] = [
  "id",
  "isEnabled",
  "isProcessing",
  "riskDetail",
  "riskLastUpdatedDateTime",
  "riskLevel",
  "riskState",
  "displayName",
  "appId",
  "servicePrincipalType",
  "servicePrincipalId",
  "initiatedBy",
  "activity",
];
export const KNOWN_SP_DETECTION_FIELDS: readonly string[] = [
  "id",
  "requestId",
  "correlationId",
  "riskEventType",
  "riskState",
  "riskLevel",
  "riskDetail",
  "source",
  "detectionTimingType",
  "activity",
  "ipAddress",
  "location",
  "activityDateTime",
  "detectedDateTime",
  "lastUpdatedDateTime",
  "servicePrincipalId",
  "servicePrincipalDisplayName",
  "appId",
  "keyIds",
  "additionalInfo",
];
const RISKY_USER_KNOWN = new Set(KNOWN_RISKY_USER_FIELDS);
const RISK_DETECTION_KNOWN = new Set(KNOWN_RISK_DETECTION_FIELDS);
const RISKY_SP_KNOWN = new Set(KNOWN_RISKY_SP_FIELDS);
const RISKY_SP_HISTORY_KNOWN = new Set(KNOWN_RISKY_SP_HISTORY_FIELDS);
const SP_DETECTION_KNOWN = new Set(KNOWN_SP_DETECTION_FIELDS);

// Compact list rows: identifier, who is at risk, and the triage state.
const DEFAULT_RISKY_USER_LIST_SELECT = ["id", "userPrincipalName", "riskLevel", "riskState"];
const DEFAULT_RISK_DETECTION_LIST_SELECT = ["id", "detectedDateTime", "userPrincipalName", "riskLevel"];
// Show rows: the full reviewed server set for one object.
const DEFAULT_RISKY_USER_SHOW_SELECT = [...KNOWN_RISKY_USER_FIELDS];
const DEFAULT_RISK_DETECTION_SHOW_SELECT = [...KNOWN_RISK_DETECTION_FIELDS];
export const DEFAULT_RISKY_USER_SCOPES = ["https://graph.microsoft.com/IdentityRiskyUser.Read.All"];
export const DEFAULT_RISK_DETECTION_SCOPES = ["https://graph.microsoft.com/IdentityRiskEvent.Read.All"];
export const DEFAULT_RISKY_SP_SCOPES = ["https://graph.microsoft.com/IdentityRiskyServicePrincipal.Read.All"];
const TRUNCATE_AT = 500;

export type RiskFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, flag: string, known: Set<string>, knownList: readonly string[], help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: RiskFlags, profile: AnyProfile, defaults: readonly string[], help: string): string[] | undefined {
  if (profile.mode === "application") {
    if (flags.scopes !== undefined) {
      throw new AxiError(
        "Application profiles use the configured Graph .default audience; delegated scopes are unavailable",
        "VALIDATION_ERROR",
        [help, "mg-axi profile show --profile <name>"],
      );
    }
    return undefined;
  }
  if (flags.scopes === undefined) return [...defaults];
  const scopes = String(flags.scopes)
    .split(",")
    .map(scope => scope.trim())
    .filter(scope => scope.length > 0);
  return scopes;
}

function isoInstant(raw: unknown, flag: string, help: string): string {
  const text = String(raw).trim();
  const match = /^(\d{4}-\d{2}-\d{2}T(?:[01]\d|2[0-3]):[0-5]\d:[0-5]\d)(?:\.(\d+))?(Z|[+-](?:[01]\d|2[0-3]):[0-5]\d)$/.exec(text);
  const local = match === null ? NaN : Date.parse(`${match[1]}Z`);
  const parsed = match === null ? NaN : Date.parse(`${match[1]}${match[3]}`);
  if (Number.isNaN(local) || Number.isNaN(parsed) || new Date(local).toISOString().slice(0, 19) !== match![1]
    || new Date(parsed).toISOString().length !== 24) {
    throw new AxiError(`--${flag} needs an ISO-8601 instant such as 2026-09-01T00:00:00Z`, "VALIDATION_ERROR", [help]);
  }
  const fraction = (match![2] ?? "").replace(/0+$/, "").padEnd(3, "0");
  return `${new Date(parsed).toISOString().slice(0, 19)}.${fraction}Z`;
}

// Detection collection queries carry an explicit time bound; there is no
// unbounded default. Returns undefined when the caller passed no time or
// filter flags so a cursor resume restores the saved query untouched.
function boundedFilter(dateField: string, flags: RiskFlags, help: string): string | undefined {
  const since = flags.since === undefined ? undefined : isoInstant(flags.since, "since", help);
  const until = flags.until === undefined ? undefined : isoInstant(flags.until, "until", help);
  if (since !== undefined && until !== undefined && until.slice(0, -1) <= since.slice(0, -1)) {
    throw new AxiError("--until must be after --since", "VALIDATION_ERROR", [help]);
  }
  const parts: string[] = [];
  if (since !== undefined) parts.push(`${dateField} ge ${since}`);
  if (until !== undefined) parts.push(`${dateField} le ${until}`);
  if (flags.filter !== undefined) {
    const filter = String(flags.filter).trim();
    if (!filter.length) throw new AxiError("--filter needs a non-empty OData expression", "VALIDATION_ERROR", [help]);
    let quoted = false;
    let depth = 0;
    for (const char of filter) {
      if (char === "'") quoted = !quoted;
      else if (!quoted && char === "(") depth++;
      else if (!quoted && char === ")" && --depth < 0) break;
    }
    if (quoted || depth !== 0) throw new AxiError("--filter needs balanced parentheses and quoted strings to preserve the time bounds", "VALIDATION_ERROR", [help]);
    parts.push(`(${filter})`);
  }
  return parts.length ? parts.join(" and ") : undefined;
}

function truncateValue(value: unknown, full: boolean): { value: unknown; truncated: boolean } {
  const truncate = (text: string) => full || text.length <= TRUNCATE_AT
    ? { value: text, truncated: false }
    : { value: `${text.slice(0, TRUNCATE_AT)}... (truncated, ${text.length} chars total)`, truncated: true };
  if (typeof value === "string") return truncate(value);
  if (Array.isArray(value)) {
    const projected = value.map(entry => truncateValue(entry, full));
    return { value: projected.map(entry => entry.value), truncated: projected.some(entry => entry.truncated) };
  }
  if (value !== null && typeof value === "object") {
    const projected = Object.entries(value).map(([key, entry]) => [key, truncateValue(entry, full)] as const);
    return {
      value: Object.fromEntries(projected.map(([key, entry]) => [key, entry.value])),
      truncated: projected.some(([, entry]) => entry.truncated),
    };
  }
  return { value, truncated: false };
}

// Local projection preserves Graph's null/missing distinction: an explicit
// null stays null, an absent property stays absent and is never synthesized.
// A null detection correlationId means no sign-in is associated, not a
// missing join key; generic and hidden risk values are the server's limited
// view and are never rewritten into a fuller claim.
function project(
  row: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  const projected: Record<string, unknown> = {};
  let truncated = false;
  for (const field of fields) {
    if (!Object.hasOwn(source, field)) continue;
    const result = truncateValue(source[field], full);
    projected[field] = result.value;
    truncated = truncated || result.truncated;
  }
  return { row: projected, truncated };
}

function selectedFields(
  flags: RiskFlags,
  defaults: string[],
  known: Set<string>,
  knownList: readonly string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, "select", known, knownList, help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, "fields", known, knownList, help);
  const missing = fields.find(field => !select.includes(field));
  if (missing) {
    throw new AxiError(`--fields ${missing} was not fetched; request it with --select`, "VALIDATION_ERROR", [help]);
  }
  return { select, fields };
}

function profileHint(profileName: string): string {
  return `--profile ${shellValue(profileName)}`;
}

function shellValue(value: string): string {
  return /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
}

function fullHint(noun: string, action: string, flags: RiskFlags, profileName: string): string {
  const args = Object.entries({ ...flags, ...(flags.cursor === undefined ? {} : { cursor: "-" }), profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi entra ${noun} ${action} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the roles and licensing that actually unlock it, because a
// 403 alone never says which prerequisite is missing.
export function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

export const RISKY_USER_DENIAL_HINTS = [
  "Risky-user reads need IdentityRiskyUser.Read.All plus a supported directory role: Global Reader, Security Operator, Security Reader or Security Administrator for delegated access, or admin-consented IdentityRiskyUser.Read.All for application access",
  "The riskyUsers API requires a Microsoft Entra ID P2 licence; data availability follows Entra retention policies",
];

const RISK_DETECTION_DENIAL_HINTS = [
  "Risk-detection reads need IdentityRiskEvent.Read.All plus a supported directory role: Global Reader, Security Operator, Security Reader or Security Administrator for delegated access, or admin-consented IdentityRiskEvent.Read.All for application access",
  "Risk detection needs a P1 or P2 licence; premium detections report riskEventType generic without P2 detail, so generic is the server's limited view rather than the real type",
];

const RISKY_SP_DENIAL_HINTS = [
  "Risky-service-principal reads need IdentityRiskyServicePrincipal.Read.All plus a supported directory role: Global Reader, Security Operator, Security Reader or Security Administrator for delegated access, or admin-consented IdentityRiskyServicePrincipal.Read.All for application access",
  "The riskyServicePrincipals API requires a Microsoft Entra Workload Identities Premium licence; data availability follows Entra retention policies",
];

const SP_DETECTION_DENIAL_HINTS = [
  "Service-principal-risk-detection reads need IdentityRiskEvent.Read.All plus a supported directory role: Global Reader, Security Operator, Security Reader or Security Administrator for delegated access, or admin-consented IdentityRiskEvent.Read.All for application access",
  "The servicePrincipalRiskDetection API requires a Microsoft Entra Workload Identities Premium licence; riskDetail and riskLevel report hidden without it, so hidden is the server's limited view rather than no risk",
];

// Nested history routes bind the parent risky-service-principal id before
// credentials; a missing or empty parent fails as usage, never as Graph 404.
function parentServicePrincipal(flags: RiskFlags, help: string): Record<string, string> {
  const parent = flags["service-principal"];
  if (parent === undefined || !String(parent).trim()) {
    throw new AxiError("--service-principal needs the risky-service-principal object ID owning this history", "VALIDATION_ERROR", [help]);
  }
  return { "riskyServicePrincipal-id": String(parent) };
}

interface CollectionShape {
  noun: string;
  key: string;
  totalsNoun: string;
  /** Time field bounding new queries; absent for state collections with no time bound. */
  dateField?: string;
  /** Parent path bindings for nested routes such as per-principal history. */
  pathParams?: (flags: RiskFlags, help: string) => Record<string, string>;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  defaultScopes: readonly string[];
  denialHints: string[];
  showHint: string;
  emptyHints: (profileName: string) => string[];
  extraHelp: (profileName: string) => string[];
}

const RISKY_USER_LIST: CollectionShape = {
  noun: "risky-user",
  key: "riskyUsers",
  totalsNoun: "risky users",
  known: RISKY_USER_KNOWN,
  knownList: KNOWN_RISKY_USER_FIELDS,
  defaultSelect: DEFAULT_RISKY_USER_LIST_SELECT,
  defaultScopes: DEFAULT_RISKY_USER_SCOPES,
  denialHints: RISKY_USER_DENIAL_HINTS,
  showHint: "mg-axi entra risky-user show --id <risky-user-id>",
  emptyHints: profileName => [
    `mg-axi entra risky-user list --filter <odata-filter> ${profileHint(profileName)}`,
    "0 risky users matched; the absence of results is the answer, not an error",
    "Limited results stay limited: without P2 detail the service reports what the caller may see, so never read a short list as no risk",
  ],
  extraHelp: () => [],
};

const RISK_DETECTION_LIST: CollectionShape = {
  noun: "risk-detection",
  key: "riskDetections",
  totalsNoun: "risk detections",
  dateField: "detectedDateTime",
  known: RISK_DETECTION_KNOWN,
  knownList: KNOWN_RISK_DETECTION_FIELDS,
  defaultSelect: DEFAULT_RISK_DETECTION_LIST_SELECT,
  defaultScopes: DEFAULT_RISK_DETECTION_SCOPES,
  denialHints: RISK_DETECTION_DENIAL_HINTS,
  showHint: "mg-axi entra risk-detection show --id <risk-detection-id>",
  emptyHints: profileName => [
    `mg-axi entra risk-detection list --since <earlier-iso-time> ${profileHint(profileName)}`,
    "0 risk detections matched in this window; widen --since/--until or loosen --filter - the absence of results is the answer, not an error",
  ],
  extraHelp: profileName => [
    `Correlate a detection to its sign-in with mg-axi entra sign-in list --since <activity-window-start> --filter "userPrincipalName eq '<user-principal-name>'" ${profileHint(profileName)}; there is no riskySignIns endpoint`,
    "Limited views stay limited: riskEventType generic can hide a premium detection without P2 detail, and hidden risk levels report the licence boundary instead of the level - never read them as none",
  ],
};

// Compact workload rows: identifier, which workload, and the triage state.
const DEFAULT_RISKY_SP_LIST_SELECT = ["id", "displayName", "riskLevel", "riskState"];
const DEFAULT_SP_DETECTION_LIST_SELECT = ["id", "detectedDateTime", "servicePrincipalDisplayName", "riskLevel"];
const DEFAULT_RISKY_SP_SHOW_SELECT = [...KNOWN_RISKY_SP_FIELDS];
const DEFAULT_RISKY_SP_HISTORY_SHOW_SELECT = [...KNOWN_RISKY_SP_HISTORY_FIELDS];
const DEFAULT_SP_DETECTION_SHOW_SELECT = [...KNOWN_SP_DETECTION_FIELDS];

const RISKY_SP_LIST: CollectionShape = {
  noun: "risky-service-principal",
  key: "riskyServicePrincipals",
  totalsNoun: "risky service principals",
  known: RISKY_SP_KNOWN,
  knownList: KNOWN_RISKY_SP_FIELDS,
  defaultSelect: DEFAULT_RISKY_SP_LIST_SELECT,
  defaultScopes: DEFAULT_RISKY_SP_SCOPES,
  denialHints: RISKY_SP_DENIAL_HINTS,
  showHint: "mg-axi entra risky-service-principal show --id <risky-service-principal-id>",
  emptyHints: profileName => [
    `mg-axi entra risky-service-principal list --filter <odata-filter> ${profileHint(profileName)}`,
    "0 risky service principals matched; the absence of results is the answer, not an error",
    "Limited results stay limited: without Workload Identities Premium detail the service reports what the caller may see, so never read a short list as no risk",
  ],
  extraHelp: () => [],
};

const RISKY_SP_HISTORY_LIST: CollectionShape = {
  noun: "risky-service-principal history",
  key: "riskyServicePrincipalHistory",
  totalsNoun: "history items",
  pathParams: parentServicePrincipal,
  known: RISKY_SP_HISTORY_KNOWN,
  knownList: KNOWN_RISKY_SP_HISTORY_FIELDS,
  defaultSelect: DEFAULT_RISKY_SP_LIST_SELECT,
  defaultScopes: DEFAULT_RISKY_SP_SCOPES,
  denialHints: RISKY_SP_DENIAL_HINTS,
  showHint: "mg-axi entra risky-service-principal history show --service-principal <risky-service-principal-id> --id <history-item-id>",
  emptyHints: profileName => [
    `mg-axi entra risky-service-principal history list --service-principal <risky-service-principal-id> ${profileHint(profileName)}`,
    "0 history items matched; the absence of results is the answer, not an error",
  ],
  extraHelp: () => [],
};

const SP_DETECTION_LIST: CollectionShape = {
  noun: "service-principal-risk-detection",
  key: "servicePrincipalRiskDetections",
  totalsNoun: "service principal risk detections",
  dateField: "detectedDateTime",
  known: SP_DETECTION_KNOWN,
  knownList: KNOWN_SP_DETECTION_FIELDS,
  defaultSelect: DEFAULT_SP_DETECTION_LIST_SELECT,
  defaultScopes: DEFAULT_RISK_DETECTION_SCOPES,
  denialHints: SP_DETECTION_DENIAL_HINTS,
  showHint: "mg-axi entra service-principal-risk-detection show --id <service-principal-risk-detection-id>",
  emptyHints: profileName => [
    `mg-axi entra service-principal-risk-detection list --since <earlier-iso-time> ${profileHint(profileName)}`,
    "0 service-principal risk detections matched in this window; widen --since/--until or loosen --filter - the absence of results is the answer, not an error",
  ],
  extraHelp: () => [
    "Limited views stay limited: riskDetail and riskLevel report hidden without Workload Identities Premium detail - never read hidden as none",
  ],
};

async function listRisk(
  shape: CollectionShape,
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  if (shape.dateField !== undefined) {
    const bounds = saved === undefined ? undefined : new RegExp(`^${shape.dateField} ge (\\S+)(?: and ${shape.dateField} le (\\S+))?(?: and \\(([\\s\\S]*)\\))?$`).exec(saved.$filter ?? "");
    const savedFlags: RiskFlags | undefined = bounds ? {
      since: bounds[1]!,
      ...(bounds[2] === undefined ? {} : { until: bounds[2] }),
      ...(bounds[3] === undefined ? {} : { filter: bounds[3] }),
    } : undefined;
    if ((saved === undefined && flags.since === undefined)
      || (saved !== undefined && (savedFlags === undefined || boundedFilter(shape.dateField, savedFlags, help) !== saved.$filter))) {
      throw new AxiError(`Detection queries are bounded in time: pass --since <ISO-time> to bound ${shape.dateField}`, "VALIDATION_ERROR", [
        help,
        "Start a new query with --since instead of resuming a cursor without a valid saved time bound",
        `Example: mg-axi entra ${shape.noun} list --since 2026-09-01T00:00:00Z ${profileHint(profileName)}`,
      ]);
    }
  }
  const savedSelect = saved?.$select;
  const { select, fields } = selectedFields(flags,
    savedSelect === undefined ? shape.defaultSelect : fieldList(savedSelect, "select", shape.known, shape.knownList, help),
    shape.known, shape.knownList, help);
  const scopes = scopesFor(flags, profile, shape.defaultScopes, help);
  const full = flags.full === true;
  const query: Record<string, string> = { $select: select.join(",") };
  if (shape.dateField !== undefined) {
    const filter = boundedFilter(shape.dateField, flags, help);
    if (filter !== undefined) query.$filter = filter;
  } else if (flags.filter !== undefined) {
    query.$filter = String(flags.filter);
  }
  const args: CollectArgs = { profile, operation, query, scopes };
  if (shape.pathParams !== undefined) args.params = shape.pathParams(flags, help);
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(shape.denialHints, () => session.collect(args));
  // The hint replays the caller's own time flags (or the cursor alone, which
  // restores the saved query); copying back the composed $filter would wrap
  // it a second time and conflict with the cursor context.
  const effectiveFlags: RiskFlags = { ...flags, select: result.query.$select ?? shape.defaultSelect.join(",") };
  if (result.query.$filter !== undefined && shape.dateField === undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `${shape.showHint} ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(shape.noun, "list", effectiveFlags, profileName)] : [];
  if (truncated && cursor !== undefined) truncationHints.push("Supply the original input cursor on stdin to replay this result with --full");
  if (!result.complete) {
    return {
      [shape.key]: rows,
      ...listTotals(rows.length, result.total, shape.totalsNoun, false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`, showHint, ...shape.extraHelp(profileName)],
    };
  }
  if (!rows.length) {
    return { [shape.key]: rows, ...listTotals(rows.length, result.total, shape.totalsNoun, true), complete: true, help: shape.emptyHints(profileName) };
  }
  return { [shape.key]: rows, ...listTotals(rows.length, result.total, shape.totalsNoun, true), complete: true, help: [...truncationHints, showHint, ...shape.extraHelp(profileName)] };
}

interface SingleShape {
  noun: string;
  key: string;
  param: string;
  /** Extra parent path bindings for nested routes such as per-principal history. */
  parentParams?: (flags: RiskFlags, help: string) => Record<string, string>;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  defaultScopes: readonly string[];
  denialHints: string[];
  correlationHelp: (profileName: string) => string[];
  /** Reported when the returned row carries a null correlationId. */
  noSignInNote?: string;
}

const RISKY_USER_SHOW: SingleShape = {
  noun: "risky-user",
  key: "riskyUser",
  param: "riskyUser-id",
  known: RISKY_USER_KNOWN,
  knownList: KNOWN_RISKY_USER_FIELDS,
  defaultSelect: DEFAULT_RISKY_USER_SHOW_SELECT,
  defaultScopes: DEFAULT_RISKY_USER_SCOPES,
  denialHints: RISKY_USER_DENIAL_HINTS,
  correlationHelp: profileName => [
    `List this user's detections with mg-axi entra risk-detection list --since <iso-time> --filter "userId eq '<risky-user-id>'" ${profileHint(profileName)}`,
    `List this user's sign-ins with mg-axi entra sign-in list --since <iso-time> --filter "userPrincipalName eq '<user-principal-name>'" ${profileHint(profileName)}`,
    "Risk dismissal is a separately reviewed write; this read never confirms, dismisses or remediates risk",
  ],
};

const RISK_DETECTION_SHOW: SingleShape = {
  noun: "risk-detection",
  key: "riskDetection",
  param: "riskDetection-id",
  known: RISK_DETECTION_KNOWN,
  knownList: KNOWN_RISK_DETECTION_FIELDS,
  defaultSelect: DEFAULT_RISK_DETECTION_SHOW_SELECT,
  defaultScopes: DEFAULT_RISK_DETECTION_SCOPES,
  denialHints: RISK_DETECTION_DENIAL_HINTS,
  correlationHelp: profileName => [
    `Correlate this detection to its sign-in with mg-axi entra sign-in list --since <activity-window-start> --filter "userPrincipalName eq '<user-principal-name>'" ${profileHint(profileName)}; there is no riskySignIns endpoint`,
    "Limited views stay limited: riskEventType generic can hide a premium detection without P2 detail, and hidden risk levels report the licence boundary instead of the level - never read them as none",
  ],
  noSignInNote: "This detection carries no associated sign-in (correlationId is null); sign-in correlation does not apply",
};

const RISKY_SP_SHOW: SingleShape = {
  noun: "risky-service-principal",
  key: "riskyServicePrincipal",
  param: "riskyServicePrincipal-id",
  known: RISKY_SP_KNOWN,
  knownList: KNOWN_RISKY_SP_FIELDS,
  defaultSelect: DEFAULT_RISKY_SP_SHOW_SELECT,
  defaultScopes: DEFAULT_RISKY_SP_SCOPES,
  denialHints: RISKY_SP_DENIAL_HINTS,
  correlationHelp: profileName => [
    `List this workload's detections with mg-axi entra service-principal-risk-detection list --since <iso-time> --filter "servicePrincipalId eq '<risky-service-principal-id>'" ${profileHint(profileName)}`,
    `List this workload's risk history with mg-axi entra risky-service-principal history list --service-principal <risky-service-principal-id> ${profileHint(profileName)}`,
    "Confirming compromise and dismissing risk are separately reviewed writes; this read never confirms, dismisses or remediates risk",
  ],
};

const RISKY_SP_HISTORY_SHOW: SingleShape = {
  noun: "risky-service-principal history",
  key: "riskyServicePrincipalHistoryItem",
  param: "riskyServicePrincipalHistoryItem-id",
  parentParams: parentServicePrincipal,
  known: RISKY_SP_HISTORY_KNOWN,
  knownList: KNOWN_RISKY_SP_HISTORY_FIELDS,
  defaultSelect: DEFAULT_RISKY_SP_HISTORY_SHOW_SELECT,
  defaultScopes: DEFAULT_RISKY_SP_SCOPES,
  denialHints: RISKY_SP_DENIAL_HINTS,
  correlationHelp: profileName => [
    `List this workload's detections with mg-axi entra service-principal-risk-detection list --since <iso-time> --filter "servicePrincipalId eq '<risky-service-principal-id>'" ${profileHint(profileName)}`,
  ],
};

const SP_DETECTION_SHOW: SingleShape = {
  noun: "service-principal-risk-detection",
  key: "servicePrincipalRiskDetection",
  param: "servicePrincipalRiskDetection-id",
  known: SP_DETECTION_KNOWN,
  knownList: KNOWN_SP_DETECTION_FIELDS,
  defaultSelect: DEFAULT_SP_DETECTION_SHOW_SELECT,
  defaultScopes: DEFAULT_RISK_DETECTION_SCOPES,
  denialHints: SP_DETECTION_DENIAL_HINTS,
  correlationHelp: profileName => [
    `List this workload's detections with mg-axi entra service-principal-risk-detection list --since <iso-time> --filter "servicePrincipalId eq '<service-principal-id>'" ${profileHint(profileName)}`,
    "Limited views stay limited: riskDetail and riskLevel report hidden without Workload Identities Premium detail - never read hidden as none",
  ],
  noSignInNote: "This detection carries no associated sign-in (correlationId is null); sign-in correlation does not apply",
};

async function showRisk(
  shape: SingleShape,
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, shape.defaultSelect, shape.known, shape.knownList, help);
  const scopes = scopesFor(flags, profile, shape.defaultScopes, help);
  const full = flags.full === true;
  const raw = await withGuidance(shape.denialHints, () => session.execute({
    profile,
    operation,
    params: { ...(shape.parentParams === undefined ? {} : shape.parentParams(flags, help)), [shape.param]: String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed risk body", "GRAPH_ERROR", [
      "Single-object reads carry one object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  const hints = [...shape.correlationHelp(profileName)];
  // A null correlationId is the server's statement that no sign-in is
  // associated, so sign-in correlation explicitly does not apply.
  if (shape.noSignInNote !== undefined && Object.hasOwn(row, "correlationId") && row.correlationId === null) {
    hints.push(shape.noSignInNote);
  }
  if (truncated) return { [shape.key]: row, help: [fullHint(shape.noun, "show", flags, profileName), ...hints] };
  return { [shape.key]: row, help: hints };
}

export async function listRiskyUsers(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listRisk(RISKY_USER_LIST, session, flags, profile, operation, help, profileName);
}

export async function showRiskyUser(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showRisk(RISKY_USER_SHOW, session, flags, profile, operation, help, profileName);
}

export async function listRiskDetections(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listRisk(RISK_DETECTION_LIST, session, flags, profile, operation, help, profileName);
}

export async function showRiskDetection(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showRisk(RISK_DETECTION_SHOW, session, flags, profile, operation, help, profileName);
}

export async function listRiskyServicePrincipals(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listRisk(RISKY_SP_LIST, session, flags, profile, operation, help, profileName);
}

export async function showRiskyServicePrincipal(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showRisk(RISKY_SP_SHOW, session, flags, profile, operation, help, profileName);
}

export async function listRiskyServicePrincipalHistory(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listRisk(RISKY_SP_HISTORY_LIST, session, flags, profile, operation, help, profileName);
}

export async function showRiskyServicePrincipalHistory(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showRisk(RISKY_SP_HISTORY_SHOW, session, flags, profile, operation, help, profileName);
}

export async function listServicePrincipalRiskDetections(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listRisk(SP_DETECTION_LIST, session, flags, profile, operation, help, profileName);
}

export async function showServicePrincipalRiskDetection(
  session: GraphSession,
  flags: RiskFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showRisk(SP_DETECTION_SHOW, session, flags, profile, operation, help, profileName);
}
