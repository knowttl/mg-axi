import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// READ-05: the Entra sign-in and directory-audit read mapping behind
// `mg-axi entra sign-in list/show` and `mg-axi entra directory-audit
// list/show`. Operation construction stays beside its command; the shared
// session owns URLs, credentials, paging, retries and error translation, and
// the SDK owns TOON rendering. This module only maps flags to session calls,
// bounds every collection query in time, and projects rows for compact output.
//
// Reviewed against the v1.0 signin-list and directoryaudit-list operation
// documentation on 2026-10-04. Both collections read D/A AuditLog.Read.All;
// delegated sign-in reads additionally accept Global Reader, Reports Reader,
// Security Administrator, Security Operator or Security Reader, while
// delegated directory-audit reads accept Reports Reader, Security
// Administrator or Security Reader. Graph omits appliedConditionalAccessPolicies
// without a CA-data permission/role, so an absent property is reported as
// unavailable rather than invented. Sign-in/audit Graph reporting carries a
// conservative P1/P2 deployment prerequisite; retention and premium fields
// stay separate constraints.

// Every property this slice may request or display. Anything else fails
// before credentials so typos never become misleading server queries.
export const KNOWN_SIGNIN_FIELDS: readonly string[] = [
  "id",
  "createdDateTime",
  "userId",
  "userPrincipalName",
  "userDisplayName",
  "appDisplayName",
  "appId",
  "ipAddress",
  "location",
  "status",
  "conditionalAccessStatus",
  "appliedConditionalAccessPolicies",
  "riskDetail",
  "riskLevelAggregated",
  "riskLevelDuringSignIn",
  "riskState",
  "resourceDisplayName",
  "resourceId",
  "clientAppUsed",
];
export const KNOWN_AUDIT_FIELDS: readonly string[] = [
  "id",
  "activityDateTime",
  "activityDisplayName",
  "category",
  "loggedByService",
  "operationType",
  "result",
  "resultReason",
  "correlationId",
  "initiatedBy",
  "targetResources",
];
const SIGNIN_KNOWN = new Set(KNOWN_SIGNIN_FIELDS);
const AUDIT_KNOWN = new Set(KNOWN_AUDIT_FIELDS);

// Compact list rows: identifier, time, who/what happened, and outcome.
const DEFAULT_SIGNIN_LIST_SELECT = ["id", "createdDateTime", "userPrincipalName", "appDisplayName"];
const DEFAULT_AUDIT_LIST_SELECT = ["id", "activityDateTime", "activityDisplayName", "result"];
// Show rows: the full reviewed server set for one object.
const DEFAULT_SIGNIN_SHOW_SELECT = [...KNOWN_SIGNIN_FIELDS];
const DEFAULT_AUDIT_SHOW_SELECT = [...KNOWN_AUDIT_FIELDS];
export const DEFAULT_DELEGATED_SCOPES = ["https://graph.microsoft.com/AuditLog.Read.All"];
const TRUNCATE_AT = 500;

// Graph omits this property without CA-data access instead of failing, so an
// absent value reports the missing prerequisite rather than an empty result.
const SIGNIN_UNAVAILABLE: Readonly<Record<string, string>> = {
  appliedConditionalAccessPolicies:
    "unavailable: Graph omits CA policy detail without CA-data access - both modes need Policy.Read.All, Policy.Read.ConditionalAccess or Policy.ReadWrite.ConditionalAccess in addition to AuditLog.Read.All; delegated also needs Conditional Access Administrator, Global Reader, Security Administrator or Security Reader. For delegated access, log in and repeat this read with --scopes https://graph.microsoft.com/AuditLog.Read.All,https://graph.microsoft.com/Policy.Read.All",
};

export type LogFlags = Record<string, string | boolean>;

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

function scopesFor(flags: LogFlags, profile: AnyProfile, help: string): string[] | undefined {
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
  if (flags.scopes === undefined) return [...DEFAULT_DELEGATED_SCOPES];
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

// Every collection query carries an explicit time bound; there is no
// unbounded default. Returns undefined when the caller passed no time or
// filter flags so a cursor resume restores the saved query untouched.
function boundedFilter(dateField: string, flags: LogFlags, help: string): string | undefined {
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
// null stays null, an absent base property stays absent and is never
// synthesized. Properties Graph omits without an entitlement report the
// missing prerequisite instead.
function project(
  row: unknown,
  fields: string[],
  unavailable: Readonly<Record<string, string>>,
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  const projected: Record<string, unknown> = {};
  let truncated = false;
  for (const field of fields) {
    if (!Object.hasOwn(source, field)) {
      if (unavailable[field] !== undefined) projected[field] = unavailable[field];
      continue;
    }
    const result = truncateValue(source[field], full);
    projected[field] = result.value;
    truncated = truncated || result.truncated;
  }
  return { row: projected, truncated };
}

function selectedFields(
  flags: LogFlags,
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

function fullHint(noun: string, action: string, flags: LogFlags, profileName: string): string {
  const args = Object.entries({ ...flags, ...(flags.cursor === undefined ? {} : { cursor: "-" }), profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi entra ${noun} ${action} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the roles and licensing that actually unlock it, because a
// 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const SIGNIN_DENIAL_HINTS = [
  "Sign-in reads need AuditLog.Read.All plus a supported directory role: Global Reader, Reports Reader, Security Administrator, Security Operator or Security Reader for delegated access, or admin-consented AuditLog.Read.All for application access",
  "Sign-in Graph reporting carries a conservative P1/P2 deployment prerequisite; retention and premium fields stay separate constraints",
];

const AUDIT_DENIAL_HINTS = [
  "Directory-audit reads need AuditLog.Read.All plus a supported directory role: Reports Reader, Security Administrator or Security Reader for delegated access, or admin-consented AuditLog.Read.All for application access",
  "Audit Graph reporting carries a conservative P1/P2 deployment prerequisite; retention stays a separate constraint",
];

interface CollectionShape {
  noun: string;
  key: string;
  dateField: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  unavailable: Readonly<Record<string, string>>;
  denialHints: string[];
  showHint: string;
  emptyHints: (profileName: string) => string[];
}

const SIGNIN_LIST: CollectionShape = {
  noun: "sign-in",
  key: "signIns",
  dateField: "createdDateTime",
  known: SIGNIN_KNOWN,
  knownList: KNOWN_SIGNIN_FIELDS,
  defaultSelect: DEFAULT_SIGNIN_LIST_SELECT,
  unavailable: SIGNIN_UNAVAILABLE,
  denialHints: SIGNIN_DENIAL_HINTS,
  showHint: "mg-axi entra sign-in show --id <sign-in-id>",
  emptyHints: profileName => [
    `mg-axi entra sign-in list --since <earlier-iso-time> ${profileHint(profileName)}`,
    "0 sign-ins matched in this window; widen --since/--until or loosen --filter - the absence of results is the answer, not an error",
  ],
};

const AUDIT_LIST: CollectionShape = {
  noun: "directory-audit",
  key: "directoryAudits",
  dateField: "activityDateTime",
  known: AUDIT_KNOWN,
  knownList: KNOWN_AUDIT_FIELDS,
  defaultSelect: DEFAULT_AUDIT_LIST_SELECT,
  unavailable: {},
  denialHints: AUDIT_DENIAL_HINTS,
  showHint: "mg-axi entra directory-audit show --id <directory-audit-id>",
  emptyHints: profileName => [
    `mg-axi entra directory-audit list --since <earlier-iso-time> ${profileHint(profileName)}`,
    "0 directory audits matched in this window; widen --since/--until or loosen --filter - the absence of results is the answer, not an error",
  ],
};

async function listLogs(
  shape: CollectionShape,
  session: GraphSession,
  flags: LogFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  if (cursor === undefined && flags.since === undefined) {
    throw new AxiError(`Log queries are bounded in time: pass --since <ISO-time> to bound ${shape.dateField}`, "VALIDATION_ERROR", [
      help,
      `Example: mg-axi entra ${shape.noun} list --since 2026-09-01T00:00:00Z ${profileHint(profileName)}`,
    ]);
  }
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const savedSelect = saved?.$select;
  const { select, fields } = selectedFields(flags,
    savedSelect === undefined ? shape.defaultSelect : fieldList(savedSelect, "select", shape.known, shape.knownList, help),
    shape.known, shape.knownList, help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const filter = boundedFilter(shape.dateField, flags, help);
  const query: Record<string, string> = { $select: select.join(",") };
  const explicitFilter = filter !== undefined;
  if (explicitFilter) query.$filter = filter;
  const args: CollectArgs = { profile, operation, query, scopes };
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
  const effectiveFlags: LogFlags = { ...flags, select: result.query.$select ?? shape.defaultSelect.join(",") };
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, shape.unavailable, full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `${shape.showHint} ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(shape.noun, "list", effectiveFlags, profileName)] : [];
  if (truncated && cursor !== undefined) truncationHints.push("Supply the original input cursor on stdin to replay this result with --full");
  if (!result.complete) {
    return {
      [shape.key]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`, showHint],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return { [shape.key]: rows, count, help: shape.emptyHints(profileName) };
  }
  return { [shape.key]: rows, count, help: [...truncationHints, showHint] };
}

interface SingleShape {
  noun: string;
  key: string;
  param: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  unavailable: Readonly<Record<string, string>>;
  denialHints: string[];
  idFlag: string;
}

const SIGNIN_SHOW: SingleShape = {
  noun: "sign-in",
  key: "signIn",
  param: "signIn-id",
  known: SIGNIN_KNOWN,
  knownList: KNOWN_SIGNIN_FIELDS,
  defaultSelect: DEFAULT_SIGNIN_SHOW_SELECT,
  unavailable: SIGNIN_UNAVAILABLE,
  denialHints: SIGNIN_DENIAL_HINTS,
  idFlag: "<sign-in-id>",
};

const AUDIT_SHOW: SingleShape = {
  noun: "directory-audit",
  key: "directoryAudit",
  param: "directoryAudit-id",
  known: AUDIT_KNOWN,
  knownList: KNOWN_AUDIT_FIELDS,
  defaultSelect: DEFAULT_AUDIT_SHOW_SELECT,
  unavailable: {},
  denialHints: AUDIT_DENIAL_HINTS,
  idFlag: "<directory-audit-id>",
};

async function showLog(
  shape: SingleShape,
  session: GraphSession,
  flags: LogFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, shape.defaultSelect, shape.known, shape.knownList, help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(shape.denialHints, () => session.execute({
    profile,
    operation,
    params: { [shape.param]: String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed log body", "GRAPH_ERROR", [
      "Single-object reads carry one object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, shape.unavailable, full);
  if (truncated) return { [shape.key]: row, help: [fullHint(shape.noun, "show", flags, profileName)] };
  return { [shape.key]: row };
}

export async function listSignIns(
  session: GraphSession,
  flags: LogFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listLogs(SIGNIN_LIST, session, flags, profile, operation, help, profileName);
}

export async function showSignIn(
  session: GraphSession,
  flags: LogFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showLog(SIGNIN_SHOW, session, flags, profile, operation, help, profileName);
}

export async function listDirectoryAudits(
  session: GraphSession,
  flags: LogFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listLogs(AUDIT_LIST, session, flags, profile, operation, help, profileName);
}

export async function showDirectoryAudit(
  session: GraphSession,
  flags: LogFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showLog(AUDIT_SHOW, session, flags, profile, operation, help, profileName);
}
