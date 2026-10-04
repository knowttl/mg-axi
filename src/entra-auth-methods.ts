import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// READ-04: the Entra authentication-method and registration-report read
// mapping behind `mg-axi entra user authentication-method list` and
// `mg-axi entra registration list`. Operation construction stays beside its
// command; the shared session owns URLs, credentials, paging, retries and
// error translation, and the SDK owns TOON rendering. This module only maps
// flags to session calls and projects rows for compact output.
//
// Reviewed against the v1.0 authentication-list-methods and
// authenticationmethodsroot-list-userregistrationdetails operation
// documentation on 2026-10-04. Per-user method reads take D/A
// UserAuthenticationMethod.Read.All for other users (delegated self-reads use
// UserAuthenticationMethod.Read); delegated callers acting on another user
// additionally need Global Reader, Authentication Administrator or Privileged
// Authentication Administrator. The registration report reads D/A
// AuditLog.Read.All; delegated callers additionally need Reports Reader,
// Security Reader, Security Administrator or Global Reader.
//
// Upstream discourages iterating per-user methods across the tenant for
// auditing; aggregate coverage belongs to the registration report. This slice
// therefore serves one named user per method-list call and offers no tenant
// scan: there is no command that loops users, and the standing help points
// aggregate questions at the report. The report itself does not work for
// disabled users, so absence from it is never proof of no MFA; that gap rides
// in the leaf help and every report output.
//
// Phone numbers are protected values: the shared session replaces them with
// the redaction marker before collection buffering and on cursor decode, so
// neither output nor resume cursors ever carry one. Method rows additionally
// preserve @odata.type (naming the method kind) without making it selectable.
// No method registration or deletion is constructed here.

// Every authentication-method property this slice may request or display,
// matching the reviewed API-01 field set. Anything else fails before
// credentials so typos never become misleading server queries.
export const KNOWN_METHOD_FIELDS: readonly string[] = [
  "id",
  "createdDateTime",
  "displayName",
  "phoneNumber",
  "phoneType",
  "smsSignInState",
  "emailAddress",
];
const KNOWN_METHODS = new Set(KNOWN_METHOD_FIELDS);

// Every registration-report property this slice may request or display,
// matching the reviewed API-01 field set.
export const KNOWN_REPORT_FIELDS: readonly string[] = [
  "id",
  "userPrincipalName",
  "userDisplayName",
  "userType",
  "isAdmin",
  "isMfaRegistered",
  "isMfaCapable",
  "isPasswordlessCapable",
  "isSsprRegistered",
  "isSsprEnabled",
  "isSsprCapable",
  "userPreferredMethodForSecondaryAuthentication",
  "lastUpdatedDateTime",
];
const KNOWN_REPORT = new Set(KNOWN_REPORT_FIELDS);
// @odata.type is preserved on method rows without being selectable: it names
// the method kind (phone, email, fido2 and the other derived types).
const METHOD_TYPE_PROPERTY = "@odata.type";

// Compact method rows: identifier, display name and registration time. The
// method kind rides along via @odata.type; phoneNumber stays out of the
// default selection and is redacted whenever it is fetched.
const DEFAULT_METHOD_SELECT = ["id", "displayName", "createdDateTime"];
// Compact report rows: identifier, user keys and the headline MFA posture.
const DEFAULT_REPORT_SELECT = ["id", "userPrincipalName", "userDisplayName", "isMfaRegistered"];
// UserAuthenticationMethod.Read.All covers other-user reads in both modes;
// delegated self-reads may use UserAuthenticationMethod.Read instead.
export const DEFAULT_METHOD_SCOPES = ["https://graph.microsoft.com/UserAuthenticationMethod.Read.All"];
// AuditLog.Read.All covers the registration report in both modes.
export const DEFAULT_REPORT_SCOPES = ["https://graph.microsoft.com/AuditLog.Read.All"];
const TRUNCATE_AT = 500;

// The only operations this slice ever binds. Method registration, deletion
// and every other route are refused before credentials.
const READ_OPERATIONS: Readonly<Record<string, string>> = {
  "GET:/users/{user-id}/authentication/methods": "method",
  "GET:/reports/authenticationMethods/userRegistrationDetails": "report",
};

function checkReadOperation(operation: SessionOperation, help: string): void {
  // Version stays the session's decision: beta reads remain preview-gated
  // there, exactly like every other named slice. This gate only ensures a
  // registration/deletion route or any other non-read can never reach
  // credentials from this module.
  const route = `${operation.method}:${operation.path}`;
  if (operation.method !== "GET" || !Object.hasOwn(READ_OPERATIONS, route)) {
    throw new AxiError(`Refused non-read route ${operation.id}: READ-04 serves only the catalogued per-user method list and registration report`, "VALIDATION_ERROR", [
      help,
      "Method registration and deletion belong to no read slice; aggregate coverage belongs to the registration report",
    ]);
  }
}

export type AuthMethodFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
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

function scopesFor(flags: AuthMethodFlags, profile: AnyProfile, defaults: readonly string[], help: string): string[] | undefined {
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

function truncateValue(value: unknown, full: boolean): { value: unknown; truncated: boolean } {
  if (typeof value === "string" && !full && value.length > TRUNCATE_AT) {
    return { value: `${value.slice(0, TRUNCATE_AT)}... (truncated, ${value.length} chars total)`, truncated: true };
  }
  return { value, truncated: false };
}

// Local projection preserves Graph's null/missing distinction: an explicit
// null stays null, an absent property stays absent and is never synthesized.
// Redacted phone numbers arrive from the session as the redaction marker and
// pass through unchanged; truncation never hides the marker.
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

// Method projection additionally preserves @odata.type so the method kind
// survives local projection.
function projectMethod(
  row: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  const { row: projected, truncated } = project(row, fields, full);
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  if (typeof source[METHOD_TYPE_PROPERTY] === "string") projected[METHOD_TYPE_PROPERTY] = source[METHOD_TYPE_PROPERTY];
  return { row: projected, truncated };
}

function selectedFields(
  flags: AuthMethodFlags,
  defaults: string[],
  known: Set<string>,
  knownList: readonly string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, known, knownList, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, known, knownList, "fields", help);
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

function fullHint(command: string, flags: AuthMethodFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the roles that actually unlock it, because a 403 alone
// never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const METHOD_DENIAL_HINTS = [
  "Authentication-method reads need UserAuthenticationMethod.Read.All plus Global Reader, Authentication Administrator or Privileged Authentication Administrator for delegated access to another user (delegated self-reads use UserAuthenticationMethod.Read), or admin-consented UserAuthenticationMethod.Read.All for application access",
  "Phone numbers stay redacted for every role; an Authentication Administrator additionally sees masked phone numbers server-side",
];

const REPORT_DENIAL_HINTS = [
  "Registration-report reads need AuditLog.Read.All plus Reports Reader, Security Reader, Security Administrator or Global Reader for delegated access, or admin-consented AuditLog.Read.All for application access",
  "The report does not cover disabled users; a 403 is a grant or role denial, never evidence about a disabled account",
];

interface CollectionShape {
  command: string;
  key: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  defaultScopes: readonly string[];
  denialHints: string[];
  standing: (profileName: string) => string[];
  emptyNote: string;
  preserveType: boolean;
}

const METHOD_LIST: CollectionShape = {
  command: "entra user authentication-method list",
  key: "authenticationMethods",
  known: KNOWN_METHODS,
  knownList: KNOWN_METHOD_FIELDS,
  defaultSelect: DEFAULT_METHOD_SELECT,
  defaultScopes: DEFAULT_METHOD_SCOPES,
  denialHints: METHOD_DENIAL_HINTS,
  standing: profileName => [
    "Targeted inspection of one named user only; for tenant-wide MFA coverage use "
      + `mg-axi entra registration list ${profileHint(profileName)}, never a scan across users`,
  ],
  emptyNote: "0 authentication methods matched; the absence of results is the answer, not an error",
  preserveType: true,
};

const REPORT_LIST: CollectionShape = {
  command: "entra registration list",
  key: "registrationDetails",
  known: KNOWN_REPORT,
  knownList: KNOWN_REPORT_FIELDS,
  defaultSelect: DEFAULT_REPORT_SELECT,
  defaultScopes: DEFAULT_REPORT_SCOPES,
  denialHints: REPORT_DENIAL_HINTS,
  standing: () => [
    "The registration report does not cover disabled users; absence from this report is not proof of no MFA",
  ],
  emptyNote: "0 registration rows matched; the absence of results is the answer, not an error",
  preserveType: false,
};

async function listCollection(
  shape: CollectionShape,
  session: GraphSession,
  flags: AuthMethodFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  params?: Record<string, string>,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const savedSelect = saved?.$select;
  const { select, fields } = selectedFields(flags,
    savedSelect === undefined ? shape.defaultSelect : fieldList(savedSelect, shape.known, shape.knownList, "select", help),
    shape.known, shape.knownList, help);
  const savedFilter = saved?.$filter;
  const filter = flags.filter === undefined ? savedFilter : String(flags.filter);
  const scopes = scopesFor(flags, profile, shape.defaultScopes, help);
  const full = flags.full === true;
  // Neither collection documents a method-specific query contract, so
  // --filter passes through as plain $filter with no $count or
  // ConsistencyLevel attached, exactly like the role reads.
  const query: Record<string, string> = { $select: select.join(",") };
  if (filter !== undefined) query.$filter = filter;
  const args: CollectArgs = { profile, operation, query, scopes };
  if (params) args.params = params;
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(shape.denialHints, () => session.collect(args));
  const effectiveFlags: AuthMethodFlags = { ...flags, select: result.query.$select ?? shape.defaultSelect.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = shape.preserveType ? projectMethod(row, fields, full) : project(row, fields, full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const truncationHints = truncated ? [fullHint(shape.command, effectiveFlags, profileName)] : [];
  const standing = shape.standing(profileName);
  if (!result.complete) {
    return {
      [shape.key]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor <cursor-from-output> ${profileHint(profileName)}`, ...standing],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return { [shape.key]: rows, count, help: [shape.emptyNote, ...standing] };
  }
  return { [shape.key]: rows, count, help: [...truncationHints, ...standing] };
}

export async function listAuthenticationMethods(
  session: GraphSession,
  flags: AuthMethodFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const user = String(flags.user);
  if (!user.trim()) throw new AxiError("--user needs the target user ID or UPN", "VALIDATION_ERROR", [help]);
  return listCollection(METHOD_LIST, session, flags, profile, operation, help, profileName, { "user-id": user });
}

export async function listRegistrationDetails(
  session: GraphSession,
  flags: AuthMethodFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(REPORT_LIST, session, flags, profile, operation, help, profileName);
}
