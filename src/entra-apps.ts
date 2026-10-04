import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { SAFE_CREDENTIAL_FIELDS } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// READ-07: the Entra application and service-principal read mapping behind
// `mg-axi entra application list/show`, `mg-axi entra service-principal
// list/show` and the two `owner list` relationship reads. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 application-list, application-get,
// servicePrincipal-list, servicePrincipal-get, application-list-owners and
// servicePrincipal-list-owners operation documentation on 2026-10-04.
// Delegated and application reads take Application.Read.All; the object id
// (`id`) is never conflated with the client ID (`appId`) - both ride in the
// compact list defaults so callers can see the distinction. Consent grants
// (appRoleAssignments, oauth2PermissionGrants) are READ-08 and stay out of
// this slice; only owners are projected as relationships.
//
// Credential safety belongs to the shared session's SAFE_CREDENTIAL_FIELDS
// projection, applied before collection buffering and again on cursor decode.
// Local projection uses that same allowlist for the displayed view.
// The six GET operations below are the only ones this module ever binds;
// secret-minting routes (addPassword, addKey) and every write stay refused before
// credentials via checkReadOperation.

// Every application property this slice may request or display. Anything
// else fails before credentials so typos never become misleading queries.
export const KNOWN_APP_FIELDS: readonly string[] = [
  "id",
  "appId",
  "displayName",
  "signInAudience",
  "publisherDomain",
  "createdDateTime",
  "keyCredentials",
  "passwordCredentials",
];
const KNOWN_APPS = new Set(KNOWN_APP_FIELDS);

// Every service-principal property this slice may request or display.
export const KNOWN_SP_FIELDS: readonly string[] = [
  "id",
  "appId",
  "displayName",
  "servicePrincipalType",
  "accountEnabled",
  "appOwnerOrganizationId",
  "appRoleAssignmentRequired",
  "preferredSingleSignOnMode",
  "loginUrl",
  "keyCredentials",
  "passwordCredentials",
];
const KNOWN_SPS = new Set(KNOWN_SP_FIELDS);

// Owner rows are directoryObjects of mixed types; richer per-type fields
// need single-object reads, so only these stay selectable.
export const KNOWN_OWNER_FIELDS: readonly string[] = ["id", "displayName", "mail"];
const KNOWN_OWNERS = new Set(KNOWN_OWNER_FIELDS);
// @odata.type is preserved on owner rows without being selectable: it names
// the owner kind (user, servicePrincipal, application).
const OWNER_TYPE_PROPERTY = "@odata.type";

// Compact rows: object id, client ID and display name. Both identifiers ride
// together so appId is never mistaken for the object id.
const DEFAULT_APP_LIST_SELECT = ["id", "appId", "displayName"];
const DEFAULT_SP_LIST_SELECT = ["id", "appId", "displayName"];
// Show rows: the full reviewed set including credential expiry metadata.
const DEFAULT_APP_SHOW_SELECT = [...KNOWN_APP_FIELDS];
const DEFAULT_SP_SHOW_SELECT = [...KNOWN_SP_FIELDS];
// Compact owner rows: identifier, display name and address.
const DEFAULT_OWNER_SELECT = ["id", "displayName", "mail"];
// Application.Read.All covers applications, service principals and owners in
// both modes; delegated reads default to it unless --scopes overrides it.
export const DEFAULT_DELEGATED_SCOPES = ["https://graph.microsoft.com/Application.Read.All"];
const TRUNCATE_AT = 500;

// The only operations this slice ever binds. Anything else - secret-minting
// routes, writes, consent grants - is refused before credentials.
const READ_OPERATIONS: Readonly<Record<string, string>> = {
  "GET:/applications": "application",
  "GET:/applications/{application-id}": "application",
  "GET:/applications/{application-id}/owners": "application",
  "GET:/servicePrincipals": "service-principal",
  "GET:/servicePrincipals/{servicePrincipal-id}": "service-principal",
  "GET:/servicePrincipals/{servicePrincipal-id}/owners": "service-principal",
};

function checkReadOperation(operation: SessionOperation, help: string): void {
  // Version stays the session's decision: beta reads remain preview-gated
  // there, exactly like every other named slice. This gate only ensures a
  // fabricated secret-minting or write route can never reach credentials.
  const route = `${operation.method}:${operation.path}`;
  if (operation.method !== "GET" || !Object.hasOwn(READ_OPERATIONS, route)) {
    throw new AxiError(`Refused non-read route ${operation.id}: READ-07 serves only the six catalogued application and service-principal GETs`, "VALIDATION_ERROR", [
      help,
      "Secret-minting routes (addPassword, addKey) and writes are never constructed here; consent grants belong to READ-08",
    ]);
  }
}

export type AppFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown application property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known application properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: AppFlags, profile: AnyProfile, help: string): string[] | undefined {
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

function truncateValue(value: unknown, full: boolean): { value: unknown; truncated: boolean } {
  const truncate = (text: string) => full || text.length <= TRUNCATE_AT
    ? { value: text, truncated: false }
    : { value: `${text.slice(0, TRUNCATE_AT)}... (truncated, ${text.length} chars total)`, truncated: true };
  if (typeof value === "string") return truncate(value);
  if (Array.isArray(value) && value.every(entry => typeof entry === "string")) {
    const projected = value.map(truncate);
    return { value: projected.map(entry => entry.value), truncated: projected.some(entry => entry.truncated) };
  }
  return { value, truncated: false };
}

// Credential entries collapse to expiry metadata: only the safe subfields
// survive, each still subject to text truncation. A non-object entry is
// dropped rather than reinterpreted.
function projectCredentials(value: unknown, full: boolean): { value: unknown[]; truncated: boolean } {
  if (!Array.isArray(value)) return { value: [], truncated: false };
  const projected: unknown[] = [];
  let truncated = false;
  for (const entry of value) {
    if (entry === null || typeof entry !== "object" || Array.isArray(entry)) continue;
    const safe: Record<string, unknown> = {};
    for (const field of SAFE_CREDENTIAL_FIELDS) {
      if (!Object.hasOwn(entry, field)) continue;
      const result = truncateValue((entry as Record<string, unknown>)[field], full);
      safe[field] = result.value;
      truncated = truncated || result.truncated;
    }
    projected.push(safe);
  }
  return { value: projected, truncated };
}

// Local projection preserves Graph's null/missing distinction: an explicit
// null stays null, an absent property stays absent and is never synthesized.
// Credential collections additionally collapse to expiry metadata, so secret
// subfields can never reach output.
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
    if (field === "keyCredentials" || field === "passwordCredentials") {
      const result = projectCredentials(source[field], full);
      projected[field] = result.value;
      truncated = truncated || result.truncated;
      continue;
    }
    const result = truncateValue(source[field], full);
    projected[field] = result.value;
    truncated = truncated || result.truncated;
  }
  return { row: projected, truncated };
}

// Owner projection additionally preserves @odata.type so the owner kind
// survives local projection.
function projectOwner(
  row: unknown,
  select: string[],
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean; limitedInfo: boolean } {
  const { row: projected, truncated } = project(row, fields, full);
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  if (typeof source[OWNER_TYPE_PROPERTY] === "string") projected[OWNER_TYPE_PROPERTY] = source[OWNER_TYPE_PROPERTY];
  const descriptive = [...new Set([...select, ...Object.keys(source)])]
    .filter(field => field !== "id" && !field.startsWith("@"));
  const limitedInfo = Object.hasOwn(source, "id") && descriptive.length > 0
    && descriptive.every(field => source[field] === null || source[field] === undefined || !Object.hasOwn(source, field));
  return { row: projected, truncated, limitedInfo };
}

function selectedFields(
  flags: AppFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
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

function fullHint(command: string, flags: AppFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
  filter: string | undefined;
}

function collectionCommon(
  session: GraphSession,
  flags: AppFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    known, knownList,
    saved?.["$select"] === undefined ? defaults : fieldList(saved["$select"], known, knownList, "select", help), help);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, scopes: scopesFor(flags, profile, help), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: AppFlags,
  help: string,
): CollectArgs {
  const query: Record<string, string> = { $select: common.select.join(",") };
  if (common.filter !== undefined) {
    query.$filter = common.filter;
    query.$count = "true";
  }
  const args: CollectArgs = { profile, operation, query, scopes: common.scopes };
  if (common.filter !== undefined) args.consistencyLevel = "eventual";
  if (common.cursor !== undefined) args.cursor = common.cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  return args;
}

function resumeHint(profileName: string): string {
  return `Resume losslessly with the same flags plus --cursor <cursor-from-output> ${profileHint(profileName)}`;
}

async function listCollection(
  session: GraphSession,
  flags: AppFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  command: string,
  key: string,
  emptyNoun: string,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  showHint: string,
  params?: Record<string, string>,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const common = collectionCommon(session, flags, known, knownList, defaults, operation, help, profile);
  const args = collectArgs(profile, operation, common, flags, help);
  if (params) args.params = params;
  const result = await session.collect(args);
  const effectiveFlags: AppFlags = { ...flags, select: result.query.$select ?? defaults.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const truncationHints = truncated ? [fullHint(command, effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      [key]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return {
      [key]: rows,
      count,
      help: [
        `0 ${emptyNoun} matched; the absence of results is the answer, not an error`,
        showHint,
      ],
    };
  }
  return { [key]: rows, count, help: [...truncationHints, showHint] };
}

async function showSingle(
  session: GraphSession,
  flags: AppFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  command: string,
  key: string,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  param: string,
  flag: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const { select, fields } = selectedFields(flags, known, knownList, defaults, help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const raw = await session.execute({
    profile,
    operation,
    params: { [param]: String(flags[flag]) },
    query: { $select: select.join(",") },
    scopes,
  });
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed application body", "GRAPH_ERROR", [
      "Single-object reads carry one object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { [key]: row, help: [fullHint(command, flags, profileName)] };
  return { [key]: row };
}

export async function listApplications(
  session: GraphSession,
  flags: AppFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(session, flags, profile, operation, help, profileName,
    "entra application list", "applications", "applications",
    KNOWN_APPS, KNOWN_APP_FIELDS, DEFAULT_APP_LIST_SELECT,
    `mg-axi entra application show --id <application-object-id> ${profileHint(profileName)}`);
}

export async function showApplication(
  session: GraphSession,
  flags: AppFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(session, flags, profile, operation, help, profileName,
    "entra application show", "application",
    KNOWN_APPS, KNOWN_APP_FIELDS, DEFAULT_APP_SHOW_SELECT, "application-id", "id");
}

export async function listServicePrincipals(
  session: GraphSession,
  flags: AppFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(session, flags, profile, operation, help, profileName,
    "entra service-principal list", "servicePrincipals", "service principals",
    KNOWN_SPS, KNOWN_SP_FIELDS, DEFAULT_SP_LIST_SELECT,
    `mg-axi entra service-principal show --id <service-principal-object-id> ${profileHint(profileName)}`);
}

export async function showServicePrincipal(
  session: GraphSession,
  flags: AppFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(session, flags, profile, operation, help, profileName,
    "entra service-principal show", "servicePrincipal",
    KNOWN_SPS, KNOWN_SP_FIELDS, DEFAULT_SP_SHOW_SELECT, "servicePrincipal-id", "id");
}

async function listOwners(
  session: GraphSession,
  flags: AppFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  command: string,
  param: string,
  flag: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const owner = String(flags[flag]);
  if (!owner.trim()) throw new AxiError(`--${flag} needs the owning object ID`, "VALIDATION_ERROR", [help]);
  const common = collectionCommon(session, flags, KNOWN_OWNERS, KNOWN_OWNER_FIELDS, DEFAULT_OWNER_SELECT, operation, help, profile);
  const args = collectArgs(profile, operation, common, flags, help);
  args.params = { [param]: owner };
  const result = await session.collect(args);
  const effectiveFlags: AppFlags = { ...flags, select: result.query.$select ?? DEFAULT_OWNER_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  let limitedInfo = 0;
  for (const row of result.value) {
    const projected = projectOwner(row, common.select, common.fields, common.full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
    if (projected.limitedInfo) limitedInfo += 1;
  }
  const truncationHints = truncated ? [fullHint(command, effectiveFlags, profileName)] : [];
  const limitedHints = limitedInfo > 0
    ? [`${limitedInfo} of ${rows.length} rows have no non-null selected descriptive properties; this may reflect limited read consent or unset properties`]
    : [];
  if (!result.complete) {
    return {
      owners: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, ...limitedHints, resumeHint(profileName)],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return {
      owners: rows,
      count,
      help: [...limitedHints, "0 owners matched; the absence of results is the answer, not an error"],
    };
  }
  return { owners: rows, count, help: [...truncationHints, ...limitedHints] };
}

export async function listApplicationOwners(
  session: GraphSession,
  flags: AppFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listOwners(session, flags, profile, operation, help, profileName,
    "entra application owner list", "application-id", "application");
}

export async function listServicePrincipalOwners(
  session: GraphSession,
  flags: AppFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listOwners(session, flags, profile, operation, help, profileName,
    "entra service-principal owner list", "servicePrincipal-id", "service-principal");
}
