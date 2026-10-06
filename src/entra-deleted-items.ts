import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 deleted-items subfamily: the read mapping behind
// `mg-axi entra deleted-user|deleted-group|deleted-application|
// deleted-service-principal|deleted-administrative-unit list/count` and
// `mg-axi entra deleted-item show`. Operation construction stays beside its
// command; the shared session owns URLs, credentials, paging, retries and
// error translation, and the SDK owns TOON rendering. This module only maps
// flags to session calls and projects rows for compact output.
//
// Reviewed against the v1.0 directory-deleteditems-list and
// directory-deleteditems-get operation documentation on 2026-10-05. The list
// page documents per-type least-privilege permissions (delegated and
// application share the same scope per type; personal Microsoft accounts are
// not supported): user takes User.Read.All, group takes Group.Read.All,
// application and servicePrincipal take Application.Read.All, and
// administrativeUnit takes AdministrativeUnit.Read.All. No delegated role
// and no P1/P2 prerequisite are stated for these reads. The OData cast is a
// required part of the list URI, so untyped list/count commands do not exist
// and device casts carry no documented v1.0 permission contract and stay
// out. The $count scalar follows its typed list's permission contract.
// Deleted users are personal data and deleted applications carry credential
// metadata, so defaults stay minimal and key/password credential collections
// are never selectable: GET never returns secret values and success
// redaction still applies. Restore, permanent delete, the POST
// lookup/validation actions and beta stay out. No mutation lives here.
//
// Show reads the single documented untyped get
// (GET /directory/deletedItems/{directoryObject-id}), which can return any
// deletable type. Its $select is therefore limited to the cross-type safe
// properties valid on every deletable type; per-type detail comes from the
// typed lists. The object's type is unknown before the read, so show takes
// no default scope: delegated callers pass the scope matching the object's
// type as --scopes. The @odata.type discriminator rides along automatically
// (envelope metadata, never selectable) so rows still name their kind.

// Every deleted-user property this slice may request or display. Phones,
// photos and other PII beyond what identifies the restore target stay out;
// anything else fails before credentials.
export const KNOWN_DELETED_USER_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "userPrincipalName",
  "mail",
  "userType",
  "deletedDateTime",
];
// Every deleted-group property this slice may request or display. Soft
// deleted security groups report securityEnabled false through a known
// upstream limitation, so groupTypes names the real kind.
export const KNOWN_DELETED_GROUP_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "mail",
  "mailNickname",
  "groupTypes",
  "visibility",
  "securityEnabled",
  "deletedDateTime",
];
// Every deleted-application property this slice may request or display.
// appId (client ID) is distinct from the object id. keyCredentials and
// passwordCredentials are never selectable: credential metadata stays
// minimal and secret values can never appear.
export const KNOWN_DELETED_APPLICATION_FIELDS: readonly string[] = [
  "id",
  "appId",
  "displayName",
  "publisherDomain",
  "signInAudience",
  "deletedDateTime",
];
// Every deleted-service-principal property this slice may request or
// display. appId (client ID) is distinct from the object id; secret-bearing
// fields are never selectable.
export const KNOWN_DELETED_SERVICE_PRINCIPAL_FIELDS: readonly string[] = [
  "id",
  "appId",
  "displayName",
  "servicePrincipalType",
  "accountEnabled",
  "deletedDateTime",
];
// Every deleted-administrative-unit property this slice may request or
// display.
export const KNOWN_DELETED_ADMINISTRATIVE_UNIT_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "description",
  "visibility",
  "membershipType",
  "deletedDateTime",
];
// Every deleted-item show property this slice may request or display: the
// cross-type safe intersection, valid on whatever deletable type the
// untyped get returns. Per-type detail comes from the typed lists.
export const KNOWN_DELETED_SHOW_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "deletedDateTime",
];
// The subtype discriminator rides on every row without being selectable: it
// names the concrete kind (user, group, application and the rest).
const DELETED_TYPE_PROPERTY = "@odata.type";

// Compact list rows carry the identifier, the display name and the deletion
// timestamp; applications and service principals additionally carry appId so
// the client ID is never conflated with the object id.
export const DEFAULT_DELETED_USER_LIST_SELECT = ["id", "displayName", "deletedDateTime"];
export const DEFAULT_DELETED_GROUP_LIST_SELECT = ["id", "displayName", "deletedDateTime"];
export const DEFAULT_DELETED_APPLICATION_LIST_SELECT = ["id", "appId", "displayName", "deletedDateTime"];
export const DEFAULT_DELETED_SERVICE_PRINCIPAL_LIST_SELECT = ["id", "appId", "displayName", "deletedDateTime"];
export const DEFAULT_DELETED_ADMINISTRATIVE_UNIT_LIST_SELECT = ["id", "displayName", "deletedDateTime"];
// Show rows carry the full cross-type safe set.
const DEFAULT_SHOW_SELECT = [...KNOWN_DELETED_SHOW_FIELDS];
// Delegated list/count defaults are per-type least privilege; application
// profiles use their configured .default audience and reject --scopes. All
// four scopes are already in the shared READ_SCOPES allowlist, and every
// one is a read scope, so no new scope and no deferral are needed here.
export const DELETED_ITEM_SCOPES = {
  user: ["https://graph.microsoft.com/User.Read.All"],
  group: ["https://graph.microsoft.com/Group.Read.All"],
  application: ["https://graph.microsoft.com/Application.Read.All"],
  "service-principal": ["https://graph.microsoft.com/Application.Read.All"],
  "administrative-unit": ["https://graph.microsoft.com/AdministrativeUnit.Read.All"],
} as const satisfies Record<string, readonly string[]>;
export type DeletedItemType = keyof typeof DELETED_ITEM_SCOPES;
// Every scope a deleted-item show caller may need to pass, one per deletable
// type. Show takes no default because the object's type is unknown before
// the read.
export const DELETED_SHOW_SCOPE_CHOICES: readonly string[] = [
  "https://graph.microsoft.com/User.Read.All",
  "https://graph.microsoft.com/Group.Read.All",
  "https://graph.microsoft.com/Application.Read.All",
  "https://graph.microsoft.com/AdministrativeUnit.Read.All",
];
const TRUNCATE_AT = 500;

const TYPE_FIELDS: Readonly<Record<DeletedItemType, readonly string[]>> = {
  user: KNOWN_DELETED_USER_FIELDS,
  group: KNOWN_DELETED_GROUP_FIELDS,
  application: KNOWN_DELETED_APPLICATION_FIELDS,
  "service-principal": KNOWN_DELETED_SERVICE_PRINCIPAL_FIELDS,
  "administrative-unit": KNOWN_DELETED_ADMINISTRATIVE_UNIT_FIELDS,
};
const TYPE_LIST_DEFAULTS: Readonly<Record<DeletedItemType, readonly string[]>> = {
  user: DEFAULT_DELETED_USER_LIST_SELECT,
  group: DEFAULT_DELETED_GROUP_LIST_SELECT,
  application: DEFAULT_DELETED_APPLICATION_LIST_SELECT,
  "service-principal": DEFAULT_DELETED_SERVICE_PRINCIPAL_LIST_SELECT,
  "administrative-unit": DEFAULT_DELETED_ADMINISTRATIVE_UNIT_LIST_SELECT,
};

// The only operations this slice ever binds. Anything else - the untyped
// list/count (upstream requires the cast), device casts (no documented v1.0
// permission contract), typed-cast singles (covered by the untyped show),
// POST lookup/validation actions, restore, permanent delete, beta - is
// refused before credentials.
const LIST_ROUTES: Readonly<Record<string, DeletedItemType>> = {
  "GET:/directory/deletedItems/graph.user": "user",
  "GET:/directory/deletedItems/graph.group": "group",
  "GET:/directory/deletedItems/graph.application": "application",
  "GET:/directory/deletedItems/graph.servicePrincipal": "service-principal",
  "GET:/directory/deletedItems/graph.administrativeUnit": "administrative-unit",
};
const COUNT_ROUTES: Readonly<Record<string, DeletedItemType>> = {
  "GET:/directory/deletedItems/graph.user/$count": "user",
  "GET:/directory/deletedItems/graph.group/$count": "group",
  "GET:/directory/deletedItems/graph.application/$count": "application",
  "GET:/directory/deletedItems/graph.servicePrincipal/$count": "service-principal",
  "GET:/directory/deletedItems/graph.administrativeUnit/$count": "administrative-unit",
};
const SHOW_ROUTE = "GET:/directory/deletedItems/{directoryObject-id}";

function routeOf(operation: SessionOperation): string {
  return `${operation.method}:${operation.path}`;
}

export type DeletedItemFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: readonly string[], flag: string, help: string): string[] {
  const knownSet = new Set(known);
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!knownSet.has(part)) {
      throw new AxiError(`Unknown deleted-item property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known deleted-item properties: ${known.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function listScopesFor(flags: DeletedItemFlags, type: DeletedItemType, profile: AnyProfile, help: string): string[] | undefined {
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
  if (flags.scopes === undefined) return [...DELETED_ITEM_SCOPES[type]];
  const scopes = String(flags.scopes)
    .split(",")
    .map(scope => scope.trim())
    .filter(scope => scope.length > 0);
  return scopes;
}

// Show takes no default scope: the untyped get can return any deletable
// type, so the caller passes the scope matching the object's type.
function showScopesFor(flags: DeletedItemFlags, profile: AnyProfile, help: string): string[] | undefined {
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
  if (flags.scopes === undefined) {
    throw new AxiError("Deleted-item show needs an explicit --scopes matching the object's type", "VALIDATION_ERROR", [
      help,
      `Pass one of: ${DELETED_SHOW_SCOPE_CHOICES.join(", ")}`,
      "List the recycle bin first: the @odata.type discriminator names each row's kind",
    ]);
  }
  const scopes = String(flags.scopes)
    .split(",")
    .map(scope => scope.trim())
    .filter(scope => scope.length > 0);
  if (!scopes.length) throw new AxiError("--scopes needs at least one scope", "VALIDATION_ERROR", [help]);
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

// Local projection preserves Graph's null/missing distinction: an explicit
// null stays null, an absent property stays absent and is never synthesized.
// The subtype discriminator rides along automatically when the server sent
// it; unreviewed properties never do, so secrets or key material any
// subtype carries stay out.
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
  if (typeof source[DELETED_TYPE_PROPERTY] === "string") projected[DELETED_TYPE_PROPERTY] = source[DELETED_TYPE_PROPERTY];
  return { row: projected, truncated };
}

function selectedFields(
  flags: DeletedItemFlags,
  known: readonly string[],
  defaults: readonly string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, known, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, known, "fields", help);
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

function fullHint(command: string, flags: DeletedItemFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each type
// adds the scope that actually unlocks it, because a 403 alone never says
// which prerequisite is missing. No delegated role and no P1/P2
// prerequisite are stated for these reads.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

function denialHints(type: DeletedItemType): string[] {
  const scope = DELETED_ITEM_SCOPES[type][0];
  return [
    `Deleted-${type} reads need ${scope} for delegated or application access; delegated callers pass it as --scopes`,
    "Personal Microsoft accounts are not supported for deleted-item reads",
    "No delegated role or P1/P2 prerequisite is stated for deleted-item reads; never diagnose role or licence solely from HTTP 403",
  ];
}

const SHOW_DENIAL_HINTS = [
  "Deleted-item show needs the scope matching the object's type (User.Read.All, Group.Read.All, Application.Read.All or AdministrativeUnit.Read.All): delegated callers pass it as --scopes, application access needs the admin-consented scope on the configured .default audience",
  "Personal Microsoft accounts are not supported for deleted-item reads",
  "No delegated role or P1/P2 prerequisite is stated for this read; never diagnose role or licence solely from HTTP 403",
];

const SHOW_SCOPE_NOTE = "Show takes no default scope: pass the scope matching the object's type as --scopes; list the recycle bin first and read each row's @odata.type kind";

function typeCommand(type: DeletedItemType, verb: "list" | "count"): string {
  return `mg-axi entra deleted-${type} ${verb}`;
}

function showCommand(profileName: string): string {
  return `mg-axi entra deleted-item show --id <object-id> ${profileHint(profileName)}`;
}

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
}

// Restoring the saved select keeps cursor resumes lossless when --select is
// omitted. The typed lists document no reviewed $filter/$orderby contract
// here, so the catalogue declares no such flags and strict input validation
// refuses them before credentials instead of forwarding an unreviewed query.
function collectionCommon(
  session: GraphSession,
  flags: DeletedItemFlags,
  type: DeletedItemType,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const known = TYPE_FIELDS[type];
  const defaults = TYPE_LIST_DEFAULTS[type];
  const { select, fields } = selectedFields(flags,
    known,
    saved?.["$select"] === undefined ? [...defaults] : fieldList(saved["$select"], known, "select", help), help);
  return { cursor, select, fields, scopes: listScopesFor(flags, type, profile, help), full: flags.full === true };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: DeletedItemFlags,
  help: string,
): CollectArgs {
  const query: Record<string, string> = { $select: common.select.join(",") };
  const args: CollectArgs = { profile, operation, query, scopes: common.scopes };
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

function deletedId(flags: DeletedItemFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the deleted object ID", "VALIDATION_ERROR", [help]);
  return id;
}

function listType(operation: SessionOperation, help: string): DeletedItemType {
  const type = LIST_ROUTES[routeOf(operation)];
  if (operation.method !== "GET" || type === undefined) {
    throw new AxiError(`Refused non-read route ${operation.id}: deleted-item lists serve only the five catalogued v1.0 typed deletedItems GETs`, "VALIDATION_ERROR", [
      help,
      "Untyped lists need the OData cast upstream requires; device casts, POST lookups, restore and beta need their own review",
    ]);
  }
  return type;
}

function countType(operation: SessionOperation, help: string): DeletedItemType {
  const type = COUNT_ROUTES[routeOf(operation)];
  if (operation.method !== "GET" || type === undefined) {
    throw new AxiError(`Refused non-read route ${operation.id}: deleted-item counts serve only the five catalogued v1.0 typed deletedItems $count GETs`, "VALIDATION_ERROR", [
      help,
      "Untyped counts need the OData cast upstream requires; device casts and beta need their own review",
    ]);
  }
  return type;
}

function checkShowOperation(operation: SessionOperation, help: string): void {
  if (operation.method !== "GET" || routeOf(operation) !== SHOW_ROUTE) {
    throw new AxiError(`Refused non-read route ${operation.id}: deleted-item show serves only the catalogued v1.0 untyped deletedItems get`, "VALIDATION_ERROR", [
      help,
      "Typed-cast singles are covered by this same untyped get; restore, permanent delete and beta need their own review",
    ]);
  }
}

export async function listDeletedItems(
  session: GraphSession,
  flags: DeletedItemFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const type = listType(operation, help);
  const common = collectionCommon(session, flags, type, operation, help, profile);
  const result = await withGuidance(denialHints(type), () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: DeletedItemFlags = { ...flags, select: result.query.$select ?? TYPE_LIST_DEFAULTS[type].join(",") };
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const command = typeCommand(type, "list");
  const showHint = showCommand(profileName);
  const groupNote = type === "group"
    ? ["Soft-deleted security groups report securityEnabled false through a known upstream limitation; read groupTypes to name the real kind"]
    : [];
  const appNote = type === "application" || type === "service-principal"
    ? ["appId (client ID) is distinct from the object id; credential collections are never selectable and GET never returns secret values"]
    : [];
  const emptyNote = `0 deleted ${type === "service-principal" ? "service principals" : `${type}s`} matched; the absence of results is the answer, not an error`;
  const truncationHints = truncated ? [fullHint(command, effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      deletedItems: rows,
      ...listTotals(rows.length, result.total, "deleted items", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, ...groupNote, ...appNote],
    };
  }
  if (!rows.length) {
    return { deletedItems: rows, ...listTotals(rows.length, result.total, "deleted items", true), complete: true, help: [emptyNote, ...groupNote, ...appNote] };
  }
  return { deletedItems: rows, ...listTotals(rows.length, result.total, "deleted items", true), complete: true, help: [...truncationHints, showHint, ...groupNote, ...appNote] };
}

export async function showDeletedItem(
  session: GraphSession,
  flags: DeletedItemFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkShowOperation(operation, help);
  const { select, fields } = selectedFields(flags, KNOWN_DELETED_SHOW_FIELDS, DEFAULT_SHOW_SELECT, help);
  const scopes = showScopesFor(flags, profile, help);
  const full = flags.full === true;
  const id = deletedId(flags, help);
  const raw = await withGuidance(SHOW_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "directoryObject-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed deleted-item body", "GRAPH_ERROR", [
      "Single-object reads carry one directoryObject; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  const helpHints: string[] = [...(truncated ? [fullHint("entra deleted-item show", flags, profileName)] : []), SHOW_SCOPE_NOTE];
  if (typeof row["@odata.type"] === "string" && row["@odata.type"].includes("group")) {
    helpHints.push("Soft-deleted security groups report securityEnabled false through a known upstream limitation; read groupTypes to name the real kind");
  }
  return { deletedItem: row, help: helpHints };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute and accepts only
// a non-negative integer. There is no --filter/--select/--limit contract on
// the count: the catalogue declares no such flags and strict input
// validation refuses them before credentials.
export async function countDeletedItems(
  session: GraphSession,
  flags: DeletedItemFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const type = countType(operation, help);
  const scopes = listScopesFor(flags, type, profile, help);
  const raw = await withGuidance(denialHints(type), () => session.execute({ profile, operation, scopes, scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed deleted-item count body", "GRAPH_ERROR", [
      "Deleted-item counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  const label = type === "service-principal" ? "service principals" : `${type}s`;
  const count = { returned: raw, complete: true };
  return raw === 0
    ? { count, help: [`0 deleted ${label} in the recycle bin; the absence of results is the answer, not an error`] }
    : { count };
}
