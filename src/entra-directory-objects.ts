import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 directory-objects subfamily: the read mapping behind
// `mg-axi entra directory-object list/show/count`. Operation construction
// stays beside its command; the shared session owns URLs, credentials,
// paging, retries and error translation, and the SDK owns TOON rendering.
// This module only maps flags to session calls and projects rows for compact
// output.
//
// Reviewed against the v1.0 directoryobject-get operation documentation and
// the directoryObject resource reference on 2026-10-05. There is no dedicated
// List operation page: the resource Methods table documents Get, delta,
// Delete and the lookup/validation functions only, so the collection GET
// follows the sibling get contract and the resource reference (both name
// Directory.Read.All), with $select-only queries because collection filter
// support is unreviewed. Both reads take delegated or application
// Directory.Read.All, the documented least privilege in each mode; personal
// Microsoft accounts are not supported. No delegated role and no P1/P2
// prerequisite are stated for these reads. The $count scalar carries no
// operation-level documentation page and follows the same Directory.Read.All
// contract. The delta() sync, every POST lookup/validation action, beta and
// every mutation stay out. No mutation lives here.
//
// Secrecy by construction: directory objects are polymorphic, so only the
// base-type properties (id, deletedDateTime) are ever requested or
// projected; $expand is not offered and subtype secrets or credentials any
// subtype carries can never appear. @odata.type rides along automatically
// (envelope metadata, never selectable) so rows still name their subtype.

// Every directory-object property this slice may request or display,
// matching the reviewed base resource. Anything else fails before
// credentials. @odata.type is preserved separately and is never selectable.
export const KNOWN_DIRECTORY_OBJECT_FIELDS: readonly string[] = [
  "id",
  "deletedDateTime",
];
const KNOWN_OBJECTS = new Set(KNOWN_DIRECTORY_OBJECT_FIELDS);
// The subtype discriminator rides on every row without being selectable: it
// names the concrete kind (user, group, servicePrincipal and the rest).
const OBJECT_TYPE_PROPERTY = "@odata.type";

// Compact list rows carry the identifier only; show rows carry the full
// reviewed base set. The discriminator is attached to both automatically.
const DEFAULT_LIST_SELECT = ["id"];
const DEFAULT_SHOW_SELECT = [...KNOWN_DIRECTORY_OBJECT_FIELDS];
// Directory.Read.All covers list, show and count in both modes; delegated
// reads default to it unless --scopes overrides it. It is already in the
// shared READ_SCOPES allowlist, and it is a read scope, so no new scope and
// no deferral are needed here.
export const DEFAULT_DIRECTORY_OBJECT_SCOPES = ["https://graph.microsoft.com/Directory.Read.All"];
const TRUNCATE_AT = 500;

// The only operations this slice ever binds. Anything else - delta sync,
// lookup/validation actions, writes - is refused before credentials.
const READ_OPERATIONS: Readonly<Record<string, string>> = {
  "GET:/directoryObjects": "directory-object",
  "GET:/directoryObjects/{directoryObject-id}": "directory-object",
  "GET:/directoryObjects/$count": "directory-object",
};

function checkReadOperation(operation: SessionOperation, help: string): void {
  const route = `${operation.method}:${operation.path}`;
  if (operation.method !== "GET" || !Object.hasOwn(READ_OPERATIONS, route)) {
    throw new AxiError(`Refused non-read route ${operation.id}: directory-object serves only the three catalogued v1.0 directoryObjects GETs`, "VALIDATION_ERROR", [
      help,
      "Delta sync, POST lookup/validation actions and writes are never constructed here; beta needs its own review",
    ]);
  }
}

export type DirectoryObjectFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown directory-object property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known directory-object properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: DirectoryObjectFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
// it; subtype properties never do, so secrets any subtype carries stay out.
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
  if (typeof source[OBJECT_TYPE_PROPERTY] === "string") projected[OBJECT_TYPE_PROPERTY] = source[OBJECT_TYPE_PROPERTY];
  return { row: projected, truncated };
}

function selectedFields(
  flags: DirectoryObjectFlags,
  defaults: string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, KNOWN_OBJECTS, KNOWN_DIRECTORY_OBJECT_FIELDS, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, KNOWN_OBJECTS, KNOWN_DIRECTORY_OBJECT_FIELDS, "fields", help);
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

function fullHint(command: string, flags: DirectoryObjectFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the scope, roles and licensing that actually unlock it,
// because a 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const DIRECTORY_OBJECT_DENIAL_HINTS = [
  "Directory-object reads need Directory.Read.All for delegated or application access; delegated callers pass it as --scopes",
  "Personal Microsoft accounts are not supported for directory-object reads",
  "No delegated role or P1/P2 prerequisite is stated for directory-object reads; never diagnose role or licence solely from HTTP 403",
];

const EMPTY_NOTE = "0 directory objects returned; every tenant directory carries objects, so verify the profile tenant before treating this as an empty directory";

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
}

// Restoring the saved select keeps cursor resumes lossless when --select is
// omitted. The collection documents no List operation page, so it offers no
// --filter: strict input validation refuses the flag before credentials
// instead of forwarding an unreviewed query.
function collectionCommon(
  session: GraphSession,
  flags: DirectoryObjectFlags,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    saved?.["$select"] === undefined ? DEFAULT_LIST_SELECT : fieldList(saved["$select"], KNOWN_OBJECTS, KNOWN_DIRECTORY_OBJECT_FIELDS, "select", help), help);
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_DIRECTORY_OBJECT_SCOPES, profile, help), full: flags.full === true };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: DirectoryObjectFlags,
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

function objectId(flags: DirectoryObjectFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the directory-object object ID", "VALIDATION_ERROR", [help]);
  return id;
}

function singleResult(
  raw: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed directory-object body", "GRAPH_ERROR", [
      "Single-object reads carry one directoryObject; treat anything else as unknown, not empty",
    ]);
  }
  return project(raw, fields, full);
}

export async function listDirectoryObjects(
  session: GraphSession,
  flags: DirectoryObjectFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const common = collectionCommon(session, flags, operation, help, profile);
  const result = await withGuidance(DIRECTORY_OBJECT_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: DirectoryObjectFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIST_SELECT.join(",") };
  const directoryObjects: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    directoryObjects.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra directory-object show --id <object-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra directory-object list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      directoryObjects,
      count: { returned: directoryObjects.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  const count = { returned: directoryObjects.length, complete: true };
  if (!directoryObjects.length) {
    return { directoryObjects, count, help: [EMPTY_NOTE] };
  }
  return { directoryObjects, count, help: [...truncationHints, showHint] };
}

export async function showDirectoryObject(
  session: GraphSession,
  flags: DirectoryObjectFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const { select, fields } = selectedFields(flags, DEFAULT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_DIRECTORY_OBJECT_SCOPES, profile, help);
  const full = flags.full === true;
  const id = objectId(flags, help);
  const raw = await withGuidance(DIRECTORY_OBJECT_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "directoryObject-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full);
  if (truncated) return { directoryObject: row, help: [fullHint("entra directory-object show", flags, profileName)] };
  return { directoryObject: row };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute and accepts only
// a non-negative integer. There is no --filter/--select/--limit contract on
// the count: the catalogue declares no such flags and strict input
// validation refuses them before credentials.
export async function countDirectoryObjects(
  session: GraphSession,
  flags: DirectoryObjectFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  checkReadOperation(operation, help);
  const scopes = scopesFor(flags, DEFAULT_DIRECTORY_OBJECT_SCOPES, profile, help);
  const raw = await withGuidance(DIRECTORY_OBJECT_DENIAL_HINTS, () => session.execute({ profile, operation, scopes, scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed directory-object count body", "GRAPH_ERROR", [
      "Directory-object counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  const count = { returned: raw, complete: true };
  return raw === 0 ? { count, help: [EMPTY_NOTE] } : { count };
}
