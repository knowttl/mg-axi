import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 on-premises-synchronization subfamily: the read mapping behind
// `mg-axi entra on-premises-synchronization list/show`. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 onpremisesdirectorysynchronization-get
// operation documentation and the onPremisesDirectorySynchronization
// resource reference on 2026-10-05, corroborated by the sibling
// onpremisesdirectorysynchronization-update page (same permission table,
// same Global Administrator-only role, same unsupported application and
// personal-account rows). Both reads take delegated
// OnPremDirectorySynchronization.Read.All; Graph documents no supported
// application permission for this operation, so application profiles are
// refused before credentials rather than sent to a certain denial.
// Delegated callers additionally need Global Administrator, the only
// supported role (a custom role with a supported permission also works).
// Personal Microsoft accounts are not supported. No P1/P2 prerequisite is
// stated for these reads. Every reviewed read documents $select only, so
// neither leaf offers --filter. The $count scalar stays scheduled for a
// later counts subfamily, beta stays out ("beta needs its own review") and
// no mutation lives here.
//
// Secrecy by construction: the reviewed set carries the tenant id plus the
// configuration and features objects only, and the shared session still
// redacts secret-shaped values (credential, secret, key and token shapes)
// before buffering or output, so sync configuration never leaks keys or
// passwords through these reads.

// Every on-premises-synchronization property this slice may request or
// display, matching the reviewed resource. Anything else fails before
// credentials.
export const KNOWN_SYNC_FIELDS: readonly string[] = [
  "id",
  "configuration",
  "features",
];
const KNOWN_SYNCS = new Set(KNOWN_SYNC_FIELDS);

// The resource carries only three properties, so compact list rows and show
// rows share the full reviewed set: the tenant id plus the sync
// configuration and feature flags agents need next.
const DEFAULT_LIST_SELECT = [...KNOWN_SYNC_FIELDS];
const DEFAULT_SHOW_SELECT = [...KNOWN_SYNC_FIELDS];
// Delegated default; application profiles are refused before credentials
// because Graph documents no supported application permission here.
export const DEFAULT_SYNC_SCOPES = ["https://graph.microsoft.com/OnPremDirectorySynchronization.Read.All"];
const TRUNCATE_AT = 500;

export type SyncFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown on-premises-synchronization property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known on-premises-synchronization properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: SyncFlags, defaults: string[], profile: AnyProfile, help: string): string[] {
  if (profile.mode === "application") {
    throw new AxiError(
      "On-premises-synchronization reads need a delegated profile; Graph documents no supported application permission for this operation",
      "VALIDATION_ERROR",
      [help, "mg-axi profile create --name <name> --tenant <tenant-id> --client <client-id> --cloud commercial"],
    );
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
  flags: SyncFlags,
  defaults: string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, KNOWN_SYNCS, KNOWN_SYNC_FIELDS, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, KNOWN_SYNCS, KNOWN_SYNC_FIELDS, "fields", help);
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

function fullHint(command: string, flags: SyncFlags, profileName: string): string {
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

const SYNC_DENIAL_HINTS = [
  "On-premises-synchronization reads need delegated OnPremDirectorySynchronization.Read.All plus Global Administrator, the only supported Entra role for this operation (a custom role with a supported permission also works); Graph documents no supported application permission, so application profiles are refused before credentials",
  "Personal Microsoft accounts are not supported for on-premises-synchronization reads",
  "No P1/P2 prerequisite is stated for on-premises-synchronization reads; never diagnose licence solely from HTTP 403",
];

const EMPTY_NOTE = "An empty on-premises-synchronization list may mean on-premises directory sync is not configured for the tenant; the absence of rows is the answer, not an error";

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[];
  full: boolean;
}

// Restoring the saved select keeps cursor resumes lossless when --select is
// omitted. The list documents $select only, so it offers no --filter:
// strict input validation refuses the flag before credentials instead of
// forwarding a misleading request.
function collectionCommon(
  session: GraphSession,
  flags: SyncFlags,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    saved?.["$select"] === undefined ? DEFAULT_LIST_SELECT : fieldList(saved["$select"], KNOWN_SYNCS, KNOWN_SYNC_FIELDS, "select", help), help);
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_SYNC_SCOPES, profile, help), full: flags.full === true };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: SyncFlags,
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

function syncId(flags: SyncFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the on-premises-synchronization object ID (the Entra tenant ID)", "VALIDATION_ERROR", [help]);
  return id;
}

function singleResult(
  raw: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed on-premises-synchronization body", "GRAPH_ERROR", [
      "Single-synchronization reads carry one onPremisesDirectorySynchronization object; treat anything else as unknown, not empty",
    ]);
  }
  return project(raw, fields, full);
}

export async function listSynchronizations(
  session: GraphSession,
  flags: SyncFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, operation, help, profile);
  const result = await withGuidance(SYNC_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: SyncFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIST_SELECT.join(",") };
  const synchronizations: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    synchronizations.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra on-premises-synchronization show --id <synchronization-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra on-premises-synchronization list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      synchronizations,
      count: { returned: synchronizations.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  const count = { returned: synchronizations.length, complete: true };
  if (!synchronizations.length) {
    return { synchronizations, count, help: [EMPTY_NOTE] };
  }
  return { synchronizations, count, help: [...truncationHints, showHint] };
}

export async function showSynchronization(
  session: GraphSession,
  flags: SyncFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, DEFAULT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_SYNC_SCOPES, profile, help);
  const full = flags.full === true;
  const id = syncId(flags, help);
  const raw = await withGuidance(SYNC_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "onPremisesDirectorySynchronization-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full);
  if (truncated) return { synchronization: row, help: [fullHint("entra on-premises-synchronization show", flags, profileName)] };
  return { synchronization: row };
}
