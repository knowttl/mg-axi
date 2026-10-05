import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { REDACTED } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-03 data-policy-operations subfamily: the read mapping behind
// `mg-axi entra data-policy-operation list/show/count`. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 datapolicyoperation-get operation documentation
// and the dataPolicyOperation resource reference on 2026-10-05,
// corroborated by the sibling user-exportpersonaldata page (same
// User.Export.All least-privileged scope, same Company Administrator-only
// delegated role, same unsupported personal-account row). Every read takes
// D/A User.Export.All plus User.Read.All; delegated callers additionally
// need Company Administrator, the privileged role documented for export
// reads. Personal Microsoft accounts are not supported. No P1/P2
// prerequisite is stated for these reads. The single GET documents $select
// only (plus unreviewed $expand, which stays out); no list operation
// documentation page exists, so the list shares the single-get permission
// contract and resource shape and offers $select only with no --filter.
// The $count scalar takes no --filter/--select/--limit/--cursor. Beta stays
// out ("beta needs its own review") and no mutation lives here: export
// submission (POST /users/{id}/exportPersonalData) belongs to no read slice.
//
// Secrecy by construction: storageLocation carries the Azure Storage SAS URL
// where export data lands. It stays selectable so agents can confirm an
// export completed, but its value is always replaced with the redaction
// marker before projection, so blob URLs and signed links can never reach
// stdout, errors, previews, logs or cursors through this slice. The shared
// session still redacts secret-shaped values before buffering or output.

// Every data-policy-operation property this slice may request or display,
// matching the reviewed resource. Anything else fails before credentials.
export const KNOWN_DATA_POLICY_FIELDS: readonly string[] = [
  "id",
  "completedDateTime",
  "status",
  "storageLocation",
  "userId",
  "submittedDateTime",
  "progress",
];
const KNOWN_DATA_POLICIES = new Set(KNOWN_DATA_POLICY_FIELDS);

// Compact list rows: identifier, lifecycle status, subject and submission
// time agents need to decide what to inspect next. storageLocation stays out
// of the default selection: it is always redacted and only confirms a
// completed export on show.
const DEFAULT_LIST_SELECT = ["id", "status", "userId", "submittedDateTime"];
// Show rows: the full reviewed set, with storageLocation always redacted.
const DEFAULT_SHOW_SELECT = [...KNOWN_DATA_POLICY_FIELDS];
// Both scopes are documented least-privileged for the single GET in both
// modes; delegated callers pass them as --scopes, application access needs
// both admin-consented on the configured .default audience.
export const DEFAULT_DATA_POLICY_SCOPES = [
  "https://graph.microsoft.com/User.Export.All",
  "https://graph.microsoft.com/User.Read.All",
];
const TRUNCATE_AT = 500;

export type DataPolicyFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown data-policy-operation property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known data-policy-operation properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: DataPolicyFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
// storageLocation is always replaced with the redaction marker when present
// as a string, so export blob URLs and signed links never reach output even
// when the server URL carries no sig= sentinel. Truncation never hides the
// marker.
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
    if (field === "storageLocation" && typeof source[field] === "string") {
      projected[field] = REDACTED;
      continue;
    }
    const result = truncateValue(source[field], full);
    projected[field] = result.value;
    truncated = truncated || result.truncated;
  }
  return { row: projected, truncated };
}

function selectedFields(
  flags: DataPolicyFlags,
  defaults: string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, KNOWN_DATA_POLICIES, KNOWN_DATA_POLICY_FIELDS, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, KNOWN_DATA_POLICIES, KNOWN_DATA_POLICY_FIELDS, "fields", help);
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

function fullHint(command: string, flags: DataPolicyFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the scopes, roles and licensing that actually unlock it,
// because a 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const DATA_POLICY_DENIAL_HINTS = [
  "Data-policy-operation reads need User.Export.All plus User.Read.All: delegated callers pass both as --scopes, application access needs both admin-consented on the configured .default audience",
  "Delegated callers additionally need Company Administrator, the privileged role documented for export reads",
  "Personal Microsoft accounts are not supported for data-policy-operation reads",
  "No P1/P2 prerequisite is stated for data-policy-operation reads; never diagnose licence solely from HTTP 403",
];

const WORKFORCE_NOTE = "Workforce tenant context only; export submission (POST /users/{id}/exportPersonalData) belongs to no read slice and is never sent here";

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
}

// Restoring the saved select keeps cursor resumes lossless when --select is
// omitted. The list offers $select only with no --filter: no list operation
// documentation page exists and the single GET documents $select only, so
// strict input validation refuses --filter before credentials instead of
// forwarding a misleading request.
function collectionCommon(
  session: GraphSession,
  flags: DataPolicyFlags,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    saved?.["$select"] === undefined ? DEFAULT_LIST_SELECT : fieldList(saved["$select"], KNOWN_DATA_POLICIES, KNOWN_DATA_POLICY_FIELDS, "select", help), help);
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_DATA_POLICY_SCOPES, profile, help), full: flags.full === true };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: DataPolicyFlags,
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

function operationId(flags: DataPolicyFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the data-policy-operation ID", "VALIDATION_ERROR", [help]);
  return id;
}

function singleResult(
  raw: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed data-policy-operation body", "GRAPH_ERROR", [
      "Single-operation reads carry one dataPolicyOperation object; treat anything else as unknown, not empty",
    ]);
  }
  return project(raw, fields, full);
}

export async function listDataPolicyOperations(
  session: GraphSession,
  flags: DataPolicyFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, operation, help, profile);
  const result = await withGuidance(DATA_POLICY_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: DataPolicyFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIST_SELECT.join(",") };
  const operations: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    operations.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra data-policy-operation show --id <operation-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra data-policy-operation list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      dataPolicyOperations: operations,
      count: { returned: operations.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, WORKFORCE_NOTE],
    };
  }
  const count = { returned: operations.length, complete: true };
  if (!operations.length) {
    return {
      dataPolicyOperations: operations,
      count,
      help: ["0 data-policy operations matched; the absence of results is the answer, not an error", WORKFORCE_NOTE],
    };
  }
  return { dataPolicyOperations: operations, count, help: [...truncationHints, showHint, WORKFORCE_NOTE] };
}

export async function showDataPolicyOperation(
  session: GraphSession,
  flags: DataPolicyFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, DEFAULT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_DATA_POLICY_SCOPES, profile, help);
  const full = flags.full === true;
  const id = operationId(flags, help);
  const raw = await withGuidance(DATA_POLICY_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "dataPolicyOperation-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full);
  if (truncated) return { dataPolicyOperation: row, help: [fullHint("entra data-policy-operation show", flags, profileName), WORKFORCE_NOTE] };
  return { dataPolicyOperation: row, help: [WORKFORCE_NOTE] };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute and accepts only
// a non-negative integer. There is no --filter/--select/--limit/--cursor
// contract on the count: the catalogue declares no such flags and strict
// input validation refuses them before credentials.
export async function countDataPolicyOperations(
  session: GraphSession,
  flags: DataPolicyFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_DATA_POLICY_SCOPES, profile, help);
  const raw = await withGuidance(DATA_POLICY_DENIAL_HINTS, () => session.execute({ profile, operation, scopes, scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed data-policy-operation count body", "GRAPH_ERROR", [
      "Data-policy-operation counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  return { count: { returned: raw, complete: true }, help: [WORKFORCE_NOTE] };
}
