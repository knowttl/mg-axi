import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-04 multi-tenant-organization subfamily: the read mapping behind
// `mg-axi entra multi-tenant-organization show`,
// `mg-axi entra multi-tenant-organization join-request show`,
// `mg-axi entra multi-tenant-organization tenant list` and
// `mg-axi entra multi-tenant-organization tenant count`. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 multiTenantOrganization get, the
// joinRequestRecord get, the tenant list and the multiTenantOrganization,
// joinRequestRecord and member resource pages on 2026-10-05. Every read
// takes D/A MultiTenantOrganization.Read.All as the full-property scope;
// delegated callers pass it as --scopes while the lower-privileged delegated
// MultiTenantOrganization.ReadBasic.All returns displayName and tenantId
// only; application callers need Read.All admin-consented on the configured
// .default audience. Delegated callers additionally need a supported Entra
// role (Security Reader or Global Reader are least-privileged). Personal
// Microsoft accounts are not supported. These reads run in the commercial
// Global service only; sovereign clouds are not supported. The list
// documents $select and $filter only, so --filter passes through as plain
// $filter with no $count or ConsistencyLevel contract. The $count route
// returns a text/plain integer scalar with no --filter/--select/--limit
// contract. A tenant outside any multitenant organization reads the
// container as state inactive with null properties, which is an answer
// rather than an error; at most one multitenant organization exists per
// tenant. The single-member read stays scheduled: its documented least
// privilege is the write scope MultiTenantOrganization.ReadWrite.All in
// both modes. The tenant-lookup functions ship as tenant-information show
// under the mg-ext-04e function-argument contract. Beta stays out. No multi-tenant-organization mutation exists
// in this slice: creation, update, member add/remove and join acceptance
// are writes.

// Every organization property this slice may request or display, matching
// the reviewed multiTenantOrganization resource. Anything else fails before
// credentials.
export const KNOWN_MTO_FIELDS: readonly string[] = [
  "createdDateTime",
  "description",
  "displayName",
  "id",
  "state",
];
// Every join-request property this slice may request or display, matching
// the reviewed joinRequestRecord resource.
export const KNOWN_MTO_JOIN_REQUEST_FIELDS: readonly string[] = [
  "addedByTenantId",
  "id",
  "memberState",
  "role",
  "transitionDetails",
];
// Every member property this slice may request or display, matching the
// reviewed member resource.
export const KNOWN_MTO_TENANT_FIELDS: readonly string[] = [
  "addedByTenantId",
  "addedDateTime",
  "displayName",
  "joinedDateTime",
  "role",
  "state",
  "tenantId",
  "transitionDetails",
];
const KNOWN_MTO = new Set(KNOWN_MTO_FIELDS);
const KNOWN_JOIN_REQUEST = new Set(KNOWN_MTO_JOIN_REQUEST_FIELDS);
const KNOWN_TENANTS = new Set(KNOWN_MTO_TENANT_FIELDS);

// Singleton rows: the full reviewed set; the container carries only five
// scalar properties and the join request five.
const DEFAULT_MTO_SHOW_SELECT = [...KNOWN_MTO_FIELDS];
const DEFAULT_JOIN_REQUEST_SHOW_SELECT = [...KNOWN_MTO_JOIN_REQUEST_FIELDS];
// Compact member rows: the member tenant identifier and name, the role that
// decides management attention and the state that decides participation.
const DEFAULT_TENANT_LIST_SELECT = ["tenantId", "displayName", "role", "state"];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_MTO_SCOPES = ["https://graph.microsoft.com/MultiTenantOrganization.Read.All"];
const TRUNCATE_AT = 500;

export type MultiTenantOrganizationFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown multi-tenant-organization property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: MultiTenantOrganizationFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
  flags: MultiTenantOrganizationFlags,
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

function fullHint(command: string, flags: MultiTenantOrganizationFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the scope, roles and prerequisites that actually unlock it,
// because a 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const MTO_DENIAL_HINTS = [
  "Multi-tenant-organization reads need MultiTenantOrganization.Read.All, passed as --scopes for delegated access or admin-consented for application access; the lower-privileged delegated MultiTenantOrganization.ReadBasic.All returns displayName and tenantId only",
  "Delegated multi-tenant-organization reads additionally need Security Reader or Global Reader, the least-privileged supported Entra roles",
  "Multi-tenant-organization reads run in the commercial Global service; sovereign clouds are not supported for these reads",
  "Personal Microsoft accounts are not supported for multi-tenant-organization reads",
  "Multi-tenant-organization participation needs Entra ID P1; never diagnose licence solely from HTTP 403",
];

const MTO_NOTE = "At most one multitenant organization exists per tenant; a tenant outside any multitenant organization reads empty or inactive, which is an answer rather than an error";

interface Singleton {
  command: string;
  known: Set<string>;
  knownList: readonly string[];
  showSelect: string[];
  rowKey: string;
  noun: string;
}

const ORGANIZATION: Singleton = {
  command: "entra multi-tenant-organization show",
  known: KNOWN_MTO,
  knownList: KNOWN_MTO_FIELDS,
  showSelect: DEFAULT_MTO_SHOW_SELECT,
  rowKey: "multiTenantOrganization",
  noun: "multitenant organization",
};

const JOIN_REQUEST: Singleton = {
  command: "entra multi-tenant-organization join-request show",
  known: KNOWN_JOIN_REQUEST,
  knownList: KNOWN_MTO_JOIN_REQUEST_FIELDS,
  showSelect: DEFAULT_JOIN_REQUEST_SHOW_SELECT,
  rowKey: "multiTenantOrganizationJoinRequest",
  noun: "join request",
};

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
  filter: string | undefined;
}

// Restoring the saved select/filter keeps cursor resumes lossless when
// --select or --filter is omitted. The tenant list passes --filter through
// as plain $filter with no $count or ConsistencyLevel contract.
function collectionCommon(
  session: GraphSession,
  flags: MultiTenantOrganizationFlags,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    KNOWN_TENANTS, KNOWN_MTO_TENANT_FIELDS,
    saved?.["$select"] === undefined ? DEFAULT_TENANT_LIST_SELECT : fieldList(saved["$select"], KNOWN_TENANTS, KNOWN_MTO_TENANT_FIELDS, "select", help), help);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_MTO_SCOPES, profile, help), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: MultiTenantOrganizationFlags,
  help: string,
): CollectArgs {
  const query: Record<string, string> = { $select: common.select.join(",") };
  if (common.filter !== undefined) query.$filter = common.filter;
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

async function showSingleton(
  session: GraphSession,
  flags: MultiTenantOrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  resource: Singleton,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, resource.known, resource.knownList, resource.showSelect, help);
  const scopes = scopesFor(flags, DEFAULT_MTO_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(MTO_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError(`Graph returned a malformed multi-tenant-organization ${resource.noun} body`, "GRAPH_ERROR", [
      `Single-${resource.noun} reads carry one ${resource.noun} object; treat anything else as unknown, not empty`,
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  // The container answers state inactive with null properties when the
  // tenant joined no multitenant organization; that absence is the answer.
  const inactive = resource === ORGANIZATION && row.state === "inactive";
  const hints = [
    ...(truncated ? [fullHint(resource.command, flags, profileName)] : []),
    ...(inactive ? ["This tenant is not part of any multitenant organization (state inactive); the absence of membership is the answer, not an error"] : []),
  ];
  if (hints.length) return { [resource.rowKey]: row, help: [...hints, MTO_NOTE] };
  return { [resource.rowKey]: row };
}

export async function showMultiTenantOrganization(
  session: GraphSession,
  flags: MultiTenantOrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingleton(session, flags, profile, operation, help, profileName, ORGANIZATION);
}

export async function showMultiTenantOrganizationJoinRequest(
  session: GraphSession,
  flags: MultiTenantOrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingleton(session, flags, profile, operation, help, profileName, JOIN_REQUEST);
}

export async function listMultiTenantOrganizationTenants(
  session: GraphSession,
  flags: MultiTenantOrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, operation, help, profile);
  const result = await withGuidance(MTO_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: MultiTenantOrganizationFlags = { ...flags, select: result.query.$select ?? DEFAULT_TENANT_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const countHint = `mg-axi entra multi-tenant-organization tenant count ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra multi-tenant-organization tenant list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      multiTenantOrganizationTenants: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), countHint, MTO_NOTE],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return {
      multiTenantOrganizationTenants: rows,
      count,
      help: ["0 tenants matched; the absence of results is the answer, not an error", MTO_NOTE],
    };
  }
  return { multiTenantOrganizationTenants: rows, count, help: [...truncationHints, countHint, MTO_NOTE] };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute and accepts only
// a non-negative integer. There is no --filter/--select/--limit contract on
// the count: the catalogue declares no such flags and strict input
// validation refuses them before credentials.
export async function countMultiTenantOrganizationTenants(
  session: GraphSession,
  flags: MultiTenantOrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_MTO_SCOPES, profile, help);
  const raw = await withGuidance(MTO_DENIAL_HINTS, () => session.execute({ profile, operation, scopes }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed multi-tenant-organization tenant count body", "GRAPH_ERROR", [
      "Tenant counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  return { multiTenantOrganizationTenantCount: raw };
}
