import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-04 delegated-admin subfamily: the read mapping behind
// `mg-axi entra delegated-admin-customer list/show` and
// `mg-axi entra delegated-admin-relationship list/show`. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 delegatedAdminCustomer list/get, the
// delegatedAdminRelationship list/get and both resource pages on
// 2026-10-05. Every read takes D/A DelegatedAdminRelationship.Read.All as
// least privilege; the scope is already in the shared READ_SCOPES allowlist,
// so no new scope and no deferral is needed here. Delegated callers pass it
// as --scopes; application callers need it admin-consented on the configured
// .default audience. The list/get pages state no Entra role and no P1/P2
// prerequisite for these reads, and personal Microsoft accounts are not
// supported. Reads run in the partner tenant: delegatedAdminCustomer objects
// are created by the system when a relationship exists and deleted when none
// remain, so a non-partner tenant lists zero customers, which is an answer
// rather than an error. Both lists document $select, $filter, $top,
// $orderby, $count and $skipToken; $top supports up to 300 objects. --filter
// passes through as plain $filter with no $count or ConsistencyLevel
// contract. The $count scalars, the tenantRelationship container root, the
// serviceManagementDetails navigation and the relationship accessAssignments,
// operations and requests navigations stay scheduled for later EXT-04
// subfamilies, as do the multi-tenant-organization and tenant-lookup
// functions. Beta stays out. No delegated-admin mutation exists in this
// slice: relationship creation, approval and termination are writes.

// Every customer property this slice may request or display, matching the
// reviewed delegatedAdminCustomer resource. Anything else fails before
// credentials.
export const KNOWN_DELEGATED_ADMIN_CUSTOMER_FIELDS: readonly string[] = [
  "displayName",
  "id",
  "tenantId",
];
// Every relationship property this slice may request or display, matching
// the reviewed delegatedAdminRelationship resource. The list/get responses
// may carry a resellerDelegatedAdminRelationship subtype with unreviewed
// extras (isPartnerConsentPending, indirectProviderTenantId); projection
// drops them, so subtype rows render on the base set only.
export const KNOWN_DELEGATED_ADMIN_RELATIONSHIP_FIELDS: readonly string[] = [
  "accessDetails",
  "activatedDateTime",
  "autoExtendDuration",
  "createdDateTime",
  "customer",
  "displayName",
  "duration",
  "endDateTime",
  "id",
  "lastModifiedDateTime",
  "status",
];
const KNOWN_CUSTOMERS = new Set(KNOWN_DELEGATED_ADMIN_CUSTOMER_FIELDS);
const KNOWN_RELATIONSHIPS = new Set(KNOWN_DELEGATED_ADMIN_RELATIONSHIP_FIELDS);

// Compact customer rows: the customer identifier, the tenant name and the
// tenant id that names the same customer.
const DEFAULT_CUSTOMER_LIST_SELECT = ["id", "displayName", "tenantId"];
// Show rows: the full reviewed customer set.
const DEFAULT_CUSTOMER_SHOW_SELECT = [...KNOWN_DELEGATED_ADMIN_CUSTOMER_FIELDS];
// Compact relationship rows: the relationship identifier and name, the
// lifecycle status, the customer participant and the expiry that decides
// renewal attention.
const DEFAULT_RELATIONSHIP_LIST_SELECT = ["id", "displayName", "status", "customer", "endDateTime"];
// Show rows: the full reviewed relationship set.
const DEFAULT_RELATIONSHIP_SHOW_SELECT = [...KNOWN_DELEGATED_ADMIN_RELATIONSHIP_FIELDS];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_DELEGATED_ADMIN_SCOPES = ["https://graph.microsoft.com/DelegatedAdminRelationship.Read.All"];
const TRUNCATE_AT = 500;

export type DelegatedAdminFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown delegated-admin property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: DelegatedAdminFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
  flags: DelegatedAdminFlags,
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

function fullHint(command: string, flags: DelegatedAdminFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the scope and partner prerequisite that actually unlock it,
// because a 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const DELEGATED_ADMIN_DENIAL_HINTS = [
  "Delegated-admin reads need DelegatedAdminRelationship.Read.All, passed as --scopes for delegated access or admin-consented for application access",
  "Delegated-admin reads run in the partner tenant; a non-partner tenant reads empty, which is an answer rather than an error",
  "Personal Microsoft accounts are not supported for delegated-admin reads",
  "No P1/P2 prerequisite is stated for delegated-admin reads; never diagnose licence solely from HTTP 403",
];

const PARTNER_NOTE = "Delegated-admin reads run in the partner tenant; a non-partner tenant lists zero customers, which is an answer rather than an error";

interface Resource {
  noun: string;
  command: string;
  known: Set<string>;
  knownList: readonly string[];
  listSelect: string[];
  showSelect: string[];
  idFlag: string;
  idDescription: string;
}

const CUSTOMER: Resource = {
  noun: "customer",
  command: "entra delegated-admin-customer list",
  known: KNOWN_CUSTOMERS,
  knownList: KNOWN_DELEGATED_ADMIN_CUSTOMER_FIELDS,
  listSelect: DEFAULT_CUSTOMER_LIST_SELECT,
  showSelect: DEFAULT_CUSTOMER_SHOW_SELECT,
  idFlag: "id",
  idDescription: "Delegated-admin customer identifier",
};

const RELATIONSHIP: Resource = {
  noun: "relationship",
  command: "entra delegated-admin-relationship list",
  known: KNOWN_RELATIONSHIPS,
  knownList: KNOWN_DELEGATED_ADMIN_RELATIONSHIP_FIELDS,
  listSelect: DEFAULT_RELATIONSHIP_LIST_SELECT,
  showSelect: DEFAULT_RELATIONSHIP_SHOW_SELECT,
  idFlag: "id",
  idDescription: "Delegated-admin relationship identifier",
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
// --select or --filter is omitted. Both lists pass --filter through as plain
// $filter with no $count or ConsistencyLevel contract.
function collectionCommon(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  resource: Resource,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    resource.known, resource.knownList,
    saved?.["$select"] === undefined ? resource.listSelect : fieldList(saved["$select"], resource.known, resource.knownList, "select", help), help);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_DELEGATED_ADMIN_SCOPES, profile, help), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: DelegatedAdminFlags,
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

function resourceId(flags: DelegatedAdminFlags, resource: Resource, help: string): string {
  const id = String(flags[resource.idFlag]);
  if (!id.trim()) throw new AxiError(`--${resource.idFlag} needs the delegated-admin ${resource.noun} identifier`, "VALIDATION_ERROR", [help]);
  return id;
}

async function listResource(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  resource: Resource,
  collectionKey: string,
  showCommand: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, resource, operation, help, profile);
  const result = await withGuidance(DELEGATED_ADMIN_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: DelegatedAdminFlags = { ...flags, select: result.query.$select ?? resource.listSelect.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi ${showCommand} --id <${resource.noun}-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(resource.command, effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      [collectionKey]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, PARTNER_NOTE],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return {
      [collectionKey]: rows,
      count,
      help: [`0 ${resource.noun}s matched; the absence of results is the answer, not an error`, PARTNER_NOTE],
    };
  }
  return { [collectionKey]: rows, count, help: [...truncationHints, showHint, PARTNER_NOTE] };
}

async function showResource(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  resource: Resource,
  rowKey: string,
  showCommand: string,
  paramName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, resource.known, resource.knownList, resource.showSelect, help);
  const scopes = scopesFor(flags, DEFAULT_DELEGATED_ADMIN_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(DELEGATED_ADMIN_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { [paramName]: resourceId(flags, resource, help) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError(`Graph returned a malformed delegated-admin ${resource.noun} body`, "GRAPH_ERROR", [
      `Single-${resource.noun} reads carry one ${resource.noun} object; treat anything else as unknown, not empty`,
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { [rowKey]: row, help: [fullHint(showCommand, flags, profileName)] };
  return { [rowKey]: row };
}

export async function listDelegatedAdminCustomers(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listResource(session, flags, profile, operation, help, profileName,
    CUSTOMER, "delegatedAdminCustomers", "entra delegated-admin-customer show");
}

export async function showDelegatedAdminCustomer(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showResource(session, flags, profile, operation, help, profileName,
    CUSTOMER, "delegatedAdminCustomer", "entra delegated-admin-customer show", "delegatedAdminCustomer-id");
}

export async function listDelegatedAdminRelationships(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listResource(session, flags, profile, operation, help, profileName,
    RELATIONSHIP, "delegatedAdminRelationships", "entra delegated-admin-relationship show");
}

export async function showDelegatedAdminRelationship(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showResource(session, flags, profile, operation, help, profileName,
    RELATIONSHIP, "delegatedAdminRelationship", "entra delegated-admin-relationship show", "delegatedAdminRelationship-id");
}
