import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-04 delegated-admin subfamily: the read mapping behind
// `mg-axi entra delegated-admin-customer list/show`,
// `mg-axi entra delegated-admin-relationship list/show` and the navigation
// reads `mg-axi entra delegated-admin-relationship
// list-access-assignments/show-access-assignment/list-operations/show-operation/list-requests/show-request`
// and `mg-axi entra delegated-admin-customer
// list-service-management-details/show-service-management-detail`. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 delegatedAdminCustomer list/get, the
// delegatedAdminRelationship list/get, the access-assignment, operation and
// request list/get pages, the serviceManagementDetails list page, the
// delegatedAdminCustomer and delegatedAdminRelationship resource pages and
// the delegatedAdminAccessAssignment, delegatedAdminRelationshipOperation,
// delegatedAdminRelationshipRequest and delegatedAdminServiceManagementDetail
// resource pages on 2026-10-05. The serviceManagementDetail single has no
// REST page; its route exists alongside the documented list (the PowerShell
// Get cmdlet takes both the customer and the detail identifiers), so the
// single runs queryless like the list. Every read takes D/A
// DelegatedAdminRelationship.Read.All as least privilege; the scope is
// already in the shared READ_SCOPES allowlist, so no new scope and no
// deferral is needed here. Delegated callers pass it as --scopes;
// application callers need it admin-consented on the configured .default
// audience. None of the reviewed pages states an Entra role or a P1/P2
// prerequisite for these reads, and personal Microsoft accounts are not
// supported. Reads run in the partner tenant: delegatedAdminCustomer objects
// are created by the system when a relationship exists and deleted when none
// remain, so a non-partner tenant lists zero customers, which is an answer
// rather than an error. The access-assignment, operation and request lists
// document $select, $filter, $top and more; $top supports up to 300 objects.
// --filter passes through as plain $filter with no $count or ConsistencyLevel
// contract. serviceManagementDetails documents no query parameters, so both
// service-management-detail reads run queryless and project locally. The
// $count scalars and the tenantRelationship container root stay scheduled
// for later EXT-04 subfamilies, as do the multi-tenant-organization reads
// (shipped separately) and the tenant-lookup functions (shipped as
// tenant-information show under mg-ext-04e). Beta stays out. No delegated-admin mutation exists in this slice:
// relationship creation, approval and termination are writes.

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
// Every access-assignment property this slice may request or display,
// matching the reviewed delegatedAdminAccessAssignment resource. Anything
// else fails before credentials.
export const KNOWN_DELEGATED_ADMIN_ACCESS_ASSIGNMENT_FIELDS: readonly string[] = [
  "accessContainer",
  "accessDetails",
  "createdDateTime",
  "id",
  "lastModifiedDateTime",
  "status",
];
// Every relationship-operation property this slice may request or display,
// matching the reviewed delegatedAdminRelationshipOperation resource. The
// data payload is a JSON-encoded string and truncates like any long text.
export const KNOWN_DELEGATED_ADMIN_OPERATION_FIELDS: readonly string[] = [
  "createdDateTime",
  "data",
  "id",
  "lastModifiedDateTime",
  "operationType",
  "status",
];
// Every relationship-request property this slice may request or display,
// matching the reviewed delegatedAdminRelationshipRequest resource.
export const KNOWN_DELEGATED_ADMIN_REQUEST_FIELDS: readonly string[] = [
  "action",
  "createdDateTime",
  "id",
  "lastModifiedDateTime",
  "status",
];
// Every service-management-detail property this slice may display, matching
// the reviewed delegatedAdminServiceManagementDetail resource. Graph
// documents no query parameters here, so rows always arrive whole and
// --fields projects them locally.
export const KNOWN_DELEGATED_ADMIN_SERVICE_MANAGEMENT_DETAIL_FIELDS: readonly string[] = [
  "id",
  "serviceManagementUrl",
  "serviceName",
];
const KNOWN_ACCESS_ASSIGNMENTS = new Set(KNOWN_DELEGATED_ADMIN_ACCESS_ASSIGNMENT_FIELDS);
const KNOWN_OPERATIONS = new Set(KNOWN_DELEGATED_ADMIN_OPERATION_FIELDS);
const KNOWN_REQUESTS = new Set(KNOWN_DELEGATED_ADMIN_REQUEST_FIELDS);
const KNOWN_SERVICE_MANAGEMENT_DETAILS = new Set(KNOWN_DELEGATED_ADMIN_SERVICE_MANAGEMENT_DETAIL_FIELDS);

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

interface Resource {
  noun: string;
  command: string;
  known: Set<string>;
  knownList: readonly string[];
  listSelect: string[];
  showSelect: string[];
  idFlag: string;
  idDescription: string;
  showPlaceholder: string;
  // Nested navigations bind the parent collection through its own flag;
  // top-level customers and relationships leave this undefined.
  parent?: { flag: string; placeholder: string; label: string };
  // serviceManagementDetails documents no query parameters, so both reads
  // run queryless and project rows locally; every other read requests
  // --select and passes --filter through as plain $filter.
  queryless?: boolean;
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
  showPlaceholder: "delegatedAdminCustomer-id",
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
  showPlaceholder: "delegatedAdminRelationship-id",
};

const RELATIONSHIP_PARENT = { flag: "id", placeholder: "delegatedAdminRelationship-id", label: "relationship" };
const CUSTOMER_PARENT = { flag: "id", placeholder: "delegatedAdminCustomer-id", label: "customer" };

const ACCESS_ASSIGNMENT: Resource = {
  noun: "access assignment",
  command: "entra delegated-admin-relationship list-access-assignments",
  known: KNOWN_ACCESS_ASSIGNMENTS,
  knownList: KNOWN_DELEGATED_ADMIN_ACCESS_ASSIGNMENT_FIELDS,
  listSelect: ["id", "status", "accessContainer", "accessDetails"],
  showSelect: [...KNOWN_DELEGATED_ADMIN_ACCESS_ASSIGNMENT_FIELDS],
  idFlag: "assignment-id",
  idDescription: "Delegated-admin access-assignment identifier",
  showPlaceholder: "delegatedAdminAccessAssignment-id",
  parent: RELATIONSHIP_PARENT,
};

const OPERATION: Resource = {
  noun: "relationship operation",
  command: "entra delegated-admin-relationship list-operations",
  known: KNOWN_OPERATIONS,
  knownList: KNOWN_DELEGATED_ADMIN_OPERATION_FIELDS,
  listSelect: ["id", "operationType", "status", "lastModifiedDateTime"],
  showSelect: [...KNOWN_DELEGATED_ADMIN_OPERATION_FIELDS],
  idFlag: "operation-id",
  idDescription: "Delegated-admin relationship-operation identifier",
  showPlaceholder: "delegatedAdminRelationshipOperation-id",
  parent: RELATIONSHIP_PARENT,
};

const REQUEST: Resource = {
  noun: "relationship request",
  command: "entra delegated-admin-relationship list-requests",
  known: KNOWN_REQUESTS,
  knownList: KNOWN_DELEGATED_ADMIN_REQUEST_FIELDS,
  listSelect: ["id", "action", "status", "lastModifiedDateTime"],
  showSelect: [...KNOWN_DELEGATED_ADMIN_REQUEST_FIELDS],
  idFlag: "request-id",
  idDescription: "Delegated-admin relationship-request identifier",
  showPlaceholder: "delegatedAdminRelationshipRequest-id",
  parent: RELATIONSHIP_PARENT,
};

const SERVICE_MANAGEMENT_DETAIL: Resource = {
  noun: "service-management detail",
  command: "entra delegated-admin-customer list-service-management-details",
  known: KNOWN_SERVICE_MANAGEMENT_DETAILS,
  knownList: KNOWN_DELEGATED_ADMIN_SERVICE_MANAGEMENT_DETAIL_FIELDS,
  listSelect: [...KNOWN_DELEGATED_ADMIN_SERVICE_MANAGEMENT_DETAIL_FIELDS],
  showSelect: [...KNOWN_DELEGATED_ADMIN_SERVICE_MANAGEMENT_DETAIL_FIELDS],
  idFlag: "detail-id",
  idDescription: "Delegated-admin service-management-detail identifier",
  showPlaceholder: "delegatedAdminServiceManagementDetail-id",
  parent: CUSTOMER_PARENT,
  queryless: true,
};

function partnerNote(resource: Resource): string {
  return `Delegated-admin reads run in the partner tenant; a non-partner tenant lists zero ${resource.noun}s, which is an answer rather than an error`;
}

function parentBinding(flags: DelegatedAdminFlags, resource: Resource, help: string): Record<string, string> {
  if (!resource.parent) return {};
  const id = String(flags[resource.parent.flag]);
  if (!id.trim()) throw new AxiError(`--${resource.parent.flag} needs the delegated-admin ${resource.parent.label} identifier`, "VALIDATION_ERROR", [help]);
  return { [resource.parent.placeholder]: id };
}

// Queryless navigations take no --select or --filter: Graph documents no
// query parameters for them, so the catalogue declares no such flags and
// this refusal is the backstop before credentials.
function refuseQuery(flags: DelegatedAdminFlags, resource: Resource, command: string, help: string): void {
  if (resource.queryless === true && (flags.select !== undefined || flags.filter !== undefined)) {
    throw new AxiError(`${command} takes no --select or --filter; Graph documents no query parameters for this navigation`, "VALIDATION_ERROR", [help]);
  }
}

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
  filter: string | undefined;
  params: Record<string, string>;
  bare: boolean;
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
  const params = parentBinding(flags, resource, help);
  if (resource.queryless === true) {
    refuseQuery(flags, resource, resource.command, help);
    // Rows arrive whole; --fields projects them locally from the full set.
    const fields = flags.fields === undefined ? [...resource.listSelect] : fieldList(flags.fields, resource.known, resource.knownList, "fields", help);
    return { cursor, select: [...resource.listSelect], fields, scopes: scopesFor(flags, DEFAULT_DELEGATED_ADMIN_SCOPES, profile, help), full: flags.full === true, filter: undefined, params, bare: true };
  }
  const { select, fields } = selectedFields(flags,
    resource.known, resource.knownList,
    saved?.["$select"] === undefined ? resource.listSelect : fieldList(saved["$select"], resource.known, resource.knownList, "select", help), help);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_DELEGATED_ADMIN_SCOPES, profile, help), full: flags.full === true, filter, params, bare: false };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: DelegatedAdminFlags,
  help: string,
): CollectArgs {
  // Queryless navigations send no query parameters at all: Graph documents
  // none for them, so rows arrive whole and project locally.
  const query: Record<string, string> = common.bare ? {} : { $select: common.select.join(",") };
  if (!common.bare && common.filter !== undefined) query.$filter = common.filter;
  const args: CollectArgs = { profile, operation, params: common.params, query, scopes: common.scopes };
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
  // Queryless hints never echo --select/--filter: those flags are refused.
  const effectiveFlags: DelegatedAdminFlags = { ...flags };
  if (!common.bare) {
    effectiveFlags.select = result.query.$select ?? resource.listSelect.join(",");
    if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  }
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showArgs = resource.parent === undefined
    ? `--id <${resource.noun}-id>`
    : `--${resource.parent.flag} <${resource.parent.label}-id> --${resource.idFlag} <${resource.idFlag}>`;
  const showHint = `mg-axi ${showCommand} ${showArgs} ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(resource.command, effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      [collectionKey]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, partnerNote(resource)],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return {
      [collectionKey]: rows,
      count,
      help: [`0 ${resource.noun}s matched; the absence of results is the answer, not an error`, partnerNote(resource)],
    };
  }
  return { [collectionKey]: rows, count, help: [...truncationHints, showHint, partnerNote(resource)] };
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
): Promise<Record<string, unknown>> {
  const queryless = resource.queryless === true;
  if (queryless) refuseQuery(flags, resource, showCommand, help);
  const { select, fields } = queryless
    ? { select: [...resource.showSelect], fields: flags.fields === undefined ? [...resource.showSelect] : fieldList(flags.fields, resource.known, resource.knownList, "fields", help) }
    : selectedFields(flags, resource.known, resource.knownList, resource.showSelect, help);
  const scopes = scopesFor(flags, DEFAULT_DELEGATED_ADMIN_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(DELEGATED_ADMIN_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { ...parentBinding(flags, resource, help), [resource.showPlaceholder]: resourceId(flags, resource, help) },
    ...(queryless ? {} : { query: { $select: select.join(",") } }),
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
    CUSTOMER, "delegatedAdminCustomer", "entra delegated-admin-customer show");
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
    RELATIONSHIP, "delegatedAdminRelationship", "entra delegated-admin-relationship show");
}

export async function listDelegatedAdminAccessAssignments(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listResource(session, flags, profile, operation, help, profileName,
    ACCESS_ASSIGNMENT, "delegatedAdminAccessAssignments", "entra delegated-admin-relationship show-access-assignment");
}

export async function showDelegatedAdminAccessAssignment(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showResource(session, flags, profile, operation, help, profileName,
    ACCESS_ASSIGNMENT, "delegatedAdminAccessAssignment", "entra delegated-admin-relationship show-access-assignment");
}

export async function listDelegatedAdminOperations(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listResource(session, flags, profile, operation, help, profileName,
    OPERATION, "delegatedAdminRelationshipOperations", "entra delegated-admin-relationship show-operation");
}

export async function showDelegatedAdminOperation(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showResource(session, flags, profile, operation, help, profileName,
    OPERATION, "delegatedAdminRelationshipOperation", "entra delegated-admin-relationship show-operation");
}

export async function listDelegatedAdminRequests(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listResource(session, flags, profile, operation, help, profileName,
    REQUEST, "delegatedAdminRelationshipRequests", "entra delegated-admin-relationship show-request");
}

export async function showDelegatedAdminRequest(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showResource(session, flags, profile, operation, help, profileName,
    REQUEST, "delegatedAdminRelationshipRequest", "entra delegated-admin-relationship show-request");
}

export async function listDelegatedAdminServiceManagementDetails(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listResource(session, flags, profile, operation, help, profileName,
    SERVICE_MANAGEMENT_DETAIL, "delegatedAdminServiceManagementDetails", "entra delegated-admin-customer show-service-management-detail");
}

export async function showDelegatedAdminServiceManagementDetail(
  session: GraphSession,
  flags: DelegatedAdminFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showResource(session, flags, profile, operation, help, profileName,
    SERVICE_MANAGEMENT_DETAIL, "delegatedAdminServiceManagementDetail", "entra delegated-admin-customer show-service-management-detail");
}
