import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-02 entitlement-management first part: the read-only catalog and
// access-package mapping. See README.md for the supported commands and
// usage. Operation construction stays beside its command; the shared
// session owns URLs, credentials, paging, retries and error translation,
// and the SDK owns TOON rendering. This module only maps flags to session
// calls and projects rows for compact output.
//
// Reviewed against the v1.0 entitlementmanagement-list-catalogs,
// accesspackagecatalog-get, entitlementmanagement-list-accesspackages,
// accesspackage-get and entitlementmanagement-list-assignmentpolicies
// operation documentation, the accessPackageAssignmentPolicy resource
// contract, and the pinned v1.0 metadata property sets for
// accessPackageCatalog, accessPackage, accessPackageAssignmentPolicy and
// accessPackageResourceRoleScope, on 2026-10-05. All twelve reads take D/A
// EntitlementManagement.Read.All as the least privileged choice; delegated
// personal Microsoft accounts are not supported. Delegated callers
// additionally need a supported Entra role with catalog visibility (Global
// Reader and Identity Governance Administrator are among the supported
// roles). The four lists document the general OData query parameters, so
// --filter passes through as plain $filter with no $count or
// ConsistencyLevel contract ($search/$expand/$orderby stay unreviewed and
// $top/$skip stay server-side paging concerns); singles document $select
// only and $count routes take no query at all. Navigation properties
// (a catalog's accessPackages/resources/scopes/roles, a package's catalog
// and incompatible sets, a policy's questions/extensions, a role scope's
// role/scope links) need $expand and stay out: they belong to later slices,
// never to these reads. Assignment, request, approval and subject reads
// carry personal data and belong to a later part, never to these reads.
// Entitlement management needs P2 or ID Governance depending on capability,
// never one uniform licence.

// Catalogs expose the package containers; serviceDefault catalogs are
// platform-owned while userManaged catalogs are tenant-created. Only scalar
// properties carry a reviewed $select contract; accessPackages, resources,
// resourceScopes, resourceRoles and customWorkflowExtensions are navigation
// properties and stay out.
export const KNOWN_CATALOG_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "description",
  "catalogType",
  "state",
  "isExternallyVisible",
  "createdDateTime",
  "modifiedDateTime",
];
// Access packages expose the assignable bundles; isHidden marks packages
// withheld from request listings. The catalog link and incompatible sets
// are navigation properties and stay out.
export const KNOWN_ACCESS_PACKAGE_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "description",
  "isHidden",
  "createdDateTime",
  "modifiedDateTime",
];
// Assignment policies expose who may request, approval and review settings
// for one access package. The nested settings objects project as values;
// questions, custom-extension stages, accessPackage and catalog links are
// navigation properties and stay out.
export const KNOWN_ASSIGNMENT_POLICY_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "description",
  "allowedTargetScope",
  "automaticRequestSettings",
  "createdDateTime",
  "expiration",
  "modifiedDateTime",
  "notificationSettings",
  "requestApprovalSettings",
  "requestorSettings",
  "reviewSettings",
  "specificAllowedTargets",
];
// Resource-role scopes expose the role-plus-scope pairs one access package
// grants. Only the pairing identity carries a reviewed $select contract;
// the role and scope links are navigation properties and stay out.
export const KNOWN_RESOURCE_ROLE_SCOPE_FIELDS: readonly string[] = [
  "id",
  "createdDateTime",
];
const CATALOG_KNOWN = new Set(KNOWN_CATALOG_FIELDS);
const ACCESS_PACKAGE_KNOWN = new Set(KNOWN_ACCESS_PACKAGE_FIELDS);
const ASSIGNMENT_POLICY_KNOWN = new Set(KNOWN_ASSIGNMENT_POLICY_FIELDS);
const RESOURCE_ROLE_SCOPE_KNOWN = new Set(KNOWN_RESOURCE_ROLE_SCOPE_FIELDS);

// Compact rows: identifiers plus container state, package visibility,
// policy scope or pairing identity.
const DEFAULT_CATALOG_LIST_SELECT = ["id", "displayName", "state", "catalogType"];
const DEFAULT_CATALOG_SHOW_SELECT = [...KNOWN_CATALOG_FIELDS];
const DEFAULT_ACCESS_PACKAGE_LIST_SELECT = ["id", "displayName", "isHidden"];
const DEFAULT_ACCESS_PACKAGE_SHOW_SELECT = [...KNOWN_ACCESS_PACKAGE_FIELDS];
const DEFAULT_ASSIGNMENT_POLICY_LIST_SELECT = ["id", "displayName", "allowedTargetScope"];
const DEFAULT_ASSIGNMENT_POLICY_SHOW_SELECT = [...KNOWN_ASSIGNMENT_POLICY_FIELDS];
const DEFAULT_RESOURCE_ROLE_SCOPE_LIST_SELECT = ["id", "createdDateTime"];
const DEFAULT_RESOURCE_ROLE_SCOPE_SHOW_SELECT = [...KNOWN_RESOURCE_ROLE_SCOPE_FIELDS];
export const DEFAULT_ENTITLEMENT_MANAGEMENT_SCOPES = ["https://graph.microsoft.com/EntitlementManagement.Read.All"];
const TRUNCATE_AT = 500;

export type EntitlementManagementFlags = Record<string, string | boolean>;

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

function scopesFor(flags: EntitlementManagementFlags, profile: AnyProfile, defaults: readonly string[], help: string): string[] | undefined {
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
  if (full) return { value, truncated: false };
  if (typeof value === "string" && value.length > TRUNCATE_AT) {
    return { value: `${value.slice(0, TRUNCATE_AT)}... (truncated, ${value.length} chars total)`, truncated: true };
  }
  if (Array.isArray(value)) {
    const results = value.map(item => truncateValue(item, full));
    return { value: results.map(result => result.value), truncated: results.some(result => result.truncated) };
  }
  if (value !== null && typeof value === "object") {
    let truncated = false;
    const entries = Object.entries(value).map(([key, item]) => {
      const result = truncateValue(item, full);
      truncated = truncated || result.truncated;
      return [key, result.value];
    });
    return { value: Object.fromEntries(entries), truncated };
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
  flags: EntitlementManagementFlags,
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

function fullHint(command: string, flags: EntitlementManagementFlags, profileName: string): string {
  const args = Object.entries({ ...flags, ...(flags.cursor === undefined ? {} : { cursor: "-" }), profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause;
// entitlement management adds the scope, roles and licensing that actually
// unlock these reads, because a 403 alone never says which prerequisite is
// missing.
function withGuidance<T>(run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...ENTITLEMENT_MANAGEMENT_DENIAL_HINTS, ...error.suggestions]);
    }
    throw error;
  });
}

const ENTITLEMENT_MANAGEMENT_DENIAL_HINTS = [
  "Entitlement-management reads need EntitlementManagement.Read.All plus a supported Entra role with catalog visibility for delegated access (Global Reader and Identity Governance Administrator are among the supported roles), or admin-consented EntitlementManagement.Read.All for application access; delegated personal Microsoft accounts are not supported",
  "Entitlement management needs P2 or ID Governance depending on capability, not one uniform licence",
];

interface CollectionShape {
  command: string;
  key: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  emptyNote: string;
  standing: (profileName: string) => string[];
}

const CATALOG_LIST: CollectionShape = {
  command: "entra entitlement catalog list",
  key: "catalogs",
  known: CATALOG_KNOWN,
  knownList: KNOWN_CATALOG_FIELDS,
  defaultSelect: DEFAULT_CATALOG_LIST_SELECT,
  emptyNote: "0 entitlement catalogs matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Catalogs are package containers: a catalog never carries its access packages; list those on the package reads",
    `Show one catalog: mg-axi entra entitlement catalog show --id <catalog-id> ${profileHint(profileName)}`,
  ],
};

const ACCESS_PACKAGE_LIST: CollectionShape = {
  command: "entra entitlement access-package list",
  key: "accessPackages",
  known: ACCESS_PACKAGE_KNOWN,
  knownList: KNOWN_ACCESS_PACKAGE_FIELDS,
  defaultSelect: DEFAULT_ACCESS_PACKAGE_LIST_SELECT,
  emptyNote: "0 access packages matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Access packages are the assignable bundles; assignments and requests belong to a later part, never to these reads",
    `Show one access package: mg-axi entra entitlement access-package show --id <access-package-id> ${profileHint(profileName)}`,
  ],
};

const ASSIGNMENT_POLICY_LIST: CollectionShape = {
  command: "entra entitlement assignment-policy list",
  key: "assignmentPolicies",
  known: ASSIGNMENT_POLICY_KNOWN,
  knownList: KNOWN_ASSIGNMENT_POLICY_FIELDS,
  defaultSelect: DEFAULT_ASSIGNMENT_POLICY_LIST_SELECT,
  emptyNote: "0 assignment policies matched; an access package without policies has none, which is the answer, not an error",
  standing: profileName => [
    "Assignment policies name who may request and how approval and review run; questions and custom-extension stages belong to a later slice",
    `Show one policy: mg-axi entra entitlement assignment-policy show --access-package <access-package-id> --id <policy-id> ${profileHint(profileName)}`,
  ],
};

const RESOURCE_ROLE_SCOPE_LIST: CollectionShape = {
  command: "entra entitlement resource-role-scope list",
  key: "resourceRoleScopes",
  known: RESOURCE_ROLE_SCOPE_KNOWN,
  knownList: KNOWN_RESOURCE_ROLE_SCOPE_FIELDS,
  defaultSelect: DEFAULT_RESOURCE_ROLE_SCOPE_LIST_SELECT,
  emptyNote: "0 resource-role scopes matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Resource-role scopes are the role-plus-scope pairs a package grants; the linked role and scope detail needs $expand and belongs to a later slice",
    `Show one resource-role scope: mg-axi entra entitlement resource-role-scope show --access-package <access-package-id> --id <scope-id> ${profileHint(profileName)}`,
  ],
};

function accessPackageId(flags: EntitlementManagementFlags, help: string): string {
  const value = String(flags["access-package"] ?? "");
  if (!value.trim()) throw new AxiError("--access-package needs the parent access-package ID", "VALIDATION_ERROR", [help]);
  return value;
}

async function listCollection(
  shape: CollectionShape,
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  params: Record<string, string>,
  resumeHint: string,
): Promise<Record<string, unknown>> {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const savedSelect = saved?.$select;
  const { select, fields } = selectedFields(flags,
    savedSelect === undefined ? shape.defaultSelect : fieldList(savedSelect, "select", shape.known, shape.knownList, help),
    shape.known, shape.knownList, help);
  const savedFilter = saved?.$filter;
  const filter = flags.filter === undefined ? savedFilter : String(flags.filter);
  const scopes = scopesFor(flags, profile, DEFAULT_ENTITLEMENT_MANAGEMENT_SCOPES, help);
  const full = flags.full === true;
  // These lists document the general OData query parameters, so --filter
  // passes through as plain $filter with no $count or ConsistencyLevel;
  // $search/$expand/$orderby stay unreviewed and $top/$skip stay
  // server-side paging concerns.
  const query: Record<string, string> = { $select: select.join(",") };
  if (filter !== undefined) query.$filter = filter;
  const args: CollectArgs = { profile, operation, params, query, scopes };
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(() => session.collect(args));
  const effectiveFlags: EntitlementManagementFlags = { ...flags, select: result.query.$select ?? shape.defaultSelect.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const truncationHints = truncated ? [fullHint(shape.command, effectiveFlags, profileName)] : [];
  if (truncated && cursor !== undefined) truncationHints.push("Supply the original input cursor on stdin to replay this result with --full");
  const standing = [...shape.standing(profileName)];
  if (!result.complete) {
    return {
      [shape.key]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint, ...standing],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return { [shape.key]: rows, count, help: [shape.emptyNote, ...standing] };
  }
  return { [shape.key]: rows, count, help: [...truncationHints, ...standing] };
}

async function showOne(
  key: string,
  command: string,
  known: Set<string>,
  knownList: readonly string[],
  defaultSelect: string[],
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  params: Record<string, string>,
  malformed: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, defaultSelect, known, knownList, help);
  const scopes = scopesFor(flags, profile, DEFAULT_ENTITLEMENT_MANAGEMENT_SCOPES, help);
  const full = flags.full === true;
  const raw = await withGuidance(() => session.execute({
    profile,
    operation,
    params,
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError(malformed, "GRAPH_ERROR", [
      "Single-object reads carry one object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { [key]: row, help: [fullHint(command, flags, profileName)] };
  return { [key]: row };
}

async function countCollection(
  key: string,
  emptyNote: string,
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, profile, DEFAULT_ENTITLEMENT_MANAGEMENT_SCOPES, help);
  // The $count route returns a text/plain integer scalar rather than a
  // JSON collection, so the leaf reads it through session.execute with
  // scalar mode and accepts only a non-negative integer. The $count route
  // carries no operation-level documentation page and takes no
  // --filter/--select/--limit/--cursor.
  const raw = await withGuidance(() => session.execute({ profile, operation, scopes, scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError(`Graph returned a malformed ${key} count body`, "GRAPH_ERROR", [
      "Counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  const count = { returned: raw, complete: true };
  return raw === 0 ? { count, help: [emptyNote] } : { count };
}

export async function listCatalogs(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(CATALOG_LIST, session, flags, profile, operation, help, profileName, {},
    `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showCatalog(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("catalog", "entra entitlement catalog show",
    CATALOG_KNOWN, KNOWN_CATALOG_FIELDS, DEFAULT_CATALOG_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessPackageCatalog-id": String(flags.id) },
    "Graph returned a malformed entitlement catalog body");
}

export async function countCatalogs(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return countCollection("catalog", CATALOG_LIST.emptyNote, session, flags, profile, operation, help, profileName);
}

export async function listAccessPackages(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ACCESS_PACKAGE_LIST, session, flags, profile, operation, help, profileName, {},
    `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showAccessPackage(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("accessPackage", "entra entitlement access-package show",
    ACCESS_PACKAGE_KNOWN, KNOWN_ACCESS_PACKAGE_FIELDS, DEFAULT_ACCESS_PACKAGE_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessPackage-id": String(flags.id) },
    "Graph returned a malformed access package body");
}

export async function countAccessPackages(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return countCollection("access package", ACCESS_PACKAGE_LIST.emptyNote, session, flags, profile, operation, help, profileName);
}

export async function listAssignmentPolicies(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const accessPackage = accessPackageId(flags, help);
  return listCollection(ASSIGNMENT_POLICY_LIST, session, flags, profile, operation, help, profileName,
    { "accessPackage-id": accessPackage },
    `Resume losslessly with the same --access-package and flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showAssignmentPolicy(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const accessPackage = accessPackageId(flags, help);
  return showOne("assignmentPolicy", "entra entitlement assignment-policy show",
    ASSIGNMENT_POLICY_KNOWN, KNOWN_ASSIGNMENT_POLICY_FIELDS, DEFAULT_ASSIGNMENT_POLICY_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessPackage-id": accessPackage, "accessPackageAssignmentPolicy-id": String(flags.id) },
    "Graph returned a malformed assignment policy body");
}

export async function countAssignmentPolicies(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const accessPackage = accessPackageId(flags, help);
  const scopes = scopesFor(flags, profile, DEFAULT_ENTITLEMENT_MANAGEMENT_SCOPES, help);
  const raw = await withGuidance(() => session.execute({
    profile,
    operation,
    params: { "accessPackage-id": accessPackage },
    scopes,
    scalar: true,
  }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed assignment policy count body", "GRAPH_ERROR", [
      "Counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  const count = { returned: raw, complete: true };
  return raw === 0 ? { count, help: [ASSIGNMENT_POLICY_LIST.emptyNote] } : { count };
}

export async function listResourceRoleScopes(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const accessPackage = accessPackageId(flags, help);
  return listCollection(RESOURCE_ROLE_SCOPE_LIST, session, flags, profile, operation, help, profileName,
    { "accessPackage-id": accessPackage },
    `Resume losslessly with the same --access-package and flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showResourceRoleScope(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const accessPackage = accessPackageId(flags, help);
  return showOne("resourceRoleScope", "entra entitlement resource-role-scope show",
    RESOURCE_ROLE_SCOPE_KNOWN, KNOWN_RESOURCE_ROLE_SCOPE_FIELDS, DEFAULT_RESOURCE_ROLE_SCOPE_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessPackage-id": accessPackage, "accessPackageResourceRoleScope-id": String(flags.id) },
    "Graph returned a malformed resource-role scope body");
}

export async function countResourceRoleScopes(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const accessPackage = accessPackageId(flags, help);
  const scopes = scopesFor(flags, profile, DEFAULT_ENTITLEMENT_MANAGEMENT_SCOPES, help);
  const raw = await withGuidance(() => session.execute({
    profile,
    operation,
    params: { "accessPackage-id": accessPackage },
    scopes,
    scalar: true,
  }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed resource-role scope count body", "GRAPH_ERROR", [
      "Counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  const count = { returned: raw, complete: true };
  return raw === 0 ? { count, help: [RESOURCE_ROLE_SCOPE_LIST.emptyNote] } : { count };
}
