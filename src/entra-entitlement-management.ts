import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-02 entitlement-management reads: the read-only catalog and
// access-package mapping plus the assignment and request reads. See
// README.md for the supported commands and usage. Operation construction
// stays beside its command; the shared session owns URLs, credentials,
// paging, retries and error translation, and the SDK owns TOON rendering.
// This module only maps flags to session calls and projects rows for
// compact output.
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
// never to these reads. Approval and subject reads carry personal data and
// belong to later parts, never to these reads. Entitlement management needs
// P2 or ID Governance depending on capability, never one uniform licence.
//
// The assignment and assignment-request reads below are the personal-data
// part deferred above. Reviewed against the v1.0
// entitlementmanagement-list-assignments, accesspackageassignment-get,
// entitlementmanagement-list-assignmentrequests and
// accesspackageassignmentrequest-get operation documentation and the
// accessPackageAssignment and accessPackageAssignmentRequest resource
// contracts, on 2026-10-05. All six reads take D/A
// EntitlementManagement.Read.All as the least privileged choice; delegated
// personal Microsoft accounts are not supported. Delegated callers
// additionally need Catalog reader (least privileged), Catalog creator or
// Access package manager, or a supported Entra role (Security Reader,
// Global Reader, Compliance Administrator, Security Administrator or
// Identity Governance Administrator). The two lists document $select,
// $filter and $expand, so --filter passes through as plain $filter with no
// $count or ConsistencyLevel contract; the singles document $select with
// $expand, and the $count routes take no query at all. The wire resources
// carry no target or access-package linkage as scalars, so each list and
// show sends one fixed documented $expand set in code (target and
// accessPackage for assignments; accessPackage and assignment for
// requests) and projects only flattened linkage identifiers from it:
// targetId, targetDisplayName and accessPackageId on assignments,
// accessPackageId and assignmentId on requests. No user-supplied $expand
// exists anywhere: --select names only reviewed scalar or flattened fields,
// and navigation names (target, accessPackage, assignment, requestor,
// assignmentPolicy) fail as unknown properties. The expanded payloads
// arrive whole and are projected locally, so defaults stay minimal:
// request justification text and answers ride only behind an explicit
// --select, and the requestor subject, the assignment policy link and the
// filterByCurrentUser/additionalAccess functions belong to later slices.
// No mutation (no request creation, cancellation, approval or reprocessing)
// and no beta.

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
// Assignments expose one subject's delivered package grant. The wire
// resource carries only scalars; target and access-package linkage arrives
// through the fixed $expand and is projected as the flattened targetId,
// targetDisplayName and accessPackageId fields below, never as nested
// subject or package objects.
export const KNOWN_ASSIGNMENT_FIELDS: readonly string[] = [
  "id",
  "state",
  "status",
  "expiredDateTime",
  "schedule",
  "customExtensionCalloutInstances",
  "targetId",
  "targetDisplayName",
  "accessPackageId",
];
// Assignment requests expose who asked for what and how processing ended.
// Justification text and answers are personal-data payloads: reviewed here
// so an explicit --select can fetch them, never in any default. The
// requestor subject, the resulting-assignment detail beyond its id and the
// access-package detail beyond its id belong to later slices.
export const KNOWN_ASSIGNMENT_REQUEST_FIELDS: readonly string[] = [
  "id",
  "requestType",
  "state",
  "status",
  "justification",
  "schedule",
  "createdDateTime",
  "completedDateTime",
  "answers",
  "customExtensionCalloutInstances",
  "accessPackageId",
  "assignmentId",
];
const CATALOG_KNOWN = new Set(KNOWN_CATALOG_FIELDS);
const ACCESS_PACKAGE_KNOWN = new Set(KNOWN_ACCESS_PACKAGE_FIELDS);
const ASSIGNMENT_POLICY_KNOWN = new Set(KNOWN_ASSIGNMENT_POLICY_FIELDS);
const RESOURCE_ROLE_SCOPE_KNOWN = new Set(KNOWN_RESOURCE_ROLE_SCOPE_FIELDS);
const ASSIGNMENT_KNOWN = new Set(KNOWN_ASSIGNMENT_FIELDS);
const ASSIGNMENT_REQUEST_KNOWN = new Set(KNOWN_ASSIGNMENT_REQUEST_FIELDS);
// Flattened linkage identifiers are always fetched through the fixed
// $expand, never through server $select, so they are stripped from the
// request selection but always count as fetched for --fields.
const ASSIGNMENT_SYNTHETIC: ReadonlySet<string> = new Set(["targetId", "targetDisplayName", "accessPackageId"]);
const ASSIGNMENT_REQUEST_SYNTHETIC: ReadonlySet<string> = new Set(["accessPackageId", "assignmentId"]);

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
// Assignment lists default to the linkage that makes a grant row actionable
// (which subject holds which package, and in what state) plus its dates;
// lifecycle detail stays behind an explicit --select.
const DEFAULT_ASSIGNMENT_LIST_SELECT = ["id", "state", "targetId", "targetDisplayName", "accessPackageId", "expiredDateTime", "schedule"];
const DEFAULT_ASSIGNMENT_SHOW_SELECT = [...KNOWN_ASSIGNMENT_FIELDS];
// Request lists default to the linkage plus the request-type discriminator:
// without userAdd/adminAdd/remove context a request row is uninterpretable.
// Justification text and answers stay behind an explicit --select, and the
// show default keeps that exclusion even though it is the full row view.
const DEFAULT_ASSIGNMENT_REQUEST_LIST_SELECT = ["id", "requestType", "state", "accessPackageId", "assignmentId", "createdDateTime", "completedDateTime", "schedule"];
const DEFAULT_ASSIGNMENT_REQUEST_SHOW_SELECT = ["id", "requestType", "state", "status", "schedule", "createdDateTime", "completedDateTime", "customExtensionCalloutInstances", "accessPackageId", "assignmentId"];
// One fixed documented $expand per family, set in code and never from user
// input. The list and get pages document exactly these expansions.
const ASSIGNMENT_EXPAND = "target,accessPackage";
const ASSIGNMENT_REQUEST_EXPAND = "accessPackage,assignment";
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
// The optional links view merges flattened $expand linkage beside the
// top-level scalars before that check, so absent linkage stays absent.
function project(
  row: unknown,
  fields: string[],
  full: boolean,
  links?: (source: Record<string, unknown>) => Record<string, unknown>,
): { row: Record<string, unknown>; truncated: boolean } {
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  const view = links === undefined ? source : { ...source, ...links(source) };
  const projected: Record<string, unknown> = {};
  let truncated = false;
  for (const field of fields) {
    if (!Object.hasOwn(view, field)) continue;
    const result = truncateValue(view[field], full);
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
  synthetic?: ReadonlySet<string>,
): { select: string[]; fields: string[]; fetch: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, "select", known, knownList, help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, "fields", known, knownList, help);
  // Flattened linkage rides the fixed $expand, so it always counts as
  // fetched even when the narrowed --select omits it.
  const missing = fields.find(field => !select.includes(field) && !(synthetic?.has(field) ?? false));
  if (missing) {
    throw new AxiError(`--fields ${missing} was not fetched; request it with --select`, "VALIDATION_ERROR", [help]);
  }
  // Synthetic names are never server $select properties; an all-linkage
  // selection omits $select and the server returns its default shape.
  const fetch = synthetic === undefined ? [...select] : select.filter(field => !synthetic.has(field));
  return { select, fields, fetch };
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
function withGuidance<T>(run: () => Promise<T>, hints: readonly string[] = ENTITLEMENT_MANAGEMENT_DENIAL_HINTS): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const ENTITLEMENT_MANAGEMENT_DENIAL_HINTS = [
  "Entitlement-management reads need EntitlementManagement.Read.All plus a supported Entra role with catalog visibility for delegated access (Global Reader and Identity Governance Administrator are among the supported roles), or admin-consented EntitlementManagement.Read.All for application access; delegated personal Microsoft accounts are not supported",
  "Entitlement management needs P2 or ID Governance depending on capability, not one uniform licence",
];

// Assignment and request reads document their own least-privileged
// delegation roles, so denials name Catalog reader first rather than the
// catalog-visibility roles above.
const ASSIGNMENT_DENIAL_HINTS = [
  "Assignment and request reads need EntitlementManagement.Read.All plus Catalog reader (least privileged), Catalog creator or Access package manager, or a supported Entra role (Security Reader, Global Reader, Compliance Administrator, Security Administrator or Identity Governance Administrator) for delegated access, or admin-consented EntitlementManagement.Read.All (a supported role alone also works) for application access; delegated personal Microsoft accounts are not supported",
  "Entitlement management needs P2 or ID Governance depending on capability, not one uniform licence",
];

function objectField(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

function stringField(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

// The fixed $expand returns whole target and access-package objects; only
// the flattened linkage identifiers below survive projection, never subject
// contact detail or package contents.
function flattenAssignmentLinks(source: Record<string, unknown>): Record<string, unknown> {
  const links: Record<string, unknown> = {};
  const target = objectField(source.target);
  const targetId = target === undefined ? undefined : stringField(target.id);
  if (targetId !== undefined) links.targetId = targetId;
  const targetDisplayName = target === undefined ? undefined : stringField(target.displayName);
  if (targetDisplayName !== undefined) links.targetDisplayName = targetDisplayName;
  const accessPackageId = stringField(objectField(source.accessPackage)?.id);
  if (accessPackageId !== undefined) links.accessPackageId = accessPackageId;
  return links;
}

// Requests flatten to bare linkage ids: the requestor subject and the
// expanded assignment and package detail stay out.
function flattenRequestLinks(source: Record<string, unknown>): Record<string, unknown> {
  const links: Record<string, unknown> = {};
  const accessPackageId = stringField(objectField(source.accessPackage)?.id);
  if (accessPackageId !== undefined) links.accessPackageId = accessPackageId;
  const assignmentId = stringField(objectField(source.assignment)?.id);
  if (assignmentId !== undefined) links.assignmentId = assignmentId;
  return links;
}

interface CollectionShape {
  command: string;
  key: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  // Flattened linkage identifiers ride a fixed $expand: always sent,
  // never a user flag, stripped from server $select, always projectable.
  synthetic?: ReadonlySet<string>;
  expand?: string;
  links?: (source: Record<string, unknown>) => Record<string, unknown>;
  hints?: readonly string[];
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
    "Access packages are the assignable bundles; assignment and request detail lives on the assignment reads, never on these reads",
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

const ASSIGNMENT_LIST: CollectionShape = {
  command: "entra entitlement assignment list",
  key: "assignments",
  known: ASSIGNMENT_KNOWN,
  knownList: KNOWN_ASSIGNMENT_FIELDS,
  defaultSelect: DEFAULT_ASSIGNMENT_LIST_SELECT,
  synthetic: ASSIGNMENT_SYNTHETIC,
  expand: ASSIGNMENT_EXPAND,
  links: flattenAssignmentLinks,
  hints: ASSIGNMENT_DENIAL_HINTS,
  emptyNote: "0 assignments matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Assignments are delivered package grants: rows carry flattened linkage (targetId, targetDisplayName, accessPackageId) from one fixed documented $expand, never nested subject or package detail",
    "Callers with only catalog-scoped roles must filter to one access package, e.g. --filter \"accessPackage/id eq '<access-package-id>'\"",
    `Show one assignment: mg-axi entra entitlement assignment show --id <assignment-id> ${profileHint(profileName)}`,
  ],
};

const ASSIGNMENT_REQUEST_LIST: CollectionShape = {
  command: "entra entitlement assignment-request list",
  key: "assignmentRequests",
  known: ASSIGNMENT_REQUEST_KNOWN,
  knownList: KNOWN_ASSIGNMENT_REQUEST_FIELDS,
  defaultSelect: DEFAULT_ASSIGNMENT_REQUEST_LIST_SELECT,
  synthetic: ASSIGNMENT_REQUEST_SYNTHETIC,
  expand: ASSIGNMENT_REQUEST_EXPAND,
  links: flattenRequestLinks,
  hints: ASSIGNMENT_DENIAL_HINTS,
  emptyNote: "0 assignment requests matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Requests carry flattened linkage (accessPackageId, assignmentId) from one fixed documented $expand; justification text and answers ride only behind an explicit --select, and the requestor subject belongs to a later slice",
    "Callers with only catalog-scoped roles must filter to one access package, e.g. --filter \"accessPackage/id eq '<access-package-id>'\"",
    `Show one request: mg-axi entra entitlement assignment-request show --id <request-id> ${profileHint(profileName)}`,
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
  const isSyntheticResume = shape.synthetic !== undefined && cursor !== undefined && flags.select === undefined && flags.fields === undefined;
  const { select, fields, fetch } = selectedFields(flags,
    savedSelect === undefined ? (isSyntheticResume ? [...shape.synthetic!] : shape.defaultSelect) : fieldList(savedSelect, "select", shape.known, shape.knownList, help),
    shape.known, shape.knownList, help, shape.synthetic);
  if (isSyntheticResume && savedSelect !== undefined) {
    const stripped = shape.defaultSelect.filter(field => !shape.synthetic!.has(field));
    const stored = fieldList(savedSelect, "select", shape.known, shape.knownList, help);
    if (stored.length === stripped.length && stripped.every(field => stored.includes(field))) {
      for (const field of shape.defaultSelect) if (!select.includes(field)) select.push(field);
      fields.splice(0, fields.length, ...select);
    }
  }
  const savedFilter = saved?.$filter;
  const filter = flags.filter === undefined ? savedFilter : String(flags.filter);
  const scopes = scopesFor(flags, profile, DEFAULT_ENTITLEMENT_MANAGEMENT_SCOPES, help);
  const full = flags.full === true;
  // These lists document the general OData query parameters, so --filter
  // passes through as plain $filter with no $count or ConsistencyLevel;
  // $search/$expand/$orderby stay unreviewed and $top/$skip stay
  // server-side paging concerns. Assignment and request lists additionally
  // send their one fixed documented $expand (never a user flag); resumed
  // queries replay the stored $select verbatim so resume never conflicts.
  const sendSelect = cursor === undefined
    ? fetch
    : saved?.$select === undefined ? [] : fieldList(saved.$select, "select", shape.known, shape.knownList, help);
  const query: Record<string, string> = {};
  if (sendSelect.length) query.$select = sendSelect.join(",");
  if (shape.expand !== undefined) query.$expand = shape.expand;
  if (filter !== undefined) query.$filter = filter;
  const args: CollectArgs = { profile, operation, params, query, scopes };
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(() => session.collect(args), shape.hints);
  const effectiveFlags: EntitlementManagementFlags = { ...flags, ...(result.query.$select === undefined ? {} : { select: select.join(",") }) };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, full, shape.links);
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
  extra?: {
    expand?: string;
    links?: (source: Record<string, unknown>) => Record<string, unknown>;
    synthetic?: ReadonlySet<string>;
    hints?: readonly string[];
  },
): Promise<Record<string, unknown>> {
  const { select, fields, fetch } = selectedFields(flags, defaultSelect, known, knownList, help, extra?.synthetic);
  const scopes = scopesFor(flags, profile, DEFAULT_ENTITLEMENT_MANAGEMENT_SCOPES, help);
  const full = flags.full === true;
  const query: Record<string, string> = {};
  if (fetch.length) query.$select = fetch.join(",");
  if (extra?.expand !== undefined) query.$expand = extra.expand;
  const raw = await withGuidance(() => session.execute({
    profile,
    operation,
    params,
    query,
    scopes,
  }), extra?.hints);
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError(malformed, "GRAPH_ERROR", [
      "Single-object reads carry one object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full, extra?.links);
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
  hints?: readonly string[],
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, profile, DEFAULT_ENTITLEMENT_MANAGEMENT_SCOPES, help);
  // The $count route returns a text/plain integer scalar rather than a
  // JSON collection, so the leaf reads it through session.execute with
  // scalar mode and accepts only a non-negative integer. The $count route
  // carries no operation-level documentation page and takes no
  // --filter/--select/--limit/--cursor.
  const raw = await withGuidance(() => session.execute({ profile, operation, scopes, scalar: true }), hints);
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

export async function listAssignments(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ASSIGNMENT_LIST, session, flags, profile, operation, help, profileName, {},
    `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showAssignment(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("assignment", "entra entitlement assignment show",
    ASSIGNMENT_KNOWN, KNOWN_ASSIGNMENT_FIELDS, DEFAULT_ASSIGNMENT_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessPackageAssignment-id": String(flags.id) },
    "Graph returned a malformed assignment body",
    { expand: ASSIGNMENT_EXPAND, links: flattenAssignmentLinks, synthetic: ASSIGNMENT_SYNTHETIC, hints: ASSIGNMENT_DENIAL_HINTS });
}

export async function countAssignments(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return countCollection("assignment", ASSIGNMENT_LIST.emptyNote, session, flags, profile, operation, help, profileName, ASSIGNMENT_DENIAL_HINTS);
}

export async function listAssignmentRequests(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ASSIGNMENT_REQUEST_LIST, session, flags, profile, operation, help, profileName, {},
    `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showAssignmentRequest(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("assignmentRequest", "entra entitlement assignment-request show",
    ASSIGNMENT_REQUEST_KNOWN, KNOWN_ASSIGNMENT_REQUEST_FIELDS, DEFAULT_ASSIGNMENT_REQUEST_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessPackageAssignmentRequest-id": String(flags.id) },
    "Graph returned a malformed assignment request body",
    { expand: ASSIGNMENT_REQUEST_EXPAND, links: flattenRequestLinks, synthetic: ASSIGNMENT_REQUEST_SYNTHETIC, hints: ASSIGNMENT_DENIAL_HINTS });
}

export async function countAssignmentRequests(
  session: GraphSession,
  flags: EntitlementManagementFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return countCollection("assignment request", ASSIGNMENT_REQUEST_LIST.emptyNote, session, flags, profile, operation, help, profileName, ASSIGNMENT_DENIAL_HINTS);
}
