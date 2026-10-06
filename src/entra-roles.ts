import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// READ-09: the Entra directory-role and PIM read mapping behind
// `mg-axi entra directory-role list/show`, `mg-axi entra directory-role-template
// list/show`, `mg-axi entra directory-role member list`, `mg-axi entra
// directory-role scoped-member list/show`, `mg-axi entra scoped-role-membership
// list/show`, `mg-axi entra role-assignment list`, `mg-axi entra pim eligible
// list` and `mg-axi entra pim active list`.
// Operation construction stays beside its command; the shared session owns
// URLs, credentials, paging, retries and error translation, and the SDK owns
// TOON rendering. This module only maps flags to session calls and projects
// rows for compact output.
//
// Reviewed against the v1.0 directoryrole-list, directoryrole-get,
// directoryroletemplate-list, directoryroletemplate-get,
// directoryrole-list-members, directoryrole-list-scopedmembers,
// scopedrolemembership-list, scopedrolemembership-get,
// rbacapplication-list-roleassignments,
// rbacapplication-list-roleeligibilityscheduleinstances and
// rbacapplication-list-roleassignmentscheduleinstances operation
// documentation on 2026-10-04 (roles, assignments and PIM) and 2026-10-06
// (templates, members, scoped members and scoped memberships). Directory-role reads, role
// assignments, templates, members, scoped members and scoped memberships
// read D/A RoleManagement.Read.Directory; eligible PIM reads
// D/A RoleEligibilitySchedule.Read.Directory; active PIM reads D/A
// RoleAssignmentSchedule.Read.Directory. Delegated callers additionally need
// a supported administrator role per operation. None of these collections
// documents an advanced-query contract, so --filter passes through as plain
// $filter with no $count or ConsistencyLevel attached.
//
// The views cover: directoryRoles are activated role
// instances only (a role appears here after activation, never before);
// directoryRoleTemplates describe every role definition whether activated or
// not; members are the principals assigned to one activated role (@odata.type
// names each member kind: user, group, device or service principal); scoped
// members bind one role to one principal within one directory scope;
// scopedRoleMemberships are the administrative-unit-scoped admin assignments
// (directory-wide assignments live on roleAssignments);
// roleAssignments include direct and PIM-activated assignments; eligible schedule
// instances are PIM-eligible but not active; active schedule instances cover
// both directly assigned (assignmentType Assigned) and activated eligible
// (assignmentType Activated) assignments. memberType names how the instance
// reaches the principal (for example Direct). Activation, assignment and
// every other PIM mutation belongs to later write slices, never to these
// reads. Built-in roles are base inventory; custom role assignments need P1;
// PIM reads need P2 or ID Governance.

// Eligible and active instances need separate reviewed property sets:
// only active instances expose assignmentType.
export const KNOWN_ROLE_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "description",
  "roleTemplateId",
  "deletedDateTime",
];
export const KNOWN_ASSIGNMENT_FIELDS: readonly string[] = [
  "id",
  "principalId",
  "roleDefinitionId",
  "directoryScopeId",
  "appScopeId",
];
export const KNOWN_ELIGIBLE_FIELDS: readonly string[] = [
  "id",
  "principalId",
  "roleDefinitionId",
  "directoryScopeId",
  "appScopeId",
  "memberType",
  "startDateTime",
  "endDateTime",
];
export const KNOWN_ACTIVE_FIELDS: readonly string[] = [...KNOWN_ELIGIBLE_FIELDS, "assignmentType"];
export const KNOWN_TEMPLATE_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "description",
];
export const KNOWN_ROLE_MEMBER_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "mail",
  "userPrincipalName",
];
export const KNOWN_SCOPED_MEMBER_FIELDS: readonly string[] = [
  "id",
  "principalId",
  "roleId",
  "directoryScopeId",
  "administrativeUnitId",
];
const ROLE_KNOWN = new Set(KNOWN_ROLE_FIELDS);
const ASSIGNMENT_KNOWN = new Set(KNOWN_ASSIGNMENT_FIELDS);
const ELIGIBLE_KNOWN = new Set(KNOWN_ELIGIBLE_FIELDS);
const ACTIVE_KNOWN = new Set(KNOWN_ACTIVE_FIELDS);
const TEMPLATE_KNOWN = new Set(KNOWN_TEMPLATE_FIELDS);
const MEMBER_KNOWN = new Set(KNOWN_ROLE_MEMBER_FIELDS);
const SCOPED_KNOWN = new Set(KNOWN_SCOPED_MEMBER_FIELDS);

// Compact rows: identifiers plus the correlation keys. For built-in roles
// the unified roleDefinitionId matches the directory-role roleTemplateId.
const DEFAULT_ROLE_LIST_SELECT = ["id", "displayName", "description", "roleTemplateId"];
const DEFAULT_ROLE_SHOW_SELECT = [...KNOWN_ROLE_FIELDS];
const DEFAULT_ASSIGNMENT_SELECT = ["id", "principalId", "roleDefinitionId", "directoryScopeId"];
const DEFAULT_ELIGIBLE_SELECT = ["id", "principalId", "roleDefinitionId", "memberType"];
const DEFAULT_ACTIVE_SELECT = ["id", "principalId", "roleDefinitionId", "assignmentType", "memberType"];
const DEFAULT_TEMPLATE_LIST_SELECT = ["id", "displayName", "description"];
const DEFAULT_TEMPLATE_SHOW_SELECT = [...KNOWN_TEMPLATE_FIELDS];
const DEFAULT_MEMBER_SELECT = ["id", "displayName"];
const DEFAULT_SCOPED_LIST_SELECT = ["id", "principalId", "roleId", "directoryScopeId"];
const DEFAULT_SCOPED_SHOW_SELECT = [...KNOWN_SCOPED_MEMBER_FIELDS];
export const DEFAULT_ROLE_SCOPES = ["https://graph.microsoft.com/RoleManagement.Read.Directory"];
export const DEFAULT_ELIGIBLE_SCOPES = ["https://graph.microsoft.com/RoleEligibilitySchedule.Read.Directory"];
export const DEFAULT_ACTIVE_SCOPES = ["https://graph.microsoft.com/RoleAssignmentSchedule.Read.Directory"];
const TRUNCATE_AT = 500;

export type RoleFlags = Record<string, string | boolean>;

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

function scopesFor(flags: RoleFlags, profile: AnyProfile, defaults: readonly string[], help: string): string[] | undefined {
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
  if (typeof value === "string" && !full && value.length > TRUNCATE_AT) {
    return { value: `${value.slice(0, TRUNCATE_AT)}... (truncated, ${value.length} chars total)`, truncated: true };
  }
  return { value, truncated: false };
}

// Local projection preserves Graph's null/missing distinction: an explicit
// null stays null, an absent property stays absent and is never synthesized.
function project(
  row: unknown,
  fields: string[],
  full: boolean,
  keepODataType = false,
): { row: Record<string, unknown>; truncated: boolean } {
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  const projected: Record<string, unknown> = {};
  // Heterogeneous member rows carry their kind in @odata.type, which is
  // structural rather than a selectable property, so it rides along first.
  if (keepODataType && typeof source["@odata.type"] === "string") {
    projected["@odata.type"] = source["@odata.type"];
  }
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
  flags: RoleFlags,
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

function fullHint(command: string, flags: RoleFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the roles and licensing that actually unlock it, because a
// 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const ROLE_DENIAL_HINTS = [
  "Directory-role reads need RoleManagement.Read.Directory plus a supported directory role (for example Global Reader, Security Reader or Privileged Role Administrator) for delegated access, or admin-consented RoleManagement.Read.Directory for application access",
  "Built-in roles are base inventory; custom role assignments need P1",
];

const ASSIGNMENT_DENIAL_HINTS = [
  "Role-assignment reads need RoleManagement.Read.Directory plus Directory Readers, Global Reader or Privileged Role Administrator for delegated access, or admin-consented RoleManagement.Read.Directory for application access",
  "Built-in role assignments are base inventory; custom role assignments need P1",
];

const ELIGIBLE_DENIAL_HINTS = [
  "Eligible PIM reads need RoleEligibilitySchedule.Read.Directory plus Global Reader, Security Operator, Security Reader, Security Administrator or Privileged Role Administrator for delegated access, or admin-consented RoleEligibilitySchedule.Read.Directory for application access",
  "PIM reads need P2 or ID Governance, not only P2",
];

const ACTIVE_DENIAL_HINTS = [
  "Active PIM reads need RoleAssignmentSchedule.Read.Directory plus Global Reader, Security Operator, Security Reader, Security Administrator or Privileged Role Administrator for delegated access, or admin-consented RoleAssignmentSchedule.Read.Directory for application access",
  "PIM reads need P2 or ID Governance, not only P2",
];

const SCOPED_DENIAL_HINTS = [
  "Scoped membership reads need RoleManagement.Read.Directory plus Global Reader, Security Reader or Privileged Role Administrator for delegated access, or admin-consented RoleManagement.Read.Directory for application access",
  "Administrative-unit-scoped memberships need P1",
];

interface CollectionShape {
  command: string;
  key: string;
  noun: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  defaultScopes: readonly string[];
  denialHints: string[];
  scopeNote: string;
  emptyNote: string;
  parent?: { param: string; flag: string; label: string };
  keepODataType?: boolean;
}

const ROLE_LIST: CollectionShape = {
  command: "entra directory-role list",
  key: "directoryRoles",
  noun: "directory roles",
  known: ROLE_KNOWN,
  knownList: KNOWN_ROLE_FIELDS,
  defaultSelect: DEFAULT_ROLE_LIST_SELECT,
  defaultScopes: DEFAULT_ROLE_SCOPES,
  denialHints: ROLE_DENIAL_HINTS,
  scopeNote: "Lists activated roles only; a role appears here after activation, never before",
  emptyNote: "0 activated roles matched; the absence of results is the answer, not an error",
};

const ASSIGNMENT_LIST: CollectionShape = {
  command: "entra role-assignment list",
  key: "roleAssignments",
  noun: "role assignments",
  known: ASSIGNMENT_KNOWN,
  knownList: KNOWN_ASSIGNMENT_FIELDS,
  defaultSelect: DEFAULT_ASSIGNMENT_SELECT,
  defaultScopes: DEFAULT_ROLE_SCOPES,
  denialHints: ASSIGNMENT_DENIAL_HINTS,
  scopeNote: "Current role-assignment inventory includes direct and PIM-activated assignments; use pim active list for assignmentType Assigned versus Activated",
  emptyNote: "0 current role assignments matched; the absence of results is the answer, not an error",
};

const PIM_ELIGIBLE: CollectionShape = {
  command: "entra pim eligible list",
  key: "eligibleAssignments",
  noun: "eligible assignments",
  known: ELIGIBLE_KNOWN,
  knownList: KNOWN_ELIGIBLE_FIELDS,
  defaultSelect: DEFAULT_ELIGIBLE_SELECT,
  defaultScopes: DEFAULT_ELIGIBLE_SCOPES,
  denialHints: ELIGIBLE_DENIAL_HINTS,
  scopeNote: "Eligible assignments are not active; activation is a PIM workflow outside these reads",
  emptyNote: "0 eligible assignments matched; the absence of results is the answer, not an error",
};

const PIM_ACTIVE: CollectionShape = {
  command: "entra pim active list",
  key: "activeAssignments",
  noun: "active assignments",
  known: ACTIVE_KNOWN,
  knownList: KNOWN_ACTIVE_FIELDS,
  defaultSelect: DEFAULT_ACTIVE_SELECT,
  defaultScopes: DEFAULT_ACTIVE_SCOPES,
  denialHints: ACTIVE_DENIAL_HINTS,
  scopeNote: "Active distinguishes directly assigned (assignmentType Assigned) and activated eligible (assignmentType Activated) assignments",
  emptyNote: "0 active assignments matched; the absence of results is the answer, not an error",
};

const TEMPLATE_LIST: CollectionShape = {
  command: "entra directory-role-template list",
  key: "directoryRoleTemplates",
  noun: "directory-role templates",
  known: TEMPLATE_KNOWN,
  knownList: KNOWN_TEMPLATE_FIELDS,
  defaultSelect: DEFAULT_TEMPLATE_LIST_SELECT,
  defaultScopes: DEFAULT_ROLE_SCOPES,
  denialHints: ROLE_DENIAL_HINTS,
  scopeNote: "Templates describe every role definition; activation state lives on directoryRoles, which appear only after activation",
  emptyNote: "0 directory-role templates matched; the absence of results is the answer, not an error",
};

const MEMBER_LIST: CollectionShape = {
  command: "entra directory-role member list",
  key: "roleMembers",
  noun: "role members",
  known: MEMBER_KNOWN,
  knownList: KNOWN_ROLE_MEMBER_FIELDS,
  defaultSelect: DEFAULT_MEMBER_SELECT,
  defaultScopes: DEFAULT_ROLE_SCOPES,
  denialHints: ROLE_DENIAL_HINTS,
  scopeNote: "Members are the principals assigned to the activated role; @odata.type names each member kind (user, group, device or service principal)",
  emptyNote: "0 role members matched; the absence of results is the answer, not an error",
  parent: { param: "directoryRole-id", flag: "role", label: "activated directory-role object ID whose members are listed" },
  keepODataType: true,
};

const SCOPED_LIST: CollectionShape = {
  command: "entra directory-role scoped-member list",
  key: "scopedMembers",
  noun: "scoped members",
  known: SCOPED_KNOWN,
  knownList: KNOWN_SCOPED_MEMBER_FIELDS,
  defaultSelect: DEFAULT_SCOPED_LIST_SELECT,
  defaultScopes: DEFAULT_ROLE_SCOPES,
  denialHints: SCOPED_DENIAL_HINTS,
  scopeNote: "Scoped members bind one role to one principal within one directory scope; administrative-unit scopes need P1",
  emptyNote: "0 scoped members matched; the absence of results is the answer, not an error",
  parent: { param: "directoryRole-id", flag: "role", label: "activated directory-role object ID whose scoped members are listed" },
};

const MEMBERSHIP_LIST: CollectionShape = {
  command: "entra scoped-role-membership list",
  key: "scopedRoleMemberships",
  noun: "scoped role memberships",
  known: SCOPED_KNOWN,
  knownList: KNOWN_SCOPED_MEMBER_FIELDS,
  defaultSelect: DEFAULT_SCOPED_LIST_SELECT,
  defaultScopes: DEFAULT_ROLE_SCOPES,
  denialHints: SCOPED_DENIAL_HINTS,
  scopeNote: "Scoped role memberships are administrative-unit-scoped admin assignments; directory-wide assignments live on roleAssignments",
  emptyNote: "0 scoped role memberships matched; the absence of results is the answer, not an error",
};

async function listCollection(
  shape: CollectionShape,
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  let params: Record<string, string> | undefined;
  if (shape.parent !== undefined) {
    const parentId = String(flags[shape.parent.flag] ?? "");
    if (!parentId.trim()) throw new AxiError(`--${shape.parent.flag} needs the ${shape.parent.label}`, "VALIDATION_ERROR", [help]);
    params = { [shape.parent.param]: parentId };
  }
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const savedSelect = saved?.$select;
  const { select, fields } = selectedFields(flags,
    savedSelect === undefined ? shape.defaultSelect : fieldList(savedSelect, "select", shape.known, shape.knownList, help),
    shape.known, shape.knownList, help);
  const savedFilter = saved?.$filter;
  const filter = flags.filter === undefined ? savedFilter : String(flags.filter);
  const scopes = scopesFor(flags, profile, shape.defaultScopes, help);
  const full = flags.full === true;
  // These collections document $select/$filter/$expand only, so --filter
  // passes through as plain $filter with no $count or ConsistencyLevel.
  const query: Record<string, string> = { $select: select.join(",") };
  if (filter !== undefined) query.$filter = filter;
  const args: CollectArgs = { profile, operation, query, scopes };
  if (params !== undefined) args.params = params;
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(shape.denialHints, () => session.collect(args));
  const effectiveFlags: RoleFlags = { ...flags, select: result.query.$select ?? shape.defaultSelect.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, full, shape.keepODataType === true);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = shape.key === "directoryRoles" ? `mg-axi entra directory-role show --id <role-id> ${profileHint(profileName)}`
    : shape.key === "directoryRoleTemplates" ? `mg-axi entra directory-role-template show --id <template-id> ${profileHint(profileName)}`
    : undefined;
  const correlateHint = shape.key === "directoryRoles"
    ? undefined
    : `For built-in roles, roleDefinitionId matches the directory-role roleTemplateId: mg-axi entra directory-role list ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(shape.command, effectiveFlags, profileName)] : [];
  const standing = [shape.scopeNote, ...(showHint === undefined ? [] : [showHint]), ...(correlateHint === undefined ? [] : [correlateHint])];
  if (!result.complete) {
    return {
      [shape.key]: rows,
      ...listTotals(rows.length, result.total, shape.noun, false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor <cursor-from-output> ${profileHint(profileName)}`, ...standing],
    };
  }
  if (!rows.length) {
    return { [shape.key]: rows, ...listTotals(rows.length, result.total, shape.noun, true), complete: true, help: [shape.emptyNote, ...standing] };
  }
  return { [shape.key]: rows, ...listTotals(rows.length, result.total, shape.noun, true), complete: true, help: [...truncationHints, ...standing] };
}

export async function listDirectoryRoles(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ROLE_LIST, session, flags, profile, operation, help, profileName);
}

export async function showDirectoryRole(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, DEFAULT_ROLE_SHOW_SELECT, ROLE_KNOWN, KNOWN_ROLE_FIELDS, help);
  const scopes = scopesFor(flags, profile, DEFAULT_ROLE_SCOPES, help);
  const full = flags.full === true;
  const raw = await withGuidance(ROLE_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "directoryRole-id": String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed directory-role body", "GRAPH_ERROR", [
      "Single-role reads carry one role object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { directoryRole: row, help: [fullHint("entra directory-role show", flags, profileName)] };
  return { directoryRole: row };
}

export async function listRoleAssignments(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ASSIGNMENT_LIST, session, flags, profile, operation, help, profileName);
}

export async function listPimEligible(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(PIM_ELIGIBLE, session, flags, profile, operation, help, profileName);
}

export async function listPimActive(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(PIM_ACTIVE, session, flags, profile, operation, help, profileName);
}

interface SingleShape {
  command: string;
  key: string;
  noun: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  defaultScopes: readonly string[];
  denialHints: string[];
}

const TEMPLATE_SINGLE: SingleShape = {
  command: "entra directory-role-template show",
  key: "directoryRoleTemplate",
  noun: "directory-role template",
  known: TEMPLATE_KNOWN,
  knownList: KNOWN_TEMPLATE_FIELDS,
  defaultSelect: DEFAULT_TEMPLATE_SHOW_SELECT,
  defaultScopes: DEFAULT_ROLE_SCOPES,
  denialHints: ROLE_DENIAL_HINTS,
};

const SCOPED_SINGLE: SingleShape = {
  command: "entra directory-role scoped-member show",
  key: "scopedMember",
  noun: "scoped member",
  known: SCOPED_KNOWN,
  knownList: KNOWN_SCOPED_MEMBER_FIELDS,
  defaultSelect: DEFAULT_SCOPED_SHOW_SELECT,
  defaultScopes: DEFAULT_ROLE_SCOPES,
  denialHints: SCOPED_DENIAL_HINTS,
};

const MEMBERSHIP_SINGLE: SingleShape = {
  command: "entra scoped-role-membership show",
  key: "scopedRoleMembership",
  noun: "scoped role membership",
  known: SCOPED_KNOWN,
  knownList: KNOWN_SCOPED_MEMBER_FIELDS,
  defaultSelect: DEFAULT_SCOPED_SHOW_SELECT,
  defaultScopes: DEFAULT_ROLE_SCOPES,
  denialHints: SCOPED_DENIAL_HINTS,
};

async function showSingle(
  shape: SingleShape,
  params: Record<string, string>,
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, shape.defaultSelect, shape.known, shape.knownList, help);
  const scopes = scopesFor(flags, profile, shape.defaultScopes, help);
  const full = flags.full === true;
  const raw = await withGuidance(shape.denialHints, () => session.execute({
    profile,
    operation,
    params,
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError(`Graph returned a malformed ${shape.noun} body`, "GRAPH_ERROR", [
      `Single-${shape.noun} reads carry one object; treat anything else as unknown, not empty`,
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { [shape.key]: row, help: [fullHint(shape.command, flags, profileName)] };
  return { [shape.key]: row };
}

export async function listDirectoryRoleTemplates(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(TEMPLATE_LIST, session, flags, profile, operation, help, profileName);
}

export async function showDirectoryRoleTemplate(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(TEMPLATE_SINGLE, { "directoryRoleTemplate-id": String(flags.id) }, session, flags, profile, operation, help, profileName);
}

export async function listRoleMembers(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(MEMBER_LIST, session, flags, profile, operation, help, profileName);
}

export async function listScopedMembers(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(SCOPED_LIST, session, flags, profile, operation, help, profileName);
}

export async function showScopedMember(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const roleId = String(flags.role ?? "");
  if (!roleId.trim()) throw new AxiError("--role needs the activated directory-role object ID owning the scoped member", "VALIDATION_ERROR", [help]);
  return showSingle(SCOPED_SINGLE, { "directoryRole-id": roleId, "scopedRoleMembership-id": String(flags.id) }, session, flags, profile, operation, help, profileName);
}

export async function listScopedRoleMemberships(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(MEMBERSHIP_LIST, session, flags, profile, operation, help, profileName);
}

export async function showScopedRoleMembership(
  session: GraphSession,
  flags: RoleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(MEMBERSHIP_SINGLE, { "scopedRoleMembership-id": String(flags.id) }, session, flags, profile, operation, help, profileName);
}
