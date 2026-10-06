import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-02 group PIM schedule reads: the read-only schedule and instance
// mapping behind `mg-axi entra pim group-assignment-schedule list/show`,
// `mg-axi entra pim group-assignment-instance list/show`,
// `mg-axi entra pim group-eligibility-schedule list/show`,
// `mg-axi entra pim group-eligibility-instance list/show` and
// `mg-axi entra pim group-eligibility-request list/show`.
// Operation construction stays beside its command; the shared session owns
// URLs, credentials, paging, retries and error translation, and the SDK owns
// TOON rendering. This module only maps flags to session calls and projects
// rows for compact output.
//
// Reviewed against the v1.0 privilegedaccessgroup-list-assignmentschedules,
// privilegedaccessgroupassignmentschedule-get,
// privilegedaccessgroup-list-assignmentscheduleinstances,
// privilegedaccessgroupassignmentscheduleinstance-get,
// privilegedaccessgroup-list-eligibilityschedules,
// privilegedaccessgroupeligibilityschedule-get,
// privilegedaccessgroup-list-eligibilityscheduleinstances,
// privilegedaccessgroupeligibilityscheduleinstance-get,
// privilegedaccessgroup-list-eligibilityschedulerequests and
// privilegedaccessgroupeligibilityschedulerequest-get operation
// documentation and the privilegedAccessGroupAssignmentSchedule,
// privilegedAccessGroupAssignmentScheduleInstance,
// privilegedAccessGroupEligibilitySchedule,
// privilegedAccessGroupEligibilityScheduleInstance and
// privilegedAccessGroupEligibilityScheduleRequest resource contracts on
// 2026-10-06, with a second check against the Graph permissions reference
// and the ID Governance licensing fundamentals. Assignment reads take D/A
// PrivilegedAssignmentSchedule.Read.AzureADGroup; eligibility reads take
// D/A PrivilegedEligibilitySchedule.Read.AzureADGroup. Delegated callers
// additionally need owner/member of the group or a supported directory
// role (Global Reader or Privileged Role Administrator for role-assignable
// groups; Global Reader, Directory Writer, Groups Administrator, Identity
// Governance Administrator or User Administrator otherwise), scoped at
// directory level; delegated personal Microsoft accounts are not
// supported. PIM for Groups needs P2 or ID Governance.
//
// Every list requires $filter (eq) scoped to a groupId or a principalId,
// so --filter is required and must name one of them; the requirement is
// validated before credentials. Lists document $select/$filter/$expand, so
// --filter passes through as plain $filter with no $count or
// ConsistencyLevel attached; singles document $select with $expand.
// These collections document no $count contract, so totals stay honest as
// count { returned, complete } like the extended role PIM reads rather
// than the shared list-totals helper, which applies where $count=true is
// sent. Deferred with reason, never built on a write scope:
// assignmentScheduleRequests list/get take
// PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup as least privilege.
// Out of scope: approvals, filterByCurrentUser functions, nested
// group/principal/activatedUsing/targetSchedule navigation (scalar linkage
// ids such as assignmentScheduleId stay in), resources, every
// count/ref/cast tail, all mutations and beta.

// Schedules carry the full lifecycle (scheduleInfo, status, audit stamps);
// instances carry one provisioned window plus its parent schedule linkage;
// requests carry the ask (action, justification, scheduleInfo, ticketInfo)
// plus its outcome (status, approvalId, targetScheduleId).
export const KNOWN_ASSIGNMENT_SCHEDULE_FIELDS: readonly string[] = [
  "id",
  "accessId",
  "assignmentType",
  "memberType",
  "principalId",
  "groupId",
  "status",
  "scheduleInfo",
  "createdDateTime",
  "modifiedDateTime",
  "createdUsing",
];
export const KNOWN_ASSIGNMENT_INSTANCE_FIELDS: readonly string[] = [
  "id",
  "accessId",
  "assignmentType",
  "memberType",
  "principalId",
  "groupId",
  "assignmentScheduleId",
  "startDateTime",
  "endDateTime",
];
export const KNOWN_ELIGIBILITY_SCHEDULE_FIELDS: readonly string[] = [
  "id",
  "accessId",
  "memberType",
  "principalId",
  "groupId",
  "status",
  "scheduleInfo",
  "createdDateTime",
  "modifiedDateTime",
  "createdUsing",
];
export const KNOWN_ELIGIBILITY_INSTANCE_FIELDS: readonly string[] = [
  "id",
  "accessId",
  "memberType",
  "principalId",
  "groupId",
  "eligibilityScheduleId",
  "startDateTime",
  "endDateTime",
];
// Requests narrow to the ask plus its outcome: createdBy is an identitySet
// for a later slice and customData is documented as unused, so neither is
// selectable here. Justification is free text and rides only behind an
// explicit --select, never in a default.
export const KNOWN_ELIGIBILITY_REQUEST_FIELDS: readonly string[] = [
  "id",
  "accessId",
  "action",
  "status",
  "principalId",
  "groupId",
  "justification",
  "scheduleInfo",
  "ticketInfo",
  "createdDateTime",
  "completedDateTime",
  "approvalId",
  "targetScheduleId",
  "isValidationOnly",
];
const ASSIGNMENT_SCHEDULE_KNOWN = new Set(KNOWN_ASSIGNMENT_SCHEDULE_FIELDS);
const ASSIGNMENT_INSTANCE_KNOWN = new Set(KNOWN_ASSIGNMENT_INSTANCE_FIELDS);
const ELIGIBILITY_SCHEDULE_KNOWN = new Set(KNOWN_ELIGIBILITY_SCHEDULE_FIELDS);
const ELIGIBILITY_INSTANCE_KNOWN = new Set(KNOWN_ELIGIBILITY_INSTANCE_FIELDS);
const ELIGIBILITY_REQUEST_KNOWN = new Set(KNOWN_ELIGIBILITY_REQUEST_FIELDS);

// Compact rows: identifiers plus the correlation keys. groupId and
// principalId are the required scoping keys; accessId names member versus
// owner; assignmentType names Assigned versus Activated; memberType names
// how the row reaches the principal (for example Direct).
const DEFAULT_ASSIGNMENT_SCHEDULE_LIST_SELECT = ["id", "principalId", "groupId", "accessId", "assignmentType"];
const DEFAULT_ASSIGNMENT_INSTANCE_LIST_SELECT = ["id", "principalId", "groupId", "accessId", "assignmentType"];
const DEFAULT_ELIGIBILITY_SCHEDULE_LIST_SELECT = ["id", "principalId", "groupId", "accessId", "memberType"];
const DEFAULT_ELIGIBILITY_INSTANCE_LIST_SELECT = ["id", "principalId", "groupId", "accessId", "memberType"];
const DEFAULT_ELIGIBILITY_REQUEST_LIST_SELECT = ["id", "action", "status", "principalId", "groupId", "accessId"];
const DEFAULT_ELIGIBILITY_REQUEST_SHOW_SELECT = KNOWN_ELIGIBILITY_REQUEST_FIELDS.filter(field => field !== "justification");
export const DEFAULT_ASSIGNMENT_SCOPES = ["https://graph.microsoft.com/PrivilegedAssignmentSchedule.Read.AzureADGroup"];
export const DEFAULT_ELIGIBILITY_SCOPES = ["https://graph.microsoft.com/PrivilegedEligibilitySchedule.Read.AzureADGroup"];
const TRUNCATE_AT = 500;

export type GroupPimFlags = Record<string, string | boolean>;

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

function scopesFor(flags: GroupPimFlags, profile: AnyProfile, defaults: readonly string[], help: string): string[] | undefined {
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

// Request justification is free text and scheduleInfo/ticketInfo are nested
// objects, so truncation walks values recursively; --full returns them whole.
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
  flags: GroupPimFlags,
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

function fullHint(command: string, flags: GroupPimFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// family adds the scope, roles and licensing that actually unlock it,
// because a 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const ASSIGNMENT_DENIAL_HINTS = [
  "Group PIM assignment reads need PrivilegedAssignmentSchedule.Read.AzureADGroup plus owner/member of the group or a supported directory role (Global Reader or Privileged Role Administrator for role-assignable groups; Global Reader, Directory Writer, Groups Administrator, Identity Governance Administrator or User Administrator otherwise, scoped at directory level) for delegated access, or admin-consented PrivilegedAssignmentSchedule.Read.AzureADGroup for application access; delegated personal Microsoft accounts are not supported",
  "PIM for Groups needs P2 or ID Governance, not only P2",
];

const ELIGIBILITY_DENIAL_HINTS = [
  "Group PIM eligibility reads need PrivilegedEligibilitySchedule.Read.AzureADGroup plus owner/member of the group or a supported directory role (Global Reader or Privileged Role Administrator for role-assignable groups; Global Reader, Directory Writer, Groups Administrator, Identity Governance Administrator or User Administrator otherwise, scoped at directory level) for delegated access, or admin-consented PrivilegedEligibilitySchedule.Read.AzureADGroup for application access; delegated personal Microsoft accounts are not supported",
  "PIM for Groups needs P2 or ID Governance, not only P2",
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
}

const ASSIGNMENT_SCHEDULE_LIST: CollectionShape = {
  command: "entra pim group-assignment-schedule list",
  key: "assignmentSchedules",
  noun: "assignment schedules",
  known: ASSIGNMENT_SCHEDULE_KNOWN,
  knownList: KNOWN_ASSIGNMENT_SCHEDULE_FIELDS,
  defaultSelect: DEFAULT_ASSIGNMENT_SCHEDULE_LIST_SELECT,
  defaultScopes: DEFAULT_ASSIGNMENT_SCOPES,
  denialHints: ASSIGNMENT_DENIAL_HINTS,
  scopeNote: "Assignment schedules govern group membership or ownership over time; instances are the provisioned windows, never the schedule itself",
  emptyNote: "0 assignment schedules matched; the absence of results is the answer, not an error",
};

const ASSIGNMENT_INSTANCE_LIST: CollectionShape = {
  command: "entra pim group-assignment-instance list",
  key: "assignmentScheduleInstances",
  noun: "assignment instances",
  known: ASSIGNMENT_INSTANCE_KNOWN,
  knownList: KNOWN_ASSIGNMENT_INSTANCE_FIELDS,
  defaultSelect: DEFAULT_ASSIGNMENT_INSTANCE_LIST_SELECT,
  defaultScopes: DEFAULT_ASSIGNMENT_SCOPES,
  denialHints: ASSIGNMENT_DENIAL_HINTS,
  scopeNote: "Assignment instances are provisioned membership or ownership windows; assignmentType Assigned names direct grants and Activated names activated eligible grants",
  emptyNote: "0 assignment instances matched; the absence of results is the answer, not an error",
};

const ELIGIBILITY_SCHEDULE_LIST: CollectionShape = {
  command: "entra pim group-eligibility-schedule list",
  key: "eligibilitySchedules",
  noun: "eligibility schedules",
  known: ELIGIBILITY_SCHEDULE_KNOWN,
  knownList: KNOWN_ELIGIBILITY_SCHEDULE_FIELDS,
  defaultSelect: DEFAULT_ELIGIBILITY_SCHEDULE_LIST_SELECT,
  defaultScopes: DEFAULT_ELIGIBILITY_SCOPES,
  denialHints: ELIGIBILITY_DENIAL_HINTS,
  scopeNote: "Eligibility schedules govern who may activate group membership or ownership; they are not active grants",
  emptyNote: "0 eligibility schedules matched; the absence of results is the answer, not an error",
};

const ELIGIBILITY_INSTANCE_LIST: CollectionShape = {
  command: "entra pim group-eligibility-instance list",
  key: "eligibilityScheduleInstances",
  noun: "eligibility instances",
  known: ELIGIBILITY_INSTANCE_KNOWN,
  knownList: KNOWN_ELIGIBILITY_INSTANCE_FIELDS,
  defaultSelect: DEFAULT_ELIGIBILITY_INSTANCE_LIST_SELECT,
  defaultScopes: DEFAULT_ELIGIBILITY_SCOPES,
  denialHints: ELIGIBILITY_DENIAL_HINTS,
  scopeNote: "Eligibility instances are provisioned eligibility windows; activation itself is a PIM workflow outside these reads",
  emptyNote: "0 eligibility instances matched; the absence of results is the answer, not an error",
};

const ELIGIBILITY_REQUEST_LIST: CollectionShape = {
  command: "entra pim group-eligibility-request list",
  key: "eligibilityScheduleRequests",
  noun: "eligibility requests",
  known: ELIGIBILITY_REQUEST_KNOWN,
  knownList: KNOWN_ELIGIBILITY_REQUEST_FIELDS,
  defaultSelect: DEFAULT_ELIGIBILITY_REQUEST_LIST_SELECT,
  defaultScopes: DEFAULT_ELIGIBILITY_SCOPES,
  denialHints: ELIGIBILITY_DENIAL_HINTS,
  scopeNote: "Eligibility requests carry the ask and its outcome; justification text rides only behind an explicit --select",
  emptyNote: "0 eligibility requests matched; the absence of results is the answer, not an error",
};

// Graph requires every one of these lists to be scoped to a groupId or a
// principalId, so --filter must name one of them; validated before any
// credential or transport is touched, with cursor resumes replaying the
// stored filter.
function requiredFilter(flags: GroupPimFlags, savedFilter: string | undefined, help: string): string {
  const filter = flags.filter === undefined ? savedFilter : String(flags.filter);
  if (filter === undefined || (!filter.includes("groupId") && !filter.includes("principalId"))) {
    throw new AxiError("--filter needs a groupId or principalId eq clause scoping the list", "VALIDATION_ERROR", [
      help,
      "For example --filter \"groupId eq '<group-id>'\" or --filter \"principalId eq '<principal-id>'\"",
    ]);
  }
  return filter;
}

async function listCollection(
  shape: CollectionShape,
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const savedSelect = saved?.$select;
  const { select, fields } = selectedFields(flags,
    savedSelect === undefined ? shape.defaultSelect : fieldList(savedSelect, "select", shape.known, shape.knownList, help),
    shape.known, shape.knownList, help);
  const filter = requiredFilter(flags, saved?.$filter, help);
  const scopes = scopesFor(flags, profile, shape.defaultScopes, help);
  const full = flags.full === true;
  // These lists document $select/$filter/$expand, so --filter passes
  // through as plain $filter with no $count or ConsistencyLevel attached;
  // $expand stays unreviewed and is never sent.
  const query: Record<string, string> = { $select: select.join(","), $filter: filter };
  const args: CollectArgs = { profile, operation, query, scopes };
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(shape.denialHints, () => session.collect(args));
  const effectiveFlags: GroupPimFlags = { ...flags, select: result.query.$select ?? shape.defaultSelect.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const truncationHints = truncated ? [fullHint(shape.command, effectiveFlags, profileName)] : [];
  const standing = [shape.scopeNote];
  if (!result.complete) {
    return {
      [shape.key]: rows,
      ...listTotals(rows.length, result.total, shape.noun, false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same --filter and flags plus --cursor <cursor-from-output> ${profileHint(profileName)}`, ...standing],
    };
  }
  if (!rows.length) {
    return { [shape.key]: rows, ...listTotals(rows.length, result.total, shape.noun, true), complete: true, help: [shape.emptyNote, ...standing] };
  }
  return { [shape.key]: rows, ...listTotals(rows.length, result.total, shape.noun, true), complete: true, help: [...truncationHints, ...standing] };
}

async function showOne(
  key: string,
  command: string,
  known: Set<string>,
  knownList: readonly string[],
  defaultSelect: string[],
  defaultScopes: readonly string[],
  denialHints: string[],
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  params: Record<string, string>,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, defaultSelect, known, knownList, help);
  const scopes = scopesFor(flags, profile, defaultScopes, help);
  const full = flags.full === true;
  const raw = await withGuidance(denialHints, () => session.execute({
    profile,
    operation,
    params,
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError(`Graph returned a malformed ${key} body`, "GRAPH_ERROR", [
      "Single-object reads carry one object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { [key]: row, help: [fullHint(command, flags, profileName)] };
  return { [key]: row };
}

export async function listGroupAssignmentSchedules(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ASSIGNMENT_SCHEDULE_LIST, session, flags, profile, operation, help, profileName);
}

export async function showGroupAssignmentSchedule(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("assignmentSchedule", "entra pim group-assignment-schedule show",
    ASSIGNMENT_SCHEDULE_KNOWN, KNOWN_ASSIGNMENT_SCHEDULE_FIELDS, [...KNOWN_ASSIGNMENT_SCHEDULE_FIELDS],
    DEFAULT_ASSIGNMENT_SCOPES, ASSIGNMENT_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName,
    { "privilegedAccessGroupAssignmentSchedule-id": String(flags.id) });
}

export async function listGroupAssignmentInstances(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ASSIGNMENT_INSTANCE_LIST, session, flags, profile, operation, help, profileName);
}

export async function showGroupAssignmentInstance(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("assignmentScheduleInstance", "entra pim group-assignment-instance show",
    ASSIGNMENT_INSTANCE_KNOWN, KNOWN_ASSIGNMENT_INSTANCE_FIELDS, [...KNOWN_ASSIGNMENT_INSTANCE_FIELDS],
    DEFAULT_ASSIGNMENT_SCOPES, ASSIGNMENT_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName,
    { "privilegedAccessGroupAssignmentScheduleInstance-id": String(flags.id) });
}

export async function listGroupEligibilitySchedules(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ELIGIBILITY_SCHEDULE_LIST, session, flags, profile, operation, help, profileName);
}

export async function showGroupEligibilitySchedule(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("eligibilitySchedule", "entra pim group-eligibility-schedule show",
    ELIGIBILITY_SCHEDULE_KNOWN, KNOWN_ELIGIBILITY_SCHEDULE_FIELDS, [...KNOWN_ELIGIBILITY_SCHEDULE_FIELDS],
    DEFAULT_ELIGIBILITY_SCOPES, ELIGIBILITY_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName,
    { "privilegedAccessGroupEligibilitySchedule-id": String(flags.id) });
}

export async function listGroupEligibilityInstances(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ELIGIBILITY_INSTANCE_LIST, session, flags, profile, operation, help, profileName);
}

export async function showGroupEligibilityInstance(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("eligibilityScheduleInstance", "entra pim group-eligibility-instance show",
    ELIGIBILITY_INSTANCE_KNOWN, KNOWN_ELIGIBILITY_INSTANCE_FIELDS, [...KNOWN_ELIGIBILITY_INSTANCE_FIELDS],
    DEFAULT_ELIGIBILITY_SCOPES, ELIGIBILITY_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName,
    { "privilegedAccessGroupEligibilityScheduleInstance-id": String(flags.id) });
}

export async function listGroupEligibilityRequests(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(ELIGIBILITY_REQUEST_LIST, session, flags, profile, operation, help, profileName);
}

export async function showGroupEligibilityRequest(
  session: GraphSession,
  flags: GroupPimFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("eligibilityScheduleRequest", "entra pim group-eligibility-request show",
    ELIGIBILITY_REQUEST_KNOWN, KNOWN_ELIGIBILITY_REQUEST_FIELDS, [...DEFAULT_ELIGIBILITY_REQUEST_SHOW_SELECT],
    DEFAULT_ELIGIBILITY_SCOPES, ELIGIBILITY_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName,
    { "privilegedAccessGroupEligibilityScheduleRequest-id": String(flags.id) });
}
