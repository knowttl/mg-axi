import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-02 access-reviews subfamily: the read-only access-review mapping
// behind `mg-axi entra access-review definition list/show`,
// `mg-axi entra access-review instance list/show` and
// `mg-axi entra access-review decision list`. Operation construction stays
// beside its command; the shared session owns URLs, credentials, paging,
// retries and error translation, and the SDK owns TOON rendering. This
// module only maps flags to session calls and projects rows for compact
// output.
//
// Reviewed against the v1.0 accessreviewset-list-definitions,
// accessreviewscheduledefinition-get,
// accessreviewscheduledefinition-list-instances, accessreviewinstance-get,
// accessreviewinstance-list-decisions,
// accessreviewinstancedecisionitem-get,
// accessreviewinstance-list-contactedreviewers,
// accessreviewinstance-list-stages and accessreviewstage-get operation
// documentation, plus the accessReviewReviewer and accessReviewStage
// resource references, on 2026-10-04. All ten reads take D/A
// AccessReview.Read.All as the least privileged choice; delegated personal
// Microsoft accounts are not supported. Delegated callers additionally need
// a supported administrator role: group or app reviews need the review
// creator, Global Reader, Security Reader, User Administrator, Identity
// Governance Administrator or Security Administrator, while Entra-role
// reviews need Security Reader, Identity Governance Administrator,
// Privileged Role Administrator or Security Administrator. Definition,
// instance and decision collections document $select with plain $filter
// ($orderby/$skip/$top stay server-side paging concerns); contacted-reviewer
// lists add $orderby to the documented set and stage lists document $filter
// (eq only), both still passed through as plain $filter with no $count or
// ConsistencyLevel attached. Singles document $select only; the
// contacted-reviewer single reuses its list's resource contract because the
// resource carries no relationships and the list returns all nested
// properties. None of these collections documents an advanced-query
// contract.
//
// Definitions are review schedules (a series: one recurrence creates one
// instance per reviewed resource, and a one-time review creates one
// instance per resource); instances are the occurrences under review; each
// reviewed principal or resource in an instance carries one decision item.
// Listing decisions never approves, denies or applies anything: submitting,
// stopping and every other access-review mutation belongs to later slices,
// never to these reads. Access reviews need P2 or ID Governance depending
// on capability, never one uniform licence.

// Definitions expose the series schedule; the deprecated backupReviewers
// alias, createdBy, stageSettings and additionalNotificationRecipients carry
// no documented $select contract and stay out of the reviewed set.
export const KNOWN_DEFINITION_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "status",
  "createdDateTime",
  "lastModifiedDateTime",
  "descriptionForAdmins",
  "descriptionForReviewers",
  "scope",
  "instanceEnumerationScope",
  "reviewers",
  "fallbackReviewers",
  "settings",
];
// Instances expose occurrence state; errors carry no documented $select
// contract and stay out of the reviewed set.
export const KNOWN_INSTANCE_FIELDS: readonly string[] = [
  "id",
  "startDateTime",
  "endDateTime",
  "status",
  "scope",
  "reviewers",
  "fallbackReviewers",
];
// Decisions expose review outcomes; target is the reviewed
// principal-or-resource item shown in the documented list response, while
// appliedBy, applyDescription, justification, permission, principalLink and
// resource carry no documented $select contract and stay out.
export const KNOWN_DECISION_FIELDS: readonly string[] = [
  "id",
  "accessReviewId",
  "decision",
  "recommendation",
  "reviewedDateTime",
  "reviewedBy",
  "appliedDateTime",
  "applyResult",
  "principal",
  "resourceLink",
  "target",
];
// Contacted reviewers are the reviewer identities recorded on one instance,
// whether or not they were notified; the resource carries no relationships
// and the list returns all nested properties, so list and show share one
// reviewed set.
export const KNOWN_CONTACTED_REVIEWER_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "userPrincipalName",
  "createdDateTime",
];
// Stages are the sequential phases of one instance (up to three when the
// definition sets stageSettings); each stage carries its own reviewer
// scopes, and durationInDays is not a stage property (it folds into
// endDateTime). Decisions hang off each stage but belong to a later slice.
export const KNOWN_STAGE_FIELDS: readonly string[] = [
  "id",
  "startDateTime",
  "endDateTime",
  "status",
  "reviewers",
  "fallbackReviewers",
];
const DEFINITION_KNOWN = new Set(KNOWN_DEFINITION_FIELDS);
const INSTANCE_KNOWN = new Set(KNOWN_INSTANCE_FIELDS);
const DECISION_KNOWN = new Set(KNOWN_DECISION_FIELDS);
const CONTACTED_REVIEWER_KNOWN = new Set(KNOWN_CONTACTED_REVIEWER_FIELDS);
const STAGE_KNOWN = new Set(KNOWN_STAGE_FIELDS);

// Compact rows: identifiers plus schedule state, occurrence or stage window,
// reviewer identity or review outcome. The definition id is the schedule;
// accessReviewId on a decision names its parent instance.
const DEFAULT_DEFINITION_LIST_SELECT = ["id", "displayName", "status"];
const DEFAULT_DEFINITION_SHOW_SELECT = [...KNOWN_DEFINITION_FIELDS];
const DEFAULT_INSTANCE_LIST_SELECT = ["id", "status", "startDateTime", "endDateTime"];
const DEFAULT_INSTANCE_SHOW_SELECT = [...KNOWN_INSTANCE_FIELDS];
const DEFAULT_DECISION_SELECT = ["id", "accessReviewId", "decision", "recommendation"];
const DEFAULT_DECISION_SHOW_SELECT = [...KNOWN_DECISION_FIELDS];
const DEFAULT_CONTACTED_REVIEWER_LIST_SELECT = ["id", "displayName", "userPrincipalName"];
const DEFAULT_CONTACTED_REVIEWER_SHOW_SELECT = [...KNOWN_CONTACTED_REVIEWER_FIELDS];
const DEFAULT_STAGE_LIST_SELECT = ["id", "status", "startDateTime", "endDateTime"];
const DEFAULT_STAGE_SHOW_SELECT = [...KNOWN_STAGE_FIELDS];
export const DEFAULT_ACCESS_REVIEW_SCOPES = ["https://graph.microsoft.com/AccessReview.Read.All"];
const TRUNCATE_AT = 500;

export type AccessReviewFlags = Record<string, string | boolean>;

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

function scopesFor(flags: AccessReviewFlags, profile: AnyProfile, defaults: readonly string[], help: string): string[] | undefined {
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
  flags: AccessReviewFlags,
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

function fullHint(command: string, flags: AccessReviewFlags, profileName: string): string {
  const args = Object.entries({ ...flags, ...(flags.cursor === undefined ? {} : { cursor: "-" }), profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; access
// reviews add the roles and licensing that actually unlock them, because a
// 403 alone never says which prerequisite is missing.
function withGuidance<T>(run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...ACCESS_REVIEW_DENIAL_HINTS, ...error.suggestions]);
    }
    throw error;
  });
}

const ACCESS_REVIEW_DENIAL_HINTS = [
  "Access-review reads need AccessReview.Read.All plus a supported Microsoft Entra role for delegated access (group or app reviews: the review creator, Global Reader, Security Reader, User Administrator, Identity Governance Administrator or Security Administrator; Entra-role reviews: Security Reader, Identity Governance Administrator, Privileged Role Administrator or Security Administrator), or admin-consented AccessReview.Read.All for application access; delegated personal Microsoft accounts are not supported",
  "Access reviews need P2 or ID Governance depending on capability, not one uniform licence",
];

interface CollectionShape {
  command: string;
  key: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  denialScopeNote: string;
  emptyNote: string;
  standing: (profileName: string) => string[];
}

const DEFINITION_LIST: CollectionShape = {
  command: "entra access-review definition list",
  key: "definitions",
  known: DEFINITION_KNOWN,
  knownList: KNOWN_DEFINITION_FIELDS,
  defaultSelect: DEFAULT_DEFINITION_LIST_SELECT,
  denialScopeNote: "Definitions are review schedules (a series), never their occurrences; instances hang off each definition",
  emptyNote: "0 access-review definitions matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Definitions are review schedules (a series): each recurrence creates one instance per reviewed resource, and a one-time review creates one instance per resource",
    `Show one schedule: mg-axi entra access-review definition show --id <definition-id> ${profileHint(profileName)}`,
  ],
};

const INSTANCE_LIST: CollectionShape = {
  command: "entra access-review instance list",
  key: "instances",
  known: INSTANCE_KNOWN,
  knownList: KNOWN_INSTANCE_FIELDS,
  defaultSelect: DEFAULT_INSTANCE_LIST_SELECT,
  denialScopeNote: "Instances are occurrences of one definition schedule, never schedules themselves",
  emptyNote: "0 access-review instances matched; the absence of results is the answer, not an error",
  standing: () => [
    "Instances are occurrences of one definition schedule: each recurrence and each reviewed resource gets its own instance",
  ],
};

const DECISION_LIST: CollectionShape = {
  command: "entra access-review decision list",
  key: "decisions",
  known: DECISION_KNOWN,
  knownList: KNOWN_DECISION_FIELDS,
  defaultSelect: DEFAULT_DECISION_SELECT,
  denialScopeNote: "Decisions are read-only here: listing never approves, denies or applies anything",
  emptyNote: "0 access-review decisions matched; an instance with no reviewed principals has no decision items, which is the answer, not an error",
  standing: () => [
    "Decisions are read-only here: listing never approves, denies or applies anything; submitting decisions belongs to a later slice",
  ],
};

const CONTACTED_REVIEWER_LIST: CollectionShape = {
  command: "entra access-review contacted-reviewer list",
  key: "contactedReviewers",
  known: CONTACTED_REVIEWER_KNOWN,
  knownList: KNOWN_CONTACTED_REVIEWER_FIELDS,
  defaultSelect: DEFAULT_CONTACTED_REVIEWER_LIST_SELECT,
  denialScopeNote: "Contacted reviewers are the reviewer identities recorded on one instance, whether or not they were notified; they never carry review outcomes",
  emptyNote: "0 contacted reviewers matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Contacted reviewers are identities recorded on one instance, never review outcomes or notification receipts",
    `Show one contacted reviewer: mg-axi entra access-review contacted-reviewer show --definition <definition-id> --instance <instance-id> --id <reviewer-id> ${profileHint(profileName)}`,
  ],
};

const STAGE_LIST: CollectionShape = {
  command: "entra access-review stage list",
  key: "stages",
  known: STAGE_KNOWN,
  knownList: KNOWN_STAGE_FIELDS,
  defaultSelect: DEFAULT_STAGE_LIST_SELECT,
  denialScopeNote: "Stages are sequential phases of one instance (up to three when stageSettings is defined), never definitions or decisions",
  emptyNote: "0 access-review stages matched; an instance without stageSettings has no stages, which is the answer, not an error",
  standing: profileName => [
    "Stages are sequential phases of one instance: each stage carries its own reviewer scopes, and per-stage decisions belong to a later slice",
    `Show one stage: mg-axi entra access-review stage show --definition <definition-id> --instance <instance-id> --id <stage-id> ${profileHint(profileName)}`,
  ],
};

function parentId(flags: AccessReviewFlags, flag: string, noun: string, help: string): string {
  const value = String(flags[flag] ?? "");
  if (!value.trim()) throw new AxiError(`--${flag} needs the parent ${noun} ID`, "VALIDATION_ERROR", [help]);
  return value;
}

async function listCollection(
  shape: CollectionShape,
  session: GraphSession,
  flags: AccessReviewFlags,
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
  const scopes = scopesFor(flags, profile, DEFAULT_ACCESS_REVIEW_SCOPES, help);
  const full = flags.full === true;
  // These collections document $select/$filter only (plus server-side
  // $orderby/$skip/$top paging concerns), so --filter passes through as
  // plain $filter with no $count or ConsistencyLevel.
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
  const effectiveFlags: AccessReviewFlags = { ...flags, select: result.query.$select ?? shape.defaultSelect.join(",") };
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
  const standing = [shape.denialScopeNote, ...shape.standing(profileName)];
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
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  params: Record<string, string>,
  malformed: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, defaultSelect, known, knownList, help);
  const scopes = scopesFor(flags, profile, DEFAULT_ACCESS_REVIEW_SCOPES, help);
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

export async function listDefinitions(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(DEFINITION_LIST, session, flags, profile, operation, help, profileName, {},
    `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showDefinition(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("definition", "entra access-review definition show",
    DEFINITION_KNOWN, KNOWN_DEFINITION_FIELDS, DEFAULT_DEFINITION_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessReviewScheduleDefinition-id": String(flags.id) },
    "Graph returned a malformed access-review definition body");
}

export async function listInstances(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const definition = parentId(flags, "definition", "access-review definition", help);
  return listCollection(INSTANCE_LIST, session, flags, profile, operation, help, profileName,
    { "accessReviewScheduleDefinition-id": definition },
    `Resume losslessly with the same --definition and flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showInstance(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const definition = parentId(flags, "definition", "access-review definition", help);
  return showOne("instance", "entra access-review instance show",
    INSTANCE_KNOWN, KNOWN_INSTANCE_FIELDS, DEFAULT_INSTANCE_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessReviewScheduleDefinition-id": definition, "accessReviewInstance-id": String(flags.id) },
    "Graph returned a malformed access-review instance body");
}

export async function listDecisions(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const definition = parentId(flags, "definition", "access-review definition", help);
  const instance = parentId(flags, "instance", "access-review instance", help);
  return listCollection(DECISION_LIST, session, flags, profile, operation, help, profileName,
    { "accessReviewScheduleDefinition-id": definition, "accessReviewInstance-id": instance },
    `Resume losslessly with the same --definition, --instance and flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showDecision(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const definition = parentId(flags, "definition", "access-review definition", help);
  const instance = parentId(flags, "instance", "access-review instance", help);
  return showOne("decision", "entra access-review decision show",
    DECISION_KNOWN, KNOWN_DECISION_FIELDS, DEFAULT_DECISION_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessReviewScheduleDefinition-id": definition, "accessReviewInstance-id": instance, "accessReviewInstanceDecisionItem-id": String(flags.id) },
    "Graph returned a malformed access-review decision body");
}

export async function listContactedReviewers(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const definition = parentId(flags, "definition", "access-review definition", help);
  const instance = parentId(flags, "instance", "access-review instance", help);
  return listCollection(CONTACTED_REVIEWER_LIST, session, flags, profile, operation, help, profileName,
    { "accessReviewScheduleDefinition-id": definition, "accessReviewInstance-id": instance },
    `Resume losslessly with the same --definition, --instance and flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showContactedReviewer(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const definition = parentId(flags, "definition", "access-review definition", help);
  const instance = parentId(flags, "instance", "access-review instance", help);
  return showOne("contactedReviewer", "entra access-review contacted-reviewer show",
    CONTACTED_REVIEWER_KNOWN, KNOWN_CONTACTED_REVIEWER_FIELDS, DEFAULT_CONTACTED_REVIEWER_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessReviewScheduleDefinition-id": definition, "accessReviewInstance-id": instance, "accessReviewReviewer-id": String(flags.id) },
    "Graph returned a malformed contacted-reviewer body");
}

export async function listStages(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const definition = parentId(flags, "definition", "access-review definition", help);
  const instance = parentId(flags, "instance", "access-review instance", help);
  return listCollection(STAGE_LIST, session, flags, profile, operation, help, profileName,
    { "accessReviewScheduleDefinition-id": definition, "accessReviewInstance-id": instance },
    `Resume losslessly with the same --definition, --instance and flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`);
}

export async function showStage(
  session: GraphSession,
  flags: AccessReviewFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const definition = parentId(flags, "definition", "access-review definition", help);
  const instance = parentId(flags, "instance", "access-review instance", help);
  return showOne("stage", "entra access-review stage show",
    STAGE_KNOWN, KNOWN_STAGE_FIELDS, DEFAULT_STAGE_SHOW_SELECT,
    session, flags, profile, operation, help, profileName,
    { "accessReviewScheduleDefinition-id": definition, "accessReviewInstance-id": instance, "accessReviewStage-id": String(flags.id) },
    "Graph returned a malformed access-review stage body");
}
