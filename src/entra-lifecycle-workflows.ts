import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-02 lifecycle-workflows reads: the read-only workflow, workflow
// template, task definition and tenant-settings first part. See README.md
// for the supported commands and usage. Operation construction stays beside
// its command; the shared session owns URLs, credentials, paging, retries
// and error translation, and the SDK owns TOON rendering. This module only
// maps flags to session calls and projects rows for compact output.
//
// Reviewed against the v1.0 lifecycleworkflowscontainer-list-workflows,
// workflow-get, lifecycleworkflowscontainer-list-workflowtemplates,
// workflowtemplate-get, lifecycleworkflowscontainer-list-taskdefinitions,
// taskdefinition-get and lifecyclemanagementsettings-get operation
// documentation and the pinned v1.0 metadata property sets for workflow,
// workflowTemplate, taskDefinition and lifecycleManagementSettings, on
// 2026-10-05. Workflow list/show take D/A
// LifecycleWorkflows-Workflow.ReadBasic.All as the least privileged choice
// (LifecycleWorkflows-Workflow.Read.All or LifecycleWorkflows.Read.All for
// richer detail); template, task-definition and settings reads take D/A
// LifecycleWorkflows.Read.All. Delegated personal Microsoft accounts are
// not supported. Delegated callers additionally need Global Reader or
// Lifecycle Workflows Administrator. The three lists document $filter (and
// $select), so --filter passes through as plain $filter with no $count or
// ConsistencyLevel contract ($search/$orderby/$expand stay unreviewed and
// $top/$skip stay server-side paging concerns); workflow, task-definition
// and settings singles document $select only. The template get documents no
// query parameters, so the template show projects whole rows with no
// --select. Navigation and execution detail (workflow tasks, createdBy,
// lastModifiedBy, previewScope, executionScope, runs, user/subject/task
// processing results, versions, insights, deleted items, custom task
// extensions) needs $expand or its own sub-read and stays out: runs and
// processing results belong to a later part, never to these reads.
// Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra
// Suite (every governed user, not only administrators). No mutation (no
// workflow create/update/delete/activate, no settings update) and no beta.

// Workflows expose the configured joiner/mover/leaver automations. Only
// scalar properties carry a reviewed $select contract; tasks, runs,
// processing results, versions, scopes and createdBy/lastModifiedBy links
// are navigation properties and stay out (tasks ride expanded by default on
// the get and are dropped in local projection).
export const KNOWN_WORKFLOW_FIELDS: readonly string[] = [
  "id",
  "category",
  "createdDateTime",
  "deletedDateTime",
  "description",
  "displayName",
  "executionConditions",
  "isEnabled",
  "isSchedulingEnabled",
  "lastModifiedDateTime",
  "nextScheduleRunDateTime",
  "version",
  "quarantineDetails",
  "settings",
];
// Workflow templates expose the pre-configured Entra scenarios new
// workflows are built from. The tasks relationship is expanded by default
// on the wire and is dropped here: template tasks belong to a later slice.
export const KNOWN_WORKFLOW_TEMPLATE_FIELDS: readonly string[] = [
  "id",
  "category",
  "description",
  "displayName",
  "executionConditions",
];
// Task definitions expose the built-in tasks workflows are constructed
// from. They carry no relationships.
export const KNOWN_TASK_DEFINITION_FIELDS: readonly string[] = [
  "id",
  "category",
  "continueOnError",
  "description",
  "displayName",
  "parameters",
  "version",
];
// Tenant settings expose the workflow schedule interval and the email and
// quarantine configuration. They carry no relationships; the update action
// belongs to no read slice.
export const KNOWN_LIFECYCLE_SETTINGS_FIELDS: readonly string[] = [
  "workflowScheduleIntervalInHours",
  "emailSettings",
  "quarantineConfiguration",
];
const WORKFLOW_KNOWN = new Set(KNOWN_WORKFLOW_FIELDS);
const WORKFLOW_TEMPLATE_KNOWN = new Set(KNOWN_WORKFLOW_TEMPLATE_FIELDS);
const TASK_DEFINITION_KNOWN = new Set(KNOWN_TASK_DEFINITION_FIELDS);
const LIFECYCLE_SETTINGS_KNOWN = new Set(KNOWN_LIFECYCLE_SETTINGS_FIELDS);

// Compact rows: identifiers plus category and enablement state.
const DEFAULT_WORKFLOW_LIST_SELECT = ["id", "displayName", "category", "isEnabled", "isSchedulingEnabled"];
const DEFAULT_WORKFLOW_SHOW_SELECT = [...KNOWN_WORKFLOW_FIELDS];
const DEFAULT_WORKFLOW_TEMPLATE_LIST_SELECT = ["id", "displayName", "category"];
const DEFAULT_WORKFLOW_TEMPLATE_SHOW_SELECT = [...KNOWN_WORKFLOW_TEMPLATE_FIELDS];
const DEFAULT_TASK_DEFINITION_LIST_SELECT = ["id", "displayName", "category", "version"];
const DEFAULT_TASK_DEFINITION_SHOW_SELECT = [...KNOWN_TASK_DEFINITION_FIELDS];
const DEFAULT_LIFECYCLE_SETTINGS_SELECT = [...KNOWN_LIFECYCLE_SETTINGS_FIELDS];
export const DEFAULT_WORKFLOW_SCOPES = ["https://graph.microsoft.com/LifecycleWorkflows-Workflow.ReadBasic.All"];
export const DEFAULT_LIFECYCLE_SCOPES = ["https://graph.microsoft.com/LifecycleWorkflows.Read.All"];
const TRUNCATE_AT = 500;

export type LifecycleWorkflowsFlags = Record<string, string | boolean>;

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

function scopesFor(flags: LifecycleWorkflowsFlags, profile: AnyProfile, defaults: readonly string[], help: string): string[] | undefined {
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
  flags: LifecycleWorkflowsFlags,
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

function fullHint(command: string, flags: LifecycleWorkflowsFlags, profileName: string): string {
  const args = Object.entries({ ...flags, ...(flags.cursor === undefined ? {} : { cursor: "-" }), profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause;
// lifecycle workflows adds the scope, roles and licensing that actually
// unlock these reads, because a 403 alone never says which prerequisite is
// missing.
function withGuidance<T>(run: () => Promise<T>, hints: readonly string[] = LIFECYCLE_DENIAL_HINTS): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const LIFECYCLE_DENIAL_HINTS = [
  "Lifecycle-workflow reads need LifecycleWorkflows.Read.All plus Global Reader or Lifecycle Workflows Administrator for delegated access, or admin-consented LifecycleWorkflows.Read.All for application access; delegated personal Microsoft accounts are not supported",
  "Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite (every governed user, not only administrators)",
];

// Workflow reads document their own least-privileged scope, so denials name
// ReadBasic first with the richer read alternatives beside it.
const WORKFLOW_DENIAL_HINTS = [
  "Workflow reads need LifecycleWorkflows-Workflow.ReadBasic.All (least privileged; LifecycleWorkflows-Workflow.Read.All or LifecycleWorkflows.Read.All for richer detail) plus Global Reader or Lifecycle Workflows Administrator for delegated access, or admin-consented LifecycleWorkflows-Workflow.ReadBasic.All for application access; delegated personal Microsoft accounts are not supported",
  "Lifecycle workflows need Microsoft Entra ID Governance or Microsoft Entra Suite (every governed user, not only administrators)",
];

interface CollectionShape {
  command: string;
  key: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  defaultScopes: readonly string[];
  hints: readonly string[];
  emptyNote: string;
  standing: (profileName: string) => string[];
}

const WORKFLOW_LIST: CollectionShape = {
  command: "entra lifecycle workflow list",
  key: "workflows",
  known: WORKFLOW_KNOWN,
  knownList: KNOWN_WORKFLOW_FIELDS,
  defaultSelect: DEFAULT_WORKFLOW_LIST_SELECT,
  defaultScopes: DEFAULT_WORKFLOW_SCOPES,
  hints: WORKFLOW_DENIAL_HINTS,
  emptyNote: "0 workflows matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Workflows are the configured joiner/mover/leaver automations: a workflow row never carries its tasks, runs or processing results; those belong to later slices",
    `Show one workflow: mg-axi entra lifecycle workflow show --id <workflow-id> ${profileHint(profileName)}`,
  ],
};

const WORKFLOW_TEMPLATE_LIST: CollectionShape = {
  command: "entra lifecycle workflow-template list",
  key: "workflowTemplates",
  known: WORKFLOW_TEMPLATE_KNOWN,
  knownList: KNOWN_WORKFLOW_TEMPLATE_FIELDS,
  defaultSelect: DEFAULT_WORKFLOW_TEMPLATE_LIST_SELECT,
  defaultScopes: DEFAULT_LIFECYCLE_SCOPES,
  hints: LIFECYCLE_DENIAL_HINTS,
  emptyNote: "0 workflow templates matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Workflow templates are the pre-configured Entra scenarios new workflows are built from; template tasks belong to a later slice",
    `Show one workflow template: mg-axi entra lifecycle workflow-template show --id <template-id> ${profileHint(profileName)}`,
  ],
};

const TASK_DEFINITION_LIST: CollectionShape = {
  command: "entra lifecycle task-definition list",
  key: "taskDefinitions",
  known: TASK_DEFINITION_KNOWN,
  knownList: KNOWN_TASK_DEFINITION_FIELDS,
  defaultSelect: DEFAULT_TASK_DEFINITION_LIST_SELECT,
  defaultScopes: DEFAULT_LIFECYCLE_SCOPES,
  hints: LIFECYCLE_DENIAL_HINTS,
  emptyNote: "0 task definitions matched; the absence of results is the answer, not an error",
  standing: profileName => [
    "Task definitions are the built-in tasks workflows are constructed from; they carry no relationships",
    `Show one task definition: mg-axi entra lifecycle task-definition show --id <task-definition-id> ${profileHint(profileName)}`,
  ],
};

async function listCollection(
  shape: CollectionShape,
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
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
  const savedFilter = saved?.$filter;
  const filter = flags.filter === undefined ? savedFilter : String(flags.filter);
  const scopes = scopesFor(flags, profile, shape.defaultScopes, help);
  const full = flags.full === true;
  // These lists document $filter (and $select), so --filter passes through
  // as plain $filter with no $count or ConsistencyLevel contract;
  // $search/$orderby/$expand stay unreviewed and $top/$skip stay
  // server-side paging concerns. Resumed queries replay the stored $select
  // verbatim so resume never conflicts.
  const sendSelect = cursor === undefined
    ? select
    : saved?.$select === undefined ? [] : fieldList(saved.$select, "select", shape.known, shape.knownList, help);
  const query: Record<string, string> = {};
  if (sendSelect.length) query.$select = sendSelect.join(",");
  if (filter !== undefined) query.$filter = filter;
  const args: CollectArgs = { profile, operation, params: {}, query, scopes };
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(() => session.collect(args), shape.hints);
  const effectiveFlags: LifecycleWorkflowsFlags = { ...flags, ...(result.query.$select === undefined ? {} : { select: select.join(",") }) };
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
  const resumeHint = `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`;
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
  defaultScopes: readonly string[],
  hints: readonly string[],
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  params: Record<string, string>,
  malformed: string,
  queryless: boolean = false,
): Promise<Record<string, unknown>> {
  if (queryless && flags.select !== undefined) {
    throw new AxiError("Graph documents no query parameters on this read; whole rows are always returned", "VALIDATION_ERROR", [help]);
  }
  const { select, fields } = queryless
    ? { select: [...defaultSelect], fields: flags.fields === undefined ? [...defaultSelect] : fieldList(flags.fields, "fields", known, knownList, help) }
    : selectedFields(flags, defaultSelect, known, knownList, help);
  const scopes = scopesFor(flags, profile, defaultScopes, help);
  const full = flags.full === true;
  // Singles document $select only (the template get documents no query
  // parameters at all and arrives whole); $expand stays out everywhere:
  // tasks, createdBy, lastModifiedBy and previewScope belong to later
  // slices, never to these reads.
  const query: Record<string, string> = {};
  if (!queryless && select.length) query.$select = select.join(",");
  const raw = await withGuidance(() => session.execute({
    profile,
    operation,
    params,
    query,
    scopes,
  }), hints);
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
  defaultScopes: readonly string[],
  hints: readonly string[],
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  _profileName: string,
): Promise<Record<string, unknown>> {
  const scopes = scopesFor(flags, profile, defaultScopes, help);
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

export async function listWorkflows(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(WORKFLOW_LIST, session, flags, profile, operation, help, profileName);
}

export async function showWorkflow(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("workflow", "entra lifecycle workflow show",
    WORKFLOW_KNOWN, KNOWN_WORKFLOW_FIELDS, DEFAULT_WORKFLOW_SHOW_SELECT,
    DEFAULT_WORKFLOW_SCOPES, WORKFLOW_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName,
    { "workflow-id": String(flags.id) },
    "Graph returned a malformed lifecycle workflow body");
}

export async function countWorkflows(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return countCollection("workflow", WORKFLOW_LIST.emptyNote, DEFAULT_WORKFLOW_SCOPES, WORKFLOW_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName);
}

export async function listWorkflowTemplates(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(WORKFLOW_TEMPLATE_LIST, session, flags, profile, operation, help, profileName);
}

export async function showWorkflowTemplate(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("workflowTemplate", "entra lifecycle workflow-template show",
    WORKFLOW_TEMPLATE_KNOWN, KNOWN_WORKFLOW_TEMPLATE_FIELDS, DEFAULT_WORKFLOW_TEMPLATE_SHOW_SELECT,
    DEFAULT_LIFECYCLE_SCOPES, LIFECYCLE_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName,
    { "workflowTemplate-id": String(flags.id) },
    "Graph returned a malformed workflow template body",
    true);
}

export async function countWorkflowTemplates(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return countCollection("workflow template", WORKFLOW_TEMPLATE_LIST.emptyNote, DEFAULT_LIFECYCLE_SCOPES, LIFECYCLE_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName);
}

export async function listTaskDefinitions(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(TASK_DEFINITION_LIST, session, flags, profile, operation, help, profileName);
}

export async function showTaskDefinition(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("taskDefinition", "entra lifecycle task-definition show",
    TASK_DEFINITION_KNOWN, KNOWN_TASK_DEFINITION_FIELDS, DEFAULT_TASK_DEFINITION_SHOW_SELECT,
    DEFAULT_LIFECYCLE_SCOPES, LIFECYCLE_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName,
    { "taskDefinition-id": String(flags.id) },
    "Graph returned a malformed task definition body");
}

export async function countTaskDefinitions(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return countCollection("task definition", TASK_DEFINITION_LIST.emptyNote, DEFAULT_LIFECYCLE_SCOPES, LIFECYCLE_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName);
}

export async function showLifecycleSettings(
  session: GraphSession,
  flags: LifecycleWorkflowsFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showOne("settings", "entra lifecycle settings show",
    LIFECYCLE_SETTINGS_KNOWN, KNOWN_LIFECYCLE_SETTINGS_FIELDS, DEFAULT_LIFECYCLE_SETTINGS_SELECT,
    DEFAULT_LIFECYCLE_SCOPES, LIFECYCLE_DENIAL_HINTS,
    session, flags, profile, operation, help, profileName,
    {},
    "Graph returned a malformed lifecycle settings body");
}
