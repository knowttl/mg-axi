import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 group lifecycle subfamily: the read mapping behind
// `mg-axi entra group-lifecycle-policy list/show/count` and
// `mg-axi entra group-setting-template list/show/count`. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 grouplifecyclepolicy-list,
// grouplifecyclepolicy-get, groupsettingtemplate-list,
// groupsettingtemplate-get and groupsettingtemplate resource documentation
// on 2026-10-05. Lifecycle-policy reads take D/A Directory.Read.All; the
// operation pages state no delegated role prerequisite and no P1/P2
// prerequisite. Setting-template reads take D/A GroupSettings.Read.All;
// delegated callers additionally need a supported Entra role (Directory
// Readers or Global Reader are the least-privileged roles). Personal
// Microsoft accounts are not supported on either family, and no P1/P2
// prerequisite is stated for either. Lifecycle-policy lists document the
// general OData query parameters, so --filter passes through as plain
// $filter with no $count or ConsistencyLevel contract. Template lists
// document $select only ($filter is not supported), so the template list
// offers no --filter. The $count routes carry no operation-level
// documentation page and take no --filter/--select/--limit/--cursor.
// groupSettingTemplates/delta() stays out: delta-token sync needs its own
// paging and change-tracking contract. Beta policies and templates and the
// POST lookup actions (getByIds, validateProperties, membership checks)
// belong to later EXT-01 subfamilies. No lifecycle or template mutation
// exists in this slice.

// Every lifecycle-policy property this slice may request or display,
// matching the reviewed resource. Anything else fails before credentials.
export const KNOWN_LIFECYCLE_FIELDS: readonly string[] = [
  "id",
  "alternateNotificationEmails",
  "groupLifetimeInDays",
  "managedGroupTypes",
];
const KNOWN_LIFECYCLE = new Set(KNOWN_LIFECYCLE_FIELDS);

// Every setting-template property this slice may request or display,
// matching the reviewed resource. Anything else fails before credentials.
export const KNOWN_TEMPLATE_FIELDS: readonly string[] = [
  "id",
  "deletedDateTime",
  "displayName",
  "description",
  "values",
];
const KNOWN_TEMPLATES = new Set(KNOWN_TEMPLATE_FIELDS);

// Compact lifecycle rows: the policy id, the expiry window and the group
// types the policy governs.
const DEFAULT_LIFECYCLE_LIST_SELECT = ["id", "groupLifetimeInDays", "managedGroupTypes"];
// Show rows: the full reviewed lifecycle set.
const DEFAULT_LIFECYCLE_SHOW_SELECT = [...KNOWN_LIFECYCLE_FIELDS];
// Compact template rows: the template id and the display name/description
// that decide which template a setting is built from.
const DEFAULT_TEMPLATE_LIST_SELECT = ["id", "displayName", "description"];
// Show rows: the full reviewed template set, including the values
// collection carrying setting names, types and defaults.
const DEFAULT_TEMPLATE_SHOW_SELECT = [...KNOWN_TEMPLATE_FIELDS];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_LIFECYCLE_SCOPES = ["https://graph.microsoft.com/Directory.Read.All"];
export const DEFAULT_TEMPLATE_SCOPES = ["https://graph.microsoft.com/GroupSettings.Read.All"];
const TRUNCATE_AT = 500;

export type GroupLifecycleFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown group lifecycle property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: GroupLifecycleFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
// Template `values` rows are objects, so they pass through untruncated;
// long descriptions still truncate with a --full escape hatch.
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
  flags: GroupLifecycleFlags,
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

function fullHint(command: string, flags: GroupLifecycleFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the scope, roles and licensing that actually unlock it,
// because a 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const LIFECYCLE_DENIAL_HINTS = [
  "Lifecycle-policy reads need Directory.Read.All for delegated or application access; delegated callers pass it as --scopes",
  "Personal Microsoft accounts are not supported for lifecycle-policy reads",
  "No delegated role or P1/P2 prerequisite is stated for lifecycle-policy reads; never diagnose role or licence solely from HTTP 403",
];

const TEMPLATE_DENIAL_HINTS = [
  "Setting-template reads need GroupSettings.Read.All plus a supported Entra role for delegated access (Directory Readers or Global Reader are the least-privileged roles), or admin-consented GroupSettings.Read.All for application access",
  "Personal Microsoft accounts are not supported for setting-template reads",
  "No P1/P2 prerequisite is stated for setting-template reads; never diagnose licence solely from HTTP 403",
];

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
  filter: string | undefined;
}

// Restoring the saved select/filter keeps cursor resumes lossless when
// --select or --filter is omitted. The lifecycle list passes --filter
// through as plain $filter with no $count or ConsistencyLevel contract;
// the template list documents $select only and offers no --filter flag, so
// strict input validation refuses the flag before credentials instead of
// forwarding a misleading request.
function collectionCommon(
  session: GraphSession,
  flags: GroupLifecycleFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  defaultScopes: string[],
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
  withFilter: boolean,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    known, knownList,
    saved?.["$select"] === undefined ? defaults : fieldList(saved["$select"], known, knownList, "select", help), help);
  const filter = withFilter
    ? (flags.filter === undefined ? saved?.["$filter"] : String(flags.filter))
    : undefined;
  return { cursor, select, fields, scopes: scopesFor(flags, defaultScopes, profile, help), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: GroupLifecycleFlags,
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

function policyId(flags: GroupLifecycleFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the lifecycle policy identifier", "VALIDATION_ERROR", [help]);
  return id;
}

function templateId(flags: GroupLifecycleFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the setting template identifier", "VALIDATION_ERROR", [help]);
  return id;
}

function singleResult(
  raw: unknown,
  fields: string[],
  full: boolean,
  kind: string,
): { row: Record<string, unknown>; truncated: boolean } {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError(`Graph returned a malformed ${kind} body`, "GRAPH_ERROR", [
      `Single-${kind} reads carry one ${kind} object; treat anything else as unknown, not empty`,
    ]);
  }
  return project(raw, fields, full);
}

export async function listLifecyclePolicies(
  session: GraphSession,
  flags: GroupLifecycleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_LIFECYCLE, KNOWN_LIFECYCLE_FIELDS,
    DEFAULT_LIFECYCLE_LIST_SELECT, DEFAULT_LIFECYCLE_SCOPES, operation, help, profile, true);
  const result = await withGuidance(LIFECYCLE_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: GroupLifecycleFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIFECYCLE_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const policies: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    policies.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra group-lifecycle-policy show --id <policy-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra group-lifecycle-policy list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      groupLifecyclePolicies: policies,
      ...listTotals(policies.length, result.total, "lifecycle policies", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  if (!policies.length) {
    return {
      groupLifecyclePolicies: policies,
      ...listTotals(policies.length, result.total, "lifecycle policies", true),
      complete: true,
      help: ["0 lifecycle policies matched; the absence of results is the answer, not an error"],
    };
  }
  return { groupLifecyclePolicies: policies, ...listTotals(policies.length, result.total, "lifecycle policies", true), complete: true, help: [...truncationHints, showHint] };
}

export async function showLifecyclePolicy(
  session: GraphSession,
  flags: GroupLifecycleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_LIFECYCLE, KNOWN_LIFECYCLE_FIELDS, DEFAULT_LIFECYCLE_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_LIFECYCLE_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(LIFECYCLE_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "groupLifecyclePolicy-id": policyId(flags, help) },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "lifecycle policy");
  if (truncated) return { groupLifecyclePolicy: row, help: [fullHint("entra group-lifecycle-policy show", flags, profileName)] };
  return { groupLifecyclePolicy: row };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute and accepts only
// a non-negative integer. There is no --filter/--select/--limit contract on
// the count: the catalogue declares no such flags and strict input
// validation refuses them before credentials.
export async function countLifecyclePolicies(
  session: GraphSession,
  flags: GroupLifecycleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_LIFECYCLE_SCOPES, profile, help);
  const raw = await withGuidance(LIFECYCLE_DENIAL_HINTS, () => session.execute({ profile, operation, scopes }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed lifecycle-policy count body", "GRAPH_ERROR", [
      "Lifecycle-policy counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  return { groupLifecyclePolicyCount: raw };
}

export async function listSettingTemplates(
  session: GraphSession,
  flags: GroupLifecycleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_TEMPLATES, KNOWN_TEMPLATE_FIELDS,
    DEFAULT_TEMPLATE_LIST_SELECT, DEFAULT_TEMPLATE_SCOPES, operation, help, profile, false);
  const result = await withGuidance(TEMPLATE_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: GroupLifecycleFlags = { ...flags, select: result.query.$select ?? DEFAULT_TEMPLATE_LIST_SELECT.join(",") };
  const templates: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    templates.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra group-setting-template show --id <template-id> ${profileHint(profileName)}`;
  const filterNote = "Template lists offer no --filter: Graph documents $select only for /groupSettingTemplates";
  const truncationHints = truncated ? [fullHint("entra group-setting-template list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      groupSettingTemplates: templates,
      ...listTotals(templates.length, result.total, "setting templates", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, filterNote],
    };
  }
  if (!templates.length) {
    return {
      groupSettingTemplates: templates,
      ...listTotals(templates.length, result.total, "setting templates", true),
      complete: true,
      help: [filterNote, "0 setting templates matched; the absence of results is the answer, not an error"],
    };
  }
  return { groupSettingTemplates: templates, ...listTotals(templates.length, result.total, "setting templates", true), complete: true, help: [...truncationHints, showHint, filterNote] };
}

export async function showSettingTemplate(
  session: GraphSession,
  flags: GroupLifecycleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_TEMPLATES, KNOWN_TEMPLATE_FIELDS, DEFAULT_TEMPLATE_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_TEMPLATE_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(TEMPLATE_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "groupSettingTemplate-id": templateId(flags, help) },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "setting template");
  if (truncated) return { groupSettingTemplate: row, help: [fullHint("entra group-setting-template show", flags, profileName)] };
  return { groupSettingTemplate: row };
}

export async function countSettingTemplates(
  session: GraphSession,
  flags: GroupLifecycleFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_TEMPLATE_SCOPES, profile, help);
  const raw = await withGuidance(TEMPLATE_DENIAL_HINTS, () => session.execute({ profile, operation, scopes }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed setting-template count body", "GRAPH_ERROR", [
      "Setting-template counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  return { groupSettingTemplateCount: raw };
}
