import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// READ-02: the Entra group read mapping behind `mg-axi entra group list/show`
// and `mg-axi entra group member list` / `member-of list` (each with an
// opt-in `--transitive` flat view). Operation construction stays beside its
// command; the shared session owns URLs, credentials, paging, retries and
// error translation, and the SDK owns TOON rendering. This module only maps
// flags to session calls and projects rows for compact output.
//
// Reviewed against the v1.0 group-list, group-get, group-list-members,
// group-list-transitivemembers, group-list-memberof and
// group-list-transitivememberof operation documentation on 2026-10-04.
// Group list/get upstream tables currently mark the write permission
// Group-NestingSupport.ReadWrite.All least-privileged; the review records the
// supported read alternatives instead (GroupMember.Read.All for basics,
// Group.Read.All for further properties). Member and memberOf reads take
// GroupMember.ReadBasic.All minimum and GroupMember.Read.All for richer
// access; hidden-membership groups additionally need Member.Read.Hidden and,
// for delegated callers, a supported role with hidden-member read permission.
// The server omits hidden members it will not disclose rather than failing,
// so completion describes pagination, never visibility. Application callers
// with narrow consent receive limited-information rows carrying only
// @odata.type and id with other properties null; those rows are preserved,
// never reinterpreted as empty. The v1.0 members route has a known issue
// omitting service principals; the documented workarounds are beta or
// $expand, and neither is used silently here - the limitation rides along as
// an explicit warning instead. Role-assignable groups surface through
// isAssignableToRole; changing their membership needs role-management
// permission and belongs to WRITE-01, never to this slice.

// Every group property this slice may request or display, matching the
// reviewed raw surface. Anything else fails before credentials.
export const KNOWN_GROUP_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "description",
  "mail",
  "mailEnabled",
  "mailNickname",
  "securityEnabled",
  "groupTypes",
  "visibility",
  "classification",
  "isAssignableToRole",
  "createdDateTime",
  "expirationDateTime",
  "renewedDateTime",
  "membershipRule",
  "membershipRuleProcessingState",
];
const KNOWN_GROUPS = new Set(KNOWN_GROUP_FIELDS);

// Relationship rows are directoryObjects of mixed types; richer per-type
// member fields need single-object reads, so only these stay selectable.
export const KNOWN_MEMBER_FIELDS: readonly string[] = ["id", "displayName", "mail"];
const KNOWN_MEMBERS = new Set(KNOWN_MEMBER_FIELDS);
// @odata.type is preserved on relationship rows without being selectable: it
// names the member kind and marks limited-information rows.
const MEMBER_TYPE_PROPERTY = "@odata.type";

// Compact group rows: identifier, display name, address and kind.
const DEFAULT_LIST_SELECT = ["id", "displayName", "mail", "groupTypes"];
// Show rows: the full reviewed group set, including the role-assignable flag.
const DEFAULT_SHOW_SELECT = [...KNOWN_GROUP_FIELDS];
// Compact relationship rows: identifier and display name.
const DEFAULT_MEMBER_SELECT = ["id", "displayName"];
// GroupMember.Read.All covers group basics and member reads in both modes;
// hidden members need Member.Read.Hidden and richer group properties may
// need Group.Read.All, both passed explicitly via --scopes.
export const DEFAULT_DELEGATED_SCOPES = ["https://graph.microsoft.com/GroupMember.Read.All"];
const TRUNCATE_AT = 500;

// Direct relationship operation to its transitive flat-view counterpart.
// --transitive selects between these two catalogued routes; beta and $expand
// workarounds are never selected implicitly.
export const TRANSITIVE_OPERATION: Readonly<Record<string, string>> = {
  "GET:/groups/{group-id}/members": "GET:/groups/{group-id}/transitiveMembers",
  "GET:/groups/{group-id}/memberOf": "GET:/groups/{group-id}/transitiveMemberOf",
};

// Sourced v1.0 limitation: service principals may be absent from direct
// member results even when pagination completes.
export const SERVICE_PRINCIPAL_WARNING =
  "Microsoft Graph v1.0 may omit service principals from group members; completed pagination does not establish complete membership.";

export type GroupFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown group property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known group properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: GroupFlags, profile: AnyProfile, help: string): string[] | undefined {
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
  if (flags.scopes === undefined) return [...DEFAULT_DELEGATED_SCOPES];
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

// Relationship projection additionally preserves @odata.type so the member
// kind and limited-information rows survive local projection.
function projectMember(
  row: unknown,
  select: string[],
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean; limitedInfo: boolean } {
  const { row: projected, truncated } = project(row, fields, full);
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  if (typeof source[MEMBER_TYPE_PROPERTY] === "string") projected[MEMBER_TYPE_PROPERTY] = source[MEMBER_TYPE_PROPERTY];
  const descriptive = [...new Set([...select, ...Object.keys(source)])]
    .filter(field => field !== "id" && !field.startsWith("@"));
  const limitedInfo = Object.hasOwn(source, "id") && descriptive.length > 0
    && descriptive.every(field => source[field] === null || source[field] === undefined || !Object.hasOwn(source, field));
  return { row: projected, truncated, limitedInfo };
}

function selectedFields(
  flags: GroupFlags,
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

function fullHint(command: string, flags: GroupFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
  filter: string | undefined;
}

// Relationship $filter runs through advanced query on every group route:
// $filter on members/memberOf/transitive collections requires
// ConsistencyLevel eventual, and group list tolerates it, so an effective
// filter always carries eventual consistency. Restoring the saved filter
// keeps cursor resumes lossless when --filter is omitted.
function collectionCommon(
  session: GraphSession,
  flags: GroupFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    known, knownList,
    saved?.["$select"] === undefined ? defaults : fieldList(saved["$select"], known, knownList, "select", help), help);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, scopes: scopesFor(flags, profile, help), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: GroupFlags,
  help: string,
): CollectArgs {
  const query: Record<string, string> = { $select: common.select.join(",") };
  if (common.filter !== undefined) {
    query.$filter = common.filter;
    query.$count = "true";
  }
  const args: CollectArgs = { profile, operation, query, scopes: common.scopes };
  if (common.filter !== undefined) args.consistencyLevel = "eventual";
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

export async function listGroups(
  session: GraphSession,
  flags: GroupFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_GROUPS, KNOWN_GROUP_FIELDS, DEFAULT_LIST_SELECT, operation, help, profile);
  const result = await session.collect(collectArgs(profile, operation, common, flags, help));
  const effectiveFlags: GroupFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const groups: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    groups.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra group show --id <group-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra group list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      groups,
      count: { returned: groups.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  const count = { returned: groups.length, complete: true };
  if (!groups.length) {
    return {
      groups,
      count,
      help: [
        `mg-axi entra group list --filter <odata-filter> ${profileHint(profileName)}`,
        "0 groups matched; the absence of results is the answer, not an error",
      ],
    };
  }
  return { groups, count, help: [...truncationHints, showHint] };
}

export async function showGroup(
  session: GraphSession,
  flags: GroupFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_GROUPS, KNOWN_GROUP_FIELDS, DEFAULT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const raw = await session.execute({
    profile,
    operation,
    params: { "group-id": String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  });
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed group body", "GRAPH_ERROR", [
      "Single-group reads carry one group object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  const helpHints: string[] = truncated ? [fullHint("entra group show", flags, profileName)] : [];
  if (row["isAssignableToRole"] === true) {
    helpHints.push("Role-assignable group: membership changes need role-management permission and belong to WRITE-01, never to this read");
  }
  if (helpHints.length) return { group: row, help: helpHints };
  return { group: row };
}

async function listRelationship(
  session: GraphSession,
  flags: GroupFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  command: string,
  key: string,
  emptyNoun: string,
  transitive: boolean,
  directWarning: string | undefined,
): Promise<Record<string, unknown>> {
  const group = String(flags.group);
  if (!group.trim()) throw new AxiError("--group needs the group object ID", "VALIDATION_ERROR", [help]);
  const common = collectionCommon(session, flags, KNOWN_MEMBERS, KNOWN_MEMBER_FIELDS, DEFAULT_MEMBER_SELECT, operation, help, profile);
  const args = collectArgs(profile, operation, common, flags, help);
  args.params = { "group-id": group };
  const result = await session.collect(args);
  const effectiveFlags: GroupFlags = { ...flags, select: result.query.$select ?? DEFAULT_MEMBER_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  let limitedInfo = 0;
  for (const row of result.value) {
    const projected = projectMember(row, common.select, common.fields, common.full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
    if (projected.limitedInfo) limitedInfo += 1;
  }
  const scopeHint = transitive
    ? `Direct relationships only: mg-axi ${command} --group ${shellValue(group)} ${profileHint(profileName)}`
    : `Flat nested view: mg-axi ${command} --group ${shellValue(group)} --transitive ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(command, effectiveFlags, profileName)] : [];
  const limitedHints = limitedInfo > 0
    ? [`${limitedInfo} of ${rows.length} rows have no non-null selected descriptive properties; this may reflect limited read consent or unset properties`]
    : [];
  const hiddenHint = "Hidden members are omitted without Member.Read.Hidden; completion describes pagination, not visibility";
  const warnings = directWarning === undefined ? undefined : { warnings: [directWarning] };
  if (!result.complete) {
    return {
      [key]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      ...warnings,
      help: [...truncationHints, ...limitedHints, hiddenHint, resumeHint(profileName), scopeHint],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return {
      [key]: rows,
      count,
      ...warnings,
      help: [
        ...limitedHints,
        hiddenHint,
        `0 ${emptyNoun} matched; the absence of results is the answer, not an error`,
        scopeHint,
      ],
    };
  }
  return { [key]: rows, count, ...warnings, help: [...truncationHints, ...limitedHints, hiddenHint, scopeHint] };
}

export async function listGroupMembers(
  session: GraphSession,
  flags: GroupFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const transitive = flags.transitive === true;
  // The v1.0 service-principal omission is sourced for direct members only;
  // transitive results carry no such claim.
  return listRelationship(session, flags, profile, operation, help, profileName,
    "entra group member list", "members", "members", transitive, transitive ? undefined : SERVICE_PRINCIPAL_WARNING);
}

export async function listGroupMemberOf(
  session: GraphSession,
  flags: GroupFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const transitive = flags.transitive === true;
  return listRelationship(session, flags, profile, operation, help, profileName,
    "entra group member-of list", "memberOf", "memberships", transitive, undefined);
}
