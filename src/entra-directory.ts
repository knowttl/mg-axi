import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// READ-10: the directory-device and administrative-unit read mapping behind
// `mg-axi entra device list/show`, `mg-axi entra administrative-unit
// list/show` and `mg-axi entra administrative-unit member list`. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 device-list, device-get,
// directory-list-administrativeunits and administrativeunit-list-members
// operation documentation on 2026-10-04. Device list/get read D/A
// Device.Read.All; delegated callers additionally need a supported directory
// role (for example Global Reader, Security Reader or Device Managers).
// Administrative-unit list and member reads take D/A
// AdministrativeUnit.Read.All; delegated list callers need a member-user
// account or a supported role with Privileged Role Administrator the
// least-privileged role, while member reads distinguish Directory Readers
// (basic properties and members) from Global Reader (all properties
// including members). Hidden AU memberships additionally need
// Member.Read.Hidden and the server omits what the caller cannot see rather
// than failing. Application callers with narrow consent receive
// limited-information rows carrying only @odata.type and id with other
// properties null; those rows are preserved, never reinterpreted as empty.
// Directory devices are Entra directory objects: Intune managed devices and
// device actions are a separately authorized surface and never this command.
// Base device reads carry no P1/P2 prerequisite; scoped AU administration
// needs P1, members are Free, and dynamic membership needs additional P1
// licensing. No device or AU mutation exists in this slice.

// Every device property this slice may request or display, matching the
// reviewed raw surface. Anything else fails before credentials.
export const KNOWN_DEVICE_FIELDS: readonly string[] = [
  "id",
  "deviceId",
  "displayName",
  "operatingSystem",
  "operatingSystemVersion",
  "trustType",
  "isCompliant",
  "isManaged",
  "accountEnabled",
  "createdDateTime",
  "approximateLastSignInDateTime",
  "manufacturer",
  "model",
];
const KNOWN_DEVICES = new Set(KNOWN_DEVICE_FIELDS);

// Every administrative-unit property this slice may request or display,
// matching the reviewed raw surface. Anything else fails before credentials.
export const KNOWN_AU_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "description",
  "visibility",
  "membershipType",
  "membershipRule",
];
const KNOWN_AUS = new Set(KNOWN_AU_FIELDS);

// AU members are directoryObjects of mixed types (users, groups, devices);
// richer per-type fields need single-object reads, so only these stay
// selectable.
export const KNOWN_AU_MEMBER_FIELDS: readonly string[] = ["id", "displayName", "mail"];
const KNOWN_AU_MEMBERS = new Set(KNOWN_AU_MEMBER_FIELDS);
// @odata.type is preserved on member rows without being selectable: it names
// the member kind and marks limited-information rows.
const MEMBER_TYPE_PROPERTY = "@odata.type";

// Compact device rows: identifier, display name, platform and enabled state.
const DEFAULT_DEVICE_LIST_SELECT = ["id", "displayName", "operatingSystem", "accountEnabled"];
// Show rows: the full reviewed device set, including the directory deviceId
// that is distinct from the object id.
const DEFAULT_DEVICE_SHOW_SELECT = [...KNOWN_DEVICE_FIELDS];
// Compact AU rows: identifier, display name, scope visibility and kind.
const DEFAULT_AU_LIST_SELECT = ["id", "displayName", "visibility", "membershipType"];
// Show rows: the full reviewed AU set, including the membership rule.
const DEFAULT_AU_SHOW_SELECT = [...KNOWN_AU_FIELDS];
// Compact member rows: identifier and display name.
const DEFAULT_AU_MEMBER_SELECT = ["id", "displayName"];
// Delegated defaults are operation-specific; hidden AU members need an
// explicit --scopes set including Member.Read.Hidden. Application profiles
// use their configured .default audience and reject --scopes.
export const DEFAULT_DEVICE_SCOPES = ["https://graph.microsoft.com/Device.Read.All"];
export const DEFAULT_AU_SCOPES = ["https://graph.microsoft.com/AdministrativeUnit.Read.All"];
const TRUNCATE_AT = 500;

export type DirectoryFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown directory property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: DirectoryFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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

// Member projection additionally preserves @odata.type so the member kind
// and limited-information rows survive local projection.
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
  flags: DirectoryFlags,
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

function fullHint(command: string, flags: DirectoryFlags, profileName: string): string {
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

const DEVICE_DENIAL_HINTS = [
  "Device reads need Device.Read.All plus a supported directory role for delegated access (for example Global Reader, Security Reader or Device Managers), or admin-consented Device.Read.All for application access",
  "Directory devices are Entra directory objects; Intune managed devices and device actions are a separately authorized surface, never this command",
  "Base device reads carry no P1/P2 prerequisite; never diagnose licence solely from HTTP 403",
];

const AU_LIST_DENIAL_HINTS = [
  "Administrative-unit lists need AdministrativeUnit.Read.All plus a member-user account or a supported directory role for delegated access (Privileged Role Administrator is the least-privileged role), or admin-consented AdministrativeUnit.Read.All for application access",
  "Scoped administration needs P1, members are Free, and dynamic membership needs additional P1 licensing; never diagnose licence solely from HTTP 403",
];

const AU_SHOW_DENIAL_HINTS = [
  "Administrative-unit reads need AdministrativeUnit.Read.All plus a member-user account or a supported directory role for delegated access (Directory Readers for basic properties, Global Reader for all properties), or admin-consented AdministrativeUnit.Read.All for application access",
  "Scoped administration needs P1, members are Free, and dynamic membership needs additional P1 licensing; never diagnose licence solely from HTTP 403",
];

const AU_MEMBER_DENIAL_HINTS = [
  "Administrative-unit member reads need AdministrativeUnit.Read.All (Directory Readers for basic properties and members, Global Reader for all properties including members), or admin-consented AdministrativeUnit.Read.All for application access",
  "Hidden memberships additionally need Member.Read.Hidden; the server omits what the caller cannot see rather than failing",
  "Scoped administration needs P1, members are Free, and dynamic membership needs additional P1 licensing; never diagnose licence solely from HTTP 403",
];

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
  filter: string | undefined;
}

// $filter on directory collections runs through advanced query: an effective
// filter always carries $count=true with ConsistencyLevel eventual.
// Restoring the saved filter keeps cursor resumes lossless when --filter is
// omitted.
function collectionCommon(
  session: GraphSession,
  flags: DirectoryFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  defaultScopes: string[],
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
  return { cursor, select, fields, scopes: scopesFor(flags, defaultScopes, profile, help), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: DirectoryFlags,
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

export async function listDevices(
  session: GraphSession,
  flags: DirectoryFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_DEVICES, KNOWN_DEVICE_FIELDS,
    DEFAULT_DEVICE_LIST_SELECT, DEFAULT_DEVICE_SCOPES, operation, help, profile);
  const result = await withGuidance(DEVICE_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: DirectoryFlags = { ...flags, select: result.query.$select ?? DEFAULT_DEVICE_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const devices: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    devices.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra device show --id <device-id> ${profileHint(profileName)}`;
  const scopeHint = "Directory devices only; Intune managed devices are a separately authorized surface";
  const truncationHints = truncated ? [fullHint("entra device list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      devices,
      count: { returned: devices.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, scopeHint],
    };
  }
  const count = { returned: devices.length, complete: true };
  if (!devices.length) {
    return {
      devices,
      count,
      help: [
        `mg-axi entra device list --filter <odata-filter> ${profileHint(profileName)}`,
        "0 devices matched; the absence of results is the answer, not an error",
        scopeHint,
      ],
    };
  }
  return { devices, count, help: [...truncationHints, showHint, scopeHint] };
}

export async function showDevice(
  session: GraphSession,
  flags: DirectoryFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_DEVICES, KNOWN_DEVICE_FIELDS, DEFAULT_DEVICE_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_DEVICE_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(DEVICE_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "device-id": String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed device body", "GRAPH_ERROR", [
      "Single-device reads carry one device object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { device: row, help: [fullHint("entra device show", flags, profileName)] };
  return { device: row };
}

export async function listAdministrativeUnits(
  session: GraphSession,
  flags: DirectoryFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_AUS, KNOWN_AU_FIELDS,
    DEFAULT_AU_LIST_SELECT, DEFAULT_AU_SCOPES, operation, help, profile);
  const result = await withGuidance(AU_LIST_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: DirectoryFlags = { ...flags, select: result.query.$select ?? DEFAULT_AU_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const units: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    units.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra administrative-unit show --id <administrative-unit-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra administrative-unit list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      administrativeUnits: units,
      count: { returned: units.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  const count = { returned: units.length, complete: true };
  if (!units.length) {
    return {
      administrativeUnits: units,
      count,
      help: [
        `mg-axi entra administrative-unit list --filter <odata-filter> ${profileHint(profileName)}`,
        "0 administrative units matched; the absence of results is the answer, not an error",
      ],
    };
  }
  return { administrativeUnits: units, count, help: [...truncationHints, showHint] };
}

export async function showAdministrativeUnit(
  session: GraphSession,
  flags: DirectoryFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_AUS, KNOWN_AU_FIELDS, DEFAULT_AU_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_AU_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(AU_SHOW_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "administrativeUnit-id": String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed administrative-unit body", "GRAPH_ERROR", [
      "Single-unit reads carry one administrativeUnit object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  const helpHints: string[] = truncated ? [fullHint("entra administrative-unit show", flags, profileName)] : [];
  if (row["membershipType"] === "Dynamic") {
    helpHints.push("Dynamic membership needs additional P1 licensing; scoped administration needs P1 while members are Free");
  }
  if (helpHints.length) return { administrativeUnit: row, help: helpHints };
  return { administrativeUnit: row };
}

export async function listAdministrativeUnitMembers(
  session: GraphSession,
  flags: DirectoryFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const unit = String(flags["administrative-unit"]);
  if (!unit.trim()) throw new AxiError("--administrative-unit needs the administrative-unit object ID", "VALIDATION_ERROR", [help]);
  const common = collectionCommon(session, flags, KNOWN_AU_MEMBERS, KNOWN_AU_MEMBER_FIELDS,
    DEFAULT_AU_MEMBER_SELECT, DEFAULT_AU_SCOPES, operation, help, profile);
  const args = collectArgs(profile, operation, common, flags, help);
  args.params = { "administrativeUnit-id": unit };
  const result = await withGuidance(AU_MEMBER_DENIAL_HINTS, () => session.collect(args));
  const effectiveFlags: DirectoryFlags = { ...flags, select: result.query.$select ?? DEFAULT_AU_MEMBER_SELECT.join(",") };
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
  const truncationHints = truncated ? [fullHint("entra administrative-unit member list", effectiveFlags, profileName)] : [];
  const limitedHints = limitedInfo > 0
    ? [`${limitedInfo} of ${rows.length} rows have no non-null selected descriptive properties; this may reflect limited read consent or unset properties`]
    : [];
  const hiddenHint = "Hidden memberships are omitted without Member.Read.Hidden; completion describes pagination, not visibility";
  if (!result.complete) {
    return {
      members: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, ...limitedHints, hiddenHint, resumeHint(profileName)],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return {
      members: rows,
      count,
      help: [
        ...limitedHints,
        hiddenHint,
        "0 members matched; the absence of results is the answer, not an error",
      ],
    };
  }
  return { members: rows, count, help: [...truncationHints, ...limitedHints, hiddenHint] };
}
