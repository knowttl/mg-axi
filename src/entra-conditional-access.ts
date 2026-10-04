import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// READ-03: the Conditional Access read mapping behind
// `mg-axi entra conditional-access policy list/show` and
// `mg-axi entra conditional-access named-location list/show` as separate
// az-style grammar. Operation construction stays beside its command; the
// shared session owns URLs, credentials, paging, retries and error
// translation, and the SDK owns TOON rendering. This module only maps flags
// to session calls and projects rows for compact output. No policy mutation
// lives here; that belongs to WRITE-04.
//
// Reviewed against the v1.0 conditionalaccessroot-list-policies,
// conditionalaccesspolicy-get, conditionalaccessroot-list-namedlocations and
// countrynamedlocation-get operation documentation on 2026-10-04. All four
// read D/A Policy.Read.All; delegated callers additionally need a supported
// directory role (Conditional Access Administrator, Global Reader, Global
// Secure Access Administrator, Security Administrator or Security Reader).
// Conditional Access needs P1; risk-based Conditional Access needs P2. The
// reviewed raw surface in src/api.ts carries exactly these routes, fields
// and access choices; the named commands below reuse that contract.

// Every policy property this slice may request or display, matching the
// reviewed raw surface. Anything else fails before credentials.
export const KNOWN_POLICY_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "state",
  "createdDateTime",
  "modifiedDateTime",
  "conditions",
  "grantControls",
  "sessionControls",
];
// Every named-location property this slice may request or display, matching
// the reviewed raw surface. Anything else fails before credentials.
export const KNOWN_LOCATION_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "createdDateTime",
  "modifiedDateTime",
  "isTrusted",
  "ipRanges",
  "countriesAndRegions",
  "includeUnknownCountriesAndRegions",
];
const POLICY_KNOWN = new Set(KNOWN_POLICY_FIELDS);
const LOCATION_KNOWN = new Set(KNOWN_LOCATION_FIELDS);
// @odata.type is preserved on named-location rows without being selectable:
// it names the location kind (ipNamedLocation versus countryNamedLocation).
const LOCATION_TYPE_PROPERTY = "@odata.type";

// Compact policy rows: identifier, name and enforcement state. The state
// rides in the list because an enabled policy and a disabled one read the
// same but act differently.
const DEFAULT_POLICY_LIST_SELECT = ["id", "displayName", "state"];
// Fetch the reviewed condition and control blocks for policy inspection;
// local projection and text truncation still apply unless overridden.
const DEFAULT_POLICY_SHOW_SELECT = [...KNOWN_POLICY_FIELDS];
// Compact location rows: identifier and name; the kind rides as @odata.type.
const DEFAULT_LOCATION_LIST_SELECT = ["id", "displayName"];
// Show rows: the full reviewed location set.
const DEFAULT_LOCATION_SHOW_SELECT = [...KNOWN_LOCATION_FIELDS];
// Policy.Read.All covers both families in both modes.
export const DEFAULT_DELEGATED_SCOPES = ["https://graph.microsoft.com/Policy.Read.All"];
const TRUNCATE_AT = 500;

export type ConditionalAccessFlags = Record<string, string | boolean>;

function fieldList(
  raw: unknown,
  flag: string,
  known: Set<string>,
  knownList: readonly string[],
  noun: string,
  help: string,
): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown ${noun} property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known ${noun} properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: ConditionalAccessFlags, profile: AnyProfile, help: string): string[] | undefined {
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

// Policy conditions and location ranges nest, so truncation recurses through
// objects and arrays. Every truncated value carries a --full hint in help.
function truncateValue(value: unknown, full: boolean): { value: unknown; truncated: boolean } {
  const truncate = (text: string) => full || text.length <= TRUNCATE_AT
    ? { value: text, truncated: false }
    : { value: `${text.slice(0, TRUNCATE_AT)}... (truncated, ${text.length} chars total)`, truncated: true };
  if (typeof value === "string") return truncate(value);
  if (Array.isArray(value)) {
    const projected = value.map(entry => truncateValue(entry, full));
    return { value: projected.map(entry => entry.value), truncated: projected.some(entry => entry.truncated) };
  }
  if (value !== null && typeof value === "object") {
    const projected = Object.entries(value).map(([key, entry]) => [key, truncateValue(entry, full)] as const);
    return {
      value: Object.fromEntries(projected.map(([key, entry]) => [key, entry.value])),
      truncated: projected.some(([, entry]) => entry.truncated),
    };
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

// Named-location projection additionally preserves @odata.type so the
// location kind survives local projection.
function projectLocation(
  row: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  const { row: projected, truncated } = project(row, fields, full);
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  if (typeof source[LOCATION_TYPE_PROPERTY] === "string") projected[LOCATION_TYPE_PROPERTY] = source[LOCATION_TYPE_PROPERTY];
  return { row: projected, truncated };
}

function selectedFields(
  flags: ConditionalAccessFlags,
  defaults: string[],
  known: Set<string>,
  knownList: readonly string[],
  noun: string,
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, "select", known, knownList, noun, help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, "fields", known, knownList, noun, help);
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

function fullHint(command: string, flags: ConditionalAccessFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each family
// adds the roles and licensing that actually unlock it, because a 403 alone
// never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const POLICY_DENIAL_HINTS = [
  "Conditional-access policy reads need Policy.Read.All plus a supported directory role: Conditional Access Administrator, Global Reader, Global Secure Access Administrator, Security Administrator or Security Reader for delegated access, or admin-consented Policy.Read.All for application access",
  "Conditional Access needs P1; risk-based Conditional Access needs P2",
];

const LOCATION_DENIAL_HINTS = [
  "Named-location reads need Policy.Read.All plus a supported directory role: Conditional Access Administrator, Global Reader, Global Secure Access Administrator, Security Administrator or Security Reader for delegated access, or admin-consented Policy.Read.All for application access",
  "Conditional Access and named locations need P1; risk-based Conditional Access needs P2",
];

interface CollectionShape {
  command: string;
  key: string;
  noun: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  denialHints: string[];
  showHint: string;
  emptyHint: string;
  preserveType: boolean;
}

const POLICY_LIST: CollectionShape = {
  command: "entra conditional-access policy list",
  key: "policies",
  noun: "policy",
  known: POLICY_KNOWN,
  knownList: KNOWN_POLICY_FIELDS,
  defaultSelect: DEFAULT_POLICY_LIST_SELECT,
  denialHints: POLICY_DENIAL_HINTS,
  showHint: "mg-axi entra conditional-access policy show --id <policy-id>",
  emptyHint: "0 conditional-access policies matched; the absence of results is the answer, not an error",
  preserveType: false,
};

const LOCATION_LIST: CollectionShape = {
  command: "entra conditional-access named-location list",
  key: "namedLocations",
  noun: "named-location",
  known: LOCATION_KNOWN,
  knownList: KNOWN_LOCATION_FIELDS,
  defaultSelect: DEFAULT_LOCATION_LIST_SELECT,
  denialHints: LOCATION_DENIAL_HINTS,
  showHint: "mg-axi entra conditional-access named-location show --id <named-location-id>",
  emptyHint: "0 named locations matched; the absence of results is the answer, not an error",
  preserveType: true,
};

async function listCollection(
  shape: CollectionShape,
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const savedSelect = cursor === undefined ? undefined : session.cursorQuery(operation, cursor).$select;
  const { select, fields } = selectedFields(flags,
    savedSelect === undefined ? shape.defaultSelect : fieldList(savedSelect, "select", shape.known, shape.knownList, shape.noun, help),
    shape.known, shape.knownList, shape.noun, help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const query: Record<string, string> = { $select: select.join(",") };
  if (flags.filter !== undefined) query.$filter = String(flags.filter);
  const args: CollectArgs = { profile, operation, query, scopes };
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(shape.denialHints, () => session.collect(args));
  const effectiveFlags: ConditionalAccessFlags = { ...flags, select: result.query.$select ?? shape.defaultSelect.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = shape.preserveType ? projectLocation(row, fields, full) : project(row, fields, full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `${shape.showHint} ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(shape.command, effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      [shape.key]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor <cursor-from-output> ${profileHint(profileName)}`, showHint],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return {
      [shape.key]: rows,
      count,
      help: [
        `mg-axi ${shape.command} --filter <odata-filter> ${profileHint(profileName)}`,
        shape.emptyHint,
      ],
    };
  }
  return { [shape.key]: rows, count, help: [...truncationHints, showHint] };
}

interface SingleShape {
  command: string;
  key: string;
  param: string;
  noun: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  denialHints: string[];
  preserveType: boolean;
}

const POLICY_SHOW: SingleShape = {
  command: "entra conditional-access policy show",
  key: "policy",
  param: "conditionalAccessPolicy-id",
  noun: "policy",
  known: POLICY_KNOWN,
  knownList: KNOWN_POLICY_FIELDS,
  defaultSelect: DEFAULT_POLICY_SHOW_SELECT,
  denialHints: POLICY_DENIAL_HINTS,
  preserveType: false,
};

const LOCATION_SHOW: SingleShape = {
  command: "entra conditional-access named-location show",
  key: "namedLocation",
  param: "namedLocation-id",
  noun: "named-location",
  known: LOCATION_KNOWN,
  knownList: KNOWN_LOCATION_FIELDS,
  defaultSelect: DEFAULT_LOCATION_SHOW_SELECT,
  denialHints: LOCATION_DENIAL_HINTS,
  preserveType: true,
};

async function showSingle(
  shape: SingleShape,
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, shape.defaultSelect, shape.known, shape.knownList, shape.noun, help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(shape.denialHints, () => session.execute({
    profile,
    operation,
    params: { [shape.param]: String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed Conditional Access body", "GRAPH_ERROR", [
      "Single-object reads carry one object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = shape.preserveType ? projectLocation(raw, fields, full) : project(raw, fields, full);
  if (truncated) return { [shape.key]: row, help: [fullHint(shape.command, flags, profileName)] };
  return { [shape.key]: row };
}

export async function listPolicies(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(POLICY_LIST, session, flags, profile, operation, help, profileName);
}

export async function showPolicy(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(POLICY_SHOW, session, flags, profile, operation, help, profileName);
}

export async function listNamedLocations(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(LOCATION_LIST, session, flags, profile, operation, help, profileName);
}

export async function showNamedLocation(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(LOCATION_SHOW, session, flags, profile, operation, help, profileName);
}
