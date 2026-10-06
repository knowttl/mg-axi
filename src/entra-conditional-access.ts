import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// READ-03: the Conditional Access read mapping behind
// `mg-axi entra conditional-access policy list/show` and
// `mg-axi entra conditional-access named-location list/show` as separate
// az-style grammar, plus the remaining v1.0 named collections under
// `identity/conditionalAccess`: authentication-strength policies, each
// policy's combination configurations, authentication method modes and
// templates. Operation construction stays beside its command; the
// shared session owns URLs, credentials, paging, retries and error
// translation, and the SDK owns TOON rendering. This module only maps flags
// to session calls and projects rows for compact output. No policy mutation
// lives here; that belongs to WRITE-04.
//
// Reviewed against the v1.0 conditionalaccessroot-list-policies,
// conditionalaccesspolicy-get, conditionalaccessroot-list-namedlocations and
// countrynamedlocation-get operation documentation on 2026-10-04, and the
// authenticationstrengthroot-list-policies,
// authenticationstrengthpolicy-get,
// authenticationstrengthpolicy-list-combinationconfigurations,
// authenticationcombinationconfiguration-get,
// authenticationstrengthroot-list-authenticationmethodmodes,
// authenticationmethodmodedetail-get, conditionalaccessroot-list-templates
// and conditionalaccesstemplate-get operation documentation on 2026-10-06.
// Policy and named-location reads need D/A Policy.Read.All; the
// authentication-strength family (policies, combination configurations and
// authentication method modes) reads D/A Policy.Read.AuthenticationMethod
// instead, and template reads D/A Policy.Read.All. Delegated callers
// additionally need a supported directory role (Conditional Access
// Administrator, Security Administrator or Security Reader for the
// strength family; those plus Global Reader and Global Secure Access
// Administrator for policies, named locations and templates).
// Conditional Access needs P1; risk-based Conditional Access needs P2. The
// reviewed raw surface in src/api.ts carries exactly these routes, fields
// and access choices; the named commands below reuse that contract.
//
// Deferred with reason: the parameterless
// `authenticationStrength/policies/{id}/usage()` binds no placeholder, so
// it never passes through the session's validated function-argument
// binding (FUNCTION_ARGUMENT_BINDINGS covers only named-parameter
// segments); widening the request-path guard for one function is out of
// scope. Authentication context class references and the deleted
// policy/named-location collections land as the follow-up split: the
// former needs an AuthenticationContext.Read.All allowlist addition and
// the latter its own sourced review.

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
// Every authentication-strength policy property this slice may request or
// display, matching the reviewed raw surface.
export const KNOWN_STRENGTH_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "description",
  "policyType",
  "requirementsSatisfied",
  "allowedCombinations",
  "createdDateTime",
  "modifiedDateTime",
];
// Every combination-configuration property this slice may request or
// display: the base type carries only an identifier and the combinations
// it applies to; subtype detail rides on @odata.type instead.
export const KNOWN_COMBO_FIELDS: readonly string[] = [
  "id",
  "appliesToCombinations",
];
// Every authentication method mode property this slice may request or
// display, matching the reviewed raw surface.
export const KNOWN_MODE_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "authenticationMethod",
];
// Every Conditional Access template property this slice may request or
// display, matching the reviewed raw surface.
export const KNOWN_TEMPLATE_FIELDS: readonly string[] = [
  "id",
  "name",
  "description",
  "scenarios",
  "details",
];
const STRENGTH_KNOWN = new Set(KNOWN_STRENGTH_FIELDS);
const COMBO_KNOWN = new Set(KNOWN_COMBO_FIELDS);
const MODE_KNOWN = new Set(KNOWN_MODE_FIELDS);
const TEMPLATE_KNOWN = new Set(KNOWN_TEMPLATE_FIELDS);
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
// Compact strength rows: identifier, name and the built-in/custom kind.
const DEFAULT_STRENGTH_LIST_SELECT = ["id", "displayName", "policyType"];
// Fetch the full reviewed strength set for inspection.
const DEFAULT_STRENGTH_SHOW_SELECT = [...KNOWN_STRENGTH_FIELDS];
// Compact combination rows: the identifier; the kind rides as @odata.type.
const DEFAULT_COMBO_LIST_SELECT = ["id"];
// Show rows: the full reviewed base set.
const DEFAULT_COMBO_SHOW_SELECT = [...KNOWN_COMBO_FIELDS];
// Compact mode rows: identifier and name.
const DEFAULT_MODE_LIST_SELECT = ["id", "displayName"];
// Show rows: the full reviewed mode set.
const DEFAULT_MODE_SHOW_SELECT = [...KNOWN_MODE_FIELDS];
// Compact template rows: identifier and name.
const DEFAULT_TEMPLATE_LIST_SELECT = ["id", "name"];
// Show rows: the full reviewed template set.
const DEFAULT_TEMPLATE_SHOW_SELECT = [...KNOWN_TEMPLATE_FIELDS];
// Policy.Read.All covers policies, named locations and templates in both modes.
export const DEFAULT_DELEGATED_SCOPES = ["https://graph.microsoft.com/Policy.Read.All"];
// Policy.Read.AuthenticationMethod is the least-privileged read for the
// authentication-strength family in both modes.
export const DEFAULT_STRENGTH_SCOPES = ["https://graph.microsoft.com/Policy.Read.AuthenticationMethod"];
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

function scopesFor(flags: ConditionalAccessFlags, profile: AnyProfile, help: string, defaults: readonly string[] = DEFAULT_DELEGATED_SCOPES): string[] | undefined {
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
  const args = Object.entries({ ...flags, ...(flags.cursor === undefined ? {} : { cursor: "-" }), profile: profileName, full: true })
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

// The authentication-strength family (strength policies, combination
// configurations, authentication method modes) reads the least-privileged
// Policy.Read.AuthenticationMethod instead of Policy.Read.All, with the
// narrower Conditional Access Administrator / Security Administrator /
// Security Reader role set.
const STRENGTH_DENIAL_HINTS = [
  "Authentication-strength reads need Policy.Read.AuthenticationMethod plus a supported directory role: Conditional Access Administrator, Security Administrator or Security Reader for delegated access, or admin-consented Policy.Read.AuthenticationMethod for application access",
  "Conditional Access needs P1",
];

const TEMPLATE_DENIAL_HINTS = [
  "Conditional Access template reads need Policy.Read.All plus a supported directory role: Conditional Access Administrator, Global Reader, Global Secure Access Administrator, Security Administrator or Security Reader for delegated access, or admin-consented Policy.Read.All for application access",
  "Conditional Access needs P1",
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
  showHint: string;
  emptyHint: string;
  preserveType: boolean;
  // Parent placeholder binding for nested collections (combination
  // configurations hang under one strength policy): the flag is
  // re-validated on every call so resumes rebind the same resource.
  parent?: { param: string; flag: string; label: string };
  // New lists render through the shared list-totals helper; shapes
  // without it keep their historical count object untouched.
  totalsNoun?: string;
}

const POLICY_LIST: CollectionShape = {
  command: "entra conditional-access policy list",
  key: "policies",
  noun: "policy",
  known: POLICY_KNOWN,
  knownList: KNOWN_POLICY_FIELDS,
  defaultSelect: DEFAULT_POLICY_LIST_SELECT,
  defaultScopes: DEFAULT_DELEGATED_SCOPES,
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
  defaultScopes: DEFAULT_DELEGATED_SCOPES,
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
  const scopes = scopesFor(flags, profile, help, shape.defaultScopes);
  const full = flags.full === true;
  const query: Record<string, string> = { $select: select.join(",") };
  if (flags.filter !== undefined) query.$filter = String(flags.filter);
  const args: CollectArgs = { profile, operation, query, scopes };
  if (shape.parent !== undefined) {
    const raw = flags[shape.parent.flag];
    if (raw === undefined || !String(raw).trim()) {
      throw new AxiError(`--${shape.parent.flag} needs the ${shape.parent.label}`, "VALIDATION_ERROR", [help]);
    }
    args.params = { [shape.parent.param]: String(raw) };
  }
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
  if (truncated && cursor !== undefined) truncationHints.push("Supply the original input cursor on stdin to replay this result with --full");
  // Totals lists share the uniform N-of-M line; historical shapes keep
  // their count object byte-for-byte for the concurrent totals lane.
  if (shape.totalsNoun !== undefined) {
    if (!result.complete) {
      return {
        [shape.key]: rows,
        ...listTotals(rows.length, result.total, shape.totalsNoun, false),
        complete: false,
        reason: result.reason,
        cursor: result.cursor,
        help: [...truncationHints, `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`, showHint],
      };
    }
    if (!rows.length) {
      return {
        [shape.key]: rows,
        ...listTotals(rows.length, result.total, shape.totalsNoun, true),
        complete: true,
        help: [
          `mg-axi ${shape.command} --filter <odata-filter> ${profileHint(profileName)}`,
          shape.emptyHint,
        ],
      };
    }
    return {
      [shape.key]: rows,
      ...listTotals(rows.length, result.total, shape.totalsNoun, true),
      complete: true,
      help: [...truncationHints, showHint],
    };
  }
  if (!result.complete) {
    return {
      [shape.key]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor - ${profileHint(profileName)} and supply the returned cursor on stdin`, showHint],
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
  defaultScopes: readonly string[];
  denialHints: string[];
  preserveType: boolean;
  parent?: { param: string; flag: string; label: string };
}

const POLICY_SHOW: SingleShape = {
  command: "entra conditional-access policy show",
  key: "policy",
  param: "conditionalAccessPolicy-id",
  noun: "policy",
  known: POLICY_KNOWN,
  knownList: KNOWN_POLICY_FIELDS,
  defaultSelect: DEFAULT_POLICY_SHOW_SELECT,
  defaultScopes: DEFAULT_DELEGATED_SCOPES,
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
  defaultScopes: DEFAULT_DELEGATED_SCOPES,
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
  const scopes = scopesFor(flags, profile, help, shape.defaultScopes);
  const full = flags.full === true;
  const params: Record<string, string> = { [shape.param]: String(flags.id) };
  if (shape.parent !== undefined) {
    const raw = flags[shape.parent.flag];
    if (raw === undefined || !String(raw).trim()) {
      throw new AxiError(`--${shape.parent.flag} needs the ${shape.parent.label}`, "VALIDATION_ERROR", [help]);
    }
    params[shape.parent.param] = String(raw);
  }
  const raw = await withGuidance(shape.denialHints, () => session.execute({
    profile,
    operation,
    params,
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

const STRENGTH_LIST: CollectionShape = {
  command: "entra conditional-access auth-strength-policy list",
  key: "authStrengthPolicies",
  noun: "auth-strength-policy",
  known: STRENGTH_KNOWN,
  knownList: KNOWN_STRENGTH_FIELDS,
  defaultSelect: DEFAULT_STRENGTH_LIST_SELECT,
  defaultScopes: DEFAULT_STRENGTH_SCOPES,
  denialHints: STRENGTH_DENIAL_HINTS,
  showHint: "mg-axi entra conditional-access auth-strength-policy show --id <auth-strength-policy-id>",
  emptyHint: "0 authentication-strength policies matched; the absence of results is the answer, not an error",
  preserveType: false,
  totalsNoun: "auth-strength policies",
};

const COMBO_LIST: CollectionShape = {
  command: "entra conditional-access combination-configuration list",
  key: "combinationConfigurations",
  noun: "combination-configuration",
  known: COMBO_KNOWN,
  knownList: KNOWN_COMBO_FIELDS,
  defaultSelect: DEFAULT_COMBO_LIST_SELECT,
  defaultScopes: DEFAULT_STRENGTH_SCOPES,
  denialHints: STRENGTH_DENIAL_HINTS,
  showHint: "mg-axi entra conditional-access combination-configuration show --policy <auth-strength-policy-id> --id <combination-configuration-id>",
  emptyHint: "0 combination configurations matched; the absence of results is the answer, not an error",
  preserveType: true,
  parent: { param: "authenticationStrengthPolicy-id", flag: "policy", label: "authentication-strength policy ID owning the combination configurations" },
  totalsNoun: "combination configurations",
};

const MODE_LIST: CollectionShape = {
  command: "entra conditional-access auth-method-mode list",
  key: "authMethodModes",
  noun: "auth-method-mode",
  known: MODE_KNOWN,
  knownList: KNOWN_MODE_FIELDS,
  defaultSelect: DEFAULT_MODE_LIST_SELECT,
  defaultScopes: DEFAULT_STRENGTH_SCOPES,
  denialHints: STRENGTH_DENIAL_HINTS,
  showHint: "mg-axi entra conditional-access auth-method-mode show --id <auth-method-mode-id>",
  emptyHint: "0 authentication method modes matched; the absence of results is the answer, not an error",
  preserveType: false,
  totalsNoun: "auth-method modes",
};

const TEMPLATE_LIST: CollectionShape = {
  command: "entra conditional-access template list",
  key: "templates",
  noun: "template",
  known: TEMPLATE_KNOWN,
  knownList: KNOWN_TEMPLATE_FIELDS,
  defaultSelect: DEFAULT_TEMPLATE_LIST_SELECT,
  defaultScopes: DEFAULT_DELEGATED_SCOPES,
  denialHints: TEMPLATE_DENIAL_HINTS,
  showHint: "mg-axi entra conditional-access template show --id <template-id>",
  emptyHint: "0 Conditional Access templates matched; the absence of results is the answer, not an error",
  preserveType: false,
  totalsNoun: "templates",
};

const STRENGTH_SHOW: SingleShape = {
  command: "entra conditional-access auth-strength-policy show",
  key: "authStrengthPolicy",
  param: "authenticationStrengthPolicy-id",
  noun: "auth-strength-policy",
  known: STRENGTH_KNOWN,
  knownList: KNOWN_STRENGTH_FIELDS,
  defaultSelect: DEFAULT_STRENGTH_SHOW_SELECT,
  defaultScopes: DEFAULT_STRENGTH_SCOPES,
  denialHints: STRENGTH_DENIAL_HINTS,
  preserveType: false,
};

const COMBO_SHOW: SingleShape = {
  command: "entra conditional-access combination-configuration show",
  key: "combinationConfiguration",
  param: "authenticationCombinationConfiguration-id",
  noun: "combination-configuration",
  known: COMBO_KNOWN,
  knownList: KNOWN_COMBO_FIELDS,
  defaultSelect: DEFAULT_COMBO_SHOW_SELECT,
  defaultScopes: DEFAULT_STRENGTH_SCOPES,
  denialHints: STRENGTH_DENIAL_HINTS,
  preserveType: true,
  parent: { param: "authenticationStrengthPolicy-id", flag: "policy", label: "authentication-strength policy ID owning the combination configuration" },
};

const MODE_SHOW: SingleShape = {
  command: "entra conditional-access auth-method-mode show",
  key: "authMethodMode",
  param: "authenticationMethodModeDetail-id",
  noun: "auth-method-mode",
  known: MODE_KNOWN,
  knownList: KNOWN_MODE_FIELDS,
  defaultSelect: DEFAULT_MODE_SHOW_SELECT,
  defaultScopes: DEFAULT_STRENGTH_SCOPES,
  denialHints: STRENGTH_DENIAL_HINTS,
  preserveType: false,
};

const TEMPLATE_SHOW: SingleShape = {
  command: "entra conditional-access template show",
  key: "template",
  param: "conditionalAccessTemplate-id",
  noun: "template",
  known: TEMPLATE_KNOWN,
  knownList: KNOWN_TEMPLATE_FIELDS,
  defaultSelect: DEFAULT_TEMPLATE_SHOW_SELECT,
  defaultScopes: DEFAULT_DELEGATED_SCOPES,
  denialHints: TEMPLATE_DENIAL_HINTS,
  preserveType: false,
};

export async function listAuthStrengthPolicies(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(STRENGTH_LIST, session, flags, profile, operation, help, profileName);
}

export async function showAuthStrengthPolicy(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(STRENGTH_SHOW, session, flags, profile, operation, help, profileName);
}

export async function listCombinationConfigurations(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(COMBO_LIST, session, flags, profile, operation, help, profileName);
}

export async function showCombinationConfiguration(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(COMBO_SHOW, session, flags, profile, operation, help, profileName);
}

export async function listAuthMethodModes(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(MODE_LIST, session, flags, profile, operation, help, profileName);
}

export async function showAuthMethodMode(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(MODE_SHOW, session, flags, profile, operation, help, profileName);
}

export async function listTemplates(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(TEMPLATE_LIST, session, flags, profile, operation, help, profileName);
}

export async function showTemplate(
  session: GraphSession,
  flags: ConditionalAccessFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showSingle(TEMPLATE_SHOW, session, flags, profile, operation, help, profileName);
}
