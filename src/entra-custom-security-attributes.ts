import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 custom-security-attributes subfamily: the read mapping behind
// `mg-axi entra attribute-set list/show/count`,
// `mg-axi entra custom-security-attribute-definition list/show/count` and
// `mg-axi entra allowed-value list/show/count`. Operation construction stays
// beside its command; the shared session owns URLs, credentials, paging,
// retries and error translation, and the SDK owns TOON rendering. This
// module only maps flags to session calls and projects rows for compact
// output.
//
// Reviewed against the v1.0 directory-list-attributesets, attributeset-get,
// directory-list-customsecurityattributedefinitions,
// customsecurityattributedefinition-get,
// customsecurityattributedefinition-list-allowedvalues and allowedvalue-get
// operation documentation and the attributeSet,
// customSecurityAttributeDefinition and allowedValue resource references on
// 2026-10-05. Every read takes D/A CustomSecAttributeDefinition.Read.All;
// delegated callers additionally need a custom-security-attribute role, and
// even Global Administrators have no custom-security-attribute access by
// default (Attribute Definition Reader or Attribute Definition
// Administrator work for every read in this slice). Personal Microsoft
// accounts are not supported. No P1/P2 prerequisite is stated for these
// reads. The attribute-set list documents $select, $top and $orderby only
// ($filter is not supported), so it offers no --filter; the definition
// list documents $select, $top and $filter (eq), so --filter passes through
// as plain $filter with no $count or ConsistencyLevel contract. The
// allowed-values list documents $select only, so it offers no --filter.
// Definition $expand (inline allowedValues) is not reviewed here: allowed
// values have their own list/show commands. The $count routes carry no
// operation-level documentation page; only the definition count accepts
// --filter (its collection documents $filter), the other two take no
// --filter/--select/--limit/--cursor. Beta attribute sets, definitions and
// allowed values and every mutation stay out. No mutation lives here.

// Every attribute-set property this slice may request or display, matching
// the reviewed resource. Anything else fails before credentials.
export const KNOWN_ATTRIBUTE_SET_FIELDS: readonly string[] = [
  "id",
  "description",
  "maxAttributesPerSet",
];
const KNOWN_ATTRIBUTE_SETS = new Set(KNOWN_ATTRIBUTE_SET_FIELDS);

// Every definition property this slice may request or display, matching the
// reviewed resource. The allowedValues navigation property is never
// projected here; it has its own list/show commands. Anything else fails
// before credentials.
export const KNOWN_CUSTOM_SECURITY_DEFINITION_FIELDS: readonly string[] = [
  "attributeSet",
  "description",
  "id",
  "isCollection",
  "isSearchable",
  "name",
  "status",
  "type",
  "usePreDefinedValuesOnly",
];
const KNOWN_DEFINITIONS = new Set(KNOWN_CUSTOM_SECURITY_DEFINITION_FIELDS);

// Every allowed-value property this slice may request or display, matching
// the reviewed resource. Anything else fails before credentials.
export const KNOWN_ALLOWED_VALUE_FIELDS: readonly string[] = [
  "id",
  "isActive",
];
const KNOWN_ALLOWED_VALUES = new Set(KNOWN_ALLOWED_VALUE_FIELDS);

// Attribute sets carry only three short properties, so lists and shows
// share the full reviewed set.
const DEFAULT_ATTRIBUTE_SET_LIST_SELECT = [...KNOWN_ATTRIBUTE_SET_FIELDS];
const DEFAULT_ATTRIBUTE_SET_SHOW_SELECT = [...KNOWN_ATTRIBUTE_SET_FIELDS];
// Compact definition rows: the identifier, the owning set and name, the
// availability status and the data type that decide whether the definition
// is usable.
const DEFAULT_DEFINITION_LIST_SELECT = ["id", "attributeSet", "name", "status", "type"];
// Show rows: the full reviewed definition set.
const DEFAULT_DEFINITION_SHOW_SELECT = [...KNOWN_CUSTOM_SECURITY_DEFINITION_FIELDS];
// Allowed values carry only two short properties, so lists and shows share
// the full reviewed set.
const DEFAULT_ALLOWED_VALUE_LIST_SELECT = [...KNOWN_ALLOWED_VALUE_FIELDS];
const DEFAULT_ALLOWED_VALUE_SHOW_SELECT = [...KNOWN_ALLOWED_VALUE_FIELDS];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_CUSTOM_SECURITY_SCOPES = ["https://graph.microsoft.com/CustomSecAttributeDefinition.Read.All"];
const TRUNCATE_AT = 500;

export type CustomSecurityFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string, noun: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown custom-security-attribute property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known ${noun} properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: CustomSecurityFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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

function selectedFields(
  flags: CustomSecurityFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  help: string,
  noun: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, known, knownList, "select", help, noun);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, known, knownList, "fields", help, noun);
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

function fullHint(command: string, flags: CustomSecurityFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the scope, roles and licensing that actually unlock it,
// because a 403 alone never says which prerequisite is missing. Custom
// security attributes are unusual: even Global Administrators have no
// access by default, so the hints name the attribute roles plainly.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const CUSTOM_SECURITY_DENIAL_HINTS = [
  "Custom-security-attribute reads need CustomSecAttributeDefinition.Read.All for delegated or application access; delegated callers pass it as --scopes",
  "Delegated reads additionally need a custom-security-attribute role even for Global Administrators: Attribute Definition Reader or Attribute Definition Administrator work for every read in this slice (some reads also accept Attribute Assignment Reader or Attribute Assignment Administrator); by default Global Administrator and other administrator roles have no custom-security-attribute access",
  "Personal Microsoft accounts are not supported for custom-security-attribute reads",
  "No P1/P2 prerequisite is stated for custom-security-attribute reads; never diagnose licence solely from HTTP 403",
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
// --select or --filter is omitted. Only the definition list documents
// $filter, so only it passes --filter through as plain $filter with no
// $count or ConsistencyLevel contract; strict input validation refuses the
// flag on the other lists before credentials instead of forwarding a
// misleading request.
function collectionCommon(
  session: GraphSession,
  flags: CustomSecurityFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  noun: string,
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
    saved?.["$select"] === undefined ? defaults : fieldList(saved["$select"], known, knownList, "select", help, noun), help, noun);
  const filter = withFilter
    ? (flags.filter === undefined ? saved?.["$filter"] : String(flags.filter))
    : undefined;
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_CUSTOM_SECURITY_SCOPES, profile, help), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: CustomSecurityFlags,
  help: string,
  params?: Record<string, string>,
): CollectArgs {
  const query: Record<string, string> = { $select: common.select.join(",") };
  if (common.filter !== undefined) query.$filter = common.filter;
  const args: CollectArgs = { profile, operation, query, scopes: common.scopes };
  if (params !== undefined) args.params = params;
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

function attributeSetId(flags: CustomSecurityFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the attribute-set identifier", "VALIDATION_ERROR", [help]);
  return id;
}

function definitionId(flags: CustomSecurityFlags, help: string, flag = "id"): string {
  const id = String(flags[flag]);
  if (!id.trim()) throw new AxiError(`--${flag} needs the custom-security-attribute-definition identifier`, "VALIDATION_ERROR", [help]);
  return id;
}

function allowedValueId(flags: CustomSecurityFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the allowed-value identifier", "VALIDATION_ERROR", [help]);
  return id;
}

function definitionParams(flags: CustomSecurityFlags, help: string): Record<string, string> {
  return { "customSecurityAttributeDefinition-id": definitionId(flags, help, "definition") };
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

export async function listAttributeSets(
  session: GraphSession,
  flags: CustomSecurityFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_ATTRIBUTE_SETS, KNOWN_ATTRIBUTE_SET_FIELDS,
    DEFAULT_ATTRIBUTE_SET_LIST_SELECT, "attribute-set", operation, help, profile, false);
  const result = await withGuidance(CUSTOM_SECURITY_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: CustomSecurityFlags = { ...flags, select: result.query.$select ?? DEFAULT_ATTRIBUTE_SET_LIST_SELECT.join(",") };
  const sets: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    sets.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra attribute-set show --id <set-id> ${profileHint(profileName)}`;
  const filterNote = "Attribute-set lists offer no --filter: Graph documents $select, $top and $orderby only for /directory/attributeSets";
  const truncationHints = truncated ? [fullHint("entra attribute-set list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      attributeSets: sets,
      count: { returned: sets.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, filterNote],
    };
  }
  const count = { returned: sets.length, complete: true };
  if (!sets.length) {
    return {
      attributeSets: sets,
      count,
      help: [filterNote, "0 attribute sets matched; the absence of results is the answer, not an error"],
    };
  }
  return { attributeSets: sets, count, help: [...truncationHints, showHint, filterNote] };
}

export async function showAttributeSet(
  session: GraphSession,
  flags: CustomSecurityFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_ATTRIBUTE_SETS, KNOWN_ATTRIBUTE_SET_FIELDS, DEFAULT_ATTRIBUTE_SET_SHOW_SELECT, help, "attribute-set");
  const scopes = scopesFor(flags, DEFAULT_CUSTOM_SECURITY_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(CUSTOM_SECURITY_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "attributeSet-id": attributeSetId(flags, help) },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "attribute set");
  if (truncated) return { attributeSet: row, help: [fullHint("entra attribute-set show", flags, profileName)] };
  return { attributeSet: row };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute with scalar mode
// and accepts only a non-negative integer. There is no
// --filter/--select/--limit contract on the count: the catalogue declares
// no such flags and strict input validation refuses them before
// credentials.
export async function countAttributeSets(
  session: GraphSession,
  flags: CustomSecurityFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_CUSTOM_SECURITY_SCOPES, profile, help);
  const raw = await withGuidance(CUSTOM_SECURITY_DENIAL_HINTS, () => session.execute({ profile, operation, scopes, scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed attribute-set count body", "GRAPH_ERROR", [
      "Attribute-set counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  return { count: { returned: raw, complete: true } };
}

export async function listCustomSecurityAttributeDefinitions(
  session: GraphSession,
  flags: CustomSecurityFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_DEFINITIONS, KNOWN_CUSTOM_SECURITY_DEFINITION_FIELDS,
    DEFAULT_DEFINITION_LIST_SELECT, "definition", operation, help, profile, true);
  const result = await withGuidance(CUSTOM_SECURITY_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: CustomSecurityFlags = { ...flags, select: result.query.$select ?? DEFAULT_DEFINITION_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const definitions: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    definitions.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra custom-security-attribute-definition show --id <definition-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra custom-security-attribute-definition list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      customSecurityAttributeDefinitions: definitions,
      count: { returned: definitions.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  const count = { returned: definitions.length, complete: true };
  if (!definitions.length) {
    return {
      customSecurityAttributeDefinitions: definitions,
      count,
      help: ["0 custom security attribute definitions matched; the absence of results is the answer, not an error"],
    };
  }
  return { customSecurityAttributeDefinitions: definitions, count, help: [...truncationHints, showHint] };
}

export async function showCustomSecurityAttributeDefinition(
  session: GraphSession,
  flags: CustomSecurityFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_DEFINITIONS, KNOWN_CUSTOM_SECURITY_DEFINITION_FIELDS, DEFAULT_DEFINITION_SHOW_SELECT, help, "definition");
  const scopes = scopesFor(flags, DEFAULT_CUSTOM_SECURITY_SCOPES, profile, help);
  const full = flags.full === true;
  const id = definitionId(flags, help);
  const raw = await withGuidance(CUSTOM_SECURITY_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "customSecurityAttributeDefinition-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "custom security attribute definition");
  const helpHints: string[] = [
    ...(truncated ? [fullHint("entra custom-security-attribute-definition show", flags, profileName)] : []),
    `mg-axi entra allowed-value list --definition ${shellValue(id)} ${profileHint(profileName)}`,
  ];
  return { customSecurityAttributeDefinition: row, help: helpHints };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute with scalar mode
// and accepts only a non-negative integer. --filter narrows the count
// server-side; there is no --select/--limit/--cursor contract on the count.
export async function countCustomSecurityAttributeDefinitions(
  session: GraphSession,
  flags: CustomSecurityFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_CUSTOM_SECURITY_SCOPES, profile, help);
  const query: Record<string, string> = {};
  if (flags.filter !== undefined) query.$filter = String(flags.filter);
  const raw = await withGuidance(CUSTOM_SECURITY_DENIAL_HINTS, () => session.execute({ profile, operation, query, scopes, scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed definition count body", "GRAPH_ERROR", [
      "Definition counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  if (flags.filter !== undefined) {
    return {
      count: { returned: raw, complete: true },
      help: [`Count reflects --filter ${shellValue(String(flags.filter))}; drop --filter for the tenant total`],
    };
  }
  return { count: { returned: raw, complete: true } };
}

export async function listAllowedValues(
  session: GraphSession,
  flags: CustomSecurityFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_ALLOWED_VALUES, KNOWN_ALLOWED_VALUE_FIELDS,
    DEFAULT_ALLOWED_VALUE_LIST_SELECT, "allowed-value", operation, help, profile, false);
  const definition = definitionId(flags, help, "definition");
  const result = await withGuidance(CUSTOM_SECURITY_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help, definitionParams(flags, help))));
  const effectiveFlags: CustomSecurityFlags = { ...flags, select: result.query.$select ?? DEFAULT_ALLOWED_VALUE_LIST_SELECT.join(",") };
  const values: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    values.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra allowed-value show --definition ${shellValue(definition)} --id <value-id> ${profileHint(profileName)}`;
  const filterNote = "Allowed-value lists offer no --filter: Graph documents $select only for allowedValues";
  const truncationHints = truncated ? [fullHint("entra allowed-value list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      allowedValues: values,
      count: { returned: values.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, filterNote],
    };
  }
  const count = { returned: values.length, complete: true };
  if (!values.length) {
    return {
      allowedValues: values,
      count,
      help: [filterNote, "0 allowed values matched; the definition may allow free-form values instead of predefined ones"],
    };
  }
  return { allowedValues: values, count, help: [...truncationHints, showHint, filterNote] };
}

export async function showAllowedValue(
  session: GraphSession,
  flags: CustomSecurityFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_ALLOWED_VALUES, KNOWN_ALLOWED_VALUE_FIELDS, DEFAULT_ALLOWED_VALUE_SHOW_SELECT, help, "allowed-value");
  const scopes = scopesFor(flags, DEFAULT_CUSTOM_SECURITY_SCOPES, profile, help);
  const full = flags.full === true;
  const params = { ...definitionParams(flags, help), "allowedValue-id": allowedValueId(flags, help) };
  const raw = await withGuidance(CUSTOM_SECURITY_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params,
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "allowed value");
  if (truncated) return { allowedValue: row, help: [fullHint("entra allowed-value show", flags, profileName)] };
  return { allowedValue: row };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute with scalar mode
// and accepts only a non-negative integer. There is no
// --filter/--select/--limit contract on the count: the catalogue declares
// no such flags and strict input validation refuses them before
// credentials.
export async function countAllowedValues(
  session: GraphSession,
  flags: CustomSecurityFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_CUSTOM_SECURITY_SCOPES, profile, help);
  const raw = await withGuidance(CUSTOM_SECURITY_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: definitionParams(flags, help),
    scopes,
    scalar: true,
  }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed allowed-value count body", "GRAPH_ERROR", [
      "Allowed-value counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  return { count: { returned: raw, complete: true } };
}
