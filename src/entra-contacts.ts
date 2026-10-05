import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 contacts subfamily plus EXT-01m/EXT-01n navigation: the read mapping
// behind `mg-axi entra contact list/show/count`, the manager/directReports
// navigation reads (`show-manager`, `list/show/count-direct-reports`) and the
// memberOf/transitiveMemberOf membership reads (`list/show/count-member-of`).
// Operation construction stays beside its command; the shared session owns URLs, credentials, paging,
// retries and error translation, and the SDK owns TOON rendering. This
// module only maps flags to session calls and projects rows for compact
// output.
//
// Reviewed against the v1.0 orgcontact-list, orgcontact-get and orgContact
// resource documentation on 2026-10-05. Both reads take delegated or
// application OrgContact.Read.All, the documented least privilege in each
// mode (Directory.Read.All and Directory.ReadWrite.All are documented only
// as higher-privileged alternatives); personal Microsoft accounts are not
// supported. Delegated callers additionally need a supported Entra role:
// Directory Readers reads basic properties, and Global Reader, Directory
// Writers, Intune Administrator or User Administrator also work. No P1/P2
// prerequisite is stated for these reads. The $count scalar carries no
// operation-level documentation page and follows the same OrgContact.Read.All
// contract; the list page's $count example sends ConsistencyLevel eventual,
// so the scalar does too. Membership reads are reviewed against the v1.0
// orgcontact-list-memberof and orgcontact-list-transitivememberof operation
// documentation plus the orgContact resource reference on 2026-10-05 (see the
// EXT-01n block below): direct reads take delegated or application
// OrgContact.Read.All, the documented least privilege in each mode
// (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are
// documented only as higher-privileged alternatives), while transitive reads
// take delegated or application OrgContact.Read.All and Group.Read.All
// together (Directory.Read.All is the documented higher-privileged
// alternative); the same supported Entra roles, personal-account exclusion
// and no-P1/P2 prerequisite apply. The delta() sync, the serviceProvisioningErrors and
// onPremisesSyncBehavior reads (unavailable with no documented permission
// contract), every POST lookup action, beta and every mutation stay out. No mutation lives here.
//
// Secrecy by construction: contacts are personal data, so default rows carry
// only id, displayName, mail and companyName; identifying fields beyond that
// need an explicit --select. Top-level list and show request and project
// only flat scalar properties; $expand is not offered, so navigation objects
// can never appear there, and the nested phones/addresses collections need
// their own projection review. Navigation rows are directory objects with
// the @odata.type discriminator plus minimal fields (see the EXT-01m block
// below). Membership rows are directory objects too: the @odata.type
// discriminator plus id and displayName by default (see the EXT-01n block
// below).

// Every contact property this slice may request or display, matching the
// reviewed resource. Anything else fails before credentials.
export const KNOWN_CONTACT_FIELDS: readonly string[] = [
  "id",
  "deletedDateTime",
  "companyName",
  "department",
  "displayName",
  "givenName",
  "surname",
  "mail",
  "mailNickname",
  "jobTitle",
  "onPremisesLastSyncDateTime",
  "onPremisesSyncEnabled",
  "proxyAddresses",
];
const KNOWN_CONTACTS = new Set(KNOWN_CONTACT_FIELDS);

// Compact rows stay minimal on purpose: contacts are personal data, so list
// and show both default to id, displayName, mail and companyName only.
const DEFAULT_LIST_SELECT = ["id", "displayName", "mail", "companyName"];
const DEFAULT_SHOW_SELECT = [...DEFAULT_LIST_SELECT];
// OrgContact.Read.All covers list, show and count in both modes; delegated
// reads default to it unless --scopes overrides it. It is new to the shared
// READ_SCOPES allowlist, and it is a read scope, so no deferral is needed
// here.
export const DEFAULT_CONTACT_SCOPES = ["https://graph.microsoft.com/OrgContact.Read.All"];
const TRUNCATE_AT = 500;

// The only operations this slice ever binds: the three catalogued v1.0
// top-level contacts GETs, the ten EXT-01m manager/directReports navigation
// GETs and the eighteen EXT-01n memberOf/transitiveMemberOf membership GETs
// (lists, $count scalars, singles and typed casts). Anything else - delta
// sync, error/sync reads, POST lookup actions, writes - is refused before
// credentials.
const READ_OPERATIONS: Readonly<Record<string, string>> = {
  "GET:/contacts": "contact",
  "GET:/contacts/{orgContact-id}": "contact",
  "GET:/contacts/$count": "contact",
  "GET:/contacts/{orgContact-id}/manager": "contact-navigation",
  "GET:/contacts/{orgContact-id}/directReports": "contact-navigation",
  "GET:/contacts/{orgContact-id}/directReports/$count": "contact-navigation",
  "GET:/contacts/{orgContact-id}/directReports/graph.orgContact": "contact-navigation",
  "GET:/contacts/{orgContact-id}/directReports/graph.orgContact/$count": "contact-navigation",
  "GET:/contacts/{orgContact-id}/directReports/graph.user": "contact-navigation",
  "GET:/contacts/{orgContact-id}/directReports/graph.user/$count": "contact-navigation",
  "GET:/contacts/{orgContact-id}/directReports/{directoryObject-id}": "contact-navigation",
  "GET:/contacts/{orgContact-id}/directReports/{directoryObject-id}/graph.orgContact": "contact-navigation",
  "GET:/contacts/{orgContact-id}/directReports/{directoryObject-id}/graph.user": "contact-navigation",
  "GET:/contacts/{orgContact-id}/memberOf": "contact-membership",
  "GET:/contacts/{orgContact-id}/memberOf/$count": "contact-membership",
  "GET:/contacts/{orgContact-id}/memberOf/graph.administrativeUnit": "contact-membership",
  "GET:/contacts/{orgContact-id}/memberOf/graph.administrativeUnit/$count": "contact-membership",
  "GET:/contacts/{orgContact-id}/memberOf/graph.group": "contact-membership",
  "GET:/contacts/{orgContact-id}/memberOf/graph.group/$count": "contact-membership",
  "GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}": "contact-membership",
  "GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}/graph.administrativeUnit": "contact-membership",
  "GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}/graph.group": "contact-membership",
  "GET:/contacts/{orgContact-id}/transitiveMemberOf": "contact-membership",
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/$count": "contact-membership",
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.administrativeUnit": "contact-membership",
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.administrativeUnit/$count": "contact-membership",
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.group": "contact-membership",
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.group/$count": "contact-membership",
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}": "contact-membership",
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}/graph.administrativeUnit": "contact-membership",
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}/graph.group": "contact-membership",
};

function checkReadOperation(operation: SessionOperation, help: string): void {
  const route = `${operation.method}:${operation.path}`;
  if (operation.method !== "GET" || !Object.hasOwn(READ_OPERATIONS, route)) {
    throw new AxiError(`Refused non-read route ${operation.id}: contact serves only the catalogued v1.0 top-level, manager/directReports and memberOf/transitiveMemberOf membership GETs`, "VALIDATION_ERROR", [
      help,
      "Delta sync, error/sync reads, POST lookup actions and writes are never constructed here; beta needs its own review",
    ]);
  }
}

export type ContactFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown contact property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known contact properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: ContactFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
  flags: ContactFlags,
  defaults: string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, KNOWN_CONTACTS, KNOWN_CONTACT_FIELDS, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, KNOWN_CONTACTS, KNOWN_CONTACT_FIELDS, "fields", help);
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

function fullHint(command: string, flags: ContactFlags, profileName: string): string {
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

const EMPTY_NOTE = "0 organizational contacts returned; tenants without synced or Exchange-provisioned contacts carry none, so the absence of rows is the answer, not an error";

const CONTACT_DENIAL_HINTS = [
  "Contact reads need OrgContact.Read.All for delegated or application access (Directory.Read.All is a documented higher-privileged alternative); delegated callers pass it as --scopes",
  "Delegated contact reads additionally need a supported Entra role: Directory Readers reads basic properties, and Global Reader, Directory Writers, Intune Administrator or User Administrator also work",
  "Personal Microsoft accounts are not supported for contact reads",
  "No P1/P2 prerequisite is stated for contact reads; never diagnose role or licence solely from HTTP 403",
];

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  filter: string | undefined;
  scopes: string[] | undefined;
  full: boolean;
}

// Restoring the saved select and filter keeps cursor resumes lossless when
// --select/--filter are omitted. --filter passes through as plain $filter
// with $count=true and ConsistencyLevel eventual, the documented
// advanced-query contract for directory-object filtering; $search and
// $orderby stay unreviewed and strict input validation refuses them before
// credentials.
function collectionCommon(
  session: GraphSession,
  flags: ContactFlags,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    saved?.["$select"] === undefined ? DEFAULT_LIST_SELECT : fieldList(saved["$select"], KNOWN_CONTACTS, KNOWN_CONTACT_FIELDS, "select", help), help);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, filter, scopes: scopesFor(flags, DEFAULT_CONTACT_SCOPES, profile, help), full: flags.full === true };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: ContactFlags,
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

function contactId(flags: ContactFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the organizational-contact object ID", "VALIDATION_ERROR", [help]);
  return id;
}

function singleResult(
  raw: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed contact body", "GRAPH_ERROR", [
      "Single-contact reads carry one orgContact; treat anything else as unknown, not empty",
    ]);
  }
  return project(raw, fields, full);
}

export async function listContacts(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const common = collectionCommon(session, flags, operation, help, profile);
  const result = await withGuidance(CONTACT_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: ContactFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const contacts: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    contacts.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra contact show --id <contact-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra contact list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      contacts,
      count: { returned: contacts.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  const count = { returned: contacts.length, complete: true };
  if (!contacts.length) {
    return {
      contacts,
      count,
      help: [
        `mg-axi entra contact list --filter <odata-filter> ${profileHint(profileName)}`,
        "0 contacts matched; the absence of results is the answer, not an error",
      ],
    };
  }
  return { contacts, count, help: [...truncationHints, showHint] };
}

export async function showContact(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const { select, fields } = selectedFields(flags, DEFAULT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_CONTACT_SCOPES, profile, help);
  const full = flags.full === true;
  const id = contactId(flags, help);
  const raw = await withGuidance(CONTACT_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "orgContact-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full);
  if (truncated) return { contact: row, help: [fullHint("entra contact show", flags, profileName)] };
  return { contact: row };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute and accepts only
// a non-negative integer. The list page's $count example sends
// ConsistencyLevel eventual, so the scalar declares it too. There is no
// --filter/--select/--limit contract on the count: the catalogue declares no
// such flags and strict input validation refuses them before credentials.
export async function countContacts(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  checkReadOperation(operation, help);
  const scopes = scopesFor(flags, DEFAULT_CONTACT_SCOPES, profile, help);
  const raw = await withGuidance(CONTACT_DENIAL_HINTS, () => session.execute({ profile, operation, scopes, consistencyLevel: "eventual", scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed contact count body", "GRAPH_ERROR", [
      "Contact counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  const count = { returned: raw, complete: true };
  return raw === 0 ? { count, help: [EMPTY_NOTE] } : { count };
}

// EXT-01m navigation (second part): the read mapping behind
// `mg-axi entra contact show-manager`, `list-direct-reports`,
// `show-direct-report` and `count-direct-reports`. Results are directory
// objects of mixed user/contact type, so rows carry the @odata.type
// discriminator plus minimal personal-data fields by default (id and
// displayName); mail needs an explicit --select and richer per-type fields
// need the subtype's single-object reads.
//
// Reviewed against the v1.0 orgcontact-get-manager and
// orgcontact-list-directreports operation documentation plus the orgContact
// resource reference on 2026-10-05. Both reads take delegated or
// application OrgContact.Read.All, the documented least privilege in each
// mode (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are
// documented only as higher-privileged alternatives); personal Microsoft
// accounts are not supported. Delegated callers additionally need a
// supported Entra role: Directory Readers reads basic properties, and
// Global Reader, Directory Writers, Intune Administrator or User
// Administrator also work. No P1/P2 prerequisite is stated for these reads.
// The directReports page documents $select only, so $filter/$search/$top
// stay unreviewed and strict input validation refuses them before
// credentials. The $count scalars carry no operation-level documentation
// page and follow the contacts $count contract (ConsistencyLevel eventual).
// The typed casts (graph.user, graph.orgContact) select one subtype through
// the documented OData-cast route with the same permission contract. The
// memberOf/transitiveMemberOf membership reads live in the EXT-01n block
// below; serviceProvisioningErrors and onPremisesSyncBehavior stay out as
// unavailable with no documented permission contract. Application callers with narrow consent receive
// limited-information rows carrying only @odata.type and id; those rows are
// preserved, never reinterpreted as empty.

// Every navigation property this slice may request or display. Users and
// contacts both carry these; anything else fails before credentials.
export const KNOWN_NAV_FIELDS: readonly string[] = ["id", "displayName", "mail"];
const KNOWN_NAV = new Set(KNOWN_NAV_FIELDS);
// @odata.type is preserved on navigation rows without being selectable: it
// names the directory-object kind and marks limited-information rows.
const NAV_TYPE_PROPERTY = "@odata.type";
// Compact navigation rows stay minimal on purpose: id and displayName only.
const DEFAULT_NAV_SELECT = ["id", "displayName"];
// The --as values mirror the Graph subtype names in the cast routes.
export const NAV_CAST_VALUES: readonly string[] = ["user", "orgContact"];

// Base direct-report operation to its typed-cast counterparts. --as selects
// between these catalogued routes; anything else is refused before
// credentials and never falls back to the unfiltered route silently.
const NAV_CASTS: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  "GET:/contacts/{orgContact-id}/directReports": {
    user: "GET:/contacts/{orgContact-id}/directReports/graph.user",
    orgContact: "GET:/contacts/{orgContact-id}/directReports/graph.orgContact",
  },
  "GET:/contacts/{orgContact-id}/directReports/$count": {
    user: "GET:/contacts/{orgContact-id}/directReports/graph.user/$count",
    orgContact: "GET:/contacts/{orgContact-id}/directReports/graph.orgContact/$count",
  },
  "GET:/contacts/{orgContact-id}/directReports/{directoryObject-id}": {
    user: "GET:/contacts/{orgContact-id}/directReports/{directoryObject-id}/graph.user",
    orgContact: "GET:/contacts/{orgContact-id}/directReports/{directoryObject-id}/graph.orgContact",
  },
};

export function castDirectReports(base: string, raw: unknown, help: string): string {
  const value = String(raw);
  const alternate = NAV_CASTS[base]?.[value];
  if (!alternate) {
    throw new AxiError(`--as ${value} is not a direct-report cast`, "VALIDATION_ERROR", [
      help,
      `--as takes ${NAV_CAST_VALUES.join(" or ")} on direct-report list, show and count reads only`,
    ]);
  }
  return alternate;
}

// Navigation projection additionally preserves @odata.type so the
// directory-object kind and limited-information rows survive projection.
function projectNav(
  row: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean; limitedInfo: boolean } {
  const { row: projected, truncated } = project(row, fields, full);
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  if (typeof source[NAV_TYPE_PROPERTY] === "string") projected[NAV_TYPE_PROPERTY] = source[NAV_TYPE_PROPERTY];
  const descriptive = [...new Set([...fields, ...Object.keys(source)])]
    .filter(field => field !== "id" && !field.startsWith("@"));
  const limitedInfo = Object.hasOwn(source, "id") && descriptive.length > 0
    && descriptive.every(field => source[field] === null || source[field] === undefined || !Object.hasOwn(source, field));
  return { row: projected, truncated, limitedInfo };
}

interface NavCollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
}

// The directReports page documents $select only: no $filter passthrough and
// no ConsistencyLevel contract here, so strict validation refuses filter
// flags before credentials and cursor resumes restore only the $select.
function navCollectionCommon(
  session: GraphSession,
  flags: ContactFlags,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): NavCollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const defaults = saved?.["$select"] === undefined ? [...DEFAULT_NAV_SELECT] : fieldList(saved["$select"], KNOWN_NAV, KNOWN_NAV_FIELDS, "select", help);
  const { select, fields } = navFieldsFor(flags, help, defaults);
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_CONTACT_SCOPES, profile, help), full: flags.full === true };
}

function navCollectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: NavCollectionCommon,
  flags: ContactFlags,
  help: string,
): CollectArgs {
  const args: CollectArgs = { profile, operation, query: { $select: common.select.join(",") }, scopes: common.scopes };
  if (common.cursor !== undefined) args.cursor = common.cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  return args;
}

function navSingleResult(
  raw: unknown,
  fields: string[],
  full: boolean,
  noun: string,
): { row: Record<string, unknown>; truncated: boolean; limitedInfo: boolean } {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError(`Graph returned a malformed ${noun} body`, "GRAPH_ERROR", [
      `${noun} reads carry one directoryObject; treat anything else as unknown, not empty`,
    ]);
  }
  return projectNav(raw, fields, full);
}

function navFieldsFor(flags: ContactFlags, help: string, defaults: string[] = DEFAULT_NAV_SELECT): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, KNOWN_NAV, KNOWN_NAV_FIELDS, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, KNOWN_NAV, KNOWN_NAV_FIELDS, "fields", help);
  const missing = fields.find(field => !select.includes(field));
  if (missing) {
    throw new AxiError(`--fields ${missing} was not fetched; request it with --select`, "VALIDATION_ERROR", [help]);
  }
  return { select, fields };
}

export async function showContactManager(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const { select, fields } = navFieldsFor(flags, help);
  const scopes = scopesFor(flags, DEFAULT_CONTACT_SCOPES, profile, help);
  const full = flags.full === true;
  const id = contactId(flags, help);
  const raw = await withGuidance(CONTACT_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "orgContact-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = navSingleResult(raw, fields, full, "manager");
  if (truncated) return { manager: row, help: [fullHint("entra contact show-manager", flags, profileName)] };
  return { manager: row };
}

export async function listContactDirectReports(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const common = navCollectionCommon(session, flags, operation, help, profile);
  const id = contactId(flags, help);
  const args = navCollectArgs(profile, operation, common, flags, help);
  args.params = { "orgContact-id": id };
  const result = await withGuidance(CONTACT_DENIAL_HINTS, () => session.collect(args));
  const effectiveFlags: ContactFlags = { ...flags, select: result.query.$select ?? DEFAULT_NAV_SELECT.join(",") };
  const directReports: Record<string, unknown>[] = [];
  let truncated = false;
  let limitedInfo = 0;
  for (const row of result.value) {
    const projected = projectNav(row, common.fields, common.full);
    directReports.push(projected.row);
    truncated = truncated || projected.truncated;
    if (projected.limitedInfo) limitedInfo += 1;
  }
  const showHint = `mg-axi entra contact show-direct-report --id <contact-id> --report-id <report-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra contact list-direct-reports", effectiveFlags, profileName)] : [];
  const limitedHints = limitedInfo > 0
    ? [`${limitedInfo} of ${directReports.length} rows carry only type and id; this may reflect limited read consent or unset properties`]
    : [];
  if (!result.complete) {
    return {
      directReports,
      count: { returned: directReports.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, ...limitedHints, resumeHint(profileName), showHint],
    };
  }
  const count = { returned: directReports.length, complete: true };
  if (!directReports.length) {
    return {
      directReports,
      count,
      help: [...limitedHints, "0 direct reports matched; the absence of results is the answer, not an error", showHint],
    };
  }
  return { directReports, count, help: [...truncationHints, ...limitedHints, showHint] };
}

export async function showContactDirectReport(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const { select, fields } = navFieldsFor(flags, help);
  const scopes = scopesFor(flags, DEFAULT_CONTACT_SCOPES, profile, help);
  const full = flags.full === true;
  const id = contactId(flags, help);
  const reportId = String(flags["report-id"]);
  if (!reportId.trim()) throw new AxiError("--report-id needs the direct-report directory-object ID", "VALIDATION_ERROR", [help]);
  const raw = await withGuidance(CONTACT_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "orgContact-id": id, "directoryObject-id": reportId },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = navSingleResult(raw, fields, full, "direct-report");
  if (truncated) return { directReport: row, help: [fullHint("entra contact show-direct-report", flags, profileName)] };
  return { directReport: row };
}

// The $count routes return a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute and accepts only
// a non-negative integer. The scalars carry no operation-level
// documentation page and follow the contacts $count contract
// (ConsistencyLevel eventual). There is no --select/--limit contract on the
// count: strict input validation refuses those flags before credentials.
export async function countContactDirectReports(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  checkReadOperation(operation, help);
  const scopes = scopesFor(flags, DEFAULT_CONTACT_SCOPES, profile, help);
  const id = contactId(flags, help);
  const raw = await withGuidance(CONTACT_DENIAL_HINTS, () => session.execute({
    profile, operation, params: { "orgContact-id": id }, scopes, consistencyLevel: "eventual", scalar: true,
  }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed direct-report count body", "GRAPH_ERROR", [
      "Direct-report counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  const count = { returned: raw, complete: true };
  return raw === 0
    ? { count, help: ["0 direct reports returned; a contact with no reports set as their manager carries none, so the absence of rows is the answer, not an error"] }
    : { count };
}

// EXT-01n membership (third part): the read mapping behind
// `mg-axi entra contact list-member-of`, `show-member-of` and
// `count-member-of`. Results are directory objects of group or
// administrativeUnit type, so rows carry the @odata.type discriminator plus
// minimal fields by default (id and displayName); mail needs an explicit
// --select and richer per-type fields need the subtype's single-object
// reads. --transitive selects the transitiveMemberOf flat closure instead of
// direct memberOf, and --as selects one typed cast route (graph.group or
// graph.administrativeUnit).
//
// Reviewed against the v1.0 orgcontact-list-memberof and
// orgcontact-list-transitivememberof operation documentation plus the
// orgContact resource reference on 2026-10-05. Direct reads take delegated
// or application OrgContact.Read.All, the documented least privilege in each
// mode (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are
// documented only as higher-privileged alternatives). Transitive reads take
// delegated or application OrgContact.Read.All and Group.Read.All together,
// the documented least privilege in each mode (Directory.Read.All is the
// documented higher-privileged alternative); both scopes are already in the
// shared READ_SCOPES allowlist, so no deferral is needed here. Delegated
// callers additionally need a supported Entra role: Directory Readers reads
// basic properties, and Global Reader, Directory Writers, Intune
// Administrator or User Administrator also work. Personal Microsoft accounts
// are not supported and no P1/P2 prerequisite is stated for these reads.
// The list pages document $filter, $count, $select, $search and $top with
// advanced query parameters only, so --filter passes through as plain
// $filter with $count=true and ConsistencyLevel eventual while $search and
// $orderby stay unreviewed and strict input validation refuses them before
// credentials. The $count scalars and the single/cast routes carry no
// operation-level documentation page and follow the parent list contract
// ($select only on singles, ConsistencyLevel eventual on scalars). The typed
// casts (graph.group, graph.administrativeUnit) select one subtype through
// the documented OData-cast route with the same permission contract.
// Application callers with narrow consent receive limited-information rows
// carrying only @odata.type and id; those rows are preserved, never
// reinterpreted as empty.

// Direct reads default to OrgContact.Read.All; transitive reads additionally
// need Group.Read.All. Both are read scopes already allowlisted.
export const DEFAULT_MEMBER_OF_SCOPES = ["https://graph.microsoft.com/OrgContact.Read.All"];
export const DEFAULT_TRANSITIVE_MEMBER_OF_SCOPES = [
  "https://graph.microsoft.com/OrgContact.Read.All",
  "https://graph.microsoft.com/Group.Read.All",
];
// The --as values mirror the Graph subtype names in the cast routes.
export const MEMBERSHIP_CAST_VALUES: readonly string[] = ["group", "administrativeUnit"];

// Direct membership operation to its transitive flat-view counterpart.
// --transitive selects between these two catalogued routes; anything else is
// refused before credentials and never falls back silently.
const MEMBERSHIP_TRANSITIVE: Readonly<Record<string, string>> = {
  "GET:/contacts/{orgContact-id}/memberOf": "GET:/contacts/{orgContact-id}/transitiveMemberOf",
  "GET:/contacts/{orgContact-id}/memberOf/$count": "GET:/contacts/{orgContact-id}/transitiveMemberOf/$count",
  "GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}": "GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}",
};

export function transitMembership(base: string, help: string): string {
  const alternate = MEMBERSHIP_TRANSITIVE[base];
  if (!alternate) {
    throw new AxiError("--transitive is available for contact member-of reads only", "VALIDATION_ERROR", [help]);
  }
  return alternate;
}

// Uncast membership operation to its typed-cast counterparts. --as selects
// between these catalogued routes; anything else is refused before
// credentials and never falls back to the unfiltered route silently.
const MEMBERSHIP_CASTS: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  "GET:/contacts/{orgContact-id}/memberOf": {
    group: "GET:/contacts/{orgContact-id}/memberOf/graph.group",
    administrativeUnit: "GET:/contacts/{orgContact-id}/memberOf/graph.administrativeUnit",
  },
  "GET:/contacts/{orgContact-id}/memberOf/$count": {
    group: "GET:/contacts/{orgContact-id}/memberOf/graph.group/$count",
    administrativeUnit: "GET:/contacts/{orgContact-id}/memberOf/graph.administrativeUnit/$count",
  },
  "GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}": {
    group: "GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}/graph.group",
    administrativeUnit: "GET:/contacts/{orgContact-id}/memberOf/{directoryObject-id}/graph.administrativeUnit",
  },
  "GET:/contacts/{orgContact-id}/transitiveMemberOf": {
    group: "GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.group",
    administrativeUnit: "GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.administrativeUnit",
  },
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/$count": {
    group: "GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.group/$count",
    administrativeUnit: "GET:/contacts/{orgContact-id}/transitiveMemberOf/graph.administrativeUnit/$count",
  },
  "GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}": {
    group: "GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}/graph.group",
    administrativeUnit: "GET:/contacts/{orgContact-id}/transitiveMemberOf/{directoryObject-id}/graph.administrativeUnit",
  },
};

export function castMembership(base: string, raw: unknown, help: string): string {
  const value = String(raw);
  const alternate = MEMBERSHIP_CASTS[base]?.[value];
  if (!alternate) {
    throw new AxiError(`--as ${value} is not a member-of cast`, "VALIDATION_ERROR", [
      help,
      `--as takes ${MEMBERSHIP_CAST_VALUES.join(" or ")} on member-of list, show and count reads only`,
    ]);
  }
  return alternate;
}

function membershipScopesFor(flags: ContactFlags, profile: AnyProfile, help: string): string[] | undefined {
  const defaults = flags.transitive === true ? DEFAULT_TRANSITIVE_MEMBER_OF_SCOPES : DEFAULT_MEMBER_OF_SCOPES;
  return scopesFor(flags, defaults, profile, help);
}

const MEMBER_OF_DENIAL_HINTS = [
  "Contact memberOf reads need OrgContact.Read.All for delegated or application access (Directory.Read.All, Directory.ReadWrite.All and Group.Read.All are documented higher-privileged alternatives); delegated callers pass it as --scopes",
  "Delegated memberOf reads additionally need a supported Entra role: Directory Readers reads basic properties, and Global Reader, Directory Writers, Intune Administrator or User Administrator also work",
  "Personal Microsoft accounts are not supported for memberOf reads",
  "No P1/P2 prerequisite is stated for memberOf reads; never diagnose role or licence solely from HTTP 403",
];

const TRANSITIVE_MEMBER_OF_DENIAL_HINTS = [
  "Contact transitiveMemberOf reads need OrgContact.Read.All and Group.Read.All together for delegated or application access (Directory.Read.All is the documented higher-privileged alternative); delegated callers pass both as --scopes",
  "With only one of the two scopes, groups outside the granted type return limited-information rows carrying only type and id",
  "Delegated transitiveMemberOf reads additionally need a supported Entra role: Directory Readers reads basic properties, and Global Reader, Directory Writers, Intune Administrator or User Administrator also work",
  "Personal Microsoft accounts are not supported for transitiveMemberOf reads",
  "No P1/P2 prerequisite is stated for transitiveMemberOf reads; never diagnose role or licence solely from HTTP 403",
];

function membershipDenialHints(flags: ContactFlags): string[] {
  return flags.transitive === true ? TRANSITIVE_MEMBER_OF_DENIAL_HINTS : MEMBER_OF_DENIAL_HINTS;
}

interface MembershipCollectionCommon extends NavCollectionCommon {
  filter: string | undefined;
}

// The memberOf pages document $filter alongside $select with advanced query
// parameters only, so --filter passes through as plain $filter with
// $count=true and ConsistencyLevel eventual (the groups memberOf contract);
// $search and $orderby stay unreviewed. Restoring the saved select and
// filter keeps cursor resumes lossless when --select/--filter are omitted.
function membershipCollectionCommon(
  session: GraphSession,
  flags: ContactFlags,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): MembershipCollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const defaults = saved?.["$select"] === undefined ? [...DEFAULT_NAV_SELECT] : fieldList(saved["$select"], KNOWN_NAV, KNOWN_NAV_FIELDS, "select", help);
  const { select, fields } = navFieldsFor(flags, help, defaults);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, scopes: membershipScopesFor(flags, profile, help), full: flags.full === true, filter };
}

function membershipCollectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: MembershipCollectionCommon,
  flags: ContactFlags,
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

function membershipId(flags: ContactFlags, help: string): string {
  const id = flags["member-id"] === undefined ? "" : String(flags["member-id"]);
  if (!id.trim()) throw new AxiError("--member-id needs the membership directory-object ID", "VALIDATION_ERROR", [help]);
  return id;
}

export async function listContactMemberOf(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const transitive = flags.transitive === true;
  const common = membershipCollectionCommon(session, flags, operation, help, profile);
  const id = contactId(flags, help);
  const args = membershipCollectArgs(profile, operation, common, flags, help);
  args.params = { "orgContact-id": id };
  const result = await withGuidance(membershipDenialHints(flags), () => session.collect(args));
  const effectiveFlags: ContactFlags = { ...flags, select: result.query.$select ?? DEFAULT_NAV_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const memberOf: Record<string, unknown>[] = [];
  let truncated = false;
  let limitedInfo = 0;
  for (const row of result.value) {
    const projected = projectNav(row, common.fields, common.full);
    memberOf.push(projected.row);
    truncated = truncated || projected.truncated;
    if (projected.limitedInfo) limitedInfo += 1;
  }
  const noun = transitive ? "transitiveMemberOf" : "memberOf";
  const command = "entra contact list-member-of";
  const showHint = `mg-axi entra contact show-member-of --id <contact-id> --member-id <membership-id>${transitive ? " --transitive" : ""} ${profileHint(profileName)}`;
  const scopeHint = transitive
    ? `Direct memberships only: mg-axi ${command} --id ${shellValue(id)} ${profileHint(profileName)}`
    : `Flat nested view: mg-axi ${command} --id ${shellValue(id)} --transitive ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(command, effectiveFlags, profileName)] : [];
  const limitedHints = limitedInfo > 0
    ? [`${limitedInfo} of ${memberOf.length} rows carry only type and id; this may reflect limited read consent or unset properties`]
    : [];
  if (!result.complete) {
    return {
      memberOf,
      count: { returned: memberOf.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, ...limitedHints, resumeHint(profileName), scopeHint, showHint],
    };
  }
  const count = { returned: memberOf.length, complete: true };
  if (!memberOf.length) {
    return {
      memberOf,
      count,
      help: [...limitedHints, `0 ${noun} matched; the absence of results is the answer, not an error`, scopeHint, showHint],
    };
  }
  return { memberOf, count, help: [...truncationHints, ...limitedHints, scopeHint, showHint] };
}

export async function showContactMemberOf(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const { select, fields } = navFieldsFor(flags, help);
  const scopes = membershipScopesFor(flags, profile, help);
  const full = flags.full === true;
  const id = contactId(flags, help);
  const memberId = membershipId(flags, help);
  const raw = await withGuidance(membershipDenialHints(flags), () => session.execute({
    profile,
    operation,
    params: { "orgContact-id": id, "directoryObject-id": memberId },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = navSingleResult(raw, fields, full, "member-of");
  if (truncated) return { memberOf: row, help: [fullHint("entra contact show-member-of", flags, profileName)] };
  return { memberOf: row };
}

// The $count routes return a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute and accepts only
// a non-negative integer. The scalars carry no operation-level
// documentation page and follow the parent list contract (ConsistencyLevel
// eventual). There is no --select/--limit contract on the count: strict
// input validation refuses those flags before credentials.
export async function countContactMemberOf(
  session: GraphSession,
  flags: ContactFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  checkReadOperation(operation, help);
  const scopes = membershipScopesFor(flags, profile, help);
  const id = contactId(flags, help);
  const raw = await withGuidance(membershipDenialHints(flags), () => session.execute({
    profile, operation, params: { "orgContact-id": id }, scopes, consistencyLevel: "eventual", scalar: true,
  }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed member-of count body", "GRAPH_ERROR", [
      "Member-of counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  const count = { returned: raw, complete: true };
  return raw === 0
    ? { count, help: ["0 memberships returned; a contact outside every group and administrative unit carries none, so the absence of rows is the answer, not an error"] }
    : { count };
}
