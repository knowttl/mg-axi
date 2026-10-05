import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 contacts subfamily (first part): the read mapping behind
// `mg-axi entra contact list/show/count`. Operation construction stays
// beside its command; the shared session owns URLs, credentials, paging,
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
// so the scalar does too. The delta() sync, every per-contact navigation
// read (manager, directReports, memberOf, transitiveMemberOf,
// serviceProvisioningErrors, onPremisesSyncBehavior), every POST lookup
// action, beta and every mutation stay out. No mutation lives here.
//
// Secrecy by construction: contacts are personal data, so default rows carry
// only id, displayName, mail and companyName; identifying fields beyond that
// need an explicit --select. Only flat scalar properties are ever requested
// or projected; $expand is not offered, so navigation objects can never
// appear, and the nested phones/addresses/error collections need their own
// projection review (the error collections have dedicated sub-reads deferred
// to the navigation part).

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

// The only operations this slice ever binds. Anything else - delta sync,
// per-contact navigation, POST lookups, writes - is refused before
// credentials.
const READ_OPERATIONS: Readonly<Record<string, string>> = {
  "GET:/contacts": "contact",
  "GET:/contacts/{orgContact-id}": "contact",
  "GET:/contacts/$count": "contact",
};

function checkReadOperation(operation: SessionOperation, help: string): void {
  const route = `${operation.method}:${operation.path}`;
  if (operation.method !== "GET" || !Object.hasOwn(READ_OPERATIONS, route)) {
    throw new AxiError(`Refused non-read route ${operation.id}: contact serves only the three catalogued v1.0 contacts GETs`, "VALIDATION_ERROR", [
      help,
      "Delta sync, per-contact navigation, POST lookup actions and writes are never constructed here; beta needs its own review",
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
