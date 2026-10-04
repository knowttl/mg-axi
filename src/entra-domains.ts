import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 domains subfamily: the read mapping behind `mg-axi entra domain
// list/show`, the per-domain verification and service-configuration DNS
// record reads, and the top-level domain DNS record reads. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 domain-list, domain-get,
// domain-list-verificationdnsrecords and
// domain-list-serviceconfigurationrecords operation documentation on
// 2026-10-04. Every read takes D/A Domain.Read.All; delegated callers
// additionally need a supported Entra role (Domain Name Administrator or
// Global Reader are the least-privileged for the DNS record reads, and the
// domain list/get support those plus Directory Readers and other directory
// roles). Personal Microsoft accounts are not supported. No P1/P2
// prerequisite is stated for these reads. The single-record GETs and the
// top-level /domainDnsRecords GETs carry no operation-level documentation
// page; they read the same domainDnsRecord resource through the same
// permission contract, so the review cites the parent list documentation
// alongside the domainDnsRecord resource reference. Domain ids are fully
// qualified names (for example contoso.com), not object UUIDs. Graph
// documents a known issue with $search, $top and $filter on domain lists,
// so `domain list` offers no --filter; the DNS record collections document
// plain OData parameters and --filter passes through as plain $filter with
// no $count or ConsistencyLevel contract. No domain mutation exists in this
// slice: verify, forceDelete, promote, federation and shared-email surfaces
// belong to later pieces.

// Every domain property this slice may request or display, matching the
// reviewed raw surface. Anything else fails before credentials.
export const KNOWN_DOMAIN_FIELDS: readonly string[] = [
  "id",
  "authenticationType",
  "availabilityStatus",
  "isAdminManaged",
  "isDefault",
  "isInitial",
  "isRoot",
  "isVerified",
  "supportedServices",
  "passwordValidityPeriodInDays",
  "passwordNotificationWindowInDays",
  "state",
];
const KNOWN_DOMAINS = new Set(KNOWN_DOMAIN_FIELDS);

// Every DNS record property this slice may request or display. The base
// domainDnsRecord properties stay selectable on every row; the derived-type
// properties (Mx mailExchange/preference, CNAME canonicalName, SRV
// nameTarget/port/priority/protocol/service/weight, TXT text) are returned
// by the server only on rows of that record kind.
export const KNOWN_DNS_FIELDS: readonly string[] = [
  "id",
  "isOptional",
  "label",
  "recordType",
  "supportedService",
  "ttl",
  "mailExchange",
  "preference",
  "canonicalName",
  "nameTarget",
  "port",
  "priority",
  "protocol",
  "service",
  "weight",
  "text",
];
const KNOWN_DNS = new Set(KNOWN_DNS_FIELDS);
// @odata.type is preserved on DNS rows without being selectable: it names
// the derived record kind (for example #microsoft.graph.domainDnsMxRecord).
const DNS_TYPE_PROPERTY = "@odata.type";

// Compact domain rows: the fully qualified name, authentication kind and the
// verification/default state that decides the next step.
const DEFAULT_DOMAIN_LIST_SELECT = ["id", "authenticationType", "isVerified", "isDefault"];
// Show rows: the full reviewed domain set.
const DEFAULT_DOMAIN_SHOW_SELECT = [...KNOWN_DOMAIN_FIELDS];
// Compact DNS rows: identifier, zone label, record kind and service.
const DEFAULT_DNS_LIST_SELECT = ["id", "label", "recordType", "supportedService"];
// Show rows: the full reviewed base set; derived-type detail needs an
// explicit --select naming the derived property.
const DEFAULT_DNS_SHOW_SELECT = ["id", "isOptional", "label", "recordType", "supportedService", "ttl"];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_DOMAIN_SCOPES = ["https://graph.microsoft.com/Domain.Read.All"];
export const DEFAULT_DNS_SCOPES = ["https://graph.microsoft.com/Domain.Read.All"];
const TRUNCATE_AT = 500;

export type DomainFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown domain property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: DomainFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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

// DNS projection additionally preserves @odata.type so the derived record
// kind survives local projection.
function projectDns(
  row: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  const { row: projected, truncated } = project(row, fields, full);
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  if (typeof source[DNS_TYPE_PROPERTY] === "string") projected[DNS_TYPE_PROPERTY] = source[DNS_TYPE_PROPERTY];
  return { row: projected, truncated };
}

function selectedFields(
  flags: DomainFlags,
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

function fullHint(command: string, flags: DomainFlags, profileName: string): string {
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

const DOMAIN_DENIAL_HINTS = [
  "Domain reads need Domain.Read.All plus a supported Entra role for delegated access (Domain Name Administrator or Global Reader are least-privileged; Directory Readers and other directory roles are also supported for list/get), or admin-consented Domain.Read.All for application access",
  "Personal Microsoft accounts are not supported for domain reads",
  "No P1/P2 prerequisite is stated for domain reads; never diagnose licence solely from HTTP 403",
];

const DNS_DENIAL_HINTS = [
  "DNS record reads need Domain.Read.All plus a supported Entra role for delegated access (Domain Name Administrator or Global Reader are least-privileged), or admin-consented Domain.Read.All for application access",
  "Personal Microsoft accounts are not supported for DNS record reads",
  "No P1/P2 prerequisite is stated for DNS record reads; never diagnose licence solely from HTTP 403",
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
// --select or --filter is omitted. Domain lists carry no --filter flag at
// all: Graph has a documented known issue with $search, $top and $filter on
// /domains, so strict input validation refuses the flag before credentials
// instead of forwarding a misleading request.
function collectionCommon(
  session: GraphSession,
  flags: DomainFlags,
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
  flags: DomainFlags,
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

function domainName(flags: DomainFlags, help: string): string {
  const domain = String(flags.domain);
  if (!domain.trim()) throw new AxiError("--domain needs the fully qualified domain name", "VALIDATION_ERROR", [help]);
  return domain;
}

export async function listDomains(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_DOMAINS, KNOWN_DOMAIN_FIELDS,
    DEFAULT_DOMAIN_LIST_SELECT, DEFAULT_DOMAIN_SCOPES, operation, help, profile);
  const result = await withGuidance(DOMAIN_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: DomainFlags = { ...flags, select: result.query.$select ?? DEFAULT_DOMAIN_LIST_SELECT.join(",") };
  const domains: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    domains.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra domain show --id <domain-name> ${profileHint(profileName)}`;
  const filterNote = "Domain lists offer no --filter: Graph documents a known issue with $search, $top and $filter on /domains";
  const truncationHints = truncated ? [fullHint("entra domain list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      domains,
      count: { returned: domains.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, filterNote],
    };
  }
  const count = { returned: domains.length, complete: true };
  if (!domains.length) {
    return {
      domains,
      count,
      help: [filterNote, "0 domains matched; the absence of results is the answer, not an error"],
    };
  }
  return { domains, count, help: [...truncationHints, showHint, filterNote] };
}

export async function showDomain(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_DOMAINS, KNOWN_DOMAIN_FIELDS, DEFAULT_DOMAIN_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_DOMAIN_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(DOMAIN_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "domain-id": String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed domain body", "GRAPH_ERROR", [
      "Single-domain reads carry one domain object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  const helpHints: string[] = truncated ? [fullHint("entra domain show", flags, profileName)] : [];
  if (row["isVerified"] === false) {
    helpHints.push(`mg-axi entra domain verification-dns-record list --domain ${shellValue(String(flags.id))} ${profileHint(profileName)}`);
  }
  if (helpHints.length) return { domain: row, help: helpHints };
  return { domain: row };
}

interface DnsCommand {
  collectionKey: string;
  recordKey: string;
  command: string;
  showCommand: string;
  needsDomain: boolean;
}

function dnsParams(flags: DomainFlags, help: string, needsDomain: boolean): Record<string, string> {
  if (!needsDomain) return {};
  return { "domain-id": domainName(flags, help) };
}

async function listDnsRecords(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  command: DnsCommand,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_DNS, KNOWN_DNS_FIELDS,
    DEFAULT_DNS_LIST_SELECT, DEFAULT_DNS_SCOPES, operation, help, profile);
  const args = collectArgs(profile, operation, common, flags, help);
  args.params = dnsParams(flags, help, command.needsDomain);
  const result = await withGuidance(DNS_DENIAL_HINTS, () => session.collect(args));
  const effectiveFlags: DomainFlags = { ...flags, select: result.query.$select ?? DEFAULT_DNS_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const records: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = projectDns(row, common.fields, common.full);
    records.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const domainFlag = command.needsDomain ? `--domain ${shellValue(String(flags.domain))} ` : "";
  const showHint = `mg-axi ${command.showCommand} ${domainFlag}--id <record-id> ${profileHint(profileName)}`;
  const kindHint = "Rows carry @odata.type naming the derived record kind; derived-type detail needs an explicit --select naming the derived property";
  const truncationHints = truncated ? [fullHint(command.command, effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      [command.collectionKey]: records,
      count: { returned: records.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, kindHint],
    };
  }
  const count = { returned: records.length, complete: true };
  if (!records.length) {
    return {
      [command.collectionKey]: records,
      count,
      help: [kindHint, "0 DNS records matched; the absence of results is the answer, not an error"],
    };
  }
  return { [command.collectionKey]: records, count, help: [...truncationHints, showHint, kindHint] };
}

async function showDnsRecord(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  command: DnsCommand,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_DNS, KNOWN_DNS_FIELDS, DEFAULT_DNS_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_DNS_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(DNS_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { ...dnsParams(flags, help, command.needsDomain), "domainDnsRecord-id": String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed DNS record body", "GRAPH_ERROR", [
      "Single-record reads carry one domainDnsRecord object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = projectDns(raw, fields, full);
  const kindHint = "Rows carry @odata.type naming the derived record kind; derived-type detail needs an explicit --select naming the derived property";
  if (truncated) return { [command.recordKey]: row, help: [fullHint(command.showCommand, flags, profileName), kindHint] };
  return { [command.recordKey]: row, help: [kindHint] };
}

const VERIFICATION: DnsCommand = {
  collectionKey: "verificationDnsRecords",
  recordKey: "verificationDnsRecord",
  command: "entra domain verification-dns-record list",
  showCommand: "entra domain verification-dns-record show",
  needsDomain: true,
};

const SERVICE: DnsCommand = {
  collectionKey: "serviceConfigurationRecords",
  recordKey: "serviceConfigurationRecord",
  command: "entra domain service-configuration-record list",
  showCommand: "entra domain service-configuration-record show",
  needsDomain: true,
};

const TOP_LEVEL: DnsCommand = {
  collectionKey: "domainDnsRecords",
  recordKey: "domainDnsRecord",
  command: "entra domain-dns-record list",
  showCommand: "entra domain-dns-record show",
  needsDomain: false,
};

export function listVerificationDnsRecords(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listDnsRecords(session, flags, profile, operation, help, profileName, VERIFICATION);
}

export function showVerificationDnsRecord(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showDnsRecord(session, flags, profile, operation, help, profileName, VERIFICATION);
}

export function listServiceConfigurationRecords(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listDnsRecords(session, flags, profile, operation, help, profileName, SERVICE);
}

export function showServiceConfigurationRecord(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showDnsRecord(session, flags, profile, operation, help, profileName, SERVICE);
}

export function listDomainDnsRecords(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listDnsRecords(session, flags, profile, operation, help, profileName, TOP_LEVEL);
}

export function showDomainDnsRecord(
  session: GraphSession,
  flags: DomainFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showDnsRecord(session, flags, profile, operation, help, profileName, TOP_LEVEL);
}
