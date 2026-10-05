import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 certificate-auth subfamily: the read mapping behind
// `mg-axi entra certificate-auth-pki list/show/count` and
// `mg-axi entra certificate-authority list/show/count`. Operation
// construction stays beside its command; the shared session owns URLs,
// credentials, paging, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0
// publickeyinfrastructureroot-list-certificatebasedauthconfigurations,
// certificatebasedauthpki-get,
// certificatebasedauthpki-list-certificateauthorities and
// certificateauthoritydetail-get operation documentation and the
// certificateBasedAuthPki and certificateAuthorityDetail resource references
// on 2026-10-05. Every read takes D/A PublicKeyInfrastructure.Read.All;
// delegated callers additionally need Privileged Authentication Administrator
// or Authentication Administrator. Personal Microsoft accounts are not
// supported. No P1/P2 prerequisite is stated for these reads. The PKI list
// documents $filter (eq, startswith), $top, $orderby and $count, so --filter
// passes through as plain $filter with no ConsistencyLevel contract; the
// single GETs document $select only. The root /certificateBasedAuthConfiguration
// reads carry no documented v1.0 operation contract (the documented
// certificateBasedAuthConfiguration contract is org-scoped) and stay out,
// as do the org-scoped certificate-auth reads, beta routes, the upload
// action and every mutation. No mutation lives here.
//
// Secrecy by construction: certificate-authority entries carry public
// certificates only, but the base64 `certificate` blob (up to 8 KB per CA
// file) is omitted from every default select and returned only on an
// explicit --select naming it, still under the truncation marker with a
// --full escape hatch. No private key material exists on these resources.

// Every PKI property this slice may request or display, matching the
// reviewed resource. Anything else fails before credentials.
export const KNOWN_PKI_FIELDS: readonly string[] = [
  "id",
  "deletedDateTime",
  "displayName",
  "status",
  "statusDetails",
  "lastModifiedDateTime",
];
const KNOWN_PKIS = new Set(KNOWN_PKI_FIELDS);

// Every certificate-authority property this slice may request or display,
// matching the reviewed resource. `certificate` is the public CA key: it is
// selectable only through an explicit --select and never rides a default.
export const KNOWN_CA_FIELDS: readonly string[] = [
  "id",
  "deletedDateTime",
  "certificateAuthorityType",
  "certificate",
  "displayName",
  "issuer",
  "issuerSubjectKeyIdentifier",
  "createdDateTime",
  "expirationDateTime",
  "thumbprint",
  "certificateRevocationListUrl",
  "deltacertificateRevocationListUrl",
  "isIssuerHintEnabled",
];
const KNOWN_CAS = new Set(KNOWN_CA_FIELDS);

// Compact PKI rows: identifier, name and the async upload/delete status
// that decides whether the authorities below are ready.
const DEFAULT_PKI_LIST_SELECT = ["id", "displayName", "status"];
// Show rows: the full reviewed PKI set.
const DEFAULT_PKI_SHOW_SELECT = [...KNOWN_PKI_FIELDS];
// Compact CA rows: identifier, name, kind and expiry; the thumbprint and
// issuer name the authority on the show view.
const DEFAULT_CA_LIST_SELECT = ["id", "displayName", "certificateAuthorityType", "expirationDateTime"];
// Show rows: the full reviewed CA set except the public-certificate blob,
// which needs an explicit --select naming it.
const DEFAULT_CA_SHOW_SELECT = KNOWN_CA_FIELDS.filter(field => field !== "certificate");
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_CERT_AUTH_SCOPES = ["https://graph.microsoft.com/PublicKeyInfrastructure.Read.All"];
const TRUNCATE_AT = 500;

export type CertificateAuthFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string, noun: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown certificate-auth property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known ${noun} properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: CertificateAuthFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
  flags: CertificateAuthFlags,
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

function fullHint(command: string, flags: CertificateAuthFlags, profileName: string): string {
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

const CERT_AUTH_DENIAL_HINTS = [
  "Certificate-auth reads need PublicKeyInfrastructure.Read.All plus a supported Entra role for delegated access (Privileged Authentication Administrator or Authentication Administrator), or admin-consented PublicKeyInfrastructure.Read.All for application access",
  "Personal Microsoft accounts are not supported for certificate-auth reads",
  "No P1/P2 prerequisite is stated for certificate-auth reads; never diagnose licence solely from HTTP 403",
];

const EMPTY_PKI_NOTE = "An empty PKI list may mean certificate-based authentication is not configured in the tenant; the absence of rows is the answer, not an error";
const CERT_BLOB_NOTE = "CA public-certificate blobs are omitted by default; request certificate explicitly with --select certificate";

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
  filter: string | undefined;
}

// Restoring the saved select/filter keeps cursor resumes lossless when
// --select or --filter is omitted. Both collections document $filter, so
// --filter passes through as plain $filter with no ConsistencyLevel
// contract; strict input validation still refuses unknown flags before
// credentials.
function collectionCommon(
  session: GraphSession,
  flags: CertificateAuthFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  noun: string,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    known, knownList,
    saved?.["$select"] === undefined ? defaults : fieldList(saved["$select"], known, knownList, "select", help, noun), help, noun);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_CERT_AUTH_SCOPES, profile, help), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: CertificateAuthFlags,
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

function pkiId(flags: CertificateAuthFlags, help: string, flag: string): string {
  const id = String(flags[flag]);
  if (!id.trim()) throw new AxiError(`--${flag} needs the certificate-auth PKI object ID`, "VALIDATION_ERROR", [help]);
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

export async function listCertificateAuthPkis(
  session: GraphSession,
  flags: CertificateAuthFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_PKIS, KNOWN_PKI_FIELDS,
    DEFAULT_PKI_LIST_SELECT, "PKI", operation, help, profile);
  const result = await withGuidance(CERT_AUTH_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: CertificateAuthFlags = { ...flags, select: result.query.$select ?? DEFAULT_PKI_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const pkis: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    pkis.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra certificate-auth-pki show --id <pki-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra certificate-auth-pki list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      certificateAuthPkis: pkis,
      count: { returned: pkis.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, EMPTY_PKI_NOTE],
    };
  }
  const count = { returned: pkis.length, complete: true };
  if (!pkis.length) {
    return { certificateAuthPkis: pkis, count, help: [EMPTY_PKI_NOTE] };
  }
  return { certificateAuthPkis: pkis, count, help: [...truncationHints, showHint, EMPTY_PKI_NOTE] };
}

export async function showCertificateAuthPki(
  session: GraphSession,
  flags: CertificateAuthFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_PKIS, KNOWN_PKI_FIELDS, DEFAULT_PKI_SHOW_SELECT, help, "PKI");
  const scopes = scopesFor(flags, DEFAULT_CERT_AUTH_SCOPES, profile, help);
  const full = flags.full === true;
  const id = pkiId(flags, help, "id");
  const raw = await withGuidance(CERT_AUTH_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "certificateBasedAuthPki-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "certificate-auth PKI");
  const helpHints: string[] = [
    ...(truncated ? [fullHint("entra certificate-auth-pki show", flags, profileName)] : []),
    `mg-axi entra certificate-authority list --pki ${shellValue(id)} ${profileHint(profileName)}`,
  ];
  return { certificateAuthPki: row, help: helpHints };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute with scalar mode
// and accepts only a non-negative integer. --filter narrows the count
// server-side; there is no --select/--limit/--cursor contract on the count.
export async function countCertificateAuthPkis(
  session: GraphSession,
  flags: CertificateAuthFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_CERT_AUTH_SCOPES, profile, help);
  const query: Record<string, string> = {};
  if (flags.filter !== undefined) query.$filter = String(flags.filter);
  const raw = await withGuidance(CERT_AUTH_DENIAL_HINTS, () => session.execute({ profile, operation, query, scopes, scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed certificate-auth PKI count body", "GRAPH_ERROR", [
      "PKI counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  if (flags.filter !== undefined) {
    return {
      count: { returned: raw, complete: true },
      help: [`Count reflects --filter ${shellValue(String(flags.filter))}; drop --filter for the tenant total`, EMPTY_PKI_NOTE],
    };
  }
  return { count: { returned: raw, complete: true }, help: [EMPTY_PKI_NOTE] };
}

function authorityId(flags: CertificateAuthFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the certificate-authority object ID", "VALIDATION_ERROR", [help]);
  return id;
}

function authorityParams(flags: CertificateAuthFlags, help: string): Record<string, string> {
  return { "certificateBasedAuthPki-id": pkiId(flags, help, "pki") };
}

export async function listCertificateAuthorities(
  session: GraphSession,
  flags: CertificateAuthFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_CAS, KNOWN_CA_FIELDS,
    DEFAULT_CA_LIST_SELECT, "certificate-authority", operation, help, profile);
  const pki = pkiId(flags, help, "pki");
  const result = await withGuidance(CERT_AUTH_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help, authorityParams(flags, help))));
  const effectiveFlags: CertificateAuthFlags = { ...flags, select: result.query.$select ?? DEFAULT_CA_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const authorities: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    authorities.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra certificate-authority show --pki ${shellValue(pki)} --id <authority-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra certificate-authority list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      certificateAuthorities: authorities,
      count: { returned: authorities.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, CERT_BLOB_NOTE],
    };
  }
  const count = { returned: authorities.length, complete: true };
  if (!authorities.length) {
    return {
      certificateAuthorities: authorities,
      count,
      help: ["0 certificate authorities matched; the absence of results is the answer, not an error", CERT_BLOB_NOTE],
    };
  }
  return { certificateAuthorities: authorities, count, help: [...truncationHints, showHint, CERT_BLOB_NOTE] };
}

export async function showCertificateAuthority(
  session: GraphSession,
  flags: CertificateAuthFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_CAS, KNOWN_CA_FIELDS, DEFAULT_CA_SHOW_SELECT, help, "certificate-authority");
  const scopes = scopesFor(flags, DEFAULT_CERT_AUTH_SCOPES, profile, help);
  const full = flags.full === true;
  const params = { ...authorityParams(flags, help), "certificateAuthorityDetail-id": authorityId(flags, help) };
  const raw = await withGuidance(CERT_AUTH_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params,
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "certificate authority");
  const helpHints: string[] = [...(truncated ? [fullHint("entra certificate-authority show", flags, profileName)] : []), CERT_BLOB_NOTE];
  return { certificateAuthority: row, help: helpHints };
}

export async function countCertificateAuthorities(
  session: GraphSession,
  flags: CertificateAuthFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_CERT_AUTH_SCOPES, profile, help);
  const query: Record<string, string> = {};
  if (flags.filter !== undefined) query.$filter = String(flags.filter);
  const raw = await withGuidance(CERT_AUTH_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: authorityParams(flags, help),
    query,
    scopes,
    scalar: true,
  }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed certificate-authority count body", "GRAPH_ERROR", [
      "Authority counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  if (flags.filter !== undefined) {
    return {
      count: { returned: raw, complete: true },
      help: [`Count reflects --filter ${shellValue(String(flags.filter))}; drop --filter for the PKI total`, CERT_BLOB_NOTE],
    };
  }
  return { count: { returned: raw, complete: true }, help: [CERT_BLOB_NOTE] };
}
