import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-03 risk-prevention subfamily: the read mapping behind
// `mg-axi entra fraud-protection-provider list/show/count`,
// `mg-axi entra web-application-firewall-provider list/show/count` and
// `mg-axi entra web-application-firewall-verification list/show/count`
// plus `... verification provider show`. Operation construction stays
// beside its command; the shared session owns URLs, credentials, paging,
// retries and error translation, and the SDK owns TOON rendering. This
// module only maps flags to session calls and projects rows for compact
// output.
//
// Reviewed against the v1.0 riskpreventioncontainer-list-fraudprotectionproviders,
// fraudprotectionprovider-get,
// riskpreventioncontainer-list-webapplicationfirewallproviders,
// webapplicationfirewallprovider-get,
// riskpreventioncontainer-list-webapplicationfirewallverifications and
// webapplicationfirewallverificationmodel-get operation documentation and
// the fraudProtectionProvider, arkoseFraudProtectionProvider,
// humanSecurityFraudProtectionProvider, webApplicationFirewallProvider,
// akamaiWebApplicationFirewallProvider,
// cloudFlareWebApplicationFirewallProvider and
// webApplicationFirewallVerificationModel resource references on
// 2026-10-06. All six operation pages document the same access contract,
// so one scope and one denial-hint set serve all ten reads. Every read
// takes delegated RiskPreventionProviders.Read.All
// (RiskPreventionProviders.ReadWrite.All is the documented
// higher-privileged alternative); Graph documents no supported
// application permission, so application profiles are refused before
// credentials rather than sent to a certain denial. Delegated callers
// additionally need a directory role that can read risk-prevention
// configuration (Security Reader is the least-privileged read-only role;
// External ID User Flow Administrator, Application Administrator,
// Directory Reader and Security Administrator also work). Personal
// Microsoft accounts are not supported. No per-operation licence
// prerequisite is stated for these reads. The fraud list documents
// $select only, so it offers no --filter; the WAF and verification
// lists document $filter and offer it as plain $filter with no $count
// or ConsistencyLevel contract. $expand stays out everywhere: the
// verification's provider has its own sub-read and file/detail
// navigations need their own review. The three $count scalars take no
// --filter/--select/--limit/--cursor. The provider sub-read has no
// dedicated operation page; access follows the WAF provider-get
// contract and the verificationModel resource reference. Beta stays out
// ("beta needs its own review") and no mutation lives here: provider
// create/update/delete and verification submission (POST
// .../webApplicationFirewallProviders/{id}/verify) belong to no read
// slice.
//
// Secrecy by construction: API keys, client secrets and other key
// material the provider shapes carry (Arkose privateKey and publicKey,
// HUMAN serverToken, Akamai clientSecret/clientToken/accessToken,
// Cloudflare apiToken) are not selectable, so --select/--fields naming
// them fails before credentials, and rows the server returns carrying
// any of them are scrubbed before projection. Secrets therefore can
// never reach stdout, errors, previews or logs through this slice.
// publicKey is key material even though it is public: it follows the
// identity-provider certificateData precedent and is never returned.
// @odata.type is preserved on fraud and WAF rows without being
// selectable: it names the provider kind (for example
// #microsoft.graph.arkoseFraudProtectionProvider).

// Every fraud-protection-provider property this slice may request or
// display: the base properties plus the non-secret derived properties
// (Arkose clientSubDomain/verifySubDomain, HUMAN appId). Anything else
// fails before credentials.
export const KNOWN_FRAUD_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "clientSubDomain",
  "verifySubDomain",
  "appId",
];
const KNOWN_FRAUD = new Set(KNOWN_FRAUD_FIELDS);

// Every WAF-provider property this slice may request or display: the
// base properties plus the non-secret derived properties (Akamai
// hostPrefix, Cloudflare zoneId). Anything else fails before
// credentials.
export const KNOWN_WAF_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "hostPrefix",
  "zoneId",
];
const KNOWN_WAF = new Set(KNOWN_WAF_FIELDS);

// Every verification property this slice may request or display,
// matching the reviewed resource. verificationResult and
// verifiedDetails ride whole as complex objects. Anything else fails
// before credentials.
export const KNOWN_VERIFICATION_FIELDS: readonly string[] = [
  "id",
  "verifiedHost",
  "providerType",
  "verificationResult",
  "verifiedDetails",
];
const KNOWN_VERIFICATIONS = new Set(KNOWN_VERIFICATION_FIELDS);

// Secret-bearing keys are never selectable and never projected. The
// server is not expected to return them on GET, but any row carrying
// them is scrubbed before projection so key material can never reach
// output.
const SECRET_RISK_FIELDS: readonly string[] = [
  "privateKey",
  "publicKey",
  "serverToken",
  "clientSecret",
  "clientToken",
  "accessToken",
  "apiToken",
];
// @odata.type is preserved on fraud and WAF rows without being
// selectable: it names the derived provider kind.
const RISK_TYPE_PROPERTY = "@odata.type";

// Compact rows: identifier and name; the kind rides as @odata.type.
const DEFAULT_FRAUD_LIST_SELECT = ["id", "displayName"];
// Show rows: the full reviewed set.
const DEFAULT_FRAUD_SHOW_SELECT = [...KNOWN_FRAUD_FIELDS];
// Compact rows: identifier and name; the kind rides as @odata.type.
const DEFAULT_WAF_LIST_SELECT = ["id", "displayName"];
// Show rows: the full reviewed set.
const DEFAULT_WAF_SHOW_SELECT = [...KNOWN_WAF_FIELDS];
// Compact rows: identifier, verified host and provider type.
const DEFAULT_VERIFICATION_LIST_SELECT = ["id", "verifiedHost", "providerType"];
// Show rows: the full reviewed set.
const DEFAULT_VERIFICATION_SHOW_SELECT = [...KNOWN_VERIFICATION_FIELDS];
// Delegated-only reads: Graph documents no supported application
// permission, so application profiles are refused before credentials.
export const DEFAULT_RISK_SCOPES = ["https://graph.microsoft.com/RiskPreventionProviders.Read.All"];
const TRUNCATE_AT = 500;

export type RiskPreventionFlags = Record<string, string | boolean>;

interface RiskKind {
  noun: string;
  totalsNoun: string;
  known: Set<string>;
  knownList: readonly string[];
  listSelect: string[];
  showSelect: string[];
  collectionKey: string;
  singleKey: string;
  keepType: boolean;
}

const FRAUD_KIND: RiskKind = {
  noun: "fraud-protection-provider",
  totalsNoun: "fraud protection providers",
  known: KNOWN_FRAUD,
  knownList: KNOWN_FRAUD_FIELDS,
  listSelect: DEFAULT_FRAUD_LIST_SELECT,
  showSelect: DEFAULT_FRAUD_SHOW_SELECT,
  collectionKey: "fraudProtectionProviders",
  singleKey: "fraudProtectionProvider",
  keepType: true,
};

const WAF_KIND: RiskKind = {
  noun: "web-application-firewall-provider",
  totalsNoun: "web application firewall providers",
  known: KNOWN_WAF,
  knownList: KNOWN_WAF_FIELDS,
  listSelect: DEFAULT_WAF_LIST_SELECT,
  showSelect: DEFAULT_WAF_SHOW_SELECT,
  collectionKey: "webApplicationFirewallProviders",
  singleKey: "webApplicationFirewallProvider",
  keepType: true,
};

const VERIFICATION_KIND: RiskKind = {
  noun: "web-application-firewall-verification",
  totalsNoun: "web application firewall verifications",
  known: KNOWN_VERIFICATIONS,
  knownList: KNOWN_VERIFICATION_FIELDS,
  listSelect: DEFAULT_VERIFICATION_LIST_SELECT,
  showSelect: DEFAULT_VERIFICATION_SHOW_SELECT,
  collectionKey: "webApplicationFirewallVerifications",
  singleKey: "webApplicationFirewallVerification",
  keepType: false,
};

function fieldList(raw: unknown, kind: RiskKind, flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!kind.known.has(part)) {
      const secret = SECRET_RISK_FIELDS.includes(part);
      throw new AxiError(secret
        ? `--${flag} cannot request ${part}: secret-bearing provider fields are never returned`
        : `Unknown ${kind.noun} property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known ${kind.noun} properties: ${kind.knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: RiskPreventionFlags, profile: AnyProfile, help: string, noun: string): string[] {
  if (profile.mode === "application") {
    throw new AxiError(
      `${noun} reads need a delegated profile; Graph documents no supported application permission for this operation`,
      "VALIDATION_ERROR",
      [help, "mg-axi profile create --name <name> --tenant <tenant-id> --client <client-id> --cloud commercial"],
    );
  }
  if (flags.scopes === undefined) return [...DEFAULT_RISK_SCOPES];
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

// Scrubbing runs before projection: any secret-bearing key the server
// returned is dropped so it can never be selected, rendered or leaked.
function scrub(row: unknown): Record<string, unknown> {
  const source = row !== null && typeof row === "object" && !Array.isArray(row) ? (row as Record<string, unknown>) : {};
  const cleaned: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(source)) {
    if (SECRET_RISK_FIELDS.includes(key)) continue;
    cleaned[key] = value;
  }
  return cleaned;
}

// Local projection preserves Graph's null/missing distinction: an
// explicit null stays null, an absent property stays absent and is never
// synthesized. @odata.type survives projection on provider rows so the
// provider kind is never lost.
function project(
  row: unknown,
  kind: RiskKind,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  const source = scrub(row);
  const projected: Record<string, unknown> = {};
  let truncated = false;
  for (const field of fields) {
    if (!Object.hasOwn(source, field)) continue;
    const result = truncateValue(source[field], full);
    projected[field] = result.value;
    truncated = truncated || result.truncated;
  }
  if (kind.keepType && typeof source[RISK_TYPE_PROPERTY] === "string") projected[RISK_TYPE_PROPERTY] = source[RISK_TYPE_PROPERTY];
  return { row: projected, truncated };
}

function selectedFields(
  flags: RiskPreventionFlags,
  kind: RiskKind,
  defaults: string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, kind, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, kind, "fields", help);
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

function fullHint(command: string, flags: RiskPreventionFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; these
// reads add the scope, roles and licensing that actually unlock them,
// because a 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const RISK_DENIAL_HINTS = [
  "Risk-prevention reads need delegated RiskPreventionProviders.Read.All (RiskPreventionProviders.ReadWrite.All is the documented higher-privileged alternative); delegated callers pass it as --scopes",
  "Delegated reads additionally need a directory role that can read risk-prevention configuration (Security Reader is the least-privileged read-only role; External ID User Flow Administrator, Application Administrator, Directory Reader and Security Administrator also work); Graph documents no supported application permission, so application profiles are refused before credentials",
  "Personal Microsoft accounts are not supported for risk-prevention reads",
  "No per-operation licence prerequisite is stated for risk-prevention reads; never diagnose licence solely from HTTP 403",
];

const WORKFORCE_NOTE = "Workforce tenant context only; verification submission and provider create/update/delete belong to no read slice and external-customer user flows, authentication event flows and provisioning stay scheduled";

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[];
  full: boolean;
  filter: string | undefined;
}

// Restoring the saved select/filter keeps cursor resumes lossless when
// --select or --filter is omitted. The fraud list documents $select
// only, so its leaf declares no --filter flag and strict input
// validation refuses the flag before credentials; the WAF and
// verification lists document $filter, so --filter passes through as
// plain $filter with no $count or ConsistencyLevel contract.
function collectionCommon(
  session: GraphSession,
  flags: RiskPreventionFlags,
  kind: RiskKind,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags, kind,
    saved?.["$select"] === undefined ? kind.listSelect : fieldList(saved["$select"], kind, "select", help), help);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, scopes: scopesFor(flags, profile, help, "Risk-prevention"), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: RiskPreventionFlags,
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

function resourceId(flags: RiskPreventionFlags, help: string, noun: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError(`--id needs the ${noun} ID`, "VALIDATION_ERROR", [help]);
  return id;
}

function singleResult(
  raw: unknown,
  kind: RiskKind,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError(`Graph returned a malformed ${kind.noun} body`, "GRAPH_ERROR", [
      `Single-${kind.noun} reads carry one ${kind.singleKey} object; treat anything else as unknown, not empty`,
    ]);
  }
  return project(raw, kind, fields, full);
}

async function collectKind(
  session: GraphSession,
  flags: RiskPreventionFlags,
  kind: RiskKind,
  command: string,
  showCommand: string,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, kind, operation, help, profile);
  const result = await withGuidance(RISK_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: RiskPreventionFlags = { ...flags, select: result.query.$select ?? kind.listSelect.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, kind, common.fields, common.full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi ${showCommand} --id <${kind.noun}-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(command, effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      [kind.collectionKey]: rows,
      ...listTotals(rows.length, result.total, kind.totalsNoun, false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, WORKFORCE_NOTE],
    };
  }
  if (!rows.length) {
    return {
      [kind.collectionKey]: rows,
      ...listTotals(rows.length, result.total, kind.totalsNoun, true),
      complete: true,
      help: [`0 ${kind.noun}s matched; the absence of results is the answer, not an error`, WORKFORCE_NOTE],
    };
  }
  return { [kind.collectionKey]: rows, ...listTotals(rows.length, result.total, kind.totalsNoun, true), complete: true, help: [...truncationHints, showHint, WORKFORCE_NOTE] };
}

async function showKind(
  session: GraphSession,
  flags: RiskPreventionFlags,
  kind: RiskKind,
  command: string,
  param: string,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, kind, kind.showSelect, help);
  const scopes = scopesFor(flags, profile, help, "Risk-prevention");
  const full = flags.full === true;
  const id = resourceId(flags, help, kind.noun);
  const raw = await withGuidance(RISK_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { [param]: id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, kind, fields, full);
  if (truncated) return { [kind.singleKey]: row, help: [fullHint(command, flags, profileName), WORKFORCE_NOTE] };
  return { [kind.singleKey]: row, help: [WORKFORCE_NOTE] };
}

async function countKind(
  session: GraphSession,
  flags: RiskPreventionFlags,
  kind: RiskKind,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
): Promise<Record<string, unknown>> {
  const scopes = scopesFor(flags, profile, help, "Risk-prevention");
  const raw = await withGuidance(RISK_DENIAL_HINTS, () => session.execute({ profile, operation, scopes, scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError(`Graph returned a malformed ${kind.noun} count body`, "GRAPH_ERROR", [
      `${kind.noun} counts carry one non-negative integer scalar; treat anything else as unknown, not empty`,
    ]);
  }
  return { count: { returned: raw, complete: true }, help: [WORKFORCE_NOTE] };
}

export async function listFraudProtectionProviders(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return collectKind(session, flags, FRAUD_KIND,
    "entra fraud-protection-provider list", "entra fraud-protection-provider show", profile, operation, help, profileName);
}

export async function showFraudProtectionProvider(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showKind(session, flags, FRAUD_KIND,
    "entra fraud-protection-provider show", "fraudProtectionProvider-id", profile, operation, help, profileName);
}

export async function countFraudProtectionProviders(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
): Promise<Record<string, unknown>> {
  return countKind(session, flags, FRAUD_KIND, profile, operation, help);
}

export async function listWafProviders(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return collectKind(session, flags, WAF_KIND,
    "entra web-application-firewall-provider list", "entra web-application-firewall-provider show", profile, operation, help, profileName);
}

export async function showWafProvider(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showKind(session, flags, WAF_KIND,
    "entra web-application-firewall-provider show", "webApplicationFirewallProvider-id", profile, operation, help, profileName);
}

export async function countWafProviders(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
): Promise<Record<string, unknown>> {
  return countKind(session, flags, WAF_KIND, profile, operation, help);
}

export async function listWafVerifications(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return collectKind(session, flags, VERIFICATION_KIND,
    "entra web-application-firewall-verification list", "entra web-application-firewall-verification show", profile, operation, help, profileName);
}

export async function showWafVerification(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showKind(session, flags, VERIFICATION_KIND,
    "entra web-application-firewall-verification show", "webApplicationFirewallVerificationModel-id", profile, operation, help, profileName);
}

export async function countWafVerifications(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
): Promise<Record<string, unknown>> {
  return countKind(session, flags, VERIFICATION_KIND, profile, operation, help);
}

export async function showVerificationProvider(
  session: GraphSession,
  flags: RiskPreventionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, WAF_KIND, WAF_KIND.showSelect, help);
  const scopes = scopesFor(flags, profile, help, "Risk-prevention");
  const full = flags.full === true;
  const id = resourceId(flags, help, VERIFICATION_KIND.noun);
  const raw = await withGuidance(RISK_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "webApplicationFirewallVerificationModel-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, WAF_KIND, fields, full);
  if (truncated) {
    return {
      [WAF_KIND.singleKey]: row,
      help: [fullHint("entra web-application-firewall-verification provider show", flags, profileName), WORKFORCE_NOTE],
    };
  }
  return { [WAF_KIND.singleKey]: row, help: [WORKFORCE_NOTE] };
}
