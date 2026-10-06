import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 terms-of-use subfamily: the read mapping behind
// `mg-axi entra agreement list/show`,
// `mg-axi entra agreement acceptance list/show --agreement <agreement-id>`
// and `mg-axi entra agreement-acceptance list/show`. Operation construction
// stays beside its command; the shared session owns URLs, credentials,
// paging, retries and error translation, and the SDK owns TOON rendering.
// This module only maps flags to session calls and projects rows for compact
// output.
//
// Reviewed against the v1.0 termsofusecontainer-list-agreements,
// agreement-get and agreement-list-acceptances operation documentation and
// the agreement and agreementAcceptance resource references on 2026-10-05,
// corroborated by the Microsoft Entra terms-of-use feature documentation
// (same P1 prerequisite, same Security Reader read role). The legacy
// /agreements and /agreementAcceptances roots share the termsOfUse
// operation contract documented on the /identityGovernance/termsOfUse
// routes; the root acceptance list and get carry no dedicated operation
// page, so they follow the reviewed acceptance-list contract and the
// agreementAcceptance resource reference. Agreement reads take delegated
// Agreement.Read.All; acceptance reads take delegated
// AgreementAcceptance.Read least-privileged (AgreementAcceptance.Read.All
// is the documented higher-privileged alternative). Graph documents no
// supported application permission for any of these operations, so
// application profiles are refused before credentials rather than sent to
// a certain denial. Delegated callers additionally need Security Reader,
// the least-privileged supported role (Global Reader, Conditional Access
// Administrator and Security Administrator also work). Personal Microsoft
// accounts are not supported. Terms of use needs Microsoft Entra ID P1.
// The agreement list documents $select, $filter and $top; the agreement
// get documents $select only, so it offers no --filter. The nested
// acceptance list documents $select and $filter; the root acceptance list
// follows the same reviewed contract. --filter passes through as plain
// $filter with no $count or ConsistencyLevel contract. The acceptances
// $count scalar, every agreement file/localization/files/versions
// sub-read, beta and every mutation stay out. No mutation lives here.
//
// Secrecy by construction: agreement file and files navigation properties
// are never requested ($expand is not reviewed here) and file bytes are
// never downloaded or printed; only agreement metadata is projected.
// Acceptance records are personal data, so default rows stay minimal (id,
// agreementId, state, recordedDateTime) and the identifying fields
// (userDisplayName, userEmail, userPrincipalName, userId and device
// detail) need an explicit --select naming them.

// Every agreement property this slice may request or display, matching the
// reviewed resource. The acceptances, file and files relationships are
// never projected here. Anything else fails before credentials.
export const KNOWN_AGREEMENT_FIELDS: readonly string[] = [
  "displayName",
  "id",
  "isPerDeviceAcceptanceRequired",
  "isViewingBeforeAcceptanceRequired",
  "termsExpiration",
  "userReacceptRequiredFrequency",
];
const KNOWN_AGREEMENTS = new Set(KNOWN_AGREEMENT_FIELDS);

// Every acceptance property this slice may request or display, matching
// the reviewed resource. Anything else fails before credentials.
export const KNOWN_ACCEPTANCE_FIELDS: readonly string[] = [
  "agreementFileId",
  "agreementId",
  "deviceDisplayName",
  "deviceId",
  "deviceOSType",
  "deviceOSVersion",
  "expirationDateTime",
  "id",
  "recordedDateTime",
  "state",
  "userDisplayName",
  "userEmail",
  "userId",
  "userPrincipalName",
];
const KNOWN_ACCEPTANCES = new Set(KNOWN_ACCEPTANCE_FIELDS);

// Compact agreement rows: the identifier plus the internal tracking name
// agents need next.
const DEFAULT_AGREEMENT_LIST_SELECT = ["id", "displayName"];
// Show rows: the full reviewed agreement metadata set.
const DEFAULT_AGREEMENT_SHOW_SELECT = [...KNOWN_AGREEMENT_FIELDS];
// Acceptance rows stay minimal by default because they are personal data:
// the identifier, the owning agreement, the current state and when it was
// recorded. Identifying fields need an explicit --select.
const DEFAULT_ACCEPTANCE_SELECT = ["id", "agreementId", "state", "recordedDateTime"];
// Delegated defaults are operation-specific; application profiles are
// refused before credentials because Graph documents no supported
// application permission here.
export const DEFAULT_AGREEMENT_SCOPES = ["https://graph.microsoft.com/Agreement.Read.All"];
export const DEFAULT_ACCEPTANCE_SCOPES = ["https://graph.microsoft.com/AgreementAcceptance.Read"];
const TRUNCATE_AT = 500;

export type TermsOfUseFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string, noun: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown terms-of-use property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known ${noun} properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: TermsOfUseFlags, defaults: string[], profile: AnyProfile, help: string, noun: string): string[] {
  if (profile.mode === "application") {
    throw new AxiError(
      `${noun} reads need a delegated profile; Graph documents no supported application permission for this operation`,
      "VALIDATION_ERROR",
      [help, "mg-axi profile create --name <name> --tenant <tenant-id> --client <client-id> --cloud commercial"],
    );
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
  flags: TermsOfUseFlags,
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

function fullHint(command: string, flags: TermsOfUseFlags, profileName: string): string {
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

const AGREEMENT_DENIAL_HINTS = [
  "Agreement reads need delegated Agreement.Read.All; delegated callers pass it as --scopes",
  "Delegated reads additionally need Security Reader, the least-privileged supported Entra role for this operation (Global Reader, Conditional Access Administrator or Security Administrator also work); Graph documents no supported application permission, so application profiles are refused before credentials",
  "Personal Microsoft accounts are not supported for agreement reads",
  "Terms of use needs Microsoft Entra ID P1; never diagnose licence solely from HTTP 403",
];

const ACCEPTANCE_DENIAL_HINTS = [
  "Acceptance reads need delegated AgreementAcceptance.Read least-privileged (AgreementAcceptance.Read.All is the documented higher-privileged alternative); delegated callers pass one as --scopes",
  "Delegated reads additionally need Security Reader, the least-privileged supported Entra role for this operation (Global Reader, Conditional Access Administrator or Security Administrator also work); Graph documents no supported application permission, so application profiles are refused before credentials",
  "Personal Microsoft accounts are not supported for acceptance reads",
  "Terms of use needs Microsoft Entra ID P1; never diagnose licence solely from HTTP 403",
];

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[];
  full: boolean;
  filter: string | undefined;
}

// Restoring the saved select/filter keeps cursor resumes lossless when
// --select or --filter is omitted. Both lists document $filter, so
// --filter passes through as plain $filter with no $count or
// ConsistencyLevel contract; strict input validation refuses the flag on
// the show leaves before credentials instead of forwarding a misleading
// request.
function collectionCommon(
  session: GraphSession,
  flags: TermsOfUseFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  noun: string,
  defaultScopes: string[],
  denialNoun: string,
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
  return { cursor, select, fields, scopes: scopesFor(flags, defaultScopes, profile, help, denialNoun), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: TermsOfUseFlags,
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

function agreementId(flags: TermsOfUseFlags, help: string, flag = "id"): string {
  const id = String(flags[flag]);
  if (!id.trim()) throw new AxiError(`--${flag} needs the agreement identifier`, "VALIDATION_ERROR", [help]);
  return id;
}

function acceptanceId(flags: TermsOfUseFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the agreement-acceptance identifier", "VALIDATION_ERROR", [help]);
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

export async function listAgreements(
  session: GraphSession,
  flags: TermsOfUseFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_AGREEMENTS, KNOWN_AGREEMENT_FIELDS,
    DEFAULT_AGREEMENT_LIST_SELECT, "agreement", DEFAULT_AGREEMENT_SCOPES, "Agreement", operation, help, profile);
  const result = await withGuidance(AGREEMENT_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: TermsOfUseFlags = { ...flags, select: result.query.$select ?? DEFAULT_AGREEMENT_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const agreements: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    agreements.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra agreement show --id <agreement-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra agreement list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      agreements,
      ...listTotals(agreements.length, result.total, "agreements", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  if (!agreements.length) {
    return {
      agreements,
      ...listTotals(agreements.length, result.total, "agreements", true),
      complete: true,
      help: ["0 agreements matched; no terms-of-use agreements are configured for the tenant, or no row passed --filter"],
    };
  }
  return { agreements, ...listTotals(agreements.length, result.total, "agreements", true), complete: true, help: [...truncationHints, showHint] };
}

export async function showAgreement(
  session: GraphSession,
  flags: TermsOfUseFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_AGREEMENTS, KNOWN_AGREEMENT_FIELDS, DEFAULT_AGREEMENT_SHOW_SELECT, help, "agreement");
  const scopes = scopesFor(flags, DEFAULT_AGREEMENT_SCOPES, profile, help, "Agreement");
  const full = flags.full === true;
  const id = agreementId(flags, help);
  const raw = await withGuidance(AGREEMENT_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "agreement-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "agreement");
  const helpHints: string[] = [
    ...(truncated ? [fullHint("entra agreement show", flags, profileName)] : []),
    `mg-axi entra agreement acceptance list --agreement ${shellValue(id)} ${profileHint(profileName)}`,
  ];
  return { agreement: row, help: helpHints };
}

function acceptanceCollection(
  session: GraphSession,
  flags: TermsOfUseFlags,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
  agreement: string | undefined,
): { common: CollectionCommon; params: Record<string, string> | undefined } {
  const common = collectionCommon(session, flags, KNOWN_ACCEPTANCES, KNOWN_ACCEPTANCE_FIELDS,
    DEFAULT_ACCEPTANCE_SELECT, "agreement-acceptance", DEFAULT_ACCEPTANCE_SCOPES, "Acceptance", operation, help, profile);
  if (agreement === undefined) return { common, params: undefined };
  if (!agreement.trim()) throw new AxiError("--agreement needs the agreement identifier", "VALIDATION_ERROR", [help]);
  return { common, params: { "agreement-id": agreement } };
}

async function collectAcceptances(
  session: GraphSession,
  flags: TermsOfUseFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  command: string,
  showCommand: string,
  agreement: string | undefined,
): Promise<Record<string, unknown>> {
  const { common, params } = acceptanceCollection(session, flags, operation, help, profile, agreement);
  const result = await withGuidance(ACCEPTANCE_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help, params)));
  const effectiveFlags: TermsOfUseFlags = { ...flags, select: result.query.$select ?? DEFAULT_ACCEPTANCE_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const acceptances: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    acceptances.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const scope = agreement === undefined ? "" : `--agreement ${shellValue(agreement)} `;
  const showHint = `mg-axi ${showCommand} ${scope}--id <acceptance-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(command, effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      agreementAcceptances: acceptances,
      ...listTotals(acceptances.length, result.total, "agreement acceptances", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  if (!acceptances.length) {
    return {
      agreementAcceptances: acceptances,
      ...listTotals(acceptances.length, result.total, "agreement acceptances", true),
      complete: true,
      help: ["0 agreement acceptances matched; nobody recorded an acceptance in scope yet, or no row passed --filter"],
    };
  }
  return { agreementAcceptances: acceptances, ...listTotals(acceptances.length, result.total, "agreement acceptances", true), complete: true, help: [...truncationHints, showHint] };
}

async function showAcceptanceById(
  session: GraphSession,
  flags: TermsOfUseFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
  command: string,
  agreement: string | undefined,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_ACCEPTANCES, KNOWN_ACCEPTANCE_FIELDS, DEFAULT_ACCEPTANCE_SELECT, help, "agreement-acceptance");
  const scopes = scopesFor(flags, DEFAULT_ACCEPTANCE_SCOPES, profile, help, "Acceptance");
  const full = flags.full === true;
  const id = acceptanceId(flags, help);
  const params: Record<string, string> = agreement === undefined
    ? { "agreementAcceptance-id": id }
    : { "agreement-id": agreement, "agreementAcceptance-id": id };
  const raw = await withGuidance(ACCEPTANCE_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params,
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "agreement acceptance");
  if (truncated) return { agreementAcceptance: row, help: [fullHint(command, flags, profileName)] };
  return { agreementAcceptance: row };
}

export async function listAgreementAcceptances(
  session: GraphSession,
  flags: TermsOfUseFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const agreement = String(flags.agreement);
  if (!agreement.trim()) throw new AxiError("--agreement needs the agreement identifier", "VALIDATION_ERROR", [help]);
  return collectAcceptances(session, flags, profile, operation, help, profileName, "entra agreement acceptance list", "entra agreement acceptance show", agreement);
}

export async function showAgreementAcceptance(
  session: GraphSession,
  flags: TermsOfUseFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const agreement = String(flags.agreement);
  if (!agreement.trim()) throw new AxiError("--agreement needs the agreement identifier", "VALIDATION_ERROR", [help]);
  return showAcceptanceById(session, flags, profile, operation, help, profileName, "entra agreement acceptance show", agreement);
}

export async function listAcceptances(
  session: GraphSession,
  flags: TermsOfUseFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return collectAcceptances(session, flags, profile, operation, help, profileName, "entra agreement-acceptance list", "entra agreement-acceptance show", undefined);
}

export async function showAcceptance(
  session: GraphSession,
  flags: TermsOfUseFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return showAcceptanceById(session, flags, profile, operation, help, profileName, "entra agreement-acceptance show", undefined);
}
