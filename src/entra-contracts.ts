import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-04 partner contracts subfamily: the read mapping behind `mg-axi entra
// contract list/show/count`. Operation construction stays beside its command;
// the shared session owns URLs, credentials, paging, retries and error
// translation, and the SDK owns TOON rendering. This module only maps flags
// to session calls and projects rows for compact output.
//
// Reviewed against the v1.0 contract-list, contract-get and contract
// resource documentation on 2026-10-04. Every read takes D/A
// Directory.Read.All; delegated callers additionally need a supported Entra
// role (Directory Readers is the only least-privileged role supported for
// these operations). Personal Microsoft accounts are not supported. No
// P1/P2 prerequisite is stated for these reads. Contracts exist in partner
// tenants only (Cloud Solution Provider, Office 365 Syndication or Advisor
// programs); a non-partner tenant lists zero contracts, which is an answer
// rather than an error. The list documents OData query parameters with
// filtering for customerId, defaultDomainName and displayName, so --filter
// passes through as plain $filter with no $count or ConsistencyLevel
// contract. The $count route returns a text/plain integer scalar; the leaf
// reads it through session.execute and accepts only a non-negative integer,
// so no --filter/--select/--limit/--cursor contract exists on the count.
// contracts/delta() stays out: delta-token sync needs its own paging and
// change-tracking contract. Beta contracts and the POST lookup actions
// belong to later EXT-04 subfamilies. No contract mutation exists in this
// slice.

// Every contract property this slice may request or display, matching the
// reviewed resource. Anything else fails before credentials.
export const KNOWN_CONTRACT_FIELDS: readonly string[] = [
  "contractType",
  "customerId",
  "defaultDomainName",
  "displayName",
  "id",
];
const KNOWN_CONTRACTS = new Set(KNOWN_CONTRACT_FIELDS);

// Compact contract rows: the partnership id, the customer name and domain
// copies that name the customer, and the contract kind that decides the
// partner relationship.
const DEFAULT_CONTRACT_LIST_SELECT = ["id", "displayName", "contractType", "defaultDomainName"];
// Show rows: the full reviewed contract set.
const DEFAULT_CONTRACT_SHOW_SELECT = [...KNOWN_CONTRACT_FIELDS];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_CONTRACT_SCOPES = ["https://graph.microsoft.com/Directory.Read.All"];
const TRUNCATE_AT = 500;

export type ContractFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown contract property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: ContractFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
  flags: ContractFlags,
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

function fullHint(command: string, flags: ContractFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the scope, roles and partner prerequisite that actually
// unlock it, because a 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const CONTRACT_DENIAL_HINTS = [
  "Contract reads need Directory.Read.All plus a supported Entra role for delegated access (Directory Readers is the least-privileged role), or admin-consented Directory.Read.All for application access",
  "Contracts exist in partner tenants only (Cloud Solution Provider, Office 365 Syndication or Advisor programs); a non-partner tenant reads empty, which is an answer rather than an error",
  "Personal Microsoft accounts are not supported for contract reads",
  "No P1/P2 prerequisite is stated for contract reads; never diagnose licence solely from HTTP 403",
];

const PARTNER_NOTE = "Contracts exist in partner tenants only; a non-partner tenant lists zero contracts, which is an answer rather than an error";

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
  filter: string | undefined;
}

// Restoring the saved select/filter keeps cursor resumes lossless when
// --select or --filter is omitted. The contract list passes --filter through
// as plain $filter with no $count or ConsistencyLevel contract.
function collectionCommon(
  session: GraphSession,
  flags: ContractFlags,
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
  flags: ContractFlags,
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

function contractId(flags: ContractFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the contract partnership identifier", "VALIDATION_ERROR", [help]);
  return id;
}

export async function listContracts(
  session: GraphSession,
  flags: ContractFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_CONTRACTS, KNOWN_CONTRACT_FIELDS,
    DEFAULT_CONTRACT_LIST_SELECT, DEFAULT_CONTRACT_SCOPES, operation, help, profile);
  const result = await withGuidance(CONTRACT_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: ContractFlags = { ...flags, select: result.query.$select ?? DEFAULT_CONTRACT_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const contracts: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    contracts.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra contract show --id <contract-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra contract list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      contracts,
      ...listTotals(contracts.length, result.total, "contracts", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, PARTNER_NOTE],
    };
  }
  if (!contracts.length) {
    return {
      contracts,
      ...listTotals(contracts.length, result.total, "contracts", true),
      complete: true,
      help: ["0 contracts matched; the absence of results is the answer, not an error", PARTNER_NOTE],
    };
  }
  return { contracts, ...listTotals(contracts.length, result.total, "contracts", true), complete: true, help: [...truncationHints, showHint, PARTNER_NOTE] };
}

export async function showContract(
  session: GraphSession,
  flags: ContractFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_CONTRACTS, KNOWN_CONTRACT_FIELDS, DEFAULT_CONTRACT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_CONTRACT_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(CONTRACT_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "contract-id": contractId(flags, help) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed contract body", "GRAPH_ERROR", [
      "Single-contract reads carry one contract object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { contract: row, help: [fullHint("entra contract show", flags, profileName)] };
  return { contract: row };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute and accepts only
// a non-negative integer. There is no --filter/--select/--limit contract on
// the count: the catalogue declares no such flags and strict input
// validation refuses them before credentials.
export async function countContracts(
  session: GraphSession,
  flags: ContractFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_CONTRACT_SCOPES, profile, help);
  const raw = await withGuidance(CONTRACT_DENIAL_HINTS, () => session.execute({ profile, operation, scopes }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed contract count body", "GRAPH_ERROR", [
      "Contract counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  return { contractCount: raw };
}
