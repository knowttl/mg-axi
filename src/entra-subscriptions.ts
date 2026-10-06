import { AxiError } from "axi-sdk-js";
import { resolveSessionOperation, type CollectArgs, type GraphSession, type SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 directory-subscriptions subfamily: the read mapping behind
// `mg-axi entra subscription list/show/count`. Operation construction stays
// beside its command; the shared session owns URLs, credentials, paging,
// retries and error translation, and the SDK owns TOON rendering. This
// module only maps flags to session calls and projects rows for compact
// output.
//
// Reviewed against the v1.0 directory-list-subscriptions,
// companysubscription-get and companySubscription resource documentation on
// 2026-10-05. Every read takes D/A Organization.Read.All; delegated callers
// additionally need a supported Entra role (Global Reader, Directory Readers,
// or Dynamics 365 Business Central Administrator for read-only standard
// properties). Personal Microsoft accounts are not supported. No P1/P2
// prerequisite is stated for these reads. The list documents the general
// OData query parameters, so --filter passes through as plain $filter with
// no $count or ConsistencyLevel contract; the single GET documents $select
// only. The $count route carries no operation-level documentation page and
// takes no --filter/--select/--limit/--cursor. The commerceSubscriptionId
// alternate-key lookup ships on `show --commerce-subscription-id` through
// the allowlisted session function binding: the key arrives only as an
// explicit CLI flag and the session validates, OData-quotes and encodes it.
// Beta subscriptions stay out ("beta needs its own review"). No mutation
// lives here.
//
// Secrecy by construction: the reviewed companySubscription set carries
// directory identifiers and licence metadata only; ownerId and ownerTenantId
// name the account admin and partner tenant without carrying credentials or
// keys, and the shared session still redacts secret-shaped values before
// buffering or output.

// Every subscription property this slice may request or display, matching
// the reviewed resource. Anything else fails before credentials.
export const KNOWN_SUBSCRIPTION_FIELDS: readonly string[] = [
  "id",
  "commerceSubscriptionId",
  "createdDateTime",
  "isTrial",
  "nextLifecycleDateTime",
  "ownerId",
  "ownerTenantId",
  "ownerType",
  "serviceStatus",
  "skuId",
  "skuPartNumber",
  "status",
  "totalLicenses",
];
const KNOWN_SUBSCRIPTIONS = new Set(KNOWN_SUBSCRIPTION_FIELDS);

// Compact list rows: identifier, SKU name, lifecycle status and the licence
// count agents usually need next.
const DEFAULT_LIST_SELECT = ["id", "skuPartNumber", "status", "totalLicenses"];
// Show rows: the full reviewed subscription set.
const DEFAULT_SHOW_SELECT = [...KNOWN_SUBSCRIPTION_FIELDS];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_SUBSCRIPTION_SCOPES = ["https://graph.microsoft.com/Organization.Read.All"];
const TRUNCATE_AT = 500;

export type SubscriptionFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, known: Set<string>, knownList: readonly string[], flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown subscription property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known subscription properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: SubscriptionFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
  flags: SubscriptionFlags,
  defaults: string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, KNOWN_SUBSCRIPTIONS, KNOWN_SUBSCRIPTION_FIELDS, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, KNOWN_SUBSCRIPTIONS, KNOWN_SUBSCRIPTION_FIELDS, "fields", help);
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

function fullHint(command: string, flags: SubscriptionFlags, profileName: string): string {
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

const SUBSCRIPTION_DENIAL_HINTS = [
  "Subscription reads need Organization.Read.All plus a supported Entra role for delegated access (Global Reader, Directory Readers, or Dynamics 365 Business Central Administrator for read-only standard properties), or admin-consented Organization.Read.All for application access",
  "Personal Microsoft accounts are not supported for subscription reads",
  "No P1/P2 prerequisite is stated for subscription reads; never diagnose licence solely from HTTP 403",
];

const EMPTY_NOTE = "An empty subscription list may mean the tenant holds no commercial subscriptions; the absence of rows is the answer, not an error";

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
  filter: string | undefined;
}

// Restoring the saved select/filter keeps cursor resumes lossless when
// --select or --filter is omitted. The list documents the general OData
// query parameters, so --filter passes through as plain $filter with no
// ConsistencyLevel contract; strict input validation still refuses unknown
// flags before credentials.
function collectionCommon(
  session: GraphSession,
  flags: SubscriptionFlags,
  operation: SessionOperation,
  help: string,
  profile: AnyProfile,
): CollectionCommon {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    saved?.["$select"] === undefined ? DEFAULT_LIST_SELECT : fieldList(saved["$select"], KNOWN_SUBSCRIPTIONS, KNOWN_SUBSCRIPTION_FIELDS, "select", help), help);
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  return { cursor, select, fields, scopes: scopesFor(flags, DEFAULT_SUBSCRIPTION_SCOPES, profile, help), full: flags.full === true, filter };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: SubscriptionFlags,
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

function subscriptionId(flags: SubscriptionFlags, help: string): string {
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the commercial-subscription object ID", "VALIDATION_ERROR", [help]);
  return id;
}

// Exactly one subscription key selects the show route: the object ID from
// the catalogue-resolved operation, or the commerce-system alternate key
// from the allowlisted function route. The key value itself is validated by
// the session binding before credentials.
function subscriptionTarget(flags: SubscriptionFlags, operation: SessionOperation, help: string): { operation: SessionOperation; params: Record<string, string> } {
  const id = flags.id;
  const commerceId = flags["commerce-subscription-id"];
  if (id !== undefined && commerceId !== undefined) {
    throw new AxiError("--id and --commerce-subscription-id cannot be combined", "VALIDATION_ERROR", [help]);
  }
  if (commerceId !== undefined) {
    if (!String(commerceId).trim()) throw new AxiError("--commerce-subscription-id needs the commerce-system subscription ID", "VALIDATION_ERROR", [help]);
    return {
      operation: resolveSessionOperation("v1.0", "GET", "/directory/subscriptions(commerceSubscriptionId='{commerceSubscriptionId}')"),
      params: { commerceSubscriptionId: String(commerceId) },
    };
  }
  if (id === undefined) throw new AxiError("Subscription show needs exactly one of --id or --commerce-subscription-id", "VALIDATION_ERROR", [help]);
  return { operation, params: { "companySubscription-id": subscriptionId(flags, help) } };
}

function singleResult(
  raw: unknown,
  fields: string[],
  full: boolean,
): { row: Record<string, unknown>; truncated: boolean } {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed subscription body", "GRAPH_ERROR", [
      "Single-subscription reads carry one companySubscription object; treat anything else as unknown, not empty",
    ]);
  }
  return project(raw, fields, full);
}

export async function listSubscriptions(
  session: GraphSession,
  flags: SubscriptionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, operation, help, profile);
  const result = await withGuidance(SUBSCRIPTION_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: SubscriptionFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const subscriptions: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    subscriptions.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra subscription show --id <subscription-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra subscription list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      subscriptions,
      ...listTotals(subscriptions.length, result.total, "subscriptions", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint],
    };
  }
  if (!subscriptions.length) {
    const emptyHint = common.filter === undefined ? EMPTY_NOTE : "0 subscriptions matched; the absence of results is the answer, not an error";
    return { subscriptions, ...listTotals(subscriptions.length, result.total, "subscriptions", true), complete: true, help: [emptyHint] };
  }
  return { subscriptions, ...listTotals(subscriptions.length, result.total, "subscriptions", true), complete: true, help: [...truncationHints, showHint] };
}

export async function showSubscription(
  session: GraphSession,
  flags: SubscriptionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, DEFAULT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_SUBSCRIPTION_SCOPES, profile, help);
  const full = flags.full === true;
  const target = subscriptionTarget(flags, operation, help);
  const raw = await withGuidance(SUBSCRIPTION_DENIAL_HINTS, () => session.execute({
    profile,
    operation: target.operation,
    params: target.params,
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full);
  if (truncated) return { subscription: row, help: [fullHint("entra subscription show", flags, profileName)] };
  return { subscription: row };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute with scalar mode
// and accepts only a non-negative integer. The $count route carries no
// operation-level documentation page and takes no --filter/--select/--limit/--cursor.
export async function countSubscriptions(
  session: GraphSession,
  flags: SubscriptionFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_SUBSCRIPTION_SCOPES, profile, help);
  const raw = await withGuidance(SUBSCRIPTION_DENIAL_HINTS, () => session.execute({ profile, operation, scopes, scalar: true }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed subscription count body", "GRAPH_ERROR", [
      "Subscription counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
    ]);
  }
  const count = { returned: raw, complete: true };
  return raw === 0 ? { count, help: [EMPTY_NOTE] } : { count };
}
