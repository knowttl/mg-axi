import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-03 workforce identity-providers subfamily: the read mapping behind
// `mg-axi entra identity-provider list/show/count/available-types`. Operation construction
// stays beside its command; the shared session owns URLs, credentials,
// paging, retries and error translation, and the SDK owns TOON rendering.
// This module only maps flags to session calls and projects rows for compact
// output.
//
// Reviewed against the v1.0 identitycontainer-list-identityproviders,
// identityproviderbase-get and identityproviderbase-availableprovidertypes
// operation documentation cited in inventory/operations.json on 2026-10-04.
// List, get and count take D/A IdentityProvider.Read.All; delegated callers
// additionally need a directory role that can read federation configuration
// (Global Reader is the least-privileged read-only directory role).
// Personal Microsoft accounts are not supported. No per-operation licence
// prerequisite is stated for these reads; available provider types follow
// the tenant's external-identities configuration. Workforce context only:
// external-customer (B2C/External ID) user flows, authentication event
// flows, invitations, cross-tenant access, trust framework and provisioning
// are separate later EXT-03 pieces and no support for them is claimed here.
// No mutation lives here.
//
// Secrecy by construction: clientSecret (write-only on social providers)
// and certificateData (key material on Apple providers) are not selectable,
// so --select/--fields naming them fails before credentials, and rows the
// server returns carrying either key are scrubbed before projection. Secrets
// therefore can never reach stdout, errors, previews or logs through this
// slice. @odata.type is preserved on every row without being selectable: it
// names the provider kind (for example
// #microsoft.graph.socialIdentityProvider).

// Every provider property this slice may request or display, matching the
// reviewed raw surface. Anything else fails before credentials. The base
// properties ride on every row; the remaining properties are returned by the
// server only on rows of that provider kind.
export const KNOWN_PROVIDER_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "identityProviderType",
  "clientId",
];
const KNOWN_PROVIDERS = new Set(KNOWN_PROVIDER_FIELDS);
// @odata.type is preserved on provider rows without being selectable: it
// names the derived provider kind.
const PROVIDER_TYPE_PROPERTY = "@odata.type";
// Secret-bearing keys are never selectable and never projected. The server
// is not expected to return them on GET, but any row carrying them is
// scrubbed before projection so key material can never reach output.
const SECRET_PROVIDER_FIELDS: readonly string[] = ["clientSecret", "certificateData"];

// Compact rows: identifier and name; the kind rides as @odata.type.
const DEFAULT_LIST_SELECT = ["id", "displayName"];
// Show rows: the full reviewed provider set.
const DEFAULT_SHOW_SELECT = [...KNOWN_PROVIDER_FIELDS];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_PROVIDER_SCOPES = ["https://graph.microsoft.com/IdentityProvider.Read.All"];
const TRUNCATE_AT = 500;

export type IdentityProviderFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!KNOWN_PROVIDERS.has(part)) {
      const secret = SECRET_PROVIDER_FIELDS.includes(part);
      throw new AxiError(secret
        ? `--${flag} cannot request ${part}: secret-bearing provider fields are never returned`
        : `Unknown provider property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known provider properties: ${KNOWN_PROVIDER_FIELDS.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: IdentityProviderFlags, profile: AnyProfile, help: string): string[] | undefined {
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
  if (flags.scopes === undefined) return [...DEFAULT_PROVIDER_SCOPES];
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
    if (SECRET_PROVIDER_FIELDS.includes(key)) continue;
    cleaned[key] = value;
  }
  return cleaned;
}

// Local projection preserves Graph's null/missing distinction: an explicit
// null stays null, an absent property stays absent and is never synthesized.
// @odata.type survives projection so the provider kind is never lost.
function project(
  row: unknown,
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
  if (typeof source[PROVIDER_TYPE_PROPERTY] === "string") projected[PROVIDER_TYPE_PROPERTY] = source[PROVIDER_TYPE_PROPERTY];
  return { row: projected, truncated };
}

function selectedFields(
  flags: IdentityProviderFlags,
  defaults: string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, "select", help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, "fields", help);
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

function fullHint(command: string, flags: IdentityProviderFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; provider
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

const PROVIDER_DENIAL_HINTS = [
  "Identity-provider reads need IdentityProvider.Read.All: delegated callers pass it as --scopes, application access needs admin-consented IdentityProvider.Read.All on the configured .default audience",
  "Delegated callers additionally need a directory role that can read federation configuration (Global Reader is the least-privileged read-only directory role)",
  "Personal Microsoft accounts are not supported for identity-provider reads",
  "No per-operation licence prerequisite is stated for these reads; never diagnose licence solely from HTTP 403",
];

const WORKFORCE_NOTE = "Workforce tenant context only; external-customer (B2C/External ID) user flows are separate scheduled operations";

export async function listIdentityProviders(
  session: GraphSession,
  flags: IdentityProviderFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const { select, fields } = selectedFields(flags,
    saved?.["$select"] === undefined ? DEFAULT_LIST_SELECT : fieldList(saved["$select"], "select", help), help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const query: Record<string, string> = { $select: select.join(",") };
  const filter = flags.filter === undefined ? saved?.["$filter"] : String(flags.filter);
  if (filter !== undefined) query.$filter = filter;
  const args: CollectArgs = { profile, operation, query, scopes };
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(PROVIDER_DENIAL_HINTS, () => session.collect(args));
  const effectiveFlags: IdentityProviderFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const providers: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, full);
    providers.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra identity-provider show --id <provider-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra identity-provider list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      identityProviders: providers,
      ...listTotals(providers.length, result.total, "identity providers", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor <cursor-from-output> ${profileHint(profileName)}`, showHint, WORKFORCE_NOTE],
    };
  }
  if (!providers.length) {
    return {
      identityProviders: providers,
      ...listTotals(providers.length, result.total, "identity providers", true),
      complete: true,
      help: [WORKFORCE_NOTE, "0 identity providers matched; the absence of results is the answer, not an error"],
    };
  }
  return { identityProviders: providers, ...listTotals(providers.length, result.total, "identity providers", true), complete: true, help: [...truncationHints, showHint, WORKFORCE_NOTE] };
}

export async function showIdentityProvider(
  session: GraphSession,
  flags: IdentityProviderFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, DEFAULT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(PROVIDER_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "identityProviderBase-id": String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed identity-provider body", "GRAPH_ERROR", [
      "Single-provider reads carry one provider object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { identityProvider: row, help: [fullHint("entra identity-provider show", flags, profileName), WORKFORCE_NOTE] };
  return { identityProvider: row, help: [WORKFORCE_NOTE] };
}

export async function countIdentityProviders(
  session: GraphSession,
  flags: IdentityProviderFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const scopes = scopesFor(flags, profile, help);
  const query: Record<string, string> = {};
  if (flags.filter !== undefined) query.$filter = String(flags.filter);
  const raw = await withGuidance(PROVIDER_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: {},
    query,
    scopes,
    scalar: true,
  }));
  if (typeof raw !== "number" || !Number.isSafeInteger(raw)) {
    throw new AxiError("Graph returned a malformed identity-provider count body", "GRAPH_ERROR", [
      "Scalar count reads carry a plain number; treat anything else as unknown, not empty",
    ]);
  }
  if (flags.filter !== undefined) {
    return {
      count: { returned: raw, complete: true },
      help: [`Count reflects --filter ${shellValue(String(flags.filter))}; drop --filter for the tenant total`, WORKFORCE_NOTE],
    };
  }
  return { count: { returned: raw, complete: true }, help: [WORKFORCE_NOTE] };
}

export async function availableIdentityProviderTypes(
  session: GraphSession,
  flags: IdentityProviderFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
): Promise<Record<string, unknown>> {
  const scopes = scopesFor(flags, profile, help);
  const raw = await withGuidance(PROVIDER_DENIAL_HINTS, () => session.execute({ profile, operation, scopes }));
  const value = raw !== null && typeof raw === "object" && !Array.isArray(raw)
    ? (raw as Record<string, unknown>).value : undefined;
  if (!Array.isArray(value) || value.some(type => typeof type !== "string")) {
    throw new AxiError("Graph returned a malformed available identity-provider types body", "GRAPH_ERROR", [
      "Available-provider-type reads carry a value array of strings; treat anything else as unknown, not empty",
    ]);
  }
  return {
    availableProviderTypes: value,
    count: { returned: value.length, complete: true },
    help: ["Available provider types depend on tenant configuration and licensing; availability does not mean a provider is configured", WORKFORCE_NOTE],
  };
}
