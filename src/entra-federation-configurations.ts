import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 directory federation-configurations subfamily: the read mapping
// behind `mg-axi entra federation-configuration list/show/count/available-types`.
// Operation construction stays beside its command; the shared session owns
// URLs, credentials, paging, retries and error translation, and the SDK owns
// TOON rendering. This module only maps flags to session calls and projects
// rows for compact output.
//
// Reviewed against the v1.0 samlorwsfedexternaldomainfederation-list,
// samlorwsfedexternaldomainfederation-get and
// identityproviderbase-availableprovidertypes operation documentation and
// the samlOrWsFedExternalDomainFederation, samlOrWsFedProvider and
// identityProviderBase resource references on 2026-10-05. List, get and
// count take D/A Domain.Read.All least-privileged
// (Domain.ReadWrite.All is the documented higher-privileged alternative);
// available provider types take D/A IdentityProvider.Read.All
// least-privileged (IdentityProvider.ReadWrite.All higher). Delegated
// callers additionally need External Identity Provider Administrator, the
// least-privileged supported role documented for these reads. Personal
// Microsoft accounts are not supported. No per-operation licence
// prerequisite is stated for these reads. Workforce context only: domain
// federationConfiguration sub-reads (/domains/{id}/federationConfiguration),
// the domains navigation property ($expand stays out), beta routes and every
// mutation stay out. No mutation lives here.
//
// Secrecy by construction: federation configurations carry the public
// token-signing certificate only, but the base64 `signingCertificate` blob
// is omitted from every default select and returned only on an explicit
// --select naming it, still under the truncation marker with a --full
// escape hatch. No private key material exists on these resources.

// Every federation-configuration property this slice may request or
// display, matching the reviewed resource. The domains navigation property
// needs its own read ($expand is not reviewed here). Anything else fails
// before credentials.
export const KNOWN_FEDERATION_FIELDS: readonly string[] = [
  "id",
  "displayName",
  "issuerUri",
  "metadataExchangeUri",
  "passiveSignInUri",
  "preferredAuthenticationProtocol",
  "signingCertificate",
];
const KNOWN_FEDERATIONS = new Set(KNOWN_FEDERATION_FIELDS);
// @odata.type is preserved on federation rows without being selectable: it
// names the derived configuration kind.
const FEDERATION_TYPE_PROPERTY = "@odata.type";

// Compact rows: identifier and name; the kind rides as @odata.type.
const DEFAULT_LIST_SELECT = ["id", "displayName"];
// Show rows: the full reviewed set except the public-certificate blob,
// which needs an explicit --select naming it.
const DEFAULT_SHOW_SELECT = KNOWN_FEDERATION_FIELDS.filter(field => field !== "signingCertificate");
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_FEDERATION_SCOPES = ["https://graph.microsoft.com/Domain.Read.All"];
export const DEFAULT_FEDERATION_TYPES_SCOPES = ["https://graph.microsoft.com/IdentityProvider.Read.All"];
const TRUNCATE_AT = 500;

export type FederationConfigurationFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!KNOWN_FEDERATIONS.has(part)) {
      throw new AxiError(`Unknown federation-configuration property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known federation-configuration properties: ${KNOWN_FEDERATION_FIELDS.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: FederationConfigurationFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
// @odata.type survives projection so the configuration kind is never lost.
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
  if (typeof source[FEDERATION_TYPE_PROPERTY] === "string") projected[FEDERATION_TYPE_PROPERTY] = source[FEDERATION_TYPE_PROPERTY];
  return { row: projected, truncated };
}

function selectedFields(
  flags: FederationConfigurationFlags,
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

function fullHint(command: string, flags: FederationConfigurationFlags, profileName: string): string {
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

const FEDERATION_DENIAL_HINTS = [
  "Federation-configuration reads need Domain.Read.All least-privileged (Domain.ReadWrite.All is the documented higher-privileged alternative): delegated callers pass one as --scopes, application access needs admin-consented Domain.Read.All on the configured .default audience",
  "Delegated callers additionally need External Identity Provider Administrator, the least-privileged supported Entra role for these reads",
  "Personal Microsoft accounts are not supported for federation-configuration reads",
  "No per-operation licence prerequisite is stated for these reads; never diagnose licence solely from HTTP 403",
];

const TYPES_DENIAL_HINTS = [
  "Available provider types need IdentityProvider.Read.All least-privileged (IdentityProvider.ReadWrite.All is the documented higher-privileged alternative): delegated callers pass one as --scopes, application access needs admin-consented IdentityProvider.Read.All on the configured .default audience",
  "Delegated callers additionally need External Identity Provider Administrator, the least-privileged supported Entra role for this read",
  "Personal Microsoft accounts are not supported for available-types reads",
  "No per-operation licence prerequisite is stated for this read; never diagnose licence solely from HTTP 403",
];

const WORKFORCE_NOTE = "Workforce tenant context only; domain federationConfiguration sub-reads and the domains navigation property belong to separate reviews and are never sent here";
const CERT_BLOB_NOTE = "Public signing-certificate blobs are omitted by default; request signingCertificate explicitly with --select signingCertificate";

export async function listFederationConfigurations(
  session: GraphSession,
  flags: FederationConfigurationFlags,
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
  const scopes = scopesFor(flags, DEFAULT_FEDERATION_SCOPES, profile, help);
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
  const result = await withGuidance(FEDERATION_DENIAL_HINTS, () => session.collect(args));
  const effectiveFlags: FederationConfigurationFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const configurations: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, full);
    configurations.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra federation-configuration show --id <configuration-id> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("entra federation-configuration list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      federationConfigurations: configurations,
      ...listTotals(configurations.length, result.total, "federation configurations", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor <cursor-from-output> ${profileHint(profileName)}`, showHint, CERT_BLOB_NOTE, WORKFORCE_NOTE],
    };
  }
  if (!configurations.length) {
    return {
      federationConfigurations: configurations,
      ...listTotals(configurations.length, result.total, "federation configurations", true),
      complete: true,
      help: ["0 federation configurations matched; no SAML or WS-Fed federation is configured for the tenant, or no row passed --filter", CERT_BLOB_NOTE, WORKFORCE_NOTE],
    };
  }
  return { federationConfigurations: configurations, ...listTotals(configurations.length, result.total, "federation configurations", true), complete: true, help: [...truncationHints, showHint, CERT_BLOB_NOTE, WORKFORCE_NOTE] };
}

export async function showFederationConfiguration(
  session: GraphSession,
  flags: FederationConfigurationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, DEFAULT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_FEDERATION_SCOPES, profile, help);
  const full = flags.full === true;
  const id = String(flags.id);
  if (!id.trim()) throw new AxiError("--id needs the federation-configuration object ID", "VALIDATION_ERROR", [help]);
  const raw = await withGuidance(FEDERATION_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "identityProviderBase-id": id },
    query: { $select: select.join(",") },
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed federation-configuration body", "GRAPH_ERROR", [
      "Single-configuration reads carry one federation configuration object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  const helpHints: string[] = [...(truncated ? [fullHint("entra federation-configuration show", flags, profileName)] : []), CERT_BLOB_NOTE, WORKFORCE_NOTE];
  return { federationConfiguration: row, help: helpHints };
}

// The $count route returns a text/plain integer scalar rather than a JSON
// collection, so the leaf reads it through session.execute with scalar mode
// and accepts only a non-negative integer. --filter narrows the count
// server-side; there is no --select/--limit/--cursor contract on the count.
export async function countFederationConfigurations(
  session: GraphSession,
  flags: FederationConfigurationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  void profileName;
  const scopes = scopesFor(flags, DEFAULT_FEDERATION_SCOPES, profile, help);
  const query: Record<string, string> = {};
  if (flags.filter !== undefined) query.$filter = String(flags.filter);
  const raw = await withGuidance(FEDERATION_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: {},
    query,
    scopes,
    scalar: true,
  }));
  if (typeof raw !== "number" || !Number.isInteger(raw) || raw < 0) {
    throw new AxiError("Graph returned a malformed federation-configuration count body", "GRAPH_ERROR", [
      "Federation-configuration counts carry one non-negative integer scalar; treat anything else as unknown, not empty",
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

export async function availableFederationProviderTypes(
  session: GraphSession,
  flags: FederationConfigurationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
): Promise<Record<string, unknown>> {
  const scopes = scopesFor(flags, DEFAULT_FEDERATION_TYPES_SCOPES, profile, help);
  const raw = await withGuidance(TYPES_DENIAL_HINTS, () => session.execute({ profile, operation, scopes }));
  const value = raw !== null && typeof raw === "object" && !Array.isArray(raw)
    ? (raw as Record<string, unknown>).value : undefined;
  if (!Array.isArray(value) || value.some(type => typeof type !== "string")) {
    throw new AxiError("Graph returned a malformed available federation provider types body", "GRAPH_ERROR", [
      "Available-provider-type reads carry a value array of strings; treat anything else as unknown, not empty",
    ]);
  }
  return {
    availableProviderTypes: value,
    count: { returned: value.length, complete: true },
    help: ["Available provider types depend on tenant configuration and licensing; availability does not mean a federation configuration exists", WORKFORCE_NOTE],
  };
}
