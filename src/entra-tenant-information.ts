import { AxiError } from "axi-sdk-js";
import { resolveSessionOperation, type GraphSession, type SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// EXT-04e tenant-lookup subfamily: the read mapping behind
// `mg-axi entra tenant-information show --domain-name <domain>` and
// `mg-axi entra tenant-information show --tenant-id <tenant-guid>`.
// Operation construction stays beside its command; the shared session owns
// URLs, credentials, retries and error translation, and the SDK owns TOON
// rendering. This module only maps flags to session calls and projects rows
// for compact output.
//
// Reviewed against the v1.0 findTenantInformationByDomainName,
// findTenantInformationByTenantId and tenantInformation resource
// documentation on 2026-10-05. Every read takes D/A
// CrossTenantInformation.ReadBasic.All as least privilege; delegated
// callers pass it as --scopes while application callers need it
// admin-consented on the configured .default audience. The operation pages
// state no Entra role prerequisite and personal Microsoft accounts are not
// supported. No P1/P2 prerequisite is stated for these reads. The pages
// document no query parameters, so the leaf offers no --select/--fields/--
// filter/--limit/--cursor and always returns the full reviewed
// tenantInformation set. The addresses are function segments, so the key
// reaches the session only as an explicit CLI flag and the allowlisted
// session binding validates, OData-quotes and encodes it; any other value
// fails before credentials. Beta stays out. No mutation lives here.

// Every tenant-information property this slice may display, matching the
// reviewed tenantInformation resource. Anything else fails before
// credentials.
export const KNOWN_TENANT_INFORMATION_FIELDS: readonly string[] = [
  "defaultDomainName",
  "displayName",
  "federationBrandName",
  "tenantId",
];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_TENANT_INFORMATION_SCOPES = ["https://graph.microsoft.com/CrossTenantInformation.ReadBasic.All"];
const TRUNCATE_AT = 500;

export type TenantInformationFlags = Record<string, string | boolean>;

function scopesFor(flags: TenantInformationFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
  fields: readonly string[],
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

function profileHint(profileName: string): string {
  return `--profile ${shellValue(profileName)}`;
}

function shellValue(value: string): string {
  return /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
}

function fullHint(command: string, flags: TenantInformationFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the scope and prerequisites that actually unlock it,
// because a 403 alone never says which prerequisite is missing.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const TENANT_INFORMATION_DENIAL_HINTS = [
  "Tenant-information lookups need CrossTenantInformation.ReadBasic.All, passed as --scopes for delegated access or admin-consented for application access",
  "No Entra role is required for tenant-information lookups; a 403 names a missing grant, not a missing role",
  "Personal Microsoft accounts are not supported for tenant-information lookups",
  "No P1/P2 prerequisite is stated for tenant-information lookups; never diagnose licence solely from HTTP 403",
];

// Exactly one lookup key selects the function route; the value itself is
// validated by the allowlisted session binding before credentials. The
// domain operation arrives catalogue-resolved; the tenant-ID route resolves
// from the same inventory authority beside its command.
function lookupTarget(flags: TenantInformationFlags, domainOperation: SessionOperation, help: string): { operation: SessionOperation; params: Record<string, string> } {
  const domain = flags["domain-name"];
  const tenantId = flags["tenant-id"];
  if (domain !== undefined && tenantId !== undefined) {
    throw new AxiError("--domain-name and --tenant-id cannot be combined", "VALIDATION_ERROR", [help]);
  }
  if (domain !== undefined) {
    if (!String(domain).trim()) throw new AxiError("--domain-name needs the tenant domain name", "VALIDATION_ERROR", [help]);
    return { operation: domainOperation, params: { domainName: String(domain) } };
  }
  if (tenantId !== undefined) {
    if (!String(tenantId).trim()) throw new AxiError("--tenant-id needs the tenant GUID", "VALIDATION_ERROR", [help]);
    return {
      operation: resolveSessionOperation("v1.0", "GET", "/tenantRelationships/findTenantInformationByTenantId(tenantId='{tenantId}')"),
      params: { tenantId: String(tenantId) },
    };
  }
  throw new AxiError("Tenant-information show needs exactly one of --domain-name or --tenant-id", "VALIDATION_ERROR", [help]);
}

export async function showTenantInformation(
  session: GraphSession,
  flags: TenantInformationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const target = lookupTarget(flags, operation, help);
  const scopes = scopesFor(flags, DEFAULT_TENANT_INFORMATION_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(TENANT_INFORMATION_DENIAL_HINTS, () => session.execute({
    profile,
    operation: target.operation,
    params: target.params,
    scopes,
  }));
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed tenant-information body", "GRAPH_ERROR", [
      "Tenant-information lookups carry one tenantInformation object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, KNOWN_TENANT_INFORMATION_FIELDS, full);
  if (truncated) return { tenantInformation: row, help: [fullHint("entra tenant-information show", flags, profileName)] };
  return { tenantInformation: row };
}
