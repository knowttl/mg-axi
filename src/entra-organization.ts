import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

// EXT-01 organization subfamily: the read mapping behind
// `mg-axi entra organization list/show`, the default-branding read behind
// `mg-axi entra organization branding show`, and the branding-localization
// reads behind `mg-axi entra organization branding-localization
// list/show`. Operation construction stays beside its command; the shared
// session owns URLs, credentials, paging, retries and error translation,
// and the SDK owns TOON rendering. This module only maps flags to session
// calls and projects rows for compact output.
//
// Reviewed against the v1.0 organization-list, organization-get,
// organizationalbranding-get, organizationalbranding-list-localizations and
// organizationalbrandinglocalization-get operation documentation on
// 2026-10-04. Organization list/get take D User.Read for restricted basics
// (id, displayName and verifiedDomains only; every other property returns
// null) or Organization.Read.All for full metadata, and A
// Organization.Read.All; delegated callers additionally need a supported
// Entra role (Directory Readers and Global Reader are among the supported
// least-privilege roles). Branding and localization reads take D User.Read
// least-privileged (OrganizationalBranding.Read.All is the purpose-built
// delegated alternative) and A OrganizationalBranding.Read.All; delegated
// callers additionally need Global Reader or Organizational Branding
// Administrator least-privileged. Personal Microsoft accounts are not
// supported on any of these reads. No P1/P2 prerequisite is stated for the
// reads, but configuring custom branding itself needs P1/P2. Contact
// fields on the organization (businessPhones, notification mails) are
// personal data. Every reviewed read documents $select only, so none of
// these leaves offers --filter. The branding GET documents Accept-Language
// as required; the session sends `Accept-Language: 0` on that route to read
// the default branding, and locale variants come from the localizations
// collection. Stream image properties (bannerLogo, backgroundImage and
// friends) need their own binary-output contract and stay out of this
// slice, as do certificateBasedAuthConfiguration, extensions, beta-only
// settings/partnerInformation/themes and the POST lookup actions. No
// organization mutation exists in this slice.

// Every organization property this slice may request or display, matching
// the reviewed raw surface. Anything else fails before credentials.
export const KNOWN_ORGANIZATION_FIELDS: readonly string[] = [
  "id",
  "deletedDateTime",
  "businessPhones",
  "city",
  "country",
  "countryLetterCode",
  "createdDateTime",
  "defaultUsageLocation",
  "displayName",
  "isMultipleDataLocationsForServicesEnabled",
  "marketingNotificationEmails",
  "onPremisesLastSyncDateTime",
  "onPremisesSyncEnabled",
  "partnerTenantType",
  "postalCode",
  "preferredLanguage",
  "privacyProfile",
  "provisionedPlans",
  "securityComplianceNotificationMails",
  "securityComplianceNotificationPhones",
  "state",
  "street",
  "technicalNotificationMails",
  "tenantType",
  "assignedPlans",
  "verifiedDomains",
];
const KNOWN_ORGANIZATIONS = new Set(KNOWN_ORGANIZATION_FIELDS);

// Every non-Stream branding property this slice may request or display.
// Stream image properties (backgroundImage, bannerLogo, customCSS, favicon,
// headerLogo, squareLogo, squareLogoDark) are real server properties but
// need a binary-output contract this text CLI does not have; requesting one
// fails with a targeted hint instead of an unknown-property error.
export const KNOWN_BRANDING_FIELDS: readonly string[] = [
  "id",
  "backgroundColor",
  "backgroundImageRelativeUrl",
  "bannerLogoRelativeUrl",
  "cdnList",
  "signInPageText",
  "squareLogoRelativeUrl",
  "usernameHintText",
  "customAccountResetCredentialsUrl",
  "customCannotAccessYourAccountText",
  "customCannotAccessYourAccountUrl",
  "customCSSRelativeUrl",
  "customForgotMyPasswordText",
  "customPrivacyAndCookiesText",
  "customPrivacyAndCookiesUrl",
  "customResetItNowText",
  "customTermsOfUseText",
  "customTermsOfUseUrl",
  "faviconRelativeUrl",
  "headerBackgroundColor",
  "headerLogoRelativeUrl",
  "squareLogoDarkRelativeUrl",
];
const KNOWN_BRANDING = new Set(KNOWN_BRANDING_FIELDS);
const STREAM_BRANDING_FIELDS = new Set([
  "backgroundImage",
  "bannerLogo",
  "customCSS",
  "favicon",
  "headerLogo",
  "squareLogo",
  "squareLogoDark",
]);

// Compact organization rows: the tenant identity, display name, tenant kind
// and the verified domains that name the tenant on the network.
const DEFAULT_ORGANIZATION_LIST_SELECT = ["id", "displayName", "tenantType", "verifiedDomains"];
// Show rows: the full reviewed organization set, including the technical
// notification mails and privacy profile.
const DEFAULT_ORGANIZATION_SHOW_SELECT = [...KNOWN_ORGANIZATION_FIELDS];
// Branding show rows: the full reviewed non-Stream set.
const DEFAULT_BRANDING_SHOW_SELECT = [...KNOWN_BRANDING_FIELDS];
// Compact localization rows: the locale id and the sign-in text a user
// actually sees.
const DEFAULT_LOCALIZATION_LIST_SELECT = ["id", "signInPageText", "usernameHintText", "backgroundColor"];
// Localization show rows: the full reviewed non-Stream set.
const DEFAULT_LOCALIZATION_SHOW_SELECT = [...KNOWN_BRANDING_FIELDS];
// Delegated defaults are operation-specific; application profiles use their
// configured .default audience and reject --scopes.
export const DEFAULT_ORGANIZATION_SCOPES = ["https://graph.microsoft.com/Organization.Read.All"];
export const DEFAULT_BRANDING_SCOPES = ["https://graph.microsoft.com/User.Read"];
const TRUNCATE_AT = 500;

export type OrganizationFlags = Record<string, string | boolean>;

function fieldList(
  raw: unknown,
  known: Set<string>,
  knownList: readonly string[],
  flag: string,
  help: string,
  streamHint?: string,
): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (streamHint !== undefined && STREAM_BRANDING_FIELDS.has(part)) {
      throw new AxiError(`${part} is a Stream image route served by a later piece; this read returns non-Stream branding metadata only`, "VALIDATION_ERROR", [help, streamHint]);
    }
    if (!known.has(part)) {
      throw new AxiError(`Unknown organization property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: OrganizationFlags, defaults: string[], profile: AnyProfile, help: string): string[] | undefined {
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
// This matters for User.Read restricted basics, where every property outside
// id, displayName and verifiedDomains returns null.
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
  flags: OrganizationFlags,
  known: Set<string>,
  knownList: readonly string[],
  defaults: string[],
  help: string,
  streamHint?: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, known, knownList, "select", help, streamHint);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, known, knownList, "fields", help, streamHint);
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

function fullHint(command: string, flags: OrganizationFlags, profileName: string): string {
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

const ORGANIZATION_DENIAL_HINTS = [
  "Organization reads need Organization.Read.All for full metadata plus a supported Entra role for delegated access (Directory Readers and Global Reader are among the supported least-privilege roles), or admin-consented Organization.Read.All for application access; delegated User.Read returns only id, displayName and verifiedDomains with every other property null",
  "Personal Microsoft accounts are not supported for organization reads",
  "No P1/P2 prerequisite is stated for organization reads; never diagnose licence solely from HTTP 403",
];

const BRANDING_DENIAL_HINTS = [
  "Branding reads need User.Read least-privileged or the purpose-built OrganizationalBranding.Read.All (Organization.Read.All also works) plus Global Reader or Organizational Branding Administrator for delegated access, or admin-consented OrganizationalBranding.Read.All for application access",
  "Personal Microsoft accounts are not supported for branding reads",
  "Configuring custom branding needs P1/P2; a branding 404 may indicate unconfigured branding or a missing or inaccessible organization; never diagnose licence solely from HTTP 403",
];

const STREAM_HINT = "Stream image properties (backgroundImage, bannerLogo, customCSS, favicon, headerLogo, squareLogo, squareLogoDark) are served by a later piece with its own binary-output contract";

interface CollectionCommon {
  cursor: string | undefined;
  select: string[];
  fields: string[];
  scopes: string[] | undefined;
  full: boolean;
}

// Restoring the saved select keeps cursor resumes lossless when --select is
// omitted. Organization and localization lists document $select only, so
// neither leaf offers --filter: strict input validation refuses the flag
// before credentials instead of forwarding a misleading request.
function collectionCommon(
  session: GraphSession,
  flags: OrganizationFlags,
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
  return { cursor, select, fields, scopes: scopesFor(flags, defaultScopes, profile, help), full: flags.full === true };
}

function collectArgs(
  profile: AnyProfile,
  operation: SessionOperation,
  common: CollectionCommon,
  flags: OrganizationFlags,
  help: string,
): CollectArgs {
  const query: Record<string, string> = { $select: common.select.join(",") };
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

function organizationId(flags: OrganizationFlags, help: string, flag: string): string {
  const id = String(flags[flag]);
  if (!id.trim()) throw new AxiError(`--${flag} needs the tenant organization UUID`, "VALIDATION_ERROR", [help]);
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

export async function listOrganizations(
  session: GraphSession,
  flags: OrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_ORGANIZATIONS, KNOWN_ORGANIZATION_FIELDS,
    DEFAULT_ORGANIZATION_LIST_SELECT, DEFAULT_ORGANIZATION_SCOPES, operation, help, profile);
  const result = await withGuidance(ORGANIZATION_DENIAL_HINTS, () => session.collect(collectArgs(profile, operation, common, flags, help)));
  const effectiveFlags: OrganizationFlags = { ...flags, select: result.query.$select ?? DEFAULT_ORGANIZATION_LIST_SELECT.join(",") };
  const organizations: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    organizations.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra organization show --id <organization-id> ${profileHint(profileName)}`;
  const filterNote = "Organization lists offer no --filter: Graph documents $select only for /organization";
  const truncationHints = truncated ? [fullHint("entra organization list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      organizations,
      ...listTotals(organizations.length, result.total, "organizations", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, filterNote],
    };
  }
  if (!organizations.length) {
    return {
      organizations,
      ...listTotals(organizations.length, result.total, "organizations", true),
      complete: true,
      help: [filterNote, "0 organizations matched; the absence of results is the answer, not an error"],
    };
  }
  return { organizations, ...listTotals(organizations.length, result.total, "organizations", true), complete: true, help: [...truncationHints, showHint, filterNote] };
}

export async function showOrganization(
  session: GraphSession,
  flags: OrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_ORGANIZATIONS, KNOWN_ORGANIZATION_FIELDS, DEFAULT_ORGANIZATION_SHOW_SELECT, help);
  const scopes = scopesFor(flags, DEFAULT_ORGANIZATION_SCOPES, profile, help);
  const full = flags.full === true;
  const raw = await withGuidance(ORGANIZATION_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "organization-id": organizationId(flags, help, "id") },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "organization");
  const helpHints: string[] = truncated ? [fullHint("entra organization show", flags, profileName)] : [];
  if (helpHints.length) return { organization: row, help: helpHints };
  return { organization: row };
}

export async function showBranding(
  session: GraphSession,
  flags: OrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_BRANDING, KNOWN_BRANDING_FIELDS, DEFAULT_BRANDING_SHOW_SELECT, help, STREAM_HINT);
  const scopes = scopesFor(flags, DEFAULT_BRANDING_SCOPES, profile, help);
  const full = flags.full === true;
  const organization = organizationId(flags, help, "organization");
  const raw = await withGuidance(BRANDING_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "organization-id": organization },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "branding");
  const helpHints: string[] = [
    ...(truncated ? [fullHint("entra organization branding show", flags, profileName)] : []),
    `mg-axi entra organization branding-localization list --organization ${shellValue(organization)} ${profileHint(profileName)}`,
  ];
  return { branding: row, help: helpHints };
}

export async function listBrandingLocalizations(
  session: GraphSession,
  flags: OrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const common = collectionCommon(session, flags, KNOWN_BRANDING, KNOWN_BRANDING_FIELDS,
    DEFAULT_LOCALIZATION_LIST_SELECT, DEFAULT_BRANDING_SCOPES, operation, help, profile);
  const args = collectArgs(profile, operation, common, flags, help);
  const organization = organizationId(flags, help, "organization");
  args.params = { "organization-id": organization };
  const result = await withGuidance(BRANDING_DENIAL_HINTS, () => session.collect(args));
  const effectiveFlags: OrganizationFlags = { ...flags, select: result.query.$select ?? DEFAULT_LOCALIZATION_LIST_SELECT.join(",") };
  const localizations: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, common.fields, common.full);
    localizations.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra organization branding-localization show --organization ${shellValue(organization)} --id <locale-id> ${profileHint(profileName)}`;
  const filterNote = "Localization lists offer no --filter: Graph documents $select only for branding localizations";
  const truncationHints = truncated ? [fullHint("entra organization branding-localization list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      brandingLocalizations: localizations,
      ...listTotals(localizations.length, result.total, "branding localizations", false),
      complete: false,
      reason: result.reason,
      cursor: result.cursor,
      help: [...truncationHints, resumeHint(profileName), showHint, filterNote],
    };
  }
  if (!localizations.length) {
    return {
      brandingLocalizations: localizations,
      ...listTotals(localizations.length, result.total, "branding localizations", true),
      complete: true,
      help: [filterNote, "0 branding localizations matched; the absence of results is the answer, not an error"],
    };
  }
  return { brandingLocalizations: localizations, ...listTotals(localizations.length, result.total, "branding localizations", true), complete: true, help: [...truncationHints, showHint, filterNote] };
}

export async function showBrandingLocalization(
  session: GraphSession,
  flags: OrganizationFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, KNOWN_BRANDING, KNOWN_BRANDING_FIELDS, DEFAULT_LOCALIZATION_SHOW_SELECT, help, STREAM_HINT);
  const scopes = scopesFor(flags, DEFAULT_BRANDING_SCOPES, profile, help);
  const full = flags.full === true;
  const organization = organizationId(flags, help, "organization");
  const locale = String(flags.id);
  if (!locale.trim()) throw new AxiError("--id needs the branding locale id, for example fr-FR", "VALIDATION_ERROR", [help]);
  const raw = await withGuidance(BRANDING_DENIAL_HINTS, () => session.execute({
    profile,
    operation,
    params: { "organization-id": organization, "organizationalBrandingLocalization-id": locale },
    query: { $select: select.join(",") },
    scopes,
  }));
  const { row, truncated } = singleResult(raw, fields, full, "branding localization");
  if (truncated) return { brandingLocalization: row, help: [fullHint("entra organization branding-localization show", flags, profileName)] };
  return { brandingLocalization: row };
}
