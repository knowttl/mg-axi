import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// READ-01: the Entra user read mapping behind `mg-axi entra user list/show`.
// Operation construction stays beside its command; the shared session owns
// URLs, credentials, paging, retries and error translation, and the SDK owns
// TOON rendering. This module only maps flags to session calls and projects
// rows for compact output.
//
// Reviewed against the v1.0 user-list and user-get operation documentation on
// 2026-10-04. List and get return the same default property set without
// $select; accountEnabled is never default and needs User.EnableDisableAccount.All
// with User.Read.All. Delegated list reads User.ReadBasic.All for basics and
// User.Read.All for richer data; application reads User.Read.All.

// Every property this slice may request or display. Anything else fails
// before credentials so typos never become misleading server queries.
export const KNOWN_USER_FIELDS: readonly string[] = [
  "businessPhones",
  "displayName",
  "givenName",
  "id",
  "jobTitle",
  "mail",
  "mobilePhone",
  "officeLocation",
  "preferredLanguage",
  "surname",
  "userPrincipalName",
  "accountEnabled",
  "department",
  "userType",
];
const KNOWN = new Set(KNOWN_USER_FIELDS);

// Compact list rows: identifier, title, address and one status signal, all
// readable with basic consent.
const DEFAULT_LIST_SELECT = ["id", "displayName", "userPrincipalName", "mail"];
// Show rows: the full default server set, a richer view of the same user.
const DEFAULT_SHOW_SELECT = [
  "businessPhones",
  "displayName",
  "givenName",
  "id",
  "jobTitle",
  "mail",
  "mobilePhone",
  "officeLocation",
  "preferredLanguage",
  "surname",
  "userPrincipalName",
];
// One SOC-capable default covers every READ-01 property; least-privilege
// basic deployments pass --scopes explicitly instead.
export const DEFAULT_DELEGATED_SCOPES = ["https://graph.microsoft.com/User.Read.All"];
const TRUNCATE_AT = 500;

export type UserFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, flag: string, help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!KNOWN.has(part)) {
      throw new AxiError(`Unknown user property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known user properties: ${KNOWN_USER_FIELDS.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: UserFlags, profile: AnyProfile, help: string): string[] | undefined {
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
  if (flags.scopes === undefined) return [...DEFAULT_DELEGATED_SCOPES];
  const scopes = String(flags.scopes)
    .split(",")
    .map(scope => scope.trim())
    .filter(scope => scope.length > 0);
  if (
    !scopes.length ||
    scopes.some(scope => !/^https:\/\/graph\.microsoft\.com\/[A-Za-z][A-Za-z.-]+$/.test(scope) || scope.endsWith("/.default"))
  ) {
    throw new AxiError("Explicit delegated Graph scopes are required", "VALIDATION_ERROR", [
      help,
      "mg-axi login --profile <name> --scopes <comma-separated-Graph-scopes>",
    ]);
  }
  return [...new Set(scopes)].sort();
}

function truncateValue(value: unknown, full: boolean): { value: unknown; truncated: boolean } {
  if (typeof value === "string") {
    if (full || value.length <= TRUNCATE_AT) return { value, truncated: false };
    return { value: `${value.slice(0, TRUNCATE_AT)}... (truncated, ${value.length} chars total)`, truncated: true };
  }
  if (Array.isArray(value)) {
    let truncated = false;
    const projected = value.map(entry => {
      const result = truncateValue(entry, full);
      truncated = truncated || result.truncated;
      return result.value;
    });
    return { value: projected, truncated };
  }
  if (value !== null && typeof value === "object") {
    let truncated = false;
    const projected: Record<string, unknown> = {};
    for (const [key, entry] of Object.entries(value as Record<string, unknown>)) {
      const result = truncateValue(entry, full);
      truncated = truncated || result.truncated;
      projected[key] = result.value;
    }
    return { value: projected, truncated };
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

function selectedFields(flags: UserFlags, defaults: string[], help: string): { select: string[]; fields: string[] } {
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

function fullHint(action: string, flags: UserFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi entra user ${action} ${args.join(" ")}`;
}

export async function listUsers(
  session: GraphSession,
  flags: UserFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const selection = flags.cursor === undefined || flags.select !== undefined
    ? selectedFields(flags, DEFAULT_LIST_SELECT, help)
    : undefined;
  if (!selection && flags.fields !== undefined) fieldList(flags.fields, "fields", help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const query: Record<string, string> = {};
  if (selection) query.$select = selection.select.join(",");
  if (flags.filter !== undefined) query.$filter = String(flags.filter);
  const args: CollectArgs = { profile, operation, query, scopes };
  if (flags.cursor !== undefined) {
    if (!String(flags.cursor).trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
    args.cursor = String(flags.cursor);
  }
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await session.collect(args);
  const effectiveFlags: UserFlags = { ...flags, select: result.query.$select ?? DEFAULT_LIST_SELECT.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const { fields } = selectedFields(effectiveFlags, DEFAULT_LIST_SELECT, help);
  const users: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, full);
    users.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const showHint = `mg-axi entra user show --id <user-id-or-upn> ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint("list", effectiveFlags, profileName)] : [];
  if (!result.complete) {
    return {
      users,
      count: { returned: users.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor <cursor-from-output> ${profileHint(profileName)}`, showHint],
    };
  }
  const count = { returned: users.length, complete: true };
  if (!users.length) {
    return {
      users,
      count,
      help: [
        `mg-axi entra user list --filter <odata-filter> ${profileHint(profileName)}`,
        "0 users matched; the absence of results is the answer, not an error",
      ],
    };
  }
  const helpHints = [...truncationHints, showHint];
  return { users, count, help: helpHints };
}

export async function showUser(
  session: GraphSession,
  flags: UserFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  const { select, fields } = selectedFields(flags, DEFAULT_SHOW_SELECT, help);
  const scopes = scopesFor(flags, profile, help);
  const full = flags.full === true;
  const raw = await session.execute({
    profile,
    operation,
    params: { "user-id": String(flags.id) },
    query: { $select: select.join(",") },
    scopes,
  });
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new AxiError("Graph returned a malformed user body", "GRAPH_ERROR", [
      "Single-user reads carry one user object; treat anything else as unknown, not empty",
    ]);
  }
  const { row, truncated } = project(raw, fields, full);
  if (truncated) return { user: row, help: [fullHint("show", flags, profileName)] };
  return { user: row };
}
