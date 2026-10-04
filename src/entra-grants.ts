import { AxiError } from "axi-sdk-js";
import type { CollectArgs, GraphSession, SessionOperation } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

// READ-08: the Entra consent-grant read mapping behind
// `mg-axi entra service-principal oauth2-grant list` and
// `mg-axi entra service-principal app-role-assignment list` for a named
// service principal (the client). Operation construction stays beside its
// command; the shared session owns URLs, credentials, paging, retries and
// error translation, and the SDK owns TOON rendering. This module only maps
// flags to session calls and projects rows for compact output.
//
// Reviewed against the v1.0 serviceprincipal-list-oauth2permissiongrants,
// serviceprincipal-list-approleassignments, oauth2permissiongrant-resource
// and approleassignment-resource operation documentation on 2026-10-04.
// Delegated-grant reads take D/A Directory.Read.All; app-role reads take
// D/A Application.Read.All. Delegated callers additionally need a supported
// administrator role per operation. Reads never request a write-consent
// scope such as DelegatedPermissionGrant.ReadWrite.All,
// Application.ReadWrite.All or Directory.ReadWrite.All. Neither collection
// documents an advanced-query contract, so --filter passes through as plain
// $filter with no $count or ConsistencyLevel attached.
//
// The two views cover actual granted consent only: oauth2PermissionGrants
// are the delegated scopes granted to the client (consentType AllPrincipals
// for every user, Principal for one user named by principalId); an
// appRoleAssignment whose principal is the named client is an app-only
// grant of the resource API's app role. The application's requested
// permissions (requiredResourceAccess) are declared on the application
// object and are never shown here. Grant creation, revocation and consent
// belong to later write slices, never to these reads. No P1/P2 prerequisite
// is stated for these consent reads.

// Delegated-grant rows carry the full reviewed oAuth2PermissionGrant set:
// clientId is the authorizing client's object id (not appId).
export const KNOWN_GRANT_FIELDS: readonly string[] = [
  "id",
  "clientId",
  "consentType",
  "principalId",
  "resourceId",
  "scope",
];
// App-role rows carry the reviewed appRoleAssignment set minus
// deletedDateTime: listed assignments are live records, and the deleted
// timestamp is always null on them.
export const KNOWN_APP_ROLE_FIELDS: readonly string[] = [
  "id",
  "appRoleId",
  "createdDateTime",
  "principalDisplayName",
  "principalId",
  "principalType",
  "resourceDisplayName",
  "resourceId",
];
const GRANT_KNOWN = new Set(KNOWN_GRANT_FIELDS);
const APP_ROLE_KNOWN = new Set(KNOWN_APP_ROLE_FIELDS);

// Compact grant rows: the grant id, who it covers, which API it targets and
// the granted scopes. clientId is constant for one client's list, while
// principalId names the covered user on Principal grants and stays an
// explicit null on AllPrincipals grants.
const DEFAULT_GRANT_SELECT = ["id", "consentType", "principalId", "resourceId", "scope"];
// Compact app-role rows: the assignment id, the granted role and the
// resource API. principalId is constant for one client's list.
const DEFAULT_APP_ROLE_SELECT = ["id", "appRoleId", "resourceDisplayName", "resourceId"];
export const DEFAULT_GRANT_SCOPES = ["https://graph.microsoft.com/Directory.Read.All"];
export const DEFAULT_APP_ROLE_SCOPES = ["https://graph.microsoft.com/Application.Read.All"];
const TRUNCATE_AT = 500;

// The only operations this slice ever binds. Anything else - grant creation,
// revocation, consent, tenant-wide/user/group grant enumerations - is refused
// before credentials.
const READ_OPERATIONS: Readonly<Record<string, string>> = {
  "GET:/servicePrincipals/{servicePrincipal-id}/oauth2PermissionGrants": "oauth2-grant",
  "GET:/servicePrincipals/{servicePrincipal-id}/appRoleAssignments": "app-role-assignment",
};

function checkReadOperation(operation: SessionOperation, help: string): void {
  // Version stays the session's decision: beta reads remain preview-gated
  // there, exactly like every other named slice. This gate only ensures a
  // fabricated grant-mutating or unrelated route can never reach
  // credentials.
  const route = `${operation.method}:${operation.path}`;
  if (operation.method !== "GET" || !Object.hasOwn(READ_OPERATIONS, route)) {
    throw new AxiError(`Refused non-read route ${operation.id}: READ-08 serves only the two catalogued service-principal grant GETs`, "VALIDATION_ERROR", [
      help,
      "Grant creation, revocation and consent are never constructed here",
    ]);
  }
}

export type GrantFlags = Record<string, string | boolean>;

function fieldList(raw: unknown, flag: string, known: Set<string>, knownList: readonly string[], help: string): string[] {
  const parts = String(raw)
    .split(",")
    .map(part => part.trim())
    .filter(part => part.length > 0);
  if (!parts.length) throw new AxiError(`--${flag} needs at least one property`, "VALIDATION_ERROR", [help]);
  const fields: string[] = [];
  for (const part of parts) {
    if (!known.has(part)) {
      throw new AxiError(`Unknown property ${part} in --${flag}`, "VALIDATION_ERROR", [
        help,
        `Known properties: ${knownList.join(", ")}`,
      ]);
    }
    if (!fields.includes(part)) fields.push(part);
  }
  return fields;
}

function scopesFor(flags: GrantFlags, profile: AnyProfile, defaults: readonly string[], help: string): string[] | undefined {
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
  if (typeof value === "string" && !full && value.length > TRUNCATE_AT) {
    return { value: `${value.slice(0, TRUNCATE_AT)}... (truncated, ${value.length} chars total)`, truncated: true };
  }
  return { value, truncated: false };
}

// Local projection preserves Graph's null/missing distinction: an explicit
// null (for example principalId on an AllPrincipals grant) stays null, an
// absent property stays absent and is never synthesized.
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
  flags: GrantFlags,
  defaults: string[],
  known: Set<string>,
  knownList: readonly string[],
  help: string,
): { select: string[]; fields: string[] } {
  const select = flags.select === undefined ? [...defaults] : fieldList(flags.select, "select", known, knownList, help);
  const fields = flags.fields === undefined ? [...select] : fieldList(flags.fields, "fields", known, knownList, help);
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

function fullHint(command: string, flags: GrantFlags, profileName: string): string {
  const args = Object.entries({ ...flags, profile: profileName, full: true })
    .map(([name, value]) => value === true ? `--${name}` : `--${name} ${shellValue(String(value))}`);
  return `mg-axi ${command} ${args.join(" ")}`;
}

// Denials carry the session's generic grant/role/licence cause; each
// operation adds the roles and read scopes that actually unlock it, because
// a 403 alone never says which prerequisite is missing. Write-consent
// scopes are named only to warn against requesting them for reads.
function withGuidance<T>(hints: string[], run: () => Promise<T>): Promise<T> {
  return run().catch(error => {
    if (error instanceof AxiError && error.code === "GRAPH_ERROR" && /\(403\)/.test(error.message)) {
      throw new AxiError(error.message, "GRAPH_ERROR", [...hints, ...error.suggestions]);
    }
    throw error;
  });
}

const GRANT_DENIAL_HINTS = [
  "Delegated-grant reads need Directory.Read.All plus a supported directory role (for example Directory Readers, Global Reader or Application Administrator) for delegated access, or admin-consented Directory.Read.All for application access; reads never request DelegatedPermissionGrant.ReadWrite.All or Directory.ReadWrite.All",
];

const APP_ROLE_DENIAL_HINTS = [
  "App-role reads need Application.Read.All plus a supported directory role (for example Directory Readers, Application Administrator or Cloud Application Administrator) for delegated access, or admin-consented Application.Read.All for application access; reads never request Application.ReadWrite.All or Directory.ReadWrite.All",
];

interface CollectionShape {
  command: string;
  key: string;
  known: Set<string>;
  knownList: readonly string[];
  defaultSelect: string[];
  defaultScopes: readonly string[];
  denialHints: string[];
  scopeNote: string;
  emptyNote: string;
}

const OAUTH2_GRANT_LIST: CollectionShape = {
  command: "entra service-principal oauth2-grant list",
  key: "oauth2PermissionGrants",
  known: GRANT_KNOWN,
  knownList: KNOWN_GRANT_FIELDS,
  defaultSelect: DEFAULT_GRANT_SELECT,
  defaultScopes: DEFAULT_GRANT_SCOPES,
  denialHints: GRANT_DENIAL_HINTS,
  scopeNote: "These rows are granted delegated consent for the named client; the application's requested permissions (requiredResourceAccess) are declared on the application object and are not shown here",
  emptyNote: "0 delegated grants matched; the absence of results is the answer, not an error",
};

const APP_ROLE_LIST: CollectionShape = {
  command: "entra service-principal app-role-assignment list",
  key: "appRoleAssignments",
  known: APP_ROLE_KNOWN,
  knownList: KNOWN_APP_ROLE_FIELDS,
  defaultSelect: DEFAULT_APP_ROLE_SELECT,
  defaultScopes: DEFAULT_APP_ROLE_SCOPES,
  denialHints: APP_ROLE_DENIAL_HINTS,
  scopeNote: "These rows are granted app-only consent for the named client; the application's requested permissions (requiredResourceAccess) are declared on the application object and are not shown here",
  emptyNote: "0 app-role assignments matched; the absence of results is the answer, not an error",
};

async function listCollection(
  shape: CollectionShape,
  session: GraphSession,
  flags: GrantFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  checkReadOperation(operation, help);
  const client = String(flags["service-principal"] ?? "");
  if (!client.trim()) throw new AxiError("--service-principal needs the client service-principal object ID", "VALIDATION_ERROR", [help]);
  const cursor = flags.cursor === undefined ? undefined : String(flags.cursor);
  if (cursor !== undefined && !cursor.trim()) throw new AxiError("--cursor needs the opaque cursor from a partial result", "VALIDATION_ERROR", [help]);
  const saved = cursor === undefined ? undefined : session.cursorQuery(operation, cursor);
  const savedSelect = saved?.$select;
  const { select, fields } = selectedFields(flags,
    savedSelect === undefined ? shape.defaultSelect : fieldList(savedSelect, "select", shape.known, shape.knownList, help),
    shape.known, shape.knownList, help);
  const savedFilter = saved?.$filter;
  const filter = flags.filter === undefined ? savedFilter : String(flags.filter);
  const scopes = scopesFor(flags, profile, shape.defaultScopes, help);
  const full = flags.full === true;
  // These collections document $select/$filter only, so --filter passes
  // through as plain $filter with no $count or ConsistencyLevel.
  const query: Record<string, string> = { $select: select.join(",") };
  if (filter !== undefined) query.$filter = filter;
  const args: CollectArgs = {
    profile,
    operation,
    params: { "servicePrincipal-id": client },
    query,
    scopes,
  };
  if (cursor !== undefined) args.cursor = cursor;
  if (flags.all === true) {
    if (flags.limit !== undefined) throw new AxiError("--limit and --all cannot be combined", "VALIDATION_ERROR", [help]);
  } else {
    args.limit = flags.limit === undefined ? 100 : Number(flags.limit);
  }
  const result = await withGuidance(shape.denialHints, () => session.collect(args));
  const effectiveFlags: GrantFlags = { ...flags, select: result.query.$select ?? shape.defaultSelect.join(",") };
  if (result.query.$filter !== undefined) effectiveFlags.filter = result.query.$filter;
  const rows: Record<string, unknown>[] = [];
  let truncated = false;
  for (const row of result.value) {
    const projected = project(row, fields, full);
    rows.push(projected.row);
    truncated = truncated || projected.truncated;
  }
  const clientHint = `mg-axi entra service-principal show --id ${shellValue(client)} ${profileHint(profileName)}`;
  const truncationHints = truncated ? [fullHint(shape.command, effectiveFlags, profileName)] : [];
  const standing = [shape.scopeNote, clientHint];
  if (!result.complete) {
    return {
      [shape.key]: rows,
      count: { returned: rows.length, complete: false, reason: result.reason },
      cursor: result.cursor,
      help: [...truncationHints, `Resume losslessly with the same flags plus --cursor <cursor-from-output> ${profileHint(profileName)}`, ...standing],
    };
  }
  const count = { returned: rows.length, complete: true };
  if (!rows.length) {
    return { [shape.key]: rows, count, help: [shape.emptyNote, ...standing] };
  }
  return { [shape.key]: rows, count, help: [...truncationHints, ...standing] };
}

export async function listOAuth2Grants(
  session: GraphSession,
  flags: GrantFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(OAUTH2_GRANT_LIST, session, flags, profile, operation, help, profileName);
}

export async function listAppRoleAssignments(
  session: GraphSession,
  flags: GrantFlags,
  profile: AnyProfile,
  operation: SessionOperation,
  help: string,
  profileName: string,
): Promise<Record<string, unknown>> {
  return listCollection(APP_ROLE_LIST, session, flags, profile, operation, help, profileName);
}
