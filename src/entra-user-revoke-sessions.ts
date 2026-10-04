import { AxiError } from "axi-sdk-js";
import type { GraphSession, SessionOperation } from "./graph-session.js";
import { encodeGraphPathSegment } from "./graph-session.js";
import type { MutationCoordinator } from "./mutations.js";
import type { AnyProfile } from "./profiles.js";

// WRITE-03: revoke one user's sign-in sessions through an action write leaf
// (`mg-axi entra user revoke-sessions --user <id>`), via
// POST /users/{id}/revokeSignInSessions with no request body. Graph resets
// the user's signInSessionsValidFromDateTime, invalidating issued refresh
// tokens and browser session cookies so the user must sign in again, and
// answers 2xx with {"value": true}; that accepted response is the proof, so
// unlike the 204 account update there is no verification reread.
//
// Reviewed against the v1.0 user-revokeSignInSessions operation
// documentation on REVIEWED_ON. The least-privileged permission is
// User.RevokeSessions.All in both delegated and application modes; the
// documentation names no administrator-role prerequisite for this action, so
// denial guidance names only the permission pair. This is an action, not a
// desired-state write: the preview says what will happen and that it cannot
// be undone (there is no rollback), never a state diff. Two Microsoft-stated
// limits ride along in the preview: token revocation can lag a few minutes
// after the call returns, and external users are unaffected because they
// sign in through their home tenant. A timeout or 5xx after send is
// OUTCOME_UNKNOWN with no automatic replay. Every gate in src/mutations.ts
// stays in force through the coordinator: hand-enabled profile, immutable
// scope, preview, --execute, typed confirmation, durable journal
// intent/outcome, no replay, coordinator-only authorization and read-only
// enforced at send time.

export const REVIEWED_ON = "2026-10-04";

// Scope name a profile hand-enables in its writes.operations allowlist.
export const USER_REVOKE_SESSIONS_OPERATION = "entra.user.revokeSessions";

export const USER_REVOKE_SESSIONS_SCOPES = [
  "https://graph.microsoft.com/User.RevokeSessions.All",
];
export const USER_REVOKE_READ_SCOPES = ["https://graph.microsoft.com/User.ReadBasic.All"];

const PERMISSION_GUIDANCE =
  "Needs D/A User.RevokeSessions.All (least-privileged per user-revokeSignInSessions v1.0, rechecked 2026-10-04)";

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export type UserRevokeSessionsFlags = Record<string, string | boolean>;

export function userRevokeSessionsDefinition(user: string) {
  const target = user.trim();
  return {
    operation: USER_REVOKE_SESSIONS_OPERATION,
    method: "POST" as const,
    version: "v1.0",
    path: `/v1.0/users/${encodeGraphPathSegment(target)}/revokeSignInSessions`,
    effect: "disruptive" as const,
    target,
  };
}

// The target must be an accessible user before preview and again before the
// single send; anything else blocks the operation. Returns the bound object
// ID for the mutation path.
async function verifyUserTarget(
  session: GraphSession,
  profile: AnyProfile,
  readOperation: SessionOperation,
  user: string,
): Promise<string> {
  const raw = await session.execute({
    profile,
    operation: readOperation,
    params: { "user-id": user },
    query: { $select: "id" },
    ...(profile.mode === "application" ? {} : { scopes: [...USER_REVOKE_READ_SCOPES] }),
  });
  const row = raw as Record<string, unknown> | null;
  if (row === null || typeof row !== "object" || Array.isArray(row)
    || typeof row["id"] !== "string" || row["id"].toLowerCase() !== user.toLowerCase()
    || (row["@odata.type"] !== undefined && row["@odata.type"] !== "#microsoft.graph.user")) {
    throw new AxiError(`Graph did not establish user identity for ${user}`, "GRAPH_ERROR", [
      "Verify --user is an accessible user object ID; non-user objects have no sessions to revoke",
    ]);
  }
  return row["id"];
}

function executeHint(user: string, profileName: string): string {
  const shell = (value: string): string =>
    /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
  return `mg-axi entra user revoke-sessions --user ${shell(user)} --execute`
    + ` --confirm ${shell(user)}`
    + ` --profile ${shell(profileName)}`;
}

function showHint(user: string, profileName: string, profile: AnyProfile): string {
  const shell = (value: string): string =>
    /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
  return `mg-axi entra user show --id ${shell(user)} --profile ${shell(profileName)} --select id`
    + (profile.mode === "application" ? "" : ` --scopes ${USER_REVOKE_READ_SCOPES.join(",")}`);
}

const ACTION_TEXT =
  "Revoke every sign-in session for the target user: Graph resets signInSessionsValidFromDateTime, "
  + "invalidating issued refresh tokens and browser session cookies, and the user must sign in again to every consented application.";

const LIMITATIONS = [
  "Token revocation can lag a few minutes after the call returns; a 2xx response does not promise immediate universal session termination",
  "External users are unaffected: they sign in through their home tenant, which this call cannot revoke",
  "This action cannot be undone; there is no rollback",
];

export async function revokeUserSessions(args: {
  session: GraphSession;
  coordinator: MutationCoordinator;
  flags: UserRevokeSessionsFlags;
  profile: AnyProfile;
  profileName: string;
  readOperation: SessionOperation;
  help: string;
}): Promise<Record<string, unknown>> {
  const { session, coordinator, flags, profile, profileName, readOperation, help } = args;
  const user = String(flags.user ?? "").trim();
  if (!user) throw new AxiError("--user needs the user object ID", "VALIDATION_ERROR", [help]);
  if (!GUID.test(user)) {
    throw new AxiError(`--user ${user} is not a user object ID`, "VALIDATION_ERROR", [
      help,
      "Pass the user object ID (GUID); UPNs are not resolved by this command",
    ]);
  }
  const version = String(flags["api-version"] ?? "v1.0");
  if (version !== "v1.0") {
    throw new AxiError(
      `--api-version ${version} cannot mutate: WRITE-03 binds POST /users/{id}/revokeSignInSessions to v1.0; beta writes stay blocked`,
      "VALIDATION_ERROR",
      [help],
    );
  }
  if (flags.scopes !== undefined) {
    throw new AxiError(
      "This command requests only its documented write scope; caller-supplied --scopes is unavailable",
      "VALIDATION_ERROR",
      [help, "Delegated revocation requests https://graph.microsoft.com/User.RevokeSessions.All"],
    );
  }
  const requestedDefinition = userRevokeSessionsDefinition(user);
  const execute = flags.execute === true;
  // Gates first: preview refuses read-only, unscoped and beta-bound profiles
  // before any read, credential or transport.
  coordinator.preview(requestedDefinition);
  if (execute) {
    if (flags.confirm === undefined) {
      throw new AxiError(
        `blocked: disruptive POST needs --confirm '${requestedDefinition.target}'`,
        "CONFIRM_REQUIRED",
        [`Re-run with --confirm '${requestedDefinition.target}'`],
      );
    }
    if (String(flags.confirm) !== requestedDefinition.target) {
      throw new AxiError(
        `blocked: --confirm '${flags.confirm}' does not match target '${requestedDefinition.target}'`,
        "CONFIRM_MISMATCH",
        [`Re-run with --confirm '${requestedDefinition.target}'`],
      );
    }
  }
  const id = await verifyUserTarget(session, profile, readOperation, user);
  const definition = { ...requestedDefinition, path: `/v1.0/users/${encodeGraphPathSegment(id)}/revokeSignInSessions` };
  const preview = coordinator.preview(definition);
  const state = {
    operation: definition.operation,
    target: definition.target,
    method: definition.method,
    url: preview.url,
    effect: definition.effect,
    action: ACTION_TEXT,
    reversible: false,
    limitations: [...LIMITATIONS],
    noop: false,
  };
  if (!execute) {
    return {
      preview: state,
      help: [
        executeHint(user, profileName),
        `Revoking sessions is disruptive: the --execute run needs the typed --confirm value shown above`,
        showHint(user, profileName, profile),
      ],
    };
  }
  const result = await coordinator.execute(definition, {
    execute: true,
    confirm: String(flags.confirm),
    ...(profile.mode === "application" ? {} : { scopes: [...USER_REVOKE_SESSIONS_SCOPES] }),
    // An action always sends: no desired state exists to satisfy early, and
    // the fresh read below re-verifies the user before the single send.
    readState: async () => verifyUserTarget(session, profile, readOperation, id),
  });
  if (result.kind === "failed") {
    throw new AxiError(
      `Graph refused the session revocation (status ${result.status})`,
      "GRAPH_ERROR",
      [
        ...(result.status === 403 ? [PERMISSION_GUIDANCE] : []),
        `Audit ${result.auditId} recorded the refusal; read back '${id}' and never replay this intent`,
      ],
    );
  }
  if (result.kind === "unknown") {
    throw new AxiError(`Revocation for '${id}' may or may not have been applied (audit ${result.auditId}); never replay this intent`, "OUTCOME_UNKNOWN", [
      `Audit ${result.auditId} recorded the uncertain outcome; read back '${id}' with ${showHint(id, profileName, profile)} before doing anything else`,
      "Never replay this intent",
    ]);
  }
  if (result.kind !== "success") {
    // Unreachable: execute is always true above (no dry-run) and no
    // desired-state check exists to report a no-op. Fail closed rather than
    // inventing an outcome.
    throw new AxiError(`Revocation for '${id}' returned an unexpected outcome; read back the target before doing anything else`, "OUTCOME_UNKNOWN", [
      "Never replay this intent",
    ]);
  }
  return { user: { id, sessionsRevoked: true }, auditId: result.auditId };
}
