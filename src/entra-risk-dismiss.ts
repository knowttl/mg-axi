import { AxiError } from "axi-sdk-js";
import type { GraphSession, SessionOperation } from "./graph-session.js";
import type { MutationCoordinator } from "./mutations.js";
import type { AnyProfile } from "./profiles.js";

// WRITE-05: dismiss the risk of one explicitly named user through an action
// write leaf (`mg-axi entra risky-user dismiss --user <risky-user-id>`), via
// POST /identityProtection/riskyUsers/dismiss with a single-element
// { userIds: [...] } body. There is no bulk form: a comma in --user is a
// usage error and unknown flags (including any plural --user-ids) fail in
// the strict catalogue before credentials. Preview reads the user's current
// risk state through the READ-06 risky-user show route; an already-dismissed
// user is a no-op. Dismissal answers 204 with an empty body, so after a send
// the command rereads the user and reports a conflict when the state is not
// dismissed, mirroring WRITE-02.
//
// Reviewed against the v1.0 riskyuser-dismiss operation documentation on
// REVIEWED_ON: POST /identityProtection/riskyUsers/dismiss carries
// { userIds: [...] } and answers 204 No Content. The least-privileged
// permission is D/A IdentityRiskyUser.ReadWrite.All; delegated callers
// additionally need Security Administrator (least-privileged role) and the
// riskyUsers API requires a Microsoft Entra ID P2 licence. Dismissal sets
// the user's risk level to none; it is not remediation - it resets no
// credential and revokes no session. Every gate in src/mutations.ts stays in
// force through the coordinator: hand-enabled profile, immutable scope,
// preview, --execute, durable journal intent/outcome, no replay,
// coordinator-only authorization and read-only enforced at send time.

export const REVIEWED_ON = "2026-10-04";

// Scope name a profile hand-enables in its writes.operations allowlist.
export const RISKY_USER_DISMISS_OPERATION = "entra.risky-user.dismiss";

export const RISKY_USER_DISMISS_WRITE_SCOPES = ["https://graph.microsoft.com/IdentityRiskyUser.ReadWrite.All"];
export const RISKY_USER_DISMISS_READ_SCOPES = ["https://graph.microsoft.com/IdentityRiskyUser.Read.All"];

const RISK_STATE_SELECT = "id,userPrincipalName,riskLevel,riskState,riskDetail,riskLastUpdatedDateTime";

const PERMISSION_GUIDANCE =
  "Needs D/A IdentityRiskyUser.ReadWrite.All (least-privileged per riskyuser-dismiss v1.0, rechecked 2026-10-04)";
const ROLE_GUIDANCE =
  "Delegated callers additionally need Security Administrator (least-privileged role for this action)";
const LICENCE_GUIDANCE =
  "The riskyUsers API requires a Microsoft Entra ID P2 licence; without it Graph denies the dismissal";
const REMEDIATION_NOTE =
  "Dismissal is not remediation: it records the user's risk as dismissed without resetting credentials or revoking sessions";

export type RiskDismissFlags = Record<string, string | boolean>;

export type RiskDismissState = {
  id: string;
  userPrincipalName: unknown;
  riskLevel: unknown;
  riskState: unknown;
  riskDetail: unknown;
  riskLastUpdatedDateTime: unknown;
};

export function riskDismissDefinition(user: string) {
  const target = user.trim();
  return {
    operation: RISKY_USER_DISMISS_OPERATION,
    method: "POST" as const,
    version: "v1.0",
    path: "/v1.0/identityProtection/riskyUsers/dismiss",
    effect: "disruptive" as const,
    target,
    payload: { userIds: [target] },
  };
}

function parseSingleUser(raw: unknown, help: string): string {
  const user = String(raw ?? "").trim();
  if (!user) throw new AxiError("--user needs the risky-user ID to dismiss", "VALIDATION_ERROR", [help]);
  if (user.includes(",")) {
    throw new AxiError("--user takes exactly one risky-user ID; there is no bulk dismissal form", "VALIDATION_ERROR", [
      help,
      "Dismiss one user per command; each dismissal is its own audited intent",
    ]);
  }
  return user;
}

async function readRiskState(
  session: GraphSession,
  profile: AnyProfile,
  readOperation: SessionOperation,
  user: string,
): Promise<RiskDismissState> {
  const raw = await session.execute({
    profile,
    operation: readOperation,
    params: { "riskyUser-id": user },
    query: { $select: RISK_STATE_SELECT },
    ...(profile.mode === "application" ? {} : { scopes: [...RISKY_USER_DISMISS_READ_SCOPES] }),
  });
  if (raw !== null && typeof raw === "object" && !Array.isArray(raw) && "id" in raw && typeof raw.id === "string" && raw.id.length > 0) {
    const row = raw as Record<string, unknown>;
    const id: string = raw.id;
    return {
      id,
      userPrincipalName: row.userPrincipalName ?? null,
      riskLevel: row.riskLevel ?? null,
      riskState: row.riskState ?? null,
      riskDetail: row.riskDetail ?? null,
      riskLastUpdatedDateTime: row.riskLastUpdatedDateTime ?? null,
    };
  }
  throw new AxiError("Graph returned a risky user without an object ID", "GRAPH_ERROR", [
    "Read back the user with mg-axi entra risky-user show before attempting a dismissal",
  ]);
}

function executeHint(user: string, profileName: string): string {
  const shell = (value: string): string =>
    /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
  return `mg-axi entra risky-user dismiss --user ${shell(user)} --execute`
    + ` --confirm ${shell(user)}`
    + ` --profile ${shell(profileName)}`;
}

function showHint(user: string, profileName: string): string {
  const shell = (value: string): string =>
    /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
  return `mg-axi entra risky-user show --id ${shell(user)} --profile ${shell(profileName)}`;
}

export async function dismissRiskyUser(args: {
  session: GraphSession;
  coordinator: MutationCoordinator;
  flags: RiskDismissFlags;
  profile: AnyProfile;
  profileName: string;
  readOperation: SessionOperation;
  help: string;
}): Promise<Record<string, unknown>> {
  const { session, coordinator, flags, profile, profileName, readOperation, help } = args;
  const user = parseSingleUser(flags.user, help);
  const version = String(flags["api-version"] ?? "v1.0");
  if (version !== "v1.0") {
    throw new AxiError(
      `--api-version ${version} cannot mutate: WRITE-05 binds POST /identityProtection/riskyUsers/dismiss to v1.0; beta writes stay blocked`,
      "VALIDATION_ERROR",
      [help],
    );
  }
  if (flags.scopes !== undefined) {
    throw new AxiError("--scopes is not accepted: dismissal requests only the documented write scope", "VALIDATION_ERROR", [
      help,
      "Delegated dismissal requests IdentityRiskyUser.ReadWrite.All; application profiles use the configured Graph .default audience",
    ]);
  }
  const requestedDefinition = riskDismissDefinition(user);
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
  // The target is verified as a user through the READ-06 route: a failed or
  // malformed read blocks the dismissal before anything is sent.
  const state = await readRiskState(session, profile, readOperation, user);
  const id = state.id;
  const definition = { ...requestedDefinition, payload: { userIds: [id] } };
  const preview = coordinator.preview(definition);
  const dismissed = state.riskState === "dismissed";
  const view = {
    operation: definition.operation,
    target: definition.target,
    method: definition.method,
    url: preview.url,
    effect: definition.effect,
    current: state,
    desired: "dismissed",
    noop: dismissed,
  };
  if (!execute) {
    return {
      preview: view,
      help: [
        dismissed
          ? "Already dismissed; no dismissal is needed"
          : executeHint(user, profileName),
        ...(!dismissed
          ? ["Dismissing risk is disruptive: the --execute run needs the typed --confirm value shown above"]
          : []),
        REMEDIATION_NOTE,
        LICENCE_GUIDANCE,
      ],
    };
  }
  const result = await coordinator.execute(definition, {
    execute: true,
    confirm: String(flags.confirm),
    ...(profile.mode === "application" ? {} : { scopes: [...RISKY_USER_DISMISS_WRITE_SCOPES] }),
    readState: async () => (await readRiskState(session, profile, readOperation, id)).riskState,
    isNoop: currentState => currentState === "dismissed",
  });
  if (result.kind === "failed") {
    throw new AxiError(
      `Graph refused the risk dismissal (status ${result.status})`,
      "GRAPH_ERROR",
      [
        ...(result.status === 403 ? [PERMISSION_GUIDANCE, ROLE_GUIDANCE, LICENCE_GUIDANCE] : []),
        `Audit ${result.auditId} recorded the refusal; read back '${id}' and never replay this intent`,
      ],
    );
  }
  if (result.kind === "unknown") {
    throw new AxiError(`Dismissal of '${id}' may or may not have been applied (audit ${result.auditId}); never replay this intent`, "OUTCOME_UNKNOWN", [
      `Audit ${result.auditId} recorded the uncertain outcome; read back '${id}' with ${showHint(id, profileName)} before doing anything else`,
      "Never replay this intent",
    ]);
  }
  // Anything but success, failure or uncertainty is the verified no-op:
  // dry-run needs execute !== true, always set above.
  if (result.kind !== "success") {
    return { noop: true, dismissal: { user: { id, riskState: "dismissed" } } };
  }
  // Success carries no proof: dismissal answers 204 with an empty body, so
  // only the reread decides between dismissed and conflict.
  let verified: RiskDismissState | null = null;
  try {
    verified = await readRiskState(session, profile, readOperation, id);
  } catch {
    throw new AxiError(
      `Dismissal of '${user}' was sent (audit ${result.auditId}) but the verification read failed; the outcome is unknown`,
      "OUTCOME_UNKNOWN",
      [`Read back '${id}' before doing anything else; never replay this intent`],
    );
  }
  if (verified.riskState !== "dismissed") {
    throw new AxiError(
      `Dismissal of '${user}' conflicts: the reread shows riskState ${JSON.stringify(verified.riskState)} instead of dismissed (audit ${result.auditId})`,
      "WRITE_CONFLICT",
      [`Read back '${id}' with ${showHint(id, profileName)} before doing anything else`, "Never replay this intent"],
    );
  }
  return { dismissal: { user: verified }, auditId: result.auditId };
}
