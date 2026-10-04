import { AxiError } from "axi-sdk-js";
import type { GraphSession, SessionOperation } from "./graph-session.js";
import { encodeGraphPathSegment } from "./graph-session.js";
import type {
  MutationCoordinator,
  MutationTransport,
  MutationTransportRequest,
  MutationTransportResponse,
} from "./mutations.js";
import type { AnyProfile } from "./profiles.js";

// WRITE-02: enable or disable one user's account through a desired-state
// write leaf (`mg-axi entra user update --user <id> --account-enabled
// true|false`), via PATCH /users/{id} with only { accountEnabled } sent.
// Preview reads the current value through the READ-01 user route and shows
// the diff; an already-desired value is a no-op; after the PATCH a reread
// reports a conflict when the value is not what was sent.
//
// Reviewed against the v1.0 user-update operation documentation on
// REVIEWED_ON: PATCH /users/{id | userPrincipalName} carries only the
// properties to update and answers 204 No Content, so the reread is the only
// proof the change landed. The least-privileged pair for accountEnabled is
// User.EnableDisableAccount.All plus User.Read.All in both delegated and
// application modes. Sensitive-target hierarchy: delegated callers need
// Privileged Authentication Administrator for admin targets and generally
// must hold a higher-privileged administrator role than the target; app-only
// callers need the pair plus a higher-privileged admin role assignment (Who
// can perform sensitive actions). Graph enforces that hierarchy and a 403
// never says which prerequisite is missing, so this module surfaces the pair
// and the hierarchy as denial guidance and never invents a local role
// verdict. Every gate in src/mutations.ts stays in force through the
// coordinator: hand-enabled profile, immutable scope, preview, --execute,
// durable journal intent/outcome, no replay, coordinator-only authorization
// and read-only enforced at send time.

export const REVIEWED_ON = "2026-10-04";

// Scope name a profile hand-enables in its writes.operations allowlist.
export const USER_ACCOUNT_UPDATE_OPERATION = "entra.user.update";

export const USER_ACCOUNT_WRITE_SCOPES = [
  "https://graph.microsoft.com/User.EnableDisableAccount.All",
  "https://graph.microsoft.com/User.Read.All",
];
export const USER_ACCOUNT_READ_SCOPES = ["https://graph.microsoft.com/User.Read.All"];

const PERMISSION_GUIDANCE =
  "Needs D/A User.EnableDisableAccount.All plus User.Read.All for accountEnabled (least-privileged pair per user-update v1.0, rechecked 2026-10-04)";
const ROLE_GUIDANCE =
  "Sensitive-target hierarchy: delegated callers need Privileged Authentication Administrator for admin targets and must generally outrank the target; app-only callers need the pair plus a higher-privileged admin role assignment (Who can perform sensitive actions)";

export type UserUpdateFlags = Record<string, string | boolean>;

export function parseAccountEnabled(raw: unknown, help: string): boolean {
  const text = String(raw ?? "").trim().toLowerCase();
  if (text === "true") return true;
  if (text === "false") return false;
  throw new AxiError("--account-enabled must be true or false", "VALIDATION_ERROR", [help]);
}

export function userAccountDefinition(user: string, desired: boolean) {
  const target = user.trim();
  return {
    operation: USER_ACCOUNT_UPDATE_OPERATION,
    method: "PATCH" as const,
    version: "v1.0",
    path: `/v1.0/users/${encodeGraphPathSegment(target)}`,
    effect: "disruptive" as const,
    target,
    payload: { accountEnabled: desired },
  };
}

async function readUserAccount(
  session: GraphSession,
  profile: AnyProfile,
  readOperation: SessionOperation,
  user: string,
): Promise<{ id: string; accountEnabled: boolean | null }> {
  const raw = await session.execute({
    profile,
    operation: readOperation,
    params: { "user-id": user },
    query: { $select: "id,accountEnabled" },
    ...(profile.mode === "application" ? {} : { scopes: [...USER_ACCOUNT_READ_SCOPES] }),
  });
  if (raw !== null && typeof raw === "object" && !Array.isArray(raw) && "id" in raw && typeof raw.id === "string" && raw.id.length > 0) {
    const value = (raw as Record<string, unknown>).accountEnabled;
    // Null, missing or malformed is unverifiable, never a desired state.
    return { id: raw.id, accountEnabled: typeof value === "boolean" ? value : null };
  }
  throw new AxiError("Graph returned a user without an object ID", "GRAPH_ERROR", [
    "Read back the user before attempting an account update",
  ]);
}

function executeHint(user: string, desired: boolean, profileName: string): string {
  const shell = (value: string): string =>
    /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
  return `mg-axi entra user update --user ${shell(user)} --account-enabled ${desired} --execute`
    + ` --confirm ${shell(user)}`
    + ` --profile ${shell(profileName)}`;
}

function showHint(user: string, profileName: string): string {
  const shell = (value: string): string =>
    /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
  return `mg-axi entra user show --id ${shell(user)} --select id,accountEnabled --profile ${shell(profileName)}`;
}

export async function updateUserAccount(args: {
  session: GraphSession;
  coordinator: MutationCoordinator;
  flags: UserUpdateFlags;
  profile: AnyProfile;
  profileName: string;
  readOperation: SessionOperation;
  help: string;
}): Promise<Record<string, unknown>> {
  const { session, coordinator, flags, profile, profileName, readOperation, help } = args;
  const user = String(flags.user ?? "").trim();
  if (!user) throw new AxiError("--user needs a user ID or UPN", "VALIDATION_ERROR", [help]);
  const desired = parseAccountEnabled(flags["account-enabled"], help);
  const version = String(flags["api-version"] ?? "v1.0");
  if (version !== "v1.0") {
    throw new AxiError(
      `--api-version ${version} cannot mutate: WRITE-02 binds PATCH /users/{id} to v1.0; beta writes stay blocked`,
      "VALIDATION_ERROR",
      [help],
    );
  }
  const requestedDefinition = userAccountDefinition(user, desired);
  const execute = flags.execute === true;
  // Gates first: preview refuses read-only, unscoped and beta-bound profiles
  // before any read, credential or transport.
  coordinator.preview(requestedDefinition);
  if (execute) {
    if (flags.confirm === undefined) {
      throw new AxiError(
        `blocked: disruptive PATCH needs --confirm '${requestedDefinition.target}'`,
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
  const account = await readUserAccount(session, profile, readOperation, user);
  const id = account.id;
  const definition = { ...requestedDefinition, path: `/v1.0/users/${encodeGraphPathSegment(id)}` };
  const preview = coordinator.preview(definition);
  const current = account.accountEnabled;
  const state = {
    operation: definition.operation,
    target: definition.target,
    method: definition.method,
    url: preview.url,
    effect: definition.effect,
    current,
    desired,
    noop: current === desired,
  };
  if (!execute) {
    return {
      preview: state,
      help: [
        state.noop
          ? "Already in the desired state; no update is needed"
          : executeHint(user, desired, profileName),
        ...(!state.noop
          ? ["Changing account access is disruptive: the --execute run needs the typed --confirm value shown above"]
          : []),
      ],
    };
  }
  const result = await coordinator.execute(definition, {
    execute: true,
    confirm: String(flags.confirm),
    ...(profile.mode === "application" ? {} : { scopes: [...USER_ACCOUNT_WRITE_SCOPES] }),
    readState: async () => (await readUserAccount(session, profile, readOperation, id)).accountEnabled,
    isNoop: currentState => currentState === desired,
  });
  if (result.kind === "failed") {
    throw new AxiError(
      `Graph refused the account update (status ${result.status})`,
      "GRAPH_ERROR",
      [
        ...(result.status === 403 ? [PERMISSION_GUIDANCE, ROLE_GUIDANCE] : []),
        `Audit ${result.auditId} recorded the refusal; read back '${id}' and never replay this intent`,
      ],
    );
  }
  if (result.kind === "unknown") {
    throw new AxiError(`Update of '${id}' may or may not have been applied (audit ${result.auditId}); never replay this intent`, "OUTCOME_UNKNOWN", [
      `Audit ${result.auditId} recorded the uncertain outcome; read back '${id}' with ${showHint(id, profileName)} before doing anything else`,
      "Never replay this intent",
    ]);
  }
  // Anything but success, failure or uncertainty is the verified no-op:
  // dry-run needs execute !== true, always set above.
  if (result.kind !== "success") {
    return { noop: true, user: { id, accountEnabled: desired } };
  }
  // Success carries no proof: user-update answers 204 with an empty body, so
  // only the reread decides between updated and conflict.
  let verified: boolean | null;
  try {
    verified = (await readUserAccount(session, profile, readOperation, id)).accountEnabled;
  } catch {
    throw new AxiError(
      `Update of '${user}' was sent (audit ${result.auditId}) but the verification read failed; the outcome is unknown`,
      "OUTCOME_UNKNOWN",
      [`Read back '${id}' before doing anything else; never replay this intent`],
    );
  }
  if (verified !== desired) {
    throw new AxiError(
      `Update of '${user}' conflicts: the reread shows accountEnabled ${verified === null ? "unreadable" : String(verified)} instead of the sent ${String(desired)} (audit ${result.auditId})`,
      "WRITE_CONFLICT",
      [`Read back '${id}' with ${showHint(id, profileName)} before doing anything else`, "Never replay this intent"],
    );
  }
  return { user: { id, accountEnabled: desired }, auditId: result.auditId };
}

// Production mutation transport: plain HTTPS with redirects held for manual
// refusal. The coordinator re-authorizes the destination before credentials
// and before handoff; a redirect answer therefore fails closed downstream
// instead of following anywhere.
export const fetchMutationTransport: MutationTransport = async (
  request: MutationTransportRequest,
): Promise<MutationTransportResponse> => {
  const response = await fetch(request.url, {
    method: request.method,
    headers: request.headers,
    ...(request.body !== undefined ? { body: request.body } : {}),
    ...(request.signal ? { signal: request.signal } : {}),
    redirect: "manual",
  });
  const headers: Record<string, string> = {};
  response.headers.forEach((value, key) => {
    headers[key] = value;
  });
  return { status: response.status, headers, body: await response.text() };
};
