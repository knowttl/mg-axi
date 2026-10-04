import { AxiError } from "axi-sdk-js";
import type { GraphSession, SessionOperation } from "./graph-session.js";
import { encodeGraphPathSegment } from "./graph-session.js";
import type { MutationCoordinator } from "./mutations.js";
import type { AnyProfile } from "./profiles.js";

// WRITE-04: update one Conditional Access policy through a reviewed-field
// write leaf (`mg-axi entra conditional-access policy update --id
// <policy-id> --state enabled|...`), via PATCH
// /identity/conditionalAccess/policies/{id} with only explicitly reviewed
// fields sent. Preview reads the current policy through the READ-03 policy
// show route and shows the current-versus-proposed diff; an already-desired
// value set is a no-op; every execute is disruptive and needs the typed
// policy-ID confirmation.
//
// Reviewed against the v1.0 conditionalaccesspolicy-update operation
// documentation on REVIEWED_ON: PATCH carries only the properties to update
// and answers 204 No Content, so the coordinator outcome is the only proof.
// The least-privileged pair is Policy.Read.All plus
// Policy.ReadWrite.ConditionalAccess in both delegated and application
// modes. Delegated callers additionally need Conditional Access
// Administrator or Security Administrator. CA needs P1; risk-based CA needs
// P2. Graph documents no ETag or If-Match precondition for this endpoint,
// so the command sends none and promises no concurrency protection: fresh
// reads do not make the update atomic. Every gate in src/mutations.ts stays
// in force through the coordinator: hand-enabled profile, immutable scope,
// preview, --execute, durable journal intent/outcome, no replay,
// coordinator-only authorization and read-only enforced at send time.

export const REVIEWED_ON = "2026-10-04";

// Scope name a profile hand-enables in its writes.operations allowlist.
export const CA_POLICY_UPDATE_OPERATION = "entra.conditional-access.policy.update";

export const CA_POLICY_UPDATE_SCOPES = [
  "https://graph.microsoft.com/Policy.Read.All",
  "https://graph.microsoft.com/Policy.ReadWrite.ConditionalAccess",
];
export const CA_POLICY_UPDATE_READ_SCOPES = ["https://graph.microsoft.com/Policy.Read.All"];

// The only top-level policy properties this slice may send. Anything else
// (id, createdDateTime, modifiedDateTime, ...) has no flag and can never
// reach the PATCH body.
export const CA_POLICY_WRITABLE_FIELDS = ["displayName", "state", "conditions", "grantControls", "sessionControls"] as const;
export type CaPolicyWritableField = (typeof CA_POLICY_WRITABLE_FIELDS)[number];

export const CA_POLICY_STATES = ["enabled", "enabledForReportingButNotEnforced", "disabled"] as const;

const FLAG_TO_FIELD = {
  "display-name": "displayName",
  "state": "state",
  "conditions": "conditions",
  "grant-controls": "grantControls",
  "session-controls": "sessionControls",
} as const satisfies Record<string, CaPolicyWritableField>;

const ENFORCEMENT_FIELDS: readonly CaPolicyWritableField[] = ["state", "conditions", "grantControls", "sessionControls"];

const PERMISSION_GUIDANCE =
  "Needs D/A Policy.Read.All plus Policy.ReadWrite.ConditionalAccess (least-privileged pair per conditionalaccesspolicy-update v1.0, rechecked 2026-10-04)";
const ROLE_GUIDANCE =
  "Delegated callers additionally need Conditional Access Administrator or Security Administrator; CA needs P1, risk-based CA needs P2";
const CONCURRENCY_NOTE =
  "Graph documents no ETag or If-Match precondition for this endpoint, so fresh reads do not make the update atomic: another actor can change the policy between the read and the PATCH. The command sends no If-Match and promises no concurrency protection";

export type CaPolicyUpdateFlags = Record<string, string | boolean>;

export type LockoutAssessment =
  | { available: false; reason: string }
  | { available: true; level: "none" | "elevated" | "refused"; findings: string[] };

function canonicalize(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (value !== null && typeof value === "object") {
    return Object.fromEntries(
      Object.keys(value as Record<string, unknown>)
        .sort()
        .map(key => [key, canonicalize((value as Record<string, unknown>)[key])]),
    );
  }
  return value;
}

function sameValue(current: unknown, proposed: unknown): boolean {
  return JSON.stringify(canonicalize(current)) === JSON.stringify(canonicalize(proposed));
}

function parseJsonObject(raw: string, flag: string, help: string): Record<string, unknown> {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw) as unknown;
  } catch {
    throw new AxiError(`--${flag} must be a JSON object`, "VALIDATION_ERROR", [help]);
  }
  if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new AxiError(`--${flag} must be a JSON object`, "VALIDATION_ERROR", [help]);
  }
  return parsed as Record<string, unknown>;
}

// Maps exactly the reviewed-field flags to PATCH properties. Unreviewed
// Graph properties have no flag, so resolveCommand already refuses them as
// unknown flags before credentials; this parser additionally refuses empty,
// mistyped and unparsable reviewed values.
export function parsePolicyUpdate(raw: CaPolicyUpdateFlags, help: string): Partial<Record<CaPolicyWritableField, unknown>> {
  const payload: Partial<Record<CaPolicyWritableField, unknown>> = {};
  for (const [flag, field] of Object.entries(FLAG_TO_FIELD)) {
    const value = raw[flag];
    if (value === undefined) continue;
    const text = String(value);
    if (field === "displayName") {
      if (!text.trim()) throw new AxiError("--display-name needs a non-empty policy name", "VALIDATION_ERROR", [help]);
      payload[field] = text;
    } else if (field === "state") {
      if (!(CA_POLICY_STATES as readonly string[]).includes(text)) {
        throw new AxiError(`--state must be ${CA_POLICY_STATES.join("|")}`, "VALIDATION_ERROR", [help]);
      }
      payload[field] = text;
    } else {
      payload[field] = parseJsonObject(text, flag, help);
    }
  }
  if (Object.keys(payload).length === 0) {
    throw new AxiError(
      "Policy update needs at least one reviewed field: --display-name, --state, --conditions, --grant-controls or --session-controls",
      "VALIDATION_ERROR",
      [help],
    );
  }
  return payload;
}

export function caPolicyDefinition(policyId: string, payload: Partial<Record<CaPolicyWritableField, unknown>>) {
  const target = policyId.trim();
  return {
    operation: CA_POLICY_UPDATE_OPERATION,
    method: "PATCH" as const,
    version: "v1.0",
    path: `/v1.0/identity/conditionalAccess/policies/${encodeGraphPathSegment(target)}`,
    effect: "disruptive" as const,
    target,
    payload: { ...payload },
  };
}

type PolicyRecord = Record<string, unknown>;

async function readPolicy(
  session: GraphSession,
  profile: AnyProfile,
  readOperation: SessionOperation,
  policyId: string,
): Promise<PolicyRecord> {
  const raw = await session.execute({
    profile,
    operation: readOperation,
    params: { "conditionalAccessPolicy-id": policyId },
    query: { $select: "id,displayName,state,conditions,grantControls,sessionControls" },
    ...(profile.mode === "application" ? {} : { scopes: [...CA_POLICY_UPDATE_READ_SCOPES] }),
  });
  if (raw !== null && typeof raw === "object" && !Array.isArray(raw)
    && typeof (raw as PolicyRecord)["id"] === "string" && ((raw as PolicyRecord)["id"] as string).length > 0) {
    return raw as PolicyRecord;
  }
  throw new AxiError("Graph returned a policy without an object ID", "GRAPH_ERROR", [
    "Read back the policy before attempting an update",
  ]);
}

function stringArray(value: unknown): string[] | null {
  return Array.isArray(value) && value.every(entry => typeof entry === "string") ? [...value] : null;
}

// Lockout analysis over the proposed effective policy (current merged with
// the PATCH payload). An enabled policy that covers all users with no
// exclusions and block controls would lock out every admin including
// break-glass accounts, so enforcement-touching changes toward that state
// are refused. An enabled all-users policy without exclusions under milder
// controls is elevated risk and needs explicit acknowledgement. When the
// current conditions, user scope or grant controls are unreadable the
// analysis cannot run, and enforcement-touching changes stay disabled
// instead of sending a blind PATCH.
export function analyzeLockout(current: PolicyRecord, payload: Partial<Record<CaPolicyWritableField, unknown>>): LockoutAssessment {
  const effective: PolicyRecord = { ...current, ...payload };
  const state = effective["state"];
  if (typeof state !== "string") {
    return { available: false, reason: "policy state is unreadable, so lockout risk cannot be assessed" };
  }
  if (state !== "enabled") {
    return { available: true, level: "none", findings: [`policy state '${state}' does not enforce access`] };
  }
  const conditions = effective["conditions"];
  if (conditions === null || typeof conditions !== "object" || Array.isArray(conditions)) {
    return { available: false, reason: "policy conditions are unreadable, so lockout risk cannot be assessed" };
  }
  const users = (conditions as PolicyRecord)["users"];
  if (users === null || typeof users !== "object" || Array.isArray(users)) {
    return { available: false, reason: "policy user scope is unreadable, so lockout risk cannot be assessed" };
  }
  const includeUsers = stringArray((users as PolicyRecord)["includeUsers"]);
  if (includeUsers === null) {
    return { available: false, reason: "policy user scope is unreadable, so lockout risk cannot be assessed" };
  }
  const excludeUsers = stringArray((users as PolicyRecord)["excludeUsers"]);
  const excludeGroups = stringArray((users as PolicyRecord)["excludeGroups"]);
  if (excludeUsers === null || excludeGroups === null) {
    return { available: false, reason: "policy exclusions are unreadable, so lockout risk cannot be assessed" };
  }
  for (const field of ["includeGroups", "includeRoles"] as const) {
    const entries = (users as PolicyRecord)[field];
    if (entries !== undefined && stringArray(entries) === null) {
      return { available: false, reason: "policy user scope is unreadable, so lockout risk cannot be assessed" };
    }
  }
  const roles = (users as PolicyRecord)["excludeRoles"];
  const excludeRoles = roles === undefined ? [] : stringArray(roles);
  if (excludeRoles === null) {
    return { available: false, reason: "policy exclusions are unreadable, so lockout risk cannot be assessed" };
  }
  const excluded = excludeUsers.length > 0 || excludeGroups.length > 0 || excludeRoles.length > 0;
  const grantControls = effective["grantControls"];
  if (grantControls === null || typeof grantControls !== "object" || Array.isArray(grantControls)) {
    return { available: false, reason: "policy grant controls are unreadable, so lockout risk cannot be assessed" };
  }
  const builtIn = (grantControls as PolicyRecord)["builtInControls"];
  const controls = builtIn === undefined ? [] : stringArray(builtIn);
  if (controls === null) {
    return { available: false, reason: "policy grant controls are unreadable, so lockout risk cannot be assessed" };
  }
  if (includeUsers.includes("All") && !excluded && controls.includes("block")) {
    return {
      available: true,
      level: "refused",
      findings: ["enabled policy would block all users with no exclusions: every admin including break-glass access would be locked out"],
    };
  }
  return {
    available: true,
    level: "elevated",
    findings: ["enabled policy may affect admins and break-glass access; protected-account coverage cannot be established from policy targeting alone"],
  };
}

function touchesEnforcement(payload: Partial<Record<CaPolicyWritableField, unknown>>): boolean {
  return ENFORCEMENT_FIELDS.some(field => Object.hasOwn(payload, field));
}

function executeHint(policyId: string, elevated: boolean, profileName: string): string {
  const shell = (value: string): string =>
    /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
  return `mg-axi entra conditional-access policy update --id ${shell(policyId)} <reviewed field flags>`
    + ` --execute --confirm ${shell(policyId)}${elevated ? " --acknowledge-lockout-risk" : ""}`
    + ` --profile ${shell(profileName)}`;
}

function showHint(policyId: string, profileName: string): string {
  const shell = (value: string): string =>
    /^[A-Za-z0-9_.,:/@=-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
  return `mg-axi entra conditional-access policy show --id ${shell(policyId)} --profile ${shell(profileName)}`;
}

function assertLockoutGates(
  payload: Partial<Record<CaPolicyWritableField, unknown>>,
  assessment: LockoutAssessment,
  acknowledged: boolean,
): void {
  if (!touchesEnforcement(payload)) return;
  if (!assessment.available) {
    throw new AxiError(
      `Refusing policy update: ${assessment.reason}; the operation stays disabled without lockout analysis`,
      "OPERATION_BLOCKED",
      ["Read back the policy and retry only when its conditions, user scope and grant controls are verifiable"],
    );
  }
  if (assessment.level === "refused") {
    throw new AxiError(
      "Refusing policy update: the proposed policy would block all users with no exclusions, locking out every admin including break-glass access",
      "OPERATION_BLOCKED",
      [
        "Narrow conditions.users.includeUsers, add excludeUsers/excludeGroups for emergency access, or replace the block control before retrying",
        "This refusal cannot be overridden with an acknowledgement flag",
      ],
    );
  }
  if (assessment.level === "elevated" && !acknowledged) {
    throw new AxiError(
      "Policy update needs --acknowledge-lockout-risk: protected-account coverage cannot be established from policy targeting alone",
      "LOCKOUT_ACK_REQUIRED",
      ["Re-run with --acknowledge-lockout-risk after confirming emergency-access coverage"],
    );
  }
}

export async function updateCaPolicy(args: {
  session: GraphSession;
  coordinator: MutationCoordinator;
  flags: CaPolicyUpdateFlags;
  profile: AnyProfile;
  profileName: string;
  readOperation: SessionOperation;
  help: string;
}): Promise<Record<string, unknown>> {
  const { session, coordinator, flags, profile, profileName, readOperation, help } = args;
  const policyId = String(flags.id ?? "").trim();
  if (!policyId) throw new AxiError("--id needs a Conditional Access policy object ID", "VALIDATION_ERROR", [help]);
  const version = String(flags["api-version"] ?? "v1.0");
  if (version !== "v1.0") {
    throw new AxiError(
      `--api-version ${version} cannot mutate: WRITE-04 binds PATCH /identity/conditionalAccess/policies/{id} to v1.0; beta writes stay blocked`,
      "VALIDATION_ERROR",
      [help],
    );
  }
  const payload = parsePolicyUpdate(flags, help);
  const acknowledged = flags["acknowledge-lockout-risk"] === true;
  const definition = caPolicyDefinition(policyId, payload);
  const execute = flags.execute === true;
  // Gates first: preview refuses read-only, unscoped and beta-bound profiles
  // before any read, credential or transport.
  coordinator.preview(definition);
  if (execute) {
    if (flags.confirm === undefined) {
      throw new AxiError(
        `blocked: disruptive PATCH needs --confirm '${definition.target}'`,
        "CONFIRM_REQUIRED",
        [`Re-run with --confirm '${definition.target}'`],
      );
    }
    if (String(flags.confirm) !== definition.target) {
      throw new AxiError(
        `blocked: --confirm '${flags.confirm}' does not match target '${definition.target}'`,
        "CONFIRM_MISMATCH",
        [`Re-run with --confirm '${definition.target}'`],
      );
    }
  }
  const current = await readPolicy(session, profile, readOperation, policyId);
  const changes = (Object.keys(payload) as CaPolicyWritableField[]).map(field => ({
    field,
    current: Object.hasOwn(current, field) ? current[field] as unknown : null,
    proposed: payload[field] as unknown,
  }));
  // A missing property reads as absent, never as null: only an explicitly
  // stored equal value counts as already-desired.
  const noop = (Object.keys(payload) as CaPolicyWritableField[]).every(field =>
    sameValue(Object.hasOwn(current, field) ? current[field] : undefined, payload[field]));
  const assessment = analyzeLockout(current, payload);
  const previewUrl = coordinator.preview(definition).url;
  const preview = {
    operation: definition.operation,
    target: definition.target,
    method: definition.method,
    url: previewUrl,
    effect: definition.effect,
    changes,
    noop,
    lockout: assessment,
    concurrency: CONCURRENCY_NOTE,
  };
  if (!execute) {
    const elevated = assessment.available && assessment.level === "elevated";
    return {
      preview,
      help: [
        ...(noop
          ? ["Already in the desired state; no update is needed"]
          : [executeHint(policyId, elevated, profileName)]),
        ...(!noop
          ? ["Policy updates are disruptive: the --execute run needs the typed --confirm value shown above"]
          : []),
        ...(elevated && !noop ? ["Unverified admin and break-glass coverage needs --acknowledge-lockout-risk on the --execute run"] : []),
        showHint(policyId, profileName),
      ],
    };
  }
  if (noop) return { noop: true, policy: { id: current["id"] } };
  assertLockoutGates(payload, assessment, acknowledged);
  const readState = async (): Promise<PolicyRecord> => {
    const fresh = await readPolicy(session, profile, readOperation, policyId);
    assertLockoutGates(payload, analyzeLockout(fresh, payload), acknowledged);
    return fresh;
  };
  const result = await coordinator.execute(definition, {
    execute: true,
    confirm: String(flags.confirm),
    ...(profile.mode === "application" ? {} : { scopes: [...CA_POLICY_UPDATE_SCOPES] }),
    readState,
    isNoop: fresh => (Object.keys(payload) as CaPolicyWritableField[]).every(field => sameValue((fresh as PolicyRecord)[field], payload[field])),
  });
  if (result.kind === "failed") {
    throw new AxiError(
      `Graph refused the policy update (status ${result.status})`,
      "GRAPH_ERROR",
      [
        ...(result.status === 403 ? [PERMISSION_GUIDANCE, ROLE_GUIDANCE] : []),
        `Audit ${result.auditId} recorded the refusal; read back '${policyId}' and never replay this intent`,
      ],
    );
  }
  if (result.kind === "unknown") {
    throw new AxiError(`Update of '${policyId}' may or may not have been applied (audit ${result.auditId}); never replay this intent`, "OUTCOME_UNKNOWN", [
      `Audit ${result.auditId} recorded the uncertain outcome; read back '${policyId}' with ${showHint(policyId, profileName)} before doing anything else`,
      "Never replay this intent",
    ]);
  }
  // Anything but success, failure or uncertainty is the verified no-op:
  // dry-run needs execute !== true, always set above.
  if (result.kind !== "success") {
    return { noop: true, policy: { id: current["id"] } };
  }
  // Success carries no proof: policy-update answers 204 with an empty body.
  // The coordinator outcome is the only record; verify with policy show.
  return { policy: { id: current["id"] }, auditId: result.auditId, help: [showHint(policyId, profileName)] };
}
