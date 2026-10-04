import { AxiError } from "axi-sdk-js";
import { ApplicationAuth } from "./app-auth.js";
import { DelegatedAuth } from "./auth.js";
import { GraphSession, resolveSessionOperation } from "./graph-session.js";
import { createMutationCoordinator, mutationGraphFault, mutationHttpStatus, type MutationTransport } from "./mutations.js";
import type { AnyProfile } from "./profiles.js";

// WRITE-01: the first named mutation, `mg-axi entra group member add`, behind
// the WRITE-00 coordinator. One user joins one group through
// POST /groups/{group-id}/members/$ref with a directoryObjects $ref target.
//
// Reviewed against the v1.0 group-post-members operation documentation on
// 2026-10-04. Success is 204 No Content with an empty body. Graph returns
// 400 when the object is already a member, when the member type is
// unsupported, or when a recently created group has not finished replicating
// ("the source resource object or one of the objects being referenced don't
// exist" - retry later as a new command, never by replaying the intent).
// It returns 404 when the referenced object does not exist and 403 for
// unmanageable group types, missing permission, or role-assignable groups
// without role-management grants. Only the already-member 400 is a no-op;
// every other 400 stays a failure.
//
// Supported scope is deliberately narrow: non-role-assignable security and
// Microsoft 365 groups, user members only. User members take D/A
// GroupMember.ReadWrite.All; every other member type needs additional
// permissions and belongs to WRITE-N, never to this slice. Delegated callers
// additionally need an Entra role carrying
// microsoft.directory/groups/members/update (least-privileged choices include
// group owner, Directory Writers, Groups Administrator, Identity Governance
// Administrator and User Administrator). Role-assignable groups need
// RoleManagement.ReadWrite.Directory plus Privileged Role Administrator and
// are refused before sending, as are dynamic-membership and distribution
// groups, which Graph cannot add members to.

// Profile write-scope operation name. Hand-edit this exact string into the
// profile's writes.operations allowlist; no command writes that object.
export const GROUP_MEMBER_ADD_OPERATION = "mg.entra.group.member.add";
// The only delegated scope this slice requests for the mutation itself.
export const GROUP_MEMBER_ADD_SCOPES = ["https://graph.microsoft.com/GroupMember.ReadWrite.All"];
// Preview reads travel the READ-02 route with its standard read scope.
const PREVIEW_SCOPES = ["https://graph.microsoft.com/GroupMember.Read.All"];
const GROUP_SELECT = [
  "id",
  "displayName",
  "isAssignableToRole",
  "securityEnabled",
  "groupTypes",
  "mailEnabled",
  "membershipRule",
  "membershipRuleProcessingState",
];
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export type MemberAddFlags = Record<string, string | boolean>;

function fail(message: string, code: string, suggestions: string[]): never {
  throw new AxiError(message, code, suggestions);
}

function groupRef(groupId: string): string {
  return `https://graph.microsoft.com/v1.0/groups/${groupId}`;
}

function memberRef(userId: string): string {
  return `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`;
}

function executeHint(groupId: string, userId: string, profileName: string): string {
  return `mg-axi entra group member add --group ${groupId} --user ${userId} --profile ${profileName} --execute --confirm ${groupId}`;
}

function readBackHint(groupId: string, profileName: string): string {
  return `mg-axi entra group member list --group ${groupId} --profile ${profileName}`;
}

// Pre-send target refusal. Runs inside the preview read so the group is
// rechecked on the fresh read before the single send. An initial refusal
// journals nothing; a refusal after intent reservation records NOT_SENT.
function assertGroupWritable(group: unknown, groupId: string): { id: string; displayName: unknown } {
  if (group === null || typeof group !== "object" || Array.isArray(group)) {
    fail("Graph returned a malformed group body", "GRAPH_ERROR", [
      "Single-group reads carry one group object; treat anything else as unknown, not empty",
    ]);
  }
  const row = group as Record<string, unknown>;
  if (row["isAssignableToRole"] === true) {
    fail(`Refusing role-assignable group ${groupId}: WRITE-01 adds users to non-role-assignable groups only`, "OPERATION_BLOCKED", [
      "Role-assignable membership needs RoleManagement.ReadWrite.Directory and Privileged Role Administrator; that grant belongs to a later slice, never to this command",
      "Verify the group with mg-axi entra group show --id <group-id>",
    ]);
  }
  if (row["isAssignableToRole"] !== false && row["isAssignableToRole"] !== null) {
    fail(`Refusing group ${groupId}: its role-assignable state is unknown`, "OPERATION_BLOCKED", [
      "Read back the group with mg-axi entra group show --id <group-id> before doing anything else",
    ]);
  }
  if (typeof row["membershipRule"] === "string" && row["membershipRule"].trim() !== "") {
    fail(`Refusing dynamic-membership group ${groupId}: members come from its rule, not manual adds`, "OPERATION_BLOCKED", [
      "Manage membership through the group's membershipRule instead",
    ]);
  }
  const groupTypes = Array.isArray(row["groupTypes"]) ? row["groupTypes"] : [];
  if (row["securityEnabled"] !== true && !groupTypes.includes("Unified")) {
    fail(`Refusing group ${groupId}: Graph manages members of security and Microsoft 365 groups only`, "OPERATION_BLOCKED", [
      "Verify the group with mg-axi entra group show --id <group-id>",
    ]);
  }
  return { id: groupId, displayName: row["displayName"] };
}

type MembershipState = {
  group: { id: string; displayName: unknown };
  membersChecked: number;
  membersComplete: boolean;
  alreadyMember: boolean;
};

function isDuplicateReference(error: unknown): boolean {
  if (mutationHttpStatus(error) !== 400) return false;
  const fault = mutationGraphFault(error);
  return fault !== null && /already exist/i.test(`${fault.code} ${fault.message}`);
}

function failedResult(status: number, groupId: string, userId: string, profileName: string): never {
  const readBack = readBackHint(groupId, profileName);
  if (status === 400) {
    fail(`Graph refused the membership add (400) without applying it`, "GRAPH_ERROR", [
      "Verify the user object ID exists: an unknown user reports 404, while an unsupported member type or an unreplicated recently created group reports 400",
      "If the group was created moments ago, wait for replication and retry as a new command with a new intent; never replay a reserved intent",
      readBack,
    ]);
  }
  if (status === 403) {
    fail(`Graph denied the membership add (403): the grant, role, licence or policy prerequisite is missing and Graph does not say which`, "GRAPH_ERROR", [
      "User members need D/A GroupMember.ReadWrite.All; other member types need additional permissions and belong to a later slice",
      "Delegated callers additionally need an Entra role with microsoft.directory/groups/members/update (for example Groups Administrator, User Administrator, Directory Writers, Identity Governance Administrator, or group owner)",
      "A 403 never proves which prerequisite is missing; do not retry blindly",
      readBack,
    ]);
  }
  if (status === 404) {
    fail(`Graph reports the group or user as not found or inaccessible (404); Graph does not distinguish missing from hidden`, "GRAPH_ERROR", [
      "Verify both object IDs; --user takes the user object ID, never a UPN",
      readBack,
    ]);
  }
  fail(`Graph membership add returned status ${status}`, "GRAPH_ERROR", [
    `Read back the membership with ${readBack} before doing anything else; never replay this intent`,
  ]);
}

export async function addGroupMember(args: {
  session: GraphSession;
  profile: AnyProfile;
  delegated: DelegatedAuth;
  application: ApplicationAuth;
  transport: MutationTransport;
  flags: MemberAddFlags;
  profileName: string;
  help: string;
  journalPath?: string;
  clock?: () => number;
}): Promise<Record<string, unknown>> {
  const { session, profile, delegated, application, transport, flags, profileName, help } = args;
  // Strict flag validation runs before any credential or HTTP.
  const apiVersion = String(flags["api-version"] ?? "v1.0");
  if (apiVersion !== "v1.0") {
    fail(`No reviewed ${apiVersion} membership write: WRITE-01 binds to v1.0 only`, "VALIDATION_ERROR", [help]);
  }
  const groupId = String(flags.group ?? "").trim();
  if (!groupId) fail("--group needs the group object ID", "VALIDATION_ERROR", [help]);
  const userId = String(flags.user ?? "").trim();
  if (!userId) fail("--user needs the user object ID", "VALIDATION_ERROR", [help]);
  if (!GUID.test(groupId)) {
    fail(`--group ${groupId} is not a group object ID`, "VALIDATION_ERROR", [help, "Pass the group object ID (GUID); names and mail addresses are not accepted"]);
  }
  if (!GUID.test(userId)) {
    fail(`--user ${userId} is not a user object ID`, "VALIDATION_ERROR", [help, "Pass the user object ID (GUID); UPNs are not resolved by this command"]);
  }
  if (profile.mode === "application" && flags.scopes !== undefined) {
    fail("Application profiles use the configured Graph .default audience; delegated scopes are unavailable", "VALIDATION_ERROR", [
      help,
      "mg-axi profile show --profile <name>",
    ]);
  }
  let writeScopes: string[] | undefined;
  if (profile.mode === "delegated") {
    if (flags.scopes === undefined) writeScopes = [...GROUP_MEMBER_ADD_SCOPES];
    else {
      writeScopes = String(flags.scopes).split(",").map(scope => scope.trim()).filter(scope => scope.length > 0);
      if (!writeScopes.length) fail("--scopes needs at least one Graph scope", "VALIDATION_ERROR", [help]);
    }
  }
  const readScopes = profile.mode === "delegated" ? [...PREVIEW_SCOPES] : undefined;
  const userReadScopes = profile.mode === "delegated" ? ["https://graph.microsoft.com/User.ReadBasic.All"] : undefined;
  const groupOperation = resolveSessionOperation("v1.0", "GET", "/groups/{group-id}");
  const userOperation = resolveSessionOperation("v1.0", "GET", "/users/{user-id}");
  const membersOperation = resolveSessionOperation("v1.0", "GET", "/groups/{group-id}/members");

  // Desired-state read through the READ-02 route. An incomplete page window
  // cannot prove absence, so a missed member proceeds to the POST and lands
  // as a duplicate no-op instead of a failure.
  const readState = async (): Promise<MembershipState> => {
    const group = await session.execute({
      profile,
      operation: groupOperation,
      params: { "group-id": groupId },
      query: { $select: GROUP_SELECT.join(",") },
      scopes: readScopes,
    });
    const writable = assertGroupWritable(group, groupId);
    const user = await session.execute({
      profile,
      operation: userOperation,
      params: { "user-id": userId },
      query: { $select: "id" },
      scopes: userReadScopes,
    });
    const userRow = user as Record<string, unknown> | null;
    if (userRow === null || typeof userRow !== "object" || Array.isArray(userRow)
      || typeof userRow["id"] !== "string" || userRow["id"].toLowerCase() !== userId.toLowerCase()
      || (userRow["@odata.type"] !== undefined && userRow["@odata.type"] !== "#microsoft.graph.user")) {
      fail(`Graph did not establish user identity for ${userId}`, "GRAPH_ERROR", [
        "Verify --user is an accessible user object ID; non-user objects cannot be added by this command",
      ]);
    }
    const collected = await session.collect({
      profile,
      operation: membersOperation,
      params: { "group-id": groupId },
      query: { $select: "id" },
      scopes: readScopes,
    });
    const wanted = userId.toLowerCase();
    let alreadyMember = false;
    let checked = 0;
    for (const row of collected.value) {
      if (row === null || typeof row !== "object" || Array.isArray(row)) continue;
      const id = (row as Record<string, unknown>)["id"];
      checked += 1;
      if (typeof id === "string" && id.toLowerCase() === wanted) alreadyMember = true;
    }
    return { group: writable, membersChecked: checked, membersComplete: collected.complete, alreadyMember };
  };

  const coordinator = createMutationCoordinator({
    profile,
    delegated,
    application,
    transport,
    ...(args.clock === undefined ? {} : { clock: args.clock }),
    ...(args.journalPath === undefined ? {} : { journalPath: args.journalPath }),
  });
  const definition = {
    operation: GROUP_MEMBER_ADD_OPERATION,
    method: "POST" as const,
    version: "v1.0",
    path: `/v1.0/groups/${groupId}/members/$ref`,
    effect: "disruptive" as const,
    target: groupId,
    payload: { "@odata.id": memberRef(userId) },
  };
  const result = await coordinator.execute(definition, {
    ...(flags.execute === true ? { execute: true } : {}),
    ...(flags.confirm === undefined ? {} : { confirm: String(flags.confirm) }),
    ...(writeScopes === undefined ? {} : { scopes: writeScopes }),
    readState,
    isNoop: current => (current as MembershipState).alreadyMember === true,
    isDuplicate: isDuplicateReference,
  });

  const membership = { group: groupRef(groupId), user: memberRef(userId) };
  if (result.kind === "dry-run") {
    return {
      membership: { ...membership, status: "preview" },
      preview: result.preview,
      help: [executeHint(groupId, userId, profileName), readBackHint(groupId, profileName)],
    };
  }
  if (result.kind === "noop") {
    if (result.duplicate === true) {
      return {
        membership: { ...membership, status: "already-member" },
        noop: true,
        auditId: result.auditId,
        note: "Graph reported the membership already exists; no change was made",
      };
    }
    return { membership: { ...membership, status: "already-member" }, noop: true };
  }
  if (result.kind === "success") {
    return {
      membership: { ...membership, status: "added" },
      auditId: result.auditId,
      help: [readBackHint(groupId, profileName)],
    };
  }
  if (result.kind === "failed") {
    failedResult(result.status, groupId, userId, profileName);
  }
  return {
    membership: { ...membership, status: "unknown" },
    auditId: result.auditId,
    httpStatus: result.httpStatus,
    guidance: result.guidance,
    help: [readBackHint(groupId, profileName)],
  };
}
