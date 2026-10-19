import { AxiError } from "axi-sdk-js";
import { LEAVES, operationFor } from "./catalogue.js";
import { listUsers } from "./entra-users.js";
import type { GraphSession } from "./graph-session.js";
import { Profiles, type AnyProfile } from "./profiles.js";
import { authStatus } from "./auth-status.js";
import { invocation } from "./invocation.js";

// PACK-01: explicit release health check over the reviewed read-only session.
// Doctor performs one documented bounded read per profile (entra user list
// with a one-row window) and reports configuration, connectivity,
// authentication and access failures. Credential acquisition stays silent:
// DelegatedAuth.credential and ApplicationAuth.credential never open a
// browser or device-code flow, so doctor never signs in interactively.
// Doctor never auto-installs, changes profiles or enables Graph writes;
// silent token refresh stays in memory.
// See README.md for usage.
export const DOCTOR_OPERATION = "GET:/users";
export const DOCTOR_VERSION = "v1.0";
export const DOCTOR_CHECK = `entra user list --limit 1 (${DOCTOR_VERSION}:${DOCTOR_OPERATION})`;

export function doctorTargets(store: Profiles, flag?: string): string[] {
  if (flag !== undefined) return [store.resolve(flag).name];
  try {
    return [store.resolve(undefined).name];
  } catch (error) {
    if (!(error instanceof AxiError) || !["PROFILE_REQUIRED", "PROFILE_AMBIGUOUS"].includes(error.code)) throw error;
    return store.list().map(item => item.name);
  }
}

export type DoctorResult = { output: Record<string, unknown>; failed: boolean };

export async function runDoctor(args: {
  store: Profiles;
  names: string[];
  session: GraphSession;
  selected?: { name: string; profile: AnyProfile };
}): Promise<DoctorResult> {
  const { command } = invocation();
  if (!args.names.length) {
    throw new AxiError("No profiles are configured", "PROFILE_REQUIRED", [`${command} profile create --help`, `${command} setup`]);
  }
  const leaf = LEAVES.find(item => item.path === "entra user list");
  const operation = leaf === undefined ? undefined : operationFor(leaf, DOCTOR_VERSION);
  if (!operation || operation.method !== "GET") {
    throw new AxiError(`Unknown catalogued Graph operation for ${DOCTOR_CHECK}`, "VALIDATION_ERROR", [`${command} entra user list --help`]);
  }
  const profiles: Record<string, unknown>[] = [];
  for (const name of args.names) {
    const rerun = `${command} doctor --profile ${name}`;
    try {
      args.store.verifyPermissions();
      const selected = args.selected ?? args.store.resolve(name);
      authStatus(name, selected.profile);
      // The same bounded read the named leaf serves: one fetched page, one
      // surfaced row, remainder buffered in a discarded cursor. Doctor keeps
      // no cursor and accepts none; recovery reruns doctor, not the read.
      const result = (await listUsers(args.session, { limit: "1" }, selected.profile, operation, `${command} entra user list --help`, name)) as {
        users: unknown[];
        complete: boolean;
        reason?: string;
      };
      if (!result.complete && result.reason !== "row limit reached; buffered remainder is preserved in the cursor") {
        throw new AxiError(`Doctor read incomplete: ${result.reason}`, "GRAPH_ERROR", [rerun]);
      }
      profiles.push({ name, mode: selected.profile.mode, check: DOCTOR_CHECK, status: "ok",
        detail: result.users.length ? "1 user row returned" : "0 users matched; the absence of results is the answer" });
    } catch (error) {
      const failure = error instanceof AxiError
        ? error
        : new AxiError("Unable to run the doctor check", "GRAPH_ERROR", [`${command} entra user list --help`]);
      profiles.push({ name, check: DOCTOR_CHECK, status: "failed", code: failure.code, error: failure.message });
    }
  }
  const ok = profiles.filter(row => row.status === "ok");
  const failed = profiles.filter(row => row.status !== "ok");
  return {
    output: {
      config: args.store.path,
      check: `${DOCTOR_CHECK} per profile; silent credential acquisition only, no interactive sign-in, no writes`,
      count: `${ok.length} of ${profiles.length} profiles ok`,
      profiles,
      complete: failed.length === 0,
      help: failed.length
        ? failed.map(row => `[${String(row.name)}] Check the reported failure, then rerun \`${command} doctor --profile ${String(row.name)}\``)
        : [`Run \`${command} entra user list --profile ${String(ok[0]!.name)} --limit 1\` to start reading`],
    },
    failed: failed.length > 0,
  };
}
