import { AxiError } from "axi-sdk-js";
import { LEAVES, operationFor } from "./catalogue.js";
import { listUsers } from "./entra-users.js";
import type { GraphSession } from "./graph-session.js";
import { Profiles } from "./profiles.js";

// PACK-01: explicit release health check over the reviewed read-only session.
// Doctor performs one documented bounded read per profile (entra user list
// with a one-row window) and reports configuration, connectivity,
// authentication and access failures. Credential acquisition stays silent:
// DelegatedAuth.credential and ApplicationAuth.credential never open a
// browser or device-code flow, so doctor never signs in interactively.
// Doctor writes nothing: no auto-install, no profile changes, no write
// enablement. See README.md for usage.
export const DOCTOR_OPERATION = "GET:/users";
export const DOCTOR_VERSION = "v1.0";
export const DOCTOR_CHECK = `entra user list --limit 1 (${DOCTOR_VERSION}:${DOCTOR_OPERATION})`;

// Explicit --profile wins; otherwise the configured default, then the sole
// profile. With no selection among several profiles, doctor checks every
// profile instead of failing ambiguous; with none, the caller reports the
// missing configuration before any credential or HTTP work.
export function doctorTargets(store: Profiles, flag?: string): string[] {
  if (flag !== undefined) return [store.resolve(flag).name];
  const items = store.list();
  if (!items.length) return [];
  try {
    return [store.resolve(undefined).name];
  } catch {
    return items.map(item => item.name);
  }
}

export type DoctorResult = { output: Record<string, unknown>; failed: boolean };

export async function runDoctor(args: {
  store: Profiles;
  names: string[];
  session: GraphSession;
}): Promise<DoctorResult> {
  if (!args.names.length) {
    throw new AxiError("No profiles are configured", "AUTH_REQUIRED", ["mg-axi profile create --help", "mg-axi setup"]);
  }
  const leaf = LEAVES.find(item => item.path === "entra user list");
  const operation = leaf === undefined ? undefined : operationFor(leaf, DOCTOR_VERSION);
  if (!operation || operation.method !== "GET") {
    throw new AxiError(`Unknown catalogued Graph operation for ${DOCTOR_CHECK}`, "VALIDATION_ERROR", ["mg-axi entra user list --help"]);
  }
  const profiles: Record<string, unknown>[] = [];
  for (const name of args.names) {
    const rerun = `mg-axi doctor --profile ${name}`;
    try {
      const selected = args.store.resolve(name);
      // The same bounded read the named leaf serves: one fetched page, one
      // surfaced row, remainder buffered in a discarded cursor. Doctor keeps
      // no cursor and accepts none; recovery reruns doctor, not the read.
      const result = (await listUsers(args.session, { limit: "1" }, selected.profile, operation, `mg-axi ${DOCTOR_CHECK} --help`, name)) as {
        users: unknown[];
      };
      profiles.push({ name, mode: selected.profile.mode, check: DOCTOR_CHECK, status: "ok",
        detail: result.users.length ? "1 user row returned" : "0 users matched; the absence of results is the answer" });
    } catch (error) {
      const failure = error instanceof AxiError
        ? error
        : new AxiError("Unable to run the doctor check", "GRAPH_ERROR", ["mg-axi entra user list --help"]);
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
        ? failed.map(row => `[${String(row.name)}] Check the reported failure, then rerun \`mg-axi doctor --profile ${String(row.name)}\``)
        : [`Run \`mg-axi entra user list --profile ${String(ok[0]!.name)} --limit 1\` to start reading`],
    },
    failed: failed.length > 0,
  };
}
