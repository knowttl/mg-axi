import { encode } from "@toon-format/toon";
import { Profiles, type AnyProfile } from "./profiles.js";

// Session-start ambient summary behind the mg-axi-hook entry point.
// Local configuration state only: every configured profile name with its
// tenant label, auth-cache presence and effective write posture, plus the
// package version and one next-step hint. No session, transport, Graph or
// sign-in call of any kind; every failure collapses to a short record with
// exit code 0, so a missing or broken config never breaks agent startup.
//
// Auth-cache presence reads the OS credential store directly: delegated
// profiles keep their MSAL session blob under the "mg-axi" service with the
// session key below, so presence means a previous explicit login happened on
// this machine. Only presence is reported; tokens and key material never
// reach the output. Application profiles acquire tokens silently per run, so
// they report the client-credentials mechanism instead of cache state.

// Session cache key mirrors MsalProvider.application() in
// src/msal-provider.ts; change both together.
function delegatedSessionKey(profile: AnyProfile): string {
  return JSON.stringify([profile.credentialRef.key, profile.tenantId, profile.clientId, profile.cloud]);
}

export type AuthPresence = "cached" | "sign-in required" | "client-credentials" | "unknown";
export type HookOptions = {
  version: string;
  env?: NodeJS.ProcessEnv;
  configPath?: string;
  // Test seam over the true external boundary (the OS credential store).
  authPresence?: (profile: AnyProfile) => Promise<AuthPresence> | AuthPresence;
};

async function probeAuthPresence(profile: AnyProfile): Promise<AuthPresence> {
  if (profile.mode !== "delegated") return "client-credentials";
  let store: { getPassword(service: string, account: string): Promise<string | null> } | undefined;
  try {
    store = await import("keytar").then(module => module.default).catch(() => undefined);
  } catch {
    return "unknown";
  }
  if (!store) return "unknown";
  try {
    return typeof await store.getPassword("mg-axi", delegatedSessionKey(profile)) === "string" ? "cached" : "sign-in required";
  } catch {
    return "unknown";
  }
}

// Next-step hints name real catalogue leaves; test/hook.test.mjs asserts
// every `mg-axi ...` command mentioned here exists in the catalogue, so the
// hook text cannot drift from the single source.
const CONFIGURED_HINT = "Run npx -y @knowttl/mg-axi doctor --profile <name> to check a profile";
const SETUP_HINT = "Run npx -y @knowttl/mg-axi setup to find the config path and example";

export async function hookSummary(options: HookOptions): Promise<string> {
  try {
    return await renderHookSummary(options);
  } catch {
    return "mg: status unavailable\n";
  }
}

async function renderHookSummary({ version, env = process.env, configPath, authPresence = probeAuthPresence }: HookOptions): Promise<string> {
  let store: Profiles;
  let names: string[];
  try {
    store = new Profiles(configPath ?? env.MG_AXI_CONFIG);
    names = store.list().map(item => item.name);
  } catch {
    return `${encode({ mg: "configuration invalid", help: ["Check the selected config file's path, permissions and JSON syntax"] })}\n`;
  }
  if (names.length === 0) {
    return `${encode({ mg: "not configured", help: [SETUP_HINT] })}\n`;
  }
  const profiles = [];
  for (const name of names) {
    const profile = store.resolve(name).profile;
    let auth: AuthPresence = "unknown";
    try {
      auth = await authPresence(profile);
    } catch {
      auth = "unknown";
    }
    profiles.push({
      name,
      tenant: profile.tenantId,
      auth,
      writes: profile.writes?.allowWrites === true ? profile.writes.operations.join(",") : "disabled",
    });
  }
  return `${encode({ mg: "configured", version, profiles, help: [CONFIGURED_HINT] })}\n`;
}
