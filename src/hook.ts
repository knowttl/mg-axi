import { encode } from "@toon-format/toon";
import { Profiles, type AnyProfile } from "./profiles.js";
import { Credentials } from "./credentials.js";
import { PrivateFiles } from "./private-file.js";
import { resolveSecret } from "./secrets.js";
import { redactSecrets } from "./redact.js";

// Session-start ambient summary behind the mg-axi-hook entry point.
// Local configuration state only: every configured profile name with its
// tenant label, auth-cache presence and effective write posture, plus the
// package version and one next-step hint. No session, transport, Graph or
// sign-in call of any kind; every failure collapses to a short record with
// exit code 0, so a missing or broken config never breaks agent startup.
//
// Cache presence uses the same binding and private-file checks as authentication.

export type AuthPresence = "cached" | "sign-in required" | "client-credentials" | "unknown";
export type HookOptions = {
  version: string;
  env?: NodeJS.ProcessEnv;
  configPath?: string;
  files?: PrivateFiles;
  // Test seam over the true external boundary (the credential store).
  authPresence?: (profile: AnyProfile) => Promise<AuthPresence> | AuthPresence;
};

async function probeAuthPresence(profile: AnyProfile, files = new PrivateFiles()): Promise<AuthPresence> {
  if (profile.mode !== "delegated") return "client-credentials";
  try {
    return ("stored" in profile.token ? new Credentials(undefined, files).read(profile) : resolveSecret(profile.token, files)) ? "cached" : "sign-in required";
  } catch {
    return "unknown";
  }
}

// Next-step hints name real catalogue leaves; test/hook.test.ts asserts
// every `mg-axi ...` command mentioned here exists in the catalogue, so the
// hook text cannot drift from the single source.
const CONFIGURED_HINT = "Run npx -y @knowttl/mg-axi doctor --profile <name> to check a profile";
const SETUP_HINT = "Run npx -y @knowttl/mg-axi setup to find the config path and example";

export async function hookSummary(options: HookOptions): Promise<string> {
  try {
    return redactSecrets(await renderHookSummary(options));
  } catch {
    return "mg: status unavailable\n";
  }
}

async function renderHookSummary({ version, env = process.env, configPath, files = new PrivateFiles(), authPresence = profile => probeAuthPresence(profile, files) }: HookOptions): Promise<string> {
  let store: Profiles;
  let names: string[];
  try {
    store = new Profiles(configPath !== undefined && configPath.trim() !== "" ? configPath.trim() : env.MG_AXI_CONFIG !== undefined && env.MG_AXI_CONFIG.trim() !== "" ? env.MG_AXI_CONFIG.trim() : undefined, files);
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
