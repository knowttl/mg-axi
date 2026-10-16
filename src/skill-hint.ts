// One-line skill install hint for the no-args home view. Detects the agent
// host from its documented environment marker and checks whether the mg-axi
// skill is already installed for that host, at user level and at project
// level from the working directory upwards. File-existence checks only:
// this never installs, writes, fails or touches the network.
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";

export const SKILL_NAME = "mg-axi";
export const SKILL_INSTALL_COMMAND = `npx skills add knowttl/${SKILL_NAME} --skill ${SKILL_NAME}`;

interface SkillHost {
  projectDir: string;
  globalDir: string;
}

function resolveHome(env: NodeJS.ProcessEnv): string {
  return env.HOME ?? env.USERPROFILE ?? homedir();
}

// Marker variables each host sets for its own child processes, verified
// against the real hosts: CLAUDECODE (Claude Code docs), CODEX_THREAD_ID /
// CODEX_SESSION_ID (codex shell_environment.rs), PI_CODING_AGENT (pi docs),
// OPENCODE (opencode startup middleware, inherited by its shell children).
// Install locations mirror the skills CLI registry for that host.
function detectSkillHost(env: NodeJS.ProcessEnv, home: string): SkillHost | undefined {
  if (env.CLAUDECODE !== undefined)
    return { projectDir: join(".claude", "skills"), globalDir: join(env.CLAUDE_CONFIG_DIR ?? join(home, ".claude"), "skills") };
  if (env.CODEX_THREAD_ID !== undefined || env.CODEX_SESSION_ID !== undefined)
    return { projectDir: join(".agents", "skills"), globalDir: join(env.CODEX_HOME ?? join(home, ".codex"), "skills") };
  if (env.PI_CODING_AGENT !== undefined)
    return { projectDir: join(".pi", "skills"), globalDir: join(home, ".pi", "agent", "skills") };
  if (env.OPENCODE !== undefined)
    return { projectDir: join(".agents", "skills"), globalDir: join(env.XDG_CONFIG_HOME ?? join(home, ".config"), "opencode", "skills") };
  return undefined;
}

function skillInstalled(host: SkillHost, fromDir: string): boolean {
  if (existsSync(join(host.globalDir, SKILL_NAME))) return true;
  for (let dir = resolve(fromDir); ; dir = dirname(dir)) {
    if (existsSync(join(dir, host.projectDir, SKILL_NAME))) return true;
    if (dirname(dir) === dir) return false;
  }
}

export function skillHint(env: NodeJS.ProcessEnv = process.env, fromDir: string = process.cwd()): string | undefined {
  const host = detectSkillHost(env, resolveHome(env));
  if (host === undefined || skillInstalled(host, fromDir)) return undefined;
  return SKILL_INSTALL_COMMAND;
}
