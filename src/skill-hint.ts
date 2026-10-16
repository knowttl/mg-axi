// One-line skill install hint for the no-args home view. Detects the agent
// host from its documented environment marker and checks whether the mg-axi
// skill is already installed for that host, at user level and at project
// level from the working directory upwards. File-existence checks only:
// this never installs, writes, fails or touches the network.
//
// "Installed" counts both the skills-CLI install paths for the host and the
// host's own skill folders, verified against the real hosts:
// - Claude Code reads ~/.claude/skills and .claude/skills (cwd to repo root)
//   per https://code.claude.com/docs/en/skills, honouring CLAUDE_CONFIG_DIR
//   for the personal folder; its binary also reads ~/.agents/skills and
//   .agents/skills.
// - Codex reads $CODEX_HOME/skills (default ~/.codex/skills) per its binary.
// - Pi reads its user/project skills directories plus the Agent Skills
//   locations ~/.agents/skills/ and .agents/skills/ (ancestors to the repo
//   root) per its skills.md.
// - OpenCode reads ~/.config/opencode/skills, .opencode/skills,
//   ~/.claude/skills, .claude/skills, ~/.agents/skills and .agents/skills
//   per https://opencode.ai/docs/skills/ (project paths walk up to the git
//   worktree).
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";

export const SKILL_NAME = "mg-axi";
export const SKILL_INSTALL_COMMAND = `npx skills add knowttl/${SKILL_NAME} --skill ${SKILL_NAME}`;

interface SkillHost {
  userDirs: string[];
  projectDirs: string[];
}

function resolveHome(env: NodeJS.ProcessEnv): string {
  return env.HOME ?? env.USERPROFILE ?? homedir();
}

// Marker variables each host sets for its own child processes, verified
// against the real hosts: CLAUDECODE (Claude Code docs), CODEX_THREAD_ID /
// CODEX_SESSION_ID (codex shell_environment.rs), PI_CODING_AGENT (pi docs),
// OPENCODE (opencode startup middleware, inherited by its shell children).
function detectSkillHost(env: NodeJS.ProcessEnv, home: string): SkillHost | undefined {
  const agentsUser = join(home, ".agents", "skills");
  const agentsProject = join(".agents", "skills");
  if (env.CLAUDECODE !== undefined)
    return { userDirs: [join(env.CLAUDE_CONFIG_DIR ?? join(home, ".claude"), "skills"), agentsUser], projectDirs: [join(".claude", "skills"), agentsProject] };
  if (env.CODEX_THREAD_ID !== undefined || env.CODEX_SESSION_ID !== undefined)
    return { userDirs: [join(env.CODEX_HOME ?? join(home, ".codex"), "skills"), agentsUser], projectDirs: [agentsProject] };
  if (env.PI_CODING_AGENT !== undefined)
    return { userDirs: [join(home, ".pi", "agent", "skills"), agentsUser], projectDirs: [join(".pi", "skills"), agentsProject] };
  if (env.OPENCODE !== undefined)
    return {
      userDirs: [join(env.XDG_CONFIG_HOME ?? join(home, ".config"), "opencode", "skills"), join(home, ".claude", "skills"), agentsUser],
      projectDirs: [join(".opencode", "skills"), join(".claude", "skills"), agentsProject],
    };
  return undefined;
}

function skillInstalled(host: SkillHost, fromDir: string): boolean {
  if (host.userDirs.some(dir => existsSync(join(dir, SKILL_NAME)))) return true;
  for (let dir = resolve(fromDir); ; dir = dirname(dir)) {
    if (host.projectDirs.some(projectDir => existsSync(join(dir, projectDir, SKILL_NAME)))) return true;
    if (dirname(dir) === dir) return false;
  }
}

export function skillHint(env: NodeJS.ProcessEnv = process.env, fromDir: string = process.cwd()): string | undefined {
  const host = detectSkillHost(env, resolveHome(env));
  if (host === undefined || skillInstalled(host, fromDir)) return undefined;
  return SKILL_INSTALL_COMMAND;
}
