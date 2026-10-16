// One-line skill install hint for the no-args home view. Detects the agent
// host from its documented environment marker and checks whether the mg-axi
// skill is already installed for that host, at user, project (working
// directory upwards) and system level. File-existence checks only: this never
// installs, writes, fails or touches the network.
//
// The check looks only at the detected host's row, so one host's copy never
// hides another host's hint. Each row lists every documented skill folder
// that host loads, verified against the real hosts:
// - Claude Code reads its config directory ($CLAUDE_CONFIG_DIR or ~/.claude,
//   per its binary) plus ~/.agents/skills and .agents/skills (per its
//   binary); project .claude/skills loads from the working directory up to
//   the repository root (per https://code.claude.com/docs/en/skills).
// - Codex reads $HOME/.agents/skills and .agents/skills (cwd to repo root)
//   per https://learn.chatgpt.com/docs/build-skills.md, plus
//   $CODEX_HOME/skills (default ~/.codex/skills, per its binary and the
//   skills-CLI registry) and the documented /etc/codex/skills admin folder.
// - Pi reads its user/project skills directories plus the Agent Skills
//   locations ~/.agents/skills/ and .agents/skills/ (ancestors to the repo
//   root) per its skills.md.
// - OpenCode reads its config directory ($OPENCODE_CONFIG_DIR,
//   $XDG_CONFIG_HOME or ~/.config, per its binary and CLI docs) plus
//   ~/.claude/skills, ~/.agents/skills, .opencode/skills, .claude/skills and
//   .agents/skills (project paths walk up to the git worktree) per
//   https://opencode.ai/docs/skills/.
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";

export const SKILL_NAME = "mg-axi";
export const SKILL_INSTALL_COMMAND = `npx skills add knowttl/${SKILL_NAME} --skill ${SKILL_NAME}`;

export interface SkillHostRow {
  name: "claude-code" | "codex" | "pi" | "opencode";
  userDirs: string[];
  projectDirs: string[];
  systemDirs: string[];
}

function resolveHome(env: NodeJS.ProcessEnv): string {
  return env.HOME ?? env.USERPROFILE ?? homedir();
}

export function skillHostRows(env: NodeJS.ProcessEnv, home: string): SkillHostRow[] {
  const agentsUser = join(home, ".agents", "skills");
  const agentsProject = join(".agents", "skills");
  return [
    {
      name: "claude-code",
      userDirs: [join(env.CLAUDE_CONFIG_DIR ?? join(home, ".claude"), "skills"), agentsUser],
      projectDirs: [join(".claude", "skills"), agentsProject],
      systemDirs: [],
    },
    {
      name: "codex",
      userDirs: [join(env.CODEX_HOME ?? join(home, ".codex"), "skills"), agentsUser],
      projectDirs: [agentsProject],
      systemDirs: ["/etc/codex/skills"],
    },
    {
      name: "pi",
      userDirs: [join(home, ".pi", "agent", "skills"), agentsUser],
      projectDirs: [join(".pi", "skills"), agentsProject],
      systemDirs: [],
    },
    {
      name: "opencode",
      userDirs: [
        join(env.OPENCODE_CONFIG_DIR ?? env.XDG_CONFIG_HOME ?? join(home, ".config"), "opencode", "skills"),
        join(home, ".claude", "skills"),
        agentsUser,
      ],
      projectDirs: [join(".opencode", "skills"), join(".claude", "skills"), agentsProject],
      systemDirs: [],
    },
  ];
}

// Marker variables each host sets for its own child processes, verified
// against the real hosts: CLAUDECODE (Claude Code docs), CODEX_THREAD_ID /
// CODEX_SESSION_ID (codex shell_environment.rs), PI_CODING_AGENT (pi docs),
// OPENCODE (opencode startup middleware, inherited by its shell children).
export function detectSkillHost(env: NodeJS.ProcessEnv, home: string): SkillHostRow | undefined {
  const rows = skillHostRows(env, home);
  if (env.CLAUDECODE !== undefined) return rows.find(row => row.name === "claude-code");
  if (env.CODEX_THREAD_ID !== undefined || env.CODEX_SESSION_ID !== undefined) return rows.find(row => row.name === "codex");
  if (env.PI_CODING_AGENT !== undefined) return rows.find(row => row.name === "pi");
  if (env.OPENCODE !== undefined) return rows.find(row => row.name === "opencode");
  return undefined;
}

export function skillInstalled(row: SkillHostRow, fromDir: string): boolean {
  if ([...row.userDirs, ...row.systemDirs].some(dir => existsSync(join(dir, SKILL_NAME)))) return true;
  for (let dir = resolve(fromDir); ; dir = dirname(dir)) {
    if (row.projectDirs.some(projectDir => existsSync(join(dir, projectDir, SKILL_NAME)))) return true;
    if (dirname(dir) === dir) return false;
  }
}

export function skillHint(env: NodeJS.ProcessEnv = process.env, fromDir: string = process.cwd()): string | undefined {
  const host = detectSkillHost(env, resolveHome(env));
  if (host === undefined || skillInstalled(host, fromDir)) return undefined;
  return SKILL_INSTALL_COMMAND;
}
