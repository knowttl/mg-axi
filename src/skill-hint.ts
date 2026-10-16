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
//   the repository root (per https://code.claude.com/docs/en/skills). When
//   the working directory is inside a linked git worktree whose root has no
//   .claude/skills directory, Claude Code (v2.1.277+) falls back to the main
//   checkout's project skills (per https://code.claude.com/docs/en/worktrees).
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
//
// Only a direct-child <folder>/mg-axi/SKILL.md counts as installed. Folders
// listed in host settings files (for example Pi settings.json) and recursive
// nested layouts are a documented limitation and are deliberately not
// implemented.
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";

export const SKILL_NAME = "mg-axi";
export const SKILL_INSTALL_COMMAND = `npx skills add knowttl/${SKILL_NAME} --skill ${SKILL_NAME}`;

function skillFile(dir: string): string {
  return join(dir, SKILL_NAME, "SKILL.md");
}

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

// Main-checkout fallback for a linked git worktree: <worktree>/.git is a file
// reading "gitdir: <main>/.git/worktrees/<name>". Applies only when the
// worktree root has no .claude/skills directory of its own.
function claudeWorktreeSkill(fromDir: string): boolean {
  for (let dir = resolve(fromDir); ; dir = dirname(dir)) {
    let gitdir: string | undefined;
    try {
      const text = readFileSync(join(dir, ".git"), "utf8");
      const prefix = "gitdir: ";
      if (text.startsWith(prefix)) gitdir = resolve(dir, text.slice(prefix.length).trim());
    } catch {
      gitdir = undefined;
    }
    if (gitdir !== undefined) {
      const splitAt = gitdir.replace(/\\/g, "/").indexOf("/.git/worktrees/");
      if (splitAt === -1 || existsSync(join(dir, ".claude", "skills"))) return false;
      return existsSync(skillFile(join(gitdir.slice(0, splitAt), ".claude", "skills")));
    }
    if (existsSync(join(dir, ".git")) || dirname(dir) === dir) return false;
  }
}

export function skillInstalled(row: SkillHostRow, fromDir: string): boolean {
  if ([...row.userDirs, ...row.systemDirs].some(dir => existsSync(skillFile(dir)))) return true;
  for (let dir = resolve(fromDir); ; dir = dirname(dir)) {
    if (row.projectDirs.some(projectDir => existsSync(skillFile(join(dir, projectDir))))) return true;
    if (dirname(dir) === dir) break;
  }
  if (row.name === "claude-code") return claudeWorktreeSkill(fromDir);
  return false;
}

export function skillHint(env: NodeJS.ProcessEnv = process.env, fromDir: string = process.cwd()): string | undefined {
  const host = detectSkillHost(env, resolveHome(env));
  if (host === undefined || skillInstalled(host, fromDir)) return undefined;
  return SKILL_INSTALL_COMMAND;
}
