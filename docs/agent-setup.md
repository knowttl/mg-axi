# Agent setup

Install the [generated skill](../skills/mg-axi/SKILL.md), which includes the complete executable command table:

```sh
npx skills add knowttl/mg-axi --skill mg-axi
```

For optional session context in Claude Code, Codex and OpenCode, use the global installation:

```sh
mg-axi setup hooks
```

Restart the agent session.
Hooks inject local profile names, tenant labels, authentication-cache presence, write policy and version; they make no Graph or sign-in calls.
Skills and hooks are independent discovery options and grant no permissions.
`setup` only displays guidance; only `setup hooks` installs integrations.
The hook installer rejects temporary, linked-worktree and validation-build targets.

There is no CLI hook-uninstall command.
To remove the integration, delete only the `mg-axi-hook` SessionStart entries from `~/.claude/settings.json` and `~/.codex/hooks.json`, and `~/.config/opencode/plugins/axi-mg-axi-hook.js`.
Preserve other hooks and shared Codex hook settings.


## Discovery defaults

Skill discovery uses HOME, then USERPROFILE, then the operating-system home directory.
CLAUDE_CONFIG_DIR defaults to ~/.claude; CODEX_HOME defaults to ~/.codex.
OpenCode skill discovery uses OPENCODE_CONFIG_DIR, then XDG_CONFIG_HOME, then ~/.config, with opencode/skills appended.
CLAUDECODE, CODEX_THREAD_ID, CODEX_SESSION_ID, PI_CODING_AGENT and OPENCODE detect the active agent by presence.
PATH is used to identify the installed executable.
