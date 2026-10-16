import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { skillHostRows, skillInstalled } from "../dist/skill-hint.js";

const row = overrides => ({ name: "fixture", userDirs: [], projectDirs: [], systemDirs: [], ...overrides });

test("system folders count as installed without a walk-up", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-skill-hint-"));
  try {
    mkdirSync(join(dir, "mg-axi"), { recursive: true });
    writeFileSync(join(dir, "mg-axi", "SKILL.md"), "---\nname: mg-axi\n---\n");
    assert.equal(skillInstalled(row({ systemDirs: [dir] }), dir), true);
    assert.equal(skillInstalled(row({ systemDirs: [join(dir, "missing")] }), dir), false);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test("codex row lists the documented admin skills folder", () => {
  const codex = skillHostRows({}, "/nonexistent-home").find(host => host.name === "codex");
  assert.deepEqual(codex.systemDirs, ["/etc/codex/skills"]);
});

test("config-directory overrides reshape their rows", () => {
  const rows = skillHostRows(
    { CLAUDE_CONFIG_DIR: "/custom-claude", CODEX_HOME: "/custom-codex", XDG_CONFIG_HOME: "/custom-config", OPENCODE_CONFIG_DIR: "/custom-opencode" },
    "/nonexistent-home",
  );
  const byName = Object.fromEntries(rows.map(host => [host.name, host]));
  assert.equal(byName["claude-code"].userDirs[0], join("/custom-claude", "skills"));
  assert.equal(byName["codex"].userDirs[0], join("/custom-codex", "skills"));
  assert.equal(byName["opencode"].userDirs[0], join("/custom-opencode", "opencode", "skills"));
});
