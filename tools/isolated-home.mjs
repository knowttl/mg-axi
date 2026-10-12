import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

// Validation commands must never inherit the user's agent configuration.
const home = mkdtempSync(join(tmpdir(), "mg-axi-validation-"));
try {
  const [command, ...args] = process.argv.slice(2);
  if (!command) throw new Error("Expected a validation command");
  const result = spawnSync(command, args, {
    stdio: "inherit",
    env: {
      ...process.env,
      HOME: home,
      USERPROFILE: home,
      XDG_CONFIG_HOME: join(home, ".config"),
      XDG_CACHE_HOME: join(home, ".cache"),
      XDG_DATA_HOME: join(home, ".local", "share"),
      CLAUDE_CONFIG_DIR: join(home, ".claude"),
      OPENCODE_CONFIG_DIR: join(home, ".config", "opencode"),
      NODE_ENV: "test",
    },
  });
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} finally {
  rmSync(home, { recursive: true, force: true });
}
