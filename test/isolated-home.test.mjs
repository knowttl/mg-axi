import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";

for (const marker of ["CLAUDECODE", "CODEX_THREAD_ID", "CODEX_SESSION_ID", "PI_CODING_AGENT", "OPENCODE"]) {
  test(`validation commands do not inherit ${marker}`, () => {
    const result = spawnSync(process.execPath, [resolve("tools/isolated-home.mjs"), process.execPath, "--input-type=module", "-e", `
      import assert from "node:assert/strict";
      import { executeArgv } from "./dist/cli.js";
      import { skillHomeHints } from "./dist/docs.js";
      assert.equal(process.env[${JSON.stringify(marker)}], undefined);
      assert.deepEqual((await executeArgv([])).help, skillHomeHints());
    `, resolve("dist/bin/mg-axi.js")], {
      encoding: "utf8", timeout: 10000,
      env: { ...process.env, [marker]: "fixture-agent-session" },
    });
    assert.equal(result.status, 0, result.stderr || result.stdout);
  });
}

test("validation commands cannot write to the inherited home or XDG config", () => {
  const sentinel = mkdtempSync(join(tmpdir(), "mg-axi-owner-sentinel-"));
  try {
    writeFileSync(join(sentinel, "owner"), "untouched");
    const result = spawnSync(process.execPath, [resolve("tools/isolated-home.mjs"), process.execPath, "--input-type=module", "-e", `
      import { mkdirSync, writeFileSync } from "node:fs";
      import { homedir } from "node:os";
      import { join } from "node:path";
      for (const dir of [homedir(), process.env.XDG_CONFIG_HOME, process.env.XDG_STATE_HOME, process.env.XDG_CACHE_HOME, process.env.APPDATA, process.env.LOCALAPPDATA]) {
        mkdirSync(dir, { recursive: true });
        writeFileSync(join(dir, "validation-write"), "test");
      }
      if (process.env.MG_AXI_CONFIG !== undefined || process.env.MG_AXI_TOKEN !== undefined || process.env.AZURE_FEDERATED_TOKEN_FILE !== undefined) throw new Error("Credential environment leaked");
    `], {
      encoding: "utf8", timeout: 10000,
      env: { ...process.env, HOME: sentinel, USERPROFILE: sentinel, XDG_CONFIG_HOME: sentinel, XDG_STATE_HOME: sentinel, XDG_CACHE_HOME: sentinel, APPDATA: sentinel, LOCALAPPDATA: sentinel, MG_AXI_CONFIG: join(sentinel, "owner-config.json"), MG_AXI_TOKEN: "fixture-owner-token", AZURE_FEDERATED_TOKEN_FILE: join(sentinel, "owner-assertion") },
    });
    assert.equal(result.status, 0, result.stderr);
    assert.deepEqual(readdirSync(sentinel), ["owner"]);
    assert.equal(readFileSync(join(sentinel, "owner"), "utf8"), "untouched");
  } finally { rmSync(sentinel, { recursive: true, force: true }); }
});
