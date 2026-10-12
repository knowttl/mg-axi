import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";

test("validation commands cannot write to the inherited home or XDG config", () => {
  const sentinel = mkdtempSync(join(tmpdir(), "mg-axi-owner-sentinel-"));
  try {
    writeFileSync(join(sentinel, "owner"), "untouched");
    const result = spawnSync(process.execPath, [resolve("tools/isolated-home.mjs"), process.execPath, "--input-type=module", "-e", `
      import { mkdirSync, writeFileSync } from "node:fs";
      import { homedir } from "node:os";
      import { join } from "node:path";
      for (const dir of [homedir(), process.env.XDG_CONFIG_HOME]) {
        mkdirSync(dir, { recursive: true });
        writeFileSync(join(dir, "validation-write"), "test");
      }
    `], {
      encoding: "utf8", timeout: 10000,
      env: { ...process.env, HOME: sentinel, USERPROFILE: sentinel, XDG_CONFIG_HOME: sentinel },
    });
    assert.equal(result.status, 0, result.stderr);
    assert.deepEqual(readdirSync(sentinel), ["owner"]);
    assert.equal(readFileSync(join(sentinel, "owner"), "utf8"), "untouched");
  } finally { rmSync(sentinel, { recursive: true, force: true }); }
});
