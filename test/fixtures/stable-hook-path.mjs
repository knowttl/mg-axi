import fs from "node:fs";
import { syncBuiltinESMExports } from "node:module";
import { mock } from "node:test";

// Model a permanent installed executable while keeping every write in the
// disposable HOME. Fake canonical paths and, when requested, a worktree marker.
const realpathSync = fs.realpathSync;
mock.method(fs, "realpathSync", path => String(path).endsWith("/mg-axi-hook.js")
  ? "/opt/mg-axi/dist/bin/mg-axi-hook.js"
  : realpathSync(path));

if (process.env.MG_AXI_TEST_WORKTREE === "1") {
  const existsSync = fs.existsSync;
  const statSync = fs.statSync;
  mock.method(fs, "existsSync", path => String(path) === "/opt/mg-axi/.git" || existsSync(path));
  mock.method(fs, "statSync", path => String(path) === "/opt/mg-axi/.git" ? { isFile: () => true } : statSync(path));
}
syncBuiltinESMExports();
