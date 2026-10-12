import * as fs from "node:fs";
import { mock } from "node:test";

// Model a permanent installed executable while keeping every write in the
// disposable HOME. Only the filesystem's canonical-path boundary is faked.
mock.module("node:fs", {
  namedExports: {
    ...fs,
    realpathSync: path => String(path).endsWith("/mg-axi-hook.js")
      ? "/opt/mg-axi/dist/bin/mg-axi-hook.js"
      : fs.realpathSync(path),
  },
});
