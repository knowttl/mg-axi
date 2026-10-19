import { accessSync, constants, existsSync, realpathSync, statSync } from "node:fs";
import { delimiter, dirname, join, resolve } from "node:path";

function quote(value: string): string {
  return /^[A-Za-z0-9_./:\\-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`;
}

// Prefer a stable package invocation for npm's disposable cache. A PATH name
// is safe only when it resolves to this executable, not another installation.
export function invocation(entry = process.argv[1], env = process.env): { command: string; checkout: boolean } {
  if (!entry) return { command: "npx -y @knowttl/mg-axi", checkout: false };
  const target = realpathSync(entry);
  if (target.split(/[/\\]/).includes("_npx")) return { command: "npx -y @knowttl/mg-axi", checkout: false };
  const checkout = existsSync(join(dirname(target), "..", "..", ".git"));
  for (const directory of (env.PATH ?? "").split(delimiter)) {
    const candidate = join(directory, process.platform === "win32" ? "mg-axi.cmd" : "mg-axi");
    try {
      accessSync(candidate, constants.X_OK);
      if (!statSync(candidate).isFile()) continue;
      if (realpathSync(candidate) === target) return { command: "mg-axi", checkout: false };
      break;
    } catch { /* Missing or unrelated PATH entries do not identify this binary. */ }
  }
  const script = target === resolve("dist/bin/mg-axi.js") ? "dist/bin/mg-axi.js" : target;
  return { command: `${quote(process.execPath)} ${quote(script)}`, checkout };
}
