// Runs before compilation: discovery dispositions identify the shipped surface.
import { readFileSync, writeFileSync } from "node:fs";

const inventory = JSON.parse(readFileSync(new URL("../inventory/operations.json", import.meta.url), "utf8"));
const target = new URL("../inventory/runtime.json", import.meta.url);
const operations = inventory.operations
  .filter(row => row.disposition === "named-command" || row.disposition === "reviewed-raw-read")
  .map(({ id, method, path, version, disposition, owningSlice }) => ({ id, method, path, version, disposition, owningSlice }))
  .sort((a, b) => a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
const output = `${JSON.stringify({ operations }, null, 2)}\n`;
if (process.argv.includes("--check")) {
  if (readFileSync(target, "utf8") !== output) {
    process.stderr.write("Stale runtime index. Run: node tools/generate-runtime-index.mjs\n");
    process.exitCode = 1;
  }
} else {
  writeFileSync(target, output);
}
