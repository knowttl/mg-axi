// Regenerate PACK-01 release records from their single sources in src/.
// Usage: node tools/generate-docs.mjs [--check]
// --check exits 1 when a committed file differs instead of rewriting it.
// Build first: this imports the compiled catalogue and inventory readers.
import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { capabilityDocument, skillDocument } from "../dist/docs.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const skillPath = join(root, "skills", "mg-axi", "SKILL.md");
const coveragePath = join(root, "docs", "coverage.md");

const check = process.argv.includes("--check");
const skill = skillDocument();
const coverage = capabilityDocument();
if (check) {
  const stale = [
    ...(readFileSync(skillPath, "utf8").replace(/\r\n/g, "\n") === skill ? [] : [skillPath]),
    ...(readFileSync(coveragePath, "utf8").replace(/\r\n/g, "\n") === coverage ? [] : [coveragePath]),
  ];
  if (stale.length > 0) {
    process.stderr.write(`Stale generated docs: ${stale.join(", ")}. Run: node tools/generate-docs.mjs\n`);
    process.exit(1);
  }
} else {
  writeFileSync(skillPath, skill);
  writeFileSync(coveragePath, coverage);
}
