import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { LEAVES, operationFor } from "../dist/catalogue.js";
import { REVIEWED_ROUTES } from "../dist/api.js";
import { resolveSessionOperation } from "../dist/graph-session.js";
import { TRANSITIVE_OPERATION } from "../dist/entra-groups.js";
import { MAIL_FOLDER_ALTERNATES } from "../dist/mail-folders.js";
import { MAIL_MESSAGE_ALTERNATES } from "../dist/mail-messages.js";
import { MAIL_MESSAGE_SHOW_ALTERNATES } from "../dist/mail-message-show.js";
import { MAIL_METADATA_ALTERNATES } from "../dist/mail-metadata.js";
import { MAIL_WRITE_ALTERNATES } from "../dist/mail-message-writes.js";
import { MAIL_FOLDER_WRITE_ALTERNATES } from "../dist/mail-folder-writes.js";
import { MAIL_RULE_WRITE_ALTERNATES } from "../dist/mail-rule-writes.js";
import { MAIL_DOWNLOAD_ALTERNATES } from "../dist/mail-download.js";
import { castDirectReports, castMembership, transitMembership, NAV_CAST_VALUES, MEMBERSHIP_CAST_VALUES } from "../dist/entra-contacts.js";
import type { SessionOperation } from "../dist/graph-session.js";

const runtime = (JSON.parse(readFileSync(new URL("../inventory/runtime.json", import.meta.url), "utf8")) as { operations: SessionOperation[] }).operations;
const discovery = new Map((JSON.parse(readFileSync(new URL("../inventory/operations.json", import.meta.url), "utf8")) as { operations: SessionOperation[] }).operations.map(row => [row.id, row]));

test("generator works without compiled files and detects stale discovery policy", () => {
  const root = mkdtempSync(join(tmpdir(), "mg-axi-index-"));
  try {
    mkdirSync(join(root, "tools"));
    mkdirSync(join(root, "inventory"));
    const script = join(root, "tools/generate-runtime-index.ts");
    copyFileSync(new URL("../tools/generate-runtime-index.ts", import.meta.url), script);
    const row = { id: "v1.0:GET:/users", method: "GET", path: "/users", version: "v1.0", disposition: "named-command", owningSlice: "READ-01" };
    const inventory = join(root, "inventory/operations.json");
    writeFileSync(inventory, JSON.stringify({ operations: [row, { ...row, id: "beta:GET:/users", version: "beta", disposition: "scheduled" }] }));
    const generated = spawnSync(process.execPath, [script], { encoding: "utf8", timeout: 30000 });
    assert.equal(generated.status, 0, generated.stderr);
    assert.deepEqual(JSON.parse(readFileSync(join(root, "inventory/runtime.json"), "utf8")).operations, [row]);
    assert.equal(spawnSync(process.execPath, [script, "--check"]).status, 0);
    writeFileSync(inventory, JSON.stringify({ operations: [{ ...row, disposition: "deferred" }] }));
    const stale = spawnSync(process.execPath, [script, "--check"], { encoding: "utf8", timeout: 30000 });
    assert.equal(stale.status, 1);
    assert.match(stale.stderr, /Stale runtime index/);
  } finally { rmSync(root, { recursive: true, force: true }); }
});

test("runtime rows exactly cover named and raw routes including flag-selected variants", () => {
  const reachable = new Set(REVIEWED_ROUTES.map(row => row.id));
  for (const leaf of LEAVES.filter((leaf): leaf is typeof leaf & { operation: string } => !!leaf.operation)) {
    reachable.add(`v1.0:${leaf.operation}`);
    const alternate = TRANSITIVE_OPERATION[leaf.operation];
    if (alternate) reachable.add(`v1.0:${alternate}`);
    const mailAlternate = MAIL_FOLDER_ALTERNATES[leaf.operation];
    if (mailAlternate) reachable.add(`v1.0:${mailAlternate}`);
    const showAlternate = MAIL_MESSAGE_SHOW_ALTERNATES[leaf.operation];
    const metadataAlternate = MAIL_METADATA_ALTERNATES[leaf.operation];
    if (metadataAlternate) reachable.add(`v1.0:${metadataAlternate}`);
    if (showAlternate) reachable.add(`v1.0:${showAlternate}`);
    const writeAlternate = MAIL_WRITE_ALTERNATES[leaf.operation];
    if (writeAlternate) reachable.add(`v1.0:${writeAlternate}`);
    const folderWriteAlternate = MAIL_FOLDER_WRITE_ALTERNATES[leaf.operation];
    if (folderWriteAlternate) reachable.add(`v1.0:${folderWriteAlternate}`);
    const ruleWriteAlternate = MAIL_RULE_WRITE_ALTERNATES[leaf.operation];
    if (ruleWriteAlternate) reachable.add(`v1.0:${ruleWriteAlternate}`);
    const downloadAlternate = MAIL_DOWNLOAD_ALTERNATES[leaf.operation as keyof typeof MAIL_DOWNLOAD_ALTERNATES];
    for (const operation of typeof downloadAlternate === "string" ? [downloadAlternate] : Object.values(downloadAlternate ?? {})) reachable.add(`v1.0:${operation}`);
    if (leaf.path === "mail message list") {
      for (const alternate of Object.values(MAIL_MESSAGE_ALTERNATES)) reachable.add(`v1.0:${alternate}`);
    }
    if (leaf.path.startsWith("entra contact ")) {
      const bases = [leaf.operation];
      if (leaf.flags.transitive) bases.push(transitMembership(leaf.operation, ""));
      for (const base of bases) {
        reachable.add(`v1.0:${base}`);
        if (leaf.flags.as) {
          const membership = leaf.path.endsWith("-member-of");
          for (const cast of membership ? MEMBERSHIP_CAST_VALUES : NAV_CAST_VALUES) {
            reachable.add(`v1.0:${(membership ? castMembership : castDirectReports)(base, cast, "")}`);
          }
        }
      }
    }
  }
  assert.deepEqual(runtime.map(row => row.id).sort(), [...reachable].sort());
  for (const row of runtime) {
    const fields = ["id", "method", "path", "version", "disposition", "owningSlice"] as const;
    assert.deepEqual(row, Object.fromEntries(fields.map(key => [key, discovery.get(row.id)![key]])), row.id);
    assert.deepEqual(operationFor({ operation: row.id.slice("v1.0:".length) } as Parameters<typeof operationFor>[0], "v1.0"), row);
    assert.deepEqual(resolveSessionOperation(row.version, row.method, row.path), row);
  }
});

test("every discovery-only operation is absent from command and session lookup", () => {
  const shipped = new Set(runtime.map(row => row.id));
  for (const row of discovery.values()) {
    if (shipped.has(row.id)) continue;
    assert.equal(operationFor({ operation: `${row.method}:${row.path}` } as Parameters<typeof operationFor>[0], row.version), undefined);
    assert.throws(() => resolveSessionOperation(row.version, row.method, row.path), { code: "VALIDATION_ERROR" });
  }
});

test("command and session lookup share one read and protect cached rows from callers", () => {
  const result = spawnSync(process.execPath, ["--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import fs from "node:fs";
    import { syncBuiltinESMExports } from "node:module";
    const read = fs.readFileSync;
    let reads = 0;
    fs.readFileSync = function(path, ...args) {
      if (String(path).endsWith("/inventory/runtime.json")) reads++;
      return read.call(this, path, ...args);
    };
    syncBuiltinESMExports();
    const { operationFor } = await import("./dist/catalogue.js");
    const { resolveSessionOperation } = await import("./dist/graph-session.js");
    const leaf = { operation: "GET:/users" };
    const expected = operationFor(leaf, "v1.0");
    const changed = resolveSessionOperation("v1.0", "GET", "/users");
    changed.disposition = "scheduled";
    changed.path = "/me/messages";
    for (let i = 0; i < 10; i++) {
      assert.deepEqual(operationFor(leaf, "v1.0"), expected);
      assert.deepEqual(resolveSessionOperation("v1.0", "GET", "/users"), expected);
    }
    assert.equal(reads, 1);
  `], { encoding: "utf8", timeout: 30000 });
  assert.equal(result.status, 0, result.stderr);
});
