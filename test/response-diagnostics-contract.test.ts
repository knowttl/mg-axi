import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { test } from "node:test";
import { ResponseDiagnostics } from "../dist/response-diagnostics.js";
import vectors from "./fixtures/response-diagnostics-vectors.json" with { type: "json" };
import hashes from "./fixtures/response-diagnostics-sha256.json" with { type: "json" };

// Canonical source and portable vectors: knowttl/beyondtrust-axi, main.
// These bytes are the distribution contract, verified without network access.
for (const [path, hash] of Object.entries(hashes)) test(`canonical copy SHA-256: ${path}`, () => {
  assert.equal(createHash("sha256").update(readFileSync(new URL(`../${path}`, import.meta.url))).digest("hex"), hash);
});

for (const vector of vectors) test(`shared diagnostics vector: ${vector.name}`, () => {
  const redactor = { text: (text: string) => (vector.secrets ?? []).reduce((text, secret) => text.replaceAll(secret, "***redacted***"), text) };
  const diagnostics = new ResponseDiagnostics(redactor, ["RESPONSE_INVALID"]);
  diagnostics.record(vector.status, new Headers(vector.contentType ? { "content-type": vector.contentType } : {}), vector.data);
  if (vector.issues) diagnostics.validationFailed(vector.issues);
  assert.deepEqual(diagnostics.forError({ code: "RESPONSE_INVALID" }), vector.expected);
  assert.deepEqual(diagnostics.forError({ code: "GRAPH_ERROR" }), {});
});
