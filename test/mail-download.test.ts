import type { BinaryTransport } from "../dist/mail-download.js";
import { record, text, authFakes, type CredentialStub, type CredentialArgs, type ProfileConfig } from "./fixtures/typed.ts";
import type { TestContext } from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import { binaryFetchTransport, MAX_DOWNLOAD_BYTES } from "../dist/mail-download.js";

const bytes = Buffer.from([0, 255, 128, 13, 10]);
const response = (body: Uint8Array, headers: Record<string, string> = {}) => ({ status: 200, headers, body: (async function* () { yield body; })() });
function fixture(t: TestContext, attachment = false, options: {
  packs?: readonly string[];
  writes?: ProfileConfig["profiles"][string]["writes"];
  transport?: (request: Parameters<BinaryTransport>[0], count: number) => Awaited<ReturnType<BinaryTransport>>;
} = {}) {
  const directory = mkdtempSync(join(tmpdir(), "mg-axi-download-"));
  t.after(() => rmSync(directory, { recursive: true, force: true }));
  const config = join(directory, "config.json"), out = join(directory, "download"), journalPath = join(directory, "writes.log");
  const profiles = new Profiles(config);
  profiles.create("soc", "11111111-1111-4111-8111-111111111111", "22222222-2222-4222-8222-222222222222", "commercial", false);
  profiles.create("batch", "11111111-1111-4111-8111-111111111111", "22222222-2222-4222-8222-222222222222", "commercial", false, { federated: true });
  const configuration: ProfileConfig = JSON.parse(readFileSync(config, "utf8"));
  for (const profile of Object.values(configuration.profiles)) {
    profile.enabledPacks = options.packs ?? ["mail"];
    profile.writes = options.writes ?? { allowWrites: true, operations: ["mail.attachment.save", "mail.message.export"] };
  }
  writeFileSync(config, JSON.stringify(configuration));
  const calls: Parameters<BinaryTransport>[0][] = [], credentials: CredentialArgs[1][] = [];
  const auth: CredentialStub = { credential: async (profile, scopes) => { credentials.push(scopes); return { token: "opaque-fixture-token", username: "mailbox@contoso.example" }; } };
  const run = (extra: readonly string[] = [], profile = "soc", mailbox = "me") => executeArgv([
    "mail", attachment ? "attachment" : "message", attachment ? "save" : "export", "--mailbox", mailbox,
    "--id", "Message+ID=", ...(attachment ? ["--attachment", "Attachment+ID="] : []), "--out", out,
    "--profile", profile, "--config", config, ...extra,
  ], { delegated: authFakes(auth).delegated, application: authFakes(auth).application, journalPath, binaryTransport: async request => {
    calls.push(request);
    assert.equal(JSON.parse(text(readFileSync(journalPath, "utf8").trim().split("\n").at(-1))).kind, "intent");
    if (options.transport) return options.transport(request, calls.length);
    return request.url.includes("?$select") ? response(Buffer.from(JSON.stringify({ "@odata.type": "#microsoft.graph.fileAttachment", size: bytes.length })))
      : response(bytes, { "content-type": "application/octet-stream", "content-length": String(bytes.length) });
  }, transport: async () => { throw new Error("Unexpected text transport"); }, mutationTransport: async () => { throw new Error("Unexpected mutation"); } });
  const execute = ["--execute", "--confirm", attachment ? "Attachment+ID=" : "Message+ID="];
  return { run, execute, out, calls, credentials, journalPath, directory };
}

for (const attachment of [false, true]) {
  test(`${attachment ? "attachment" : "MIME"} preview has no credentials, file or audit`, async t => {
    const f = fixture(t, attachment);
    assert.deepEqual(record(await f.run()), { status: "dry-run" });
    assert.equal(f.credentials.length, 0); assert.equal(f.calls.length, 0);
    assert.equal(existsSync(f.out), false); assert.equal(existsSync(f.journalPath), false);
  });
  for (const [profile, mailbox] of [["soc", "me"], ["soc", "mailbox@contoso.example"], ["batch", "mailbox@contoso.example"]] as const) {
    test(`${attachment ? "attachment" : "MIME"} saves exact bytes with ${profile}/${mailbox}`, async t => {
      const f = fixture(t, attachment);
      const result = record(await f.run(f.execute, profile, mailbox));
      assert.deepEqual(result, { path: f.out, size: bytes.length, sha256: createHash("sha256").update(bytes).digest("hex"), contentType: "application/octet-stream", kind: attachment ? "attachment" : "mime" });
      assert.deepEqual(readFileSync(f.out), bytes); assert.equal(statSync(f.out).mode & 0o777, 0o600);
      assert.deepEqual(f.credentials, [profile === "soc" ? ["https://graph.microsoft.com/Mail.Read"] : undefined]);
      assert.equal(f.calls.length, attachment ? 2 : 1);
      assert.ok(f.calls.at(-1)!.url.endsWith(attachment ? "/attachments/Attachment%2BID%3D/$value" : "/messages/Message%2BID%3D/$value"));
      const journal = readFileSync(f.journalPath, "utf8");
      assert.deepEqual(journal.trim().split("\n").map(line => record(JSON.parse(line)).kind), ["intent", "outcome"]);
      assert.ok(!/Message|Attachment|contoso|opaque-fixture/.test(journal));
    });
  }
}

for (const [name, options, extra, profile, mailbox] of [
  ["mail pack disabled", { packs: ["entra"] }, null],
  ["writes disabled", { writes: { allowWrites: false, operations: ["mail.message.export"] } }, null],
  ["operation not opted in", { writes: { allowWrites: true, operations: ["other.operation"] } }, null],
  ["missing confirmation", {}, ["--execute"]],
  ["wrong confirmation", {}, ["--execute", "--confirm", "wrong"]],
  ["application me", {}, null, "batch", "me"],
] as const) test(`${name} refuses before credentials or network`, async t => {
  const f = fixture(t, false, options);
  await assert.rejects(f.run(extra ?? f.execute, profile, mailbox));
  assert.equal(f.credentials.length, 0); assert.equal(f.calls.length, 0); assert.equal(existsSync(f.out), false);
});

for (const symlink of [false, true]) test(`existing ${symlink ? "symlink" : "file"} is never overwritten`, async t => {
  const f = fixture(t);
  const original = join(f.directory, "original"); writeFileSync(original, "keep");
  if (symlink) symlinkSync(original, f.out); else writeFileSync(f.out, "keep");
  await assert.rejects(f.run(f.execute));
  assert.equal(f.credentials.length, 0); assert.equal(f.calls.length, 0); assert.equal(readFileSync(original, "utf8"), "keep");
  assert.equal(readFileSync(f.out, "utf8"), "keep");
});

test("delegated foreign mailbox is refused before network", async t => {
  const f = fixture(t);
  await assert.rejects(f.run(f.execute, "soc", "other@contoso.example"));
  assert.equal(f.calls.length, 0); assert.equal(existsSync(f.out), false);
});

test("forced read-only refuses downloads before credentials", async t => {
  const prior = process.env.MG_AXI_READ_ONLY;
  t.after(() => { if (prior === undefined) delete process.env.MG_AXI_READ_ONLY; else process.env.MG_AXI_READ_ONLY = prior; });
  process.env.MG_AXI_READ_ONLY = "1";
  const f = fixture(t);
  await assert.rejects(f.run(f.execute));
  assert.equal(f.credentials.length, 0); assert.equal(f.calls.length, 0);
});

test("item attachments can be saved", async t => {
  const f = fixture(t, true, { transport: request => request.url.includes("?$select")
    ? response(Buffer.from(JSON.stringify({ "@odata.type": "#microsoft.graph.itemAttachment", size: bytes.length })))
    : response(bytes, { "content-type": "message/rfc822" }) });
  assert.equal(record(await f.run(f.execute)).contentType, "message/rfc822");
  assert.deepEqual(readFileSync(f.out), bytes);
});

test("exactly 50 MiB is accepted with the streaming hash", async t => {
  const chunk = Buffer.alloc(1024 * 1024, 7);
  const hash = createHash("sha256");
  for (let i = 0; i < 50; i++) hash.update(chunk);
  const f = fixture(t, false, { transport: () => ({ status: 200, headers: { "content-length": String(MAX_DOWNLOAD_BYTES) },
    body: (async function* () { for (let i = 0; i < 50; i++) yield chunk; })() }) });
  const result = record(await f.run(f.execute));
  assert.equal(result.size, MAX_DOWNLOAD_BYTES); assert.equal(result.sha256, hash.digest("hex"));
  assert.equal(statSync(f.out).size, MAX_DOWNLOAD_BYTES);
});

test("binary fetch preserves bytes and refuses automatic redirects", async t => {
  const previous = globalThis.fetch;
  t.after(() => { globalThis.fetch = previous; });
  globalThis.fetch = async (url, options) => {
    assert.equal(url, "https://graph.microsoft.com/v1.0/me/messages/fixture/$value");
    assert.equal(options!.redirect, "manual"); assert.equal(options!.method, "GET");
    return new Response(bytes, { headers: { "Content-Type": "message/rfc822" } });
  };
  const result = await binaryFetchTransport({ method: "GET", url: "https://graph.microsoft.com/v1.0/me/messages/fixture/$value", headers: {}, signal: new AbortController().signal });
  const chunks = [];
  for await (const chunk of result.body) chunks.push(chunk);
  assert.deepEqual(Buffer.concat(chunks), bytes); assert.equal(result.headers["content-type"], "message/rfc822");
});

for (const type of ["#microsoft.graph.referenceAttachment", "#unknown"]) test(`refuses ${type} before binary request`, async t => {
  const f = fixture(t, true, { transport: () => response(Buffer.from(JSON.stringify({ "@odata.type": type, size: 1 }))) });
  await assert.rejects(f.run(f.execute)); assert.equal(f.calls.length, 1); assert.equal(existsSync(f.out), false);
});

for (const [name, transport] of [
  ["reported attachment overflow", () => response(Buffer.from(JSON.stringify({ "@odata.type": "#microsoft.graph.fileAttachment", size: MAX_DOWNLOAD_BYTES + 1 })))],
  ["reported body overflow", () => response(bytes, { "content-length": String(MAX_DOWNLOAD_BYTES + 1) })],
  ["stream overflow", () => ({ status: 200, headers: {}, body: (async function* () { for (let i = 0; i < 51; i++) yield Buffer.alloc(1024 * 1024); })() })],
  ["stream failure", () => ({ status: 200, headers: {}, body: (async function* () { yield bytes; throw new Error("sensitive mail content"); })() })],
  ["length mismatch", () => response(bytes, { "content-length": "6" })],
  ["redirect", () => ({ ...response(bytes), status: 302, headers: { location: "https://untrusted.example" } })],
  ["throttle", () => ({ ...response(bytes), status: 429 })],
  ["server error", () => ({ ...response(bytes), status: 503 })],
] as const) test(`${name} removes partial file and never retries`, async t => {
  const f = fixture(t, name === "reported attachment overflow", { transport });
  await assert.rejects(f.run(f.execute), error => !text(record(error).message).includes("sensitive mail content"));
  assert.equal(f.calls.length, 1); assert.equal(existsSync(f.out), false);
  assert.equal(JSON.parse(text(readFileSync(f.journalPath, "utf8").trim().split("\n").at(-1))).outcome, "FAILED");
});
