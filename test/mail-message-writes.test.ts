import type { MutationTransportRequest, MutationTransportResponse } from "../dist/mutations.js";
import { record, records, text, strings, authFakes, type CredentialStub, type CredentialArgs, type ProfileConfig } from "./fixtures/typed.ts";
import type { TransportRequest } from "../dist/graph-session.js";
import type { TestContext } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync, existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { decode, encode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const mailbox = "mailbox@contoso.example";
const actions = [
  { action: "tag", flags: ["--add", "Reviewed", "--remove", "Pending"], field: "categories", current: ["Pending", "Keep"], desired: ["Keep", "Reviewed"] },
  { action: "flag", flags: ["--status", "flagged"], field: "flag", current: { flagStatus: "notFlagged" }, desired: { flagStatus: "flagged" } },
  { action: "mark", flags: ["--state", "read"], field: "isRead", current: false, desired: true },
];
const transfers = [
  { action: "move", flags: ["--destination", "Folder+ID="], destination: "Folder+ID=" },
  { action: "copy", flags: ["--destination", "archive"], destination: "archive" },
  { action: "trash", flags: [], destination: "deleteditems" },
];
function fixture(t: TestContext, spec: { action: string; flags: readonly string[]; field?: string; current?: unknown }, options: {
  username?: (count: number) => string;
  read?: (id: string, count: number) => unknown;
  write?: (request: MutationTransportRequest, count: number) => MutationTransportResponse;
} = {}) {
  const directory = mkdtempSync(join(tmpdir(), "mg-axi-mail-writes-"));
  const previousExitCode = process.exitCode;
  t.after(() => { process.exitCode = previousExitCode; rmSync(directory, { recursive: true, force: true }); });
  const config = join(directory, "config.json"), journalPath = join(directory, "writes.log");
  const profiles = new Profiles(config);
  const tenant = "11111111-1111-4111-8111-111111111111", client = "22222222-2222-4222-8222-222222222222";
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  const configuration: ProfileConfig = JSON.parse(readFileSync(config, "utf8"));
  for (const profile of Object.values(configuration.profiles)) {
    profile.enabledPacks = ["mail"];
    profile.writes = { allowWrites: true, operations: [...actions, ...transfers].map(item => `mail.message.${item.action}`) };
  }
  writeFileSync(config, JSON.stringify(configuration));
  const reads: TransportRequest[] = [], writes: MutationTransportRequest[] = [], credentials: { profile: CredentialArgs[0]; scopes: CredentialArgs[1] }[] = [], events: string[] = [];
  const auth: CredentialStub = { credential: async (profile, scopes) => {
    credentials.push({ profile, scopes });
    return { token: "opaque-fixture-token", tenantId: tenant, clientId: client, accountId: "fixture", username: options.username?.(credentials.length) ?? mailbox };
  } };
  const run = (extra: readonly string[] = [], profile = "soc", target = "me", ids = "A+ID=,B+ID=") => executeArgv([
    "mail", "message", spec.action, "--mailbox", target, "--ids", ids, ...spec.flags, "--profile", profile, "--config", config, ...extra,
  ], { delegated: authFakes(auth).delegated, application: authFakes(auth).application, journalPath,
    transport: async request => {
      reads.push(request); events.push("read");
      const id = decodeURIComponent(new URL(request.url).pathname.split("/").at(-1)!);
      return { status: 200, headers: {}, body: JSON.stringify(options.read?.(id, reads.length) ?? { id, [spec.field!]: spec.current, body: { content: "unreviewed-body" } }) };
    }, mutationTransport: async request => {
      writes.push(request); events.push("write");
      assert.equal(JSON.parse(text(readFileSync(journalPath, "utf8").trim().split("\n").at(-1))).kind, "intent");
      return options.write ? options.write(request, writes.length) : { status: 200, headers: {}, body: '{"body":{"content":"unreviewed-response"}}' };
    },
  });
  const journal = () => existsSync(journalPath) ? readFileSync(journalPath, "utf8").trim().split("\n").map(line => record(JSON.parse(line))) : [];
  return { run, reads, writes, credentials, events, journal, config, configuration, journalPath };
}

for (const spec of transfers) {
  test(`${spec.action} previews only the target and destination`, async t => {
    const f = fixture(t, spec);
    const result = record(decode(encode(record(await f.run()))));
    assert.deepEqual(records(result.results).map(row => row.status), ["dry-run", "dry-run"]);
    assert.deepEqual(JSON.parse(text(record(records(result.results)[0]!.preview).currentState)), { id: "A+ID=" });
    assert.deepEqual(JSON.parse(text(record(records(result.results)[0]!.preview).proposedChange)), { destinationId: spec.destination });
    assert.equal(f.writes.length, 0); assert.equal(f.journal().length, 0);
    assert.ok(!JSON.stringify(result).includes("unreviewed-"));
    assert.equal(new URL(f.reads[0]!.url).searchParams.get("$select"), "id");
  });
  for (const [profile, target] of [["soc", "me"], ["soc", mailbox], ["batch", mailbox]] as const) test(`${spec.action} posts once per ID and returns the new ID in ${profile}/${target}`, async t => {
    const f = fixture(t, spec, { write: (_request, count) => ({ status: 201, headers: {}, body: JSON.stringify({ id: `new-${count}`, body: { content: "unreviewed-response" }, subject: "private-subject" }) }) });
    const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="], profile, target));
    assert.deepEqual(records(result.results).map(row => [row.id, row.newId, row.status]), [["A+ID=", "new-1", "success"], ["B+ID=", "new-2", "success"]]);
    assert.deepEqual(f.events, ["read", "read", "write", "read", "read", "write"]);
    assert.deepEqual(f.journal().map(row => row.kind), ["intent", "outcome", "intent", "outcome"]);
    assert.deepEqual(f.journal().filter(row => row.kind === "outcome").map(row => row.outcome), ["SUCCESS", "SUCCESS"]);
    for (const request of f.writes) {
      assert.equal(request.method, "POST");
      assert.deepEqual(JSON.parse(text(request.body)), { destinationId: spec.destination });
      assert.match(request.url, target === "me" ? /\/me\/messages\/[AB]%2BID%3D\/(move|copy)$/ : /\/users\/mailbox%40contoso.example\/messages\/[AB]%2BID%3D\/(move|copy)$/);
      assert.equal(new URL(request.url).pathname.split("/").at(-1), spec.action === "copy" ? "copy" : "move");
    }
    assert.equal(f.credentials.filter(item => item.scopes?.includes("https://graph.microsoft.com/Mail.ReadWrite")).length, profile === "soc" ? 2 : 0);
    assert.ok(!JSON.stringify(result).includes("unreviewed-")); assert.ok(!JSON.stringify(result).includes("private-subject"));
  });
  for (const status of [403, 429, 503]) test(`${spec.action} stops on ${status} without retry`, async t => {
    const f = fixture(t, spec, { write: () => ({ status, headers: { "retry-after": "0" }, body: '{"error":{"message":"private-echo"}}' }) });
    const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="]));
    assert.deepEqual(records(result.results).map(row => row.status), [status === 503 ? "unknown" : "failed", "not-attempted"]);
    assert.equal(f.writes.length, 1); assert.equal(result.complete, false); assert.equal(process.exitCode, 1);
    assert.ok(!JSON.stringify(result).includes("private-echo"));
  });
  test(`${spec.action} refuses confirmation and operation mismatches before calls`, async t => {
    const f = fixture(t, spec);
    await assert.rejects(f.run(["--execute", "--confirm", "B+ID=,A+ID="]), { code: "CONFIRM_MISMATCH" });
    f.configuration.profiles.soc!.writes.operations = ["mail.message.mark"];
    writeFileSync(f.config, JSON.stringify(f.configuration));
    await assert.rejects(f.run(), { code: "OPERATION_NOT_WRITABLE" });
    assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
  });
  test(`${spec.action} timeout stops the batch with an unknown audited outcome`, async t => {
    const f = fixture(t, spec, { write: () => { throw new Error("fixture timeout"); } });
    const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="]));
    assert.deepEqual(records(result.results).map(row => row.status), ["unknown", "not-attempted"]);
    assert.equal(f.writes.length, 1); assert.equal(f.journal().at(-1)!.outcome, "OUTCOME_UNKNOWN");
  });
}

for (const spec of transfers.slice(0, 2)) {
  for (const destination of ["", "../folder", "folder/child", "a,b", " folder", "folder?x=y"]) test(`${spec.action} refuses unsafe destination ${destination} before calls`, async t => {
    const f = fixture(t, { ...spec, flags: ["--destination", destination] });
    await assert.rejects(f.run(), { code: "VALIDATION_ERROR" });
    assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
  });
}

test("trash refuses destination overrides before calls", async t => {
  const f = fixture(t, transfers[2]!);
  await assert.rejects(f.run(["--destination", "archive"]));
  assert.equal(f.credentials.length, 0); assert.equal(f.writes.length, 0);
});

for (const body of ["{}", "[]", "null", '{"id":42}', '{"id":""}', "invalid json"]) test(`move missing response ID stops after accepted mutation (${body})`, async t => {
  const f = fixture(t, transfers[0]!, { write: () => ({ status: 201, headers: {}, body }) });
  const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="]));
  assert.deepEqual(records(result.results).map(row => row.status), ["unknown", "not-attempted"]);
  assert.equal(f.writes.length, 1); assert.equal(result.complete, false); assert.equal(process.exitCode, 1);
  assert.ok(records(result.results)[0]!.auditId);
});

test("accepted HTML mail mutation exposes only a safe unreadable warning", async t => {
  const f = fixture(t, actions[2]!, { write: () => ({ status: 200, headers: {}, body: "<html>synthetic-login opaque-fixture-token</html>" }) });
  const result = record(decode(encode(record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="])))));
  assert.deepEqual(records(result.results).map(row => row.status), ["success", "success"]);
  assert.deepEqual(records(result.results).map(row => row.responseUnreadable), [
    { code: "RESPONSE_INVALID", httpStatus: 200 },
    { code: "RESPONSE_INVALID", httpStatus: 200 },
  ]);
  assert.doesNotMatch(JSON.stringify(result), /synthetic-login|opaque-fixture-token|<html>/);
  assert.deepEqual(f.journal().filter(row => row.kind === "outcome").map(row => row.outcome), ["SUCCESS", "SUCCESS"]);
  assert.equal(f.writes.length, 2);
});

test("copy new ID truncates with full escape hatch", async t => {
  const f = fixture(t, transfers[1]!, { write: () => ({ status: 201, headers: {}, body: JSON.stringify({ id: "x".repeat(501) }) }) });
  const preview = record(await f.run(["--execute", "--confirm", "a"], "soc", "me", "a"));
  assert.match(text(records(preview.results)[0]!.newId), /truncated/);
  const full = record(await f.run(["--execute", "--confirm", "a", "--full"], "soc", "me", "a"));
  assert.equal(records(full.results)[0]!.newId, "x".repeat(501));
});

test("move returned ID retains redaction even with full", async t => {
  const f = fixture(t, transfers[0]!, { write: () => ({ status: 201, headers: {}, body: JSON.stringify({ id: "-----BEGIN PRIVATE KEY-----fixture-----END PRIVATE KEY-----" }) }) });
  const result = record(await f.run(["--execute", "--confirm", "a", "--full"], "soc", "me", "a"));
  assert.ok(!JSON.stringify(result).includes("PRIVATE KEY"));
  assert.equal(records(result.results)[0]!.status, "success");
});

for (const spec of actions) {
  test(`${spec.action} previews only reviewed state without writes or audit`, async t => {
    const f = fixture(t, spec);
    const result = record(decode(encode(record(await f.run()))));
    assert.deepEqual(records(result.results).map(row => row.status), ["dry-run", "dry-run"]);
    assert.deepEqual(JSON.parse(text(record(records(result.results)[0]!.preview).proposedChange)), { [spec.field]: spec.desired });
    assert.equal(f.writes.length, 0);
    assert.equal(f.journal().length, 0);
    assert.equal(f.reads.length, 2);
    assert.ok(!JSON.stringify(result).includes("unreviewed-"));
  });
  for (const [profile, target] of [["soc", "me"], ["soc", mailbox], ["batch", mailbox]] as const) test(`${spec.action} writes sequentially and audits in ${profile}/${target}`, async t => {
    const f = fixture(t, spec);
    const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="], profile, target));
    assert.deepEqual(records(result.results).map(row => row.status), ["success", "success"]);
    assert.deepEqual(f.events, ["read", "read", "write", "read", "read", "write"]);
    assert.deepEqual(f.journal().map(row => row.kind), ["intent", "outcome", "intent", "outcome"]);
    assert.deepEqual(f.journal().filter(row => row.kind === "outcome").map(row => row.outcome), ["SUCCESS", "SUCCESS"]);
    assert.equal(new Set(f.journal().map(row => row.id)).size, 2);
    for (const request of f.writes) {
      assert.equal(request.method, "PATCH");
      assert.deepEqual(JSON.parse(text(request.body)), { [spec.field]: spec.desired });
      assert.match(request.url, target === "me" ? /\/me\/messages\/[AB]%2BID%3D$/ : /\/users\/mailbox%40contoso.example\/messages\/[AB]%2BID%3D$/);
    }
    assert.deepEqual(f.credentials.filter(item => item.scopes?.includes("https://graph.microsoft.com/Mail.ReadWrite")).length, profile === "soc" ? 2 : 0);
    assert.ok(!JSON.stringify(result).includes("unreviewed-"));
    assert.ok(!JSON.stringify(f.journal()).includes(`"${spec.field}":`));
  });
  test(`${spec.action} desired state is a no-op`, async t => {
    const f = fixture(t, spec, { read: id => ({ id, [spec.field]: spec.desired }) });
    const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="]));
    assert.deepEqual(records(result.results).map(row => row.status), ["noop", "noop"]);
    assert.equal(f.writes.length, 0);
    assert.equal(f.journal().length, 0);
  });
  for (const [args, ids, code] of [
    [["--execute"], "a,b", "CONFIRM_REQUIRED"],
    [["--execute", "--confirm", "b,a"], "a,b", "CONFIRM_MISMATCH"],
    [["--execute", "--confirm", "a"], "a,b", "CONFIRM_MISMATCH"],
    [[], "a,a", "VALIDATION_ERROR"], [[], "", "VALIDATION_ERROR"], [[], "a,", "VALIDATION_ERROR"],
    [[], "a/../b", "VALIDATION_ERROR"], [[], Array.from({ length: 21 }, (_, index) => `id${index}`).join(","), "VALIDATION_ERROR"],
    [["--api-version", "beta"], "a", "VALIDATION_ERROR"], [["--scopes", "Mail.ReadWrite"], "a", "VALIDATION_ERROR"],
  ] as const) test(`${spec.action} refuses invalid selection or confirmation before calls (${code}/${args.join(" ")}/${ids})`, async t => {
    const f = fixture(t, spec);
    await assert.rejects(f.run(args, "soc", "me", ids), { code });
    assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
  });
}

for (const [name, change, code] of [
  ["pack", (profile: ProfileConfig["profiles"][string]) => { profile.enabledPacks = ["entra"]; }, "POLICY_DENIED"],
  ["writes", (profile: ProfileConfig["profiles"][string]) => { profile.writes.allowWrites = false; }, "WRITES_DISABLED"],
  ["operation", (profile: ProfileConfig["profiles"][string]) => { profile.writes.operations = ["mail.message.move"]; }, "OPERATION_NOT_WRITABLE"],
] as const) test(`mail write ${name} gate refuses before calls`, async t => {
  const f = fixture(t, actions[0]!);
  change(f.configuration.profiles.soc!); writeFileSync(f.config, JSON.stringify(f.configuration));
  await assert.rejects(f.run(), { code });
  assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
});

for (const status of [400, 403, 429, 503]) test(`mail batch stops on ${status} without retrying or leaking errors`, async t => {
  const f = fixture(t, actions[2]!, { write: () => ({ status, headers: { "retry-after": "0" }, body: '{"error":{"message":"private-echo"}}' }) });
  const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="]));
  assert.deepEqual(records(result.results).map(row => row.status), [status === 503 ? "unknown" : "failed", "not-attempted"]);
  assert.equal(result.complete, false); assert.equal(process.exitCode, 1);
  assert.equal(f.writes.length, 1); assert.equal(f.reads.length, 2);
  assert.ok(!JSON.stringify(result).includes("private-echo"));
});

test("mail timeout remains unknown and stops before the next ID", async t => {
  const f = fixture(t, actions[2]!, { write: () => { throw new Error("fixture timeout"); } });
  const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="]));
  assert.deepEqual(records(result.results).map(row => row.status), ["unknown", "not-attempted"]);
  assert.deepEqual(f.journal().map(row => row.outcome).filter(Boolean), ["OUTCOME_UNKNOWN"]);
  assert.equal(f.writes.length, 1);
});

test("outcome storage failure after send is unknown and stops the batch", async t => {
  const f = fixture(t, actions[2]!, { write: () => {
    rmSync(f.journalPath); mkdirSync(f.journalPath);
    return { status: 200, headers: {}, body: "{}" };
  } });
  const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="]));
  assert.deepEqual(records(result.results).map(row => row.status), ["unknown", "not-attempted"]);
  assert.equal(f.writes.length, 1);
});

test("tag aborts on concurrent category edits without overwriting them", async t => {
  const f = fixture(t, actions[0]!, { read: (id, count) => ({ id, categories: count === 1 ? ["Keep"] : ["Keep", "Concurrent"] }) });
  const result = record(await f.run(["--execute", "--confirm", "A+ID=,B+ID="]));
  assert.deepEqual(records(result.results).map(row => row.status), ["failed", "not-attempted"]);
  assert.equal(records(result.results)[0]!.code, "STATE_CHANGED"); assert.equal(f.writes.length, 0);
  assert.equal(f.journal().at(-1)!.outcome, "NOT_SENT");
});

test("mark skips sending if a fresh read already has the desired state", async t => {
  const f = fixture(t, actions[2]!, { read: (id, count) => ({ id, isRead: count > 1 }) });
  const result = record(await f.run(["--execute", "--confirm", "a"], "soc", "me", "a"));
  assert.equal(records(result.results)[0]!.status, "noop"); assert.equal(f.writes.length, 0);
  assert.equal(f.journal().at(-1)!.outcome, "NOT_SENT");
});

test("explicit mailbox identity is checked again on write credential acquisition", async t => {
  const f = fixture(t, actions[2]!, { username: count => count > 2 ? "other@contoso.example" : mailbox });
  const result = record(await f.run(["--execute", "--confirm", "a"], "soc", mailbox, "a"));
  assert.equal(records(result.results)[0]!.code, "POLICY_DENIED"); assert.equal(f.writes.length, 0);
  assert.equal(f.journal().at(-1)!.outcome, "NOT_SENT");
});

test("mail selection accepts the 20-ID boundary", async t => {
  const f = fixture(t, actions[2]!);
  const ids = Array.from({ length: 20 }, (_, index) => `id${index}`).join(",");
  assert.equal(records(record(await f.run([], "soc", "me", ids)).results).length, 20);
  assert.equal(f.reads.length, 20);
});

test("tag preview truncates at 500 characters and full preserves redaction", async t => {
  const f = fixture(t, actions[0]!, { read: id => ({ id, categories: ["x".repeat(501), "-----BEGIN PRIVATE KEY-----fixture-----END PRIVATE KEY-----"] }) });
  const preview = record(await f.run([], "soc", "me", "a"));
  assert.match(text(record(records(preview.results)[0]!.preview).currentState), /truncated/); assert.match(strings(preview.help).join(" "), /--full/);
  const full = record(await f.run(["--full"], "soc", "me", "a"));
  assert.ok(text(record(records(full.results)[0]!.preview).currentState).includes("x".repeat(501)));
  assert.ok(!JSON.stringify(full).includes("PRIVATE KEY"));
});

test("tag refuses to write category names redacted by the read boundary", async t => {
  const f = fixture(t, actions[0]!, { read: id => ({ id, categories: ["-----BEGIN PRIVATE KEY-----fixture-----END PRIVATE KEY-----"] }) });
  const result = record(await f.run(["--execute", "--confirm", "a"], "soc", "me", "a"));
  assert.equal(records(result.results)[0]!.code, "POLICY_DENIED"); assert.equal(f.writes.length, 0);
});

test("forced read-only refuses all mail write requests before credentials", async t => {
  const f = fixture(t, actions[0]!);
  const previous = process.env.MG_AXI_READ_ONLY;
  process.env.MG_AXI_READ_ONLY = "1";
  t.after(() => { if (previous === undefined) delete process.env.MG_AXI_READ_ONLY; else process.env.MG_AXI_READ_ONLY = previous; });
  await assert.rejects(f.run(), { code: "WRITES_DISABLED" });
  assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
});

test("application me and delegated other-mailbox writes are refused", async t => {
  const f = fixture(t, actions[2]!);
  await assert.rejects(f.run([], "batch", "me"), { code: "POLICY_DENIED" });
  assert.equal(f.credentials.length, 0);
  const result = record(await f.run([], "soc", "other@contoso.example"));
  assert.equal(records(result.results)[0]!.code, "POLICY_DENIED"); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
});
