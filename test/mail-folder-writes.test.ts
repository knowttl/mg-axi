import type { MutationTransportRequest, MutationTransportResponse } from "../dist/mutations.js";
import { record, records, text, authFakes, type CredentialStub, type CredentialArgs, type ProfileConfig } from "./fixtures/typed.ts";
import type { TransportRequest, TransportResponse } from "../dist/graph-session.js";
import type { TestContext } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync, existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { decode, encode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const mailbox = "mailbox@contoso.example";
const commands = [
  { noun: "folder", action: "create", flags: ["--name", "Reviewed"], confirm: "Reviewed", scope: "Mail.ReadWrite", readScope: "Mail.ReadBasic", payload: { displayName: "Reviewed" } },
  { noun: "folder", action: "delete", flags: ["--ids", "Custom+ID=,Other+ID="], confirm: "Custom+ID=,Other+ID=", scope: "Mail.ReadWrite", readScope: "Mail.ReadBasic" },
  { noun: "category", action: "create", flags: ["--name", "Reviewed", "--color", "preset4"], confirm: "Reviewed", scope: "MailboxSettings.ReadWrite", readScope: "MailboxSettings.Read", payload: { displayName: "Reviewed", color: "preset4" } },
];
function fixture(t: TestContext, spec: (typeof commands)[number], options: {
  read?: (id: string, count: number, request: TransportRequest) => TransportResponse | undefined;
  write?: (request: MutationTransportRequest, count: number) => MutationTransportResponse;
} = {}) {
  const directory = mkdtempSync(join(tmpdir(), "mg-axi-folder-writes-"));
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
    profile.writes = { allowWrites: true, operations: commands.map(item => `mail.${item.noun}.${item.action}`) };
  }
  writeFileSync(config, JSON.stringify(configuration));
  const reads: TransportRequest[] = [], writes: MutationTransportRequest[] = [], credentials: { profile: CredentialArgs[0]; scopes: CredentialArgs[1] }[] = [];
  const auth: CredentialStub = { credential: async (profile, scopes) => {
    credentials.push({ profile, scopes });
    return { token: "opaque-fixture-token", tenantId: tenant, clientId: client, accountId: "fixture", username: mailbox };
  } };
  const run = (extra: readonly string[] = [], profile = "soc", target = "me", flags: readonly string[] = spec.flags) => executeArgv([
    "mail", spec.noun, spec.action, "--mailbox", target, ...flags, "--profile", profile, "--config", config, ...extra,
  ], { delegated: authFakes(auth).delegated, application: authFakes(auth).application, journalPath,
    transport: async request => {
      reads.push(request);
      const url = new URL(request.url), id = decodeURIComponent(url.pathname.split("/").at(-1)!);
      const body = id === "masterCategories" ? { value: [] } : { id: id.endsWith("ID=") ? id : `protected-${id}`, totalItemCount: 0, childFolderCount: 0, body: "unreviewed-state" };
      return options.read?.(id, reads.length, request) ?? { status: 200, headers: {}, body: JSON.stringify(body) };
    }, mutationTransport: async request => {
      writes.push(request);
      assert.equal(JSON.parse(text(readFileSync(journalPath, "utf8").trim().split("\n").at(-1))).kind, "intent");
      return options.write?.(request, writes.length) ?? { status: spec.action === "delete" ? 204 : 201, headers: {}, body: spec.action === "delete" ? "" : '{"id":"new-id","body":"unreviewed-response"}' };
    },
  });
  const journal = () => existsSync(journalPath) ? readFileSync(journalPath, "utf8").trim().split("\n").map(line => record(JSON.parse(line))) : [];
  return { run, reads, writes, credentials, journal, config, configuration };
}
const response = (body: unknown) => ({ status: 200, headers: {}, body: JSON.stringify(body) });

for (const spec of commands) {
  const label = `${spec.noun} ${spec.action}`;
  test(`${label} previews reviewed state without sending or auditing`, async t => {
    const f = fixture(t, spec);
    const result = record(decode(encode(record(await f.run()))));
    assert.ok(records(result.results).every(row => row.status === "dry-run"));
    assert.equal(f.writes.length, 0); assert.equal(f.journal().length, 0);
    assert.ok(!JSON.stringify(result).includes("unreviewed-"));
    if (spec.payload) assert.deepEqual(JSON.parse(text(record(records(result.results)[0]!.preview).proposedChange)), spec.payload);
  });
  for (const [profile, target] of [["soc", "me"], ["soc", mailbox], ["batch", mailbox]] as const) test(`${label} executes and audits in ${profile}/${target}`, async t => {
    const f = fixture(t, spec);
    const result = record(await f.run(["--execute", "--confirm", spec.confirm], profile, target));
    const count = spec.action === "delete" ? 2 : 1;
    assert.deepEqual(records(result.results).map(row => row.status), Array(count).fill("success"));
    assert.equal(f.writes.length, count);
    assert.deepEqual(f.journal().map(row => row.kind), Array.from({ length: count }, () => ["intent", "outcome"]).flat());
    assert.ok(f.journal().filter(row => row.kind === "outcome").every(row => row.outcome === "SUCCESS"));
    for (const request of f.writes) {
      assert.equal(request.method, spec.action === "delete" ? "DELETE" : "POST");
      assert.equal(request.body === undefined ? undefined : JSON.parse(request.body).displayName, spec.payload?.displayName);
      if (spec.payload) assert.deepEqual(JSON.parse(text(request.body)), spec.payload);
      assert.ok(request.url.startsWith(`https://graph.microsoft.com/v1.0/${target === "me" ? "me" : "users/mailbox%40contoso.example"}/`));
      assert.ok(request.url.endsWith(spec.noun === "category" ? "/outlook/masterCategories" : spec.action === "create" ? "/mailFolders" : "ID%3D"));
    }
    assert.equal(f.credentials.filter(item => item.scopes?.includes(`https://graph.microsoft.com/${spec.scope}`)).length, profile === "soc" ? count : 0);
    if (profile === "soc") assert.ok(f.credentials.filter(item => !item.scopes?.includes(`https://graph.microsoft.com/${spec.scope}`)).every(item => item.scopes?.[0] === `https://graph.microsoft.com/${spec.readScope}`));
    assert.ok(!JSON.stringify(result).includes("unreviewed-"));
    assert.ok(!JSON.stringify(f.journal()).includes("displayName"));
    if (spec.payload) assert.equal(records(result.results)[0]!.newId, "new-id");
  });
  for (const [extra, code] of [[["--execute"], "CONFIRM_REQUIRED"], [["--execute", "--confirm", "wrong"], "CONFIRM_MISMATCH"]] as const) test(`${label} refuses ${code} before credentials`, async t => {
    const f = fixture(t, spec);
    await assert.rejects(f.run(extra), { code });
    assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
  });
  for (const [gate, change, code] of [
    ["pack", (p: ProfileConfig["profiles"][string]) => { p.enabledPacks = ["entra"]; }, "POLICY_DENIED"],
    ["writes", (p: ProfileConfig["profiles"][string]) => { p.writes.allowWrites = false; }, "WRITES_DISABLED"],
    ["operation", (p: ProfileConfig["profiles"][string]) => { p.writes.operations = ["mail.message.mark"]; }, "OPERATION_NOT_WRITABLE"],
  ] as const) test(`${label} enforces ${gate} before credentials`, async t => {
    const f = fixture(t, spec);
    change(f.configuration.profiles.soc!); writeFileSync(f.config, JSON.stringify(f.configuration));
    await assert.rejects(f.run(), { code });
    assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
  });
  for (const status of [403, 429, 503]) test(`${label} sends once on ${status} and suppresses error bodies`, async t => {
    const f = fixture(t, spec, { write: () => ({ status, headers: {}, body: '{"error":{"message":"private-echo"}}' }) });
    const result = record(await f.run(["--execute", "--confirm", spec.confirm]));
    assert.equal(records(result.results)[0]!.status, status === 503 ? "unknown" : "failed");
    assert.equal(f.writes.length, 1); assert.equal(process.exitCode, 1);
    assert.ok(!JSON.stringify(result).includes("private-echo"));
    if (spec.action === "delete") assert.equal(records(result.results)[1]!.status, "not-attempted");
  });
  test(`${label} refuses application me and delegated other mailbox`, async t => {
    const f = fixture(t, spec);
    await assert.rejects(f.run([], "batch", "me"), { code: "POLICY_DENIED" });
    try { const result = record(await f.run([], "soc", "other@contoso.example")); assert.equal(records(result.results)[0]!.code, "POLICY_DENIED"); }
    catch (error) { assert.equal(record(error).code, "POLICY_DENIED"); }
    assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
  });
}

for (const counts of [{ totalItemCount: 1, childFolderCount: 0 }, { totalItemCount: 0, childFolderCount: 1 }, { totalItemCount: -1, childFolderCount: 0 }, { totalItemCount: null, childFolderCount: 0 }, {}]) test(`folder delete refuses unsafe counts ${JSON.stringify(counts)}`, async t => {
  const f = fixture(t, commands[1]!, { read: id => id.endsWith("ID=") ? response({ id, ...counts }) : undefined });
  const result = record(await f.run(["--execute", "--confirm", commands[1]!.confirm]));
  assert.equal(records(result.results)[0]!.status, "failed"); assert.equal(records(result.results)[1]!.status, "not-attempted");
  assert.equal(f.writes.length, 0);
});

test("folder delete detects new children on the fresh read and journals NOT_SENT", async t => {
  const f = fixture(t, commands[1]!, { read: (id, count) => id === "Custom+ID=" ? response({ id, totalItemCount: 0, childFolderCount: count > 18 ? 1 : 0 }) : undefined });
  const result = record(await f.run(["--execute", "--confirm", commands[1]!.confirm]));
  assert.equal(records(result.results)[0]!.code, "POLICY_DENIED"); assert.equal(f.writes.length, 0);
  assert.equal(f.journal().at(-1)!.outcome, "NOT_SENT");
});
for (const alias of ["inbox", "INBOX", "deleteditems", "recoverableitemsdeletions", "msgfolderroot"]) test(`folder delete refuses ${alias} before credentials`, async t => {
  const f = fixture(t, commands[1]!);
  await assert.rejects(f.run([], "soc", "me", ["--ids", alias]), { code: "POLICY_DENIED" });
  assert.equal(f.credentials.length, 0);
});
test("folder delete refuses opaque well-known IDs even when localized", async t => {
  const f = fixture(t, commands[1]!, { read: id => id === "inbox" ? response({ id: "Custom+ID=", displayName: "Posteingang" }) : undefined });
  await assert.rejects(f.run(), { code: "POLICY_DENIED" }); assert.equal(f.writes.length, 0);
});
test("folder delete fails closed when protected IDs cannot be read", async t => {
  const f = fixture(t, commands[1]!, { read: id => id === "clutter" ? { status: 404, headers: {}, body: "private-echo" } : undefined });
  await assert.rejects(f.run(), { code: "GRAPH_ERROR" }); assert.equal(f.writes.length, 0);
});
for (const [flags, code] of [
  [["--ids", "a,a"], "VALIDATION_ERROR"], [["--ids", "../a"], "VALIDATION_ERROR"],
  [["--ids", Array.from({ length: 21 }, (_, i) => `id${i}`).join(",")], "VALIDATION_ERROR"],
] as const) test("folder delete validates the entire selection before requests", async t => {
  const f = fixture(t, commands[1]!);
  await assert.rejects(f.run([], "soc", "me", flags), { code }); assert.equal(f.credentials.length, 0);
});
for (const spec of [commands[0]!, commands[2]!]) {
  test(`${spec!.noun} create rejects invalid names before requests`, async t => {
    const f = fixture(t, spec);
    await assert.rejects(f.run([], "soc", "me", spec!.noun === "category" ? ["--name", " bad", "--color", "preset4"] : ["--name", " bad"]), { code: "VALIDATION_ERROR" });
    assert.equal(f.credentials.length, 0);
  });
  test(`${spec!.noun} create missing returned ID is unknown`, async t => {
    const f = fixture(t, spec, { write: () => ({ status: 201, headers: {}, body: "{}" }) });
    const result = record(await f.run(["--execute", "--confirm", spec!.confirm]));
    assert.equal(records(result.results)[0]!.status, "unknown"); assert.equal(f.writes.length, 1);
  });
  test(`${spec!.noun} create truncates previews and returned IDs with full escape hatch`, async t => {
    const f = fixture(t, spec, { write: () => ({ status: 201, headers: {}, body: JSON.stringify({ id: "x".repeat(501) }) }) });
    assert.match(text(records(record(await f.run(["--execute", "--confirm", spec!.confirm])).results)[0]!.newId), /truncated/);
    assert.equal(records(record(await f.run(["--execute", "--confirm", spec!.confirm, "--full"])).results)[0]!.newId, "x".repeat(501));
  });
}
for (const color of ["preset25", "red", "preset01"]) test(`category create refuses color ${color}`, async t => {
  const f = fixture(t, commands[2]!);
  await assert.rejects(f.run([], "soc", "me", ["--name", "Reviewed", "--color", color]), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentials.length, 0);
});
test("category create existing name and color is a no-op", async t => {
  const f = fixture(t, commands[2]!, { read: () => response({ value: [{ id: "existing", displayName: "Reviewed", color: "preset4" }] }) });
  const result = record(await f.run(["--execute", "--confirm", "Reviewed"]));
  assert.equal(records(result.results)[0]!.status, "noop"); assert.equal(f.writes.length, 0); assert.equal(f.journal().length, 0);
});
test("category create detects a concurrently created category without sending", async t => {
  const f = fixture(t, commands[2]!, { read: (_id, count) => response({ value: count === 1 ? [] : [{ id: "existing", displayName: "Reviewed", color: "preset4" }] }) });
  const result = record(await f.run(["--execute", "--confirm", "Reviewed"]));
  assert.equal(records(result.results)[0]!.status, "noop"); assert.equal(f.writes.length, 0); assert.equal(f.journal().at(-1)!.outcome, "NOT_SENT");
});
test("category create refuses a conflicting existing color", async t => {
  const f = fixture(t, commands[2]!, { read: () => response({ value: [{ id: "existing", displayName: "Reviewed", color: "preset5" }] }) });
  const result = record(await f.run(["--execute", "--confirm", "Reviewed"]));
  assert.equal(records(result.results)[0]!.code, "POLICY_DENIED"); assert.equal(f.writes.length, 0);
});
test("category create refuses incomplete duplicate checks without following nextLink", async t => {
  const f = fixture(t, commands[2]!, { read: () => response({ value: [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/outlook/masterCategories?$skiptoken=next" }) });
  const result = record(await f.run(["--execute", "--confirm", "Reviewed"]));
  assert.equal(records(result.results)[0]!.code, "POLICY_DENIED"); assert.equal(f.reads.length, 1); assert.equal(f.writes.length, 0);
});
