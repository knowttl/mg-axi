import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync, existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { decode, encode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const mailbox = "mailbox@contoso.example";
const rule = { displayName: "Reviewed", sequence: 1, isEnabled: true, conditions: { subjectContains: ["Review"] }, actions: { markAsRead: true } };
const response = body => ({ status: 200, headers: {}, body: JSON.stringify(body) });
function fixture(t, action, options = {}) {
  const directory = mkdtempSync(join(tmpdir(), "mg-axi-rule-writes-"));
  const previousExitCode = process.exitCode;
  t.after(() => { process.exitCode = previousExitCode; rmSync(directory, { recursive: true, force: true }); });
  const config = join(directory, "config.json"), journalPath = join(directory, "writes.log"), file = join(directory, "rule.json");
  writeFileSync(file, JSON.stringify(options.payload ?? (action === "create" ? rule : { isEnabled: false })));
  const profiles = new Profiles(config);
  const tenant = "11111111-1111-4111-8111-111111111111", client = "22222222-2222-4222-8222-222222222222";
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  const configuration = JSON.parse(readFileSync(config, "utf8"));
  for (const profile of Object.values(configuration.profiles)) {
    profile.enabledPacks = ["mail"];
    profile.writes = { allowWrites: true, operations: ["mail.rule.create", "mail.rule.update", "mail.rule.delete"] };
  }
  writeFileSync(config, JSON.stringify(configuration));
  const reads = [], writes = [], credentials = [];
  const auth = { credential: async (profile, scopes) => {
    credentials.push({ profile, scopes });
    return { token: "opaque-fixture-token", tenantId: tenant, clientId: client, accountId: "fixture", username: mailbox };
  } };
  const confirm = action === "create" ? "Reviewed" : "Rule+ID=,Other+ID=";
  const run = (extra = [], profile = "soc", target = "me", ids = "Rule+ID=,Other+ID=") => executeArgv([
    "mail", "rule", action, "--mailbox", target, ...(action === "create" ? [] : ["--ids", ids]),
    ...(action === "delete" ? [] : ["--file", file]), "--profile", profile, "--config", config, ...extra,
  ], { delegated: auth, application: auth, journalPath,
    transport: async request => {
      reads.push(request);
      const id = decodeURIComponent(new URL(request.url).pathname.split("/").at(-1));
      return options.read?.(id, reads.length) ?? response(action === "create" ? { value: [] } : { id, isReadOnly: false, actions: { markAsRead: true }, body: "unreviewed-state" });
    }, mutationTransport: async request => {
      writes.push(request);
      assert.equal(JSON.parse(readFileSync(journalPath, "utf8").trim().split("\n").at(-1)).kind, "intent");
      return options.write?.() ?? { status: action === "delete" ? 204 : action === "create" ? 201 : 200, headers: {}, body: action === "delete" ? "" : '{"id":"new-id","body":"unreviewed-response"}' };
    },
  });
  const journal = () => existsSync(journalPath) ? readFileSync(journalPath, "utf8").trim().split("\n").map(line => JSON.parse(line)) : [];
  return { run, reads, writes, credentials, journal, config, configuration, file, confirm };
}

for (const action of ["create", "update", "delete"]) {
  test(`rule ${action} previews without writes or audit`, async t => {
    const f = fixture(t, action);
    const result = decode(encode(await f.run()));
    assert.ok(result.results.every(row => row.status === "dry-run"));
    assert.equal(f.writes.length, 0); assert.equal(f.journal().length, 0);
    assert.ok(!JSON.stringify(result).includes("unreviewed-"));
  });
  for (const [profile, target] of [["soc", "me"], ["soc", mailbox], ["batch", mailbox]]) test(`rule ${action} executes with ${profile}/${target}`, async t => {
    const f = fixture(t, action);
    const result = await f.run(["--execute", "--confirm", f.confirm], profile, target);
    const count = action === "create" ? 1 : 2;
    assert.deepEqual(result.results.map(row => row.status), Array(count).fill("success"));
    assert.equal(f.writes.length, count); assert.equal(f.reads.length, count * 2);
    assert.deepEqual(f.journal().map(row => row.kind), Array.from({ length: count }, () => ["intent", "outcome"]).flat());
    for (const request of f.writes) {
      assert.equal(request.method, action === "create" ? "POST" : action === "update" ? "PATCH" : "DELETE");
      assert.ok(request.url.startsWith(`https://graph.microsoft.com/v1.0/${target === "me" ? "me" : "users/mailbox%40contoso.example"}/mailFolders/inbox/messageRules`));
      assert.deepEqual(request.body === undefined ? undefined : JSON.parse(request.body), action === "delete" ? undefined : action === "create" ? rule : { isEnabled: false });
    }
    if (profile === "soc") {
      assert.equal(f.credentials.filter(row => row.scopes?.[0] === "https://graph.microsoft.com/MailboxSettings.ReadWrite").length, count);
      assert.ok(f.credentials.every(row => ["https://graph.microsoft.com/MailboxSettings.Read", "https://graph.microsoft.com/MailboxSettings.ReadWrite"].includes(row.scopes?.[0])));
    } else assert.ok(f.credentials.every(row => row.scopes === undefined));
    assert.ok(!JSON.stringify(result).includes("unreviewed-"));
    assert.ok(!JSON.stringify(f.journal()).includes("subjectContains"));
  });
  for (const [extra, code] of [[["--execute"], "CONFIRM_REQUIRED"], [["--execute", "--confirm", "wrong"], "CONFIRM_MISMATCH"], [["--api-version", "beta"], "VALIDATION_ERROR"]]) test(`rule ${action} refuses ${code} before credentials`, async t => {
    const f = fixture(t, action);
    await assert.rejects(f.run(extra), { code });
    assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
  });
  for (const [gate, change, code] of [
    ["pack", p => { p.enabledPacks = ["entra"]; }, "POLICY_DENIED"],
    ["writes", p => { p.writes.allowWrites = false; }, "WRITES_DISABLED"],
    ["operation", p => { p.writes.operations = ["mail.message.mark"]; }, "OPERATION_NOT_WRITABLE"],
  ]) test(`rule ${action} enforces ${gate} before credentials`, async t => {
    const f = fixture(t, action);
    change(f.configuration.profiles.soc); writeFileSync(f.config, JSON.stringify(f.configuration));
    await assert.rejects(f.run(), { code });
    assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
  });
  for (const status of [403, 429, 503]) test(`rule ${action} sends once on ${status} and stops the batch`, async t => {
    const f = fixture(t, action, { write: () => ({ status, headers: {}, body: '{"error":{"message":"private-echo"}}' }) });
    const result = await f.run(["--execute", "--confirm", f.confirm]);
    assert.equal(result.results[0].status, status === 503 ? "unknown" : "failed");
    assert.equal(f.writes.length, 1); assert.equal(process.exitCode, 1);
    assert.ok(!JSON.stringify(result).includes("private-echo"));
    if (action !== "create") assert.equal(result.results[1].status, "not-attempted");
  });
  test(`rule ${action} restricts mailbox identity`, async t => {
    const f = fixture(t, action);
    await assert.rejects(f.run([], "batch", "me"), { code: "POLICY_DENIED" });
    const result = await f.run([], "soc", "other@contoso.example");
    assert.equal(result.results[0].code, "POLICY_DENIED");
    assert.equal(f.reads.length, 0); assert.equal(f.writes.length, 0);
  });
}
for (const payload of [[], {}, { id: "bad" }, { sequence: -1 }, { isEnabled: "true" }, { conditions: { unknown: true } }, { conditions: { withinSizeRange: { minimumSize: 10, maximumSize: 1 } } }, { conditions: { fromAddresses: [{ emailAddress: { address: "x", secret: "bad" } }] } }, { actions: { markAsRead: "true" } }, ...["forwardTo", "forwardAsAttachmentTo", "redirectTo", "permanentDelete"].map(key => ({ actions: { [key]: false } }))]) {
  test(`rule file rejects invalid shape ${JSON.stringify(payload)}`, async t => {
    const f = fixture(t, "update", { payload });
    await assert.rejects(f.run(), { code: "VALIDATION_ERROR" });
    assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0);
  });
}
for (const content of ["{bad-json-private", " ".repeat(65537), Buffer.from([0xff])]) test("rule file rejects malformed, oversized or invalid UTF-8 input without echo", async t => {
  const f = fixture(t, "create"); writeFileSync(f.file, content);
  await assert.rejects(f.run(), error => error.code === "VALIDATION_ERROR" && !error.message.includes("bad-json-private"));
  assert.equal(f.credentials.length, 0);
});
test("create requires a complete rule", async t => {
  const f = fixture(t, "create", { payload: { displayName: "Reviewed" } });
  await assert.rejects(f.run(), { code: "VALIDATION_ERROR" }); assert.equal(f.credentials.length, 0);
});
for (const action of ["update", "delete"]) {
  test(`rule ${action} refuses read-only rules`, async t => {
    const f = fixture(t, action, { read: id => response({ id, isReadOnly: true, actions: {} }) });
    const result = await f.run(["--execute", "--confirm", f.confirm]);
    assert.equal(result.results[0].code, "POLICY_DENIED"); assert.equal(f.writes.length, 0);
  });
  for (const ids of ["a,a", "a/../b", Array.from({ length: 21 }, (_, i) => `id${i}`).join(",")]) test(`rule ${action} validates target selection`, async t => {
    const f = fixture(t, action);
    await assert.rejects(f.run([], "soc", "me", ids), { code: "VALIDATION_ERROR" }); assert.equal(f.credentials.length, 0);
  });
}
for (const actions of [{ forwardTo: [{ emailAddress: { address: "external@contoso.example" } }] }, { permanentDelete: true }]) test("update refuses unsafe actions added before the fresh read", async t => {
  const f = fixture(t, "update", { read: (id, count) => response({ id, isReadOnly: false, actions: count === 1 ? {} : actions }) });
  const result = await f.run(["--execute", "--confirm", f.confirm]);
  assert.equal(result.results[0].code, "POLICY_DENIED"); assert.equal(f.writes.length, 0);
  assert.equal(f.journal().at(-1).outcome, "NOT_SENT");
});
test("delete can remove a rule containing prohibited actions", async t => {
  const f = fixture(t, "delete", { read: id => response({ id, isReadOnly: false, actions: { permanentDelete: true } }) });
  assert.ok((await f.run(["--execute", "--confirm", f.confirm])).results.every(row => row.status === "success"));
});
test("create reports unknown when Graph omits the new ID", async t => {
  const f = fixture(t, "create", { write: () => response({}) });
  assert.equal((await f.run(["--execute", "--confirm", f.confirm])).results[0].status, "unknown"); assert.equal(f.writes.length, 1);
});
test("rule previews truncate with a full escape hatch", async t => {
  const f = fixture(t, "update", { payload: { conditions: { subjectContains: ["x".repeat(501)] } } });
  assert.match((await f.run()).results[0].preview.proposedChange, /truncated/);
  assert.deepEqual(JSON.parse((await f.run(["--full"])).results[0].preview.proposedChange), { conditions: { subjectContains: ["x".repeat(501)] } });
});
test("rule output redacts protected values even with full output", async t => {
  const secret = "-----BEGIN PRIVATE KEY-----fixture-----END PRIVATE KEY-----";
  const f = fixture(t, "create", { payload: { ...rule, displayName: secret, conditions: { subjectContains: [secret] } } });
  const result = await f.run(["--full"]);
  assert.ok(!JSON.stringify(result).includes(secret));
  assert.match(result.results[0].id, /redacted/);
});
test("rule writes honor forced read-only before credentials", async t => {
  const previous = process.env.MG_AXI_READ_ONLY;
  t.after(() => { if (previous === undefined) delete process.env.MG_AXI_READ_ONLY; else process.env.MG_AXI_READ_ONLY = previous; });
  process.env.MG_AXI_READ_ONLY = "1";
  const f = fixture(t, "delete");
  await assert.rejects(f.run(), { code: "WRITES_DISABLED" });
  assert.equal(f.credentials.length, 0); assert.equal(f.reads.length, 0);
});
