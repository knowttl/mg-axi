import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { decode, encode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import type { TransportRequest, TransportResponse } from "../dist/graph-session.js";
import type { TestContext } from "node:test";
import { record, text, strings, authFakes, type CredentialStub } from "./fixtures/typed.ts";

const mailbox = "mailbox@contoso.example";
const cases = [
  { command: ["attachment", "list"], args: ["--id", "A+ID="], noun: "attachments", path: "/messages/A%2BID%3D/attachments", scope: "Mail.Read",
    source: { id: "a", name: "fixture.txt", contentType: "text/plain", size: 12, isInline: false, "@odata.type": "#microsoft.graph.fileAttachment", contentBytes: "private-bytes", item: { body: "private-body" } },
    expected: { id: "a", name: "fixture.txt", contentType: "text/plain", size: 12, isInline: false, "@odata.type": "#microsoft.graph.fileAttachment" } },
  { command: ["settings", "show"], args: [], noun: "settings", path: "/mailboxSettings", scope: "MailboxSettings.Read",
    source: { timeZone: "UTC", language: { locale: "en-US", unreviewed: "private-extra" }, automaticRepliesSetting: { status: "disabled", internalReplyMessage: "fixture reply" } },
    expected: { timeZone: "UTC", language: { locale: "en-US" }, automaticRepliesSetting: { status: "disabled", internalReplyMessage: "fixture reply" } } },
  { command: ["rule", "list"], args: [], noun: "rules", path: "/mailFolders/inbox/messageRules", scope: "MailboxSettings.Read",
    source: { id: "r", displayName: "Fixture rule", sequence: 1, isEnabled: true, actions: { forwardTo: [] } },
    expected: { id: "r", displayName: "Fixture rule", sequence: 1, isEnabled: true } },
  { command: ["rule", "show"], args: ["--id", "A+ID="], noun: "rule", path: "/mailFolders/inbox/messageRules/A%2BID%3D", scope: "MailboxSettings.Read",
    source: { id: "r", conditions: { subjectContains: ["fixture"], unreviewed: "private-extra" }, exceptions: {}, actions: { forwardTo: [{ emailAddress: { address: mailbox, unreviewed: "private-extra" } }], permanentDelete: true, unreviewed: "private-extra" } },
    expected: { id: "r", conditions: { subjectContains: ["fixture"] }, exceptions: {}, actions: { forwardTo: [{ emailAddress: { address: mailbox } }], permanentDelete: true } } },
  { command: ["category", "list"], args: [], noun: "categories", path: "/outlook/masterCategories", scope: "MailboxSettings.Read",
    source: { id: "c", displayName: "Fixture", color: "preset0" }, expected: { id: "c", displayName: "Fixture", color: "preset0" } },
];
function fixture(t: TestContext, spec: (typeof cases)[number], response?: (request: TransportRequest, count: number) => TransportResponse) {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-mail-metadata-"));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const config = join(dir, "config.json");
  const profiles = new Profiles(config);
  const tenant = "11111111-1111-4111-8111-111111111111";
  const client = "22222222-2222-4222-8222-222222222222";
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  const data: { profiles: Record<string, { enabledPacks: string[] }> } = JSON.parse(readFileSync(config, "utf8"));
  for (const profile of Object.values(data.profiles)) profile.enabledPacks = ["mail"];
  writeFileSync(config, JSON.stringify(data));
  const requests: TransportRequest[] = [], credentials: Parameters<CredentialStub["credential"]>[] = [];
  const auth: CredentialStub = { credential: async (...args) => {
    credentials.push(args);
    return { token: "opaque-fixture-token", tenantId: tenant, clientId: client, accountId: "fixture", username: mailbox };
  } };
  const run = (args: string[] = [], profile = "soc", target = mailbox) => executeArgv(["mail", ...spec.command, ...spec.args, "--mailbox", target, "--profile", profile, "--config", config, ...args], {
    ...authFakes(auth), transport: async request => {
      requests.push(request);
      return response ? response(request, requests.length) : { status: 200, headers: {}, body: JSON.stringify(spec.command[1] === "show" ? spec.source : { value: [spec.source] }) };
    },
  });
  return { run, requests, credentials, config };
}

for (const spec of cases) {
  for (const [profile, target] of [["soc", "me"], ["soc", mailbox], ["batch", mailbox]]) test(`${spec.command.join(" ")} routes and projects with ${profile}/${target}`, async t => {
    const f = fixture(t, spec);
    const result = record(decode(encode(record(await f.run([], profile, target)))));
    assert.deepEqual(result[spec.noun], spec.command[1] === "show" ? spec.expected : [spec.expected]);
    assert.equal(new URL(f.requests[0]!.url).pathname, `/v1.0/${target === "me" ? "me" : "users/mailbox%40contoso.example"}${spec.path}`);
    assert.deepEqual(f.credentials[0]![1], profile === "soc" ? [`https://graph.microsoft.com/${spec.scope}`] : undefined);
    assert.equal(f.requests[0]!.method, "GET");
    assert.ok(!JSON.stringify(result).includes("private-"));
    if (spec.command[1] === "list") assert.equal(new URL(f.requests[0]!.url).searchParams.get("$top"), "100");
  });
  for (const args of [["--api-version", "beta"], ["--fields", "contentBytes"], ["--scopes", "Mail.ReadWrite"], ["--all"]]) test(`${spec.command.join(" ")} invalid flags refuse without calls: ${args.join(" ")}`, async t => {
    const f = fixture(t, spec);
    await assert.rejects(f.run(args), { code: "VALIDATION_ERROR" });
    assert.equal(f.credentials.length, 0);
    assert.equal(f.requests.length, 0);
  });
  for (const [profile, target, calls] of [["batch", "me", 0], ["soc", "other@contoso.example", 1]] as const) test(`${spec.command.join(" ")} denies ${profile}/${target}`, async t => {
    const f = fixture(t, spec);
    await assert.rejects(f.run([], profile, target), { code: "POLICY_DENIED" });
    assert.equal(f.credentials.length, calls);
    assert.equal(f.requests.length, 0);
  });
  test(`${spec.command.join(" ")} requires opt-in before credentials`, async t => {
    const f = fixture(t, spec);
    const data = JSON.parse(readFileSync(f.config, "utf8"));
    data.profiles.soc.enabledPacks = ["entra"];
    writeFileSync(f.config, JSON.stringify(data));
    await assert.rejects(f.run(), { code: "POLICY_DENIED" });
    assert.equal(f.credentials.length, 0);
    assert.equal(f.requests.length, 0);
  });
}

for (const spec of cases.filter(spec => spec.command[1] === "list")) test(`${spec.command.join(" ")} preserves bounded cursor and rejects changed fields`, async t => {
  const f = fixture(t, spec, (request, count) => ({ status: 200, headers: {}, body: JSON.stringify({ value: [spec.source], ...(count === 1 ? { "@odata.nextLink": `${request.url}&$skip=17` } : {}) }) }));
  const first = record(await f.run(["--limit", "1"]));
  assert.equal(first.complete, false);
  assert.deepEqual(JSON.parse(Buffer.from(text(first.cursor), "base64url").toString("utf8")).buffered, []);
  await assert.rejects(f.run(["--limit", "1", "--fields", "id", "--cursor", text(first.cursor)]), { code: "VALIDATION_ERROR" });
  assert.equal(f.requests.length, 1);
  assert.equal((record(await f.run(["--limit", "1", "--cursor", text(first.cursor)]))).complete, true);
  assert.equal(new URL(f.requests[1]!.url).searchParams.get("$skip"), "17");
});

test("settings nested text truncates, full retains redaction and signed timezone bias", async t => {
  const spec = cases[1]!;
  const secret = "-----BEGIN PRIVATE KEY-----fixture-----END PRIVATE KEY-----";
  const f = fixture(t, spec, () => ({ status: 200, headers: {}, body: JSON.stringify({ automaticRepliesSetting: { internalReplyMessage: "x".repeat(501), externalReplyMessage: secret }, workingHours: { timeZone: { bias: -60 } } }) }));
  const first = record(await f.run());
  assert.match(text(record(record(first.settings).automaticRepliesSetting).internalReplyMessage), /truncated, 501 chars total/);
  assert.match(strings(first.help).join(" "), /--full/);
  const full = record(await f.run(["--full"]));
  assert.equal(text(record(record(full.settings).automaticRepliesSetting).internalReplyMessage).length, 501);
  assert.equal(record(record(full.settings).automaticRepliesSetting).externalReplyMessage, "***redacted***");
  assert.equal(record(record(record(full.settings).workingHours).timeZone).bias, -60);
});

for (const spec of cases.filter(spec => spec.command[1] === "list")) test(`${spec.command.join(" ")} reports empty results`, async t => {
  const f = fixture(t, spec, () => ({ status: 200, headers: {}, body: '{"value":[]}' }));
  assert.match(strings(record(await f.run()).help).join(" "), new RegExp(`0 ${spec.noun} returned`));
});

test("attachment list refuses oversized pages and malformed metadata", async t => {
  const spec = cases[0]!;
  const oversized = fixture(t, spec, () => ({ status: 200, headers: {}, body: JSON.stringify({ value: [spec.source, spec.source] }) }));
  await assert.rejects(oversized.run(["--limit", "1"]), { code: "GRAPH_ERROR" });
  const malformed = fixture(t, spec, () => ({ status: 200, headers: {}, body: '{"value":[{"size":"12"}]}' }));
  await assert.rejects(malformed.run(), { code: "GRAPH_ERROR" });
});
