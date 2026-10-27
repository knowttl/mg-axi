import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const commands = [
  ["folder", "list"], ["folder", "show", "--id", "inbox"], ["message", "list"],
  ["message", "show", "--id", "a", "--body", "none"], ["attachment", "list", "--id", "a"],
  ["settings", "show"], ["rule", "list"], ["rule", "show", "--id", "a"], ["category", "list"],
];
function fixture(t) {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-mail-mailbox-"));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const config = join(dir, "config.json");
  const profiles = new Profiles(config);
  const tenant = "11111111-1111-4111-8111-111111111111", client = "22222222-2222-4222-8222-222222222222";
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  const data = JSON.parse(readFileSync(config, "utf8")); data.profiles.batch.enabledPacks = ["mail"];
  writeFileSync(config, JSON.stringify(data));
  const requests = [], credentials = [];
  const run = (command, mailbox) => executeArgv(["mail", ...command, "--mailbox", mailbox, "--profile", "batch", "--config", config], {
    application: { credential: async () => { credentials.push(true); return { token: "opaque-fixture-token", tenantId: tenant, clientId: client }; } },
    transport: async request => { requests.push(request); return { status: 200, headers: {}, body: JSON.stringify(command[1] === "list" ? { value: [] } : { id: "a" }) }; },
  });
  return { run, requests, credentials };
}

for (const command of commands) {
  for (const mailbox of ["o'brien@contoso.example", "name#^~!_-@contoso.example", "user+tag@contoso.example"]) test(`${command.join(" ")} accepts and encodes permitted UPN ${mailbox}`, async t => {
    const f = fixture(t);
    await f.run(command, mailbox);
    assert.equal(f.requests.length, 1);
    const url = new URL(f.requests[0].url);
    assert.equal(url.pathname.split("/")[3], encodeURIComponent(mailbox));
    assert.equal(url.hash, "");
  });
  for (const mailbox of ["name.@contoso.example", "name@-contoso.example", "name@contoso..example", "name/extra@contoso.example", "name%2fextra@contoso.example", "name?query@contoso.example", "name@@contoso.example", "name @contoso.example"]) test(`${command.join(" ")} rejects invalid mailbox without calls: ${mailbox}`, async t => {
    const f = fixture(t);
    await assert.rejects(f.run(command, mailbox), { code: "VALIDATION_ERROR" });
    assert.equal(f.credentials.length, 0); assert.equal(f.requests.length, 0);
  });
}
