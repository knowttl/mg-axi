import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { spawnSync } from "node:child_process";
import { encode, decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";
import { GraphSession, resolveSessionOperation } from "../dist/graph-session.js";
import { HUNTING_OPERATION, HUNTING_SCOPE } from "../dist/security-hunting.js";
import { skillCommandTable } from "../dist/docs.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const query = 'EmailEvents | where Timestamp > ago(1d) | project Timestamp, Subject // end comment';
const schema = [{ name: "Subject", type: "String" }];

function fixture(t, { kql = query, body = { schema, results: [{ Subject: "fixture event" }] }, status = 200, enabled = true } = {}) {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-hunting-"));
  const config = join(dir, "config.json");
  const file = join(dir, "hunt.kql");
  const profiles = new Profiles(config);
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  const value = JSON.parse(readFileSync(config, "utf8"));
  if (enabled) for (const profile of Object.values(value.profiles)) profile.enabledPacks = ["security"];
  writeFileSync(config, JSON.stringify(value));
  writeFileSync(file, kql);
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const requests = [];
  const credentials = [];
  const auth = mode => ({ credential: async (...args) => {
    credentials.push({ mode, args });
    return { token: "synthetic-hunting-token", tenantId: tenant, clientId: client, expiresAt: Date.now() + 3600000 };
  } });
  const overrides = { delegated: auth("delegated"), application: auth("application"), transport: async request => {
    requests.push(request);
    return { status, headers: {}, body: JSON.stringify(body) };
  } };
  const run = (args = [], profile = "soc") => executeArgv(["security", "hunting", "query", "--config", config, "--profile", profile, "--file", file, ...args], overrides);
  return { run, requests, credentials, file, config, profiles, overrides };
}

for (const profile of ["soc", "batch"]) test(`hunting runs the bounded read-only POST using ${profile} fixture credentials`, async t => {
  const f = fixture(t);
  const result = await f.run([], profile);
  assert.equal(result.complete, true);
  assert.equal(result.returned, 1);
  assert.deepEqual(decode(encode(result)).results, [{ Subject: "fixture event" }]);
  const [request] = f.requests;
  assert.equal(request.method, "POST");
  assert.equal(request.url, "https://graph.microsoft.com/v1.0/security/runHuntingQuery");
  assert.equal(request.headers["Content-Type"], "application/json; charset=utf-8");
  const body = JSON.parse(request.body);
  assert.equal(body.Timespan, "P1D");
  assert.match(body.Query, /\n\| take 101$/);
  assert.ok(!body.Query.includes("//"));
  assert.deepEqual(f.credentials[0].args[1], profile === "soc" ? [HUNTING_SCOPE] : undefined);
  assert.equal(f.requests.length, 1);
  assert.ok(!JSON.stringify(result).includes("synthetic-hunting-token"));
});

test("hunting reports overflow without a cursor and bounds full output", async t => {
  const f = fixture(t, { body: { schema, results: [{ Subject: "a" }, { Subject: "b" }] } });
  const result = await f.run(["--limit", "1", "--timespan", "PT6H", "--full"]);
  assert.equal(result.complete, false);
  assert.equal(result.reason, "row-limit");
  assert.equal(result.returned, 1);
  assert.equal(result.cursor, undefined);
  assert.match(result.help.join(" "), /no server cursor/);
  assert.deepEqual(JSON.parse(f.requests[0].body), { Query: `${query.split("//")[0]} \n| take 2`, Timespan: "PT6H" });
});

test("hunting recursively truncates after redaction and full retains redaction", async t => {
  const long = "x".repeat(650);
  const f = fixture(t, { body: { schema, results: [{ Subject: long, nested: { note: long, token: "fixture-secret" }, url: "https://example.test/?sig=fixture-signature" }] } });
  const result = await f.run();
  assert.match(result.results[0].Subject, /truncated, 650 chars total/);
  assert.match(result.results[0].nested.note, /truncated/);
  assert.match(result.help.join(" "), /--full/);
  assert.ok(!JSON.stringify(result).includes("fixture-secret"));
  assert.ok(!JSON.stringify(result).includes("fixture-signature"));
  const full = await f.run(["--full"]);
  assert.equal(full.results[0].Subject, long);
  assert.equal(full.results[0].nested.token, "***redacted***");
  assert.equal(full.results[0].url, "***redacted***");
});

test("hunting gives a definitive empty result", async t => {
  const f = fixture(t, { body: { schema, results: [] } });
  const result = await f.run();
  assert.equal(result.complete, true);
  assert.equal(result.returned, 0);
  assert.match(result.help.join(" "), /0 hunting rows matched/);
});

test("hunting returns all syntax diagnostics before credentials", async t => {
  const f = fixture(t, { kql: "EmailEvents\n| wher Subject != 0 | take 5" });
  await assert.rejects(f.run(), error => {
    assert.equal(error.code, "VALIDATION_ERROR");
    assert.match(error.message, /KQL syntax error at line 2, column 3/);
    assert.deepEqual(error.diagnostics, [
      { line: 2, column: 3, code: "KS176", token: "wher", message: "Query operator expected." },
      { line: 2, column: 8, code: "KS005", token: "Subject", message: "Expected: ;" },
    ]);
    return true;
  });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

test("malformed parentheses report syntax rather than an unsupported stage", async t => {
  const f = fixture(t, { kql: "EmailEvents | where Timestamp > ago(1d" });
  await assert.rejects(f.run(), error => /KQL syntax error/.test(error.message) && error.diagnostics[0].code === "KS005");
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

test("hunting policy errors locate the operator and list allowed operators", async t => {
  const f = fixture(t, { kql: "EmailEvents\n// comment\n| join (DeviceEvents) on DeviceId" });
  await assert.rejects(f.run(), error => {
    assert.match(error.message, /read-only allowlist: where, project/);
    assert.deepEqual(error.diagnostics[0], {
      line: 3, column: 3, code: "READ_ONLY_ALLOWLIST", token: "join",
      message: "Operator is not in the read-only allowlist: where, project, project-away, project-rename, project-reorder, extend, summarize, distinct, count, take, limit, top, sort, order",
    });
    return true;
  });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

test("hunting CLI emits capped TOON diagnostics with usage exit and no query literals", t => {
  const f = fixture(t, { kql: 'EmailEvents | wher Subject == "private-literal" | wher x | wher y | wher z' });
  const result = spawnSync(process.execPath, ["dist/bin/mg-axi.js", "security", "hunting", "query", "--config", f.config, "--profile", "soc", "--file", f.file], { encoding: "utf8" });
  assert.equal(result.status, 2);
  assert.equal(result.stderr, "");
  const output = decode(result.stdout);
  assert.equal(output.code, "VALIDATION_ERROR");
  assert.equal(output.diagnostics.length, 5);
  assert.match(output.help.join(" "), /Showing 5 of/);
  assert.ok(!result.stdout.includes("private-literal"));
  assert.ok(!result.stdout.includes("EmailEvents |"));
});

for (const kind of ["Semantic", "Syntax"]) test(`hunting exposes only ${kind} class and position from a synthetic 400`, async t => {
  const f = fixture(t, { status: 400, body: { error: { message: `${kind} error: 'private-literal' at line 2, column 17: secret-unquoted-tail` } } });
  await assert.rejects(f.run(), error => {
    assert.equal(error.code, "GRAPH_ERROR");
    assert.equal(error.message, `The server rejected the query (${kind.toLowerCase()} error): line 2, column 17`);
    assert.ok(!JSON.stringify(error).includes("private-literal"));
    assert.ok(!JSON.stringify(error).includes("secret-unquoted-tail"));
    return true;
  });
  assert.equal(f.requests.length, 1);
});

for (const message of ["unrecognized body", "Semantic error: 'line 2, column 17'", "Semantic error at line 0, column 1", "Semantic error at line 1, column 0"]) test(`hunting keeps unknown server shapes opaque: ${message}`, async t => {
  const f = fixture(t, { status: 400, body: { error: { message } } });
  await assert.rejects(f.run(), error => /^Graph failed .*\(400\)$/.test(error.message));
});

for (const kql of [
  ".show tables", ".delete table EmailEvents records", "EmailEvents; DeviceEvents", "let x = EmailEvents; x",
  "set query_results_cache_max_age = time(1h); EmailEvents", "EmailEvents | evaluate http_request('https://example.test')",
  "externaldata(x:string)['https://example.test']", "EmailEvents | invoke storedFunction()", "storedFunction",
  "EmailEvents | where customFunction(Subject)", "EmailEvents | where foo.tostring(Subject)",
  "EmailEvents | join (DeviceEvents) on DeviceId", "EmailEvents | where (DeviceEvents | count)",
  "EmailEvents | where Timestamp > ago(1d", "EmailEvents |", "EmailEvents | where", "EmailEvents /* hidden */",
  "EmailEvents | where true /* hidden */", "EmailEvents | where customFunction/* hidden */(Subject)",
  "EmailEvents | where Subject == 'unterminated", "", "EmailEvents\n.drop table EmailEvents",
]) test(`hunting refuses unsupported query before credentials: ${JSON.stringify(kql)}`, async t => {
  const f = fixture(t, { kql });
  await assert.rejects(f.run(), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

for (const kql of [
  "EmailEvents", "// comment\nUrlClickEvents | count",
  'EmailEvents | where Subject contains "; .drop table foo | evaluate x()" | take 5',
  'DeviceProcessEvents | where FileName =~ "powershell.exe" | project Timestamp, FileName | order by Timestamp desc',
  'EmailEvents | where Subject in ("a", "b") | summarize Total=count() by bin(Timestamp, 1h)',
]) test(`hunting accepts a read-only pipeline: ${JSON.stringify(kql)}`, async t => {
  const f = fixture(t, { kql });
  await f.run();
  assert.equal(f.requests.length, 1);
  assert.ok(JSON.parse(f.requests[0].body).Query.endsWith("\n| take 101"));
});

for (const args of [["--limit", "10001"], ["--limit", "0"], ["--timespan", "P31D"], ["--timespan", "PT0S"], ["--timespan", "P"], ["--timespan", "PT"], ["--api-version", "beta"], ["--cursor", "x"], ["--query", "EmailEvents"], ["--scopes", HUNTING_SCOPE]]) {
  test(`hunting refuses unsupported bounds and flags: ${args[0]} ${args[1]}`, async t => {
    const f = fixture(t);
    await assert.rejects(f.run(args), { code: "VALIDATION_ERROR" });
    assert.equal(f.credentials.length, 0);
    assert.equal(f.requests.length, 0);
  });
}

test("hunting requires explicit security pack before credentials", async t => {
  const f = fixture(t, { enabled: false });
  await assert.rejects(f.run(), { code: "POLICY_DENIED" });
  assert.equal(f.credentials.length, 0);
});

test("hunting file errors never expose file contents or underlying errors", async t => {
  const f = fixture(t);
  rmSync(f.file);
  await assert.rejects(f.run(), error => error.code === "VALIDATION_ERROR" && error.message === "Cannot read the KQL file");
  assert.equal(f.credentials.length, 0);
});

for (const status of [302, 400, 401, 403, 429, 503]) test(`hunting does not leak server query echoes, redirect or retry (${status})`, async t => {
  const f = fixture(t, { status, body: { error: { message: "unregistered-secret-query-literal" } } });
  await assert.rejects(f.run(), error => !error.message.includes("unregistered-secret-query-literal"));
  assert.equal(f.requests.length, 1);
});

for (const body of [{}, { schema, results: null }, { schema: [{}], results: [] }, { schema, results: [null] }]) test(`hunting refuses malformed success: ${JSON.stringify(body)}`, async t => {
  const f = fixture(t, { body });
  await assert.rejects(f.run(), { code: "GRAPH_ERROR" });
});

test("raw api still rejects security before credentials", async t => {
  const f = fixture(t);
  await assert.rejects(executeArgv(["api", "get", "/security/runHuntingQuery", "--config", f.config, "--profile", "soc"], f.overrides), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

test("session cannot bypass hunting validation or send arbitrary POST bodies", async t => {
  const f = fixture(t);
  const session = new GraphSession(f.overrides);
  await assert.rejects(session.execute({ profile: f.profiles.resolve("soc").profile, operation: resolveSessionOperation("v1.0", "POST", HUNTING_OPERATION.split(":POST:")[1]),
    hunting: { query: ".show tables", timespan: "P1D", limit: 100 }, scopes: [HUNTING_SCOPE] }), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
  assert.match(skillCommandTable(), /mg-axi security hunting query` \| native \| read/);
});

test("collection paging cannot bypass the hunting request contract", async t => {
  const f = fixture(t);
  const session = new GraphSession(f.overrides);
  await assert.rejects(session.collect({ profile: f.profiles.resolve("soc").profile,
    operation: resolveSessionOperation("v1.0", "POST", "/security/microsoft.graph.security.runHuntingQuery"), scopes: [HUNTING_SCOPE] }), { code: "VALIDATION_ERROR" });
  assert.equal(f.credentials.length, 0);
  assert.equal(f.requests.length, 0);
});

test("hunting rejects an oversized response instead of reporting success", async t => {
  const f = fixture(t, { body: { schema, results: [{ Subject: "x".repeat(5_000_000) }] } });
  await assert.rejects(f.run(), error => error.code === "GRAPH_ERROR" && /5 MB/.test(error.message));
});

test("hunting times out through the shared clock without retrying", async t => {
  const f = fixture(t);
  let now = 0;
  const clock = { now: () => now, sleep: async ms => { now += ms; } };
  const session = new GraphSession(f.overrides);
  await assert.rejects(session.execute({ profile: f.profiles.resolve("soc").profile,
    operation: resolveSessionOperation("v1.0", "POST", "/security/microsoft.graph.security.runHuntingQuery"),
    hunting: { query, timespan: "P1D", limit: 100 }, scopes: [HUNTING_SCOPE], clock }),
  error => error.code === "GRAPH_ERROR" && /30-second deadline/.test(error.message));
  assert.equal(now, 30_000);
  assert.equal(f.requests.length, 0);
});
