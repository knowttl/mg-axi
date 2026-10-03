import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, cpSync, mkdirSync, readFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { performance } from "node:perf_hooks";
import { decode } from "@toon-format/toon";

const bin = resolve("dist/bin/mg-axi.js");
function run(args, executable = bin) {
  return spawnSync(process.execPath, [executable, ...args], {
    encoding: "utf8", input: "", timeout: 10000,
    env: { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot },
  });
}

test("home reports unavailable tenant state on stdout without credentials", () => {
  const result = run([]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  const output = decode(result.stdout);
  assert.equal(output.bin.replace(/^~(?=[/\\])/, homedir()), bin);
  assert.match(output.tenant, /unavailable/);
  assert.equal(output.domains[0].status, "scheduled");
});

test("top help lists only the shell catalogue", () => {
  const result = run(["--help"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /entra user list/);
  assert.match(result.stdout, /login/);
  assert.doesNotMatch(result.stdout, /Upgrade/);
});

test("leaf help includes flags, defaults and examples without requiring an ID", () => {
  const result = run(["entra", "user", "show", "--help"]);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /--id.*required/);
  assert.match(result.stdout, /default: v1.0/);
  assert.doesNotMatch(result.stdout, /--limit/);
});

for (const [name, args, error] of [
  ["unknown flag", ["entra", "user", "list", "--limt", "2"], /unknown flag --limt/],
  ["unknown flag alongside help", ["entra", "user", "list", "--help", "--typo"], /unknown flag/],
  ["incompatible row controls", ["entra", "user", "list", "--all", "--limit", "2"], /cannot be combined/],
  ["invalid limit", ["entra", "user", "list", "--limit", "0"], /positive safe integer/],
  ["invalid version", ["entra", "user", "list", "--api-version", "v2"], /v1.0 or beta/],
  ["missing required ID", ["entra", "user", "show"], /--id is required/],
  ["empty ID", ["entra", "user", "show", "--id="], /non-empty/],
  ["extra positional", ["entra", "user", "show", "--id", "fixture", "extra"], /unexpected/],
  ["duplicate flag", ["entra", "user", "list", "--all", "--all"], /duplicate/],
  ["boolean value", ["entra", "user", "list", "--all=false"], /does not take a value/],
  ["literal help after delimiter", ["home", "--", "--help"], /unexpected/],
  ["unpublished command", ["api", "GET", "/users"], /unknown/],
  ["prototype name", ["constructor"], /unknown/],
  ["trailing version", ["home", "--version"], /unknown flag/],
]) test(`${name} fails with usage output`, () => {
  const result = run(args);
  assert.equal(result.status, 2);
  assert.equal(result.stderr, "");
  const output = decode(result.stdout);
  assert.match(output.error, error);
  assert.equal(output.code, "VALIDATION_ERROR");
  assert.ok(output.help.length);
});

for (const version of ["v1.0", "beta"]) test(`valid ${version} leaf reports its INV-01 owner without execution`, () => {
  const result = run(["entra", "user", "show", "--id", "fixture", "--api-version", version]);
  assert.equal(result.status, 1);
  assert.equal(result.stderr, "");
  assert.match(decode(result.stdout).error, /scheduled \(READ-01\)/);
});

test("version probes work with the command graph absent", () => {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-version-"));
  try {
    mkdirSync(join(dir, "dist", "bin"), { recursive: true });
    cpSync("dist/bin/mg-axi.js", join(dir, "dist/bin/mg-axi.js"));
    cpSync("dist/version.js", join(dir, "dist/version.js"));
    cpSync("package.json", join(dir, "package.json"));
    cpSync("node_modules", join(dir, "node_modules"), { recursive: true, dereference: true });
    for (const flag of ["-v", "-V", "--version"]) {
      const result = run([flag], join(dir, "dist/bin/mg-axi.js"));
      assert.equal(result.status, 0);
      assert.equal(result.stderr, "");
      assert.equal(result.stdout, `${JSON.parse(readFileSync("package.json", "utf8")).version}\n`);
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("version latency stays near the Node startup floor", () => {
  const measure = args => {
    const start = performance.now();
    const result = spawnSync(process.execPath, args, { encoding: "utf8", input: "", timeout: 10000 });
    assert.equal(result.status, 0);
    return performance.now() - start;
  };
  const floor = measure(["-e", "console.log(1)"]);
  const version = measure([bin, "--version"]);
  assert.ok(version < floor * 6, `version ${version}ms exceeds Node floor ${floor}ms by more than 6x`);
});
