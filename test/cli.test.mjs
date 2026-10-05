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
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-cli-"));
  try {
    return spawnSync(process.execPath, [executable, ...args], {
      encoding: "utf8", input: "", timeout: 10000,
      env: { PATH: process.env.PATH, SystemRoot: process.env.SystemRoot, MG_AXI_CONFIG: join(dir, "config.json") },
    });
  } finally { rmSync(dir, { recursive: true, force: true }); }
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
  assert.match(result.stdout, /entra group list/);
  assert.match(result.stdout, /entra application list/);
  assert.match(result.stdout, /entra service-principal list/);
  assert.match(result.stdout, /entra directory-role list/);
  assert.match(result.stdout, /entra user authentication-method list/);
  assert.match(result.stdout, /entra registration list/);
  assert.match(result.stdout, /entra service-principal oauth2-grant list/);
  assert.match(result.stdout, /entra service-principal app-role-assignment list/);
  assert.match(result.stdout, /entra pim active list/);
  assert.match(result.stdout, /entra device list/);
  assert.match(result.stdout, /entra administrative-unit list/);
  assert.match(result.stdout, /entra organization list/);
  assert.match(result.stdout, /entra organization branding show/);
  assert.match(result.stdout, /entra organization branding-localization list/);
  assert.match(result.stdout, /entra domain list/);
  assert.match(result.stdout, /entra domain verification-dns-record list/);
  assert.match(result.stdout, /entra domain service-configuration-record list/);
  assert.match(result.stdout, /entra domain-dns-record list/);
  assert.match(result.stdout, /entra certificate-auth-pki list/);
  assert.match(result.stdout, /entra certificate-authority list/);
  assert.match(result.stdout, /entra subscription list/);
  assert.match(result.stdout, /entra subscription show/);
  assert.match(result.stdout, /entra subscription count/);
  assert.match(result.stdout, /entra on-premises-synchronization list/);
  assert.match(result.stdout, /entra on-premises-synchronization show/);
  assert.match(result.stdout, /entra group-lifecycle-policy list/);
  assert.match(result.stdout, /entra group-lifecycle-policy show/);
  assert.match(result.stdout, /entra group-lifecycle-policy count/);
  assert.match(result.stdout, /entra group-setting-template list/);
  assert.match(result.stdout, /entra group-setting-template show/);
  assert.match(result.stdout, /entra group-setting-template count/);
  assert.match(result.stdout, /entra attribute-set list/);
  assert.match(result.stdout, /entra attribute-set show/);
  assert.match(result.stdout, /entra attribute-set count/);
  assert.match(result.stdout, /entra custom-security-attribute-definition list/);
  assert.match(result.stdout, /entra custom-security-attribute-definition show/);
  assert.match(result.stdout, /entra custom-security-attribute-definition count/);
  assert.match(result.stdout, /entra allowed-value list/);
  assert.match(result.stdout, /entra allowed-value show/);
  assert.match(result.stdout, /entra allowed-value count/);
  assert.match(result.stdout, /entra contract list/);
  assert.match(result.stdout, /entra contract show/);
  assert.match(result.stdout, /entra contract count/);
  assert.match(result.stdout, /entra conditional-access policy list/);
  assert.match(result.stdout, /entra conditional-access named-location list/);
  assert.match(result.stdout, /entra group member add/);
  assert.match(result.stdout, /entra user revoke-sessions/);
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
  ["missing required role ID", ["entra", "directory-role", "show"], /--id is required/],
  ["missing required policy ID", ["entra", "conditional-access", "policy", "show"], /--id is required/],
  ["missing required location ID", ["entra", "conditional-access", "named-location", "show"], /--id is required/],
  ["missing required group for member add", ["entra", "group", "member", "add"], /--group is required/],
  ["missing required user for member add", ["entra", "group", "member", "add", "--group", "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"], /--user is required/],
  ["unknown flag on member add", ["entra", "group", "member", "add", "--transitive"], /unknown flag/],
  ["missing required application", ["entra", "application", "owner", "list"], /--application is required/],
  ["missing required user", ["entra", "user", "authentication-method", "list"], /--user is required/],
  ["missing required organization ID", ["entra", "organization", "show"], /--id is required/],
  ["missing required organization for branding", ["entra", "organization", "branding", "show"], /--organization is required/],
  ["missing required organization for localizations", ["entra", "organization", "branding-localization", "list"], /--organization is required/],
  ["organization list rejects filters", ["entra", "organization", "list", "--filter", "displayName eq 'Contoso'"], /unknown flag --filter/],
  ["missing required domain ID", ["entra", "domain", "show"], /--id is required/],
  ["missing required domain for verification records", ["entra", "domain", "verification-dns-record", "list"], /--domain is required/],
  ["missing required domain for service records", ["entra", "domain", "service-configuration-record", "list"], /--domain is required/],
  ["domain list rejects filters", ["entra", "domain", "list", "--filter", "isVerified eq true"], /unknown flag --filter/],
  ["missing required lifecycle policy ID", ["entra", "group-lifecycle-policy", "show"], /--id is required/],
  ["lifecycle policy count rejects limits", ["entra", "group-lifecycle-policy", "count", "--limit", "5"], /unknown flag --limit/],
  ["missing required setting template ID", ["entra", "group-setting-template", "show"], /--id is required/],
  ["setting template list rejects filters", ["entra", "group-setting-template", "list", "--filter", "displayName eq 'Group.Unified'"], /unknown flag --filter/],
  ["setting template count rejects limits", ["entra", "group-setting-template", "count", "--limit", "5"], /unknown flag --limit/],
  ["missing required contract ID", ["entra", "contract", "show"], /--id is required/],
  ["contract count rejects limits", ["entra", "contract", "count", "--limit", "5"], /unknown flag --limit/],
  ["missing required service-principal", ["entra", "service-principal", "owner", "list"], /--service-principal is required/],
  ["missing required grant client", ["entra", "service-principal", "oauth2-grant", "list"], /--service-principal is required/],
  ["missing required app-role client", ["entra", "service-principal", "app-role-assignment", "list"], /--service-principal is required/],
  ["transitive on application list", ["entra", "application", "list", "--transitive"], /unknown flag/],
  ["missing required administrative unit", ["entra", "administrative-unit", "member", "list"], /--administrative-unit is required/],
  ["transitive on group list", ["entra", "group", "list", "--transitive"], /unknown flag/],
  ["transitive on device list", ["entra", "device", "list", "--transitive"], /unknown flag/],
  ["transitive on policy list", ["entra", "conditional-access", "policy", "list", "--transitive"], /unknown flag/],
  ["empty ID", ["entra", "user", "show", "--id="], /non-empty/],
  ["extra positional", ["entra", "user", "show", "--id", "fixture", "extra"], /unexpected/],
  ["duplicate flag", ["entra", "user", "list", "--all", "--all"], /duplicate/],
  ["boolean value", ["entra", "user", "list", "--all=false"], /does not take a value/],
  ["literal help after delimiter", ["home", "--", "--help"], /unexpected/],
  ["non-GET api verb", ["api", "delete", "/users"], /reviewed GET reads only/],
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

for (const version of ["v1.0", "beta"]) test(`valid ${version} leaf resolves its profile before execution`, () => {
  const result = run(["entra", "user", "show", "--id", "fixture", "--api-version", version]);
  assert.equal(result.status, 1);
  assert.equal(result.stderr, "");
  const output = decode(result.stdout);
  assert.match(output.error, /No configured profile selected/);
  assert.equal(output.code, "AUTH_REQUIRED");
});

test("user list leaf help advertises the read journey flags", () => {
  const result = run(["entra", "user", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /--select/);
  assert.match(result.stdout, /--cursor/);
  assert.match(result.stdout, /--full/);
  assert.match(result.stdout, /--filter/);
});

test("group member list leaf help advertises the relationship flags", () => {
  const result = run(["entra", "group", "member", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /--group.*required/);
  assert.match(result.stdout, /--transitive/);
  assert.match(result.stdout, /--cursor/);
});

test("pim active list leaf help advertises the state distinction and flags", () => {
  const result = run(["entra", "pim", "active", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /directly assigned plus activated eligible/);
  assert.match(result.stdout, /assignmentType Assigned versus Activated/);
  assert.match(result.stdout, /--filter/);
  assert.match(result.stdout, /--cursor/);
});

test("role-assignment list leaf help describes current inventory and points to state classification", () => {
  const result = run(["entra", "role-assignment", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /including direct and PIM-activated assignments/);
  assert.match(result.stdout, /pim active list for assignmentType Assigned versus Activated/);
});

test("pim eligible list leaf help marks eligibility as not active", () => {
  const result = run(["entra", "pim", "eligible", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /not active/);
  assert.match(result.stdout, /RoleEligibilitySchedule\.Read\.Directory/);
});

test("access-review definition list leaf help marks schedules distinct from instances", () => {
  const result = run(["entra", "access-review", "definition", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /review schedules.*never their occurrences/);
  assert.match(result.stdout, /--cursor/);
  assert.match(result.stdout, /AccessReview\.Read\.All/);
});

test("access-review decision list leaf help marks decisions read-only with parent flags", () => {
  const result = run(["entra", "access-review", "decision", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /never approves, denies or applies anything/);
  assert.match(result.stdout, /--definition.*required/);
  assert.match(result.stdout, /--instance.*required/);
});

test("organization list leaf help names the singular organization route", () => {
  const result = run(["entra", "organization", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /--filter is unsupported on \/organization /);
  assert.ok(!result.stdout.includes("/organizations"));
});

test("access-review decision show leaf help marks the single read-only with parent flags", () => {
  const result = run(["entra", "access-review", "decision", "show", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /never approves, denies or applies anything/);
  assert.match(result.stdout, /--definition.*required/);
  assert.match(result.stdout, /--instance.*required/);
  assert.match(result.stdout, /--id.*required/);
});

test("access-review contacted-reviewer list leaf help marks recorded identities with parent flags", () => {
  const result = run(["entra", "access-review", "contacted-reviewer", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /whether or not notified, never review outcomes/);
  assert.match(result.stdout, /--definition.*required/);
  assert.match(result.stdout, /--instance.*required/);
  assert.match(result.stdout, /--cursor/);
});

test("access-review contacted-reviewer show leaf help marks the identity set", () => {
  const result = run(["entra", "access-review", "contacted-reviewer", "show", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /full reviewed identity set/);
  assert.match(result.stdout, /--id.*required/);
});

test("access-review stage list leaf help marks sequential phases and eq-only filters", () => {
  const result = run(["entra", "access-review", "stage", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /sequential phases/);
  assert.match(result.stdout, /eq only/);
  assert.match(result.stdout, /--definition.*required/);
  assert.match(result.stdout, /--instance.*required/);
});

test("access-review stage show leaf help marks reviewer scopes with later-slice decisions", () => {
  const result = run(["entra", "access-review", "stage", "show", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /reviewer scopes/);
  assert.match(result.stdout, /per-stage decisions belong to a later slice/);
});

test("administrative-unit member list leaf help advertises the relationship flags", () => {
  const result = run(["entra", "administrative-unit", "member", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /--administrative-unit.*required/);
  assert.match(result.stdout, /--cursor/);
  assert.match(result.stdout, /Member\.Read\.Hidden/);
});

test("oauth2-grant list leaf help marks rows as granted consent with read-only scopes", () => {
  const result = run(["entra", "service-principal", "oauth2-grant", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /granted consent, distinct from the application's requested permissions/);
  assert.match(result.stdout, /Directory\.Read\.All/);
  assert.match(result.stdout, /never request write-consent scopes/);
  assert.match(result.stdout, /--service-principal.*required/);
});

test("app-role-assignment list leaf help marks rows as granted consent with read-only scopes", () => {
  const result = run(["entra", "service-principal", "app-role-assignment", "list", "--help"]);
  assert.equal(result.status, 0);
  assert.equal(result.stderr, "");
  assert.match(result.stdout, /granted consent, distinct from the application's requested permissions/);
  assert.match(result.stdout, /Application\.Read\.All/);
  assert.match(result.stdout, /--service-principal.*required/);
});

test("group member list without a profile fails operationally on stdout", () => {
  const result = run(["entra", "group", "member", "list", "--group", "fixture"]);
  assert.equal(result.status, 1);
  assert.equal(result.stderr, "");
  const output = decode(result.stdout);
  assert.equal(output.code, "AUTH_REQUIRED");
  assert.ok(output.help.length);
});

test("directory-role show without a profile fails operationally on stdout", () => {
  const result = run(["entra", "directory-role", "show", "--id", "fixture"]);
  assert.equal(result.status, 1);
  assert.equal(result.stderr, "");
  const output = decode(result.stdout);
  assert.equal(output.code, "AUTH_REQUIRED");
  assert.ok(output.help.length);
});

test("user list without a profile fails operationally on stdout", () => {
  const result = run(["entra", "user", "list"]);
  assert.equal(result.status, 1);
  assert.equal(result.stderr, "");
  const output = decode(result.stdout);
  assert.equal(output.code, "AUTH_REQUIRED");
  assert.ok(output.help.length);
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
