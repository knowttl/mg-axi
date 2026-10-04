import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { decode } from "@toon-format/toon";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const groupScopes = ["https://graph.microsoft.com/GroupMember.Read.All"];

const g1 = {
  id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
  displayName: "Engineering",
  description: "Product engineering",
  mail: "eng@contoso.com",
  mailEnabled: true,
  mailNickname: "eng",
  securityEnabled: false,
  groupTypes: ["Unified"],
  visibility: "Private",
  classification: null,
  isAssignableToRole: false,
  createdDateTime: "2024-01-01T00:00:00Z",
  expirationDateTime: null,
  renewedDateTime: "2024-06-01T00:00:00Z",
  membershipRule: null,
  membershipRuleProcessingState: null,
};
const g2 = {
  id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  displayName: "Helpdesk Admins",
  description: null,
  mail: null,
  mailEnabled: false,
  mailNickname: "helpdesk-admins",
  securityEnabled: true,
  groupTypes: [],
  visibility: null,
  classification: null,
  isAssignableToRole: true,
  createdDateTime: "2024-02-01T00:00:00Z",
  expirationDateTime: null,
  renewedDateTime: null,
  membershipRule: null,
  membershipRuleProcessingState: null,
};
const g3 = {
  id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  displayName: "Niche",
  mail: "niche@contoso.com",
  groupTypes: ["Unified"],
};
const groups = [g1, g2, g3];

const mUser = { "@odata.type": "#microsoft.graph.user", id: "dddddddd-dddd-4ddd-8ddd-dddddddddddd", displayName: "Adele Vance" };
const mGroup = { "@odata.type": "#microsoft.graph.group", id: g2.id, displayName: "Helpdesk Admins" };
const mLimited = { "@odata.type": "#microsoft.graph.servicePrincipal", id: "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee" };
const mNullName = { "@odata.type": "#microsoft.graph.user", id: "ffffffff-ffff-4fff-8fff-ffffffffffff", displayName: null };
const members = [mUser, mGroup, mLimited, mNullName];

const p1 = { "@odata.type": "#microsoft.graph.group", id: "99999999-9999-4999-8999-999999999999", displayName: "All Staff" };
const memberOf = [p1];

function setupProfiles() {
  const dir = mkdtempSync(join(tmpdir(), "mg-axi-read-02-"));
  const previous = process.env.MG_AXI_CONFIG;
  process.env.MG_AXI_CONFIG = join(dir, "config.json");
  const profiles = new Profiles();
  profiles.create("soc", tenant, client, "commercial", false);
  profiles.create("batch", tenant, client, "commercial", false, { federated: true });
  return { dir, previous };
}

function teardownProfiles(state) {
  process.env.MG_AXI_CONFIG = state.previous;
  rmSync(state.dir, { recursive: true, force: true });
}

function credentialService(mode, calls) {
  return {
    credential: async (...args) => {
      calls.push(args);
      return {
        token: `opaque-fixture-${mode}-token`,
        expiresAt: Date.now() + 3_600_000,
        tenantId: tenant,
        clientId: client,
        ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
      };
    },
  };
}

function transport(handler) {
  const requests = [];
  const send = async request => {
    requests.push(request);
    const response = await handler(request);
    return { headers: {}, body: "", ...response };
  };
  return { requests, send };
}

function json(status, body, headers = {}) {
  return { status, headers, body: JSON.stringify(body) };
}

function groupTransport() {
  return transport(request => {
    const url = new URL(request.url);
    const path = url.pathname;
    if (path === "/v1.0/groups") {
      if (url.searchParams.has("$skiptoken")) return json(200, { value: [g3] });
      return json(200, {
        value: [g1, g2],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/groups?%24skiptoken=page2",
      });
    }
    const single = /^\/v1\.0\/groups\/([^/]+)$/.exec(path);
    if (single) {
      const found = groups.find(group => group.id === decodeURIComponent(single[1]));
      return found ? json(200, found) : json(404, { error: { code: "Request_ResourceNotFound", message: "no such group" } });
    }
    const rel = /^\/v1\.0\/groups\/[^/]+\/(members|transitiveMembers|memberOf|transitiveMemberOf)$/.exec(path);
    if (rel) {
      const rows = rel[1] === "members" || rel[1] === "transitiveMembers" ? members : memberOf;
      return json(200, { value: rows });
    }
    return json(404, { error: { code: "Unknown", message: "unexpected route" } });
  });
}

function overridesFor(mode, handler, calls = []) {
  const fixture = handler ?? groupTransport();
  const credential = credentialService(mode, calls);
  return {
    requests: fixture.requests,
    calls,
    overrides: {
      transport: fixture.send,
      delegated: mode === "delegated" ? credential : credentialService("delegated", []),
      application: mode === "application" ? credential : credentialService("application", []),
    },
  };
}

function runGroupCli(args, state, mode, denied = false) {
  return spawnSync(process.execPath, [
    "--experimental-test-module-mocks", "--disable-warning=ExperimentalWarning",
    "--import", pathToFileURL(resolve("test/fixtures/read-cli.mjs")).href, resolve("dist/bin/mg-axi.js"), ...args,
  ], {
    encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000,
    env: {
      HOME: state.dir, USERPROFILE: state.dir, PATH: process.env.PATH, SystemRoot: process.env.SystemRoot,
      MG_AXI_CONFIG: join(state.dir, "config.json"),
      MG_AXI_READ_FIXTURE: JSON.stringify({ mode, directory: [], denied, scopes: groupScopes[0], groups, members, memberOf }),
    },
  });
}

for (const [mode, profile] of [["delegated", "soc"], ["application", "batch"]]) {
  test(`${mode} lists groups with compact rows preserving null and missing`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group", "list", "--profile", profile], overrides);
      assert.deepEqual(result.groups, [
        { id: g1.id, displayName: "Engineering", mail: "eng@contoso.com", groupTypes: ["Unified"] },
        { id: g2.id, displayName: "Helpdesk Admins", mail: null, groupTypes: [] },
        { id: g3.id, displayName: "Niche", mail: "niche@contoso.com", groupTypes: ["Unified"] },
      ]);
      assert.deepEqual(result.count, { returned: 3, complete: true });
      assert.ok(result.help.some(hint => hint.includes("entra group show --id <group-id>")));
      assert.ok(requests.every(request => request.headers.Authorization === `Bearer opaque-fixture-${mode}-token`));
      assert.ok(requests[0].url.startsWith("https://graph.microsoft.com/v1.0/groups?"));
      assert.ok(!JSON.stringify(result).includes(`opaque-fixture-${mode}-token`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} shows a group with richer properties`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group", "show", "--id", g1.id, "--profile", profile], overrides);
      assert.deepEqual(result.group, g1);
      assert.equal(result.help, undefined);
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} show marks role-assignable groups without mutating them`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group", "show", "--id", g2.id, "--profile", profile], overrides);
      assert.equal(result.group.isAssignableToRole, true);
      assert.ok(result.help.some(hint => hint.includes("Role-assignable group")));
      assert.ok(result.help.some(hint => hint.includes("mg-axi entra group member add")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists direct members with kinds and the service-principal warning`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--profile", profile,
        "--select", "id,displayName"], overrides);
      assert.deepEqual(result.members, [
        mUser,
        mGroup,
        { "@odata.type": "#microsoft.graph.servicePrincipal", id: mLimited.id },
        mNullName,
      ]);
      assert.deepEqual(result.count, { returned: 4, complete: true });
      assert.ok(result.warnings.some(warning => warning.includes("service principals")));
      assert.ok(result.warnings.some(warning => warning.includes("complete membership")));
      assert.ok(result.help.some(hint => hint.includes("Member.Read.Hidden")));
      assert.ok(result.help.some(hint => hint.includes("--transitive")));
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/groups/${g1.id}/members`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} transitive members flatten nesting without the v1.0 warning`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--transitive", "--profile", profile], overrides);
      assert.equal(result.members.length, 4);
      assert.equal(result.warnings, undefined);
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/groups/${g1.id}/transitiveMembers`));
      assert.ok(result.help.some(hint => hint.includes("Direct relationships only")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} lists memberOf and transitiveMemberOf`, async () => {
    const state = setupProfiles();
    try {
      const { requests, overrides } = overridesFor(mode);
      const direct = await executeArgv(["entra", "group", "member-of", "list", "--group", g1.id, "--profile", profile], overrides);
      assert.deepEqual(direct.memberOf, memberOf);
      assert.deepEqual(direct.count, { returned: 1, complete: true });
      assert.equal(direct.warnings, undefined);
      assert.ok(new URL(requests[0].url).pathname.endsWith(`/groups/${g1.id}/memberOf`));
      const transitive = await executeArgv(["entra", "group", "member-of", "list", "--group", g1.id, "--transitive", "--profile", profile], overrides);
      assert.deepEqual(transitive.memberOf, memberOf);
      assert.ok(new URL(requests[1].url).pathname.endsWith(`/groups/${g1.id}/transitiveMemberOf`));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} preserves limited-information members with a count hint`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const result = await executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--profile", profile], overrides);
      const limited = result.members.filter(row => row.id === mLimited.id || row.id === mNullName.id);
      assert.equal(limited.length, 2);
      assert.deepEqual(limited[0], { id: mLimited.id, "@odata.type": "#microsoft.graph.servicePrincipal" });
      assert.ok(result.help.some(hint => hint.includes("2 of 4 rows have no non-null selected descriptive properties")));
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} resumes a capped member list through its opaque cursor`, async () => {
    const state = setupProfiles();
    try {
      const { overrides } = overridesFor(mode);
      const first = await executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--profile", profile, "--limit", "1"], overrides);
      assert.deepEqual(first.members.map(row => row.id), [mUser.id]);
      assert.equal(first.count.complete, false);
      assert.equal(typeof first.cursor, "string");
      const second = await executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--profile", profile, "--cursor", first.cursor], overrides);
      assert.deepEqual(second.members.map(row => row.id), [mGroup.id, mLimited.id, mNullName.id]);
      assert.deepEqual(second.count, { returned: 3, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} denied group reads surface as operational failures`, async () => {
    const state = setupProfiles();
    try {
      const denied = transport(() => json(403, { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } }));
      const missing = transport(() => json(404, { error: { code: "Request_ResourceNotFound", message: "gone" } }));
      const overrides = {
        transport: denied.send,
        delegated: credentialService("delegated", []),
        application: credentialService("application", []),
      };
      await assert.rejects(executeArgv(["entra", "group", "list", "--profile", profile], overrides), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /grant, role, licence/.test(error.message);
      });
      await assert.rejects(executeArgv(["entra", "group", "show", "--id", g1.id, "--profile", profile],
        { ...overrides, transport: missing.send }), error => {
        assert.equal(error.code, "GRAPH_ERROR");
        return /not found or inaccessible/.test(error.message);
      });
    } finally {
      teardownProfiles(state);
    }
  });

  test(`${mode} executable lists groups, shows one and lists members`, () => {
    const state = setupProfiles();
    try {
      const listed = runGroupCli(["entra", "group", "list", "--profile", profile], state, mode);
      assert.equal(listed.status, 0, listed.stdout);
      assert.equal(listed.stderr, "");
      const groupsOut = decode(listed.stdout);
      assert.deepEqual(groupsOut.groups.map(group => group.id), [g1.id, g2.id, g3.id]);
      assert.deepEqual(groupsOut.count, { returned: 3, complete: true });
      assert.ok(!listed.stdout.includes(`opaque-fixture-${mode}-token`));

      const shown = runGroupCli(["entra", "group", "show", "--id", g1.id, "--profile", profile], state, mode);
      assert.equal(shown.status, 0, shown.stdout);
      assert.deepEqual(decode(shown.stdout).group, g1);

      const membersOut = runGroupCli(["entra", "group", "member", "list", "--group", g1.id, "--profile", profile], state, mode);
      assert.equal(membersOut.status, 0, membersOut.stdout);
      const membersDecoded = decode(membersOut.stdout);
      assert.equal(membersDecoded.members.length, 4);
      assert.ok(membersDecoded.warnings.some(warning => warning.includes("service principals")));
      assert.ok(!membersOut.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });

  test(`${mode} executable denied group reads fail operationally on stdout`, () => {
    const state = setupProfiles();
    try {
      const result = runGroupCli(["entra", "group", "member", "list", "--group", g1.id, "--profile", profile], state, mode, true);
      assert.equal(result.status, 1, result.stdout);
      assert.equal(result.stderr, "");
      const output = decode(result.stdout);
      assert.equal(output.code, "GRAPH_ERROR");
      assert.match(output.error, /grant, role, licence or policy/);
      assert.equal(output.members, undefined);
      assert.ok(!result.stdout.includes(`opaque-fixture-${mode}-token`));
    } finally { teardownProfiles(state); }
  });
}

for (const [relationship, route, transitive] of [
  ["member", "members", false],
  ["member", "transitiveMembers", true],
  ["member-of", "memberOf", false],
  ["member-of", "transitiveMemberOf", true],
]) {
  test(`${route} filtering carries count and consistency through resume`, async () => {
    const state = setupProfiles();
    try {
      const seen = [];
      const fixture = transport(request => {
        seen.push({ url: request.url, headers: request.headers });
        assert.equal(new URL(request.url).searchParams.get("$count"), "true");
        if (new URL(request.url).searchParams.has("$skiptoken")) {
          return json(200, { value: [{ id: g2.id, displayName: "Helpdesk Admins" }] });
        }
        return json(200, {
          value: [{ id: g1.id, displayName: "Engineering" }],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0/groups/${g1.id}/${route}?%24count=true&%24skiptoken=next`,
        });
      });
      const { overrides } = overridesFor("delegated", fixture);
      const command = ["entra", "group", relationship, "list", "--group", g1.id, ...(transitive ? ["--transitive"] : [])];
      const filtered = await executeArgv([...command,
        "--profile", "soc", "--filter", "startswith(displayName,'A')", "--limit", "1"], overrides);
      assert.equal(seen[0].headers.ConsistencyLevel, "eventual");
      assert.equal(filtered.count.complete, false);
      const resumed = await executeArgv([...command,
        "--profile", "soc", "--cursor", filtered.cursor], overrides);
      assert.equal(seen[1].headers.ConsistencyLevel, "eventual");
      assert.deepEqual(resumed.count.complete, true);
      await assert.rejects(executeArgv([...command,
        "--profile", "soc", "--cursor", filtered.cursor, "--filter", "startswith(displayName,'B')"], overrides),
      { code: "VALIDATION_ERROR" });
      const plain = transport(request => {
        seen.push({ url: request.url, headers: request.headers });
        return json(200, { value: [{ id: g1.id, displayName: "Engineering" }] });
      });
      const unfiltered = await executeArgv(["entra", "group", "list", "--profile", "soc", "--limit", "1"],
        overridesFor("delegated", plain).overrides);
      assert.equal(seen[seen.length - 1].headers.ConsistencyLevel, undefined);
      assert.equal(new URL(seen[seen.length - 1].url).searchParams.has("$count"), false);
      assert.deepEqual(unfiltered.count, { returned: 1, complete: true });
    } finally {
      teardownProfiles(state);
    }
  });

  for (const [name, row, select, fields, expectedHint] of [
    ["readable group projected to nullable mail", { ...mGroup, mail: null }, "id,displayName,mail", "id,mail", undefined],
    ["readable name returned beyond selection", { ...mGroup, mail: null }, "id,mail", "id,mail", undefined],
    ["null name projected to id", mNullName, "id,displayName", "id", "1 of 1 rows have no non-null selected descriptive properties"],
    ["missing name projected to id", mLimited, "id,displayName", "id", "1 of 1 rows have no non-null selected descriptive properties"],
    ["nullable mail alone", { id: mGroup.id, "@odata.type": mGroup["@odata.type"], mail: null }, "id,mail", "id,mail", "1 of 1 rows have no non-null selected descriptive properties"],
    ["id-only selection", mLimited, "id", "id", undefined],
  ]) {
    for (const limit of ["1", "2"]) {
      test(`${route} ${name} assesses fetched fields in ${limit === "1" ? "partial" : "complete"} output`, async () => {
        const state = setupProfiles();
        try {
          const fixture = transport(() => json(200, { value: [row, { ...mUser, mail: "adele@contoso.com" }] }));
          const { overrides } = overridesFor("application", fixture);
          const result = await executeArgv(["entra", "group", relationship, "list", "--group", g1.id,
            ...(transitive ? ["--transitive"] : []), "--profile", "batch", "--select", select, "--fields", fields, "--limit", limit], overrides);
          const hints = result.help.filter(hint => hint.includes("selected descriptive properties"));
          assert.deepEqual(hints, expectedHint === undefined ? [] : [
            `${expectedHint.replace("1 of 1", `1 of ${limit}`)}; this may reflect limited read consent or unset properties`,
          ]);
          assert.equal(result.count.complete, limit === "2");
        } finally {
          teardownProfiles(state);
        }
      });
    }
  }
}

test("direct and transitive cursors do not cross resume", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    const first = await executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--profile", "soc", "--limit", "1"], overrides);
    assert.equal(first.count.complete, false);
    await assert.rejects(executeArgv(["entra", "group", "member", "list", "--group", g1.id,
      "--profile", "soc", "--cursor", first.cursor, "--transitive"], overrides), { code: "VALIDATION_ERROR" });
  } finally {
    teardownProfiles(state);
  }
});

test("group list $filter also carries ConsistencyLevel eventual", async () => {
  const state = setupProfiles();
  try {
    const seen = [];
    const fixture = transport(request => {
      seen.push(request.headers);
      return json(200, { value: [] });
    });
    const { overrides } = overridesFor("delegated", fixture);
    await executeArgv(["entra", "group", "list", "--profile", "soc", "--filter", "securityEnabled eq true"], overrides);
    assert.equal(seen[0].ConsistencyLevel, "eventual");
  } finally {
    teardownProfiles(state);
  }
});

test("application mode rejects delegated scopes before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("application");
    await assert.rejects(
      executeArgv(["entra", "group", "list", "--profile", "batch", "--scopes", groupScopes[0]], overrides),
      error => {
        assert.equal(error.code, "VALIDATION_ERROR");
        return /Graph \.default audience/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--profile", "batch", "--scopes", groupScopes[0]], overrides),
      { code: "VALIDATION_ERROR" },
    );
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("unknown properties and unfetched fields fail before HTTP", async () => {
  const state = setupProfiles();
  try {
    const { requests, overrides } = overridesFor("delegated");
    await assert.rejects(executeArgv(["entra", "group", "list", "--profile", "soc", "--select", "id,aboutMe"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "group", "list", "--profile", "soc", "--fields", "membershipRule"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--profile", "soc", "--select", "id,department"], overrides), { code: "VALIDATION_ERROR" });
    await assert.rejects(executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--profile", "soc", "--cursor", "not-a-cursor"], overrides), { code: "VALIDATION_ERROR" });
    assert.equal(requests.length, 0);
  } finally {
    teardownProfiles(state);
  }
});

test("group relationship flags validate before profiles or HTTP", async () => {
  await assert.rejects(executeArgv(["entra", "group", "member", "list"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "group", "member", "list", "--group="]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--id", g1.id]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "group", "list", "--transitive"]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "group", "show", "--transitive", "--id", g1.id]), { code: "VALIDATION_ERROR" });
  await assert.rejects(executeArgv(["entra", "group", "show"]), { code: "VALIDATION_ERROR" });
});

test("group reads stay preview-gated on beta", async () => {
  const state = setupProfiles();
  try {
    const { overrides } = overridesFor("delegated");
    await assert.rejects(
      executeArgv(["entra", "group", "list", "--profile", "soc", "--api-version", "beta", "--limit", "1"], overrides),
      error => {
        assert.equal(error.code, "POLICY_DENIED");
        return /preview-enabled/.test(error.message);
      },
    );
    await assert.rejects(
      executeArgv(["entra", "group", "member", "list", "--group", g1.id, "--profile", "soc", "--api-version", "beta"], overrides),
      { code: "POLICY_DENIED" },
    );
  } finally {
    teardownProfiles(state);
  }
});
