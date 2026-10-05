import assert from "node:assert/strict";
import { Socket } from "node:net";
import { mock } from "node:test";

const { mode, contacts, denied, manager, reports, memberships } = JSON.parse(process.env.MG_AXI_READ_FIXTURE);
const allowedDelegated = new Set(["https://graph.microsoft.com/OrgContact.Read.All", "https://graph.microsoft.com/Group.Read.All"]);
const [major, minor] = process.versions.node.split(".").map(Number);
const exportOption = major >= 26 || (major === 25 && minor >= 9) || (major === 24 && minor >= 15)
  ? "exports" : "namedExports";
const noNetwork = () => { throw new Error("Network access is disabled in the offline read journey"); };
mock.method(globalThis, "fetch", noNetwork);
mock.method(Socket.prototype, "connect", noNetwork);

function credential(profile, scopes, expectedMode) {
  assert.equal(mode, expectedMode);
  assert.equal(profile.mode, expectedMode);
  if (mode === "application") {
    assert.deepEqual(scopes, ["https://graph.microsoft.com/.default"]);
  } else {
    assert.ok(scopes.length >= 1 && scopes.length <= 2, `Unexpected delegated scope count ${scopes.length}`);
    for (const scope of scopes) assert.ok(allowedDelegated.has(scope), `Unexpected delegated scope ${scope}`);
  }
  return {
    token: `opaque-fixture-${mode}-token`,
    expiresAt: Date.now() + 3_600_000,
    tenantId: profile.tenantId,
    clientId: profile.clientId,
    ...(mode === "delegated" ? { accountId: "synthetic-account" } : {}),
  };
}

const base = "/v1.0/contacts";

function single(rows, pathname) {
  const found = rows.find(row => pathname.endsWith(`/${row.id}`));
  assert.ok(found, `Unexpected contact route ${pathname}`);
  return found;
}

function castRows(rows, cast) {
  if (cast === undefined) return rows;
  assert.ok(cast === "graph.user" || cast === "graph.orgContact" || cast === "graph.group" || cast === "graph.administrativeUnit", `Unexpected cast ${cast}`);
  return rows.filter(row => row["@odata.type"] === `#microsoft.graph.${cast.slice("graph.".length)}`);
}

// Navigation routes under one contact: manager single, directReports
// collection/casts/singles and their $count scalars, plus the EXT-01n
// memberOf/transitiveMemberOf membership collection/casts/singles and their
// $count scalars. Unknown contact ids 404 like the top-level single route.
function navRoute(url) {
  const rest = url.pathname.slice(`${base}/`.length).split("/");
  const [contactId, head, ...tail] = rest;
  assert.ok(contacts.some(row => row.id === contactId), `Unexpected contact route ${url.pathname}`);
  if (head === "memberOf" || head === "transitiveMemberOf") {
    assert.ok(memberships !== undefined, `No membership fixture for ${url.pathname}`);
    return membershipRoute(url, head, tail);
  }
  assert.ok(manager !== undefined && reports !== undefined, `No navigation fixture for ${url.pathname}`);
  if (head === "manager" && tail.length === 0) return manager;
  if (head === "directReports" && tail.length === 0) return { value: reports };
  if (head === "directReports" && tail.length === 1 && tail[0] === "$count") {
    assert.equal(url.searchParams.size, 0);
    return { scalar: reports.length };
  }
  if (head === "directReports" && tail.length === 1 && tail[0].startsWith("graph.")) {
    return { value: castRows(reports, tail[0]) };
  }
  if (head === "directReports" && tail.length === 2 && tail[1] === "$count" && tail[0].startsWith("graph.")) {
    return { scalar: castRows(reports, tail[0]).length };
  }
  if (head === "directReports" && (tail.length === 1 || (tail.length === 2 && tail[1].startsWith("graph.")))) {
    const [reportId, cast] = tail;
    const candidates = cast === undefined ? reports : castRows(reports, cast);
    const found = candidates.find(row => row.id === reportId);
    assert.ok(found, `Unexpected contact route ${url.pathname}`);
    return found;
  }
  assert.fail(`Unexpected contact route ${url.pathname}`);
}

// Membership routes under one contact: the memberOf/transitiveMemberOf
// collection/casts/singles and their $count scalars, mirroring the
// directReports shapes above.
function membershipRoute(url, head, tail) {
  if (tail.length === 0) return { value: memberships };
  if (tail.length === 1 && tail[0] === "$count") {
    assert.equal(url.searchParams.size, 0);
    return { scalar: memberships.length };
  }
  if (tail.length === 1 && tail[0].startsWith("graph.")) {
    return { value: castRows(memberships, tail[0]) };
  }
  if (tail.length === 2 && tail[1] === "$count" && tail[0].startsWith("graph.")) {
    return { scalar: castRows(memberships, tail[0]).length };
  }
  if (tail.length === 1 || (tail.length === 2 && tail[1].startsWith("graph."))) {
    const [memberId, cast] = tail;
    const candidates = cast === undefined ? memberships : castRows(memberships, cast);
    const found = candidates.find(row => row.id === memberId);
    assert.ok(found, `Unexpected contact route ${url.pathname}`);
    return found;
  }
  assert.fail(`Unexpected contact route ${url.pathname}`);
}

mock.module(new URL("../../dist/msal-provider.js", import.meta.url), {
  [exportOption]: { MsalProvider: class {
    storage = "session-only";
    async login() { throw new Error("Interactive authentication is disabled"); }
    async silent(profile, scopes) { return credential(profile, scopes, "delegated"); }
  } },
});
mock.module(new URL("../../dist/msal-app-provider.js", import.meta.url), {
  [exportOption]: { MsalApplicationProvider: class {
    storage = "session-only";
    async acquire(profile, scopes) { return credential(profile, scopes, "application"); }
  } },
});
mock.module(new URL("../../dist/api.js", import.meta.url), {
  [exportOption]: { ...await import("../../dist/api.js"), fetchTransport: async request => {
    assert.equal(request.headers.Authorization, `Bearer opaque-fixture-${mode}-token`);
    const url = new URL(request.url);
    assert.equal(url.origin, "https://graph.microsoft.com");
    let body;
    let status = 200;
    if (denied) {
      status = 403;
      body = { error: { code: "Authorization_RequestDenied", message: "insufficient grants" } };
    } else if (url.pathname === `${base}/$count`) {
      assert.equal(request.headers.ConsistencyLevel, "eventual");
      return { status, headers: { "Content-Type": "text/plain" }, body: String(contacts.length) };
    } else if (url.pathname === base) {
      body = { value: contacts };
    } else if (url.pathname.startsWith(`${base}/`) && url.pathname.split("/").length > 4) {
      const nav = navRoute(url);
      if (nav !== null && typeof nav === "object" && Object.hasOwn(nav, "scalar")) {
        assert.equal(request.headers.ConsistencyLevel, "eventual");
        return { status, headers: { "Content-Type": "text/plain" }, body: String(nav.scalar) };
      }
      body = nav;
    } else {
      body = single(contacts, url.pathname);
    }
    return { status, headers: {}, body: JSON.stringify(body) };
  } },
});
