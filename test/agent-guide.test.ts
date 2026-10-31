import { record, text, authFakes, type OutputRecord } from "./fixtures/typed.ts";
import type { TransportRequest } from "../dist/graph-session.js";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { executeArgv } from "../dist/cli.js";
import { Profiles } from "../dist/profiles.js";

const guide = readFileSync(new URL("../docs/agent-guide.md", import.meta.url), "utf8");
const examples = [...guide.matchAll(/\| `mg-axi ([^`]+)` \|/g)].map(match => match[1]!);
const cases: [string, string, OutputRecord][] = [
  ["/users", "users", { id: "synthetic-user", displayName: "Synthetic Analyst", userPrincipalName: "analyst@example.invalid" }],
  ["/users/synthetic-user", "user", { id: "synthetic-user", displayName: "Synthetic Analyst", userType: "Member" }],
  ["/groups/synthetic-group/transitiveMembers", "members", { id: "synthetic-user", displayName: "Synthetic Analyst", "@odata.type": "#microsoft.graph.user" }],
  ["/auditLogs/signIns", "signIns", { id: "synthetic-sign-in", createdDateTime: "2026-09-01T12:00:00Z", userId: "synthetic-user", status: { errorCode: 50076 } }],
  ["/auditLogs/signIns/synthetic-sign-in", "signIn", { id: "synthetic-sign-in", userId: "synthetic-user", status: { errorCode: 50076 } }],
  ["/auditLogs/directoryAudits", "directoryAudits", { id: "synthetic-audit", activityDateTime: "2026-09-01T12:00:00Z", activityDisplayName: "Update user", result: "success" }],
  ["/identityProtection/riskyUsers", "riskyUsers", { id: "synthetic-user", riskLevel: "high", riskState: "atRisk" }],
  ["/identityProtection/riskDetections/synthetic-detection", "riskDetection", { id: "synthetic-detection", userId: "synthetic-user", activityDateTime: "2026-09-01T12:00:00Z", requestId: "synthetic-sign-in", correlationId: "synthetic-correlation" }],
  ["/applications", "applications", { id: "synthetic-app", appId: "33333333-3333-4333-8333-333333333333", displayName: "Synthetic App" }],
  ["/servicePrincipals/synthetic-sp/oauth2PermissionGrants", "oauth2PermissionGrants", { id: "synthetic-grant", clientId: "synthetic-sp", resourceId: "synthetic-resource", scope: "User.Read" }],
  ["/security/runHuntingQuery", "results", { Timestamp: "2026-09-01T12:00:00Z", DeviceName: "synthetic-device", FileName: "synthetic.exe" }],
];

assert.equal(examples.length, cases.length, "every guide example needs an offline journey");

for (const [index, [path, key, row]] of cases.entries()) {
  test(`guide example works with a synthetic profile: ${examples[index]}`, async t => {
    const dir = mkdtempSync(join(tmpdir(), "mg-axi-agent-guide-"));
    t.after(() => rmSync(dir, { recursive: true, force: true }));
    const config = join(dir, "config.json");
    new Profiles(config).create("soc", "11111111-1111-4111-8111-111111111111", "22222222-2222-4222-8222-222222222222", "commercial", false);
    const value = JSON.parse(readFileSync(config, "utf8"));
    value.profiles.soc.enabledPacks = ["entra", "security"];
    writeFileSync(config, JSON.stringify(value));
    const file = join(dir, "hunt.kql");
    writeFileSync(file, guide.match(/```kql\n([\s\S]*?)```/)![1]!);
    // Guide examples use only double-quoted filter arguments and plain words.
    const args = [...examples[index]!.matchAll(/"([^"]*)"|(\S+)/g)].map(match => match[1] ?? match[2]!);
    const fileIndex = args.indexOf("--file");
    if (fileIndex !== -1) args[fileIndex + 1] = file;
    const requests: TransportRequest[] = [];
    const correlating = key === "riskDetection";
    const result = record(await executeArgv([...args, "--config", config], {
      delegated: authFakes({ credential: async profile => ({ token: "synthetic-token", tenantId: profile.tenantId, clientId: profile.clientId, accountId: "synthetic-account", expiresAt: Date.now() + 3600000 }) }).delegated,
      transport: async request => {
        requests.push(request);
        if (correlating && requests.length === 2) {
          assert.equal(new URL(request.url).pathname, "/v1.0/auditLogs/signIns");
          assert.equal(request.method, "GET");
          return { status: 200, headers: {}, body: JSON.stringify({ value: [{ id: row!.requestId, userId: row!.userId, createdDateTime: row!.activityDateTime, correlationId: row!.correlationId }] }) };
        }
        assert.equal(new URL(request.url).pathname, `/v1.0${path}`);
        const hunting = path === "/security/runHuntingQuery";
        assert.equal(request.method, hunting ? "POST" : "GET");
        const body = hunting ? { schema: Object.keys(row).map(name => ({ name, type: "String" })), results: [row] }
          : / show /.test(text(examples[index])) || correlating ? row : { value: [row] };
        return { status: 200, headers: {}, body: JSON.stringify(body) };
      },
    }));
    assert.equal(requests.length, correlating ? 2 : 1);
    assert.deepEqual(result[key], / show /.test(text(examples[index])) || correlating ? row : [row]);
    if (correlating) assert.equal(result.status, "exact");
  });
}
