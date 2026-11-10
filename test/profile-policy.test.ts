import assert from "node:assert/strict";
import { test } from "node:test";
import { validateApplicationProfile, validateDelegatedProfile } from "../dist/profiles.js";
import { record } from "./fixtures/typed.ts";

const tenant = "11111111-1111-4111-8111-111111111111";
const client = "22222222-2222-4222-8222-222222222222";
const delegated = { mode: "delegated", tenantId: tenant, clientId: client, cloud: "commercial", token: { stored: true } };
const application = { mode: "application", tenantId: tenant, clientId: client, cloud: "commercial", auth: "federated", token: { env: "MG_AXI_TOKEN" } };

test("denials stay closed when the profile omits them", () => {
  for (const profile of [validateDelegatedProfile(delegated), validateApplicationProfile(application)]) {
    assert.deepEqual(profile.sensitiveAreas, []);
    assert.equal(profile.rawBatch, undefined);
    assert.equal(profile.rawWrites, undefined);
  }
});

test("a host-injected profile carries sensitive areas and rawBatch", () => {
  const injected = { sensitiveAreas: ["files", "mail"], rawBatch: true };
  for (const profile of [validateDelegatedProfile({ ...delegated, ...injected }), validateApplicationProfile({ ...application, ...injected })]) {
    assert.deepEqual(profile.sensitiveAreas, ["files", "mail"]);
    assert.equal(profile.rawBatch, true);
    assert.throws(() => (profile.sensitiveAreas as string[]).push("x"), TypeError);
  }
});

for (const [name, policy] of Object.entries({
  "unknown sensitive area": { sensitiveAreas: ["calendar"] },
  "duplicate sensitive area": { sensitiveAreas: ["files", "files"] },
  "non-array sensitive areas": { sensitiveAreas: "files" },
  "non-boolean rawBatch": { rawBatch: "true" },
})) test(`profile validation rejects ${name}`, () => {
  assert.throws(() => validateDelegatedProfile({ ...delegated, ...policy }), error => record(error).code === "CONFIG_INVALID");
});
