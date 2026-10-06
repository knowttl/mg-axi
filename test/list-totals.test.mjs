import assert from "node:assert/strict";
import { test } from "node:test";
import { listTotals } from "../dist/list-totals.js";

for (const [name, shown, total, noun, complete, expected] of [
  ["known total names returned of available", 3, 10, "groups", false, { total: 10, count: "3 of 10 groups" }],
  ["known total on the final resumed window still names the total", 2, 10, "groups", true, { total: 10, count: "2 of 10 groups" }],
  ["matching total collapses to the row count", 3, 3, "groups", true, { total: 3, count: "3 groups" }],
  ["known total on a partial read always names the total", 3, 3, "groups", false, { total: 3, count: "3 of 3 groups" }],
  ["unknown total on a complete read names its rows", 4, null, "members", true, { total: null, count: "4 members" }],
  ["unknown total on a partial read says more is available", 1, null, "members", false, { total: null, count: "1 members shown, more available" }],
  ["undefined total resumes as unknown", 2, undefined, "memberships", true, { total: null, count: "2 memberships" }],
  ["empty complete read stays definitive", 0, 0, "groups", true, { total: 0, count: "0 groups" }],
  ["empty unknown read stays definitive", 0, null, "groups", true, { total: null, count: "0 groups" }],
]) {
  test(`listTotals ${name}`, () => {
    assert.deepEqual(listTotals(shown, total, noun, complete), expected);
  });
}

for (const [name, args] of [
  ["negative shown", [-1, null, "groups", true]],
  ["fractional shown", [1.5, null, "groups", true]],
  ["negative total", [1, -1, "groups", true]],
  ["fractional total", [1, 1.5, "groups", true]],
  ["empty noun", [1, null, "", true]],
]) {
  test(`listTotals rejects ${name}`, () => {
    assert.throws(() => listTotals(...args), error => error.code === "VALIDATION_ERROR");
  });
}
