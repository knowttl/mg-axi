import assert from "node:assert/strict";
import { test } from "node:test";
import { resolveSessionOperation } from "../dist/graph-session.js";
import { ResponseShapeError } from "../dist/response-shape-error.js";

for (const [method, path] of [
  ["GET", "/users/{user-id}/mailFolders"],
  ["POST", "/security/microsoft.graph.security.runHuntingQuery"],
  ["GET", "/directory/attributeSets"],
  ["GET", "/directory/customSecurityAttributeDefinitions"],
] as const) test(`response diagnostics hide every key on ${path}`, () => {
  const operation = resolveSessionOperation("v1.0", method, path);
  const response = { status: 200, headers: { "content-type": "application/json; private=ignored" }, body: "" };
  const error = new ResponseShapeError("Graph response is malformed", operation, response,
    { PrivateIdentifier: [{ AnotherIdentifier: "private-value" }] }, [{ path: ["PrivateIdentifier", 0, "AnotherIdentifier"], expected: "number" }]);
  assert.deepEqual(error.diagnostics, {
    status: 200, contentType: "application/json",
    shape: [{ path: "$", type: "object" }, { path: '$["<key>"]', type: "array", count: 1 }, { path: '$["<key>"][*]', type: "object" }, { path: '$["<key>"][*]["<key>"]', type: "string", expected: "number" }],
    shapeTruncated: false,
    issues: [{ path: '$["<key>"][*]["<key>"]', received: "string", expected: "number" }],
    issuePaths: ['$["<key>"][*]["<key>"]'], issuePathsTruncated: false,
  });
  assert.doesNotMatch(JSON.stringify(error), /PrivateIdentifier|AnotherIdentifier|private-value/);
});
