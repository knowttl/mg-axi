import assert from "node:assert/strict";
import { once } from "node:events";
import { createServer } from "node:http";
import { test } from "node:test";
import { fetchTransport } from "../dist/api.js";
import type { AddressInfo } from "node:net";
import { record, text } from "./fixtures/typed.ts";

test("fetch transport sends hunting JSON and refuses oversized response streams", async () => {
  const requests: { method: string | undefined; body: string }[] = [];
  const server = createServer(async (request, response) => {
    const chunks = [];
    for await (const chunk of request) chunks.push(chunk);
    requests.push({ method: request.method, body: Buffer.concat(chunks).toString("utf8") });
    response.writeHead(200, { "Content-Type": "application/json" });
    response.end("x".repeat(5_000_001));
  });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  try {
    const body = JSON.stringify({ Query: "EmailEvents | take 2", Timespan: "P1D" });
    await assert.rejects(fetchTransport({ method: "POST", url: `http://127.0.0.1:${(server.address() as AddressInfo).port}/hunting`,
      headers: { Authorization: "Bearer offline-fixture", "Content-Type": "application/json; charset=utf-8" }, body }),
    error => record(error).code === "GRAPH_ERROR" && /5 MB/.test(text(record(error).message)));
    assert.deepEqual(requests, [{ method: "POST", body }]);
  } finally {
    await new Promise<void>((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
  }
});

test("fetch transport returns redirects for session authorization without following them", async () => {
  const requests: (string | undefined)[] = [];
  const server = createServer((request, response) => {
    requests.push(request.url);
    if (request.url === "/allowed") {
      response.writeHead(302, { location: "/denied" });
      response.end("redirect body");
    } else {
      response.writeHead(200);
      response.end("unauthorized response");
    }
  });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  try {
    const result = await fetchTransport({
      method: "GET",
      url: `http://127.0.0.1:${(server.address() as AddressInfo).port}/allowed`,
      headers: { Authorization: "Bearer offline-fixture" },
    });
    assert.equal(result.status, 302);
    assert.equal(result.headers.location, "/denied");
    assert.equal(result.body, "redirect body");
    assert.deepEqual(requests, ["/allowed"]);
  } finally {
    await new Promise<void>((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
  }
});
