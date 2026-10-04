import assert from "node:assert/strict";
import { once } from "node:events";
import { createServer } from "node:http";
import { test } from "node:test";
import { fetchTransport } from "../dist/graph-transport.js";

test("fetch transport returns redirects for session authorization without following them", async () => {
  const requests = [];
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
      url: `http://127.0.0.1:${server.address().port}/allowed`,
      headers: { Authorization: "Bearer offline-fixture" },
    });
    assert.equal(result.status, 302);
    assert.equal(result.headers.location, "/denied");
    assert.equal(result.body, "redirect body");
    assert.deepEqual(requests, ["/allowed"]);
  } finally {
    await new Promise((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
  }
});
