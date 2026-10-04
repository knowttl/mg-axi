import type { GraphTransport } from "./graph-session.js";

// Production HTTP adapter for the shared Graph session: it sends exactly what
// the session authorizes and returns status, headers and body unmodified.
// Credential attachment, URL authorization and error translation stay in the
// session; tests substitute fixture transports for this seam.
export const fetchTransport: GraphTransport = async request => {
  const response = await fetch(request.url, {
    method: request.method,
    headers: request.headers,
    signal: request.signal,
  });
  const headers: Record<string, string> = Object.create(null);
  response.headers.forEach((value, name) => {
    headers[name] = value;
  });
  return { status: response.status, headers, body: await response.text() };
};
