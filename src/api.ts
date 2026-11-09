import { AxiError } from "axi-sdk-js";
import { ApplicationAuth } from "./app-auth.js";
import { DelegatedAuth } from "./auth.js";
import { DEFAULT_MAX_BYTES, GraphSession, rawReadOperation, type GraphTransport } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

const TRUNCATE_AT = 500;
const RAW_VERSIONS = ["v1.0", "beta"];
const TRUNCATED_HINT = "Strings truncated at 500 chars; re-run with --full";
const PAGING_HINT = "Resume the same path, profile and scopes with --cursor - and supply the cursor token on stdin; use --limit <rows> or --all";

function rawPath(pathname: string): string {
  const segments = pathname.split("/").slice(1);
  const invalid = !pathname.startsWith("/") || /:\/\/|[?#\\]/.test(pathname) || segments.length === 0
    || segments.some(segment => segment === "" || segment === "." || segment === "..");
  if (invalid) {
    throw new AxiError(`Refused raw path '${pathname || "(empty)"}'`, "VALIDATION_ERROR", [
      "Pass a server-relative Graph path such as /users; query belongs in --odata 'k=v&k2=v2'",
      "The command never takes a host, scheme or headers; the session owns the Graph destination",
    ]);
  }
  return pathname;
}

function parseQuery(raw: string | undefined): Record<string, string> {
  if (!raw) return {};
  const text = raw.startsWith("?") ? raw.slice(1) : raw;
  const params = new URLSearchParams(text);
  const out: Record<string, string> = Object.create(null);
  for (const [key, value] of params) {
    if (Object.hasOwn(out, key)) throw new AxiError(`Duplicate query key ${key}`, "VALIDATION_ERROR", ["Pass each query key once: --odata 'k=v&k2=v2'"]);
    out[key] = value;
  }
  if (!Object.keys(out).length && text.trim() !== "") throw new AxiError(`Invalid --odata '${raw}'`, "VALIDATION_ERROR", ["Example: --odata '$top=5&$select=id,displayName'"]);
  return out;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && Object.getPrototypeOf(value) === Object.prototype;
}

// String truncation at the output boundary, matching the 500-char rule used
// by every Entra command. --full disables truncation; redaction already ran
// in the session and row caps are unaffected. Plain objects and arrays only.
function truncateForOutput(value: unknown, full: boolean): { value: unknown; truncated: boolean } {
  if (typeof value === "string") {
    if (full || value.length <= TRUNCATE_AT) return { value, truncated: false };
    return { value: `${value.slice(0, TRUNCATE_AT)}... (truncated, ${value.length} chars total)`, truncated: true };
  }
  if (Array.isArray(value)) {
    let truncated = false;
    const shaped = value.map(item => {
      const child = truncateForOutput(item, full);
      truncated = truncated || child.truncated;
      return child.value;
    });
    return { value: shaped, truncated };
  }
  if (isRecord(value)) {
    let truncated = false;
    const shaped: Record<string, unknown> = {};
    for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
      const next = truncateForOutput(child, full);
      truncated = truncated || next.truncated;
      shaped[key] = next.value;
    }
    return { value: shaped, truncated };
  }
  return { value, truncated: false };
}

function singleOutput(body: unknown, full: boolean): Record<string, unknown> {
  const shaped = truncateForOutput(body, full);
  const record = isRecord(shaped.value) ? shaped.value : { value: shaped.value };
  return shaped.truncated ? { ...record, help: [TRUNCATED_HINT] } : record;
}

export interface ApiGetArgs {
  path: string;
  apiVersion: string;
  odata?: string;
  cursor?: string;
  scopes?: string;
  /** Row cap; undefined with --all follows pages within the session budget. */
  limit?: number;
  full?: boolean;
  profile: AnyProfile;
}

export interface ApiDeps {
  delegated: DelegatedAuth;
  application: ApplicationAuth;
  transport: GraphTransport;
}

export async function runApiGet(args: ApiGetArgs, deps: ApiDeps): Promise<Record<string, unknown>> {
  const version = args.apiVersion || "v1.0";
  const pathname = rawPath(args.path ?? "");
  if (!RAW_VERSIONS.includes(version)) {
    throw new AxiError(`Unsupported raw api version ${version}`, "VALIDATION_ERROR", ["Use --api-version v1.0 or beta"]);
  }
  if (args.profile.mode === "application" && args.scopes !== undefined) {
    throw new AxiError("Application profiles use the configured Graph .default audience; delegated scopes are unavailable", "VALIDATION_ERROR", [
      "mg-axi profile show --profile <name>",
    ]);
  }
  const operation = rawReadOperation(version, pathname);
  const scopes = args.scopes === undefined ? undefined : args.scopes.split(",").map(scope => scope.trim()).filter(scope => scope.length > 0);
  const session = new GraphSession(deps);
  const request = { profile: args.profile, operation, raw: true as const, params: {}, scopes };
  const query = args.cursor !== undefined && args.odata === undefined ? undefined : parseQuery(args.odata);
  if (query !== undefined && Object.hasOwn(query, "$skiptoken")) {
    throw new AxiError("Server paging state cannot be supplied; page with --all", "VALIDATION_ERROR", ["mg-axi api get --help"]);
  }
  const full = !!args.full;
  if (pathname.endsWith("/$count")) {
    return { count: await session.execute({ ...request, query, scalar: true, consistencyLevel: "eventual" }) as number };
  }
  const consistencyLevel = query && (Object.hasOwn(query, "$search") || String(query.$count).toLowerCase() === "true") ? "eventual" as const : undefined;
  const collected = await session.collect({ ...request, query, consistencyLevel, limit: args.limit, cursor: args.cursor, singleResource: true });
  if (collected.singleResource && collected.complete && collected.value.length === 1) return singleOutput(collected.value[0], full);
  const rows = truncateForOutput(collected.value, full);
  const metadata = truncateForOutput(collected.metadata ?? {}, full);
  const page = { ...(metadata.value as Record<string, unknown>), returned: collected.value.length, complete: collected.complete, value: rows.value };
  if (collected.complete) return { ...page, ...(rows.truncated ? { help: [TRUNCATED_HINT] } : {}) };
  const help = rows.truncated ? [PAGING_HINT, TRUNCATED_HINT] : [PAGING_HINT];
  return { ...page, reason: collected.reason, cursor: collected.cursor, help };
}

// Minimal HTTPS transport: sends only what the session authorized (GET on the
// commercial Graph host with session-owned headers) and returns status,
// headers and text. Timeouts and cancellation arrive through the session's
// AbortSignal; credential attachment happened before this call.
export const fetchTransport: GraphTransport = async request => {
  const response = await fetch(request.url, { method: request.method, body: request.body, headers: request.headers, signal: request.signal, redirect: "manual" });
  const headers: Record<string, string> = {};
  response.headers.forEach((value, key) => {
    headers[key] = value;
  });
  if (request.method === "POST" && response.body) {
    const reader = response.body.getReader();
    const chunks: Uint8Array[] = [];
    let bytes = 0;
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        bytes += value.length;
        if (bytes > DEFAULT_MAX_BYTES) {
          await reader.cancel();
          throw new AxiError("Hunting response exceeds the 5 MB client budget", "GRAPH_ERROR", ["Reduce --limit or project fewer columns"]);
        }
        chunks.push(value);
      }
    } finally {
      reader.releaseLock();
    }
    return { status: response.status, headers, body: Buffer.concat(chunks).toString("utf8") };
  }
  return { status: response.status, headers, body: await response.text() };
};
