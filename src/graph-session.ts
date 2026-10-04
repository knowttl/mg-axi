import { createHash, randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import { setTimeout as delay } from "node:timers/promises";
import { AxiError } from "axi-sdk-js";
import { ApplicationAuth } from "./app-auth.js";
import { DelegatedAuth } from "./auth.js";
import {
  validateApplicationProfile,
  validateDelegatedProfile,
  validateProfile,
  type AnyProfile,
} from "./profiles.js";

// CORE-01: the single path that owns Graph URL construction, operation
// authorization, credential attachment and response validation. Command
// handlers call execute() with a catalogued inventory operation and bound
// parameters; they never see tokens and never touch a raw fetch.

// Redaction marker and secret shapes follow az-axi's redact.ts conventions;
// Graph error pairs never carry az-axi's keyName/value envelope, so only the
// key-name and secret-value rules are adapted here.
export const REDACTED = "***redacted***";
export const SAFE_CREDENTIAL_FIELDS: readonly string[] = ["keyId", "displayName", "startDateTime", "endDateTime"];
// READ-04 protected values: authentication-method phone numbers are PII the
// server itself masks for the Authentication Administrator role, so the
// session replaces them with the marker for every role before collection
// buffering, execute output and cursor decode.
export const GRAPH_HOST = "graph.microsoft.com";
const READ_SCOPES = new Set([
  "AccessReview.Read.All",
  "AdministrativeUnit.Read.All",
  "Application.Read.All",
  "AuditLog.Read.All",
  "CustomSecAttributeDefinition.Read.All",
  "DelegatedAdminRelationship.Read.All",
  "Device.Read.All",
  "Directory.Read.All",
  "Domain.Read.All",
  "EntitlementManagement.Read.All",
  "Group.Read.All",
  "GroupMember.Read.All",
  "GroupMember.ReadBasic.All",
  "IdentityRiskEvent.Read.All",
  "IdentityRiskyUser.Read.All",
  "LicenseAssignment.Read.All",
  "Member.Read.Hidden",
  "Organization.Read.All",
  "OrganizationalBranding.Read.All",
  "Policy.Read.All",
  "Policy.Read.AuthenticationMethod",
  "Policy.Read.ConditionalAccess",
  "PrivilegedEligibilitySchedule.Read.AzureADGroup",
  "RoleAssignmentSchedule.Read.Directory",
  "RoleEligibilitySchedule.Read.Directory",
  "RoleManagement.Read.Directory",
  "Synchronization.Read.All",
  "User.Read",
  "User.Read.All",
  "User.ReadBasic.All",
  "UserAuthenticationMethod.Read",
  "UserAuthenticationMethod.Read.All",
].map(scope => `https://${GRAPH_HOST}/${scope}`));
// Conservative read-query allowlist. Per-operation review (READ slices) can
// extend it; unknown keys fail closed here. $search/$count=true need eventual
// consistency (see checkQueryContext); $skiptoken carries paging state.
export const ALLOWED_QUERY_KEYS: readonly string[] = ["$select", "$filter", "$top", "$orderby", "$count", "$skiptoken", "$search", "$expand"];
const ALLOWED_KEYS = new Set(ALLOWED_QUERY_KEYS);
// Only these dispositions may execute. `scheduled` stays executable at this
// layer so READ slices need no session change; catalogue/CLI gating still
// keeps unshipped commands out of reach. Everything else fails closed.
const EXECUTABLE_DISPOSITIONS = new Set(["named-command", "reviewed-raw-read", "scheduled"]);
const MAX_REDIRECTS = 3;
const MAX_QUERY_VALUE = 1024;
// CORE-02 ceilings: single Retry-After waits above this never retry; the
// collection budget bounds total requests, bytes and wall-clock time.
export const MAX_RETRY_AFTER_MS = 10_000;
export const DEFAULT_MAX_REQUESTS = 20;
export const DEFAULT_MAX_BYTES = 5_000_000;
export const MAX_CURSOR_BYTES = 16_000_000;
export const DEFAULT_DEADLINE_MS = 30_000;
// Bounded digest of visited continuations carried in cursors so cycles
// across resumes end as partial, never complete.
export const MAX_SEEN = 64;

export interface SessionOperation {
  id: string;
  method: string;
  path: string;
  version: string;
  disposition: string;
  owningSlice: string | null;
}

export interface TransportRequest {
  method: "GET";
  url: string;
  headers: Record<string, string>;
  signal?: AbortSignal;
}

export interface TransportResponse {
  status: number;
  headers: Record<string, string>;
  body: string;
  receivedBodyBytes?: number;
}

// True external seam: tests substitute fixture transports, later slices wire
// real HTTPS. The transport sends what it is given; only the session attaches
// credentials, and only after the destination re-authorizes.
export type GraphTransport = (request: TransportRequest) => Promise<TransportResponse>;

// Fake-clock seam: production uses systemClock, tests inject a fake that
// advances instantly and still honors AbortSignal.
export interface Clock {
  now(): number;
  sleep(ms: number, signal?: AbortSignal): Promise<void>;
}

export const systemClock: Clock = {
  now: () => Date.now(),
  sleep: (ms, signal) => (signal ? delay(ms, undefined, { signal }) : delay(ms)) as Promise<void>,
};

export interface ExecuteArgs {
  profile: AnyProfile;
  operation: SessionOperation;
  params?: Record<string, string>;
  query?: Record<string, string>;
  // Delegated callers declare explicit Graph scopes; application profiles use
  // the configured .default audience and reject caller scopes.
  scopes?: string[];
  // Advanced queries needing eventual consistency declare it explicitly;
  // $search and $count=true fail closed without it.
  consistencyLevel?: "eventual";
  signal?: AbortSignal;
  clock?: Clock;
}

export interface CollectionBudget {
  maxRequests?: number;
  maxBytes?: number;
  deadlineMs?: number;
}

export interface CollectArgs extends ExecuteArgs {
  // Row cap: the remainder of an already fetched page is buffered into the
  // cursor, never discarded.
  limit?: number;
  budget?: CollectionBudget;
  // Opaque resume from a previous partial result.
  cursor?: string;
}

export interface CollectionResult {
  value: unknown[];
  query: Record<string, string>;
  complete: boolean;
  reason?: string;
  cursor?: string;
  requests: number;
  bytes: number;
}

let operationsById: Map<string, SessionOperation> | undefined;

function findOperation(id: string): SessionOperation | undefined {
  if (!operationsById) {
    const inventory = JSON.parse(readFileSync(new URL("../inventory/operations.json", import.meta.url), "utf8")) as {
      operations: SessionOperation[];
    };
    operationsById = new Map(inventory.operations.map(row => [row.id, row]));
  }
  return operationsById.get(id);
}

export function resolveSessionOperation(version: string, method: string, route: string): SessionOperation {
  const id = `${version}:${method}:${route}`;
  const operation = findOperation(id);
  if (!operation) {
    throw new AxiError(`Unknown catalogued Graph operation ${id}`, "VALIDATION_ERROR", [
      "Use a route from inventory/operations.json; the session never builds uncatalogued URLs",
    ]);
  }
  return { ...operation };
}

const SECRET_KEY = /(password|passwd|secret|token|credential|sas|authorization|accountkey)/i;
const KEY_SUFFIX = /[a-z0-9](key|keys)$/i;
const SECRET_VALUE = [
  /AccountKey=/i,
  /SharedAccessKey/i,
  /SharedAccessSignature/i,
  /[?&]sig=/i,
  /-----BEGIN/,
  /\beyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]*/,
];

function secretKey(key: string): boolean {
  if (key === "customAccountResetCredentialsUrl" || key === "customForgotMyPasswordText") return false;
  const name = key.toLowerCase().replace(/[-_]/g, "");
  return name === "sas" || name === "authorization" || SECRET_KEY.test(name) || KEY_SUFFIX.test(name);
}

// Shared redaction for reads, mutation previews and audit metadata: secret
// key names and sentinel values become the marker before buffering, output
// or journaling. WRITE-00 reuses this so previews never leak secrets.
export function redactGraphValue(value: unknown): unknown {
  return redact(value);
}

function redact(value: unknown): unknown {
  if (typeof value === "string") return SECRET_VALUE.some(pattern => pattern.test(value)) ? REDACTED : value;
  if (Array.isArray(value)) return value.map(redact);
  if (value === null || typeof value !== "object") return value;
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>).map(([key, child]) => {
      if (key === "keyCredentials" || key === "passwordCredentials") {
        const entries = Array.isArray(child) ? child : [];
        return [key, entries
          .filter(entry => entry !== null && typeof entry === "object" && !Array.isArray(entry))
          .map(entry => redact(Object.fromEntries(SAFE_CREDENTIAL_FIELDS
            .filter(field => Object.hasOwn(entry, field))
            .map(field => [field, entry[field]]))))];
      }
      return [key, typeof child === "string" && (secretKey(key) || key === "phoneNumber") ? REDACTED : redact(child)];
    }),
  );
}

function splitPath(path: string): string[] {
  return path.split("/").filter(segment => segment.length > 0);
}

function bindingName(segment: string): string | null {
  return /^\{([^{}]+)\}$/.exec(segment)?.[1] ?? null;
}

// Match an authorized path against the operation route template, extracting
// the encoded binding for each placeholder. Literals compare
// case-insensitively, matching Graph routing; bindings stay exact.
function matchRoute(route: string, pathname: string): Record<string, string> | null {
  const template = splitPath(route);
  const actual = splitPath(pathname);
  if (template.length !== actual.length) return null;
  const params: Record<string, string> = Object.create(null);
  for (let i = 0; i < template.length; i++) {
    const slot = template[i]!;
    const segment = actual[i]!;
    const name = bindingName(slot);
    if (!name && /[{}]/.test(slot)) return null;
    if (name) params[name] = segment;
    else if (slot.toLowerCase() !== segment.toLowerCase()) return null;
  }
  return params;
}

// Initial sensitive-area map. Mail and file content stay denied unless the
// profile explicitly enables the area; consent and directory metadata are not
// content areas. Later slices refine this table with their reviewed contracts.
function sensitiveArea(route: string): string | null {
  const segments = splitPath(route).map(segment => segment.toLowerCase());
  if (segments.some(segment => segment === "messages" || segment === "mailfolders")) return "mail";
  if (segments.some(segment => segment === "drive" || segment === "drives" || segment === "file" || segment === "files") || segments[0] === "shares") return "files";
  return null;
}

function checkOperation(operation: SessionOperation): void {
  if (
    !operation ||
    typeof operation.id !== "string" ||
    typeof operation.method !== "string" ||
    typeof operation.path !== "string" ||
    !operation.path.startsWith("/") ||
    typeof operation.version !== "string" ||
    typeof operation.disposition !== "string" ||
    !(typeof operation.owningSlice === "string" || operation.owningSlice === null)
  ) {
    throw new AxiError("Graph operations must be catalogued inventory rows", "VALIDATION_ERROR", [
      "Resolve the operation from inventory/operations.json before executing it",
    ]);
  }
}

const UNSAFE_PARAM = /[%/?\\{}]/;
const CONTROL = /[\s\x00-\x1f\x7f]/;

export function encodeGraphPathSegment(value: string): string {
  if (typeof value !== "string" || !value.length || value === "." || value === ".." || UNSAFE_PARAM.test(value) || CONTROL.test(value)) {
    throw new AxiError("Invalid path parameter", "VALIDATION_ERROR", [
      "Bind one resource identifier per placeholder; encoded separators and traversal are rejected",
    ]);
  }
  if (value.startsWith("$")) {
    throw new AxiError("OData reserved segments cannot be resource identifiers", "VALIDATION_ERROR", [
      "Bind a resource identifier that does not start with $; $count, $value and $ref are reserved route segments",
    ]);
  }
  return encodeURIComponent(value);
}

function buildPath(route: string, params: Record<string, string>): string {
  const used = new Set<string>();
  const path = splitPath(route)
    .map(segment => {
      const name = bindingName(segment);
      if (!name) {
        if (/[{}]/.test(segment)) {
          throw new AxiError(`Unsupported path template ${route}`, "VALIDATION_ERROR", [
            "Use a catalogued route with whole-segment placeholders",
          ]);
        }
        return segment;
      }
      if (!Object.hasOwn(params, name)) {
        throw new AxiError(`Missing path parameter {${name}} for ${route}`, "VALIDATION_ERROR", [
          "Bind one resource identifier per placeholder from the catalogued route",
        ]);
      }
      used.add(name);
      const value = params[name]!;
      return encodeGraphPathSegment(value);
    })
    .join("/");
  const extra = Object.keys(params).find(name => !used.has(name) && splitPath(route).every(segment => bindingName(segment) !== name));
  if (extra) throw new AxiError(`Unknown path parameter {${extra}} for ${route}`, "VALIDATION_ERROR", ["Bind only the placeholders in the catalogued route"]);
  return path;
}

function checkQueryEntry(key: string, value: string): void {
  if (!ALLOWED_KEYS.has(key)) {
    throw new AxiError(`Unsupported query key ${key}`, "VALIDATION_ERROR", [
      `Supported query keys: ${ALLOWED_QUERY_KEYS.join(", ")}`,
    ]);
  }
  if (typeof value !== "string" || !value.length || value.length > MAX_QUERY_VALUE || /[\x00-\x1f\x7f]/.test(value)) {
    throw new AxiError(`Invalid query value for ${key}`, "VALIDATION_ERROR", ["Keep query values short printable strings"]);
  }
}

function buildQuery(query: Record<string, string>): string {
  const pairs = Object.entries(query).map(([key, value]) => {
    checkQueryEntry(key, value);
    return `${encodeURIComponent(key)}=${encodeURIComponent(value)}`;
  });
  pairs.sort();
  return pairs.length ? `?${pairs.join("&")}` : "";
}

function denied(operation: SessionOperation, reason: string): AxiError {
  return new AxiError(`Graph continuation for ${operation.id} is denied: ${reason}`, "POLICY_DENIED", [
    "Redirects and continuations re-authorize as the same catalogued operation before credentials are attached",
    `Expected route ${operation.path} on https://${GRAPH_HOST}/${operation.version} with the bound resource and supported query keys`,
  ]);
}

// Re-authorize a redirect or nextLink target as the same catalogued operation:
// same commercial host, same version, same route template, identical bound
// resource values and only allowed query keys. Returns the canonical URL.
// CORE-02 reuses this when following @odata.nextLink pages.
export function authorizeUrl(operation: SessionOperation, params: Record<string, string>, raw: string, base?: string, consistencyLevel?: "eventual"): string {
  let url: URL;
  try {
    url = new URL(raw, base ?? `https://${GRAPH_HOST}/${operation.version}/`);
  } catch {
    throw denied(operation, "the target is not a parsable URL");
  }
  if (url.protocol !== "https:" || url.hostname.toLowerCase() !== GRAPH_HOST || url.port || url.username || url.password) {
    throw denied(operation, "the target leaves the commercial Graph host");
  }
  const prefix = `/${operation.version}/`;
  if (!url.pathname.startsWith(prefix)) throw denied(operation, `the target leaves version ${operation.version}`);
  const bindings = matchRoute(operation.path, url.pathname.slice(prefix.length));
  if (!bindings) throw denied(operation, `the target leaves the catalogued route ${operation.path}`);
  for (const [name, actual] of Object.entries(bindings)) {
    const expected = params[name];
    if (typeof expected !== "string" || actual.toLowerCase() !== encodeGraphPathSegment(expected).toLowerCase()) {
      throw denied(operation, `the target changes the bound resource {${name}}`);
    }
  }
  const queryKeys = new Set<string>();
  for (const [key, value] of url.searchParams) {
    if (queryKeys.has(key)) throw denied(operation, `the target repeats query ${key}`);
    queryKeys.add(key);
    try {
      checkQueryEntry(key, value);
    } catch {
      throw denied(operation, `the target carries unsupported query ${key}`);
    }
  }
  checkQueryContext(Object.fromEntries(url.searchParams), consistencyLevel);
  url.hash = "";
  return url.toString();
}

function header(headers: Record<string, string>, name: string): string | undefined {
  for (const [key, value] of Object.entries(headers)) if (key.toLowerCase() === name) return value;
  return undefined;
}

function faultBody(body: string): string {
  if (!body) return "";
  try {
    const parsed = redact(JSON.parse(body)) as { error?: { code?: unknown; message?: unknown } };
    const code = typeof parsed?.error?.code === "string" ? parsed.error.code : null;
    if (!code) return "";
    const message = typeof parsed?.error?.message === "string" ? parsed.error.message.slice(0, 300) : "";
    return `: graph ${code}${message ? ` ${message}` : ""}`;
  } catch {
    return "";
  }
}

function checkQueryContext(query: Record<string, string>, consistencyLevel: "eventual" | undefined): void {
  if (consistencyLevel !== undefined && consistencyLevel !== "eventual") {
    throw new AxiError(`Unsupported consistency ${consistencyLevel}`, "VALIDATION_ERROR", [
      "Advanced queries use ConsistencyLevel eventual or no header",
    ]);
  }
  const needsEventual = Object.hasOwn(query, "$search") || String(query["$count"] ?? "").toLowerCase() === "true";
  if (Object.hasOwn(query, "$expand")) {
    if (needsEventual) {
      throw new AxiError("Advanced queries cannot use $expand", "VALIDATION_ERROR", ["Use $expand separately from $search and $count=true"]);
    }
    for (const relationship of query["$expand"]!.split(",")) {
      const path = relationship.trim();
      if (!/^[a-zA-Z][a-zA-Z0-9]*(\/[a-zA-Z][a-zA-Z0-9]*)*$/.test(path)) {
        throw new AxiError("Unsupported $expand relationship", "VALIDATION_ERROR", ["Expand explicit relationship paths without nested query options"]);
      }
      if (sensitiveArea(path)) {
        throw new AxiError("Sensitive relationships cannot be expanded", "POLICY_DENIED", ["Mail and file content stay disabled for this profile"]);
      }
    }
  }
  if (needsEventual && consistencyLevel !== "eventual") {
    throw new AxiError("Advanced query needs ConsistencyLevel eventual", "VALIDATION_ERROR", [
      "Pass consistencyLevel eventual with $search or $count=true; unsupported combinations fail before credentials",
    ]);
  }
}

function checkLimit(limit: number | undefined): void {
  if (limit !== undefined && (!Number.isSafeInteger(limit) || limit <= 0)) {
    throw new AxiError(`Invalid row limit ${limit}`, "VALIDATION_ERROR", ["Pass a positive safe integer limit; the remainder of a fetched page is buffered, never discarded"]);
  }
}

function checkBudget(budget: CollectionBudget | undefined): { maxRequests: number; maxBytes: number; deadlineMs: number } {
  const maxRequests = budget?.maxRequests ?? DEFAULT_MAX_REQUESTS;
  const maxBytes = budget?.maxBytes ?? DEFAULT_MAX_BYTES;
  const deadlineMs = budget?.deadlineMs ?? DEFAULT_DEADLINE_MS;
  for (const [name, value] of [["maxRequests", maxRequests], ["maxBytes", maxBytes], ["deadlineMs", deadlineMs]] as const) {
    if (!Number.isSafeInteger(value) || value <= 0) throw new AxiError(`Invalid budget ${name} ${value}`, "VALIDATION_ERROR", ["Declare positive request/byte/deadline ceilings"]);
  }
  return { maxRequests, maxBytes, deadlineMs };
}

// Retry-After is seconds or an HTTP date; absent or unparseable means the
// caller decides (bounded backoff). Dates resolve through the fake clock.
function parseRetryAfter(raw: string | undefined, now: number): number | undefined {
  if (!raw) return undefined;
  const text = raw.trim();
  if (!text) return undefined;
  const seconds = Number(text);
  if (Number.isFinite(seconds)) return Math.max(0, seconds) * 1000;
  const at = Date.parse(text);
  if (Number.isNaN(at)) return undefined;
  return Math.max(0, at - now);
}

function backoffMs(attempt: number): number {
  return Math.min(1000 * 2 ** attempt, MAX_RETRY_AFTER_MS);
}

function digestUrl(url: string): string {
  return createHash("sha256").update(url).digest("hex").slice(0, 16);
}

class DeadlineExceeded extends Error {}

async function beforeDeadline<T>(action: (signal: AbortSignal) => Promise<T>, clock: Clock, deadline: number, callerSignal?: AbortSignal): Promise<T> {
  callerSignal?.throwIfAborted();
  const remaining = deadline - clock.now();
  if (remaining <= 0) throw new DeadlineExceeded();
  const controller = new AbortController();
  const timer = new AbortController();
  const signal = callerSignal ? AbortSignal.any([callerSignal, controller.signal]) : controller.signal;
  let onAbort: () => void;
  const aborted = new Promise<never>((_, reject) => {
    onAbort = () => reject(signal.reason);
    signal.addEventListener("abort", onAbort, { once: true });
  });
  const timeout = clock.sleep(remaining, timer.signal).then(() => {
    controller.abort(new DeadlineExceeded());
    throw signal.reason;
  });
  try {
    const result = await Promise.race([Promise.resolve().then(() => {
      signal.throwIfAborted();
      return action(signal);
    }), timeout, aborted]);
    signal.throwIfAborted();
    if (clock.now() >= deadline) throw new DeadlineExceeded();
    return result;
  } finally {
    timer.abort();
    signal.removeEventListener("abort", onAbort!);
  }
}

function bodyBytes(body: string): number {
  return Buffer.byteLength(body, "utf8");
}

interface CursorState {
  next?: string;
  buffered: unknown[];
  seen: string[];
  query: Record<string, string>;
  consistencyLevel?: "eventual";
  context: string;
  identity: string | null;
}

function encodeCursor(operation: SessionOperation, state: CursorState): string {
  const payload = { v: 3, op: operation.id, next: state.next ?? null, buffered: state.buffered, seen: state.seen.slice(-MAX_SEEN), query: state.query, consistencyLevel: state.consistencyLevel ?? null, context: state.context, identity: state.identity };
  return Buffer.from(JSON.stringify(payload), "utf8").toString("base64url");
}

function decodeCursor(operation: SessionOperation, cursor: string): CursorState {
  if (Buffer.byteLength(cursor, "utf8") > MAX_CURSOR_BYTES) {
    throw new AxiError(`Collection cursor exceeds ${MAX_CURSOR_BYTES} bytes`, "VALIDATION_ERROR", ["Use a cursor within the supported size ceiling"]);
  }
  let payload: unknown;
  try {
    payload = JSON.parse(Buffer.from(cursor, "base64url").toString("utf8"));
  } catch {
    throw new AxiError("Invalid collection cursor", "VALIDATION_ERROR", ["Resume with the cursor from a partial result of the same operation"]);
  }
  const record = payload as { v?: unknown; op?: unknown; next?: unknown; buffered?: unknown; seen?: unknown; query?: unknown; consistencyLevel?: unknown; context?: unknown; identity?: unknown };
  if (!record || typeof record !== "object" || record.v !== 3 || record.op !== operation.id || !(record.next === null || typeof record.next === "string") || !Array.isArray(record.buffered) || !Array.isArray(record.seen) || !record.query || typeof record.query !== "object" || Array.isArray(record.query) || !(record.consistencyLevel === null || record.consistencyLevel === "eventual") || typeof record.context !== "string" || !/^[0-9a-f]{64}$/.test(record.context) || !(record.identity === null || (typeof record.identity === "string" && /^[0-9a-f]{64}$/.test(record.identity)))) {
    throw new AxiError("Invalid collection cursor", "VALIDATION_ERROR", ["Resume with the cursor from a partial result of the same operation"]);
  }
  if (record.seen.length > MAX_SEEN || record.seen.some(entry => typeof entry !== "string" || !/^[0-9a-f]{16}$/.test(entry))) {
    throw new AxiError("Invalid collection cursor", "VALIDATION_ERROR", ["Resume with the cursor from a partial result of the same operation"]);
  }
  if (record.buffered.length > 5000) throw new AxiError("Invalid collection cursor", "VALIDATION_ERROR", ["The cursor buffers at most one fetched page"]);
  if (record.identity === null && (record.buffered.length > 0 || record.seen.length > 0)) throw new AxiError("Invalid collection cursor", "VALIDATION_ERROR", ["Fetched rows and consumed pages require a bound credential identity"]);
  const query = record.query as Record<string, string>;
  const consistencyLevel = record.consistencyLevel ?? undefined;
  buildQuery(query);
  checkQueryContext(query, consistencyLevel);
  return { next: record.next ?? undefined, buffered: record.buffered.map(redact), seen: [...record.seen], query, consistencyLevel, context: record.context, identity: record.identity };
}

function nextLinkOf(body: unknown): string | undefined {
  if (!body || typeof body !== "object" || Array.isArray(body)) return undefined;
  const value = (body as Record<string, unknown>)["@odata.nextLink"];
  return typeof value === "string" && value ? value : undefined;
}

function valuesOf(operation: SessionOperation, body: unknown): unknown[] {
  if (!body || typeof body !== "object" || Array.isArray(body) || !Object.hasOwn(body, "value")) {
    throw new AxiError(`Graph returned a malformed collection body for ${operation.id}`, "GRAPH_ERROR", ["Collections carry a value array; treat anything else as unknown, not empty"]);
  }
  const value = (body as { value?: unknown }).value;
  if (!Array.isArray(value)) throw new AxiError(`Graph returned a malformed collection body for ${operation.id}`, "GRAPH_ERROR", ["Collections carry a value array; treat anything else as unknown, not empty"]);
  return value;
}

export class GraphSession {
  constructor(
    private deps: {
      delegated: DelegatedAuth;
      application: ApplicationAuth;
      transport: GraphTransport;
    },
  ) {}

  cursorQuery(operation: SessionOperation, cursor: string): Readonly<Record<string, string>> {
    return decodeCursor(operation, cursor).query;
  }

  async execute(args: ExecuteArgs): Promise<unknown> {
    checkOperation(args.operation);
    // The inventory is authoritative: only the id is trusted from the caller,
    // so a fabricated disposition can never widen access.
    const operation = findOperation(args.operation.id);
    if (!operation) {
      throw new AxiError(`Unknown catalogued Graph operation ${args.operation.id}`, "VALIDATION_ERROR", [
        "Use a route from inventory/operations.json; the session never builds uncatalogued URLs",
      ]);
    }
    const profile = validateProfile(args.profile);
    this.authorizePolicy(profile, operation);
    const query = args.query ?? {};
    checkQueryContext(query, args.consistencyLevel);
    const params = args.params ?? {};
    const url = `https://${GRAPH_HOST}/${operation.version}/${buildPath(operation.path, params)}${buildQuery(query)}`;
    const clock = args.clock ?? systemClock;
    const signal = args.signal;
    const deadline = clock.now() + DEFAULT_DEADLINE_MS;
    let token: string;
    try {
      token = await beforeDeadline(async () => profile.mode === "delegated"
        ? (await this.deps.delegated.credential(validateDelegatedProfile(profile), this.readScopes(args.scopes))).token
        : (await this.deps.application.credential(validateApplicationProfile(profile), this.applicationScopes(args.scopes))).token, clock, deadline, signal);
    } catch (error) {
      signal?.throwIfAborted();
      if (error instanceof DeadlineExceeded) throw new AxiError(`Graph deadline for ${operation.id} is exceeded`, "GRAPH_ERROR", ["Narrow the query before retrying"]);
      throw error;
    }
    return this.send(operation, params, url, token, { signal, clock, deadline, consistencyLevel: args.consistencyLevel });
  }

  // CORE-02: follow @odata.nextLink continuations through the same
  // operation re-authorization and shared success redaction. Row caps never
  // discard fetched rows: overflow stays buffered in the opaque cursor.
  // Request/byte/deadline ceilings and continuation cycles end as truthful
  // partial results (complete:false), never as complete. Throttle waits and
  // deadlines resolve through the fake clock; signal aborts between
  // requests and while waiting.
  async collect(args: CollectArgs): Promise<CollectionResult> {
    checkOperation(args.operation);
    const operation = findOperation(args.operation.id);
    if (!operation) {
      throw new AxiError(`Unknown catalogued Graph operation ${args.operation.id}`, "VALIDATION_ERROR", [
        "Use a route from inventory/operations.json; the session never builds uncatalogued URLs",
      ]);
    }
    const profile = validateProfile(args.profile);
    this.authorizePolicy(profile, operation);
    checkLimit(args.limit);
    const budget = checkBudget(args.budget);
    if (args.cursor !== undefined && (typeof args.cursor !== "string" || !args.cursor)) {
      throw new AxiError("Invalid collection cursor", "VALIDATION_ERROR", ["Resume with the cursor from a partial result of the same operation"]);
    }
    const resumed = args.cursor === undefined ? undefined : decodeCursor(operation, args.cursor);
    const query = { ...resumed?.query, ...args.query };
    const consistencyLevel = args.consistencyLevel ?? resumed?.consistencyLevel;
    checkQueryContext(query, consistencyLevel);
    if (resumed && ((args.query !== undefined && buildQuery(query) !== buildQuery(resumed.query)) || (args.consistencyLevel !== undefined && args.consistencyLevel !== resumed.consistencyLevel))) {
      throw new AxiError("Resume arguments conflict with collection cursor context", "VALIDATION_ERROR", ["Resume with the cursor's original query and consistency level, or omit those arguments"]);
    }
    const params = args.params ?? {};
    const clock = args.clock ?? systemClock;
    const signal = args.signal;
    const path = buildPath(operation.path, params);
    const context = createHash("sha256").update(JSON.stringify([
      path, profile.mode, profile.tenantId.toLowerCase(), profile.clientId.toLowerCase(), profile.cloud,
      profile.credentialRef, [...new Set(args.scopes ?? [])].sort(),
    ])).digest("hex");
    if (resumed && resumed.context !== context) throw new AxiError("Collection cursor resource or authentication context does not match", "VALIDATION_ERROR", ["Resume under the original resource bindings, profile identity, credential reference and scopes"]);
    const startUrl = `https://${GRAPH_HOST}/${operation.version}/${path}${buildQuery(query)}`;
    const resumeUrl = resumed?.next === undefined ? undefined : authorizeUrl(operation, params, resumed.next, undefined, consistencyLevel);
    const deadline = clock.now() + budget.deadlineMs;
    let pending: unknown[] = resumed ? [...resumed.buffered] : [];
    let current: string | undefined = resumed ? resumeUrl : startUrl;
    const seenList: string[] = resumed ? [...resumed.seen] : [];
    const seen = new Set(seenList);
    const remember = (url: string): void => {
      const key = digestUrl(url);
      if (seen.has(key)) return;
      seen.add(key);
      seenList.push(key);
      if (seenList.length > MAX_SEEN) {
        for (const dropped of seenList.splice(0, seenList.length - MAX_SEEN)) seen.delete(dropped);
      }
    };
    const results: unknown[] = [];
    let requests = 0;
    let bytes = 0;
    let identity = resumed?.identity ?? null;
    const partial = (reason: string, next: string | undefined, buffered: unknown[]): CollectionResult => ({
      value: results,
      query,
      complete: false,
      reason,
      cursor: encodeCursor(operation, { next, buffered, seen: [...seenList], query, consistencyLevel, context, identity }),
      requests,
      bytes,
    });
    let token: string;
    try {
      const credential = await beforeDeadline(async () => profile.mode === "delegated"
        ? this.deps.delegated.credential(validateDelegatedProfile(profile), this.readScopes(args.scopes))
        : this.deps.application.credential(validateApplicationProfile(profile), this.applicationScopes(args.scopes)), clock, deadline, signal);
      const acquiredIdentity = createHash("sha256").update(JSON.stringify([
        profile.mode, credential.tenantId.toLowerCase(), credential.clientId.toLowerCase(),
        profile.mode === "delegated" && "accountId" in credential ? credential.accountId : null,
      ])).digest("hex");
      if (identity !== null && identity !== acquiredIdentity) throw new AxiError("Collection cursor credential identity does not match", "VALIDATION_ERROR", ["Resume with the original delegated account or application identity"]);
      identity = acquiredIdentity;
      token = credential.token;
    } catch (error) {
      signal?.throwIfAborted();
      if (error instanceof DeadlineExceeded) return partial("deadline exceeded", current, pending);
      throw error;
    }
    const headersFor = (): Record<string, string> => ({
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "client-request-id": randomUUID(),
      ...(consistencyLevel === "eventual" ? { ConsistencyLevel: "eventual" } : {}),
    });
    for (;;) {
      signal?.throwIfAborted();
      if (clock.now() >= deadline) return partial("deadline exceeded", current, pending);
      while (pending.length > 0 && (args.limit === undefined || results.length < args.limit)) results.push(pending.shift()!);
      if (clock.now() >= deadline) return partial("deadline exceeded", current, pending);
      if (args.limit !== undefined && results.length >= args.limit) {
        if (pending.length === 0 && !current) return { value: results, query, complete: true, requests, bytes };
        return partial("row limit reached; buffered remainder is preserved in the cursor", current, pending);
      }
      if (!current) return { value: results, query, complete: true, requests, bytes };
      if (requests >= budget.maxRequests) return partial(`request budget exhausted after ${requests} requests`, current, pending);
      const cycleKey = digestUrl(current);
      if (seen.has(cycleKey)) return partial("continuation cycle detected; result is partial, never complete", current, pending);
      let fetchUrl = current;
      const redirects = new Set([digestUrl(fetchUrl)]);
      let hops = 0;
      let throttleAttempt = 0;
      for (;;) {
        signal?.throwIfAborted();
        if (requests >= budget.maxRequests) return partial(`request budget exhausted after ${requests} requests`, fetchUrl, pending);
        if (clock.now() >= deadline) return partial("deadline exceeded", fetchUrl, pending);
        let response: TransportResponse;
        try {
          requests += 1;
          const requestUrl = new URL(fetchUrl);
          if (operation.path === "/users/{user-id}/authentication/methods") requestUrl.searchParams.delete("$select");
          response = await beforeDeadline(signal => this.deps.transport({ method: "GET", url: requestUrl.toString(), headers: headersFor(), signal }), clock, deadline, signal);
        } catch (error) {
          signal?.throwIfAborted();
          if (error instanceof DeadlineExceeded) return partial("deadline exceeded", fetchUrl, pending);
          throw new AxiError(`Graph request for ${operation.id} failed before a response was received`, "GRAPH_ERROR", [
            `Check network access to https://${GRAPH_HOST}`,
            "Transport failures carry no Graph diagnosis; do not retry blindly",
          ]);
        }
        bytes += response.receivedBodyBytes ?? bodyBytes(response.body ?? "");
        const status = response.status;
        if (status === 301 || status === 302 || status === 303 || status === 307 || status === 308) {
          const location = header(response.headers, "location");
          if (!location) throw new AxiError(`Graph redirect for ${operation.id} is missing its target`, "GRAPH_ERROR", ["A redirect without a target cannot be re-authorized"]);
          if (hops >= MAX_REDIRECTS) throw denied(operation, `the redirect exceeds ${MAX_REDIRECTS} hops`);
          const target = authorizeUrl(operation, params, location, fetchUrl, consistencyLevel);
          if (seen.has(digestUrl(target)) || redirects.has(digestUrl(target))) return partial("continuation cycle detected; result is partial, never complete", target, pending);
          redirects.add(digestUrl(target));
          fetchUrl = target;
          hops += 1;
          if (bytes > budget.maxBytes) return partial(`byte budget exceeded after ${bytes} bytes`, fetchUrl, pending);
          continue;
        }
        if (status === 429 || status === 503) {
          if (bytes > budget.maxBytes) return partial(`byte budget exceeded after ${bytes} bytes`, fetchUrl, pending);
          const raw = header(response.headers, "retry-after");
          let wait = parseRetryAfter(raw, clock.now());
          if (wait === undefined) wait = backoffMs(throttleAttempt);
          throttleAttempt += 1;
          const detail = raw ? `; retry after ${raw}` : "";
          if (requests >= budget.maxRequests) return partial(`request budget exhausted after ${requests} requests`, fetchUrl, pending);
          if (clock.now() + wait > deadline) return partial(`deadline exceeded; throttled (${status})${detail} would overrun the budget`, fetchUrl, pending);
          if (wait > MAX_RETRY_AFTER_MS) return partial(`throttled (${status})${detail}; budget preserves ${results.length + pending.length} rows`, fetchUrl, pending);
          await clock.sleep(wait, signal);
          continue;
        }
        if (status < 200 || status >= 300) throw this.translateError(operation, response);
        if (status === 204 || status === 205) {
          remember(current);
          remember(fetchUrl);
          current = undefined;
          if (bytes > budget.maxBytes) return partial(`byte budget exceeded after ${bytes} bytes`, current, pending);
          break;
        }
        if (!response.body) throw new AxiError(`Graph returned an empty success body for ${operation.id}`, "GRAPH_ERROR", ["Empty reads are malformed; treat the result as unknown, not empty"]);
        let parsed: unknown;
        try {
          parsed = redact(JSON.parse(response.body));
        } catch {
          throw new AxiError(`Graph returned a non-JSON success body for ${operation.id}`, "GRAPH_ERROR", ["Successful reads are JSON; anything else is malformed"]);
        }
        const rows = valuesOf(operation, parsed);
        const rawNext = nextLinkOf(parsed);
        const nextUrl = rawNext ? authorizeUrl(operation, params, rawNext, fetchUrl, consistencyLevel) : undefined;
        remember(current);
        remember(fetchUrl);
        current = nextUrl;
        pending = [...pending, ...rows];
        if (bytes > budget.maxBytes) return partial(`byte budget exceeded after ${bytes} bytes`, current, pending);
        break;
      }
    }
  }

  private authorizePolicy(profile: AnyProfile, operation: SessionOperation): void {
    if (operation.method !== "GET") {
      throw new AxiError(`Graph session executes reads; ${operation.id} is not a read`, "VALIDATION_ERROR", [
        "Named mutations ship with their own coordinator slice; the session never sends them",
      ]);
    }
    // Profile-scoped gates run before the catalogued verdict, so a route that
    // is both sensitive and excluded reports the sensitive denial.
    // Every current owning slice belongs to the entra pack; a profile without
    // it cannot execute. New packs extend this gate with their own review.
    if (!profile.enabledPacks.includes("entra")) {
      throw new AxiError("Pack entra is not enabled for this profile", "POLICY_DENIED", ["mg-axi profile show --profile <name>"]);
    }
    if (operation.version === "beta" && !profile.preview) {
      throw new AxiError(`Beta Graph reads require a preview-enabled profile (${operation.id})`, "POLICY_DENIED", [
        "Enable preview explicitly for beta reads; v1.0 stays the default with no fallback",
      ]);
    }
    const route = operation.path.toLowerCase();
    if (profile.mode === "application" && (route === "/me" || route.startsWith("/me/"))) {
      throw new AxiError("Application profiles cannot use /me; bind an explicit /users/{id} route", "POLICY_DENIED", [
        "App-only calls address the resource directly; /me needs a signed-in user",
      ]);
    }
    const area = sensitiveArea(operation.path);
    if (area && !profile.sensitiveAreas.includes(area)) {
      throw new AxiError(`Sensitive area ${area} is not enabled for this profile (${operation.id})`, "POLICY_DENIED", [
        "Mail and file content stay disabled unless explicitly enabled for the profile",
      ]);
    }
    if (!EXECUTABLE_DISPOSITIONS.has(operation.disposition)) {
      throw new AxiError(`Graph operation ${operation.id} is ${operation.disposition}; it cannot be executed`, "POLICY_DENIED", [
        "Blocked, deprecated and unavailable operations stay denied even when reached by redirect",
      ]);
    }
  }

  private readScopes(scopes: string[] | undefined): string[] {
    if (!scopes?.length || scopes.some(scope => !READ_SCOPES.has(scope))) {
      throw new AxiError("Graph reads require supported read scopes", "VALIDATION_ERROR", [
        `Supported read scopes: ${[...READ_SCOPES].join(", ")}`,
      ]);
    }
    return scopes;
  }

  private applicationScopes(scopes: string[] | undefined): string[] | undefined {
    if (scopes !== undefined) {
      throw new AxiError("Application profiles use the configured Graph .default audience; delegated scopes are unavailable", "VALIDATION_ERROR", [
        "mg-axi profile show --profile <name>",
      ]);
    }
    return undefined;
  }

  private async send(operation: SessionOperation, params: Record<string, string>, url: string, token: string, opts: { signal?: AbortSignal; clock: Clock; deadline: number; consistencyLevel?: "eventual" }): Promise<unknown> {
    const clock = opts.clock;
    const signal = opts.signal;
    const deadline = opts.deadline;
    let requests = 0;
    let current = url;
    let hops = 0;
    let throttleAttempt = 0;
    const visited = new Set([current]);
    const headersFor = (): Record<string, string> => ({
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "client-request-id": randomUUID(),
      ...(opts.consistencyLevel === "eventual" ? { ConsistencyLevel: "eventual" } : {}),
      // The branding GET documents Accept-Language as required; 0 selects
      // the default branding object. Locale variants come from the
      // localizations collection instead of this header.
      ...(operation.path === "/organization/{organization-id}/branding" ? { "Accept-Language": "0" } : {}),
    });
    for (;;) {
      signal?.throwIfAborted();
      if (requests >= DEFAULT_MAX_REQUESTS) throw new AxiError(`Graph request budget for ${operation.id} is exhausted after ${requests} requests`, "GRAPH_ERROR", ["Narrow the query before retrying"]);
      if (clock.now() >= deadline) throw new AxiError(`Graph deadline for ${operation.id} is exceeded`, "GRAPH_ERROR", ["Narrow the query before retrying"]);
      let response: TransportResponse;
      try {
        response = await beforeDeadline(signal => this.deps.transport({ method: "GET", url: current, headers: headersFor(), signal }), clock, deadline, signal);
      } catch (error) {
        signal?.throwIfAborted();
        if (error instanceof DeadlineExceeded) throw new AxiError(`Graph deadline for ${operation.id} is exceeded`, "GRAPH_ERROR", ["Narrow the query before retrying"]);
        throw new AxiError(`Graph request for ${operation.id} failed before a response was received`, "GRAPH_ERROR", [
          `Check network access to https://${GRAPH_HOST}`,
          "Transport failures carry no Graph diagnosis; do not retry blindly",
        ]);
      }
      requests += 1;
      if (response.status === 301 || response.status === 302 || response.status === 303 || response.status === 307 || response.status === 308) {
        const location = header(response.headers, "location");
        if (!location) throw new AxiError(`Graph redirect for ${operation.id} is missing its target`, "GRAPH_ERROR", ["A redirect without a target cannot be re-authorized"]);
        if (hops >= MAX_REDIRECTS) throw denied(operation, `the redirect exceeds ${MAX_REDIRECTS} hops`);
        const target = authorizeUrl(operation, params, location, current, opts.consistencyLevel);
        if (visited.has(target)) throw denied(operation, "the redirect loops");
        visited.add(target);
        current = target;
        hops += 1;
        throttleAttempt = 0;
        continue;
      }
      if (response.status === 429 || response.status === 503) {
        const raw = header(response.headers, "retry-after");
        let wait = parseRetryAfter(raw, clock.now());
        if (wait === undefined) wait = backoffMs(throttleAttempt);
        throttleAttempt += 1;
        if (wait > MAX_RETRY_AFTER_MS || requests >= DEFAULT_MAX_REQUESTS || clock.now() + wait > deadline) throw this.translateError(operation, response);
        await clock.sleep(wait, signal);
        continue;
      }
      const result = this.translate(operation, response);
      signal?.throwIfAborted();
      if (clock.now() >= deadline) throw new AxiError(`Graph deadline for ${operation.id} is exceeded`, "GRAPH_ERROR", ["Narrow the query before retrying"]);
      return result;
    }
  }

  private translate(operation: SessionOperation, response: TransportResponse): unknown {
    const status = response.status;
    if (status === 204 || status === 205) return null;
    if (status >= 200 && status < 300) {
      if (!response.body) throw new AxiError(`Graph returned an empty success body for ${operation.id}`, "GRAPH_ERROR", ["Empty reads are malformed; treat the result as unknown, not empty"]);
      try {
        return redact(JSON.parse(response.body));
      } catch {
        throw new AxiError(`Graph returned a non-JSON success body for ${operation.id}`, "GRAPH_ERROR", ["Successful reads are JSON; anything else is malformed"]);
      }
    }
    throw this.translateError(operation, response);
  }

  private translateError(operation: SessionOperation, response: TransportResponse): AxiError {
    const status = response.status;
    const fault = faultBody(response.body);
    if (status === 401) {
      return new AxiError(`Graph rejected the credential for ${operation.id} (401)${fault}`, "AUTH_REQUIRED", [
        "mg-axi login --profile <name> --scopes <comma-separated-Graph-scopes>",
        "For application profiles, ask an administrator to grant application consent for the Graph .default audience",
      ]);
    }
    if (status === 403) {
      return new AxiError(
        `Graph denied ${operation.id} (403)${fault}: the grant, role, licence or policy prerequisite is missing and Graph does not say which`,
        "GRAPH_ERROR",
        [
          "Check the catalogue permission and role guidance for this operation",
          "Confirm tenant licensing for the generating feature",
          "A 403 never proves which prerequisite is missing; do not retry blindly",
        ],
      );
    }
    if (status === 404) {
      return new AxiError(`Graph reports ${operation.id} as not found or inaccessible (404)${fault}; Graph does not distinguish missing from hidden`, "GRAPH_ERROR", [
        "Verify the bound identifier; absence is not proof of nonexistence",
      ]);
    }
    if (status === 429 || status === 503) {
      const retryAfter = header(response.headers, "retry-after");
      return new AxiError(`Graph throttled ${operation.id} (${status})${retryAfter ? `; retry after ${retryAfter}` : ""}${fault}`, "GRAPH_ERROR", [
        "Bounded safe-read retries honor Retry-After within request and deadline budgets",
        "Narrow the query when throttling persists",
      ]);
    }
    return new AxiError(`Graph failed ${operation.id} (${status})${fault}`, "GRAPH_ERROR", ["Server failures may be transient; retry safe reads within a bounded budget"]);
  }
}
