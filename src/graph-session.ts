import { randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
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
export const GRAPH_HOST = "graph.microsoft.com";
// Conservative read-query allowlist. Per-operation review (READ slices) can
// extend it; unknown keys fail closed here.
export const ALLOWED_QUERY_KEYS: readonly string[] = ["$select", "$filter", "$top", "$orderby", "$count", "$skiptoken"];
const ALLOWED_KEYS = new Set(ALLOWED_QUERY_KEYS);
// Only these dispositions may execute. `scheduled` stays executable at this
// layer so READ slices need no session change; catalogue/CLI gating still
// keeps unshipped commands out of reach. Everything else fails closed.
const EXECUTABLE_DISPOSITIONS = new Set(["named-command", "reviewed-raw-read", "scheduled"]);
const MAX_REDIRECTS = 3;
const MAX_QUERY_VALUE = 1024;

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
}

export interface TransportResponse {
  status: number;
  headers: Record<string, string>;
  body: string;
}

// True external seam: tests substitute fixture transports, later slices wire
// real HTTPS. The transport sends what it is given; only the session attaches
// credentials, and only after the destination re-authorizes.
export type GraphTransport = (request: TransportRequest) => Promise<TransportResponse>;

export interface ExecuteArgs {
  profile: AnyProfile;
  operation: SessionOperation;
  params?: Record<string, string>;
  query?: Record<string, string>;
  // Delegated callers declare explicit Graph scopes; application profiles use
  // the configured .default audience and reject caller scopes.
  scopes?: string[];
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
  const name = key.toLowerCase().replace(/[-_]/g, "");
  return name === "sas" || name === "authorization" || SECRET_KEY.test(name) || KEY_SUFFIX.test(name);
}

function redact(value: unknown): unknown {
  if (typeof value === "string") return SECRET_VALUE.some(pattern => pattern.test(value)) ? REDACTED : value;
  if (Array.isArray(value)) return value.map(redact);
  if (value === null || typeof value !== "object") return value;
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>).map(([key, child]) => [
      key,
      typeof child === "string" && secretKey(key) ? REDACTED : redact(child),
    ]),
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

const UNSAFE_PARAM = /[%/?#\\]/;
const CONTROL = /[\s\x00-\x1f\x7f]/;

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
      if (typeof value !== "string" || !value.length || value === "." || value === ".." || UNSAFE_PARAM.test(value) || CONTROL.test(value)) {
        throw new AxiError(`Invalid path parameter {${name}}`, "VALIDATION_ERROR", [
          "Bind one resource identifier per placeholder; encoded separators and traversal are rejected",
        ]);
      }
      return encodeURIComponent(value);
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
export function authorizeUrl(operation: SessionOperation, params: Record<string, string>, raw: string, base?: string): string {
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
    if (typeof expected !== "string" || actual.toLowerCase() !== encodeURIComponent(expected).toLowerCase()) {
      throw denied(operation, `the target changes the bound resource {${name}}`);
    }
  }
  for (const [key, value] of url.searchParams) {
    try {
      checkQueryEntry(key, value);
    } catch {
      throw denied(operation, `the target carries unsupported query ${key}`);
    }
  }
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

export class GraphSession {
  constructor(
    private deps: {
      delegated: DelegatedAuth;
      application: ApplicationAuth;
      transport: GraphTransport;
    },
  ) {}

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
    const params = args.params ?? {};
    const url = `https://${GRAPH_HOST}/${operation.version}/${buildPath(operation.path, params)}${buildQuery(args.query ?? {})}`;
    const token =
      profile.mode === "delegated"
        ? (await this.deps.delegated.credential(validateDelegatedProfile(profile), args.scopes ?? [])).token
        : (await this.deps.application.credential(validateApplicationProfile(profile), this.applicationScopes(args.scopes))).token;
    return this.send(operation, params, url, token);
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

  private applicationScopes(scopes: string[] | undefined): string[] | undefined {
    if (scopes !== undefined) {
      throw new AxiError("Application profiles use the configured Graph .default audience; delegated scopes are unavailable", "VALIDATION_ERROR", [
        "mg-axi profile show --profile <name>",
      ]);
    }
    return undefined;
  }

  private async send(operation: SessionOperation, params: Record<string, string>, url: string, token: string): Promise<unknown> {
    const visited = new Set<string>();
    let current = url;
    for (let hop = 0; ; hop++) {
      if (visited.has(current.toLowerCase())) throw denied(operation, "the redirect loops");
      visited.add(current.toLowerCase());
      if (hop > MAX_REDIRECTS) throw denied(operation, `the redirect exceeds ${MAX_REDIRECTS} hops`);
      let response: TransportResponse;
      try {
        response = await this.deps.transport({
          method: "GET",
          url: current,
          headers: { Authorization: `Bearer ${token}`, Accept: "application/json", "client-request-id": randomUUID() },
        });
      } catch {
        throw new AxiError(`Graph request for ${operation.id} failed before a response was received`, "GRAPH_ERROR", [
          `Check network access to https://${GRAPH_HOST}`,
          "Transport failures carry no Graph diagnosis; do not retry blindly",
        ]);
      }
      if (response.status === 301 || response.status === 302 || response.status === 303 || response.status === 307 || response.status === 308) {
        const location = header(response.headers, "location");
        if (!location) throw new AxiError(`Graph redirect for ${operation.id} is missing its target`, "GRAPH_ERROR", ["A redirect without a target cannot be re-authorized"]);
        current = authorizeUrl(operation, params, location, current);
        continue;
      }
      return this.translate(operation, response);
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
    const fault = faultBody(response.body);
    if (status === 401) {
      throw new AxiError(`Graph rejected the credential for ${operation.id} (401)${fault}`, "AUTH_REQUIRED", [
        "mg-axi login --profile <name> --scopes <comma-separated-Graph-scopes>",
        "For application profiles, ask an administrator to grant application consent for the Graph .default audience",
      ]);
    }
    if (status === 403) {
      throw new AxiError(
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
      throw new AxiError(`Graph reports ${operation.id} as not found or inaccessible (404)${fault}; Graph does not distinguish missing from hidden`, "GRAPH_ERROR", [
        "Verify the bound identifier; absence is not proof of nonexistence",
      ]);
    }
    if (status === 429 || status === 503) {
      const retryAfter = header(response.headers, "retry-after");
      throw new AxiError(`Graph throttled ${operation.id} (${status})${retryAfter ? `; retry after ${retryAfter}` : ""}${fault}`, "GRAPH_ERROR", [
        "Back off for the advertised delay before retrying",
        "Bounded safe-read retries ship with CORE-02; until then never retry blindly",
      ]);
    }
    throw new AxiError(`Graph failed ${operation.id} (${status})${fault}`, "GRAPH_ERROR", ["Server failures may be transient; bounded retries ship with CORE-02"]);
  }
}
