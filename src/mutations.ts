import { closeSync, existsSync, fsyncSync, mkdirSync, openSync, readFileSync, unlinkSync, writeFileSync } from "node:fs";
import { createHash, randomUUID } from "node:crypto";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { AxiError } from "axi-sdk-js";
import { ApplicationAuth } from "./app-auth.js";
import { DelegatedAuth } from "./auth.js";
import { encodeGraphPathSegment, GRAPH_HOST, redactGraphValue } from "./graph-session.js";
import { validateApplicationProfile, validateDelegatedProfile, type AnyProfile } from "./profiles.js";

// Shared mutation coordinator for reviewed named mutation families.
// The gate order follows az-axi's write gates as the reference: read-only default,
// allowWrites plus a scope allowlist, preview, --execute, --confirm,
// --if-match and a durable journal. The shared read-only
// scope guard in graph-session.ts stays in force for reads; write scopes are
// requested only through this coordinator path.

// Forced read-only beats every profile, mirroring AZ_AXI_READ_ONLY.
export const READ_ONLY_ENV = "MG_AXI_READ_ONLY";
export const WRITE_LOG_ENV = "MG_AXI_WRITE_LOG";

export function readOnlyForced(env: NodeJS.ProcessEnv = process.env): boolean {
  return env[READ_ONLY_ENV] === "1";
}

// Mutation intents and outcomes are journaled here: MG_AXI_WRITE_LOG when
// set, otherwise ~/.mg-axi/writes.log. A blank override falls back.
export function resolveWriteLogPath(env: NodeJS.ProcessEnv = process.env): string {
  const override = env[WRITE_LOG_ENV]?.trim();
  if (override) return override;
  return join(homedir(), ".mg-axi", "writes.log");
}

export type MutationMethod = "POST" | "PUT" | "PATCH" | "DELETE";
export type MutationEffect = "write" | "disruptive";

// Named handlers bind reviewed mutations to this shape; the operation must
// fall inside the profile's configured scope.
export type MutationDefinition = {
  operation: string;
  method: MutationMethod;
  /** Graph version the path is bound to; only v1.0 mutates, beta fails closed. */
  version: string;
  /** Literal server-relative path, already including the version segment. */
  path: string;
  effect: MutationEffect;
  target: string;
  payload?: unknown;
};

export type MutationPreview = {
  operation: string;
  method: MutationMethod;
  url: string;
  effect: MutationEffect;
  target: string;
  noop: boolean;
  currentState: string | null;
  proposedChange: string | null;
};

export type MutationExecuteOptions = {
  execute?: boolean;
  confirm?: string;
  ifMatch?: string;
  /** Delegated profiles resolve credentials with these explicit Graph scopes. */
  scopes?: string[];
  readState: () => unknown | Promise<unknown>;
  isNoop?: (current: unknown) => boolean;
  intentId?: string;
};

export type MutationResult =
  | { kind: "dry-run"; preview: MutationPreview }
  | { kind: "noop"; preview: MutationPreview }
  | { kind: "success"; preview: MutationPreview; auditId: string; status: number; response: unknown }
  | { kind: "failed"; preview: MutationPreview; auditId: string; status: number }
  | { kind: "unknown"; preview: MutationPreview; auditId: string; httpStatus: number; guidance: string };

// The immutable tenant and operation scope: copied from the hand-edited
// profile when the coordinator is created. Read flags, environment profile
// overrides and later raw-read access never widen it because nothing after
// creation feeds it. Callers see only this frozen snapshot; enforcement
// state stays private to the coordinator and sender.
export type WriteScope = Readonly<{
  tenantId: string;
  allowWrites: boolean;
  operations: readonly string[];
}>;

export type MutationCoordinator = {
  readonly scope: WriteScope;
  preview(definition: MutationDefinition, current?: unknown, isNoop?: (current: unknown) => boolean): MutationPreview;
  execute(definition: MutationDefinition, options: MutationExecuteOptions): Promise<MutationResult>;
};

// Mutation-only transport seam: the read GraphTransport carries GET only, so
// mutations travel on this separate type. Tests substitute fixture
// transports; no production mutation transport ships in WRITE-00.
export interface MutationTransportRequest {
  method: MutationMethod;
  url: string;
  headers: Record<string, string>;
  body?: string;
  signal?: AbortSignal;
}

export interface MutationTransportResponse {
  status: number;
  headers: Record<string, string>;
  body: string;
}

export type MutationTransport = (request: MutationTransportRequest) => Promise<MutationTransportResponse>;

// Coordinator-minted authorization: an unforgeable capability the sender
// consumes once. There is no self-service authorize on the sender; anything
// else is refused before credentials resolve and before any HTTP call.
export type MutationAuthorization = {
  readonly nonce: string;
  readonly method: MutationMethod;
  readonly url: string;
};

export type MutationResponse = { status: number; body: unknown };

const authorizations = new WeakMap<MutationAuthorization, object>();

function consumeMutationAuthorization(authorization: MutationAuthorization, sender: object): boolean {
  if (authorizations.get(authorization) !== sender) return false;
  authorizations.delete(authorization);
  return true;
}

function assertMutationDestination(url: string): void {
  let destination: URL;
  try {
    destination = new URL(url);
  } catch {
    throw new AxiError("Refusing unparseable mutation destination", "DESTINATION_DENIED", [
      "Mutation destinations must be absolute HTTPS URLs on the commercial Graph host; no credential was sent",
    ]);
  }
  if (destination.protocol !== "https:" || destination.hostname.toLowerCase() !== GRAPH_HOST
    || destination.port || destination.username || destination.password
    || !destination.pathname.startsWith("/v1.0/")) {
    throw new AxiError("Refusing mutation destination outside the commercial Graph host", "DESTINATION_DENIED", [
      `Mutations must stay under https://${GRAPH_HOST}/v1.0/; no credential was sent`,
    ]);
  }
}

function mutationFailure(
  message: string, code: string, suggestions: string[], details: { httpStatus: number; accepted?: boolean },
): AxiError {
  return Object.assign(new AxiError(message, code, suggestions), { details });
}

export function mutationHttpStatus(error: unknown): number {
  if (error && typeof error === "object" && "details" in error) {
    const details = (error as { details?: unknown }).details;
    if (details && typeof details === "object" && "httpStatus" in details) {
      const status = (details as { httpStatus?: unknown }).httpStatus;
      if (typeof status === "number" && Number.isInteger(status) && status > 0) return status;
    }
  }
  return 0;
}

export function mutationAccepted(error: unknown): boolean {
  if (error && typeof error === "object" && "details" in error) {
    const details = (error as { details?: unknown }).details;
    if (details && typeof details === "object" && "accepted" in details) {
      return (details as { accepted?: unknown }).accepted === true;
    }
  }
  return false;
}

const mutationsNotSent = new WeakSet<object>();

export function mutationNotSent(error: unknown): boolean {
  return error instanceof Error && mutationsNotSent.has(error);
}

function decodeMutationBody(response: MutationTransportResponse): unknown {
  if (response.status === 401) {
    throw mutationFailure(`Graph rejected the mutation credential (401)`, "AUTH_REQUIRED", [
      "mg-axi login --profile <name> --scopes <comma-separated-Graph-scopes>",
      "For application profiles, ask an administrator to grant application consent for the Graph .default audience",
    ], { httpStatus: response.status });
  }
  if (response.status === 403) {
    throw mutationFailure(`Graph denied the mutation (403): the grant, role, licence or policy prerequisite is missing and Graph does not say which`,
      "GRAPH_ERROR", [
        "Check the catalogue permission and role guidance for this operation",
        "A 403 never proves which prerequisite is missing; do not retry blindly",
      ], { httpStatus: response.status });
  }
  if (response.status === 404) {
    throw mutationFailure(`Graph reports the mutation target as not found or inaccessible (404); Graph does not distinguish missing from hidden`,
      "GRAPH_ERROR", ["Verify the bound identifier; absence is not proof of nonexistence"], { httpStatus: response.status });
  }
  if (response.status < 200 || response.status > 299) {
    throw mutationFailure(`Graph mutation returned status ${response.status}`, "GRAPH_ERROR", [
      "Read back the target before doing anything else; never replay this intent",
    ], { httpStatus: response.status });
  }
  try {
    return response.body ? JSON.parse(response.body) as unknown : {};
  } catch {
    // The server accepted the mutation; only its response body was discarded.
    throw mutationFailure("Graph mutation response is not valid JSON", "RESPONSE_INVALID",
      ["The server accepted the mutation but its response body was discarded"],
      { httpStatus: response.status, accepted: true });
  }
}

// The authorized mutation sender. It shares credential providers and the
// transport seam with reads while keeping mutation policy explicit: every
// send needs a coordinator authorization minted after the write gates pass.
// The bound profile copy is private; no profile snapshot is exposed.
export function createMutationSender(args: {
  profile: AnyProfile;
  delegated: DelegatedAuth;
  application: ApplicationAuth;
  transport: MutationTransport;
}): {
  send(authorization: MutationAuthorization, options?: { body?: string; ifMatch?: string; signal?: AbortSignal; scopes?: string[] }): Promise<MutationResponse>;
} {
  const { delegated, application, transport } = args;
  const bound = args.profile.mode === "application" ? validateApplicationProfile(args.profile) : validateDelegatedProfile(args.profile);

  async function credential(scopes: string[] | undefined, signal?: AbortSignal): Promise<string> {
    signal?.throwIfAborted();
    if (bound.mode === "delegated") {
      if (scopes === undefined) {
        throw new AxiError("Delegated mutations require explicit Graph scopes", "VALIDATION_ERROR", [
          "Pass the operation's documented delegated permission using full https://graph.microsoft.com/ scope names",
        ]);
      }
      return (await delegated.credential(validateDelegatedProfile(bound), scopes)).token;
    }
    if (scopes !== undefined) {
      throw new AxiError("Application profiles use the configured Graph .default audience; delegated scopes are unavailable", "VALIDATION_ERROR", [
        "mg-axi profile show --profile <name>",
      ]);
    }
    return (await application.credential(validateApplicationProfile(bound))).token;
  }

  async function send(
    authorization: MutationAuthorization,
    options?: { body?: string; ifMatch?: string; signal?: AbortSignal; scopes?: string[] },
  ): Promise<MutationResponse> {
    let handedOff = false;
    try {
      if (!consumeMutationAuthorization(authorization, sender)) {
        throw new AxiError("Refusing mutation without coordinator authorization", "OPERATION_BLOCKED", [
          "Mutations are sent only with a coordinator authorization minted after the write gates pass; no credential was sent",
        ]);
      }
      if (!["POST", "PUT", "PATCH", "DELETE"].includes(authorization.method)) {
        throw new AxiError(`Refusing mutation with method: ${authorization.method}`, "OPERATION_BLOCKED", [
          "Mutations use POST, PUT, PATCH or DELETE only",
        ]);
      }
      // Re-check the bound destination after authorization, before credentials.
      assertMutationDestination(authorization.url);
      if (options?.ifMatch !== undefined && (!options.ifMatch.trim() || /[\x00-\x1f\x7f]/.test(options.ifMatch))) {
        throw new AxiError("Invalid If-Match value", "VALIDATION_ERROR", [
          "Provide the exact entity tag returned by the previewed read",
        ]);
      }
      options?.signal?.throwIfAborted();
      if (readOnlyForced()) {
        throw new AxiError("blocked: forced read-only disables mutations", "WRITES_DISABLED", [
          "MG_AXI_READ_ONLY=1 overrides profile write enablement",
        ]);
      }
      // Provider checks bind the acquired credential to the snapshotted
      // tenant and client; a credential for another identity fails here and
      // nothing is sent.
      const token = await credential(options?.scopes, options?.signal);
      options?.signal?.throwIfAborted();
      assertMutationDestination(authorization.url);
      if (readOnlyForced()) {
        throw new AxiError("blocked: forced read-only disables mutations", "WRITES_DISABLED", [
          "MG_AXI_READ_ONLY=1 overrides profile write enablement",
        ]);
      }
      // The request is handed to the transport here: any later failure is
      // ambiguous (bytes may have gone out) and is never replayed.
      handedOff = true;
      const response = await transport({
        method: authorization.method,
        url: authorization.url,
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          "client-request-id": randomUUID(),
          ...(options?.body !== undefined ? { "Content-Type": "application/json" } : {}),
          ...(options?.ifMatch !== undefined ? { "If-Match": options.ifMatch } : {}),
        },
        ...(options?.body !== undefined ? { body: options.body } : {}),
        ...(options?.signal ? { signal: options.signal } : {}),
      });
      options?.signal?.throwIfAborted();
      return { status: response.status, body: decodeMutationBody(response) };
    } catch (error) {
      if (!handedOff && error instanceof Error) mutationsNotSent.add(error);
      throw error;
    }
  }

  const sender = { send };
  return sender;
}

type AuditRecord = {
  kind: "intent" | "outcome";
  id: string;
  intentKey: string;
  time: string;
  profile: string;
  operation: string;
  method: MutationMethod;
  url: string;
  target: string;
  effect: MutationEffect;
  ifMatch?: string;
  httpStatus?: number;
  outcome?: string;
};

function flushDirectory(path: string): void {
  // Node cannot open directory handles on Windows; journal fsync still
  // applies, and recordAudit pre-creates the journal file there instead.
  if (process.platform === "win32") return;
  const fd = openSync(path, "r");
  try { fsyncSync(fd); } finally { closeSync(fd); }
}

function serialize(value: unknown): string | undefined {
  try {
    const serialized = JSON.stringify(value);
    if (value !== undefined && serialized === undefined) throw new Error("Not JSON");
    return serialized;
  } catch {
    throw new AxiError("Mutation state or payload is not serializable", "VALIDATION_ERROR", [
      "Provide JSON-serializable current state and payload",
    ]);
  }
}

export function createMutationCoordinator(args: {
  profile: AnyProfile;
  delegated: DelegatedAuth;
  application: ApplicationAuth;
  transport: MutationTransport;
  clock?: () => number;
  journalPath?: string;
}): MutationCoordinator {
  const { profile, delegated, application, transport } = args;
  const clock = args.clock ?? Date.now;
  const journalPath = args.journalPath ?? resolveWriteLogPath();
  const validated = profile.mode === "application" ? validateApplicationProfile(profile) : validateDelegatedProfile(profile);
  const scope: WriteScope = Object.freeze({
    tenantId: validated.tenantId,
    allowWrites: validated.writes?.allowWrites === true,
    operations: Object.freeze([...(validated.writes?.operations ?? [])]),
  });
  const sender = createMutationSender({ profile: validated, delegated, application, transport });

  function enforceWriteGates(definition: MutationDefinition): void {
    // Gate 1: forced read-only overrides every profile.
    // Gate 2: the profile must opt in through hand-edited configuration.
    if (readOnlyForced() || !scope.allowWrites) {
      throw new AxiError(
        `blocked: writes are disabled for this profile (${definition.method} ${definition.operation})`,
        "WRITES_DISABLED",
        [
          "Writes are disabled for this profile",
          "Enable named writes explicitly in the profile configuration; see README.md for the write gates",
        ],
      );
    }
    // Classification precedes scope gating, mirroring az-axi: a malformed
    // definition fails as VALIDATION_ERROR before membership is considered.
    if (!definition.operation.trim()) {
      throw new AxiError("Mutation operation must be a non-empty name", "VALIDATION_ERROR", [
        "Name the approved operation the coordinator should authorize",
      ]);
    }
    if (!definition.target.trim()) {
      throw new AxiError("Mutation target must be a non-empty name", "VALIDATION_ERROR", [
        "Name the exact target this mutation would change",
      ]);
    }
    if (definition.version !== "v1.0") {
      throw new AxiError(`Unsupported mutation version ${definition.version}`, "VALIDATION_ERROR", [
        "Mutations are bound to v1.0; beta writes stay blocked",
      ]);
    }
    // Gate 3: the operation must fall inside the profile's own configured scope.
    if (!scope.operations.includes(definition.operation)) {
      throw new AxiError(
        `blocked: operation '${definition.operation}' is outside the write scope of this profile (${definition.method} request)`,
        "OPERATION_NOT_WRITABLE",
        [
          "Writes are limited to this profile's configured operations",
          "The named write must be in the profile's writes.operations allowlist; see README.md for configuration",
        ],
      );
    }
  }

  function boundMutationUrl(definition: MutationDefinition): string {
    const prefix = "/v1.0/";
    const segments = definition.path.split("/");
    if (!definition.path.startsWith(prefix) || segments[0] !== ""
      || segments.slice(1).some(segment => !segment || segment === "." || segment === "..")
      || /[\s\x00-\x1f\x7f\\?#{}]/.test(definition.path)) {
      throw new AxiError("Refusing mutation path outside the bound operation scope", "VALIDATION_ERROR", [
        `Mutation paths must stay under ${prefix} without traversal, query strings or fragments`,
      ]);
    }
    for (const segment of segments.slice(1)) {
      let decoded: string;
      try {
        decoded = decodeURIComponent(segment);
      } catch {
        throw new AxiError("Invalid mutation path encoding", "VALIDATION_ERROR", [
          "Bind one encoded resource identifier per path segment",
        ]);
      }
      encodeGraphPathSegment(decoded);
    }
    return `https://${GRAPH_HOST}${definition.path}`;
  }

  function preview(definition: MutationDefinition, current?: unknown, isNoop?: (current: unknown) => boolean): MutationPreview {
    enforceWriteGates(definition);
    const url = boundMutationUrl(definition);
    return Object.freeze({
      operation: definition.operation,
      method: definition.method,
      url,
      effect: definition.effect,
      target: definition.target,
      noop: current === undefined || isNoop === undefined ? false : isNoop(current),
      currentState: current === undefined ? null : serialize(redactGraphValue(current))!,
      proposedChange: definition.payload === undefined ? null : serialize(redactGraphValue(definition.payload))!,
    });
  }

  // Metadata only: never payloads, headers, secrets or bodies.
  function recordAudit(record: Omit<AuditRecord, "time">): void {
    const safe = redactGraphValue({ ...record, time: new Date(clock()).toISOString() }) as AuditRecord;
    const line = JSON.stringify({ ...safe, intentKey: record.intentKey });
    try {
      const directory = dirname(journalPath);
      const firstCreated = mkdirSync(directory, { recursive: true, mode: 0o700 });
      if (firstCreated) {
        let path = directory;
        const parent = dirname(firstCreated);
        while (path !== parent) {
          flushDirectory(path);
          path = dirname(path);
        }
        flushDirectory(parent);
      }
      if (process.platform === "win32" && !existsSync(journalPath)) {
        const journal = openSync(journalPath, "a", 0o600);
        try { fsyncSync(journal); } finally { closeSync(journal); }
      }
      const lockPath = `${journalPath}.lock`;
      const lock = openSync(lockPath, "wx", 0o600);
      try {
        const created = !existsSync(journalPath);
        const journal = openSync(journalPath, "a+", 0o600);
        try {
          if (record.kind === "intent") {
            const contents = readFileSync(journal, "utf8");
            if (contents && !contents.endsWith("\n")) throw new Error("Incomplete mutation journal");
            const records = contents.split("\n").filter(Boolean)
              .map(entry => JSON.parse(entry) as AuditRecord);
            if (records.some(entry => !entry || !["intent", "outcome"].includes(entry.kind)
              || typeof entry.intentKey !== "string" || !/^[a-f0-9]{64}$/.test(entry.intentKey))) {
              throw new Error("Invalid mutation journal");
            }
            // No automatic replay: a recreated coordinator refuses an intent
            // recorded without a terminal outcome, and a duplicate intent id
            // is never reserved twice.
            if (records.some(entry => entry.kind === "intent" && entry.intentKey === record.intentKey)) {
              throw new AxiError(`blocked: intent '${record.id}' was already reserved`, "ALREADY_EXECUTED", [
                `Intent '${record.id}' requires manual reconciliation; read back its target and never resend it`,
              ]);
            }
          }
          writeFileSync(journal, `${line}\n`);
          fsyncSync(journal);
          if (created) flushDirectory(directory);
        } finally { closeSync(journal); }
      } finally {
        closeSync(lock);
        unlinkSync(lockPath);
      }
    } catch (error) {
      if (error instanceof AxiError && error.code === "ALREADY_EXECUTED") throw error;
      if (record.kind === "intent") {
        throw new AxiError(`Mutation intent '${record.id}' could not be recorded`, "INTENT_NOT_RECORDED", [
          "The mutation was not sent",
          `Check the audit path ${journalPath}; a remaining lock requires manual reconciliation before removal`,
        ]);
      }
      throw new AxiError(`Mutation outcome '${record.id}' could not be recorded`, "OUTCOME_NOT_RECORDED", [
        "The mutation may have been applied; read back the target before doing anything else",
        `Check the audit path ${journalPath}`,
      ]);
    }
  }

  function metadata(definition: MutationDefinition, seen: MutationPreview, id: string, ifMatch: string | undefined): Omit<AuditRecord, "kind" | "time"> {
    return {
      id,
      intentKey: createHash("sha256").update(id).digest("hex"),
      profile: scope.tenantId,
      operation: definition.operation,
      method: definition.method,
      url: seen.url,
      target: definition.target,
      effect: definition.effect,
      ...(ifMatch !== undefined ? { ifMatch } : {}),
    };
  }

  async function execute(definition: MutationDefinition, options: MutationExecuteOptions): Promise<MutationResult> {
    const body = serialize(definition.payload);
    definition = { ...definition, ...(body === undefined ? {} : { payload: JSON.parse(body) as unknown }) };
    options = { ...options };
    const seen = preview(definition, await options.readState(), options.isNoop);
    // Gate 4: without --execute the caller runs the dry run instead.
    if (options.execute !== true) return { kind: "dry-run", preview: seen };
    // Verified already-desired state is a no-op: nothing is sent.
    if (seen.noop) return { kind: "noop", preview: seen };
    // Gate 5: disruptive mutations need the target name back.
    if (definition.effect === "disruptive") {
      if (options.confirm === undefined) {
        throw new AxiError(
          `blocked: disruptive ${definition.method} needs --confirm '${definition.target}'`,
          "CONFIRM_REQUIRED",
          [`Re-run with --confirm '${definition.target}'`],
        );
      }
      if (options.confirm !== definition.target) {
        throw new AxiError(
          `blocked: --confirm '${options.confirm}' does not match target '${definition.target}'`,
          "CONFIRM_MISMATCH",
          [`Re-run with --confirm '${definition.target}'`],
        );
      }
    }
    const id = options.intentId ?? randomUUID();
    // A re-read alone is not atomic protection; conditional writes travel as
    // If-Match only where the endpoint supports them (WRITE-N evidence).
    const meta = metadata(definition, seen, id, options.ifMatch);
    // Failure to record intent blocks the send.
    recordAudit({ ...meta, kind: "intent" });
    // Re-read before the single send; a failed re-read aborts without sending.
    let fresh: unknown;
    try {
      fresh = await options.readState();
    } catch (error) {
      recordAudit({ ...meta, kind: "outcome", httpStatus: 0, outcome: "NOT_SENT" });
      throw error;
    }
    if (options.isNoop?.(fresh) === true) {
      recordAudit({ ...meta, kind: "outcome", httpStatus: 0, outcome: "NOT_SENT" });
      return { kind: "noop", preview: seen };
    }
    const authorization = Object.freeze({ nonce: randomUUID(), method: definition.method, url: seen.url });
    authorizations.set(authorization, sender);
    let sent: MutationResponse;
    try {
      sent = await sender.send(authorization, {
        ...(body !== undefined ? { body } : {}),
        ...(options.ifMatch !== undefined ? { ifMatch: options.ifMatch } : {}),
        ...(options.scopes !== undefined ? { scopes: options.scopes } : {}),
      });
    } catch (error) {
      if (mutationNotSent(error)) {
        recordAudit({ ...meta, kind: "outcome", httpStatus: 0, outcome: "NOT_SENT" });
        throw error;
      }
      const httpStatus = mutationHttpStatus(error);
      if (httpStatus > 0 && mutationAccepted(error)) {
        recordAudit({ ...meta, kind: "outcome", httpStatus, outcome: "SUCCESS" });
        return { kind: "success", preview: seen, auditId: id, status: httpStatus, response: {} };
      }
      if (httpStatus >= 400 && httpStatus < 500 && httpStatus !== 408) {
        recordAudit({ ...meta, kind: "outcome", httpStatus, outcome: "FAILED" });
        return { kind: "failed", preview: seen, auditId: id, status: httpStatus };
      }
      recordAudit({ ...meta, kind: "outcome", httpStatus, outcome: "OUTCOME_UNKNOWN" });
      return {
        kind: "unknown",
        preview: seen,
        auditId: id,
        httpStatus,
        guidance: `Mutation ${definition.operation} may or may not have been applied (audit ${id}); read back target '${definition.target}' before doing anything else; never replay this intent`,
      };
    }
    recordAudit({ ...meta, kind: "outcome", httpStatus: sent.status, outcome: "SUCCESS" });
    return { kind: "success", preview: seen, auditId: id, status: sent.status, response: sent.body };
  }

  return { scope, preview, execute };
}
