import { AxiError } from "axi-sdk-js";
import type { SessionOperation, TransportResponse } from "./graph-session.js";
import { isMailOperation } from "./mail-authorization.js";
import { redactSecrets } from "./redact.js";
import { ResponseDiagnostics } from "./response-diagnostics.js";
import { HUNTING_OPERATION } from "./security-hunting.js";

export class ResponseShapeError extends AxiError {
  readonly diagnostics: Record<string, unknown>;

  constructor(message: string, operation: SessionOperation, response: TransportResponse, data: unknown = response.body,
    issues?: readonly { path: readonly PropertyKey[]; expected?: unknown }[]) {
    super(message, "RESPONSE_INVALID", ["Check API compatibility and proxy behavior using the status, content type and bounded shape report; treat this result as unknown, never empty"]);
    const diagnostics = new ResponseDiagnostics({ text: redactSecrets }, [this.code]);
    diagnostics.record(response.status, new Headers(response.headers), data);
    if (issues) diagnostics.validationFailed(issues);
    this.diagnostics = diagnostics.forError(this);
    // These routes can use tenant data as otherwise ordinary identifier keys.
    if (isMailOperation(operation) || operation.id === HUNTING_OPERATION
      || /\/(?:attributeSets|customSecurityAttributeDefinitions|customSecurityAttributes)(?:\/|$)/.test(operation.path)) {
      const hideKeys = (path: string) => path.replace(/\["(?:\\.|[^"\\])*"\]/g, '["<key>"]');
      for (const field of ["shape", "issues"]) {
        const entries = this.diagnostics[field];
        if (Array.isArray(entries)) this.diagnostics[field] = entries.map(entry => ({ ...entry, path: hideKeys(entry.path) }));
      }
      const paths = this.diagnostics.issuePaths;
      if (Array.isArray(paths)) this.diagnostics.issuePaths = paths.map(hideKeys);
    }
  }
}
