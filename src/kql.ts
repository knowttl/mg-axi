/// <reference types="@kusto/language-service-next" />
import { AxiError } from "axi-sdk-js";

export interface KqlDiagnostic {
  line: number;
  column: number;
  code: string;
  token: string;
  message: string;
}

export class KqlValidationError extends AxiError {
  constructor(message: string, public readonly diagnostics: KqlDiagnostic[], help: string[]) {
    super(message, "VALIDATION_ERROR", help);
  }
}

export function position(query: string, offset: number): { line: number; column: number } {
  const lines = query.slice(0, offset).split(/\r\n|\r|\n/);
  return { line: lines.length, column: lines.at(-1)!.length + 1 };
}

// Never render source excerpts or string literals, even for local failures.
export function diagnosticToken(token: string): string {
  return /^(?:[A-Za-z_][A-Za-z_0-9-]{0,63}|[|(),.\[\]+*/%<>=!~:-])$/.test(token) ? token : "<redacted>";
}

export async function validateSyntax(query: string): Promise<void> {
  await import("@kusto/language-service-next/bridge.min.js");
  await import("@kusto/language-service-next/Kusto.Language.Bridge.min.js");
  const parsed = Kusto.Language.KustoCode.Parse(query)!;
  // The runtime spells this toArray; the published declarations say ToArray.
  const bridge = Bridge as unknown as { toArray(value: unknown): Kusto.Language.Diagnostic[] };
  const diagnostics = bridge.toArray(parsed.GetDiagnostics()!).map((item): KqlDiagnostic => ({
    ...position(query, item.Start), code: item.Code!,
    token: diagnosticToken(query.slice(item.Start, item.Start + item.Length)),
    message: item.Message!.replace(/'[^']*'|"[^"]*"|`[^`]*`/g, "<redacted>"),
  }));
  if (!diagnostics.length) return;
  const first = diagnostics[0]!;
  throw new KqlValidationError(`KQL syntax error at line ${first.line}, column ${first.column}: ${first.message} The query was not sent.`, diagnostics.slice(0, 5), [
    ...(diagnostics.length > 5 ? [`Showing 5 of ${diagnostics.length} syntax diagnostics; fix these and retry`] : []),
    "Fix the query and retry; KQL reference: https://learn.microsoft.com/kusto/query/",
  ]);
}
