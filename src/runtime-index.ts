import { readFileSync } from "node:fs";
import type { SessionOperation } from "./graph-session.js";

let operationsById: Map<string, Readonly<SessionOperation>> | undefined;

// Both command dispatch and session authorization share this process-local cache.
// Return copies so callers cannot change the authoritative policy rows.
export function findOperation(id: string): SessionOperation | undefined {
  if (!operationsById) {
    const index = JSON.parse(readFileSync(new URL("../inventory/runtime.json", import.meta.url), "utf8")) as {
      operations: SessionOperation[];
    };
    operationsById = new Map(index.operations.map(row => [row.id, Object.freeze(row)]));
  }
  const operation = operationsById.get(id);
  return operation ? { ...operation } : undefined;
}
