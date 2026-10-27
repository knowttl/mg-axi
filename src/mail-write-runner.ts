import { AxiError } from "axi-sdk-js";
import type { MutationCoordinator, MutationDefinition, MutationExecuteOptions, MutationResult } from "./mutations.js";

// Validate the entire selection before the first credential or request.
export function mailWriteTargets(flags: Record<string, string | boolean>, help: string): string[] {
  const ids = String(flags.ids ?? "").split(",");
  if (ids.length > 20 || ids.some(id => !id || id.trim() !== id || /[\x00-\x20\x7f/\\?#%$]/.test(id) || id === "." || id === "..") || new Set(ids).size !== ids.length) {
    throw new AxiError("--ids requires 1 to 20 distinct message IDs, comma-separated without whitespace or path syntax", "VALIDATION_ERROR", [help]);
  }
  if (flags.execute === true && flags.confirm !== ids.join(",")) {
    throw new AxiError("--execute requires --confirm listing every ID in exactly the same order as --ids", flags.confirm === undefined ? "CONFIRM_REQUIRED" : "CONFIRM_MISMATCH", [help]);
  }
  return ids;
}

export async function runMailWrites(args: {
  ids: string[];
  coordinator: MutationCoordinator;
  definition: (id: string) => MutationDefinition;
  prepare: (id: string) => Promise<{ payload: unknown; options: MutationExecuteOptions }>;
}): Promise<{ results: Record<string, unknown>[]; complete: boolean }> {
  const { ids, coordinator, definition, prepare } = args;
  for (const id of ids) coordinator.preview(definition(id));
  const results: Record<string, unknown>[] = [];
  let stopped = false;
  for (const id of ids) {
    if (stopped) {
      results.push({ id, status: "not-attempted" });
      continue;
    }
    try {
      const prepared = await prepare(id);
      const result: MutationResult = await coordinator.execute({ ...definition(id), payload: prepared.payload }, prepared.options);
      results.push({ id, status: result.kind, ...("auditId" in result && result.auditId ? { auditId: result.auditId } : {}),
        ...(result.kind === "dry-run" ? { preview: { ...result.preview } } : {}),
        ...(result.kind === "unknown" ? { guidance: result.guidance } : {}),
        ...(result.kind === "failed" ? { httpStatus: result.status } : {}) });
      stopped = result.kind === "failed" || result.kind === "unknown";
    } catch (error) {
      // Preserve earlier outcomes even if a read or journal failed. In
      // particular, an outcome-journal failure can follow a successful send.
      const code = error instanceof AxiError ? error.code : "GRAPH_ERROR";
      results.push({ id, status: ["OUTCOME_NOT_RECORDED", "CONFIG_INSECURE"].includes(code) ? "unknown" : "failed", code,
        guidance: "Read back this target before continuing; never replay an uncertain intent" });
      stopped = true;
    }
  }
  return { results, complete: !stopped };
}
