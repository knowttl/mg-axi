import { AxiError } from "axi-sdk-js";
import { MAIL_FOLDER_SCOPE } from "./mail-authorization.js";
import { resolveSessionOperation, type GraphSession } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

export const MAIL_FOLDER_ALTERNATES: Readonly<Record<string, string>> = {
  "GET:/users/{user-id}/mailFolders": "GET:/me/mailFolders",
  "GET:/users/{user-id}/mailFolders/{mailFolder-id}": "GET:/me/mailFolders/{mailFolder-id}",
};
const LIST_FIELDS = ["id", "displayName", "totalItemCount", "unreadItemCount"];
const SHOW_FIELDS = [...LIST_FIELDS, "parentFolderId", "childFolderCount", "isHidden"];

export async function readMailFolders(session: GraphSession, flags: Record<string, string | boolean>, profile: AnyProfile, show: boolean, help: string): Promise<Record<string, unknown>> {
  const mailbox = String(flags.mailbox);
  if (mailbox !== "me" && !/^(?:[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|[A-Za-z0-9.!_+%-]+@[A-Za-z0-9.-]+)$/i.test(mailbox)) {
    throw new AxiError("--mailbox must be a user UUID, UPN or me", "VALIDATION_ERROR", [help]);
  }
  const path = `${mailbox === "me" ? "/me" : "/users/{user-id}"}/mailFolders${show ? "/{mailFolder-id}" : ""}`;
  const operation = resolveSessionOperation("v1.0", "GET", path);
  const params: Record<string, string> = mailbox === "me" ? {} : { "user-id": mailbox };
  if (show) params["mailFolder-id"] = String(flags.id);
  const scopes = profile.mode === "delegated" ? [MAIL_FOLDER_SCOPE] : undefined;
  const defaults = show ? SHOW_FIELDS : LIST_FIELDS;
  const fields = flags.fields === undefined ? defaults : String(flags.fields).split(",");
  if (!fields.length || fields.some(field => !SHOW_FIELDS.includes(field))) throw new AxiError("Unknown folder field", "VALIDATION_ERROR", [help, `Reviewed folder fields: ${SHOW_FIELDS.join(", ")}`]);
  let truncated = false;
  const project = (raw: unknown) => {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) throw new AxiError("Graph returned a malformed folder", "GRAPH_ERROR");
    const source = raw as Record<string, unknown>;
    return Object.fromEntries(fields.filter(field => Object.hasOwn(source, field)).map(field => {
      const value = source[field];
      const expected = field === "isHidden" ? "boolean" : field.endsWith("Count") ? "number" : "string";
      if (value !== null && (typeof value !== expected || (expected === "number" && (!Number.isSafeInteger(value) || Number(value) < 0)))) {
        throw new AxiError("Graph returned malformed folder metadata", "GRAPH_ERROR");
      }
      if (typeof value === "string" && flags.full !== true && value.length > 500) {
        truncated = true;
        return [field, `${value.slice(0, 500)}... (truncated, ${value.length} chars total)`];
      }
      return [field, value];
    }));
  };
  const query = { $select: fields.join(",") };
  if (show) {
    const folder = project(await session.execute({ profile, operation, params, query, scopes }));
    return { folder, ...(truncated ? { help: ["Repeat the same command with --full for complete text; redaction remains"] } : {}) };
  }
  const limit = flags.limit === undefined ? 100 : Number(flags.limit);
  if (!Number.isSafeInteger(limit) || limit < 1 || limit > 1000) throw new AxiError("--limit must be between 1 and 1000", "VALIDATION_ERROR", [help]);
  const result = await session.collect({ profile, operation, params, query: { ...query, $top: String(limit) }, scopes, limit,
    budget: { maxRequests: 1 },
    cursor: flags.cursor === undefined ? undefined : String(flags.cursor) });
  const folders = result.value.map(project);
  return { folders, ...listTotals(folders.length, result.total, "folders", result.complete), complete: result.complete,
    ...(!result.complete ? { reason: result.reason, cursor: result.cursor } : {}),
    help: [
      ...(truncated ? ["Repeat the same command with --full for complete text; redaction remains"] : []),
      ...(!folders.length && result.complete ? ["0 folders returned directly under the mailbox root; hidden folders are excluded"] : []),
      ...(!result.complete ? ["Resume with the same flags plus --cursor <cursor-from-output>"] : []),
      "mg-axi mail folder show --mailbox <user-id-or-upn-or-me> --id <folder-id> --profile <name>",
    ] };
}
