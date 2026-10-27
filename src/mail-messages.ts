import { AxiError } from "axi-sdk-js";
import { MAIL_FOLDER_SCOPE, validateMailbox } from "./mail-authorization.js";
import { mailText } from "./mail-output.js";
import { resolveSessionOperation, type GraphSession } from "./graph-session.js";
import { listTotals } from "./list-totals.js";
import type { AnyProfile } from "./profiles.js";

export const MAIL_MESSAGE_ALTERNATES: Readonly<Record<string, string>> = {
  "GET:/users/{user-id}/messages": "GET:/me/messages",
  "GET:/me/messages": "GET:/me/mailFolders/{mailFolder-id}/messages",
  "GET:/me/mailFolders/{mailFolder-id}/messages": "GET:/users/{user-id}/mailFolders/{mailFolder-id}/messages",
};
const DEFAULT_FIELDS = ["id", "subject", "receivedDateTime", "isRead"];
const STRING_FIELDS = ["id", "subject", "receivedDateTime", "sentDateTime", "conversationId", "parentFolderId", "importance", "internetMessageId"];
const BOOLEAN_FIELDS = ["isRead", "hasAttachments", "isDraft"];
const FIELDS = [...STRING_FIELDS, ...BOOLEAN_FIELDS, "categories"];

export async function readMailMessages(session: GraphSession, flags: Record<string, string | boolean>, profile: AnyProfile, help: string): Promise<Record<string, unknown>> {
  const mailbox = String(flags.mailbox);
  validateMailbox(mailbox, help);
  if (["search", "filter", "conversation"].filter(key => flags[key] !== undefined).length > 1) throw new AxiError("--search, --filter and --conversation are mutually exclusive", "VALIDATION_ERROR", [help]);
  for (const key of ["search", "filter", "conversation", "folder"]) {
    if (flags[key] !== undefined && !String(flags[key]).trim()) throw new AxiError(`--${key} cannot be empty`, "VALIDATION_ERROR", [help]);
  }
  const fields = flags.fields === undefined ? DEFAULT_FIELDS : String(flags.fields).split(",");
  if (fields.some(field => !FIELDS.includes(field))) throw new AxiError("Unknown message field; bodies and previews are unavailable in lists", "VALIDATION_ERROR", [help, `Reviewed message fields: ${FIELDS.join(", ")}`]);
  const limit = flags.limit === undefined ? 100 : Number(flags.limit);
  if (!Number.isSafeInteger(limit) || limit < 1 || limit > 1000) throw new AxiError("--limit must be between 1 and 1000", "VALIDATION_ERROR", [help]);
  const path = `${mailbox === "me" ? "/me" : "/users/{user-id}"}${flags.folder === undefined ? "" : "/mailFolders/{mailFolder-id}"}/messages`;
  const operation = resolveSessionOperation("v1.0", "GET", path);
  const params: Record<string, string> = mailbox === "me" ? {} : { "user-id": mailbox };
  if (flags.folder !== undefined) params["mailFolder-id"] = String(flags.folder);
  const query: Record<string, string> = { $select: fields.join(","), $top: String(limit) };
  if (flags.search !== undefined) {
    const search = String(flags.search);
    query.$search = search.startsWith('"') && search.endsWith('"') ? search : `"${search}"`;
  } else if (flags.filter !== undefined) query.$filter = String(flags.filter);
  else if (flags.conversation !== undefined) query.$filter = `conversationId eq '${String(flags.conversation).replaceAll("'", "''")}'`;
  else query.$orderby = "receivedDateTime desc";
  const result = await session.collect({ profile, operation, params, query, scopes: profile.mode === "delegated" ? [MAIL_FOLDER_SCOPE] : undefined,
    limit, budget: { maxRequests: 1 }, cursor: flags.cursor === undefined ? undefined : String(flags.cursor) });
  const output = mailText(flags.full === true);
  const text = output.text;
  const messages = result.value.map(raw => {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) throw new AxiError("Graph returned a malformed message", "GRAPH_ERROR");
    const source = raw as Record<string, unknown>;
    return Object.fromEntries(fields.filter(field => Object.hasOwn(source, field)).map(field => {
      const value = source[field];
      if (value === null) return [field, value];
      if (STRING_FIELDS.includes(field) && typeof value === "string") return [field, text(value)];
      if (BOOLEAN_FIELDS.includes(field) && typeof value === "boolean") return [field, value];
      if (field === "categories" && Array.isArray(value) && value.every(item => typeof item === "string")) return [field, value.map(text)];
      throw new AxiError("Graph returned malformed message metadata", "GRAPH_ERROR");
    }));
  });
  return { messages, ...listTotals(messages.length, result.total, "messages", result.complete), complete: result.complete,
    ...(!result.complete ? { reason: result.reason, ...(result.cursor ? { cursor: result.cursor } : {}) } : {}),
    help: [
      ...output.help(),
      ...(!messages.length && result.complete ? ["0 messages matched in the selected mailbox or folder"] : []),
      ...(result.cursor ? ["Resume with the same flags plus --cursor <cursor-from-output>; cursors contain query text and must be treated as sensitive"] : []),
    ] };
}
