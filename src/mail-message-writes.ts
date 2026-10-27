import { AxiError } from "axi-sdk-js";
import { encodeGraphPathSegment, REDACTED, resolveSessionOperation, type GraphSession } from "./graph-session.js";
import { MAIL_MESSAGE_SCOPE, validateMailbox } from "./mail-authorization.js";
import { mailText } from "./mail-output.js";
import { mailWriteTargets, runMailWrites } from "./mail-write-runner.js";
import type { MutationCoordinator, MutationDefinition } from "./mutations.js";
import type { AnyProfile } from "./profiles.js";

export const MAIL_WRITE_ALTERNATES: Readonly<Record<string, string>> = {
  "PATCH:/users/{user-id}/messages/{message-id}": "PATCH:/me/messages/{message-id}",
};
export const MAIL_MESSAGE_WRITE_SCOPE = "https://graph.microsoft.com/Mail.ReadWrite";

export async function writeMailMessages(args: {
  session: GraphSession; coordinator: MutationCoordinator; profile: AnyProfile;
  flags: Record<string, string | boolean>; action: string; help: string;
}): Promise<Record<string, unknown>> {
  const { session, coordinator, profile, flags, action, help } = args;
  const mailbox = String(flags.mailbox);
  validateMailbox(mailbox, help);
  if (!profile.enabledPacks.includes("mail")) throw new AxiError("Mail pack is not enabled", "POLICY_DENIED", [help]);
  if (mailbox === "me" && profile.mode === "application") throw new AxiError("me requires delegated authentication", "POLICY_DENIED", [help]);
  const ids = mailWriteTargets(flags, help);
  let add: string[] = [], remove: string[] = [];
  const field = action === "tag" ? "categories" : action === "flag" ? "flag" : "isRead";
  if (action === "tag") {
    const categories = (value: string | boolean | undefined) => value === undefined ? [] : String(value).split(",");
    add = categories(flags.add);
    remove = categories(flags.remove);
    if ((!add.length && !remove.length) || [...add, ...remove].some(value => !value.trim() || value !== value.trim() || /[\x00-\x1f\x7f]/.test(value)) || add.some(value => remove.includes(value))) {
      throw new AxiError("Tag needs --add and/or --remove with nonempty category names; a category cannot appear in both", "VALIDATION_ERROR", [help]);
    }
  } else if (action === "flag") {
    if (!["notFlagged", "flagged", "complete"].includes(String(flags.status))) throw new AxiError("--status must be notFlagged, flagged or complete", "VALIDATION_ERROR", [help]);
  } else if (action !== "mark" || !["read", "unread"].includes(String(flags.state))) {
    throw new AxiError("--state must be read or unread", "VALIDATION_ERROR", [help]);
  }
  const root = mailbox === "me" ? "/me" : "/users/{user-id}";
  const readOperation = resolveSessionOperation("v1.0", "GET", `${root}/messages/{message-id}`);
  const definition = (id: string): MutationDefinition => ({
    operation: `mail.message.${action}`, method: "PATCH", version: "v1.0", effect: "disruptive", target: id,
    path: `/v1.0${mailbox === "me" ? "/me" : `/users/${encodeGraphPathSegment(mailbox)}`}/messages/${encodeGraphPathSegment(id)}`,
  });
  const text = mailText(flags.full === true);
  const result = await runMailWrites({ ids, coordinator, definition, prepare: async id => {
    const read = async () => {
      const raw = await session.execute({ profile, operation: readOperation,
        params: { "message-id": id, ...(mailbox === "me" ? {} : { "user-id": mailbox }) },
        query: { $select: `id,${field}` }, ...(profile.mode === "delegated" ? { scopes: [MAIL_MESSAGE_SCOPE] } : {}) });
      const row = raw as Record<string, unknown> | null;
      if (!row || row.id !== id || (field === "categories" ? !Array.isArray(row.categories) || row.categories.some(value => typeof value !== "string")
        : field === "isRead" ? typeof row.isRead !== "boolean"
          : !row.flag || typeof row.flag !== "object" || !["notFlagged", "flagged", "complete"].includes(String((row.flag as Record<string, unknown>).flagStatus)))) {
        throw new AxiError("Graph did not return the requested message state", "GRAPH_ERROR");
      }
      return { [field]: field === "flag" ? { flagStatus: (row.flag as Record<string, unknown>).flagStatus } : row[field] };
    };
    const current = await read();
    if (field === "categories" && flags.execute === true && (current.categories as string[]).includes(REDACTED)) {
      throw new AxiError("Cannot safely preserve redacted category names", "POLICY_DENIED");
    }
    const payload = field === "categories" ? { categories: [...new Set([...(current.categories as string[]).filter(value => !remove.includes(value)), ...add])] }
      : field === "flag" ? { flag: { flagStatus: String(flags.status) } } : { isRead: flags.state === "read" };
    const same = (state: unknown) => {
      if (field !== "categories") return JSON.stringify(state) === JSON.stringify(payload);
      const values = (state as { categories: string[] }).categories;
      const desired = (payload as { categories: string[] }).categories;
      return values.length === desired.length && values.every(value => desired.includes(value));
    };
    let first = true;
    return { payload, options: { execute: flags.execute === true, confirm: id,
      ...(profile.mode === "delegated" ? { scopes: [MAIL_MESSAGE_WRITE_SCOPE] } : {}),
      isNoop: same, readState: async () => {
        if (first) { first = false; return current; }
        const fresh = await read();
        // Do not overwrite another client's category edits with a stale
        // read-modify-write payload. The coordinator records NOT_SENT.
        if (field === "categories" && !same(fresh) && JSON.stringify(fresh) !== JSON.stringify(current)) throw new AxiError("Message categories changed during execution; read back before making a new request", "STATE_CHANGED");
        return fresh;
      } } };
  } });
  for (const row of result.results) {
    if (row.preview && typeof row.preview === "object") {
      const preview = row.preview as Record<string, unknown>;
      for (const key of ["currentState", "proposedChange"]) if (typeof preview[key] === "string") preview[key] = text.text(preview[key]);
    }
  }
  return { ...result, help: [...text.help(),
    "Delegated writes need Mail.ReadWrite; preview and fresh reads need Mail.Read. Application writes need scoped Exchange Application Mail.ReadWrite, with no Mail.* application grant in Entra",
    ...(flags.execute === true ? [] : ["Repeat with --execute and --confirm identical to --ids to apply this preview"]),
    ...(!result.complete ? ["Batch stopped; later IDs were not attempted. Read back failed or unknown targets before continuing"] : []),
  ] };
}
