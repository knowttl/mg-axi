import { AxiError } from "axi-sdk-js";
import { encodeGraphPathSegment, resolveSessionOperation, type GraphSession } from "./graph-session.js";
import { MAIL_FOLDER_SCOPE, MAIL_SETTINGS_SCOPE, validateMailbox } from "./mail-authorization.js";
import { mailText } from "./mail-output.js";
import { mailWriteTargets, runMailWrites } from "./mail-write-runner.js";
import type { MutationCoordinator, MutationDefinition } from "./mutations.js";
import type { AnyProfile } from "./profiles.js";

export const MAIL_FOLDER_WRITE_ALTERNATES: Readonly<Record<string, string>> = {
  "POST:/users/{user-id}/mailFolders": "POST:/me/mailFolders",
  "DELETE:/users/{user-id}/mailFolders/{mailFolder-id}": "DELETE:/me/mailFolders/{mailFolder-id}",
  "POST:/users/{user-id}/outlook/masterCategories": "POST:/me/outlook/masterCategories",
};
// Locale-independent aliases from the Graph v1.0 mailFolder contract.
const WELL_KNOWN = "archive clutter conflicts conversationhistory deleteditems drafts inbox junkemail localfailures msgfolderroot outbox recoverableitemsdeletions scheduled searchfolders sentitems serverfailures syncissues".split(" ");

export async function writeMailFolders(args: {
  session: GraphSession; coordinator: MutationCoordinator; profile: AnyProfile;
  flags: Record<string, string | boolean>; command: string; help: string;
}): Promise<Record<string, unknown>> {
  const { session, coordinator, profile, flags, command, help } = args;
  const mailbox = String(flags.mailbox);
  validateMailbox(mailbox, help);
  if (!profile.enabledPacks.includes("mail")) throw new AxiError("Mail pack is not enabled", "POLICY_DENIED", [help]);
  if (mailbox === "me" && profile.mode === "application") throw new AxiError("me requires delegated authentication", "POLICY_DENIED", [help]);
  const deleting = command === "mail folder delete", category = command === "mail category create";
  const name = String(flags.name ?? "");
  if (!deleting && (!name.trim() || name !== name.trim() || /[\x00-\x1f\x7f,]/.test(name))) throw new AxiError("--name must be nonempty, trimmed text without controls or commas", "VALIDATION_ERROR", [help]);
  if (category && !/^(?:none|preset(?:[0-9]|1[0-9]|2[0-4]))$/.test(String(flags.color))) throw new AxiError("--color must be none or preset0 through preset24", "VALIDATION_ERROR", [help]);
  const ids = deleting ? mailWriteTargets(flags, help) : [name];
  if (flags.execute === true && flags.confirm !== ids.join(",")) throw new AxiError("--execute requires --confirm repeating the exact target name or ordered --ids list", flags.confirm === undefined ? "CONFIRM_REQUIRED" : "CONFIRM_MISMATCH", [help]);
  if (deleting && ids.some(id => WELL_KNOWN.includes(id.toLowerCase()))) throw new AxiError("Well-known folders cannot be deleted", "POLICY_DENIED", [help]);
  const root = mailbox === "me" ? "/me" : "/users/{user-id}";
  const params: Record<string, string> = mailbox === "me" ? {} : { "user-id": mailbox };
  const scopes = profile.mode === "delegated" ? [category ? MAIL_SETTINGS_SCOPE : MAIL_FOLDER_SCOPE] : undefined;
  const readFolder = async (id: string, counts = false) => {
    const value = await session.execute({ profile, operation: resolveSessionOperation("v1.0", "GET", `${root}/mailFolders/{mailFolder-id}`), params: { ...params, "mailFolder-id": id }, scopes,
      query: { $select: counts ? "id,totalItemCount,childFolderCount" : "id" } });
    const row = value as Record<string, unknown> | null;
    if (!row || typeof row.id !== "string" || !row.id || (counts && (row.id !== id || ![row.totalItemCount, row.childFolderCount].every(count => Number.isSafeInteger(count) && Number(count) >= 0)))) throw new AxiError("Graph did not return the requested folder metadata", "GRAPH_ERROR");
    return counts ? { id: row.id, totalItemCount: row.totalItemCount, childFolderCount: row.childFolderCount } : { id: row.id };
  };
  const definition = (id: string): MutationDefinition => ({ operation: command.replaceAll(" ", "."), method: deleting ? "DELETE" : "POST", version: "v1.0", effect: "disruptive", target: id,
    path: `/v1.0${mailbox === "me" ? "/me" : `/users/${encodeGraphPathSegment(mailbox)}`}/${category ? "outlook/masterCategories" : `mailFolders${deleting ? `/${encodeGraphPathSegment(id)}` : ""}`}` });
  // Check every operation gate before credentials or safety reads.
  for (const id of ids) coordinator.preview(definition(id));
  if (deleting) {
    for (const alias of WELL_KNOWN) {
      // An inaccessible or missing alias is not proof that the target is safe.
      // Refuse deletion when the complete protected-ID set cannot be resolved.
      const folder = await readFolder(alias);
      if (ids.includes(folder.id)) throw new AxiError("Well-known folder IDs cannot be deleted", "POLICY_DENIED", [help]);
    }
  }
  const payload = deleting ? undefined : category ? { displayName: name, color: String(flags.color) } : { displayName: name };
  const result = await runMailWrites({ ids, coordinator, definition, includeNewId: !deleting, prepare: async id => {
    const readState = async () => {
      if (deleting) {
        const folder = await readFolder(id, true);
        if (folder.totalItemCount !== 0 || folder.childFolderCount !== 0) throw new AxiError("Only empty folders without child folders can be deleted", "POLICY_DENIED");
        return folder;
      }
      if (!category) return readFolder("msgfolderroot");
      const categories = await session.collect({ profile, operation: resolveSessionOperation("v1.0", "GET", `${root}/outlook/masterCategories`), params, scopes,
        query: { $select: "id,displayName,color", $top: "1000" }, limit: 1000, budget: { maxRequests: 1 } });
      if (!categories.complete) throw new AxiError("Category list is incomplete; cannot safely check for an existing name", "POLICY_DENIED");
      const matches = categories.value.map(value => {
        const row = value as Record<string, unknown> | null;
        if (!row || typeof row.id !== "string" || typeof row.displayName !== "string" || typeof row.color !== "string") throw new AxiError("Graph returned malformed category metadata", "GRAPH_ERROR");
        return row;
      }).filter(row => row.displayName === name);
      if (matches.length > 1 || matches.some(row => row.color !== flags.color)) throw new AxiError("A category with this name has a different color or is ambiguous; creation cannot update it", "POLICY_DENIED");
      return { exists: matches.length === 1 };
    };
    return { payload, options: { execute: flags.execute === true, confirm: id, readState,
      ...(category ? { isNoop: (current: unknown) => (current as { exists: boolean }).exists } : {}),
      ...(profile.mode === "delegated" ? { scopes: [`https://graph.microsoft.com/${category ? "MailboxSettings.ReadWrite" : "Mail.ReadWrite"}`] } : {}),
    } };
  } });
  const text = mailText(flags.full === true);
  for (const row of result.results) {
    if (typeof row.newId === "string") row.newId = text.text(row.newId);
    if (row.preview && typeof row.preview === "object") {
      const preview = row.preview as Record<string, unknown>;
      for (const key of ["currentState", "proposedChange"]) if (typeof preview[key] === "string") preview[key] = text.text(preview[key]);
    }
  }
  return { ...result, help: [...text.help(),
    category ? "Delegated reads need MailboxSettings.Read and creation needs MailboxSettings.ReadWrite; application mode needs scoped Exchange Application MailboxSettings.ReadWrite with no MailboxSettings.* application grant in Entra"
      : "Delegated reads need Mail.ReadBasic and writes need Mail.ReadWrite; application mode needs scoped Exchange Application Mail.ReadWrite with no Mail.* application grant in Entra",
    ...(flags.execute === true ? [] : ["Repeat with --execute and --confirm repeating the exact --name or ordered --ids list"]),
    ...(deleting ? ["Deletion requires every well-known folder ID to be readable and both item and child-folder counts to be zero. A fresh read cannot make deletion atomic with other mailbox clients"] : []),
  ] };
}
