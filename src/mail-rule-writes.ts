import { open } from "node:fs/promises";
import { constants } from "node:fs";
import { AxiError } from "axi-sdk-js";
import { encodeGraphPathSegment, redactGraphValue, resolveSessionOperation, type GraphSession } from "./graph-session.js";
import { MAIL_SETTINGS_SCOPE, validateMailbox } from "./mail-authorization.js";
import { mailText } from "./mail-output.js";
import { mailWriteTargets, runMailWrites } from "./mail-write-runner.js";
import type { MutationCoordinator, MutationDefinition } from "./mutations.js";
import type { AnyProfile } from "./profiles.js";

export const MAIL_RULE_WRITE_ALTERNATES: Readonly<Record<string, string>> = {
  "POST:/users/{user-id}/mailFolders/{mailFolder-id}/messageRules": "POST:/me/mailFolders/{mailFolder-id}/messageRules",
  "PATCH:/users/{user-id}/mailFolders/{mailFolder-id}/messageRules/{messageRule-id}": "PATCH:/me/mailFolders/{mailFolder-id}/messageRules/{messageRule-id}",
  "DELETE:/users/{user-id}/mailFolders/{mailFolder-id}/messageRules/{messageRule-id}": "DELETE:/me/mailFolders/{mailFolder-id}/messageRules/{messageRule-id}",
};
type ObjectValue = Record<string, unknown>;
const object = (value: unknown): value is ObjectValue => value !== null && typeof value === "object" && !Array.isArray(value);
const text = (value: unknown): boolean => typeof value === "string" && value.length > 0 && !/[\x00-\x1f\x7f]/.test(value);
const strings = (value: unknown): boolean => Array.isArray(value) && value.length <= 100 && value.every(text);
const integer = (value: unknown): boolean => Number.isInteger(value) && Number(value) >= 0 && Number(value) <= 2147483647;
const enumeration = (...values: string[]) => (value: unknown): boolean => values.includes(String(value)) && typeof value === "string";
const boolean = (value: unknown): boolean => typeof value === "boolean";
const forbidden = ["forwardTo", "forwardAsAttachmentTo", "redirectTo", "permanentDelete"];
const actions: Record<string, (value: unknown) => boolean> = {
  assignCategories: strings, copyToFolder: text, moveToFolder: text, delete: boolean,
  markAsRead: boolean, markImportance: enumeration("low", "normal", "high"), stopProcessingRules: boolean,
};
const recipients = (value: unknown): boolean => Array.isArray(value) && value.length <= 100 && value.every(row =>
  object(row) && Object.keys(row).every(key => key === "emailAddress") && object(row.emailAddress)
  && Object.keys(row.emailAddress).every(key => key === "name" || key === "address")
  && text(row.emailAddress.address) && (row.emailAddress.name === undefined || text(row.emailAddress.name)));
const predicates: Record<string, (value: unknown) => boolean> = {
  ...Object.fromEntries("bodyContains bodyOrSubjectContains categories headerContains recipientContains senderContains subjectContains".split(" ").map(key => [key, strings])),
  ...Object.fromEntries("hasAttachments isApprovalRequest isAutomaticForward isAutomaticReply isEncrypted isMeetingRequest isMeetingResponse isNonDeliveryReport isPermissionControlled isReadReceipt isSigned isVoicemail notSentToMe sentCcMe sentOnlyToMe sentToMe sentToOrCcMe".split(" ").map(key => [key, boolean])),
  fromAddresses: recipients, sentToAddresses: recipients, importance: enumeration("low", "normal", "high"),
  sensitivity: enumeration("normal", "personal", "private", "confidential"),
  messageActionFlag: enumeration("any", "call", "doNotForward", "followUp", "fyi", "forward", "noResponseNecessary", "read", "reply", "replyToAll", "review"),
  withinSizeRange: value => object(value) && Object.keys(value).length > 0
    && Object.keys(value).every(key => ["minimumSize", "maximumSize"].includes(key) && integer(value[key]))
    && (value.minimumSize === undefined || value.maximumSize === undefined || Number(value.minimumSize) <= Number(value.maximumSize)),
};
function validateFields(value: unknown, fields: Record<string, (value: unknown) => boolean>): boolean {
  return object(value) && Object.entries(value).every(([key, item]) => Object.hasOwn(fields, key) && fields[key]!(item));
}
async function readRuleFile(path: string, creating: boolean, help: string): Promise<ObjectValue> {
  let value: unknown;
  try {
    const file = await open(path, constants.O_RDONLY | constants.O_NONBLOCK);
    try {
      const stat = await file.stat();
      if (!stat.isFile() || stat.size > 65536) throw new Error();
      const buffer = Buffer.alloc(65537);
      const { bytesRead } = await file.read(buffer, 0, buffer.length, 0);
      if (bytesRead > 65536) throw new Error();
      value = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(buffer.subarray(0, bytesRead)));
    } finally { await file.close(); }
  } catch { throw new AxiError("--file must be a readable UTF-8 JSON object in a regular file, at most 64 KiB", "VALIDATION_ERROR", [help]); }
  const fields = { displayName: text, sequence: integer, isEnabled: boolean,
    conditions: (item: unknown) => validateFields(item, predicates), exceptions: (item: unknown) => validateFields(item, predicates),
    actions: (item: unknown) => validateFields(item, actions) && Object.keys(item as ObjectValue).length > 0 };
  if (!validateFields(value, fields) || !object(value) || Object.keys(value).length === 0
    || (creating && ["displayName", "sequence", "isEnabled", "conditions", "actions"].some(key => value[key] === undefined))) {
    throw new AxiError("Rule file contains missing, unknown or invalid fields; forwarding, redirect and permanentDelete actions are forbidden", "VALIDATION_ERROR", [help]);
  }
  if (value.displayName !== undefined && (String(value.displayName).trim() !== value.displayName || !String(value.displayName).trim() || String(value.displayName).includes(","))) {
    throw new AxiError("Rule displayName must be nonempty, trimmed text without commas", "VALIDATION_ERROR", [help]);
  }
  return value;
}

export async function writeMailRules(args: {
  session: GraphSession; coordinator: MutationCoordinator; profile: AnyProfile;
  flags: Record<string, string | boolean>; command: string; help: string;
}): Promise<Record<string, unknown>> {
  const { session, coordinator, profile, flags, command, help } = args;
  const mailbox = String(flags.mailbox), creating = command === "mail rule create", deleting = command === "mail rule delete";
  validateMailbox(mailbox, help);
  if (!profile.enabledPacks.includes("mail")) throw new AxiError("Mail pack is not enabled", "POLICY_DENIED", [help]);
  if (mailbox === "me" && profile.mode === "application") throw new AxiError("me requires delegated authentication", "POLICY_DENIED", [help]);
  const payload = deleting ? undefined : await readRuleFile(String(flags.file), creating, help);
  const ids = creating ? [String(payload!.displayName)] : mailWriteTargets(flags, help);
  if (flags.execute === true && flags.confirm !== ids.join(",")) throw new AxiError("--execute requires --confirm repeating the exact rule displayName or ordered --ids list", flags.confirm === undefined ? "CONFIRM_REQUIRED" : "CONFIRM_MISMATCH", [help]);
  const root = mailbox === "me" ? "/me" : "/users/{user-id}";
  const params: Record<string, string> = mailbox === "me" ? {} : { "user-id": mailbox };
  const scopes = profile.mode === "delegated" ? [MAIL_SETTINGS_SCOPE] : undefined;
  const definition = (id: string): MutationDefinition => ({ operation: command.replaceAll(" ", "."), method: creating ? "POST" : deleting ? "DELETE" : "PATCH", version: "v1.0", effect: "disruptive", target: id,
    path: `/v1.0${mailbox === "me" ? "/me" : `/users/${encodeGraphPathSegment(mailbox)}`}/mailFolders/inbox/messageRules${creating ? "" : `/${encodeGraphPathSegment(id)}`}` });
  const result = await runMailWrites({ ids, coordinator, definition, includeNewId: creating, prepare: async id => ({ payload, options: {
    execute: flags.execute === true, confirm: id,
    ...(profile.mode === "delegated" ? { scopes: ["https://graph.microsoft.com/MailboxSettings.ReadWrite"] } : {}),
    readState: async () => {
      if (creating) {
        const rules = await session.collect({ profile, operation: resolveSessionOperation("v1.0", "GET", `${root}/mailFolders/{mailFolder-id}/messageRules`), params: { ...params, "mailFolder-id": "inbox" },
          scopes, query: { $select: "id", $top: "1" }, limit: 1, budget: { maxRequests: 1 } });
        return { ruleCount: rules.value.length, complete: rules.complete };
      }
      const row = await session.execute({ profile, operation: resolveSessionOperation("v1.0", "GET", `${root}/mailFolders/{mailFolder-id}/messageRules/{messageRule-id}`), params: { ...params, "mailFolder-id": "inbox", "messageRule-id": id }, scopes,
        query: { $select: "id,isReadOnly,actions" } });
      if (!object(row) || row.id !== id || typeof row.isReadOnly !== "boolean" || !object(row.actions)) throw new AxiError("Graph returned malformed rule metadata", "GRAPH_ERROR");
      if (row.isReadOnly) throw new AxiError("Read-only rules cannot be changed", "POLICY_DENIED");
      const currentActions = row.actions;
      if (!deleting && forbidden.some(key => {
        const value = currentActions[key];
        return value !== undefined && value !== false && (!Array.isArray(value) || value.length > 0);
      })) {
        throw new AxiError("Rules with forwarding, redirect or permanent-delete actions cannot be updated; inspect or delete the rule instead", "POLICY_DENIED");
      }
      return { id: row.id, isReadOnly: row.isReadOnly };
    },
  } }) });
  const output = mailText(flags.full === true);
  for (const row of result.results) {
    for (const [key, value] of Object.entries(row)) if (typeof value === "string") row[key] = output.text(String(redactGraphValue(value)));
    if (object(row.preview)) for (const [key, value] of Object.entries(row.preview)) if (typeof value === "string") row.preview[key] = output.text(String(redactGraphValue(value)));
  }
  return { ...result, help: [...output.help(), "Rule writes require delegated MailboxSettings.Read for safety reads and MailboxSettings.ReadWrite for writes; application mode requires scoped Exchange Application MailboxSettings.ReadWrite, with no MailboxSettings.* application grant in Entra",
    ...(flags.execute === true ? [] : ["Repeat with --execute and --confirm repeating the file displayName or every --ids target in identical order"]),
  ] };
}
