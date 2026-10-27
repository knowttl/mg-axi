import { AxiError } from "axi-sdk-js";
import type { Credential } from "./auth.js";
import type { AnyProfile, AppCredential } from "./profiles.js";

export const MAIL_FOLDER_SCOPE = "https://graph.microsoft.com/Mail.ReadBasic";
export const MAIL_MESSAGE_SCOPE = "https://graph.microsoft.com/Mail.Read";
export const MAIL_SETTINGS_SCOPE = "https://graph.microsoft.com/MailboxSettings.Read";
export function isMailOperation(operation: { owningSlice: string | null }) {
  return operation.owningSlice?.startsWith("MAIL-") === true;
}

export function checkMailScopes(profile: AnyProfile, scopes: string[] | undefined, operation: { owningSlice: string | null; path: string }) {
  const scope = operation.owningSlice === "MAIL-03" || operation.path.endsWith("/attachments") ? MAIL_MESSAGE_SCOPE : operation.owningSlice === "MAIL-04" ? MAIL_SETTINGS_SCOPE : MAIL_FOLDER_SCOPE;
  if (profile.mode === "delegated" && (scopes?.length !== 1 || scopes[0] !== scope)) {
    throw new AxiError(`Mail reads require the delegated ${scope} scope`, "VALIDATION_ERROR", ["Use Mail.Read for message show and attachments, MailboxSettings.Read for settings, rules and categories, and Mail.ReadBasic for message lists and folders; shared mailbox scopes are unsupported"]);
  }
}

// Match only authenticated MSAL account metadata; access tokens stay opaque.
// Graph enforces permissions and the administrator's Exchange RBAC scope.
export function checkMailCredential(profile: AnyProfile, credential: Credential | AppCredential, mailbox: string | undefined) {
  if (profile.mode === "application" || mailbox === undefined) return;
  const account = credential as Partial<Credential>;
  if (![account.localAccountId, account.username].some(value => typeof value === "string" && value.toLowerCase() === mailbox.toLowerCase())) {
    throw new AxiError("Delegated mail access is limited to the signed-in user's own mailbox", "POLICY_DENIED", ["Use --mailbox me or the signed-in user's ID or UPN"]);
  }
}
