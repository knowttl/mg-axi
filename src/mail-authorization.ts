import { AxiError } from "axi-sdk-js";
import type { Credential } from "./auth.js";
import type { AnyProfile, AppCredential } from "./profiles.js";

export const MAIL_FOLDER_SCOPE = "https://graph.microsoft.com/Mail.ReadBasic";
export function isMailOperation(operation: { owningSlice: string | null }) {
  return operation.owningSlice?.startsWith("MAIL-") === true;
}

export function checkMailScopes(profile: AnyProfile, scopes: string[] | undefined) {
  if (profile.mode === "delegated" && (scopes?.length !== 1 || scopes[0] !== MAIL_FOLDER_SCOPE)) {
    throw new AxiError("Folder reads require the delegated Mail.ReadBasic scope", "VALIDATION_ERROR", ["Use Mail.ReadBasic for folder metadata; shared mailbox scopes are unsupported"]);
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
