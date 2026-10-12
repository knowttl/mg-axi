import { AxiError } from "axi-sdk-js";
import { Credentials } from "./credentials.js";
import type { AnyProfile } from "./profiles.js";
import { resolveSecret, secretSource, type SecretRef } from "./secrets.js";

export function authStatus(name: string, profile: AnyProfile, credentials = new Credentials()) {
  const fields: Record<string, SecretRef> = profile.mode === "delegated" ? { token: profile.token }
    : profile.auth === "certificate" ? { privateKey: profile.privateKey }
      : profile.auth === "client-secret" ? { clientSecret: profile.clientSecret } : { token: profile.token };
  const sources = Object.entries(fields).map(([field, ref]) => {
    let available = false;
    try {
      available = "stored" in ref && profile.mode === "delegated" ? credentials.read(profile, true) !== undefined : !!resolveSecret(ref, undefined, true);
    } catch (error) {
      if (!(error instanceof AxiError) || error.code !== "AUTH_REQUIRED") throw error;
    }
    return { field, source: "stored" in ref ? `stored:${credentials.path}` : secretSource(ref), available };
  });
  return { profile: name, mode: profile.mode, endpoint: "https://graph.microsoft.com", tenantId: profile.tenantId, clientId: profile.clientId, sources };
}
