import { AxiError } from "axi-sdk-js";
import { validateApplicationProfile, type AppCredential, type ApplicationProfile } from "./profiles.js";

// The configured Graph audience is the app's consented application grants, never a per-command subset.
export const graphAudience = (profile: ApplicationProfile) => `https://graph.microsoft.com/.default`;

export interface ApplicationCredentialProvider {
  acquire(profile: ApplicationProfile, scopes: string[]): Promise<AppCredential>;
  storage: "os-protected" | "session-only";
}

// This service is the application credential boundary for CORE-01; tokens never reach CLI output.
// It has no login method: application profiles never invoke interactive login or delegated fallback.
export class ApplicationAuth {
  private cache = new Map<string, AppCredential>();
  constructor(private provider: ApplicationCredentialProvider, private now: () => number = Date.now) {}
  async credential(profile: ApplicationProfile, scopes: string[] = [graphAudience(profile)]) {
    profile = validateApplicationProfile(profile);
    const requested = this.audience(profile, scopes);
    const key = this.key(profile, requested);
    const hit = this.cache.get(key);
    if (hit && hit.expiresAt > this.now() + 60_000) return { ...hit };
    this.cache.delete(key);
    try {
      const credential = this.check(profile, await this.provider.acquire(profile, requested));
      // One identity per profile: a replacement wipes rather than merges.
      this.cache.clear();
      this.cache.set(key, credential);
      return { ...credential };
    } catch (error) {
      // Provider AUTH_REQUIRED failures already name only the credential
      // reference, never its value; everything else stays fully generic.
      if (error instanceof AxiError && error.code === "AUTH_REQUIRED") throw error;
      throw this.failure();
    }
  }
  private audience(profile: ApplicationProfile, scopes: string[]) {
    const configured = [graphAudience(profile)];
    if (scopes.length !== 1 || scopes[0] !== configured[0]) throw new AxiError("Application profiles use the configured Graph .default audience; delegated scopes are unavailable", "VALIDATION_ERROR", ["mg-axi profile show --profile <name>", "Application permissions apply to the app registration, not to individual commands"]);
    return configured;
  }
  private key(profile: ApplicationProfile, scopes: string[]) { return JSON.stringify([profile.credentialRef, profile.tenantId, profile.clientId, profile.cloud, scopes]); }
  private check(profile: ApplicationProfile, credential: AppCredential) {
    if (!credential.token || credential.tenantId.toLowerCase() !== profile.tenantId.toLowerCase() || credential.clientId.toLowerCase() !== profile.clientId.toLowerCase() || !Number.isFinite(credential.expiresAt) || credential.expiresAt <= this.now() + 60_000) throw new Error("Invalid credential context");
    return { ...credential };
  }
  private failure() { return new AxiError("Application authentication failed for the configured identity; no user or device-code fallback was attempted", "AUTH_REQUIRED", ["mg-axi profile show --profile <name>", "Ask an administrator to grant application consent for the Graph .default audience on this app registration", "Confirm the certificate thumbprint or federated credential mapping on the app registration"]); }
}
