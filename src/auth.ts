import { AxiError } from "axi-sdk-js";
import { validateProfile, type DelegatedProfile } from "./profiles.js";

export type LoginMethod = "browser" | "device-code";
export type Credential = { token: string; expiresAt: number; tenantId: string; clientId: string; accountId: string };
export interface CredentialProvider {
  login(profile: DelegatedProfile, method: LoginMethod, scopes: string[]): Promise<Credential>;
  silent(profile: DelegatedProfile, scopes: string[]): Promise<Credential>;
  storage: "os-protected" | "session-only";
}

// This service is the credential boundary for CORE-01; tokens never reach CLI output.
export class DelegatedAuth {
  private cache = new Map<string, Credential>();
  constructor(private provider: CredentialProvider, private now: () => number = Date.now) {}
  async login(profile: DelegatedProfile, method: string = "browser", scopes: string[] = []) {
    profile = validateProfile(profile);
    if (method !== "browser" && method !== "device-code") throw new AxiError("Login method must be browser or device-code", "VALIDATION_ERROR", ["mg-axi login --help"]);
    if (method === "device-code" && !profile.allowDeviceCode) throw new AxiError("Device code is disabled for this profile", "VALIDATION_ERROR", ["Enable device code only when organization policy permits it; browser login is the default"]);
    const requested = this.scopes(scopes);
    const key = this.key(profile, requested);
    this.cache.clear();
    try {
      const credential = this.check(profile, await this.provider.login(profile, method, requested));
      this.cache.set(key, credential);
      return { status: "authenticated", mode: "delegated", storage: this.provider.storage };
    } catch { throw this.failure("LOGIN_FAILED"); }
  }
  async credential(profile: DelegatedProfile, scopes: string[]) {
    profile = validateProfile(profile);
    const requested = this.scopes(scopes);
    const key = this.key(profile, requested);
    const hit = this.cache.get(key);
    if (hit && hit.expiresAt > this.now() + 60_000) return { ...hit };
    this.cache.delete(key);
    try {
      const credential = this.check(profile, await this.provider.silent(profile, requested));
      this.cache.set(key, credential);
      return { ...credential };
    } catch { throw this.failure("AUTH_REQUIRED"); }
  }
  private scopes(scopes: string[]) {
    if (!scopes.length || scopes.some(scope => !/^https:\/\/graph\.microsoft\.com\/[A-Za-z][A-Za-z.]+$/.test(scope) || scope.endsWith("/.default"))) throw new AxiError("Explicit delegated Graph scopes are required", "VALIDATION_ERROR", ["mg-axi login --profile <name> --scopes <comma-separated-Graph-scopes>"]);
    return [...new Set(scopes)].sort();
  }
  private key(profile: DelegatedProfile, scopes: string[]) { return JSON.stringify([profile.credentialRef, profile.tenantId, profile.clientId, profile.cloud, scopes]); }
  private check(profile: DelegatedProfile, credential: Credential) {
    if (!credential.token || !credential.accountId || credential.tenantId.toLowerCase() !== profile.tenantId.toLowerCase() || credential.clientId.toLowerCase() !== profile.clientId.toLowerCase() || !Number.isFinite(credential.expiresAt) || credential.expiresAt <= this.now() + 60_000) throw new Error("Invalid credential context");
    return { ...credential };
  }
  private failure(code: string) { return new AxiError("Authentication unavailable for the configured identity; no interactive fallback was attempted", code, ["mg-axi login --profile <name> --scopes <comma-separated-Graph-scopes>", "Check dedicated app consent and organization sign-in policy"]); }
}
