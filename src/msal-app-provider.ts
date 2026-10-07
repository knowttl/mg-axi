import { lstatSync, readFileSync } from "node:fs";
import { ConfidentialClientApplication, type AuthenticationResult } from "@azure/msal-node";
import { AxiError } from "axi-sdk-js";
import type { ApplicationCredentialProvider } from "./app-auth.js";
import { graphAudience } from "./app-auth.js";
import type { ApplicationProfile } from "./profiles.js";

type CertificateRef = Extract<ApplicationProfile["credentialRef"], { provider: "certificate" }>;
type SecretRef = Extract<ApplicationProfile["credentialRef"], { provider: "client-secret" }>;
type FederatedRef = Extract<ApplicationProfile["credentialRef"], { provider: "federated" }>;

// Only this adapter touches application credentials; keys and assertions are referenced, never logged.
export class MsalApplicationProvider implements ApplicationCredentialProvider {
  storage: "os-protected" | "session-only" = "session-only";
  async acquire(profile: ApplicationProfile, scopes: string[]) {
    const ref = profile.credentialRef;
    const configuration = {
      auth: { clientId: profile.clientId, authority: `https://login.microsoftonline.com/${profile.tenantId}`,
        ...(ref.provider === "client-secret"
          ? { clientSecret: this.clientSecret(ref) }
          : ref.provider === "certificate"
            ? { clientCertificate: { thumbprint: ref.thumbprint, privateKey: await this.privateKey(ref) } }
            : { clientAssertion: () => this.assertion(ref) }) },
      system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } },
    };
    const result: AuthenticationResult | null = await new ConfidentialClientApplication(configuration)
      .acquireTokenByClientCredential({ scopes: scopes.length ? scopes : [graphAudience(profile)] });
    if (!result) throw new Error("Empty application token response");
    if (result.authority.toLowerCase() !== `${configuration.auth.authority.toLowerCase()}/`) throw new Error("Configured application authority mismatch");
    return { token: result.accessToken, expiresAt: result.expiresOn?.getTime() ?? NaN, tenantId: profile.tenantId, clientId: profile.clientId };
  }
  private async privateKey(ref: CertificateRef) {
    if (ref.keyFile !== undefined) return this.normalizeKey(this.readProtectedFile(ref.keyFile), ref.keyFile);
    if (ref.keyEnv !== undefined) {
      const material = process.env[ref.keyEnv];
      if (typeof material !== "string" || !material.trim()) throw this.missing(ref.keyEnv);
      return this.normalizeKey(material, ref.keyEnv);
    }
    const store = await import("keytar").then(module => module.default).catch(() => undefined);
    this.storage = store ? "os-protected" : "session-only";
    const material = await store?.getPassword("mg-axi", ref.key).catch(() => null);
    if (typeof material !== "string" || !material.trim()) throw new Error("Missing referenced certificate private key");
    return material;
  }
  private clientSecret(ref: SecretRef) {
    this.storage = "session-only";
    if (ref.secretEnv !== undefined) {
      const secret = process.env[ref.secretEnv];
      if (typeof secret !== "string" || !secret.trim()) throw this.missing(ref.secretEnv);
      return secret.trim();
    }
    if (ref.secretFile !== undefined) return this.readProtectedFile(ref.secretFile);
    throw new Error("Missing client-secret credential holder");
  }
  private async assertion(ref: FederatedRef) {
    this.storage = "session-only";
    if (ref.tokenFileEnv === undefined) {
      const file = process.env.AZURE_FEDERATED_TOKEN_FILE;
      if (!file?.trim()) throw new Error("Missing federated token source");
      // Read on every invocation: projected federated tokens rotate on disk.
      const assertion = readFileSync(file, "utf8").trim();
      if (!assertion) throw new Error("Missing federated token source");
      return assertion;
    }
    const file = process.env[ref.tokenFileEnv];
    if (!file?.trim()) throw this.missing(ref.tokenFileEnv);
    // Read on every invocation: projected federated tokens rotate on disk.
    try {
      const assertion = readFileSync(file, "utf8").trim();
      if (!assertion) throw this.missing(ref.tokenFileEnv);
      return assertion;
    } catch (error) {
      if (error instanceof AxiError) throw error;
      throw this.missing(ref.tokenFileEnv);
    }
  }
  // Key and secret files hold long-lived material: regular file, owned by the
  // current user, mode 0600, never a symlink. Relative paths resolve against
  // the process working directory at acquire time. Windows skips the
  // ownership and mode checks. Every failure names only the path, never content.
  private readProtectedFile(path: string) {
    try {
      const stat = lstatSync(path);
      if (stat.isSymbolicLink() || !stat.isFile()) throw this.missing(path);
      if (process.platform !== "win32") {
        if (typeof process.getuid === "function" && stat.uid !== process.getuid()) throw this.missing(path);
        if ((stat.mode & 0o777) !== 0o600) throw this.missing(path);
      }
      const content = readFileSync(path, "utf8").trim();
      if (!content) throw this.missing(path);
      return content;
    } catch (error) {
      if (error instanceof AxiError) throw error;
      throw this.missing(path);
    }
  }
  // PEM material from files and env vars; env vars may carry escaped "\n".
  private normalizeKey(material: string, reference: string) {
    const normalized = material.includes("\n") ? material : material.replace(/\\n/g, "\n");
    if (/ENCRYPTED PRIVATE KEY/.test(normalized)) throw new AxiError(`Application credential reference ${reference} is passphrase-protected; passphrase-protected keys are unsupported`, "AUTH_REQUIRED", ["mg-axi profile show --profile <name>", "Provision an unencrypted PEM through the referenced environment variable or file"]);
    if (!normalized.trim()) throw this.missing(reference);
    return normalized;
  }
  // Fail-closed holder error: names only the env var name or file path.
  private missing(reference: string) {
    return new AxiError(`Application credential unavailable for the configured reference ${reference}; no user or device-code fallback was attempted`, "AUTH_REQUIRED", ["mg-axi profile show --profile <name>", "Supply the referenced credential through the named environment variable or file"]);
  }
}
