import { readFileSync } from "node:fs";
import { ConfidentialClientApplication, type AuthenticationResult } from "@azure/msal-node";
import type { ApplicationCredentialProvider } from "./app-auth.js";
import { graphAudience } from "./app-auth.js";
import type { ApplicationProfile } from "./profiles.js";

// Only this adapter touches application credentials; keys and assertions are referenced, never logged.
export class MsalApplicationProvider implements ApplicationCredentialProvider {
  storage: "os-protected" | "session-only" = "session-only";
  async acquire(profile: ApplicationProfile, scopes: string[]) {
    const configuration = {
      auth: { clientId: profile.clientId, authority: `https://login.microsoftonline.com/${profile.tenantId}`,
        ...(profile.credentialRef.provider === "certificate"
          ? { clientCertificate: { thumbprint: profile.credentialRef.thumbprint, privateKey: await this.privateKey(profile.credentialRef.key) } }
          : { clientAssertion: () => this.assertion() }) },
      system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } },
    };
    const result: AuthenticationResult | null = await new ConfidentialClientApplication(configuration)
      .acquireTokenByClientCredential({ scopes: scopes.length ? scopes : [graphAudience(profile)] });
    if (!result) throw new Error("Empty application token response");
    return { token: result.accessToken, expiresAt: result.expiresOn?.getTime() ?? NaN, tenantId: result.tenantId, clientId: profile.clientId };
  }
  private async privateKey(key: string) {
    const store = await import("keytar").then(module => module.default).catch(() => undefined);
    this.storage = store ? "os-protected" : "session-only";
    const material = await store?.getPassword("mg-axi", key).catch(() => null);
    if (typeof material !== "string" || !material.trim()) throw new Error("Missing referenced certificate private key");
    return material;
  }
  private async assertion() {
    this.storage = "session-only";
    const file = process.env.AZURE_FEDERATED_TOKEN_FILE;
    if (!file?.trim()) throw new Error("Missing federated token source");
    // Read on every invocation: projected federated tokens rotate on disk.
    const assertion = readFileSync(file, "utf8").trim();
    if (!assertion) throw new Error("Missing federated token source");
    return assertion;
  }
}
