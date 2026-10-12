import { ConfidentialClientApplication, type AuthenticationResult } from "@azure/msal-node";
import { AxiError } from "axi-sdk-js";
import type { ApplicationCredentialProvider } from "./app-auth.js";
import { graphAudience } from "./app-auth.js";
import type { ApplicationProfile } from "./profiles.js";
import { PrivateFiles } from "./private-file.js";
import { resolveSecret } from "./secrets.js";
import { registerSecret } from "./redact.js";

// Application acquisition has no interactive or persistent credential fallback.
export class MsalApplicationProvider implements ApplicationCredentialProvider {
  storage = "session-only" as const;
  constructor(private files = new PrivateFiles()) {}
  async acquire(profile: ApplicationProfile, scopes: string[]) {
    const configuration = {
      auth: { clientId: profile.clientId, authority: `https://login.microsoftonline.com/${profile.tenantId}`,
        ...(profile.auth === "client-secret"
          ? { clientSecret: resolveSecret(profile.clientSecret, this.files) }
          : profile.auth === "certificate"
            ? { clientCertificate: { thumbprintSha256: profile.thumbprintSha256, privateKey: this.privateKey(profile) } }
            : { clientAssertion: () => Promise.resolve(this.assertion(profile)) }) },
      system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } },
    };
    const result: AuthenticationResult | null = await new ConfidentialClientApplication(configuration)
      .acquireTokenByClientCredential({ scopes: scopes.length ? scopes : [graphAudience(profile)] });
    if (!result) throw new Error("Empty application token response");
    if (result.authority.toLowerCase() !== `${configuration.auth.authority.toLowerCase()}/`) throw new Error("Configured application authority mismatch");
    return { token: registerSecret(result.accessToken), expiresAt: result.expiresOn?.getTime() ?? NaN, tenantId: profile.tenantId, clientId: profile.clientId };
  }
  private assertion(profile: Extract<ApplicationProfile, { auth: "federated" }>) {
    if ("env" in profile.token && profile.token.env === "MG_AXI_TOKEN" && !process.env.MG_AXI_TOKEN?.trim()) {
      const legacy = process.env.AZURE_FEDERATED_TOKEN_FILE?.trim();
      if (legacy) {
        process.stderr.write("AZURE_FEDERATED_TOKEN_FILE is deprecated; use MG_AXI_TOKEN instead\n");
        if (!this.files.absolute(legacy)) throw new AxiError("AZURE_FEDERATED_TOKEN_FILE must name an absolute file path", "CONFIG_INVALID", ["Set MG_AXI_TOKEN to the workload assertion"]);
        return resolveSecret({ file: legacy }, this.files);
      }
    }
    return resolveSecret(profile.token, this.files);
  }
  private privateKey(profile: Extract<ApplicationProfile, { auth: "certificate" }>) {
    const material = resolveSecret(profile.privateKey, this.files);
    const normalized = material.includes("\n") ? material : material.replace(/\\n/g, "\n");
    if (/ENCRYPTED PRIVATE KEY/.test(normalized) || /Proc-Type:\s*4,ENCRYPTED/.test(normalized)) throw new AxiError("Referenced privateKey is passphrase-protected; passphrase-protected keys are unsupported", "AUTH_REQUIRED", ["Provision an unencrypted PEM through the referenced environment variable or file"]);
    return registerSecret(normalized);
  }
}
