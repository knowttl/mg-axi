import { PublicClientApplication, type AuthenticationResult } from "@azure/msal-node";
import open from "open";
import { AxiError } from "axi-sdk-js";
import type { CredentialProvider, LoginMethod } from "./auth.js";
import { profileName, type DelegatedProfile } from "./profiles.js";
import { Credentials } from "./credentials.js";
import { resolveSecret } from "./secrets.js";
import { redactSecrets, registerCache, registerSecret } from "./redact.js";

export class MsalProvider implements CredentialProvider {
  storage: "protected-file" | "session-only" = "protected-file";
  private sessions = new Map<string, PublicClientApplication>();
  constructor(private store = new Credentials()) {}
  private key(profile: DelegatedProfile) { return JSON.stringify([profileName(profile), profile.tenantId, profile.clientId, profile.cloud, profile.token]); }
  private application(profile: DelegatedProfile, login = false) {
    const key = this.key(profile);
    this.storage = "stored" in profile.token ? "protected-file" : "session-only";
    const existing = this.sessions.get(key);
    if (existing && !login) return existing;
    let loaded = false;
    let snapshot: string | undefined;
    const app = new PublicClientApplication({
      auth: { clientId: profile.clientId, authority: `https://login.microsoftonline.com/${profile.tenantId}` },
      system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } },
      cache: { cachePlugin: {
        beforeCacheAccess: async context => {
          if (!loaded) {
            snapshot = "stored" in profile.token ? this.store.read(profile) : resolveSecret(profile.token, this.store.files);
            loaded = true;
          }
          // MSAL overwrites its cache from the plugin snapshot before a silent lookup.
          if (snapshot) context.tokenCache.deserialize(registerCache(snapshot));
        },
        afterCacheAccess: async context => {
          if (!context.cacheHasChanged) return;
          snapshot = registerCache(context.tokenCache.serialize());
          // Hosted commands keep refreshed tokens only in memory. Explicit login owns persistence.
          if (login) this.store.write(profile, snapshot);
        },
      } },
    });
    this.sessions.set(key, app);
    return app;
  }
  async login(profile: DelegatedProfile, method: LoginMethod, scopes: string[]) {
    if (!("stored" in profile.token)) throw new AxiError("auth login requires token: {stored: true}", "CONFIG_INVALID");
    this.sessions.delete(this.key(profile));
    this.store.remove(profile);
    const app = this.application(profile, true);
    try {
      const result = method === "browser"
        ? await app.acquireTokenInteractive({ scopes, openBrowser: async url => { await open(url); }, prompt: "select_account", errorTemplate: "Sign-in failed. Return to mg-axi for recovery guidance." })
        : await app.acquireTokenByDeviceCode({ scopes, deviceCodeCallback: response => {
          if (typeof response.userCode !== "string" || !response.userCode.trim() || typeof response.deviceCode !== "string" || !response.deviceCode.trim() || !Number.isFinite(response.expiresIn) || response.expiresIn <= 0) throw new Error("Invalid device-code challenge");
          process.stderr.write(redactSecrets(`Open https://microsoft.com/devicelogin and enter ${response.userCode}\n`));
        } });
      const credential = this.result(profile, result);
      // Subsequent silent acquisitions must not retain a persistence-enabled plugin.
      this.sessions.delete(this.key(profile));
      return credential;
    } catch (error) {
      this.sessions.delete(this.key(profile));
      this.store.remove(profile);
      throw error;
    }
  }
  async silent(profile: DelegatedProfile, scopes: string[]) {
    const app = this.application(profile);
    const accounts = (await app.getTokenCache().getAllAccounts()).filter(account => account.tenantId.toLowerCase() === profile.tenantId.toLowerCase());
    if (accounts.length !== 1) throw new Error("Explicit login required");
    const credential = this.result(profile, await app.acquireTokenSilent({ scopes, account: accounts[0]! }));
    if (credential.accountId !== accounts[0]!.homeAccountId) throw new Error("Configured account mismatch");
    return credential;
  }
  async logout(profile: DelegatedProfile) {
    this.sessions.delete(this.key(profile));
    if ("stored" in profile.token) this.store.remove(profile);
  }
  private result(profile: DelegatedProfile, result: AuthenticationResult | null) {
    if (!result?.account || result.tenantId.toLowerCase() !== profile.tenantId.toLowerCase()) throw new Error("Configured tenant mismatch");
    return { token: registerSecret(result.accessToken), expiresAt: result.expiresOn?.getTime() ?? NaN, tenantId: result.tenantId, clientId: profile.clientId, accountId: result.account.homeAccountId };
  }
}
