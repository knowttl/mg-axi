import { PublicClientApplication, type AuthenticationResult } from "@azure/msal-node";
import open from "open";
import type { CredentialProvider, LoginMethod } from "./auth.js";
import type { DelegatedProfile } from "./profiles.js";

// Only this adapter touches Microsoft authentication and OS credential storage.
export class MsalProvider implements CredentialProvider {
  storage: "os-protected" | "session-only" = "session-only";
  private sessions = new Map<string, { app: PublicClientApplication; store?: typeof import("keytar") }>();
  private async application(profile: DelegatedProfile, fresh = false) {
    const key = JSON.stringify([profile.credentialRef.key, profile.tenantId, profile.clientId, profile.cloud]);
    const existing = this.sessions.get(key);
    if (existing && !fresh) {
      this.storage = existing.store ? "os-protected" : "session-only";
      return existing.app;
    }
    const store = existing ? existing.store : await import("keytar").then(module => module.default).catch(() => undefined);
    this.storage = store ? "os-protected" : "session-only";
    if (fresh) {
      this.sessions.delete(key);
      if (store) {
        await store.deletePassword("mg-axi", key);
        if (await store.getPassword("mg-axi", key) !== null) throw new Error("Unable to wipe profile cache");
      }
    }
    const configuration = {
      auth: { clientId: profile.clientId, authority: `https://login.microsoftonline.com/${profile.tenantId}` },
      system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } },
    };
    const app = new PublicClientApplication({
      ...configuration,
      cache: store ? { cachePlugin: {
        beforeCacheAccess: async context => {
          const saved = await store.getPassword("mg-axi", key);
          if (saved) context.tokenCache.deserialize(saved);
        },
        afterCacheAccess: async context => {
          if (!context.cacheHasChanged) return;
          const saved = context.tokenCache.serialize();
          if (process.platform === "win32" && Buffer.byteLength(saved, "utf8") > 2560) throw new Error("OS credential store cannot persist this cache");
          await store.setPassword("mg-axi", key, saved);
        },
      } } : undefined,
    });
    this.sessions.set(key, { app, store });
    return app;
  }
  async login(profile: DelegatedProfile, method: LoginMethod, scopes: string[]) {
    // Each explicit login starts with a new client and a verified empty protected cache.
    const app = await this.application(profile, true);
    const result = method === "browser"
      ? await app.acquireTokenInteractive({ scopes, openBrowser: async url => { await open(url); }, prompt: "select_account", errorTemplate: "Sign-in failed. Return to mg-axi for recovery guidance." })
      : await app.acquireTokenByDeviceCode({ scopes, deviceCodeCallback: response => {
        if (typeof response.userCode !== "string" || !response.userCode.trim() || typeof response.deviceCode !== "string" || !response.deviceCode.trim() || !Number.isFinite(response.expiresIn) || response.expiresIn <= 0) throw new Error("Invalid device-code challenge");
        // The one-time user challenge belongs only to explicit login stderr.
        process.stderr.write(`Open https://microsoft.com/devicelogin and enter ${response.userCode}\n`);
      } });
    return this.result(profile, result);
  }
  async silent(profile: DelegatedProfile, scopes: string[]) {
    const app = await this.application(profile);
    const accounts = (await app.getTokenCache().getAllAccounts()).filter(account => account.tenantId.toLowerCase() === profile.tenantId.toLowerCase());
    if (accounts.length !== 1) throw new Error("Explicit login required");
    return this.result(profile, await app.acquireTokenSilent({ scopes, account: accounts[0]! }));
  }
  private result(profile: DelegatedProfile, result: AuthenticationResult | null) {
    if (!result?.account || result.tenantId.toLowerCase() !== profile.tenantId.toLowerCase()) throw new Error("Configured tenant mismatch");
    return { token: result.accessToken, expiresAt: result.expiresOn?.getTime() ?? NaN, tenantId: result.tenantId, clientId: profile.clientId, accountId: result.account.homeAccountId };
  }
}
