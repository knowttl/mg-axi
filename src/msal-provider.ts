import { PublicClientApplication, type AuthenticationResult } from "@azure/msal-node";
import open from "open";
import type { CredentialProvider, LoginMethod } from "./auth.js";
import type { DelegatedProfile } from "./profiles.js";

// Only this adapter touches Microsoft authentication and OS credential storage.
export class MsalProvider implements CredentialProvider {
  storage: "os-protected" | "session-only" = "session-only";
  private sessions = new Map<string, { app: PublicClientApplication; storage: "os-protected" | "session-only"; invalidate: () => Promise<void> }>();
  private async application(profile: DelegatedProfile) {
    const key = JSON.stringify([profile.credentialRef.key, profile.tenantId, profile.clientId, profile.cloud]);
    const existing = this.sessions.get(key);
    if (existing) {
      this.storage = existing.storage;
      return existing;
    }
    const store = await import("keytar").then(module => module.default).catch(() => undefined);
    this.storage = store ? "os-protected" : "session-only";
    const app = new PublicClientApplication({
      auth: { clientId: profile.clientId, authority: `https://login.microsoftonline.com/${profile.tenantId}` },
      system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } },
      cache: store ? { cachePlugin: {
        beforeCacheAccess: async context => {
          if (session.storage === "session-only") return;
          const saved = await store.getPassword("mg-axi", key);
          if (saved) context.tokenCache.deserialize(saved);
        },
        afterCacheAccess: async context => {
          if (!context.cacheHasChanged || session.storage === "session-only") return;
          const saved = context.tokenCache.serialize();
          if (process.platform === "win32" && Buffer.byteLength(saved, "utf8") > 2560) {
            await session.invalidate();
            session.storage = this.storage = "session-only";
            return;
          }
          await store.setPassword("mg-axi", key, saved);
        },
      } } : undefined,
    });
    const session = { app, storage: this.storage, invalidate: async () => {
      if (store) await store.deletePassword("mg-axi", key);
    } };
    this.sessions.set(key, session);
    return session;
  }
  async login(profile: DelegatedProfile, method: LoginMethod, scopes: string[]) {
    const session = await this.application(profile);
    const app = session.app;
    // A new explicit login replaces the old account, never silently selects it.
    const accounts = await app.getTokenCache().getAllAccounts();
    await session.invalidate();
    for (const account of accounts) await app.getTokenCache().removeAccount(account);
    const result = method === "browser"
      ? await app.acquireTokenInteractive({ scopes, openBrowser: async url => { await open(url); }, prompt: "select_account", errorTemplate: "Sign-in failed. Return to mg-axi for recovery guidance." })
      : await app.acquireTokenByDeviceCode({ scopes, deviceCodeCallback: response => {
        // The one-time user challenge belongs only to explicit login stderr.
        process.stderr.write(`Open https://microsoft.com/devicelogin and enter ${response.userCode}\n`);
      } });
    return this.result(profile, result);
  }
  async silent(profile: DelegatedProfile, scopes: string[]) {
    const { app } = await this.application(profile);
    const accounts = (await app.getTokenCache().getAllAccounts()).filter(account => account.tenantId.toLowerCase() === profile.tenantId.toLowerCase());
    if (accounts.length !== 1) throw new Error("Explicit login required");
    return this.result(profile, await app.acquireTokenSilent({ scopes, account: accounts[0]! }));
  }
  private result(profile: DelegatedProfile, result: AuthenticationResult | null) {
    if (!result?.account || result.tenantId.toLowerCase() !== profile.tenantId.toLowerCase()) throw new Error("Configured tenant mismatch");
    return { token: result.accessToken, expiresAt: result.expiresOn?.getTime() ?? NaN, tenantId: result.tenantId, clientId: profile.clientId, accountId: result.account.homeAccountId };
  }
}
