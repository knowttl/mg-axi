import { rmSync } from "node:fs";
import { AxiError } from "axi-sdk-js";
import { statePath } from "./paths.js";
import { PrivateFiles } from "./private-file.js";
import { profileName, type DelegatedProfile } from "./profiles.js";
import { registerCache } from "./redact.js";

type StoredCache = { tenantId: string; clientId: string; cache: string };
type Store = { version: 1; profiles: Record<string, StoredCache> };
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === "object" && !Array.isArray(value);

// The only persistent credential backend. Access tokens never leave memory.
export class Credentials {
  constructor(readonly path = statePath("credentials.json"), readonly files = new PrivateFiles()) {}
  private load(verifyAcl = false): Store {
    let text: string;
    try { text = this.files.read(this.path, true, verifyAcl); }
    catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return { version: 1, profiles: Object.create(null) };
      throw error;
    }
    try {
      const value: unknown = JSON.parse(text);
      if (!object(value) || value.version !== 1 || !object(value.profiles) || Object.keys(value).some(key => !["version", "profiles"].includes(key))) throw new Error("Invalid store");
      for (const [name, entry] of Object.entries(value.profiles)) {
        if (!/^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$/.test(name) || !object(entry) || typeof entry.tenantId !== "string" || typeof entry.clientId !== "string" || typeof entry.cache !== "string" ||
          Object.keys(entry).some(key => !["tenantId", "clientId", "cache"].includes(key)) || !object(JSON.parse(entry.cache))) throw new Error("Invalid cache");
      }
      return value as unknown as Store;
    } catch (error) {
      if (error instanceof AxiError) throw error;
      throw new AxiError(`Invalid credential store ${this.path}`, "CONFIG_INVALID", ["Restore a valid protected credentials.json before signing in again"]);
    }
  }
  private name(profile: DelegatedProfile) {
    const name = profileName(profile);
    if (!name) throw new AxiError("Stored credentials require a named profile", "CONFIG_INVALID");
    return name;
  }
  read(profile: DelegatedProfile, verifyAcl = false): string | undefined {
    const records = this.load(verifyAcl).profiles;
    const name = this.name(profile);
    const entry = Object.hasOwn(records, name) ? records[name] : undefined;
    if (!entry) return undefined;
    if (entry.tenantId.toLowerCase() !== profile.tenantId.toLowerCase() || entry.clientId.toLowerCase() !== profile.clientId.toLowerCase()) throw new AxiError("Stored credential binding does not match the selected tenant and client", "AUTH_REQUIRED", ["mg-axi auth login --profile <name> --scopes <comma-separated-Graph-scopes>"]);
    return registerCache(entry.cache);
  }
  write(profile: DelegatedProfile, serialized: string) {
    registerCache(serialized);
    const cache = JSON.parse(serialized) as Record<string, unknown>;
    delete cache.AccessToken;
    const store = this.load();
    Object.defineProperty(store.profiles, this.name(profile), { value: { tenantId: profile.tenantId, clientId: profile.clientId, cache: JSON.stringify(cache) }, enumerable: true, configurable: true, writable: true });
    this.files.write(this.path, `${JSON.stringify(store, null, 2)}\n`);
  }
  remove(profile: DelegatedProfile) {
    const store = this.load();
    const name = this.name(profile);
    if (!Object.hasOwn(store.profiles, name)) return;
    delete store.profiles[name];
    if (Object.keys(store.profiles).length) this.files.write(this.path, `${JSON.stringify(store, null, 2)}\n`);
    else rmSync(this.path);
  }
}
