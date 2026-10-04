import { mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { AxiError } from "axi-sdk-js";

export type DelegatedProfile = Readonly<{
  mode: "delegated"; tenantId: string; clientId: string; cloud: "commercial";
  enabledPacks: readonly "entra"[]; preview: boolean; sensitiveAreas: readonly string[];
  credentialRef: Readonly<{ provider: "os-or-session"; key: string }>;
  allowDeviceCode: boolean;
}>;
type Config = { version: 1; defaultProfile?: string; profiles: Record<string, DelegatedProfile> };
const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const namePattern = /^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$/;
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === "object" && !Array.isArray(value);
const exact = (value: Record<string, unknown>, keys: string[]) => Object.keys(value).every(key => keys.includes(key));
function invalid(message: string): never {
  throw new AxiError(message, "VALIDATION_ERROR", ["mg-axi profile create --help", "Correct the version-1 profile configuration; no automatic migration is performed"]);
}

export function validateProfile(value: unknown): DelegatedProfile {
  if (!object(value) || !exact(value, ["mode", "tenantId", "clientId", "cloud", "enabledPacks", "preview", "sensitiveAreas", "credentialRef", "allowDeviceCode"]) ||
    value.mode !== "delegated" || value.cloud !== "commercial" || typeof value.tenantId !== "string" || !guid.test(value.tenantId) ||
    typeof value.clientId !== "string" || !guid.test(value.clientId) || typeof value.preview !== "boolean" || typeof value.allowDeviceCode !== "boolean" ||
    !Array.isArray(value.enabledPacks) || value.enabledPacks.some(pack => pack !== "entra") || new Set(value.enabledPacks).size !== value.enabledPacks.length ||
    !Array.isArray(value.sensitiveAreas) || value.sensitiveAreas.length !== 0 ||
    !object(value.credentialRef) || !exact(value.credentialRef, ["provider", "key"]) || value.credentialRef.provider !== "os-or-session" ||
    typeof value.credentialRef.key !== "string" || !guid.test(value.credentialRef.key)) invalid("Invalid delegated profile; explicit tenant/client IDs and commercial cloud are required, with credentials only by reference");
  return Object.freeze({ mode: "delegated", tenantId: value.tenantId, clientId: value.clientId, cloud: "commercial", preview: value.preview,
    allowDeviceCode: value.allowDeviceCode, enabledPacks: Object.freeze([...value.enabledPacks] as "entra"[]), sensitiveAreas: Object.freeze([]),
    credentialRef: Object.freeze({ provider: "os-or-session", key: value.credentialRef.key }) });
}

export class Profiles {
  constructor(readonly path = process.env.MG_AXI_CONFIG ?? join(homedir(), ".mg-axi", "config.json")) {}
  private load(): Config {
    let value: unknown;
    try { value = JSON.parse(readFileSync(this.path, "utf8")); }
    catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return { version: 1, profiles: {} };
      invalid("Unable to read profile configuration; check the file permissions and version-1 JSON schema");
    }
    if (!object(value) || value.version !== 1 || !exact(value, ["version", "defaultProfile", "profiles"]) || !object(value.profiles)) invalid("Unsupported profile configuration; explicitly migrate to version 1");
    const profiles: Record<string, DelegatedProfile> = Object.create(null);
    for (const [name, profile] of Object.entries(value.profiles)) {
      if (!namePattern.test(name)) invalid("Invalid profile name");
      profiles[name] = validateProfile(profile);
    }
    if (value.defaultProfile !== undefined && (typeof value.defaultProfile !== "string" || !Object.hasOwn(profiles, value.defaultProfile))) invalid("Default profile does not exist");
    if (new Set(Object.values(profiles).map(profile => profile.credentialRef.key)).size !== Object.keys(profiles).length) invalid("Each profile must have a separate credential reference");
    return { version: 1, profiles, defaultProfile: value.defaultProfile as string | undefined };
  }
  list() { return Object.entries(this.load().profiles).map(([name, profile]) => ({ name, mode: profile.mode, tenant: profile.tenantId, cloud: profile.cloud })); }
  resolve(name?: string) {
    const config = this.load();
    const selected = name ?? config.defaultProfile;
    if (!selected || !Object.hasOwn(config.profiles, selected)) throw new AxiError("No configured profile selected", "AUTH_REQUIRED", ["mg-axi profile list", "mg-axi profile create --help"]);
    return { name: selected, profile: config.profiles[selected]! };
  }
  create(name: string, tenantId: string, clientId: string, cloud: string, allowDeviceCode: boolean) {
    if (!namePattern.test(name)) invalid("Profile name must contain only letters, digits, underscores or hyphens");
    const profile = validateProfile({ mode: "delegated", tenantId, clientId, cloud, enabledPacks: ["entra"], preview: false, sensitiveAreas: [], credentialRef: { provider: "os-or-session", key: randomUUID() }, allowDeviceCode });
    const config = this.load();
    if (Object.hasOwn(config.profiles, name)) invalid("Profile already exists; creation never replaces identity");
    config.profiles[name] = profile;
    config.defaultProfile ??= name;
    mkdirSync(dirname(this.path), { recursive: true, mode: 0o700 });
    const temporary = `${this.path}.${randomUUID()}.tmp`;
    try {
      writeFileSync(temporary, `${JSON.stringify(config, null, 2)}\n`, { mode: 0o600, flag: "wx" });
      renameSync(temporary, this.path);
    } catch { throw new AxiError("Unable to save profile configuration", "CONFIG_ERROR", ["Check configuration directory permissions"]); }
    finally { rmSync(temporary, { force: true }); }
    return { name, mode: profile.mode, tenant: profile.tenantId, cloud: profile.cloud };
  }
}
