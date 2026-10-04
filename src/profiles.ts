import { mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { AxiError } from "axi-sdk-js";

// WRITE-00 hand-edited write opt-in. Absent or { allowWrites: false } keeps
// the profile read-only. No CLI command writes this object; it is added by
// editing the configuration file, and later named writes bind to it.
export type WritePolicy = Readonly<{
  allowWrites: boolean;
  operations: readonly string[];
}>;
export type DelegatedProfile = Readonly<{
  mode: "delegated"; tenantId: string; clientId: string; cloud: "commercial";
  enabledPacks: readonly "entra"[]; preview: boolean; sensitiveAreas: readonly string[];
  credentialRef: Readonly<{ provider: "os-or-session"; key: string }>;
  allowDeviceCode: boolean;
  writes?: WritePolicy;
}>;
export type ApplicationProfile = Readonly<{
  mode: "application"; tenantId: string; clientId: string; cloud: "commercial";
  enabledPacks: readonly "entra"[]; preview: boolean; sensitiveAreas: readonly string[];
  credentialRef: Readonly<
    | { provider: "certificate"; key: string; thumbprint: string }
    | { provider: "federated"; key: string }
  >;
  allowDeviceCode: false;
  writes?: WritePolicy;
}>;
export type AnyProfile = DelegatedProfile | ApplicationProfile;
export type AppCredential = { token: string; expiresAt: number; tenantId: string; clientId: string };
type Config = { version: 1; defaultProfile?: string; profiles: Record<string, AnyProfile> };
const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const namePattern = /^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$/;
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === "object" && !Array.isArray(value);
const exact = (value: Record<string, unknown>, keys: string[]) => Object.keys(value).every(key => keys.includes(key));
function invalid(message: string): never {
  throw new AxiError(message, "VALIDATION_ERROR", ["mg-axi profile create --help", "Correct the version-1 profile configuration; no automatic migration is performed"]);
}

const thumbprint = /^[0-9a-f]{40}$/i;
const sharedKeys = ["mode", "tenantId", "clientId", "cloud", "enabledPacks", "preview", "sensitiveAreas", "credentialRef", "allowDeviceCode", "writes"];
const MAX_WRITE_OPERATIONS = 64;
const MAX_WRITE_OPERATION_LENGTH = 256;
function validateWrites(value: unknown): WritePolicy | undefined {
  if (value === undefined) return undefined;
  if (!object(value) || !exact(value, ["allowWrites", "operations"]) || typeof value.allowWrites !== "boolean" ||
    !Array.isArray(value.operations) || value.operations.length === 0 || value.operations.length > MAX_WRITE_OPERATIONS ||
    value.operations.some(entry => typeof entry !== "string" || !entry.length || entry.length > MAX_WRITE_OPERATION_LENGTH)) {
    invalid("Invalid write policy; hand-edit a writes object with boolean allowWrites and a nonempty operations array of nonempty operation names");
  }
  return Object.freeze({ allowWrites: value.allowWrites as boolean, operations: Object.freeze([...(value.operations as string[])]) });
}
const sharedIdentity = (value: Record<string, unknown>) =>
  value.cloud === "commercial" && typeof value.tenantId === "string" && guid.test(value.tenantId) &&
  typeof value.clientId === "string" && guid.test(value.clientId) && typeof value.preview === "boolean" &&
  Array.isArray(value.enabledPacks) && value.enabledPacks.every(pack => pack === "entra") && new Set(value.enabledPacks).size === value.enabledPacks.length &&
  Array.isArray(value.sensitiveAreas) && value.sensitiveAreas.length === 0;

export function validateDelegatedProfile(value: unknown): DelegatedProfile {
  if (!object(value) || !exact(value, sharedKeys) || value.mode !== "delegated" || !sharedIdentity(value) || typeof value.allowDeviceCode !== "boolean" ||
    !object(value.credentialRef) || !exact(value.credentialRef, ["provider", "key"]) || value.credentialRef.provider !== "os-or-session" ||
    typeof value.credentialRef.key !== "string" || !guid.test(value.credentialRef.key)) invalid("Invalid delegated profile; explicit tenant/client IDs and commercial cloud are required, with credentials only by reference");
  const profile = value as unknown as DelegatedProfile;
  const writes = validateWrites((value as Record<string, unknown>).writes);
  return Object.freeze({ mode: "delegated" as const, tenantId: profile.tenantId, clientId: profile.clientId, cloud: "commercial" as const, preview: profile.preview,
    allowDeviceCode: profile.allowDeviceCode, enabledPacks: Object.freeze([...profile.enabledPacks]), sensitiveAreas: Object.freeze([] as string[]),
    credentialRef: Object.freeze({ provider: "os-or-session" as const, key: profile.credentialRef.key }), ...(writes === undefined ? {} : { writes }) });
}

export function validateApplicationProfile(value: unknown): ApplicationProfile {
  if (!object(value) || !exact(value, sharedKeys) || value.mode !== "application" || !sharedIdentity(value) || value.allowDeviceCode !== false ||
    !object(value.credentialRef)) invalid("Invalid application profile; explicit tenant/client IDs and commercial cloud are required, with credentials only by reference");
  const profile = value as unknown as Omit<ApplicationProfile, "credentialRef"> & { credentialRef: Record<string, unknown> };
  const ref = profile.credentialRef;
  const writes = validateWrites((value as Record<string, unknown>).writes);
  const writeFields = writes === undefined ? {} : { writes };
  if (ref.provider === "certificate" && exact(ref, ["provider", "key", "thumbprint"]) && typeof ref.key === "string" && guid.test(ref.key) &&
    typeof ref.thumbprint === "string" && thumbprint.test(ref.thumbprint)) {
    return Object.freeze({ mode: "application" as const, tenantId: profile.tenantId, clientId: profile.clientId, cloud: "commercial" as const,
      preview: profile.preview, allowDeviceCode: false as const, enabledPacks: Object.freeze([...profile.enabledPacks]), sensitiveAreas: Object.freeze([] as string[]),
      credentialRef: Object.freeze({ provider: "certificate" as const, key: ref.key, thumbprint: ref.thumbprint.toLowerCase() }), ...writeFields });
  }
  if (ref.provider === "federated" && exact(ref, ["provider", "key"]) && typeof ref.key === "string" && guid.test(ref.key)) {
    return Object.freeze({ mode: "application" as const, tenantId: profile.tenantId, clientId: profile.clientId, cloud: "commercial" as const,
      preview: profile.preview, allowDeviceCode: false as const, enabledPacks: Object.freeze([...profile.enabledPacks]), sensitiveAreas: Object.freeze([] as string[]),
      credentialRef: Object.freeze({ provider: "federated" as const, key: ref.key }), ...writeFields });
  }
  invalid("Invalid application profile; credentialRef must be a certificate or federated reference, never inlined key material");
}

export function validateProfile(value: unknown): AnyProfile {
  if (!!value && typeof value === "object" && !Array.isArray(value) && (value as Record<string, unknown>).mode === "application") return validateApplicationProfile(value);
  return validateDelegatedProfile(value);
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
    const profiles: Record<string, AnyProfile> = Object.create(null);
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
  create(name: string, tenantId: string, clientId: string, cloud: string, allowDeviceCode: boolean, application?: { certificateThumbprint?: string; federated?: boolean }) {
    if (!namePattern.test(name)) invalid("Profile name must contain only letters, digits, underscores or hyphens");
    const certificate = application?.certificateThumbprint;
    const federated = application?.federated ?? false;
    if (application && (certificate !== undefined && (typeof certificate !== "string" || !/^[0-9a-f]{40}$/i.test(certificate)))) invalid("Application certificate profiles name a 40-hex-digit thumbprint; the private key stays in protected storage");
    if (application && ((certificate === undefined) === !federated)) invalid("Application profiles use exactly one credential: --certificate-thumbprint or --federated");
    if (application && allowDeviceCode) invalid("Application profiles never use device code; it is a delegated login method");
    const profile = application
      ? validateApplicationProfile({ mode: "application", tenantId, clientId, cloud, enabledPacks: ["entra"], preview: false, sensitiveAreas: [],
        credentialRef: certificate === undefined ? { provider: "federated", key: randomUUID() } : { provider: "certificate", key: randomUUID(), thumbprint: certificate }, allowDeviceCode: false })
      : validateDelegatedProfile({ mode: "delegated", tenantId, clientId, cloud, enabledPacks: ["entra"], preview: false, sensitiveAreas: [], credentialRef: { provider: "os-or-session", key: randomUUID() }, allowDeviceCode });
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
