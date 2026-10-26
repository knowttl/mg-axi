import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { AxiError } from "axi-sdk-js";
import { configPath, toolDirectory } from "./paths.js";
import { PrivateFiles } from "./private-file.js";
import { validateSecretRef, type SecretRef } from "./secrets.js";

export type WritePolicy = Readonly<{ allowWrites: boolean; operations: readonly string[] }>;
type SharedProfile = Readonly<{
  tenantId: string; clientId: string; cloud: "commercial";
  enabledPacks: readonly ("entra" | "security")[]; preview: boolean; sensitiveAreas: readonly string[];
  writes?: WritePolicy;
}>;
export type DelegatedProfile = SharedProfile & Readonly<{ mode: "delegated"; token: SecretRef; allowDeviceCode: boolean }>;
export type ApplicationProfile = SharedProfile & Readonly<{ mode: "application"; allowDeviceCode: false }> & (
  | Readonly<{ auth: "certificate"; privateKey: SecretRef; thumbprintSha256: string }>
  | Readonly<{ auth: "client-secret"; clientSecret: SecretRef }>
  | Readonly<{ auth: "federated"; token: SecretRef }>
);
export type AnyProfile = DelegatedProfile | ApplicationProfile;
export type AppCredential = { token: string; expiresAt: number; tenantId: string; clientId: string };
type Config = { version: 1; defaultProfile?: string; profiles: Record<string, AnyProfile> };
const names = new WeakMap<object, string>();
const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const namePattern = /^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$/;
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === "object" && !Array.isArray(value);
const exact = (value: Record<string, unknown>, keys: string[]) => Object.keys(value).every(key => keys.includes(key));
function invalid(message: string): never {
  throw new AxiError(message, "CONFIG_INVALID", ["mg-axi profile create --help", "Correct the version-1 profile configuration; no automatic migration is performed"]);
}

export function booleanEnv(name: string, env: NodeJS.ProcessEnv = process.env): boolean {
  const value = env[name]?.toLowerCase();
  if (value === undefined || value === "0" || value === "false") return false;
  if (value === "1" || value === "true") return true;
  return invalid(`Invalid boolean for ${name}; use 1, true, 0 or false`);
}

// Names belong to the selected configuration, never to its serialized profile schema.
export function profileName(profile: AnyProfile) { return names.get(profile); }
function validated<T extends AnyProfile>(source: object, profile: T): T {
  const name = names.get(source);
  if (name !== undefined) names.set(profile, name);
  return Object.freeze(profile);
}

function shared(value: Record<string, unknown>): SharedProfile {
  if (typeof value.tenantId !== "string" || !guid.test(value.tenantId) || typeof value.clientId !== "string" || !guid.test(value.clientId)) invalid("Explicit tenant and client UUIDs are required");
  if (value.cloud !== undefined && value.cloud !== "commercial") invalid("Only commercial cloud is supported");
  const packs = value.enabledPacks === undefined ? ["entra"] : value.enabledPacks;
  const sensitive = value.sensitiveAreas === undefined ? [] : value.sensitiveAreas;
  const preview = value.preview === undefined ? false : value.preview;
  if (!Array.isArray(packs) || packs.some(pack => pack !== "entra" && pack !== "security") || new Set(packs).size !== packs.length ||
    !Array.isArray(sensitive) || sensitive.length || typeof preview !== "boolean") invalid("Invalid profile policy");
  let writes: WritePolicy | undefined;
  if (value.writes !== undefined) {
    const policy = value.writes;
    if (!object(policy) || !exact(policy, ["allowWrites", "operations"]) || typeof policy.allowWrites !== "boolean" ||
      !Array.isArray(policy.operations) || !policy.operations.length || policy.operations.length > 64 ||
      policy.operations.some(entry => typeof entry !== "string" || !entry.length || entry.length > 256)) invalid("Invalid write policy; use boolean allowWrites and a nonempty operations array");
    writes = Object.freeze({ allowWrites: policy.allowWrites, operations: Object.freeze([...policy.operations] as string[]) });
  }
  return { tenantId: value.tenantId as string, clientId: value.clientId as string, cloud: "commercial", enabledPacks: Object.freeze([...packs] as ("entra" | "security")[]),
    preview: preview as boolean, sensitiveAreas: Object.freeze([]), ...(writes === undefined ? {} : { writes }) };
}
const sharedKeys = ["mode", "tenantId", "clientId", "cloud", "enabledPacks", "preview", "sensitiveAreas", "allowDeviceCode", "writes"];

export function validateDelegatedProfile(value: unknown, files?: PrivateFiles): DelegatedProfile {
  if (!object(value) || value.mode !== "delegated" || !exact(value, [...sharedKeys, "token"]) ||
    (value.allowDeviceCode !== undefined && typeof value.allowDeviceCode !== "boolean")) invalid("Invalid delegated profile; token must be a SecretRef");
  return validated(value, { ...shared(value), mode: "delegated", allowDeviceCode: value.allowDeviceCode === true, token: validateSecretRef(value.token, "token", true, files) });
}

export function validateApplicationProfile(value: unknown, files?: PrivateFiles): ApplicationProfile {
  if (!object(value) || value.mode !== "application" || (value.allowDeviceCode !== undefined && value.allowDeviceCode !== false)) invalid("Invalid application profile; device code is unavailable");
  const keys = value.auth === "certificate" ? ["privateKey", "thumbprintSha256"] : value.auth === "client-secret" ? ["clientSecret"] : value.auth === "federated" ? ["token"] : [];
  if (!keys.length || !exact(value, [...sharedKeys, "auth", ...keys])) invalid("Invalid application profile; auth must be certificate, client-secret or federated with only its credential fields");
  const base = { ...shared(value), mode: "application" as const, allowDeviceCode: false as const };
  if (value.auth === "certificate") {
    if (typeof value.thumbprintSha256 !== "string" || !/^[0-9a-f]{64}$/i.test(value.thumbprintSha256)) invalid("thumbprintSha256 must contain 64 hex digits");
    return validated(value, { ...base, auth: "certificate", thumbprintSha256: (value.thumbprintSha256 as string).toLowerCase(), privateKey: validateSecretRef(value.privateKey, "privateKey", false, files) });
  }
  if (value.auth === "client-secret") return validated(value, { ...base, auth: "client-secret", clientSecret: validateSecretRef(value.clientSecret, "clientSecret", false, files) });
  return validated(value, { ...base, auth: "federated", token: validateSecretRef(value.token, "token", false, files) });
}

export function validateProfile(value: unknown, files?: PrivateFiles): AnyProfile {
  return object(value) && value.mode === "application" ? validateApplicationProfile(value, files) : validateDelegatedProfile(value, files);
}

function envProfile(env: NodeJS.ProcessEnv): AnyProfile | undefined {
  const fields = ["MODE", "TENANT_ID", "CLIENT_ID", "AUTH", "CLOUD", "PREVIEW", "ALLOW_DEVICE_CODE", "ENABLED_PACKS", "SENSITIVE_AREAS", "WRITES"];
  if (!fields.some(field => env[`MG_AXI_${field}`] !== undefined)) return undefined;
  if (!env.MG_AXI_MODE || !env.MG_AXI_TENANT_ID || !env.MG_AXI_CLIENT_ID) invalid("Partial env profile; MG_AXI_MODE, MG_AXI_TENANT_ID and MG_AXI_CLIENT_ID are required");
  const json = (name: string) => {
    if (env[name] === undefined) return undefined;
    try { return JSON.parse(env[name]!); } catch { return invalid(`Invalid JSON field ${name}`); }
  };
  const base = { mode: env.MG_AXI_MODE, tenantId: env.MG_AXI_TENANT_ID, clientId: env.MG_AXI_CLIENT_ID, cloud: env.MG_AXI_CLOUD,
    preview: booleanEnv("MG_AXI_PREVIEW", env), allowDeviceCode: booleanEnv("MG_AXI_ALLOW_DEVICE_CODE", env),
    enabledPacks: json("MG_AXI_ENABLED_PACKS"), sensitiveAreas: json("MG_AXI_SENSITIVE_AREAS"), writes: json("MG_AXI_WRITES"),
    ...(env.MG_AXI_AUTH === undefined ? {} : { auth: env.MG_AXI_AUTH }),
    ...(env.MG_AXI_THUMBPRINT_SHA256 === undefined ? {} : { thumbprintSha256: env.MG_AXI_THUMBPRINT_SHA256 }),
    ...(env.MG_AXI_CLIENT_SECRET === undefined ? {} : { clientSecret: { env: "MG_AXI_CLIENT_SECRET" } }),
    ...(env.MG_AXI_PRIVATE_KEY === undefined ? {} : { privateKey: { env: "MG_AXI_PRIVATE_KEY" } }),
    ...(env.MG_AXI_TOKEN === undefined ? {} : { token: { env: "MG_AXI_TOKEN" } }) };
  if (base.mode === "delegated") {
    if (!env.MG_AXI_TOKEN) invalid("Partial env profile; MG_AXI_TOKEN must supply a delegated cache");
    return validateDelegatedProfile(base);
  }
  const auth = env.MG_AXI_AUTH;
  const required = auth === "certificate" ? ["PRIVATE_KEY", "THUMBPRINT_SHA256"] : auth === "client-secret" ? ["CLIENT_SECRET"] : auth === "federated" ? ["TOKEN"] : [];
  if (!required.length || required.some(field => !env[`MG_AXI_${field}`])) invalid("Partial env profile; MG_AXI_AUTH and its required credential fields must be set");
  return validateApplicationProfile(base);
}

export function movedError(kind: string, legacy: string, target: string, guidance: string): never {
  const quote = (path: string) => `'${path.replace(/'/g, "'\\''")}'`;
  const move = process.platform === "win32"
    ? `New-Item -ItemType Directory -Force -Path '${dirname(target).replace(/'/g, "''")}'; Move-Item -LiteralPath '${legacy.replace(/'/g, "''")}' -Destination '${target.replace(/'/g, "''")}'`
    : `mkdir -p ${quote(dirname(target))} && mv ${quote(legacy)} ${quote(target)}`;
  throw new AxiError(`${kind} moved from ${legacy} to ${target}`, "CONFIG_MOVED", [move, guidance]);
}

export class Profiles {
  readonly path: string;
  constructor(path?: string, private files = new PrivateFiles()) {
    this.path = path?.trim() ? path.trim() : configPath();
  }
  private moved(): never {
    const legacy = join(homedir(), ".mg-axi", "config.json");
    const target = resolve(this.path) === resolve(legacy) ? join(toolDirectory("config"), "config.json") : this.path;
    movedError("Configuration", legacy, target, "Update credentialRef to the flattened SecretRef format described in the AXI credential standard");
  }
  private load(): Config {
    const legacy = join(homedir(), ".mg-axi", "config.json");
    if (resolve(this.path) === resolve(legacy) && existsSync(legacy)) this.moved();
    let text: string;
    try { text = this.files.read(this.path, false); }
    catch (error) {
      if (error instanceof AxiError) throw error;
      if ((error as NodeJS.ErrnoException).code === "ENOENT") {
        if (existsSync(legacy)) this.moved();
        return { version: 1, profiles: {} };
      }
      throw error;
    }
    let value: unknown;
    try { value = JSON.parse(text); }
    catch { invalid("Unable to read profile configuration; check the version-1 JSON schema"); }
    if (!object(value) || value.version !== 1 || !exact(value, ["version", "defaultProfile", "profiles"]) || !object(value.profiles)) invalid("Unsupported profile configuration; explicitly migrate to version 1");
    const profiles: Record<string, AnyProfile> = Object.create(null);
    for (const [name, profile] of Object.entries(value.profiles)) {
      if (!namePattern.test(name)) invalid("Invalid profile name");
      profiles[name] = validateProfile(profile, this.files);
      names.set(profiles[name], name);
    }
    if (value.defaultProfile !== undefined && (typeof value.defaultProfile !== "string" || !Object.hasOwn(profiles, value.defaultProfile))) invalid("Default profile does not exist");
    return { version: 1, profiles, defaultProfile: value.defaultProfile as string | undefined };
  }
  list() { return Object.entries(this.load().profiles).map(([name, profile]) => ({ name, mode: profile.mode, tenant: profile.tenantId, cloud: profile.cloud })); }
  resolve(name?: string) {
    const flag = name !== undefined && name.trim() !== "" ? name.trim() : undefined;
    const envName = process.env.MG_AXI_PROFILE !== undefined && process.env.MG_AXI_PROFILE.trim() !== "" ? process.env.MG_AXI_PROFILE.trim() : undefined;
    const explicit = flag ?? envName;
    if (explicit === undefined) {
      const profile = envProfile(process.env);
      if (profile) { names.set(profile, "env"); return { name: "env", profile }; }
    }
    const config = this.load();
    const configured = Object.keys(config.profiles);
    const selected = explicit ?? config.defaultProfile ?? (configured.length === 1 ? configured[0] : undefined);
    if (!selected) throw new AxiError("No configured profile selected", configured.length ? "PROFILE_AMBIGUOUS" : "PROFILE_REQUIRED", ["mg-axi profile list", "mg-axi profile create --help"]);
    if (!Object.hasOwn(config.profiles, selected)) throw new AxiError(`Profile ${selected} does not exist`, "PROFILE_NOT_FOUND", ["mg-axi profile list"]);
    return { name: selected, profile: config.profiles[selected]! };
  }
  verifyPermissions() {
    try { this.files.check(this.path, false, true); }
    catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error; }
  }
  create(name: string, tenantId: string, clientId: string, cloud: string, allowDeviceCode: boolean, application?: { certificateThumbprint?: string; federated?: boolean; clientSecretEnv?: string; clientSecretFile?: string; certificateKeyFile?: string; certificateKeyEnv?: string; federatedTokenFile?: string }) {
    if (!namePattern.test(name)) invalid("Profile name must contain only letters, digits, underscores or hyphens");
    const base = { tenantId, clientId, cloud, allowDeviceCode };
    let profile: AnyProfile;
    if (!application) profile = validateDelegatedProfile({ ...base, mode: "delegated", token: { stored: true } }, this.files);
    else {
      const { certificateThumbprint, federated, clientSecretEnv, clientSecretFile, certificateKeyFile, certificateKeyEnv, federatedTokenFile } = application;
      if (clientSecretEnv !== undefined && clientSecretFile !== undefined || certificateKeyEnv !== undefined && certificateKeyFile !== undefined) invalid("Use exactly one env or file reference for each credential");
      if ([certificateThumbprint !== undefined, federated === true, clientSecretEnv !== undefined || clientSecretFile !== undefined].filter(Boolean).length !== 1) invalid("Select exactly one application credential kind");
      if ((certificateKeyEnv !== undefined || certificateKeyFile !== undefined) && certificateThumbprint === undefined || federatedTokenFile !== undefined && !federated) invalid("Credential holder does not match the selected application auth kind");
      profile = validateApplicationProfile({ ...base, mode: "application",
        ...(certificateThumbprint !== undefined ? { auth: "certificate", thumbprintSha256: certificateThumbprint, privateKey: certificateKeyFile !== undefined ? { file: certificateKeyFile } : { env: certificateKeyEnv ?? "MG_AXI_PRIVATE_KEY" } }
          : federated ? { auth: "federated", ...(federatedTokenFile === undefined ? {} : { token: { file: federatedTokenFile } }) }
            : { auth: "client-secret", clientSecret: clientSecretFile !== undefined ? { file: clientSecretFile } : { env: clientSecretEnv } }) }, this.files);
    }
    const config = this.load();
    if (Object.hasOwn(config.profiles, name)) invalid("Profile already exists; creation never replaces identity");
    config.profiles[name] = profile;
    config.defaultProfile ??= name;
    this.files.write(this.path, `${JSON.stringify(config, null, 2)}\n`);
    return { name, mode: profile.mode, tenant: profile.tenantId, cloud: profile.cloud };
  }
}
