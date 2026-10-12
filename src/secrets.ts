import { AxiError } from "axi-sdk-js";
import { PrivateFiles } from "./private-file.js";
import { registerSecret } from "./redact.js";

export type SecretRef = Readonly<{ env: string } | { file: string } | { stored: true }>;
export function validateSecretRef(value: unknown, field: string, stored = true, files = new PrivateFiles()): SecretRef {
  if (value === undefined) return Object.freeze({ env: `MG_AXI_${field.replace(/[A-Z]/g, letter => `_${letter}`).toUpperCase()}` });
  if (value && typeof value === "object" && !Array.isArray(value) && Object.keys(value).length === 1) {
    const ref = value as Record<string, unknown>;
    if (typeof ref.env === "string" && /^[A-Za-z_][A-Za-z0-9_]*$/.test(ref.env)) return Object.freeze({ env: ref.env });
    if (typeof ref.file === "string" && files.absolute(ref.file)) return Object.freeze({ file: ref.file });
    if (stored && ref.stored === true) return Object.freeze({ stored: true });
  }
  throw new AxiError(`Invalid ${field}; use a SecretRef, never an inline secret`, "CONFIG_INVALID", ["Use exactly one validated env name, absolute local file path, or stored: true for a delegated login"]);
}

export function secretSource(ref: SecretRef) { return "env" in ref ? `env:${ref.env}` : "file" in ref ? `file:${ref.file}` : "stored"; }
export function resolveSecret(ref: SecretRef, files = new PrivateFiles(), verifyAcl = false): string {
  let value: string | undefined;
  if ("env" in ref) value = process.env[ref.env];
  else if ("file" in ref) {
    try { value = files.read(ref.file, true, verifyAcl).trimEnd(); }
    catch (error) {
      if (error instanceof AxiError) throw error;
      if ((error as NodeJS.ErrnoException)?.code !== "ENOENT") throw error;
    }
  }
  if (!value?.trim()) throw new AxiError(`Credential unavailable from ${secretSource(ref)}`, "AUTH_REQUIRED", ["Supply the named credential reference; no interactive fallback was attempted"]);
  return registerSecret(value);
}
