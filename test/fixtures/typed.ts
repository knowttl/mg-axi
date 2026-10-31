import assert from "node:assert/strict";
import type { Credential, CredentialProvider } from "../../dist/auth.js";
import type { DispatchOverrides } from "../../dist/cli.js";
import type { AnyProfile } from "../../dist/profiles.js";

export type OutputRecord = Record<string, unknown>;
export type CredentialStub = {
  credential(profile: AnyProfile, scopes?: string[]): Promise<Pick<Credential, "token" | "tenantId" | "clientId"> & Partial<Credential>>;
};

// These fixtures exercise command/session behavior with opaque credentials,
// bypassing the auth service's cache and provider validation as before.
export function authFakes(fake: CredentialStub): Required<Pick<DispatchOverrides, "delegated" | "application">> {
  return {
    delegated: fake as NonNullable<DispatchOverrides["delegated"]>,
    application: fake as NonNullable<DispatchOverrides["application"]>,
  };
}

// Login/refresh fixtures never call logout; preserve their original surface.
export function delegatedProvider(fake: Omit<CredentialProvider, "logout">): CredentialProvider {
  return fake as CredentialProvider;
}

export function record(value: unknown): OutputRecord {
  assert.ok(value !== null && typeof value === "object" && !Array.isArray(value));
  return value as OutputRecord;
}

export function records(value: unknown): OutputRecord[] {
  assert.ok(Array.isArray(value));
  return value.map(record);
}

export function text(value: unknown): string {
  assert.equal(typeof value, "string");
  return value as string;
}

export function strings(value: unknown): string[] {
  assert.ok(Array.isArray(value));
  return value.map(text);
}
