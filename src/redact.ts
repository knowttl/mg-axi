const secrets = new Set<string>();
export function registerSecret(value: string) {
  if (!value) return value;
  for (const form of [value, JSON.stringify(value).slice(1, -1), encodeURIComponent(value)]) secrets.add(form);
  return value;
}
export function redactSecrets(value: string): string {
  for (const secret of [...secrets].sort((a, b) => b.length - a.length)) value = value.split(secret).join("***redacted***");
  return value;
}

// MSAL caches contain opaque refresh, access and ID tokens under secret fields.
export function registerCache(value: string) {
  registerSecret(value);
  const visit = (entry: unknown, key = "") => {
    if (typeof entry === "string" && /^(secret|accessToken|refreshToken|idToken)$/i.test(key)) registerSecret(entry);
    else if (entry && typeof entry === "object") for (const [name, child] of Object.entries(entry)) visit(child, name);
  };
  visit(JSON.parse(value));
  return value;
}
