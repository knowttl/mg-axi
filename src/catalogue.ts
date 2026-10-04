import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { resolve, sep } from "node:path";
import { AxiError } from "axi-sdk-js";

// Use INV-01 identities and evidence directly; discovery never grants access.
const schema = JSON.parse(readFileSync(new URL("../inventory/schema.json", import.meta.url), "utf8"));
export const API_VERSIONS: string[] = schema.$defs.version.enum;
export const DESCRIPTION = "Inspect Microsoft Graph capabilities, read-only by default";

type Flag = { description: string; value?: string; default?: string; required?: true };
type Leaf = { path: string; description: string; flags: Record<string, Flag>; examples: string[]; operation?: string };
const common = {
  profile: { value: "name", description: "Select a configured profile" },
  "api-version": { value: API_VERSIONS.join("|"), default: "v1.0", description: "Explicit API version; no fallback" },
};
const PARSE_MODES = ["delegated", "application"];
export const LEAVES: Leaf[] = [
  { path: "home", description: "Show local profile status without authenticating", flags: { profile: common.profile }, examples: ["mg-axi", "mg-axi home --profile soc"] },
  { path: "profile create", description: "Create a dedicated-app profile without signing in", flags: {
    name: { value: "name", required: true, description: "New profile name; existing identities cannot be overwritten" },
    tenant: { value: "tenant-id", required: true, description: "Explicit workforce tenant UUID" },
    client: { value: "client-id", required: true, description: "Organization-owned application UUID" },
    cloud: { value: "commercial", required: true, description: "Explicit cloud; only commercial is supported" },
    mode: { value: "delegated|application", default: "delegated", description: "Delegated analyst sign-in or application client credentials" },
    "allow-device-code": { description: "Delegated only: opt in only when permitted by organization policy; never enables fallback" },
    "certificate-thumbprint": { value: "40-hex-digits", description: "Application only: certificate thumbprint; the private key stays in protected storage" },
    federated: { description: "Application only: workload federation via the AZURE_FEDERATED_TOKEN_FILE assertion source" },
  }, examples: ["mg-axi profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial", "mg-axi profile create --name batch --tenant <tenant-id> --client <client-id> --cloud commercial --mode application --federated", "mg-axi profile create --help"] },
  { path: "profile list", description: "List configured profiles without accessing credentials", flags: {}, examples: ["mg-axi profile list", "mg-axi profile list --help"] },
  { path: "profile show", description: "Show profile policy and credential reference, never credentials", flags: {
    profile: common.profile,
  }, examples: ["mg-axi profile show --profile soc", "mg-axi profile show"] },
  { path: "login", description: "Explicit delegated login; browser by default, no automatic device-code fallback", flags: {
    profile: common.profile,
    method: { value: "browser|device-code", default: "browser", description: "Device code additionally requires profile opt-in" },
    scopes: { value: "comma-separated-Graph-scopes", required: true, description: "Explicit delegated permissions using full https://graph.microsoft.com/ scope names" },
  }, examples: ["mg-axi login --profile soc --scopes https://graph.microsoft.com/User.Read", "mg-axi login --profile soc --method device-code --scopes https://graph.microsoft.com/User.Read"] },
  { path: "entra user list", description: "List users (scheduled for READ-01; not executable yet)", operation: "GET:/users", flags: {
    ...common,
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; incompatible with --all" },
    all: { description: "Follow pages within a budget (ships with CORE-02)" },
  }, examples: ["mg-axi entra user list --help", "mg-axi entra user list --limit 100"] },
  { path: "entra user show", description: "Show a user (scheduled for READ-01; not executable yet)", operation: "GET:/users/{user-id}", flags: {
    ...common, id: { value: "user-id-or-upn", required: true, description: "User object ID or UPN" },
  }, examples: ["mg-axi entra user show --help", "mg-axi entra user show --id <user-id-or-upn>"] },
];

export function leafHelp(leaf: Leaf): string {
  return [
    `mg-axi ${leaf.path}`, leaf.description,
    ...Object.entries(leaf.flags).map(([name, flag]) => `--${name}${flag.value ? ` <${flag.value}>` : ""}: ${flag.description}${flag.default ? ` (default: ${flag.default})` : ""}${flag.required ? " (required)" : ""}`),
    "--help: Show this reference",
    ...leaf.examples,
  ].join("\n");
}

export const TOP_LEVEL_HELP = [DESCRIPTION, ...LEAVES.map(leaf => leafHelp(leaf)), "-v, -V, --version: Print the bare version"].join("\n\n");

export function resolveCommand(argv: string[]): { leaf: Leaf; flags: Record<string, string | boolean> } {
  const words = argv.slice(0, argv.findIndex(arg => arg.startsWith("-")) < 0 ? argv.length : argv.findIndex(arg => arg.startsWith("-")));
  const path = words.join(" ") || "home";
  const leaf = LEAVES.find(item => item.path === path);
  if (!leaf) throw new AxiError("unknown or incomplete command", "VALIDATION_ERROR", [TOP_LEVEL_HELP]);
  const help = leafHelp(leaf);
  const fail = (message: string): never => { throw new AxiError(message, "VALIDATION_ERROR", [help]); };
  const flags: Record<string, string | boolean> = Object.create(null);
  for (let i = words.length; i < argv.length; i++) {
    const arg = argv[i]!;
    const match = /^--([a-z-]+)(?:=(.*))?$/.exec(arg);
    if (!match) fail("unexpected argument or short flag");
    const name = match![1]!;
    const flag = Object.hasOwn(leaf.flags, name) ? leaf.flags[name] : name === "help" ? { description: "Help" } : undefined;
    if (!flag) fail(`unknown flag --${name}`);
    if (Object.hasOwn(flags, name)) fail(`duplicate flag --${name}`);
    if (flag!.value) {
      const value = match![2] ?? argv[++i];
      if (!value?.trim() || value.startsWith("-")) fail(`--${name} requires a non-empty value`);
      flags[name] = value!;
    } else {
      if (match![2] !== undefined) fail(`--${name} does not take a value`);
      flags[name] = true;
    }
  }
  if (flags["api-version"] && !API_VERSIONS.includes(String(flags["api-version"]))) fail(`--api-version must be ${API_VERSIONS.join(" or ")}`);
  if (flags.mode && !PARSE_MODES.includes(String(flags.mode))) fail("--mode must be delegated or application");
  if (flags.limit && (!/^[1-9]\d*$/.test(String(flags.limit)) || !Number.isSafeInteger(Number(flags.limit)))) fail("--limit must be a positive safe integer");
  if (flags.limit && flags.all) fail("--limit and --all cannot be combined");
  if (!flags.help) for (const [name, flag] of Object.entries(leaf.flags)) if (flag.required && !flags[name]) fail(`--${name} is required`);
  return { leaf, flags };
}

export function operationFor(leaf: Leaf, version: string) {
  const inventory = JSON.parse(readFileSync(new URL("../inventory/operations.json", import.meta.url), "utf8"));
  return inventory.operations.find((row: { id: string }) => row.id === `${version}:${leaf.operation}`);
}

// A useful local status, without treating unavailable tenant summaries as zeros.
export function home() {
  const bin = resolve(process.argv[1]!);
  return {
    bin: bin.startsWith(`${homedir()}${sep}`) ? `~${bin.slice(homedir().length)}` : bin,
    description: DESCRIPTION,
    profile: "unavailable: no profile configured",
    tenant: "unavailable: no tenant selected",
    domains: [{ name: "entra", status: "scheduled", summary: "Tenant summaries await Graph execution" }],
    help: ["mg-axi entra user list --help", "mg-axi entra user show --help"],
  };
}
