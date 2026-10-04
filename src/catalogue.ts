import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { resolve, sep } from "node:path";
import { AxiError } from "axi-sdk-js";

// Use INV-01 identities and evidence directly; discovery never grants access.
const schema = JSON.parse(readFileSync(new URL("../inventory/schema.json", import.meta.url), "utf8"));
export const API_VERSIONS: string[] = schema.$defs.version.enum;
export const DESCRIPTION = "Inspect Microsoft Graph capabilities, read-only by default";

type Flag = { description: string; value?: string; default?: string; required?: true };
type Leaf = { path: string; description: string; flags: Record<string, Flag>; examples: string[]; operation?: string; positional?: { name: string; description: string } };
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
  { path: "api get", description: "Reviewed read-only raw Graph GET (API-01, v1.0 only): users, groups, conditional access, authentication methods, audit/sign-in, risk, apps, roles/PIM, devices and administrative units; unreviewed, secret-value, mail/file-content, beta and write routes are refused before credentials", positional: { name: "path", description: "Server-relative Graph path, e.g. /users" }, flags: {
    ...common,
    query: { value: "k=v&k2=v2", description: "OData query reviewed per route ($select/$filter/$top/$orderby on collections; $select/$expand on singles)" },
    scopes: { value: "comma-separated-Graph-scopes", description: "Delegated only: explicit full https://graph.microsoft.com/ scope names; application profiles use the .default audience" },
    limit: { value: "positive-integer", default: "100", description: "Cap returned rows; use --all to follow pages within budget" },
    all: { description: "Follow @odata.nextLink pages within the request budget" },
    full: { description: "Disable 4000-character string truncation; never disables redaction or row caps" },
  }, examples: ["mg-axi api get /users --scopes https://graph.microsoft.com/User.Read.All", "mg-axi api get /groups --query '$filter=securityEnabled eq true&$top=5' --scopes https://graph.microsoft.com/GroupMember.Read.All", "mg-axi api get /identity/conditionalAccess/policies --scopes https://graph.microsoft.com/Policy.Read.All"] },
];

export function leafHelp(leaf: Leaf): string {
  return [
    `mg-axi ${leaf.path}${leaf.positional ? ` <${leaf.positional.name}>` : ""}`, leaf.description,
    ...(leaf.positional ? [`<${leaf.positional.name}>: ${leaf.positional.description}`] : []),
    ...Object.entries(leaf.flags).map(([name, flag]) => `--${name}${flag.value ? ` <${flag.value}>` : ""}: ${flag.description}${flag.default ? ` (default: ${flag.default})` : ""}${flag.required ? " (required)" : ""}`),
    "--help: Show this reference",
    ...leaf.examples,
  ].join("\n");
}

export const TOP_LEVEL_HELP = [DESCRIPTION, ...LEAVES.map(leaf => leafHelp(leaf)), "-v, -V, --version: Print the bare version"].join("\n\n");

export function resolveCommand(argv: string[]): { leaf: Leaf; flags: Record<string, string | boolean>; positional?: string } {
  const words = argv.slice(0, argv.findIndex(arg => arg.startsWith("-")) < 0 ? argv.length : argv.findIndex(arg => arg.startsWith("-")));
  const path = words.join(" ") || "home";
  let leaf = LEAVES.find(item => item.path === path);
  let positional: string | undefined;
  let flagStart = words.length;
  if (!leaf) {
    if (words[0] !== "api") throw new AxiError("unknown or incomplete command", "VALIDATION_ERROR", [TOP_LEVEL_HELP]);
    leaf = LEAVES.find(item => item.path === "api get")!;
    const help = leafHelp(leaf);
    const verb = words[1];
    if (verb === undefined) throw new AxiError("unknown or incomplete command", "VALIDATION_ERROR", [TOP_LEVEL_HELP]);
    if (verb.toLowerCase() !== "get") throw new AxiError(`mg-axi api serves reviewed GET reads only; got api ${verb}`, "VALIDATION_ERROR", [help]);
    if (words.length < 3 || words[2]!.startsWith("-")) throw new AxiError("missing path for `mg-axi api get`", "VALIDATION_ERROR", ["Example: mg-axi api get /users --scopes https://graph.microsoft.com/User.Read.All", help]);
    if (words.length > 3) throw new AxiError(`unexpected argument \`${words[3]}\` for \`mg-axi api get\``, "VALIDATION_ERROR", ["Pass one path before flags: mg-axi api get <path> [--query 'k=v']", help]);
    positional = words[2];
    flagStart = 3;
  }
  const help = leafHelp(leaf);
  const fail = (message: string): never => { throw new AxiError(message, "VALIDATION_ERROR", [help]); };
  const flags: Record<string, string | boolean> = Object.create(null);
  for (let i = flagStart; i < argv.length; i++) {
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
  if (!flags.help && leaf.positional && positional === undefined) fail(`missing ${leaf.positional.name} for \`mg-axi ${leaf.path}\``);
  return { leaf, flags, positional };
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
    help: ["mg-axi entra user list --help", "mg-axi entra user show --help", "mg-axi api get --help"],
  };
}
