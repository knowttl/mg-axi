import { LEAVES } from "./catalogue.js";
import { Profiles } from "./profiles.js";

// PACK-01: explicit release setup guidance. The setup view is offline: it
// reads the selected configuration file, describes installation and the
// supported Entra read surface, and points at the generated skill and
// capability report. It writes nothing, signs in nowhere and installs no
// hooks; ordinary commands never gain installation side effects. See
// README.md for the release notes and skills/mg-axi/SKILL.md for the
// installable skill.
// EXT-01b document gate (D2): the capability prose lists the shipped read
// families, driven from the catalogue in catalogue order, so a merged read
// slice extends the sentence without touching this template. Paths outside
// the known prefixes fall back to their two-segment leaf key and are never
// silently omitted from the list.
const READ_FAMILY_PREFIXES: ReadonlyArray<readonly [string, string]> = [
  ["entra user", "user"],
  ["entra registration", "registration"],
  ["entra group", "group"],
  ["entra directory-role", "directory-role/PIM"],
  ["entra role-assignment", "directory-role/PIM"],
  ["entra pim", "directory-role/PIM"],
  ["entra device", "device"],
  ["entra administrative-unit", "administrative-unit"],
  ["entra organization", "organization and branding"],
  ["entra domain", "domain and domain DNS"],
  ["entra domain-dns-record", "domain and domain DNS"],
  ["entra certificate-auth-pki", "certificate authentication"],
  ["entra certificate-authority", "certificate authentication"],
  ["entra subscription", "commercial subscriptions"],
  ["entra on-premises-synchronization", "on-premises synchronization"],
  ["entra agreement", "terms-of-use agreements and acceptances"],
  ["entra agreement-acceptance", "terms-of-use agreements and acceptances"],
  ["entra group-lifecycle-policy", "group lifecycle policies"],
  ["entra group-setting-template", "group setting templates"],
  ["entra attribute-set", "attribute sets and custom security attributes"],
  ["entra custom-security-attribute-definition", "attribute sets and custom security attributes"],
  ["entra allowed-value", "attribute sets and custom security attributes"],
  ["entra contract", "partner contracts"],
  ["entra sign-in", "sign-in/directory-audit"],
  ["entra directory-audit", "sign-in/directory-audit"],
  ["entra application", "application/service-principal and consent grants"],
  ["entra service-principal", "application/service-principal and consent grants"],
  ["entra risky-user", "risk"],
  ["entra risk-detection", "risk"],
  ["entra conditional-access", "Conditional Access"],
  ["entra identity-provider", "identity-provider"],
  ["entra federation-configuration", "federation configurations"],
  ["entra data-policy-operation", "data policy operations"],
  ["entra access-review", "access-review"],
];

function shippedReadFamilies(): string[] {
  const seen: string[] = [];
  for (const leaf of LEAVES) {
    if (leaf.operation === undefined || !leaf.operation.startsWith("GET:")) continue;
    const match = READ_FAMILY_PREFIXES.find(([prefix]) => leaf.path === prefix || leaf.path.startsWith(`${prefix} `));
    const label = match ? match[1] : leaf.path.split(" ").slice(0, 2).join(" ");
    if (!seen.includes(label)) seen.push(label);
  }
  return seen;
}

export function setupView(store: Profiles): Record<string, unknown> {
  const families = shippedReadFamilies();
  const items = store.list();
  const implemented = LEAVES.map(leaf => leaf.path);
  const reads = LEAVES.filter(leaf => (leaf.operation !== undefined && leaf.operation.startsWith("GET:")) || leaf.path === "api get" || leaf.path === "doctor").length;
  const writes = LEAVES.filter(leaf => leaf.operation !== undefined && !leaf.operation.startsWith("GET:")).length;
  return {
    command: "mg-axi setup",
    config: store.path,
    state: items.length ? "configured" : "unconfigured",
    profiles: items.length
      ? items.map(item => ({ name: item.name, mode: item.mode }))
      : "0 profiles configured",
    installation: [
      "corepack pnpm install --frozen-lockfile --ignore-scripts --config.confirm-modules-purge=false",
      "corepack pnpm build",
      "node dist/bin/mg-axi.js --version",
    ],
    guidance: [
      `Hand-edit ${store.path} or set MG_AXI_CONFIG to a separate configuration file; secrets stay in protected storage references, never in argv or config`,
      "node dist/bin/mg-axi.js profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial",
      "node dist/bin/mg-axi.js login --profile soc --scopes https://graph.microsoft.com/User.Read.All",
      "node dist/bin/mg-axi.js doctor",
    ],
    capabilities: {
      implemented,
      reads,
      writes,
      local: implemented.length - reads - writes,
      api: `Entra ${families.slice(0, -1).join(", ")} and ${families.at(-1)} reads plus reviewed raw api get, one gated group-membership write, one gated account enable/disable write, one gated session-revocation write and one gated risky-user dismissal write; every other operation is scheduled, blocked, deprecated or excluded`,
      report: "docs/coverage.md",
    },
    integration: "Install skills/mg-axi/SKILL.md explicitly through your agent's skill installation mechanism; setup only shows guidance and installs no skills or hooks",
    help: ["node dist/bin/mg-axi.js doctor", "node dist/bin/mg-axi.js entra user list --help", "node dist/bin/mg-axi.js api get --help"],
  };
}
