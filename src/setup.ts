import { LEAVES } from "./catalogue.js";
import { Profiles } from "./profiles.js";

// PACK-01: explicit release setup guidance. The setup view is offline: it
// reads the selected configuration file, describes installation and the
// supported Entra read surface, and points at the generated skill and
// capability report. It writes nothing, signs in nowhere and installs no
// hooks; ordinary commands never gain installation side effects. See
// README.md for the release notes and skills/mg-axi/SKILL.md for the
// installable skill.
export function setupView(store: Profiles): Record<string, unknown> {
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
      api: "Entra user, group, directory-role/PIM, device, administrative-unit, sign-in/directory-audit, application/service-principal, consent-grant, risk and Conditional Access reads plus reviewed raw api get, one gated group-membership write, one gated account enable/disable write, one gated session-revocation write and one gated risky-user dismissal write; every other operation is scheduled, blocked, deprecated or excluded",
      report: "docs/coverage.md",
    },
    integration: "Install skills/mg-axi/SKILL.md explicitly through your agent's skill installation mechanism; setup only shows guidance and installs no skills or hooks",
    help: ["node dist/bin/mg-axi.js doctor", "node dist/bin/mg-axi.js entra user list --help", "node dist/bin/mg-axi.js api get --help"],
  };
}
