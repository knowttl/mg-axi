import { readFileSync } from "node:fs";
import { DESCRIPTION, HOME_HELP, LEAVES } from "./catalogue.js";

// PACK-01: single source for generated release records. The committed
// skills/mg-axi/SKILL.md and docs/coverage.md are generated from the templates
// below, the executable catalogue and the discovery inventory, so help, skill
// and coverage cannot drift apart. tools/generate-docs.mjs writes both
// files; test/pack.test.mjs fails when they are stale.
//
// The skill file is generated in full by skillDocument() below: its trigger
// description, its next-step hints (the same hints the no-args home view
// prints) and its command table all come from the catalogue, so a new leaf
// without an area label, or a hand-edit to the committed file, fails the
// --check step instead of drifting silently.

// One row per executable leaf in catalogue order.
// Every leaf is implemented by an mg-axi handler (native). Graph
// leaves carry the read effect, except reviewed named writes.
function leafEffect(leaf: { path: string; operation?: string }): string {
  if (leaf.operation !== undefined) return leaf.operation.startsWith("GET:") ? "read" : "write";
  return leaf.path === "api get" || leaf.path === "doctor" ? "read" : "local";
}
export function skillCommandTable(): string {
  return [
    "| Command | Capability | Effect |",
    "|---|---|---|",
    ...LEAVES.map(leaf => `| \`mg-axi ${leaf.path}\` | native | ${leafEffect(leaf)} |`),
  ].join("\n");
}

// SKILL-01: the installable skill is generated in full from the catalogue.
// Area labels for the trigger description, keyed by leaf-path prefix with
// specific prefixes before general ones. The first matching prefix owns each
// read leaf: a new read leaf without coverage here throws, so the
// committed description can never silently omit shipped coverage.
const READ_AREA_LABELS: Array<[string, string]> = [
  ["entra user authentication-method ", "authentication methods"],
  ["entra registration ", "authentication methods"],
  ["entra user ", "users"],
  ["entra group ", "groups"],
  ["entra directory-role ", "roles"],
  ["entra role-assignment ", "roles"],
  ["entra pim ", "roles"],
  ["entra device ", "devices"],
  ["entra administrative-unit ", "administrative units"],
  ["entra organization ", "organization and branding"],
  ["entra certificate-auth-pki ", "certificate authentication"],
  ["entra certificate-authority ", "certificate authentication"],
  ["entra subscription ", "commercial subscriptions"],
  ["entra on-premises-synchronization ", "on-premises synchronization"],
  ["entra agreement ", "terms-of-use agreements and acceptances"],
  ["entra agreement-acceptance ", "terms-of-use agreements and acceptances"],
  ["entra directory-object ", "directory objects"],
  ["entra domain ", "domains and DNS records"],
  ["entra domain-dns-record ", "domains and DNS records"],
  ["entra group-lifecycle-policy ", "group lifecycle policies"],
  ["entra group-setting-template ", "group setting templates"],
  ["entra attribute-set ", "attribute sets and custom security attributes"],
  ["entra custom-security-attribute-definition ", "attribute sets and custom security attributes"],
  ["entra allowed-value ", "attribute sets and custom security attributes"],
  ["entra contract ", "partner contracts"],
  ["entra sign-in ", "sign-ins and audit logs"],
  ["entra directory-audit ", "sign-ins and audit logs"],
  ["entra application ", "applications and consent grants"],
  ["entra service-principal ", "applications and consent grants"],
  ["entra risky-user ", "risk"],
  ["entra risk-detection ", "risk"],
  ["entra conditional-access ", "Conditional Access"],
  ["entra identity-provider ", "identity providers"],
  ["entra data-policy-operation ", "data policy operations"],
  ["entra access-review ", "access reviews"],
  ["api get", "reviewed raw reads"],
];

// Short outcome labels for the shipped gated writes, keyed by leaf path.
// A new write leaf without an entry here throws for the same reason.
const WRITE_LABELS: Record<string, string> = {
  "entra user update": "account enable/disable",
  "entra user revoke-sessions": "session revocation",
  "entra group member add": "group-membership add",
  "entra risky-user dismiss": "risky-user dismissal",
  "entra conditional-access policy update": "CA policy update",
};

function skillReadAreas(): string[] {
  const areas: string[] = [];
  for (const leaf of LEAVES) {
    if (leafEffect(leaf) !== "read" || leaf.path === "doctor") continue;
    const match = READ_AREA_LABELS.find(([prefix]) => leaf.path.startsWith(prefix) || leaf.path === prefix.trim());
    if (!match) throw new Error(`SKILL.md has no area label for read leaf: ${leaf.path}`);
    if (!areas.includes(match[1])) areas.push(match[1]);
  }
  return areas;
}

function skillWriteLabels(): string[] {
  return LEAVES.filter(leaf => leafEffect(leaf) === "write").map(leaf => {
    const label = WRITE_LABELS[leaf.path];
    if (!label) throw new Error(`SKILL.md has no write label for write leaf: ${leaf.path}`);
    return label;
  });
}

// Terse, outcome-focused trigger description: current read coverage plus
// the shipped gated writes. Generated so EXT reads and new writes land in
// the frontmatter with the slice that ships them.
export function skillDescription(): string {
  const areas = skillReadAreas();
  const writes = skillWriteLabels();
  return `Inspect Microsoft Entra ${areas.join(", ")} and run ${writes.length} gated writes (${writes.join(", ")}) through token-efficient TOON output.`;
}

// Next-step hints printed by the no-args home view. The two leading entries mirror the
// profile/login hints localHome() in cli.ts prepends to HOME_HELP; the
// rest is HOME_HELP verbatim, so skill and home view cannot drift apart.
export function skillHomeHints(): string[] {
  return ["mg-axi profile list", "mg-axi login --help", ...HOME_HELP];
}

// The full installable skill. Durable guidance only: read-only by default,
// writes need preview plus typed confirmation, no secrets in output, the
// profile/auth model, offline tests. Long walkthroughs live in leaf --help
// and docs/ instead.
export function skillDocument(): string {
  return [
    "---",
    "name: mg-axi",
    "description: >",
    `  ${skillDescription()}`,
    "user-invocable: false",
    "---",
    "",
    "<!-- Generated by tools/generate-docs.mjs from src/catalogue.ts. Do not hand-edit. -->",
    "",
    "# mg-axi",
    "",
    `Agent-ergonomic CLI for Microsoft Graph. ${DESCRIPTION}.`,
    "Entra SOC reads through token-efficient TOON output.",
    "The generated [command table](#orientation) lists supported named writes through the gated mutation coordinator; raw API remains read-only and every other mutation is refused.",
    "See [README.md](../../README.md) for write usage, execution gates, identity pinning, permissions, target-role hierarchy and write configuration.",
    "",
    "Build and link the local checkout as documented in [README.md](../../README.md#mg-axi), then run commands non-interactively with `mg-axi <command>`.",
    "Once the package is published, `npx -y @knowttl/mg-axi <command>` also applies.",
    "Run `mg-axi doctor` first.",
    "See [README.md](../../README.md#release) for doctor profile selection, checks and failure behavior.",
    "",
    "## Orientation",
    "",
    "The exact current leaf registry is `src/catalogue.ts`.",
    "Its capability label is `native` (implemented by an mg-axi handler).",
    "Its effect is `read` for Graph read leaves and the doctor health check, `write` for the gated named writes, or `local` for home, profile, login and setup views.",
    "Doctor acquires credentials silently and contacts Graph; only its help view stays offline.",
    "The list below records current executable leaves; it makes no coverage claim for other Graph operations.",
    "See `docs/coverage.md` for the per-operation disposition records.",
    "`docs/coverage.md` and this file are generated from the same catalogue and inventory sources; `node tools/generate-docs.mjs --check` fails when they are stale.",
    "",
    "<!-- command-registry:start -->",
    skillCommandTable(),
    "<!-- command-registry:end -->",
    "",
    "Run `mg-axi <leaf-path> --help` for that leaf's accepted flags and reference.",
    "Unknown flags fail before any credential or HTTP work.",
    "",
    "## Next steps",
    "",
    "The no-args home view prints these hints alongside live profile status:",
    "",
    "```sh",
    ...skillHomeHints(),
    "```",
    "",
    "## Setup (explicit only)",
    "",
    "Follow [README.md](../../README.md#mg-axi) for checkout builds and binary linking, and [release guidance](../../README.md#release) for explicit skill installation.",
    "Create profiles explicitly and keep secrets out of argv and config files:",
    "",
    "```sh",
    "mg-axi setup                    # build steps, config path and capabilities; writes nothing",
    "mg-axi profile create --name soc --tenant <tenant-id> --client <client-id> --cloud commercial",
    "mg-axi login --profile soc --scopes https://graph.microsoft.com/User.Read.All",
    "mg-axi doctor                   # one bounded read per selected profile",
    "mg-axi entra user list --profile soc --limit 10",
    "```",
    "",
    "Configuration defaults to `~/.mg-axi/config.json`; `MG_AXI_CONFIG` selects a separate file.",
    "`mg-axi setup` shows the selected path and writes nothing.",
    "",
    "## Selecting a profile",
    "",
    "Read leaves accept `--profile <name>`.",
    "Selection uses `--profile` or the configured default.",
    "Without either, read leaves report `AUTH_REQUIRED`, even if only one profile exists.",
    "Doctor selection is documented in [README.md](../../README.md#release).",
    "Without profiles, home, setup and profile list show unconfigured state; Graph read leaves report `AUTH_REQUIRED`.",
    "",
    "## Safety",
    "",
    "Reads run through a session that authorizes only catalogued operations before silent credential acquisition; only explicit `login` opens a browser or prints a device-code challenge.",
    "Writes preview before sending and need `--execute` plus a typed `--confirm` repeating the target; an already-desired value is a no-op with exit 0, never an error.",
    "Output is TOON on stdout; diagnostics use stderr.",
    "Text truncates at 500 characters (`--full` restores text, never redaction); row caps resume through opaque cursors.",
    "Secrets stay in protected storage references, never in argv, config files or output; credential fields carry expiry metadata only.",
    "Tests run offline with fixture credentials; never point tests at a real tenant.",
    "",
  ].join("\n");
}

type InventoryRow = {
  id: string;
  disposition: string;
  owningSlice: string | null;
  reason: string;
};

// Initial write families resolve to shipped when the catalogue carries a
// leaf for one of their v1.0 operations, and stay pending while every row
// remains unimplemented. Family membership comes from the inventory
// owningSlice and shipped status comes from the catalogue, so a merged
// write slice flips its own row without touching this template. Blocked,
// deprecated and unavailable variants never count as pending.
function initialWriteStatus(rows: InventoryRow[]): { shipped: string[]; pending: string[] } {
  const families = new Map<string, string[]>();
  for (const row of rows) {
    const operation = /^v1\.0:(.+)$/.exec(row.id)?.[1];
    if (operation === undefined || row.owningSlice === null || !/^WRITE-0[1-9]$/.test(row.owningSlice)) continue;
    if (row.disposition === "intentionally-blocked" || row.disposition === "deprecated" || row.disposition === "unavailable") continue;
    families.set(row.owningSlice, [...(families.get(row.owningSlice) ?? []), operation]);
  }
  const shipped: string[] = [];
  const pending: string[] = [];
  for (const [slice, operations] of [...families.entries()].sort()) {
    const leaf = LEAVES.find(candidate => candidate.operation !== undefined && operations.includes(candidate.operation));
    if (leaf) shipped.push(`${slice} (\`mg-axi ${leaf.path}\`)`);
    else pending.push(`${slice} (${operations.map(operation => `\`${operation}\``).join(", ")})`);
  }
  return { shipped, pending };
}

const DISPOSITIONS = ["named-command", "reviewed-raw-read", "scheduled", "intentionally-blocked", "deprecated", "unavailable", "excluded"] as const;

// The capability report is generated from the catalogue: every executable
// read leaf owns one row with its inventory disposition and owning slice,
// and the counts separate implemented leaves from inventory dispositions.
// Scheduled work never counts as complete and excluded rows stay separate
// from scoped capabilities, per docs/inventory.md. There is no completeness
// percentage: the report names the supported Entra read surface, not full
// Entra coverage.
export function capabilityDocument(): string {
  const inventory = JSON.parse(readFileSync(new URL("../inventory/operations.json", import.meta.url), "utf8")) as {
    operations: InventoryRow[];
  };
  const rows = inventory.operations;
  const count = (disposition: string): number => rows.filter(row => row.disposition === disposition).length;
  const writeStatus = initialWriteStatus(rows);
  const readLeaves = LEAVES.filter(leaf => leaf.operation !== undefined && leaf.operation.startsWith("GET:"));
  const writeLeaves = LEAVES.filter(leaf => leaf.operation !== undefined && !leaf.operation.startsWith("GET:"));
  const commandRows = readLeaves.map(leaf => {
    const row = rows.find(candidate => candidate.id === `v1.0:${leaf.operation}`);
    return `| \`mg-axi ${leaf.path}\` | \`${leaf.operation}\` | ${row ? row.disposition : "no v1.0 inventory row"} | ${row?.owningSlice ?? "-"} |`;
  });
  const writeRows = writeLeaves.map(leaf => {
    const row = rows.find(candidate => candidate.id === `v1.0:${leaf.operation}`);
    return `| \`mg-axi ${leaf.path}\` | \`${leaf.operation}\` | ${row?.owningSlice ?? "-"} |`;
  });
  const deferredDomainRows = rows.filter(row => row.owningSlice === "EXT-01" && row.reason.startsWith("Deferred by firstmate R1 "));
  const deferredProviderRows = rows.filter(row => row.owningSlice === "EXT-03" && row.reason.startsWith("Deferred to a later EXT-03 "));
  const deferredDataPolicyRows = rows.filter(row => row.owningSlice === "EXT-03" && row.reason.startsWith("Deferred by firstmate data-policy-operations scope "));
  const unavailableInvitationRows = rows.filter(row => row.owningSlice === "EXT-03" && row.reason.startsWith("Marked unavailable by firstmate EXT-03b decision: "));
  const deferredOrganizationRows = rows.filter(row => row.owningSlice === "EXT-01" && row.reason.startsWith("Deferred by firstmate organization scope "));
  const unavailableCertAuthRows = rows.filter(row => row.owningSlice === "EXT-01" && row.reason.startsWith("Marked unavailable by firstmate mg-ext-01d decision: "));
  const deferredSubscriptionRows = rows.filter(row => row.owningSlice === "EXT-01" && row.reason.startsWith("Deferred by firstmate directory-subscriptions scope "));
  const deferredOnPremSyncRows = rows.filter(row => row.owningSlice === "EXT-01" && row.reason.startsWith("Deferred by firstmate on-premises-synchronization scope "));
  const deferredTermsOfUseRows = rows.filter(row => row.owningSlice === "EXT-01" && row.reason.startsWith("Deferred by firstmate terms-of-use scope "));
  const deferredDirectoryObjectRows = rows.filter(row => row.owningSlice === "EXT-01" && row.reason.startsWith("Deferred by firstmate directory-objects scope "));
  const deferredContractRows = rows.filter(row => row.owningSlice === "EXT-04" && row.reason.startsWith("Deferred by firstmate contracts scope "));
  const deferredGroupLifecycleRows = rows.filter(row => row.owningSlice === "EXT-01" && row.reason.startsWith("Deferred by firstmate group-lifecycle scope "));
  const deferredCustomSecurityRows = rows.filter(row => row.owningSlice === "EXT-01" && row.reason.startsWith("Deferred by firstmate custom-security-attributes scope "));
  const deferredAccessReviewRows = rows.filter(row => row.owningSlice === "EXT-02c");
  const readCount = readLeaves.length + 2;
  const localCount = LEAVES.length - readCount - writeLeaves.length;
  return [
    "<!-- Generated by tools/generate-docs.mjs from src/catalogue.ts and inventory/operations.json. Do not hand-edit. -->",
    "",
    "# mg-axi coverage",
    "",
    "This is the supported Entra read surface, not full Entra coverage.",
    "Only the commands below have an implemented, tested leaf; every other",
    "operation remains scheduled, blocked, deprecated or excluded until its own slice ships.",
    "Inventory dispositions are discovery-time records from INV-01; the command",
    "list below is the implemented truth and never counts scheduled work as complete.",
    "",
    "## Counts",
    "",
    `- implemented read leaves: ${readCount} (${readLeaves.length} named Entra reads plus reviewed raw api get and doctor health check)`,
    `- implemented write leaves: ${writeLeaves.length} (named gated mutations below)`,
    `- local leaves: ${localCount} (home, profile, login and setup views)`,
    ...DISPOSITIONS.map(disposition => `- inventory ${disposition}: ${count(disposition)}`),
    "",
    "## Commands",
    "",
    "| Command | Operation | Inventory (v1.0) | Owning slice |",
    "|---|---|---|---|",
    ...commandRows,
    "| `mg-axi api get` | reviewed raw reads (see src/api.ts) | reviewed-raw-read catalogue | API-01 |",
    "| `mg-axi doctor` | bounded `GET:/users` health check | uses the named user-list read | PACK-01 |",
    "",
    "## EXT-01 domain scope decisions",
    "",
    "Firstmate decision R1: approve narrowing this change to the eight v1.0 domain reads above.",
    "The operations below remain scheduled with an explicit deferred disposition to a later EXT-01 subfamily; no new commands or raw access are approved.",
    "Federation stays out because it can carry signing-certificate material.",
    "Firstmate decision R2: keep the write-family status section as the coverage fix added to scope (resolved-kept).",
    "Both firstmate decisions must also be stated in the PR body by the delivery phase.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredDomainRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-03 identity-provider scope decisions",
    "",
    "This change covers the four v1.0 workforce identity-provider reads above (list, show, count, available-types).",
    "The four beta operations below remain scheduled with an explicit deferred disposition to a later EXT-03 subfamily.",
    "Identity-provider reads support v1.0 only; beta needs its own review.",
    "Workforce context only; no external-customer (B2C/External ID tenant) support is claimed.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredProviderRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-03 data-policy-operations scope decisions",
    "",
    "This change covers the three v1.0 workforce data-policy-operation reads above (data-policy-operation list, show and count).",
    "The three beta operations below remain scheduled with an explicit deferred disposition to a later EXT-03 subfamily; no new commands or raw access are approved.",
    "Data-policy-operation reads support v1.0 only; beta needs its own review.",
    "storageLocation always renders as the redaction marker; export submission (POST /users/{id}/exportPersonalData) belongs to no read slice.",
    "Workforce context only; no external-customer support is claimed.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredDataPolicyRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-03 invitations scope decisions",
    "",
    "Firstmate decision (mg-ext-03b inbox 001): do not implement these reads and add no scope.",
    "The eight v1.0 invitation reads below carry an explicit reviewed unavailable disposition; no new commands or raw access are approved.",
    "The [v1.0 invitation resource documentation](https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/v1.0/resources/invitation.md) lists Create only in its Methods table and provides no documented GET contract for these invitation reads.",
    "Workforce B2B context only; no invitation creation or any mutation is claimed here.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...unavailableInvitationRows.map(row => `| \`${row.id}\` | ${row.disposition} | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-01 organization scope decisions",
    "",
    "The organization scope covers the five organization/branding/localization reads above; the remaining subfamilies need the separate contracts described below.",
    "The operations below remain scheduled with an explicit deferred disposition to a later EXT-01 subfamily; no new commands or raw access are approved.",
    "Stream image and CSS bytes stay out because they need a separate binary-output contract; beta-only settings, partner and theme contracts need separate review.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredOrganizationRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-01 certificate-auth scope decisions",
    "",
    "This change covers the six v1.0 PKI and certificate-authority reads above (certificate-auth-pki and certificate-authority list, show and count).",
    "The three root /certificateBasedAuthConfiguration reads below carry an explicit reviewed unavailable disposition; no new commands or raw access are approved for them.",
    "The v1.0 certificateBasedAuthConfiguration resource Methods table documents List/Create/Get/Delete only on the org-scoped routes, which belong to the deferred organization certificate-auth subfamily.",
    "Certificate-authority entries carry public certificates only; the base64 certificate blob is omitted from default selects and needs an explicit --select naming it.",
    "Certificate-auth reads support v1.0 only; beta needs its own review.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...unavailableCertAuthRows.map(row => `| \`${row.id}\` | ${row.disposition} | ${row.owningSlice} | ${row.reason} |`),
    "## EXT-01 directory-subscriptions scope decisions",
    "",
    "This change covers the three v1.0 commercial-subscription reads above (subscription list, show and count).",
    "The commerceSubscriptionId alternate-key lookup below remains scheduled with an explicit deferred disposition to a later EXT-01 subfamily; no new command or raw access is approved for it.",
    "Alternate-key function segments (key='value') are not whole-segment placeholders, so the shared session path template and the raw-route matcher cannot bind them without their own contract review.",
    "Subscription reads support v1.0 only; beta needs its own review.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredSubscriptionRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "## EXT-01 group lifecycle scope decisions",
    "",
    "This change covers the six v1.0 group lifecycle policy and group setting template reads above (list, show, count per family).",
    "The operations below remain scheduled with an explicit deferred disposition to a later EXT-01 subfamily; no new commands or raw access are approved.",
    "Beta policies and templates need their own review; delta-token sync needs its own change-tracking contract beyond the approved list/show/count reads; the POST lookup actions need their own request and projection review.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredGroupLifecycleRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-01 custom-security-attributes scope decisions",
    "",
    "This change covers the nine v1.0 attribute-set, custom-security-attribute-definition and allowed-value reads above (list, show, count per family).",
    "The nine beta operations below remain scheduled with an explicit deferred disposition to a later EXT-01 subfamily; no new commands or raw access are approved.",
    "Beta attribute sets, definitions and allowed values need their own review; definition $expand (inline allowedValues) is not reviewed here and stays on the dedicated allowed-value list/show reads.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredCustomSecurityRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-01 on-premises-synchronization scope decisions",
    "",
    "This change covers the two v1.0 on-premises-synchronization reads above (on-premises-synchronization list and show).",
    "The $count scalar and the three beta operations below remain scheduled with an explicit deferred disposition to later EXT-01 subfamilies; no new commands or raw access are approved.",
    "On-premises-synchronization reads support v1.0 only; beta needs its own review.",
    "Scalar counts need a separate query and response contract from the approved list/show reads.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredOnPremSyncRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-01 terms-of-use scope decisions",
    "",
    "This change covers the six v1.0 agreement and agreement-acceptance reads above (agreement list/show, one agreement's acceptance list/show, and tenant-wide acceptance list/show).",
    "The acceptances $count scalar, the agreement file/localization/files/versions sub-reads and the beta operations below remain scheduled with an explicit deferred disposition to later EXT-01 subfamilies; no new commands or raw access are approved.",
    "Terms-of-use reads support v1.0 only; beta needs its own review.",
    "Scalar counts need a separate query and response contract from the approved list/show reads; agreement file contents and localizations need a separate binary-output and projection contract, and file bytes are never downloaded or printed.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredTermsOfUseRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-01 directory-objects scope decisions",
    "",
    "This change covers the three v1.0 directory-object reads above (directory-object list, show and count).",
    "The delta sync, the POST lookup/validation actions and the beta operations below remain scheduled with an explicit deferred disposition to later EXT-01 subfamilies; no new commands or raw access are approved.",
    "Directory-object reads support v1.0 only; beta needs its own review.",
    "Delta-token sync needs its own change-tracking contract beyond the approved list/show/count reads; the POST lookup and validation actions need their own request and projection review.",
    "Rows are polymorphic: only the base-type properties plus the @odata.type discriminator are projected, so subtype secrets or credentials can never appear.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredDirectoryObjectRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),
    "",
    "## EXT-04 partner contracts scope decisions",
    "",
    "Firstmate scope: approve narrowing this change to the three v1.0 partner-contract reads above (list, show, count).",
    "The operations below remain scheduled with an explicit deferred disposition to a later EXT-04 subfamily; no new commands or raw access are approved.",
    "Beta contracts need their own review; delta-token sync needs its own change-tracking contract beyond the approved list/show/count reads.",
    "",
    "| Inventory operation | Disposition | Owning slice | Deferral reason |",
    "|---|---|---|---|",
    ...deferredContractRows.map(row => `| \`${row.id}\` | ${row.disposition} (deferred) | ${row.owningSlice} | ${row.reason} |`),    "## EXT-02 access-review scope decisions",
    "",
    "The access-review scope covers the ten reads above (definition, instance, decision, contacted-reviewer and stage list/show reads).",
    "The legacy accessReviews and unified alias reads below are split into the EXT-02c follow-up to limit this piece's size; beta-only operations stay deferred (\"beta needs its own review\").",
    "The history reads below are blocked: their documented least privilege is the write scope AccessReview.ReadWrite.All (no read scope), and history instances return SAS download URLs in downloadUri.",
    "No new commands or raw access are approved for any row below.",
    "",
    "| Inventory operation | Disposition | Deferral reason |",
    "|---|---|---|",
    ...deferredAccessReviewRows.map(row => `| \`${row.id}\` | ${row.disposition}: ${row.owningSlice} | ${row.reason} |`),
    "",
    "## Named writes",
    "",
    "Each write below runs the WRITE-00 mutation coordinator: hand-enabled profile, immutable scope, preview, explicit `--execute`, typed target confirmation for disruptive effects, durable journal intent/outcome and no replay. Inventory dispositions stay discovery-time records until a metadata refresh reviews them; the implemented review lives beside each command.",
    "",
    "| Command | Operation | Owning slice |",
    "|---|---|---|",
    ...writeRows,
    "",
    `Shipped initial writes: ${writeStatus.shipped.length ? writeStatus.shipped.join(", ") : "none"}.`,
    `Pending initial writes: ${writeStatus.pending.length ? writeStatus.pending.join(", ") : "none"}.`,
    "Extended families (EXT-01 through EXT-04) and the full-Entra audit (FULL-01, COMPLETE-01) own the remaining",
    "scheduled rows; see docs/build-plan.md for their dispatch.",
    "",
  ].join("\n");
}
