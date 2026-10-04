import { encode } from "@toon-format/toon";
import { AxiError, runAxiCli } from "axi-sdk-js";
import { home, leafHelp, operationFor, resolveCommand, DESCRIPTION, TOP_LEVEL_HELP } from "./catalogue.js";
import { VERSION } from "./version.js";
import { Profiles } from "./profiles.js";
import { GraphSession, MAX_CURSOR_BYTES, type GraphTransport } from "./graph-session.js";
import { listUsers, showUser } from "./entra-users.js";
import { TRANSITIVE_OPERATION, listGroupMemberOf, listGroupMembers, listGroups, showGroup } from "./entra-groups.js";
import { listSignIns, showSignIn, listDirectoryAudits, showDirectoryAudit } from "./entra-audit-logs.js";
import { listApplicationOwners, listApplications, listServicePrincipalOwners, listServicePrincipals, showApplication, showServicePrincipal } from "./entra-apps.js";
import { fetchTransport } from "./api.js";
import type { DelegatedAuth } from "./auth.js";
import type { ApplicationAuth } from "./app-auth.js";

function localHome(name?: string) {
  const output = home();
  const store = new Profiles();
  if (name || store.list().length) {
    try {
      const selected = store.resolve(name);
      output.profile = selected.name;
      output.tenant = selected.profile.tenantId;
    } catch (error) {
      if (name || !(error instanceof AxiError) || error.code !== "AUTH_REQUIRED") throw error;
      output.profile = "unavailable: no default profile selected";
    }
  }
  output.help.unshift("mg-axi profile list", "mg-axi login --help");
  return output;
}

// Test seam over true external boundaries only: the packaged dispatch stays
// identical while offline journeys substitute fixture credential services and
// transports. Production callers pass no overrides and reach MSAL + HTTPS.
export interface DispatchOverrides {
  transport?: GraphTransport;
  delegated?: DelegatedAuth;
  application?: ApplicationAuth;
}

async function readCursor(cursor: string | undefined): Promise<string | undefined> {
  if (cursor !== "-") return cursor;
  const chunks: Buffer[] = [];
  let bytes = 0;
  for await (const chunk of process.stdin) {
    const buffer = Buffer.from(chunk);
    bytes += buffer.length;
    if (bytes > MAX_CURSOR_BYTES) throw new AxiError(`Collection cursor exceeds ${MAX_CURSOR_BYTES} bytes`, "VALIDATION_ERROR", ["Use a cursor within the supported size ceiling"]);
    chunks.push(buffer);
  }
  return Buffer.concat(chunks).toString("utf8");
}

export async function executeArgv(argv: string[], overrides: DispatchOverrides = {}): Promise<string | Record<string, unknown>> {
  const { leaf, flags, positional } = resolveCommand(argv);
  if (flags.help) return leafHelp(leaf);
  if (leaf.path === "home") return localHome(flags.profile as string | undefined);
  const profiles = new Profiles();
  if (leaf.path === "profile create") {
    if (String(flags.mode ?? "delegated") !== "application" && (flags["certificate-thumbprint"] !== undefined || flags.federated)) throw new AxiError("Certificate and federated credentials belong to application profiles; pass --mode application", "VALIDATION_ERROR", [leafHelp(leaf)]);
    return profiles.create(String(flags.name), String(flags.tenant), String(flags.client), String(flags.cloud), !!flags["allow-device-code"],
    String(flags.mode ?? "delegated") === "application"
      ? { certificateThumbprint: flags["certificate-thumbprint"] === undefined ? undefined : String(flags["certificate-thumbprint"]), federated: !!flags.federated }
      : undefined);
  }
  if (leaf.path === "profile list") {
    const items = profiles.list();
    return items.length ? { profiles: items, help: ["mg-axi profile show --profile <name>", "mg-axi login --help"] } : { profiles: "0 profiles configured", help: ["mg-axi profile create --help"] };
  }
  if (leaf.path === "profile show") return profiles.resolve(flags.profile as string | undefined);
  if (leaf.path === "login") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    if (selected.profile.mode !== "delegated") throw new AxiError("Application profiles authenticate with client credentials; interactive login is unavailable", "VALIDATION_ERROR", ["mg-axi profile show --profile <name>", "Application tokens are acquired silently with the configured Graph .default audience"]);
    const { DelegatedAuth } = await import("./auth.js");
    const { MsalProvider } = await import("./msal-provider.js");
    return { profile: selected.name, ...await new DelegatedAuth(new MsalProvider()).login(selected.profile, String(flags.method ?? "browser"), String(flags.scopes).split(",")) };
  }
  if (leaf.path === "api get") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const { runApiGet } = await import("./api.js");
    const { DelegatedAuth } = await import("./auth.js");
    const { ApplicationAuth } = await import("./app-auth.js");
    const { MsalProvider } = await import("./msal-provider.js");
    const { MsalApplicationProvider } = await import("./msal-app-provider.js");
    const cursor = await readCursor(flags.cursor === undefined ? undefined : String(flags.cursor));
    return runApiGet({
      path: positional!,
      apiVersion: String(flags["api-version"] ?? "v1.0"),
      odata: flags.odata === undefined ? undefined : String(flags.odata),
      cursor,
      scopes: flags.scopes === undefined ? undefined : String(flags.scopes),
      limit: flags.all ? undefined : flags.limit === undefined ? 100 : Number(flags.limit),
      full: !!flags.full,
      profile: selected.profile,
    }, {
      delegated: new DelegatedAuth(new MsalProvider()),
      application: new ApplicationAuth(new MsalApplicationProvider()),
      transport: fetchTransport,
    });
  }
  if (leaf.path === "entra user list" || leaf.path === "entra user show") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    return leaf.path === "entra user list"
      ? listUsers(session, flags, selected.profile, operation, leafHelp(leaf), selected.name)
      : showUser(session, flags, selected.profile, operation, leafHelp(leaf), selected.name);
  }
  if (leaf.path === "entra group list" || leaf.path === "entra group show" || leaf.path === "entra group member list" || leaf.path === "entra group member-of list") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const template = flags.transitive === true
      ? (() => {
        const alternate = TRANSITIVE_OPERATION[leaf.operation!];
        if (!alternate) throw new AxiError("--transitive is available for group member and member-of lists only", "VALIDATION_ERROR", [leafHelp(leaf)]);
        return alternate;
      })()
      : leaf.operation!;
    const operation = operationFor({ ...leaf, operation: template }, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    if (leaf.path === "entra group list") return listGroups(session, flags, selected.profile, operation, leafHelp(leaf), selected.name);
    if (leaf.path === "entra group show") return showGroup(session, flags, selected.profile, operation, leafHelp(leaf), selected.name);
    return leaf.path === "entra group member list"
      ? listGroupMembers(session, flags, selected.profile, operation, leafHelp(leaf), selected.name)
      : listGroupMemberOf(session, flags, selected.profile, operation, leafHelp(leaf), selected.name);
  }
  if (leaf.path === "entra sign-in list" || leaf.path === "entra sign-in show" || leaf.path === "entra directory-audit list" || leaf.path === "entra directory-audit show" || leaf.path === "entra application list" || leaf.path === "entra application show" || leaf.path === "entra service-principal list" || leaf.path === "entra service-principal show" || leaf.path === "entra application owner list" || leaf.path === "entra service-principal owner list") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    if ((leaf.path.startsWith("entra sign-in ") || leaf.path.startsWith("entra directory-audit ")) && flags.cursor !== undefined) flags.cursor = (await readCursor(String(flags.cursor)))!;
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra sign-in list": return listSignIns(session, flags, selected.profile, operation, help, selected.name);
      case "entra sign-in show": return showSignIn(session, flags, selected.profile, operation, help, selected.name);
      case "entra directory-audit list": return listDirectoryAudits(session, flags, selected.profile, operation, help, selected.name);
      case "entra directory-audit show": return showDirectoryAudit(session, flags, selected.profile, operation, help, selected.name);
      case "entra application list": return listApplications(session, flags, selected.profile, operation, help, selected.name);
      case "entra application show": return showApplication(session, flags, selected.profile, operation, help, selected.name);
      case "entra service-principal list": return listServicePrincipals(session, flags, selected.profile, operation, help, selected.name);
      case "entra service-principal show": return showServicePrincipal(session, flags, selected.profile, operation, help, selected.name);
      case "entra application owner list": return listApplicationOwners(session, flags, selected.profile, operation, help, selected.name);
      default: return listServicePrincipalOwners(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
  throw new AxiError(`Command is not executable: ${operation?.disposition ?? "unavailable"} (${operation?.owningSlice ?? "no inventory mapping"})`, "NOT_IMPLEMENTED", [leafHelp(leaf)]);
}

export async function main() {
  await runAxiCli({
    description: DESCRIPTION,
    version: VERSION,
    topLevelHelp: TOP_LEVEL_HELP,
    // Strict resolution owns leaf help, so SDK help cannot bypass validation.
    argv: ["dispatch"],
    home: () => localHome(),
    commands: {
      dispatch: async () => {
        if (process.argv.length === 3 && process.argv[2] === "--help") return TOP_LEVEL_HELP;
        return executeArgv(process.argv.slice(2));
      },
    },
    formatError: error => ({
      output: `${encode(error instanceof AxiError ? { error: error.message, code: error.code, help: error.suggestions } : { error: "Unable to run mg-axi", help: ["mg-axi --help"] })}\n`,
      exitCode: error instanceof AxiError && error.code === "VALIDATION_ERROR" ? 2 : 1,
    }),
  });
}
