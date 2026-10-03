import { encode } from "@toon-format/toon";
import { AxiError, runAxiCli } from "axi-sdk-js";
import { home, leafHelp, operationFor, resolveCommand, DESCRIPTION, TOP_LEVEL_HELP } from "./catalogue.js";
import { VERSION } from "./version.js";
import { Profiles } from "./profiles.js";

function localHome(name?: string) {
  const output = home();
  const profiles = new Profiles();
  if (name || profiles.list().length) {
    try {
      const selected = profiles.resolve(name);
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
        const { leaf, flags } = resolveCommand(process.argv.slice(2));
        if (flags.help) return leafHelp(leaf);
        if (leaf.path === "home") return localHome(flags.profile as string | undefined);
        const profiles = new Profiles();
        if (leaf.path === "profile create") return profiles.create(String(flags.name), String(flags.tenant), String(flags.client), String(flags.cloud), !!flags["allow-device-code"]);
        if (leaf.path === "profile list") {
          const items = profiles.list();
          return items.length ? { profiles: items, help: ["mg-axi profile show --profile <name>", "mg-axi login --help"] } : { profiles: "0 profiles configured", help: ["mg-axi profile create --help"] };
        }
        if (leaf.path === "profile show") return profiles.resolve(flags.profile as string | undefined);
        if (leaf.path === "login") {
          const selected = profiles.resolve(flags.profile as string | undefined);
          const { DelegatedAuth } = await import("./auth.js");
          const { MsalProvider } = await import("./msal-provider.js");
          return { profile: selected.name, ...await new DelegatedAuth(new MsalProvider()).login(selected.profile, String(flags.method ?? "browser"), String(flags.scopes).split(",")) };
        }
        const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
        throw new AxiError(`Command is not executable: ${operation?.disposition ?? "unavailable"} (${operation?.owningSlice ?? "no inventory mapping"})`, "NOT_IMPLEMENTED", [leafHelp(leaf)]);
      },
    },
    formatError: error => ({
      output: `${encode(error instanceof AxiError ? { error: error.message, code: error.code, help: error.suggestions } : { error: "Unable to run mg-axi", help: ["mg-axi --help"] })}\n`,
      exitCode: error instanceof AxiError && error.code === "VALIDATION_ERROR" ? 2 : 1,
    }),
  });
}
