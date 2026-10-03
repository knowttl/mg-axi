import { encode } from "@toon-format/toon";
import { AxiError, runAxiCli } from "axi-sdk-js";
import { home, leafHelp, operationFor, resolveCommand, DESCRIPTION, TOP_LEVEL_HELP } from "./catalogue.js";
import { VERSION } from "./version.js";

export async function main() {
  await runAxiCli({
    description: DESCRIPTION,
    version: VERSION,
    topLevelHelp: TOP_LEVEL_HELP,
    // Strict resolution owns leaf help, so SDK help cannot bypass validation.
    argv: ["dispatch"],
    home,
    commands: {
      dispatch: async () => {
        if (process.argv.length === 3 && process.argv[2] === "--help") return TOP_LEVEL_HELP;
        const { leaf, flags } = resolveCommand(process.argv.slice(2));
        if (flags.help) return leafHelp(leaf);
        if (leaf.path === "home") return home();
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
