#!/usr/bin/env node

// Session-start hook entry point: print the local-only ambient summary and
// always exit 0 on config state, never loudly, so agent startup never breaks.
// Unknown argv fails loud with exit 2 instead of a silent wrong summary.
// Static imports are avoided so a broken install still reaches the fallback.
const args = process.argv.slice(2);
if (args.length === 1 && args[0] === "--help") {
  process.stdout.write("mg-axi-hook: print the local-only mg-axi session-start summary (configured profiles, auth-cache presence, version)\nusage: mg-axi-hook [--help]\n");
} else if (args.length > 0) {
  try {
    const { encode } = await import("@toon-format/toon");
    process.stdout.write(`${encode({ error: `unexpected argument \`${args[0]}\` for \`mg-axi-hook\``, code: "VALIDATION_ERROR", help: ["Run `mg-axi-hook` with no arguments for the session summary"] })}\n`);
  } catch {
    process.stdout.write(`error: unexpected argument \`${args[0]}\` for \`mg-axi-hook\`\n`);
  }
  process.exitCode = 2;
} else {
  try {
    const { VERSION } = await import("../version.js").catch(() => ({ VERSION: "unknown" }));
    const { hookSummary } = await import("../hook.js");
    process.stdout.write(await hookSummary({ version: VERSION }));
  } catch {
    process.stdout.write("mg: status unavailable\n");
  }
  process.exitCode = 0;
}
