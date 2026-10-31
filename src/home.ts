import { AxiError } from "axi-sdk-js";
import { GraphSession, resolveSessionOperation } from "./graph-session.js";
import { fetchTransport } from "./api.js";
import type { DispatchOverrides } from "./cli.js";
import type { home } from "./catalogue.js";
import type { Profiles } from "./profiles.js";
import { skillHint } from "./skill-hint.js";

// Home only samples collections. No continuation, retry, hunting query,
// mailbox discovery, interactive authentication or permission escalation.
export async function liveHome(output: ReturnType<typeof home>, selected: ReturnType<Profiles["resolve"]>, overrides: DispatchOverrides, config?: string) {
  const { profile, name } = selected;
  const context = [selected.source === "environment" ? "" : `--profile ${name}`,
    config?.trim() ? `--config '${config.trim().replaceAll("'", "'\\''")}'` : ""].filter(Boolean).join(" ");
  const domains: { name: string; status: string; summary: string }[] = [];
  const help = [`mg-axi home --offline ${context}`];
  let session: GraphSession | undefined;
  for (const pack of ["entra", "mail", "security"] as const) {
    if (!profile.enabledPacks.includes(pack)) {
      domains.push({ name: pack, status: "disabled", summary: "Disabled in selected profile; no Graph request made" });
      continue;
    }
    if (pack === "security" || (pack === "mail" && profile.mode === "application")) {
      domains.push({ name: pack, status: "unsupported", summary: pack === "security"
        ? "Hunting requires an explicit query; no automatic Graph read"
        : "Application mail requires an explicit mailbox; no automatic Graph read" });
      help.push(pack === "security" ? `mg-axi security hunting query ${context} --help` : `mg-axi mail folder list ${context} --help`);
      continue;
    }
    const scope = `https://graph.microsoft.com/${pack === "entra" ? "User.Read.All" : "Mail.Read"}`;
    const command = `${pack === "entra" ? `mg-axi entra user list ${context} --limit 5` : `mg-axi mail folder list --mailbox me ${context} --limit 5`}${profile.mode === "delegated" ? ` --scopes ${scope}` : ""}`;
    try {
      if (!session) {
        const delegated = overrides.delegated ?? new (await import("./auth.js")).DelegatedAuth(new (await import("./msal-provider.js")).MsalProvider());
        const application = overrides.application ?? new (await import("./app-auth.js")).ApplicationAuth(new (await import("./msal-app-provider.js")).MsalApplicationProvider());
        session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
      }
      const result = await session.collect({ profile,
        operation: resolveSessionOperation("v1.0", "GET", pack === "entra" ? "/users" : "/me/mailFolders"),
        scopes: profile.mode === "delegated" ? [scope] : undefined,
        query: { $select: "id", $top: "5" }, limit: 5,
        budget: { maxRequests: 1, maxBytes: 64_000, deadlineMs: 5_000 },
      });
      const noun = pack === "entra" ? "users" : "root folders";
      const unavailable = !result.complete && !result.value.length;
      domains.push({ name: pack, status: unavailable ? "unavailable" : "live",
        summary: unavailable ? "Graph sample incomplete; no count established"
          : `${result.value.length} ${noun} observed; total ${result.total ?? (result.complete ? result.value.length : "unknown")}; ${result.complete ? "complete" : "sample only"}` });
      help.push(command);
    } catch (error) {
      const code = error instanceof AxiError ? error.code : "GRAPH_ERROR";
      const authentication = code.startsWith("AUTH_");
      domains.push({ name: pack, status: "unavailable", summary: authentication
        ? `${code}: silent credential unavailable for existing ${profile.mode === "delegated" ? scope.split("/").at(-1) : ".default"} access; no interactive sign-in attempted`
        : `${code}: Graph read unavailable or denied; check existing grants, role, licence, policy and connectivity` });
      help.push(authentication ? `mg-axi auth status ${context}` : `${command} --help`);
    }
  }
  const hint = skillHint();
  if (hint !== undefined) help.push(hint);
  return { ...output, domains, help };
}
