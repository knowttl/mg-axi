import { readFile } from "node:fs/promises";
import { AxiError } from "axi-sdk-js";
import type { GraphSession } from "./graph-session.js";
import type { AnyProfile } from "./profiles.js";

export const HUNTING_OPERATION = "v1.0:POST:/security/microsoft.graph.security.runHuntingQuery";
export const HUNTING_SCOPE = "https://graph.microsoft.com/ThreatHunting.Read.All";
const HELP = "mg-axi security hunting query --help";
const TABLES = new Set("DeviceEvents DeviceProcessEvents DeviceNetworkEvents DeviceFileEvents DeviceFileCertificateInfo DeviceRegistryEvents DeviceLogonEvents DeviceImageLoadEvents DeviceInfo DeviceNetworkInfo DeviceTvmSoftwareInventory DeviceTvmSoftwareVulnerabilities DeviceTvmSoftwareVulnerabilitiesKB EmailEvents EmailPostDeliveryEvents EmailUrlInfo EmailAttachmentInfo UrlClickEvents IdentityLogonEvents IdentityQueryEvents IdentityDirectoryEvents CloudAppEvents AlertInfo AlertEvidence BehaviorInfo BehaviorEntities ExposureGraphNodes ExposureGraphEdges".split(" "));
const OPERATORS = new Set("where project project-away project-rename project-reorder extend summarize distinct count take limit top sort order".split(" "));
const FUNCTIONS = new Set("ago now datetime timespan time bin bin_at floor ceiling round abs iff iif case coalesce isnull isnotnull isempty isnotempty tostring toint tolong todouble tobool todatetime totimespan tolower toupper strlen substring strcat split replace_string trim trim_start trim_end extract extract_all parse_json parse_url parse_path array_length set_has_element bag_keys count countif dcount dcountif sum sumif min max avg percentile percentiles arg_max arg_min make_set make_list startofday endofday startofweek startofmonth datetime_diff datetime_add format_datetime format_timespan hash_sha256 ipv4_is_private ipv4_is_in_range".split(" "));

function invalid(message: string): never {
  throw new AxiError(message, "VALIDATION_ERROR", [HELP]);
}

// This is deliberately a restricted grammar, not a blacklist of write words.
// Only an identifier table source and allowlisted pipeline stages/functions
// survive. Strings and line comments are lexed before checking structure so
// quoted punctuation cannot split statements or escape the appended row cap.
export function huntingRequest(query: string, timespan: string, limit: number): { Query: string; Timespan: string } {
  if (!Number.isInteger(limit) || limit < 1 || limit > 10_000) invalid("--limit must be between 1 and 10000");
  const duration = /^(?:P(\d+)D|PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)$/.exec(timespan);
  const seconds = duration ? Number(duration[1] ?? 0) * 86400 + Number(duration[2] ?? 0) * 3600 + Number(duration[3] ?? 0) * 60 + Number(duration[4] ?? 0) : 0;
  if (!Number.isSafeInteger(seconds) || seconds < 1 || seconds > 30 * 86400) invalid("--timespan must be an ISO duration of 1 second to 30 days (for example P1D or PT6H)");
  if (!query.trim() || Buffer.byteLength(query, "utf8") > 100_000) invalid("KQL file must contain 1-100000 bytes of one read-only table pipeline");
  const tokens: string[] = [];
  const source: string[] = [];
  const lex = /\s+|\/\/[^\r\n]*|'(?:[^'\\\r\n]|\\[^\r\n]|'')*'|"(?:[^"\\\r\n]|\\[^\r\n]|"")*"|[A-Za-z_][A-Za-z_0-9-]*|\d+(?:\.\d+)?(?:[dhms])?|[|(),.\[\]+*/%<>=!~:-]/gy;
  let offset = 0;
  while (offset < query.length) {
    if (query.startsWith("/*", offset)) invalid("KQL block comments are unsupported; use // line comments");
    lex.lastIndex = offset;
    const match = lex.exec(query);
    if (!match) invalid("KQL must be one read-only table pipeline; statements, control commands and unsupported syntax are refused");
    const token = match![0];
    offset = lex.lastIndex;
    source.push(token.startsWith("//") ? " " : token);
    if (!/^\s|^\/\//.test(token)) tokens.push(token);
  }
  if (!TABLES.has(tokens[0] ?? "")) invalid("KQL must begin with a reviewed Defender table name listed in README");
  if (tokens.length > 1 && tokens[1] !== "|") invalid("Only a table source followed by a pipeline is supported; no statements or source functions");
  let depth = 0;
  const stack: string[] = [];
  for (let i = 1; i < tokens.length; i++) {
    const token = tokens[i]!;
    if (token === "|") {
      if (depth !== 0 || !OPERATORS.has(tokens[i + 1] ?? "")) invalid("Unsupported KQL stage; use the read-only operators listed in README");
      const operator = tokens[++i]!;
      const end = tokens[i + 1] === undefined || tokens[i + 1] === "|";
      if (operator === "count" ? !end : end) invalid("KQL pipeline stage has invalid or missing operands");
      continue;
    }
    if (token === "(" || token === "[") {
      if (token === "(" && /^[A-Za-z_]/.test(tokens[i - 1] ?? "") && !FUNCTIONS.has(tokens[i - 1]!) && !["in", "in~", "!in", "between"].includes(tokens[i - 1]!)) {
        invalid("Only the reviewed read-only KQL scalar and aggregate functions are supported");
      }
      stack.push(token);
      depth++;
    } else if (token === ")" || token === "]") {
      if (stack.pop() !== (token === ")" ? "(" : "[")) invalid("KQL has unbalanced parentheses or brackets");
      depth--;
    }
    // A dotted function call could select a stored function or plugin.
    if (token === "." && tokens[i + 2] === "(") invalid("Qualified KQL function calls are unsupported");
  }
  if (depth !== 0) invalid("KQL has unbalanced parentheses or brackets");
  // Remove comments and use a fresh line so the appended cap cannot be
  // swallowed. Preserve original token spacing inside KQL expressions.
  return { Query: `${source.join("")}\n| take ${limit + 1}`, Timespan: timespan };
}

export async function queryHunting(session: GraphSession, flags: Record<string, string | boolean>, profile: AnyProfile): Promise<Record<string, unknown>> {
  let query: string;
  try {
    query = await readFile(String(flags.file), "utf8");
  } catch {
    throw new AxiError("Cannot read the KQL file", "VALIDATION_ERROR", ["Pass --file with a readable UTF-8 KQL file", HELP]);
  }
  const timespan = String(flags.timespan ?? "P1D");
  const limit = flags.limit === undefined ? 100 : Number(flags.limit);
  // The session validates this input before credentials and constructs the
  // only permitted POST body itself.
  const raw = await session.execute({ profile, operation: {
    id: HUNTING_OPERATION, method: "POST", path: "/security/microsoft.graph.security.runHuntingQuery", version: "v1.0", disposition: "named-command", owningSlice: "SECURITY-01",
  }, hunting: { query, timespan, limit }, scopes: profile.mode === "delegated" ? [HUNTING_SCOPE] : undefined });
  const body = raw as { schema?: unknown; results?: unknown } | null;
  if (!body || !Array.isArray(body.schema) || !Array.isArray(body.results) ||
    body.schema.some(field => !field || typeof field !== "object" || typeof field.name !== "string" || typeof field.type !== "string") ||
    body.results.some(row => !row || typeof row !== "object" || Array.isArray(row))) {
    throw new AxiError("Graph returned a malformed hunting response", "GRAPH_ERROR", ["Expected schema and results arrays; treat the result as unknown"]);
  }
  let truncated = false;
  const truncate = (value: unknown): unknown => {
    if (typeof value === "string" && flags.full !== true && value.length > 500) {
      truncated = true;
      return `${value.slice(0, 500)}... (truncated, ${value.length} chars total)`;
    }
    if (Array.isArray(value)) return value.map(truncate);
    if (value && typeof value === "object") return Object.fromEntries(Object.entries(value).map(([key, child]) => [key, truncate(child)]));
    return value;
  };
  const results = truncate(body.results.slice(0, limit));
  const capped = body.results.length > limit;
  return {
    schema: truncate(body.schema), results, returned: Math.min(body.results.length, limit), timespan, complete: !capped,
    ...(capped ? { reason: "row-limit" } : {}),
    help: [
      ...(body.results.length === 0 ? ["0 hunting rows matched in the requested time window"] : []),
      ...(capped ? ["Row cap reached; no server cursor exists. Narrow --timespan or the query filters and run again"] : []),
      ...(truncated ? ["Repeat the same command with --full for complete text values; redaction and bounds remain"] : []),
    ],
  };
}
