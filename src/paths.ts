import { homedir } from "node:os";
import { isAbsolute, join, win32 } from "node:path";

export function toolDirectory(kind: "config" | "state", env: NodeJS.ProcessEnv = process.env, platform = process.platform, home = homedir()): string {
  if (platform === "win32") {
    const base = env[kind === "config" ? "APPDATA" : "LOCALAPPDATA"];
    return win32.join(base || win32.join(home, "AppData", kind === "config" ? "Roaming" : "Local"), "mg-axi");
  }
  const base = env[kind === "config" ? "XDG_CONFIG_HOME" : "XDG_STATE_HOME"];
  return join(base && isAbsolute(base) ? base : join(home, kind === "config" ? ".config" : ".local/state"), "mg-axi");
}

export function configPath(env: NodeJS.ProcessEnv = process.env): string {
  return env.MG_AXI_CONFIG ?? join(toolDirectory("config", env), "config.json");
}

export function statePath(name: string, env: NodeJS.ProcessEnv = process.env): string {
  return join(toolDirectory("state", env), name);
}
