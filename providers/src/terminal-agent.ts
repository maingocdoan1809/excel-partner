import { createHash } from "node:crypto";
import fs from "node:fs";
import type { ProviderId } from "../../protocol/src/index.js";
import { execCli } from "./base.js";

/** An MCP server to register in a terminal CLI: the command that starts it and the environment it needs. */
export interface TerminalMcpServer {
  name: string;
  command: string;
  args: string[];
  env: Record<string, string>;
}

/**
 * How a provider's own terminal CLI (the one the person runs in a shell, not the copy the product manages) adds and removes
 * an MCP server in its user-level configuration. Each adapter declares its own as `static terminal`.
 */
export interface TerminalAgentSpec {
  /** The executable looked up on PATH. */
  command: string;
  /** What the person sees when this CLI asks for something ("Claude Code"). */
  label: string;
  /** Matches the name this CLI gives itself in the MCP handshake. */
  client: RegExp;
  add(server: TerminalMcpServer): string[];
  remove(name: string): string[];
}

/** The form `copilot mcp add` and `codex mcp add` share: `<cli> mcp add <name> --env K=V -- <command> <args>`. */
export function mcpAddRemove(command: string, identity: Pick<TerminalAgentSpec, "label" | "client">): TerminalAgentSpec {
  return {
    command,
    ...identity,
    add: server => ["mcp", "add", server.name, ...Object.entries(server.env).flatMap(([key, value]) => ["--env", `${key}=${value}`]), "--", server.command, ...server.args],
    remove: name => ["mcp", "remove", name],
  };
}

/** Who is asking, as the person sees it: the label of the provider whose CLI introduced itself with this MCP client name. */
export function terminalAgentLabel(agents: Array<{ spec: TerminalAgentSpec }>, clientName?: string): string {
  const name = clientName?.trim();
  if (!name) return "Terminal agent";
  return `${agents.find(({ spec }) => spec.client.test(name))?.spec.label ?? name.slice(0, 40)} (terminal)`;
}

export type TerminalAgentOutcome = "registered" | "unchanged" | "not_installed" | "disabled" | "failed";

interface StateFile {
  /** Turn automatic registration off for every CLI. */
  disabled?: boolean;
  /** What was registered in each CLI. A server the person removed stays removed until this changes (a new release moves the files). */
  registered?: Partial<Record<ProviderId, string>>;
}

export interface TerminalSyncOptions {
  stateFile: string;
  env?: NodeJS.ProcessEnv;
  run?: typeof execCli;
}

const signature = (server: TerminalMcpServer) => createHash("sha256").update(JSON.stringify([server.command, server.args, server.env])).digest("hex");

function readState(file: string): StateFile {
  try { return JSON.parse(fs.readFileSync(file, "utf8")) as StateFile; } catch { return {}; }
}

/**
 * Registers the server in every terminal CLI found on this computer, so a terminal assistant can use it without the person
 * running anything. Safe to call on every start: a CLI that already has this exact registration is left alone, and one
 * that is not installed is skipped. Never throws; the outcome of each CLI is returned.
 */
export async function syncTerminalAgents(
  agents: Array<{ id: ProviderId; spec: TerminalAgentSpec }>,
  server: TerminalMcpServer,
  options: TerminalSyncOptions,
): Promise<Record<string, { outcome: TerminalAgentOutcome; detail?: string }>> {
  const env = options.env ?? process.env, run = options.run ?? execCli;
  const state = readState(options.stateFile);
  const results: Record<string, { outcome: TerminalAgentOutcome; detail?: string }> = {};
  if (state.disabled || env.OFFICE_PARTNER_TERMINAL_AGENTS === "0") {
    for (const { id } of agents) results[id] = { outcome: "disabled" };
    return results;
  }
  const wanted = signature(server);
  const registered = { ...state.registered };
  await Promise.all(agents.map(async ({ id, spec }) => {
    try {
      const version = await run(spec.command, ["--version"], env, 15_000);
      if (version.code !== 0) { results[id] = { outcome: "not_installed" }; return; }
      if (registered[id] === wanted) { results[id] = { outcome: "unchanged" }; return; }
      // An older registration (another release's files) is replaced; a missing one makes the removal fail, which is fine.
      await run(spec.command, spec.remove(server.name), env, 30_000);
      const added = await run(spec.command, spec.add(server), env, 30_000);
      if (added.code !== 0) { results[id] = { outcome: "failed", detail: (added.stderr || added.stdout).trim().slice(0, 300) }; return; }
      registered[id] = wanted;
      results[id] = { outcome: "registered" };
    } catch (error) {
      results[id] = { outcome: "failed", detail: error instanceof Error ? error.message : String(error) };
    }
  }));
  if (JSON.stringify(registered) !== JSON.stringify(state.registered ?? {})) {
    try { fs.writeFileSync(options.stateFile, JSON.stringify({ ...state, registered }, null, 2), { mode: 0o600 }); } catch { /* it is registered; it will be checked again next start */ }
  }
  return results;
}
