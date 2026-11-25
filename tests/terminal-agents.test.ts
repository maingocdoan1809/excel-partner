import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { loadConfig } from "../bridge/src/config.js";
import { peerServerFor } from "../bridge/src/terminal-agents.js";
import { terminalAgents } from "../providers/src/registry.js";
import { syncTerminalAgents, terminalAgentLabel, type TerminalMcpServer } from "../providers/src/terminal-agent.js";

const server: TerminalMcpServer = { name: "office-partner", command: "node", args: ["C:\\app\\peer-cli.js"], env: { OFFICE_PARTNER_DATA_DIR: "C:\\data" } };
const agents = terminalAgents();
const claude = agents.find(agent => agent.id === "claude")!;

/** A computer with some CLIs installed: records every command the sync runs. */
function computer(installed: string[], failAdd: string[] = []) {
  const calls: string[][] = [];
  const run = async (command: string, args: string[]) => {
    calls.push([command, ...args]);
    if (!installed.includes(command)) return { code: null, stdout: "", stderr: "not found", errorCode: "ENOENT" };
    return { code: args[0] === "mcp" && args[1] === "add" && failAdd.includes(command) ? 1 : 0, stdout: "", stderr: args[1] === "add" && failAdd.includes(command) ? "refused" : "" };
  };
  return { calls, run: run as never };
}
const stateFile = () => path.join(fs.mkdtempSync(path.join(os.tmpdir(), "terminal-agents-")), "terminal-agents.json");

describe("registering the peer MCP server in the terminal CLIs", () => {
  it("knows how each CLI adds a server", () => {
    expect(agents.map(agent => agent.id)).toEqual(["claude", "copilot", "codex", "antigravity"]);
    expect(claude.spec.add(server)).toEqual(["mcp", "add", "--scope", "user", "office-partner", "-e", "OFFICE_PARTNER_DATA_DIR=C:\\data", "--", "node", "C:\\app\\peer-cli.js"]);
    expect(agents.find(agent => agent.id === "copilot")!.spec.add(server)).toEqual(["mcp", "add", "office-partner", "--env", "OFFICE_PARTNER_DATA_DIR=C:\\data", "--", "node", "C:\\app\\peer-cli.js"]);
    expect(agents.find(agent => agent.id === "antigravity")!.spec.add(server)).toEqual(["mcp", "add", "--env", "OFFICE_PARTNER_DATA_DIR=C:\\data", "office-partner", "--", "node", "C:\\app\\peer-cli.js"]);
    expect(terminalAgentLabel(agents, "Antigravity ACP Server")).toBe("Antigravity CLI (terminal)");
    expect(claude.spec.remove("office-partner")).toEqual(["mcp", "remove", "--scope", "user", "office-partner"]);
  });

  it("shows the person who is asking from the name the CLI gave in the MCP handshake, each provider supplying its own label", () => {
    expect(agents.map(agent => agent.spec.label)).toEqual(["Claude Code", "Copilot CLI", "Codex CLI", "Antigravity CLI"]);
    expect(terminalAgentLabel(agents, "claude-code")).toBe("Claude Code (terminal)");
    expect(terminalAgentLabel(agents, "GitHub-Copilot-CLI")).toBe("Copilot CLI (terminal)");
    expect(terminalAgentLabel(agents, "codex-mcp-client")).toBe("Codex CLI (terminal)");
    expect(terminalAgentLabel(agents, "my-agent")).toBe("my-agent (terminal)");
    expect(terminalAgentLabel(agents, undefined)).toBe("Terminal agent");
    expect(terminalAgentLabel(agents, "  ")).toBe("Terminal agent");
  });

  it("registers in the CLIs that are installed and skips the rest", async () => {
    const pc = computer(["claude"]);
    const results = await syncTerminalAgents(agents, server, { stateFile: stateFile(), run: pc.run });
    expect(results).toEqual({ claude: { outcome: "registered" }, copilot: { outcome: "not_installed" }, codex: { outcome: "not_installed" }, antigravity: { outcome: "not_installed" } });
    expect(pc.calls.filter(call => call[0] === "claude").map(call => call.slice(1, 3).join(" "))).toEqual(["--version", "mcp remove", "mcp add"]);
    expect(pc.calls.some(call => call[0] === "copilot" && call[1] === "mcp")).toBe(false);
  });

  it("leaves a CLI alone once it has this registration, so a server the person removed stays removed", async () => {
    const file = stateFile();
    await syncTerminalAgents(agents, server, { stateFile: file, run: computer(["claude"]).run });
    const again = computer(["claude"]);
    expect((await syncTerminalAgents(agents, server, { stateFile: file, run: again.run })).claude).toEqual({ outcome: "unchanged" });
    expect(again.calls.some(call => call[1] === "mcp")).toBe(false);
  });

  it("registers again when the files moved (a new release)", async () => {
    const file = stateFile();
    await syncTerminalAgents(agents, server, { stateFile: file, run: computer(["claude"]).run });
    const moved = computer(["claude"]);
    const results = await syncTerminalAgents(agents, { ...server, args: ["C:\\new\\peer-cli.js"] }, { stateFile: file, run: moved.run });
    expect(results.claude).toEqual({ outcome: "registered" });
    expect(moved.calls.filter(call => call[0] === "claude").map(call => call.slice(1, 3).join(" "))).toEqual(["--version", "mcp remove", "mcp add"]);
  });

  it("reports a CLI that refuses without touching the others, and tries again next start", async () => {
    const file = stateFile();
    const results = await syncTerminalAgents(agents, server, { stateFile: file, run: computer(["claude", "copilot"], ["claude"]).run });
    expect(results.claude).toEqual({ outcome: "failed", detail: "refused" });
    expect(results.copilot).toEqual({ outcome: "registered" });
    expect((await syncTerminalAgents(agents, server, { stateFile: file, run: computer(["claude", "copilot"]).run })).claude).toEqual({ outcome: "registered" });
  });

  it("does nothing when the person turned it off", async () => {
    const file = stateFile();
    fs.writeFileSync(file, JSON.stringify({ disabled: true }));
    const pc = computer(["claude"]);
    expect((await syncTerminalAgents(agents, server, { stateFile: file, run: pc.run })).claude).toEqual({ outcome: "disabled" });
    expect(pc.calls).toEqual([]);
    const byEnv = computer(["claude"]);
    expect((await syncTerminalAgents(agents, server, { stateFile: stateFile(), env: { OFFICE_PARTNER_TERMINAL_AGENTS: "0" }, run: byEnv.run })).claude).toEqual({ outcome: "disabled" });
    expect(byEnv.calls).toEqual([]);
  });

  it("starts the peer server from the same release and node as the Excel MCP, with where the bridge keeps its secret", () => {
    const config = loadConfig({ tls: null, publicBaseUrl: "http://127.0.0.1:8765", configFile: path.join("C:\\data", "excel-partner.config.json"), mcp: { command: "node.exe", args: [path.join("C:\\rel", "dist", "excel_mcp", "src", "cli.js")] } });
    expect(peerServerFor(config)).toEqual({
      name: "office-partner", command: "node.exe", args: [path.join("C:\\rel", "dist", "excel_mcp", "src", "peer-cli.js")],
      env: { OFFICE_PARTNER_DATA_DIR: "C:\\data", OFFICE_PARTNER_URL: "http://127.0.0.1:8765" },
    });
  });
});
