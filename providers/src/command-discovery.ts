import spawn from "cross-spawn";
import { createInterface } from "node:readline";
import { JsonRpcPeer } from "./jsonrpc.js";
import { normalizeCommands } from "./commands.js";
import type { ProviderCommand } from "../../protocol/src/commands.js";

/** Initialize only: never submit a user prompt or request model inference. */
export function discoverClaudeCommands(command: string, env: NodeJS.ProcessEnv, cwd: string): Promise<ProviderCommand[]> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, ["-p", "--input-format", "stream-json", "--output-format", "stream-json", "--verbose",
      "--setting-sources", "project", "--tools", "", "--strict-mcp-config", "--mcp-config", '{"mcpServers":{}}'],
    { env, cwd, stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
    let settled = false;
    let lastType = "none", linesSeen = 0, invalidLines = 0;
    const finish = (error?: Error, commands?: ProviderCommand[]) => {
      if (settled) return;
      settled = true; clearTimeout(timer); lines.close(); child.stdin?.destroy(); child.kill();
      if (error) reject(error); else resolve(commands ?? []);
    };
    const timer = setTimeout(() => finish(new Error(`Claude command discovery timed out (${linesSeen} lines, ${invalidLines} non-JSON; last event: ${lastType})`)), 15_000);
    const lines = createInterface({ input: child.stdout! });
    child.stderr?.resume();
    child.stdin?.on("error", () => undefined);
    child.once("error", error => finish(error));
    child.once("exit", () => finish(new Error("Claude exited before returning its commands")));
    lines.on("line", line => {
      linesSeen++;
      let message: any;
      try { message = JSON.parse(line); } catch { invalidLines++; return; }
      lastType = String(message.type ?? "unknown");
      if (message.type !== "control_response" || message.response?.request_id !== "commands-init") return;
      if (message.response.subtype !== "success") return finish(new Error(message.response.error ?? "Claude initialization failed"));
      finish(undefined, normalizeCommands(message.response.response?.commands));
    });
    child.stdin!.write(JSON.stringify({ type: "control_request", request_id: "commands-init", request: { subtype: "initialize" } }) + "\n");
  });
}

export async function discoverAcpCommands(command: string, args: string[], env: NodeJS.ProcessEnv, cwd: string): Promise<ProviderCommand[]> {
  let resolveCommands!: (commands: ProviderCommand[]) => void;
  const received = new Promise<ProviderCommand[]>(resolve => { resolveCommands = resolve; });
  const peer = new JsonRpcPeer({ command, args, env, cwd, versionField: true,
    onNotification: (method, params: any) => {
      if (method === "session/update" && params?.update?.sessionUpdate === "available_commands_update") {
        resolveCommands(normalizeCommands(params.update.availableCommands));
      }
    }, onRequest: () => ({ outcome: { outcome: "cancelled" } }),
  });
  let timer: NodeJS.Timeout | undefined;
  try {
    await peer.request("initialize", { protocolVersion: 1, clientCapabilities: {}, clientInfo: { name: "excel-partner-commands", version: "1" } }, 15_000);
    await peer.request("session/new", { cwd, mcpServers: [] }, 15_000);
    return await Promise.race([received, new Promise<ProviderCommand[]>(resolve => { timer = setTimeout(() => resolve([]), 3_000); })]);
  } finally { clearTimeout(timer); await peer.close(); }
}
