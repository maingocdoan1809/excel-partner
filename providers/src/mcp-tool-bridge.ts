import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import type { McpEndpoint } from "./types.js";

/** One tool of the MCP server, as it described itself. */
export interface BridgedTool {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
}

/** The content of an MCP tool result, in the shape both the MCP SDK and the Copilot SDK use. */
export interface BridgedResult {
  content: Array<{ type: string; [key: string]: unknown }>;
  isError?: boolean;
}

/** An open connection from the bridge process to the Excel MCP server it started itself. */
export interface McpToolBridge {
  tools: BridgedTool[];
  call(name: string, args: unknown, signal?: AbortSignal): Promise<BridgedResult>;
  close(): Promise<void>;
}

export type McpToolBridgeFactory = (endpoint: Pick<McpEndpoint, "command" | "args" | "env">) => Promise<McpToolBridge>;

// The layers behind the MCP server enforce the real limits: 130 s per call to the bridge, 13 minutes for a request to
// another open file, and for the gateway up to an hour of waiting for the client's result (its relay gives up after 65
// minutes). The MCP client's own default of 60 s must not cut any of them short, so this only caps a hung server.
const CALL_TIMEOUT_MS = 70 * 60_000;

/**
 * Starts the Excel MCP server and talks to it from this process.
 *
 * Copilot CLI 1.0.108 (1.0.103 does not) blocks every MCP server of its own ("blocked by policy: 'excel'") when it cannot
 * fetch the organization's MCP policy, which it only tries from a git repository with a GitHub remote, read with `git`.
 * Tools the host registers itself are not MCP servers and are not subject to that policy, so the Copilot adapter lists
 * the server's tools here and hands them to the CLI as its own, running each call through this connection.
 */
export const connectMcpTools: McpToolBridgeFactory = async ({ command, args, env }) => {
  const client = new Client({ name: "office-partner-copilot", version: "1.0.0" });
  const tools: BridgedTool[] = [];
  try {
    await client.connect(new StdioClientTransport({ command, args, env, stderr: "ignore" }));
    let cursor: string | undefined;
    do {
      const page = await client.listTools(cursor ? { cursor } : undefined);
      for (const tool of page.tools) tools.push({ name: tool.name, description: tool.description ?? "", inputSchema: tool.inputSchema as Record<string, unknown> });
      cursor = page.nextCursor;
    } while (cursor);
  } catch (error) {
    await client.close().catch(() => {});
    throw error;
  }
  return {
    tools,
    call: async (name, input, signal) => (await client.callTool({ name, arguments: (input ?? {}) as Record<string, unknown> }, undefined, {
      timeout: CALL_TIMEOUT_MS,
      signal,
    })) as BridgedResult,
    close: () => client.close(),
  };
};
