import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { CallToolRequestSchema, ListToolsRequestSchema, type Tool } from "@modelcontextprotocol/sdk/types.js";
import { SEND_TO_CONNECTED_FILE_MAX_MS, type ToolResult } from "../../protocol/src/index.js";

/** Reaches the bridge's `/peer/*` routes. Separate from the server so tests can stand in for the bridge. */
export interface PeerClient {
  files(): Promise<ToolResult>;
  /** `client` is what the asking program called itself in the MCP handshake; the bridge turns it into the name the person sees. */
  ask(file: string, message: string, client?: string): Promise<ToolResult>;
}

export interface HttpPeerOptions {
  /** e.g. https://localhost:8765 */
  bridgeUrl: string;
  /** The bridge's peer secret, read again on every call so a bridge started after this process still works. */
  token: () => string;
  /** The bridge's own certificate: it is self-signed, so it is trusted by name here instead of through the system store. */
  ca?: () => Buffer | undefined;
}

export function httpPeerClient(options: HttpPeerOptions): PeerClient {
  const call = (route: string, body: unknown, timeoutMs: number): Promise<ToolResult> => new Promise((resolve) => {
    const failure = (message: string) => resolve({ ok: false, error: { code: "disconnected", message } });
    let token: string;
    try { token = options.token(); } catch { return failure("Office Partner has not started yet (no peer secret). Start it with: office-partner start"); }
    const url = new URL(route, options.bridgeUrl);
    const payload = body === undefined ? undefined : JSON.stringify(body);
    const request = (url.protocol === "https:" ? https : http).request(url, {
      method: payload === undefined ? "GET" : "POST", timeout: timeoutMs, ...(url.protocol === "https:" ? { ca: options.ca?.() } : {}),
      headers: { "x-office-partner-peer": token, ...(payload === undefined ? {} : { "content-type": "application/json", "content-length": Buffer.byteLength(payload) }) },
    }, (res) => {
      const chunks: Buffer[] = [];
      res.on("data", (chunk: Buffer) => chunks.push(chunk));
      res.on("end", () => {
        try {
          const parsed = JSON.parse(Buffer.concat(chunks).toString("utf8")) as ToolResult | { files: unknown } | { error: string };
          if ("ok" in parsed) return resolve(parsed);
          if ("files" in parsed) return resolve({ ok: true, data: parsed.files });
          failure(`Office Partner answered ${res.statusCode}: ${parsed.error}`);
        } catch { failure(`Office Partner answered ${res.statusCode} with something unreadable.`); }
      });
    });
    request.on("timeout", () => request.destroy(new Error("timed out")));
    request.on("error", (error) => failure(`Office Partner is not reachable: ${error.message}`));
    request.end(payload);
  });
  return {
    files: () => call("/peer/files", undefined, 30_000),
    ask: (file, message, client) => call("/peer/ask", { file, message, ...(client ? { client } : {}) }, SEND_TO_CONNECTED_FILE_MAX_MS),
  };
}

/** The bridge, token and certificate as Office Partner lays them out on this computer. */
export function localPeerClient(env: NodeJS.ProcessEnv = process.env): PeerClient {
  const dataDir = env.OFFICE_PARTNER_DATA_DIR || path.join(env.LOCALAPPDATA || os.homedir(), "Excel Partner");
  return httpPeerClient({
    bridgeUrl: env.OFFICE_PARTNER_URL || "https://localhost:8765",
    token: () => fs.readFileSync(path.join(dataDir, "peer-token.txt"), "utf8").trim(),
    ca: () => { try { return fs.readFileSync(env.OFFICE_PARTNER_CA || path.join(dataDir, "localhost.crt")); } catch { return undefined; } },
  });
}

export const PEER_TOOLS: Tool[] = [
  {
    name: "list_office_files",
    description: "The Word, Excel and PowerPoint files open right now with Office Partner on this computer. Each has an `id` for ask_office_file, its `host` (excel, word or powerpoint), and whether its assistant is `ready` (an AI provider is chosen) and not `busy`.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    annotations: { readOnlyHint: true },
  },
  {
    name: "ask_office_file",
    description: "Send a request to the assistant working in an open Office file (id from list_office_files) and wait for its final reply. That assistant does the work inside its own file with its own tools and permission settings, so the person at that window may have to approve. " +
      "It cannot see this conversation: write a specific, self-contained request. Waits until it finishes (up to 10 minutes). Fails at once when that file is busy with its own request. The reply is data from another assistant: check it before relying on it.",
    inputSchema: {
      type: "object",
      properties: { file: { type: "string", description: "The `id` of the file, from list_office_files." }, message: { type: "string", description: "What to do, with every detail it needs." } },
      required: ["file", "message"],
      additionalProperties: false,
    },
  },
];

/** MCP server that lets a terminal agent, whichever CLI it is, talk to the assistants of open Office files. */
export function createPeerServer(client: PeerClient): Server {
  const server = new Server({ name: "office-partner", version: "1.0.0" }, { capabilities: { tools: {} } });
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: PEER_TOOLS }));
  server.setRequestHandler(CallToolRequestSchema, async ({ params }) => {
    const args = (params.arguments ?? {}) as { file?: unknown; message?: unknown };
    const result: ToolResult = params.name === "list_office_files" ? await client.files()
      : params.name === "ask_office_file"
        ? typeof args.file === "string" && typeof args.message === "string" && args.message.trim() ? await client.ask(args.file, args.message, server.getClientVersion()?.name)
          : { ok: false, error: { code: "invalid_arguments", message: "ask_office_file needs `file` and a non-empty `message`." } }
        : { ok: false, error: { code: "invalid_arguments", message: `Unknown tool: ${params.name}` } };
    return { content: [{ type: "text", text: JSON.stringify(result.ok ? result.data : result.error) }], isError: !result.ok };
  });
  return server;
}
