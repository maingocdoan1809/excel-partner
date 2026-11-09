import { ListToolsResultSchema } from "@modelcontextprotocol/sdk/types.js";
import { SEND_TO_CONNECTED_FILE_MAX_MS, type ToolResult } from "../../protocol/src/index.js";

/** Forwards a validated Excel tool call to the bridge that owns the workbook. */
export interface ToolForwarder {
  listTools?(): Promise<import("@modelcontextprotocol/sdk/types.js").Tool[]>;
  call(tool: string, args: unknown): Promise<ToolResult>;
}

export interface HttpForwarderOptions {
  bridgeUrl: string;
  sessionId: string;
  token: string;
  timeoutMs?: number;
}

/**
 * HTTP implementation used by the stdio MCP process that a provider CLI
 * launches. It authenticates with the per-workbook session token.
 */
export class HttpToolForwarder implements ToolForwarder {
  constructor(private readonly options: HttpForwarderOptions) {}

  async listTools() {
    const response = await fetch(`${this.options.bridgeUrl}/internal/tools`, { headers: { authorization: `Bearer ${this.options.token}` }, signal: AbortSignal.timeout(20_000) });
    if (!response.ok) throw new Error(`Tool catalog returned HTTP ${response.status}`);
    return ListToolsResultSchema.parse(await response.json()).tools;
  }

  async call(tool: string, args: unknown): Promise<ToolResult> {
    const controller = new AbortController();
    // A first read may wait for the user to approve access in the pane (up to two minutes). Asking another file's
    // assistant to do something waits for that assistant to finish, which takes far longer.
    const timeoutMs = tool === "send_to_connected_file"
      ? Math.max(this.options.timeoutMs ?? 0, SEND_TO_CONNECTED_FILE_MAX_MS)
      : this.options.timeoutMs ?? 130_000;
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const res = await fetch(`${this.options.bridgeUrl}/internal/tool`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          authorization: `Bearer ${this.options.token}`,
        },
        body: JSON.stringify({ sessionId: this.options.sessionId, tool, args }),
        signal: controller.signal,
      });
      if (!res.ok) {
        const text = await safeText(res);
        return {
          ok: false,
          error: {
            code: "internal_error",
            message: `bridge responded ${res.status}: ${text}`,
          },
        };
      }
      return (await res.json()) as ToolResult;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return { ok: false, error: { code: "disconnected", message: `bridge unreachable: ${message}` } };
    } finally {
      clearTimeout(timeout);
    }
  }
}

async function safeText(res: Response): Promise<string> {
  try {
    return (await res.text()).slice(0, 500);
  } catch {
    return "";
  }
}
