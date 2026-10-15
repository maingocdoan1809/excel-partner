import type { ToolResult } from "../../protocol/src/index.js";

/** Forwards a validated Excel tool call to the bridge that owns the workbook. */
export interface ToolForwarder {
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

  async call(tool: string, args: unknown): Promise<ToolResult> {
    const controller = new AbortController();
    // A first read may wait for the user to approve access in the pane (up to two minutes).
    const timeout = setTimeout(() => controller.abort(), this.options.timeoutMs ?? 130_000);
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
