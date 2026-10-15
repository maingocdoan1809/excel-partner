import http from "node:http";
import type { AddressInfo } from "node:net";

export type FakeStep =
  | { tool: string; input: unknown }
  | { text: string; delayMs?: number };

export interface FakeAnthropicRequest {
  url: string;
  body: any;
}

/**
 * Minimal Anthropic Messages API (SSE) that replays a script. Lets the REAL
 * `claude` binary run end to end (flags, MCP wiring, permissions, stream-json
 * output) without a model or login: set ANTHROPIC_BASE_URL to `url` and
 * ANTHROPIC_API_KEY to any value.
 *
 * Step selection: the number of `tool_result` blocks in the conversation so
 * far picks the script step, so a script of [tool, tool, text] behaves like
 * "call two tools then answer". Each new user message restarts the script.
 */
export class FakeAnthropic {
  readonly requests: FakeAnthropicRequest[] = [];
  private server: http.Server | null = null;
  url = "";

  constructor(private readonly script: FakeStep[]) {}

  async start(): Promise<void> {
    this.server = http.createServer((req, res) => this.handle(req, res));
    await new Promise<void>((resolve) => this.server!.listen(0, "127.0.0.1", resolve));
    this.url = `http://127.0.0.1:${(this.server.address() as AddressInfo).port}`;
  }

  async stop(): Promise<void> {
    await new Promise<void>((resolve) => (this.server ? this.server.close(() => resolve()) : resolve()));
    this.server?.closeAllConnections?.();
  }

  private handle(req: http.IncomingMessage, res: http.ServerResponse): void {
    let raw = "";
    req.on("data", (c) => (raw += c));
    req.on("end", () => {
      if (req.method !== "POST" || !req.url?.startsWith("/v1/messages")) {
        res.writeHead(200, { "content-type": "application/json" }).end("{}");
        return;
      }
      const body = JSON.parse(raw || "{}");
      this.requests.push({ url: req.url, body });
      void this.respond(body, res);
    });
  }

  private async respond(body: any, res: http.ServerResponse): Promise<void> {
    const messages: any[] = body.messages ?? [];
    // Count tool results since the last plain user message.
    let results = 0;
    for (let i = messages.length - 1; i >= 0; i--) {
      const m = messages[i];
      const blocks = Array.isArray(m.content) ? m.content : [];
      const hasResult = blocks.some((b: any) => b.type === "tool_result");
      if (m.role === "user" && !hasResult) break;
      if (m.role === "user" && hasResult) results += blocks.filter((b: any) => b.type === "tool_result").length;
    }
    const step: FakeStep = this.script[Math.min(results, this.script.length - 1)] ?? { text: "ok" };

    // Side requests (titles, haiku helpers) get a plain short answer.
    const isMain = Array.isArray(body.tools) && body.tools.length > 0;
    res.writeHead(200, { "content-type": "text/event-stream", "cache-control": "no-cache" });
    const send = (event: string, data: unknown) => res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
    const id = `msg_${Math.random().toString(36).slice(2)}`;
    send("message_start", {
      type: "message_start",
      message: {
        id,
        type: "message",
        role: "assistant",
        model: body.model ?? "fake",
        content: [],
        stop_reason: null,
        stop_sequence: null,
        usage: { input_tokens: 1, output_tokens: 1 },
      },
    });

    if (isMain && "tool" in step) {
      const toolName = body.tools.find((t: any) => t.name === `mcp__excel__${step.tool}`)?.name ?? step.tool;
      send("content_block_start", {
        type: "content_block_start",
        index: 0,
        content_block: { type: "tool_use", id: `toolu_${Math.random().toString(36).slice(2)}`, name: toolName, input: {} },
      });
      send("content_block_delta", {
        type: "content_block_delta",
        index: 0,
        delta: { type: "input_json_delta", partial_json: JSON.stringify(step.input) },
      });
      send("content_block_stop", { type: "content_block_stop", index: 0 });
      send("message_delta", { type: "message_delta", delta: { stop_reason: "tool_use", stop_sequence: null }, usage: { output_tokens: 5 } });
    } else {
      const text = "text" in step ? step.text : "ok";
      send("content_block_start", { type: "content_block_start", index: 0, content_block: { type: "text", text: "" } });
      const mid = Math.max(1, Math.floor(text.length / 2));
      for (const part of [text.slice(0, mid), text.slice(mid)]) {
        send("content_block_delta", { type: "content_block_delta", index: 0, delta: { type: "text_delta", text: part } });
      }
      if ("delayMs" in step && step.delayMs) await new Promise((r) => setTimeout(r, step.delayMs));
      send("content_block_stop", { type: "content_block_stop", index: 0 });
      send("message_delta", { type: "message_delta", delta: { stop_reason: "end_turn", stop_sequence: null }, usage: { output_tokens: 5 } });
    }
    send("message_stop", { type: "message_stop" });
    res.end();
  }
}
