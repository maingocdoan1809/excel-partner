import { randomUUID } from "node:crypto";
import type {
  ProviderCapabilities,
  ProviderConnectionStatus,
  ProviderDetection,
  ProviderEvent,
  ProviderId,
  ToolResult,
} from "../../protocol/src/index.js";
import type { McpEndpoint, ProviderAdapter, SendTurnInput, StartSessionContext } from "./types.js";

export interface MockToolCall {
  tool: string;
  args: unknown;
}

export interface MockAdapterOptions {
  /** Provider id this adapter answers to in the registry. */
  id?: ProviderId;
  /** Tool calls the adapter runs, in order, on every turn. */
  toolCalls?: MockToolCall[];
  /** Text streamed back after the tool calls finish. */
  reply?: string;
  delayMs?: number;
}

interface MockSession {
  mcp: McpEndpoint;
  controller: AbortController;
}

/**
 * In-process adapter used by contract/integration tests and `npm run dev`.
 * It exercises the exact production path: the "CLI" reaches the bridge through
 * the MCP endpoint, which routes tool calls to the task pane over WebSocket.
 */
export class MockAdapter implements ProviderAdapter {
  readonly id: ProviderId;
  readonly capabilities: ProviderCapabilities = {
    streaming: true,
    cancel: true,
    sessionResume: true,
    mcp: true,
  };

  private readonly sessions = new Map<string, MockSession>();
  private readonly options: Required<MockAdapterOptions>;

  constructor(options: MockAdapterOptions = {}) {
    this.options = {
      id: options.id ?? "claude",
      toolCalls: options.toolCalls ?? [],
      reply: options.reply ?? "Done.",
      delayMs: options.delayMs ?? 0,
    };
    this.id = this.options.id;
  }

  async detect(): Promise<ProviderDetection> {
    return {
      provider: this.id,
      installed: true,
      version: "mock-0.1.0",
      status: "ready",
      hint: null,
      capabilities: this.capabilities,
    };
  }

  async connectionStatus(): Promise<ProviderConnectionStatus> {
    return "ready";
  }

  async startSession(ctx: StartSessionContext): Promise<string> {
    const id = `mock-${randomUUID()}`;
    this.sessions.set(id, { mcp: ctx.mcp, controller: new AbortController() });
    return id;
  }

  async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    const session = this.sessions.get(input.providerSessionId);
    const requestId = input.requestId;

    yield { type: "turn_started", requestId };

    if (!session) {
      yield { type: "error", requestId, code: "no_session", message: "mock session not found" };
      yield { type: "turn_finished", requestId, reason: "error" };
      return;
    }

    for (const call of this.options.toolCalls) {
      if (input.signal.aborted) {
        yield { type: "turn_finished", requestId, reason: "cancelled" };
        return;
      }
      const callId = randomUUID();
      yield { type: "tool_started", requestId, callId, tool: call.tool, args: call.args };
      const result = await callTool(session.mcp, call.tool, call.args);
      yield {
        type: "tool_finished",
        requestId,
        callId,
        tool: call.tool,
        ok: result.ok,
        error: result.ok ? null : result.error.message,
      };
      if (!result.ok) {
        yield { type: "error", requestId, code: result.error.code, message: result.error.message };
        yield { type: "turn_finished", requestId, reason: "error" };
        return;
      }
    }

    if (this.options.delayMs > 0) {
      await new Promise((resolve) => setTimeout(resolve, this.options.delayMs));
    }

    if (input.signal.aborted) {
      yield { type: "turn_finished", requestId, reason: "cancelled" };
      return;
    }

    yield { type: "message_delta", requestId, text: this.options.reply };
    yield { type: "message_completed", requestId, text: this.options.reply };
    yield { type: "turn_finished", requestId, reason: "completed" };
  }

  async cancelTurn(providerSessionId: string): Promise<void> {
    this.sessions.get(providerSessionId)?.controller.abort();
  }

  async closeSession(providerSessionId: string): Promise<void> {
    this.sessions.delete(providerSessionId);
  }

  async dispose(): Promise<void> {
    this.sessions.clear();
  }
}

async function callTool(endpoint: McpEndpoint, tool: string, args: unknown): Promise<ToolResult> {
  const res = await fetch(`${endpoint.bridgeUrl}/internal/tool`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${endpoint.token}`,
    },
    body: JSON.stringify({ sessionId: endpoint.sessionId, tool, args }),
  });
  if (!res.ok) {
    return {
      ok: false,
      error: { code: "internal_error", message: `bridge responded ${res.status}` },
    };
  }
  return (await res.json()) as ToolResult;
}
