import { randomBytes, randomUUID } from "node:crypto";
import type { WebSocket } from "ws";
import type { SessionAttachment } from "./attachments.js";
import {
  encodeMessage,
  type ClientInfo,
  type ProviderId,
  type ServerMessage,
  type ToolResult,
  type ToolError,
} from "../../protocol/src/index.js";

interface PendingToolCall {
  callId: string;
  tool: string;
  resolve: (result: ToolResult) => void;
  timer: NodeJS.Timeout;
}

export interface WorkbookSessionState {
  provider: ProviderId | null;
  providerSessionId: string | null;
  /** requestId of the turn currently running, if any. */
  activeRequestId: string | null;
  lastActivityAt: number;
}

/**
 * One workbook <-> one chat session <-> one provider at a time.
 *
 * The session owns the WebSocket to the task pane and is the only component
 * allowed to talk to it. Tool calls originating from a provider/MCP process are
 * turned into `tool_request` messages and resolved when the pane answers with
 * `tool_result`.
 */
export class WorkbookSession {
  readonly id: string;
  readonly sessionToken: string;
  readonly workbookConnectionId: string;
  readonly createdAt: number;
  readonly clientInfo: ClientInfo | null;

  socket: WebSocket | null = null;

  provider: ProviderId | null = null;
  providerSessionId: string | null = null;
  providerNativeSessionId: string | null = null;
  providerChatSessionId: string | null = null;
  providerModel: string | null = null;
  providerEffort: "low" | "medium" | "high" = "medium";
  activeRequestId: string | null = null;
  turnController: AbortController | null = null;
  readonly scriptControllers = new Set<AbortController>();
  lastActivityAt: number;
  /** Extra read-only paths approved in this connected workbook session. */
  readonly referenceGrants = new Set<string>();
  /** Folder containing the workbook currently open in this pane, when its local path is available. */
  defaultWorkspace: string | null = null;
  /** File bytes explicitly selected in the pane; never written to disk. */
  readonly attachments = new Map<string, SessionAttachment>();

  /** Serialises write-ish tool calls so the pane applies them in order. */
  private toolChain: Promise<unknown> = Promise.resolve();
  private readonly pending = new Map<string, PendingToolCall>();

  constructor(workbookConnectionId: string, clientInfo: ClientInfo | null = null) {
    this.id = randomUUID();
    this.sessionToken = randomBytes(32).toString("base64url");
    this.workbookConnectionId = workbookConnectionId;
    this.clientInfo = clientInfo;
    this.createdAt = Date.now();
    this.lastActivityAt = this.createdAt;
  }

  touch(): void {
    this.lastActivityAt = Date.now();
  }

  get isBusy(): boolean {
    return this.activeRequestId !== null;
  }

  /** Request ids whose `turn_finished` has been sent (bounded, for dedupe). */
  private readonly finished = new Set<string>();

  hasFinished(requestId: string): boolean {
    return this.finished.has(requestId);
  }

  /** Returns false when `turn_finished` was already sent for this request. */
  markFinished(requestId: string): boolean {
    if (this.finished.has(requestId)) return false;
    this.finished.add(requestId);
    if (this.finished.size > 200) {
      const oldest = this.finished.values().next().value as string;
      this.finished.delete(oldest);
    }
    return true;
  }

  beginTurn(requestId: string): AbortController {
    const controller = new AbortController();
    this.activeRequestId = requestId;
    this.turnController = controller;
    return controller;
  }

  endTurn(requestId: string): void {
    if (this.activeRequestId === requestId) {
      this.activeRequestId = null;
      this.turnController = null;
    }
  }

  abortTurn(): void {
    this.turnController?.abort();
    for (const controller of this.scriptControllers) controller.abort();
  }

  get isConnected(): boolean {
    return this.socket !== null && this.socket.readyState === 1;
  }

  attachSocket(socket: WebSocket): void {
    this.socket = socket;
    this.touch();
  }

  detachSocket(socket: WebSocket): void {
    if (this.socket === socket) {
      this.abortTurn();
      this.socket = null;
      this.referenceGrants.clear();
      this.defaultWorkspace = null;
      this.attachments.clear();
    }
    this.rejectAllPending({
      code: "disconnected",
      message: "workbook/task pane disconnected before the tool call completed",
    });
    this.activeRequestId = null;
  }

  send(message: ServerMessage): boolean {
    if (!this.socket || this.socket.readyState !== 1) return false;
    this.socket.send(encodeMessage(message));
    return true;
  }

  /**
   * Send a tool request to the pane and await its result. Calls are serialised
   * per session so ordering of writes is preserved.
   */
  requestTool(tool: string, args: unknown, timeoutMs: number): Promise<ToolResult> {
    const run = () => this.dispatchTool(tool, args, timeoutMs);
    const result = this.toolChain.then(run, run);
    this.toolChain = result.catch(() => undefined);
    return result;
  }

  private dispatchTool(tool: string, args: unknown, timeoutMs: number): Promise<ToolResult> {
    const callId = randomUUID();
    if (!this.isConnected) {
      return Promise.resolve({
        ok: false,
        error: { code: "disconnected", message: "task pane is not connected" },
      });
    }

    return new Promise<ToolResult>((resolve) => {
      const timer = setTimeout(() => {
        this.pending.delete(callId);
        resolve({
          ok: false,
          error: { code: "timeout", message: `tool ${tool} timed out after ${timeoutMs}ms` },
        });
      }, timeoutMs);
      timer.unref?.();

      this.pending.set(callId, { callId, tool, resolve, timer });

      const sent = this.send({
        type: "tool_request",
        requestId: this.activeRequestId ?? "no-turn",
        callId,
        tool,
        args,
        timeoutMs,
      });

      if (!sent) {
        clearTimeout(timer);
        this.pending.delete(callId);
        resolve({
          ok: false,
          error: { code: "disconnected", message: "failed to send tool_request to task pane" },
        });
      }
    });
  }

  resolveTool(callId: string, result: ToolResult): boolean {
    const entry = this.pending.get(callId);
    if (!entry) return false;
    clearTimeout(entry.timer);
    this.pending.delete(callId);
    entry.resolve(result);
    return true;
  }

  rejectAllPending(error: ToolError): void {
    for (const entry of this.pending.values()) {
      clearTimeout(entry.timer);
      entry.resolve({ ok: false, error });
    }
    this.pending.clear();
  }

  get pendingToolCount(): number {
    return this.pending.size;
  }

  toJSON(): Record<string, unknown> {
    return {
      id: this.id,
      workbookConnectionId: this.workbookConnectionId,
      provider: this.provider,
      connected: this.isConnected,
      activeRequestId: this.activeRequestId,
      pendingTools: this.pendingToolCount,
      createdAt: this.createdAt,
      lastActivityAt: this.lastActivityAt,
    };
  }
}
