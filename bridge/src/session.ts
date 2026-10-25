import { randomBytes, randomUUID } from "node:crypto";
import type { WebSocket } from "ws";
import type { ProviderSessionState } from "../../providers/src/session-manager.js";
import type { SessionAttachment } from "./attachments.js";
import {
  encodeMessage,
  type ClientInfo,
  type ProviderId,
  type ServerMessage,
  type ToolResult,
  type ToolError,
  UserInputRequestSchema,
  validateUserInputResponse,
  CANCELLED_USER_INPUT,
  type UserInputRequest,
  type UserInputResponse,
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
export class WorkbookSession implements ProviderSessionState {
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
  systemPromptRevision: string | null = null;
  capabilities = { webSearch: false, playwright: false };
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
  private readonly followups = new Map<string, { requestId: string; text: string }>();

  queueFollowup(requestId: string, id: string, text: string): void {
    if (this.followups.has(id)) return;
    if (this.followups.size >= 20 || [...this.followups.values()].reduce((total, item) => total + item.text.length, 0) + text.length > 64_000) throw new Error("Too many pending messages. Wait for the assistant to receive them.");
    this.followups.set(id, { requestId, text });
  }

  takeFollowups(requestId: string, state: "delivered" | "next_turn"): Array<{ id: string; text: string }> {
    const items: Array<{ id: string; text: string }> = [];
    for (const [id, item] of this.followups) {
      if (item.requestId !== requestId) continue;
      this.followups.delete(id);
      items.push({ id, text: item.text });
      this.send({ type: "followup_status", requestId, id, state });
    }
    return items;
  }

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
      for (const entry of this.pending.values()) {
        if (entry.tool === "__request_user_input") entry.resolve({ ok: false, error: { code: "cancelled", message: "The turn ended." } });
      }
      this.activeRequestId = null;
      this.turnController = null;
    }
  }

  abortTurn(): void {
    this.turnController?.abort();
    for (const controller of this.scriptControllers) controller.abort();
  }

  async requestUserInput(request: UserInputRequest, signal?: AbortSignal): Promise<UserInputResponse> {
    const parsed = UserInputRequestSchema.parse(request);
    const turnSignal = this.turnController?.signal;
    if (!this.activeRequestId || !turnSignal || turnSignal.aborted || signal?.aborted) return CANCELLED_USER_INPUT;
    // Native questions must not queue behind the very tool that is asking them.
    const controller = new AbortController();
    const abort = () => controller.abort();
    turnSignal.addEventListener("abort", abort, { once: true });
    signal?.addEventListener("abort", abort, { once: true });
    try {
      const result = await this.dispatchTool("__request_user_input", parsed, 600_000, this.activeRequestId, controller.signal);
      return result.ok ? validateUserInputResponse(parsed, result.data) : CANCELLED_USER_INPUT;
    } finally {
      turnSignal.removeEventListener("abort", abort);
      signal?.removeEventListener("abort", abort);
    }
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
    const requestId = this.activeRequestId;
    const signal = this.turnController?.signal;
    const run = (): Promise<ToolResult> => signal?.aborted || this.activeRequestId !== requestId
      ? Promise.resolve({ ok: false, error: { code: "cancelled", message: "The request ended before this tool could run." } })
      : this.dispatchTool(tool, args, timeoutMs, requestId);
    const result = this.toolChain.then(run, run);
    this.toolChain = result.catch(() => undefined);
    return result;
  }

  private dispatchTool(tool: string, args: unknown, timeoutMs: number, requestId: string | null, signal?: AbortSignal): Promise<ToolResult> {
    const callId = randomUUID();
    if (!this.isConnected) {
      return Promise.resolve({
        ok: false,
        error: { code: "disconnected", message: "task pane is not connected" },
      });
    }

    return new Promise<ToolResult>((resolve) => {
      let finished = false;
      const finish = (result: ToolResult) => {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        this.pending.delete(callId);
        signal?.removeEventListener("abort", abort);
        if (tool === "__request_user_input") this.send({ type: "user_input_closed", callId });
        resolve(result);
      };
      const abort = () => finish({ ok: false, error: { code: "cancelled", message: "Question cancelled." } });
      const timer = setTimeout(() => {
        finish({
          ok: false,
          error: { code: "timeout", message: `tool ${tool} timed out after ${timeoutMs}ms` },
        });
      }, timeoutMs);
      timer.unref?.();

      this.pending.set(callId, { callId, tool, resolve: finish, timer });
      signal?.addEventListener("abort", abort, { once: true });
      if (signal?.aborted) { abort(); return; }

      const sent = this.send({
        type: "tool_request",
        requestId: requestId ?? "no-turn",
        callId,
        tool,
        args,
        timeoutMs,
      });

      if (!sent) {
        clearTimeout(timer);
        this.pending.delete(callId);
        finish({
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
      capabilities: { ...this.capabilities },
      connected: this.isConnected,
      activeRequestId: this.activeRequestId,
      pendingTools: this.pendingToolCount,
      createdAt: this.createdAt,
      lastActivityAt: this.lastActivityAt,
    };
  }
}
