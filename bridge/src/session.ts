import type { ExternalMcpCatalog } from "../../runtime/src/mcp-catalog.js";
import { randomBytes, randomUUID } from "node:crypto";
import type { WebSocket } from "ws";
import type { ProviderSessionState } from "../../providers/src/session-manager.js";
import type { SessionAttachment } from "./attachments.js";
import {
  encodeMessage,
  DEFAULT_HOST,
  HOST_IDS,
  isHostId,
  type ClientInfoInput,
  type HostId,
  type ProviderId,
  type ServerMessage,
  type ToolResult,
  type ToolError,
  UserInputRequestSchema,
  validateUserInputResponse,
  CANCELLED_USER_INPUT,
  WORKBOOK_NAME_REQUEST,
  type UserInputRequest,
  type UserInputResponse,
} from "../../protocol/src/index.js";

/** How long a pane has to say what its workbook is called before the name it last gave is used. */
export const WORKBOOK_NAME_TIMEOUT_MS = 2_000;

const ENDED_BEFORE_RUN = { ok: false, error: { code: "cancelled", message: "The request ended before this tool could run." } } as const satisfies ToolResult;

interface PendingToolCall {
  callId: string;
  tool: string;
  resolve: (result: ToolResult) => void;
  timer: NodeJS.Timeout;
}

const HOST_APP_NAMES: Record<HostId, string> = { excel: "Excel", word: "Word", powerpoint: "PowerPoint" };

/** Shown to the model when the pane drops mid-call, which usually means the Office app itself crashed on that call. */
export function disconnectedMidCall(host: HostId): string {
  const app = HOST_APP_NAMES[host];
  const progress = host === "excel" ? "one sheet or a smaller range" : host === "word" ? "one section or a smaller range" : "one slide or a smaller change";
  const reopened = host === "excel" ? "workbook" : host === "word" ? "document" : "presentation";
  return `${app}/task pane disconnected while this tool call was running; ${app} may have crashed on it. ` +
    `Do NOT repeat the same call with the same arguments. Tell the user, wait until the ${reopened} is reopened, ` +
    `then retry only a smaller step (for example ${progress}) or a different approach.`;
}

/** The Excel wording; kept as the value Excel sessions use. */
export const DISCONNECTED_MID_CALL = disconnectedMidCall("excel");

/** Whether `message` is the "pane dropped mid-call" error of any host. */
export function isDisconnectedMidCall(message: string): boolean {
  return HOST_IDS.some((host) => message === disconnectedMidCall(host));
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
  /** Short name the assistant uses to point at this file (`f1`, `f2`, ...). Not a credential; see list_connected_files. */
  readonly fileId: string;
  readonly createdAt: number;
  readonly clientInfo: ClientInfoInput | null;
  /** The Office app this pane runs in, fixed at `hello`. Every tool gate reads it from here. Standalone panes are "excel". */
  readonly host: HostId;

  socket: WebSocket | null = null;

  provider: ProviderId | null = null;
  providerSessionId: string | null = null;
  providerNativeSessionId: string | null = null;
  providerChatSessionId: string | null = null;
  providerModel: string | null = null;
  systemPromptRevision: string | null = null;
  externalMcpTools: ExternalMcpCatalog = new Map();
  readonly enabledMcpServers = new Set<string>();
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
  /** Name of the workbook open in this pane, as Excel reports it; null until the pane says. */
  workbookName: string | null = null;
  /** Other files this chat may read, keyed by the other session's id, and the request (message) the user allowed it for. */
  readonly connectedReadGrants = new Map<string, string>();
  /**
   * What the current test run has done so far, keyed "sheet|row": results written to the sheet (with the columns they went
   * to) and regression scripts saved. finish_test_run checks the two against each other before the run may be called done.
   */
  readonly testRun = {
    recorded: new Map<string, { values: Record<string, string>; columns: Record<string, string> }>(),
    scripts: new Map<string, { caseId: string; outcome: "pass" | "fail_app_defect" }>(),
  };
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

  constructor(workbookConnectionId: string, clientInfo: ClientInfoInput | null = null, fileId = "f1") {
    this.id = randomUUID();
    this.fileId = fileId;
    this.sessionToken = randomBytes(32).toString("base64url");
    this.workbookConnectionId = workbookConnectionId;
    this.clientInfo = clientInfo;
    this.host = isHostId(clientInfo?.host) ? clientInfo.host : DEFAULT_HOST;
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
      this.connectedReadGrants.clear();
      this.defaultWorkspace = null;
      this.attachments.clear();
    }
    this.rejectAllPending({
      code: "disconnected",
      message: disconnectedMidCall(this.host),
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
    return this.enqueue(() => signal?.aborted || this.activeRequestId !== requestId
      ? Promise.resolve(ENDED_BEFORE_RUN)
      : this.dispatchTool(tool, args, timeoutMs, requestId));
  }

  /**
   * A read made for another workbook's chat. It waits behind this pane's own calls so Office.js never runs two at
   * once, but it belongs to no turn of this pane: this pane's own chat neither shows nor cancels it. `signal` is the
   * asker's: when it aborts, the read is dropped (or stops being waited for) so this pane's own calls are not held up.
   */
  requestPeerRead(tool: string, args: unknown, timeoutMs: number, signal?: AbortSignal): Promise<ToolResult> {
    return this.enqueue(() => signal?.aborted
      ? Promise.resolve(ENDED_BEFORE_RUN)
      : this.dispatchTool(tool, args, timeoutMs, null, signal));
  }

  /**
   * Ask the pane what its workbook is called now and remember the answer. Excel has no event for a rename (Save As), so the
   * name is read when someone needs it. This skips the tool queue on purpose: it changes nothing, and it must not wait
   * behind a long tool call. A pane that does not answer in time keeps the name it gave before. True when the name changed.
   */
  async refreshWorkbookName(timeoutMs = WORKBOOK_NAME_TIMEOUT_MS): Promise<boolean> {
    if (!this.isConnected) return false;
    const result = await this.dispatchTool(WORKBOOK_NAME_REQUEST, {}, timeoutMs, null);
    const name = result.ok ? (result.data as { name?: unknown } | null)?.name : undefined;
    if (typeof name !== "string" || !name.trim()) return false;
    const next = name.trim().slice(0, 255);
    const changed = next !== this.workbookName;
    this.workbookName = next;
    return changed;
  }

  /** Run `run` after every call queued before it, whether those succeeded or not. */
  private enqueue(run: () => Promise<ToolResult>): Promise<ToolResult> {
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
      const abort = () => finish({ ok: false, error: { code: "cancelled", message: tool === "__request_user_input" ? "Question cancelled." : ENDED_BEFORE_RUN.error.message } });
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
