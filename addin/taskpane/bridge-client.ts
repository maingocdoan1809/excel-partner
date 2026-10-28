import {
  PROTOCOL_VERSION,
  MAX_SKILL_UPLOAD_BYTES,
  decodeServerMessage,
  encodeMessage,
  type ProviderEvent,
  type ProviderId,
  type ToolCapabilities,
  type BrowserFrameMessage,
  type ProviderStatusMessage,
  type SelectionSnapshot,
  type ServerMessage,
  type ToolRequestMessage,
  type ToolResult,
  type RunSkillScriptInput,
  UserInputRequestSchema,
  CANCELLED_USER_INPUT,
  type UserInputRequest,
  type UserInputResponse,
} from "../../protocol/src/index.js";
import { executeTool } from "./executor/index.js";
import { authorizeSkillScript, describeQaApproval, executeWithPermission, type PermissionMode, type QaApproval } from "./permissions.js";
import type { ProviderLoginStatus } from "../../protocol/src/provider-login.js";

export type BridgeState = "idle" | "connecting" | "open" | "closed";
export interface WorkspaceRoot { path: string; name: string; isDefault: boolean; removable: boolean }
export interface WorkspaceEntry { path: string; name: string; type: "folder" | "file"; size: number | null; supported: boolean }
export interface WorkspaceListing { path: string; entries: WorkspaceEntry[]; truncated: boolean }
export interface UploadedAttachment { id?: string; name: string; kind: "workbook" | "image" | "file"; path: string }
export interface SkillTemplate { name: string; kind: "workbook" | "image" | "text"; size: number; path: string }
export interface AvailableSkill { name: string; description: string }
export interface InstalledSkill extends AvailableSkill { removable: boolean; enabled: boolean }
export interface SkillTextFile { file: string; content: string; revision: string }
export interface SavedReference { id: string; name: string; path: string; description: string; available: boolean }
export interface ModelCatalog { models: Array<{ id: string; name: string }>; defaultModelId: string | null }

export interface BridgeClientHandlers {
  onUserInput?(request: UserInputRequest, signal: AbortSignal, requestId: string): Promise<UserInputResponse>;
  onState(state: BridgeState, detail?: string): void;
  onHello(info: { sessionToken: string; serverTime: string; addinBuildId?: string }): void;
  onProviderStatus(message: ProviderStatusMessage): void;
  onProviderEvent(event: ProviderEvent): void;
  /** A picture of the assistant's browser; shown to the person only. */
  onBrowserFrame?(frame: BrowserFrameMessage): void;
  onError(message: string): void;
  /** A tool request from the bridge finished running in this workbook. */
  onToolExecuted?(tool: string, args: unknown, result: ToolResult, context: { callId: string; requestId: string }): void;
}

const MAX_BACKOFF_MS = 15_000;

/**
 * Owns the task pane side of the bridge protocol: handshake, reconnect, tool
 * execution and outbound messages. It never talks to a CLI directly.
 */
export class BridgeClient {
  private ws: WebSocket | null = null;
  private reconnectAttempt = 0;
  private reconnectTimer: number | null = null;
  private closedByUser = false;
  private workbookConnectionId = randomId();
  private sessionToken: string | null = null;
  private permissionMode: PermissionMode = "ask";
  private readonly questions = new Map<string, { requestId: string; controller: AbortController }>();

  constructor(
    private readonly handlers: BridgeClientHandlers,
    private readonly url: string = defaultUrl(),
  ) {}

  get connected(): boolean {
    return this.ws?.readyState === WebSocket.OPEN;
  }

  setPermissionMode(mode: PermissionMode): void {
    this.permissionMode = mode;
  }

  connect(): void {
    this.closedByUser = false;
    this.workbookConnectionId = randomId();
    this.setState("connecting");
    try {
      this.ws = new WebSocket(this.url);
    } catch (error) {
      this.handlers.onError(`cannot open bridge socket: ${String(error)}`);
      this.scheduleReconnect();
      return;
    }

    this.ws.onopen = () => {
      this.reconnectAttempt = 0;
      this.send({
        type: "hello",
        protocolVersion: PROTOCOL_VERSION,
        workbookConnectionId: this.workbookConnectionId,
        clientInfo: {
          platform: Office.context?.platform ? String(Office.context.platform) : null,
          locale: Office.context?.displayLanguage ?? null,
        },
      });
    };

    this.ws.onmessage = (event) => void this.onMessage(String(event.data));
    this.ws.onerror = () => {
      /* handled by onclose */
    };
    this.ws.onclose = () => {
      this.closeQuestions();
      this.setState("closed");
      if (!this.closedByUser) this.scheduleReconnect();
    };
  }

  disconnect(): void {
    this.closeQuestions();
    this.closedByUser = true;
    if (this.reconnectTimer !== null) window.clearTimeout(this.reconnectTimer);
    this.ws?.close(1000, "task_pane_closed");
    this.ws = null;
    this.setState("closed");
  }

  setProvider(provider: ProviderId): void {
    this.send({ type: "set_provider", provider });
  }

  sendPrompt(requestId: string, prompt: string, selection: SelectionSnapshot | null, context?: string, attachmentIds?: string[], model?: string | null, referencePaths?: string[], effort: "low" | "medium" | "high" = "medium", chatSessionId?: string, resumeSessionId?: string, capabilities?: ToolCapabilities, command?: boolean): void {
    this.send({ type: "prompt", requestId, prompt, selection, chatSessionId, resumeSessionId, capabilities, command, model: model ?? null, effort, ...(context ? { context } : {}), ...(attachmentIds?.length ? { attachmentIds } : {}), ...(referencePaths?.length ? { referencePaths } : {}) });
  }

  async listCommands(provider: ProviderId): Promise<import("../../protocol/src/commands.js").ProviderCommand[]> {
    const result = await this.sessionFetch<{ commands: import("../../protocol/src/commands.js").ProviderCommand[] }>(`/session/commands?provider=${encodeURIComponent(provider)}`);
    return result.commands;
  }

  async usageSummaries(sessionId?: string): Promise<import("../../protocol/src/index.js").UsageSummary[]> {
    const data = await this.sessionFetch<{ sessions: import("../../protocol/src/index.js").UsageSummary[] }>(
      `/session/usage${sessionId ? `?sessionId=${encodeURIComponent(sessionId)}` : ""}`);
    return data.sessions;
  }

  async uploadAttachment(file: File): Promise<UploadedAttachment> {
    if (!this.connected || !this.sessionToken) throw new Error("Bridge is not connected. Try again after reconnecting.");
    if (file.size > 30 * 1024 * 1024) throw new Error("File exceeds the 30 MB limit.");
    const response = await fetch("/session/attachment", {
      method: "POST",
      headers: { authorization: `Bearer ${this.sessionToken}`, "x-file-name": encodeURIComponent(file.name) },
      body: file,
    });
    const data = await response.json() as { id?: string; name?: string; kind?: "workbook" | "image" | "file"; path?: string; error?: string };
    if (!response.ok || !data.id || !data.name || !data.kind || !data.path) throw new Error(data.error ?? "Could not attach the file");
    return { id: data.id, name: data.name, kind: data.kind, path: data.path };
  }

  async listWorkspaces(documentUrl: string, workbookName: string): Promise<{ roots: WorkspaceRoot[]; defaultError: string | null }> {
    const query = `?documentUrl=${encodeURIComponent(documentUrl)}&workbookName=${encodeURIComponent(workbookName)}`;
    return this.sessionFetch(`/session/workspaces${query}`);
  }

  async removeWorkspace(path: string): Promise<void> {
    await this.sessionFetch("/session/workspaces", { method: "DELETE", body: JSON.stringify({ path }) });
  }

  async listWorkspaceEntries(path: string): Promise<WorkspaceListing> {
    return this.sessionFetch(`/session/workspaces/entries?path=${encodeURIComponent(path)}`);
  }

  async attachWorkspaceFile(path: string): Promise<UploadedAttachment> {
    return this.sessionFetch("/session/workspaces/attach", { method: "POST", body: JSON.stringify({ path }) });
  }

  async listSavedReferences(): Promise<SavedReference[]> {
    const data = await this.sessionFetch<{ references: SavedReference[] }>("/session/references");
    return data.references;
  }

  async saveReference(input: { id?: string; path?: string; name: string; description: string }): Promise<void> {
    await this.sessionFetch("/session/references", { method: "POST", body: JSON.stringify(input) });
  }

  async removeReference(id: string): Promise<void> {
    await this.sessionFetch("/session/references", { method: "DELETE", body: JSON.stringify({ id }) });
  }

  async listTemplateSkills(): Promise<AvailableSkill[]> {
    const data = await this.sessionFetch<{ skills: AvailableSkill[] }>("/session/skill-templates");
    return data.skills;
  }

  async listModels(provider: ProviderId): Promise<ModelCatalog> {
    return this.sessionFetch(`/session/models?provider=${encodeURIComponent(provider)}`);
  }

  async listSkills(): Promise<InstalledSkill[]> {
    const data = await this.sessionFetch<{ skills: InstalledSkill[] }>("/session/skills");
    return data.skills;
  }

  async uploadSkill(file: File): Promise<void> {
    if (file.size > MAX_SKILL_UPLOAD_BYTES) throw new Error("Skill uploads must be 5 MB or smaller.");
    await this.sessionFetch("/session/skills", { method: "POST", body: file, headers: { "x-file-name": encodeURIComponent(file.name) } });
  }

  async removeSkill(name: string): Promise<void> {
    await this.sessionFetch("/session/skills", { method: "DELETE", body: JSON.stringify({ name }) });
  }

  async setSkillEnabled(name: string, enabled: boolean): Promise<void> {
    await this.sessionFetch("/session/skills", { method: "PATCH", body: JSON.stringify({ name, enabled }) });
  }

  async skillTextFiles(name: string): Promise<string[]> {
    return (await this.sessionFetch<{ files: string[] }>(`/session/skills?name=${encodeURIComponent(name)}`)).files;
  }

  async systemPrompt(): Promise<SkillTextFile> {
    return { file: "System prompt", ...await this.sessionFetch<{ content: string; revision: string }>("/session/system-prompt") };
  }

  async saveSystemPrompt(content: string | null, revision: string): Promise<SkillTextFile> {
    return { file: "System prompt", ...await this.sessionFetch<{ content: string; revision: string }>("/session/system-prompt", { method: "PUT", body: JSON.stringify({ content, revision }) }) };
  }

  async readSkillText(name: string, file: string): Promise<SkillTextFile> {
    return (await this.sessionFetch<{ document: SkillTextFile }>(`/session/skills?name=${encodeURIComponent(name)}&file=${encodeURIComponent(file)}`)).document;
  }

  async saveSkillText(name: string, document: SkillTextFile): Promise<SkillTextFile> {
    return (await this.sessionFetch<{ document: SkillTextFile }>("/session/skills", { method: "PUT", body: JSON.stringify({ name, ...document }) })).document;
  }

  async providerLogin(method: "GET" | "POST" | "DELETE" = "GET"): Promise<ProviderLoginStatus> {
    return this.sessionFetch("/session/provider-login", { method });
  }

  async listSkillTemplates(skill: string): Promise<SkillTemplate[]> {
    const data = await this.sessionFetch<{ templates: SkillTemplate[] }>(`/session/skill-templates?skill=${encodeURIComponent(skill)}`);
    return data.templates;
  }

  async addSkillTemplate(skill: string, file: File): Promise<SkillTemplate> {
    const data = await this.sessionFetch<{ template: SkillTemplate }>("/session/skill-templates", {
      method: "POST", body: file, headers: { "x-skill-name": encodeURIComponent(skill), "x-file-name": encodeURIComponent(file.name) },
    });
    return data.template;
  }

  async deleteSkillTemplate(skill: string, name: string): Promise<void> {
    await this.sessionFetch("/session/skill-templates", { method: "DELETE", body: JSON.stringify({ skill, name }) });
  }

  private async sessionFetch<T>(url: string, options: RequestInit = {}): Promise<T> {
    if (!this.connected || !this.sessionToken) throw new Error("Bridge is not connected. Try again after reconnecting.");
    const response = await fetch(url, {
      ...options,
      headers: { authorization: `Bearer ${this.sessionToken}`, ...(typeof options.body === "string" ? { "content-type": "application/json" } : {}), ...options.headers },
      cache: "no-store",
    });
    const data = await response.json() as T & { error?: string; code?: string };
    if (!response.ok) throw Object.assign(new Error(data.error ?? `Workspace request failed (HTTP ${response.status})`), { code: data.code });
    return data;
  }

  cancel(requestId: string): void {
    this.closeQuestions(requestId);
    this.send({ type: "cancel", requestId });
  }

  queuePrompt(requestId: string, id: string, prompt: string, referencePaths: string[], attachmentIds: string[]): void {
    this.send({ type: "followup", requestId, id, prompt, referencePaths, attachmentIds });
  }

  /* ----------------------------------------------------------------------- */

  private async onMessage(raw: string): Promise<void> {
    let message: ServerMessage;
    try {
      message = decodeServerMessage(raw);
    } catch (error) {
      this.handlers.onError(`bad message from bridge: ${String(error)}`);
      return;
    }

    switch (message.type) {
      case "hello_ok":
        this.sessionToken = message.sessionToken;
        this.setState("open");
        this.handlers.onHello({
          sessionToken: message.sessionToken,
          serverTime: message.serverTime,
          addinBuildId: message.addinBuildId,
        });
        return;
      case "hello_error":
        this.handlers.onError(`bridge rejected connection: ${message.message}`);
        this.closedByUser = true;
        this.ws?.close();
        return;
      case "provider_status":
        this.handlers.onProviderStatus(message);
        return;
      case "tool_request":
        await this.handleToolRequest(message);
        return;
      case "user_input_closed":
        this.questions.get(message.callId)?.controller.abort();
        return;
      case "ping":
        this.send({ type: "pong", nonce: message.nonce });
        return;
      case "browser_frame":
        this.handlers.onBrowserFrame?.(message);
        return;
      default:
        if (message.type === "turn_finished") this.closeQuestions(message.requestId);
        this.handlers.onProviderEvent(message);
        return;
    }
  }

  private async handleToolRequest(message: ToolRequestMessage): Promise<void> {
    if (message.tool === "__request_user_input") {
      const socket = this.ws;
      const controller = new AbortController();
      this.questions.set(message.callId, { requestId: message.requestId, controller });
      const timer = window.setTimeout(() => controller.abort(), message.timeoutMs);
      let response = CANCELLED_USER_INPUT;
      try {
        response = await this.handlers.onUserInput?.(UserInputRequestSchema.parse(message.args), controller.signal, message.requestId) ?? response;
      } catch { /* Invalid, disconnected or dismissed questions are cancelled, never auto-answered. */ }
      finally {
        window.clearTimeout(timer);
        this.questions.delete(message.callId);
      }
      if (socket === this.ws) this.send({ type: "tool_result", callId: message.callId,
        result: { ok: true, data: controller.signal.aborted ? CANCELLED_USER_INPUT : response } });
      return;
    }
    const result = message.tool === "__authorize_provider_tool"
      ? { ok: true as const, data: { approved: this.permissionMode === "auto" || await this.askPermission(
        "Allow OpenCode tool?", "OpenCode requests permission to run this tool on your computer.", "Allow tool", message.timeoutMs,
        JSON.stringify(message.args, null, 2).slice(0, 8000)) } }
      : message.tool === "__authorize_skill_edit"
      ? { ok: true as const, data: { approved: this.permissionMode === "auto" || await this.askPermission("Update installed skill?",
        `Save changes to /${(message.args as { name: string }).name}/${(message.args as { file: string }).file}?`, "Save skill", message.timeoutMs,
        (message.args as { content: string }).content.slice(0, 8000)) } }
      : message.tool === "__authorize_browser"
      ? { ok: true as const, data: { approved: this.permissionMode === "auto" || await this.askPermission("Allow browser action?",
        "Playwright will perform this action in a separate browser on this computer.", "Allow browser action", message.timeoutMs,
        JSON.stringify(message.args, null, 2).slice(0, 8000)) } }
      : message.tool === "__authorize_python"
      ? this.permissionMode === "auto" ? { ok: true as const, data: { approved: true } }
        : { ok: true as const, data: { approved: await this.askPermission("Run Python on this computer?",
          "The assistant will run Python on your computer with your user permissions, including access to files and the network. Selected file copies and generated files use a temporary working folder.",
          "Run Python", message.timeoutMs, (message.args as { code: string }).code.slice(0, 8000)) } }
      : message.tool === "__authorize_skill_script"
      ? await authorizeSkillScript(this.permissionMode, message.args as RunSkillScriptInput,
        (detail, target) => this.askPermission("Allow skill script?", detail, "Run script", message.timeoutMs, target))
      : message.tool === "__authorize_qa"
      ? await this.authorizeQa(message.args as QaApproval, message.timeoutMs)
      : message.tool === "__qa_apply_results"
      ? await this.applyQaResults(message.args as { sheet: string; rows: unknown[] }, message)
      : message.tool === "__authorize_reference_path"
      ? this.permissionMode === "auto"
        ? { ok: true as const, data: { approved: true } }
        : await this.askReferencePermission(message.args as { path: string; resolvedPath?: string }, message.timeoutMs)
      : await executeWithPermission(
        message.tool,
        message.args,
        this.permissionMode,
        executeTool,
        (detail) => this.askPermission("Allow workbook change?", detail, "Allow change", message.timeoutMs),
      );
    this.send({ type: "tool_result", callId: message.callId, result });
    if (!message.tool.startsWith("__authorize_")) this.handlers.onToolExecuted?.(message.tool, message.args, result, { callId: message.callId, requestId: message.requestId });
  }

  private closeQuestions(requestId?: string): void {
    for (const question of this.questions.values()) {
      if (!requestId || question.requestId === requestId) question.controller.abort();
    }
  }

  /** Save/run/open-browser requests from the QA harness. Draft runs always ask, even in Auto mode. */
  private async authorizeQa(args: QaApproval, timeoutMs: number): Promise<ToolResult> {
    if (this.permissionMode === "auto" && !args.forceAsk) return { ok: true, data: { approved: true } };
    const { title, detail, confirm, target } = describeQaApproval(args);
    return { ok: true, data: { approved: await this.askPermission(title, detail, confirm, timeoutMs, target) } };
  }

  private async applyQaResults(args: { sheet: string; rows: unknown[] }, message: ToolRequestMessage): Promise<ToolResult> {
    if (this.permissionMode !== "auto") {
      const approved = await this.askPermission("Write QA results to the workbook?",
        `The assistant will update ${args.rows.length} test case row(s) in sheet '${args.sheet}': automation status, script path and last run result. Other cells are not touched.`,
        "Write results", message.timeoutMs);
      if (!approved) return { ok: false, error: { code: "cancelled", message: "The user declined writing QA results." } };
    }
    return executeTool("__qa_apply_results", args);
  }

  private askReferencePermission(args: { path: string; resolvedPath?: string }, timeoutMs: number): Promise<ToolResult> {
    return this.askPermission(
      "Allow read access?",
      "Allow the assistant to read this file for this chat. This grants read access to this file only.",
      "Allow read",
      timeoutMs,
      args.resolvedPath ?? args.path,
    ).then((approved) => ({ ok: true, data: { approved } }));
  }

  private askPermission(titleText: string, detail: string, allowText: string, timeoutMs: number, path?: string): Promise<boolean> {
    return new Promise((resolve) => {
      const overlay = document.createElement("div");
      overlay.className = "permission-overlay";
      overlay.setAttribute("role", "dialog");
      overlay.setAttribute("aria-modal", "true");
      overlay.setAttribute("aria-label", titleText);
      const box = document.createElement("div");
      box.className = "permission-dialog";
      const title = document.createElement("h2");
      title.textContent = titleText;
      const explanation = document.createElement("p");
      explanation.textContent = detail;
      const actions = document.createElement("div");
      actions.className = "permission-actions";
      const deny = document.createElement("button");
      deny.type = "button";
      deny.textContent = "Deny";
      const allow = document.createElement("button");
      allow.type = "button";
      allow.textContent = allowText;
      actions.append(deny, allow);
      box.append(title, explanation);
      if (path) {
        const target = document.createElement("code");
        target.textContent = path;
        box.append(target);
      }
      box.append(actions);
      overlay.append(box);
      document.body.append(overlay);
      const finish = (approved: boolean) => {
        window.clearTimeout(timer);
        overlay.remove();
        resolve(approved);
      };
      const timer = window.setTimeout(() => finish(false), Math.max(1000, timeoutMs - 1000));
      deny.addEventListener("click", () => finish(false), { once: true });
      allow.addEventListener("click", () => finish(true), { once: true });
      deny.focus();
    });
  }

  private send(message: Parameters<typeof encodeMessage>[0]): void {
    if (!this.connected) return;
    this.ws!.send(encodeMessage(message));
  }

  private setState(state: BridgeState, detail?: string): void {
    this.handlers.onState(state, detail);
  }

  private scheduleReconnect(): void {
    const delay = Math.min(500 * 2 ** this.reconnectAttempt, MAX_BACKOFF_MS);
    this.reconnectAttempt++;
    this.reconnectTimer = window.setTimeout(() => this.connect(), delay);
  }
}

function defaultUrl(): string {
  const scheme = window.location.protocol === "https:" ? "wss" : "ws";
  return `${scheme}://${window.location.host}/ws`;
}

function randomId(): string {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return [...bytes].map((b) => b.toString(16).padStart(2, "0")).join("");
}
