import {
  PROTOCOL_VERSION,
  MAX_SKILL_UPLOAD_BYTES,
  decodeServerMessage,
  encodeMessage,
  type ProviderEvent,
  type ProviderId,
  type ProviderStatusMessage,
  type SelectionSnapshot,
  type ServerMessage,
  type ToolRequestMessage,
  type ToolResult,
  type RunSkillScriptInput,
} from "../../protocol/src/index.js";
import { executeTool } from "./executor/index.js";
import { authorizeSkillScript, executeWithPermission, type PermissionMode } from "./permissions.js";
import type { ProviderLoginStatus } from "../../protocol/src/provider-login.js";

export type BridgeState = "idle" | "connecting" | "open" | "closed";
export interface WorkspaceRoot { path: string; name: string; isDefault: boolean; removable: boolean }
export interface WorkspaceEntry { path: string; name: string; type: "folder" | "file"; size: number | null; supported: boolean }
export interface WorkspaceListing { path: string; entries: WorkspaceEntry[]; truncated: boolean }
export interface UploadedAttachment { id?: string; name: string; kind: "workbook" | "image" | "file"; path: string }
export interface SkillTemplate { name: string; kind: "workbook" | "image" | "text"; size: number; path: string }
export interface AvailableSkill { name: string; description: string }
export interface InstalledSkill extends AvailableSkill { removable: boolean }
export interface SavedReference { id: string; name: string; path: string; description: string; available: boolean }
export interface ModelCatalog { models: Array<{ id: string; name: string }>; defaultModelId: string | null }

export interface BridgeClientHandlers {
  onState(state: BridgeState, detail?: string): void;
  onHello(info: { sessionToken: string; serverTime: string; addinBuildId?: string }): void;
  onProviderStatus(message: ProviderStatusMessage): void;
  onProviderEvent(event: ProviderEvent): void;
  onError(message: string): void;
  /** A tool request from the bridge finished running in this workbook. */
  onToolExecuted?(tool: string, args: unknown, result: ToolResult): void;
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
      this.setState("closed");
      if (!this.closedByUser) this.scheduleReconnect();
    };
  }

  disconnect(): void {
    this.closedByUser = true;
    if (this.reconnectTimer !== null) window.clearTimeout(this.reconnectTimer);
    this.ws?.close(1000, "task_pane_closed");
    this.ws = null;
    this.setState("closed");
  }

  setProvider(provider: ProviderId): void {
    this.send({ type: "set_provider", provider });
  }

  sendPrompt(requestId: string, prompt: string, selection: SelectionSnapshot | null, context?: string, attachmentIds?: string[], model?: string | null, referencePaths?: string[], effort: "low" | "medium" | "high" = "medium", chatSessionId?: string): void {
    this.send({ type: "prompt", requestId, prompt, selection, chatSessionId, model: model ?? null, effort, ...(context ? { context } : {}), ...(attachmentIds?.length ? { attachmentIds } : {}), ...(referencePaths?.length ? { referencePaths } : {}) });
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
    this.send({ type: "cancel", requestId });
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
      case "ping":
        this.send({ type: "pong", nonce: message.nonce });
        return;
      default:
        this.handlers.onProviderEvent(message);
        return;
    }
  }

  private async handleToolRequest(message: ToolRequestMessage): Promise<void> {
    const result = message.tool === "__authorize_python"
      ? this.permissionMode === "auto" ? { ok: true as const, data: { approved: true } }
        : { ok: true as const, data: { approved: await this.askPermission("Run Python on this computer?",
          "The assistant will run Python on your computer with your user permissions, including access to files and the network. Selected file copies and generated files use a temporary working folder.",
          "Run Python", message.timeoutMs, (message.args as { code: string }).code.slice(0, 8000)) } }
      : message.tool === "__authorize_skill_script"
      ? await authorizeSkillScript(this.permissionMode, message.args as RunSkillScriptInput,
        (detail, target) => this.askPermission("Allow skill script?", detail, "Run script", message.timeoutMs, target))
      : message.tool === "__authorize_reference_path"
      ? this.permissionMode === "auto"
        ? { ok: true as const, data: { approved: true } }
        : await this.askReferencePermission(message.args as { path: string }, message.timeoutMs)
      : await executeWithPermission(
        message.tool,
        message.args,
        this.permissionMode,
        executeTool,
        (detail) => this.askPermission("Allow workbook change?", detail, "Allow change", message.timeoutMs),
      );
    this.send({ type: "tool_result", callId: message.callId, result });
    if (!message.tool.startsWith("__authorize_")) this.handlers.onToolExecuted?.(message.tool, message.args, result);
  }

  private askReferencePermission(args: { path: string }, timeoutMs: number): Promise<ToolResult> {
    return this.askPermission(
      "Allow read access?",
      "The assistant wants to read this Excel file for this chat. It cannot change the file. Edits only affect the workbook open in Excel.",
      "Allow read",
      timeoutMs,
      args.path,
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
