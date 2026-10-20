import http from "node:http";
import https from "node:https";
import path from "node:path";
import os from "node:os";
import { createReadStream, readFileSync, renameSync, statSync, writeFileSync } from "node:fs";
import type { Duplex } from "node:stream";
import { WebSocketServer, type WebSocket, type RawData } from "ws";
import {
  PROTOCOL_VERSION,
  ProviderIdSchema,
  type ClientMessage,
  type ProviderId,
  type SelectionSnapshot,
  type ServerMessage,
} from "../../protocol/src/index.js";
import { PROJECT_ROOT, type BridgeConfig } from "./config.js";
import { UsageStore, type UsageRecord } from "./usage.js";
import { Logger } from "./logger.js";
import { ReferenceError_, ReferenceFiles, workbookFolderFromUrl } from "./reference-files.js";
import { AttachmentError, MAX_ATTACHMENT_BYTES, attachmentPath, createAttachment } from "./attachments.js";
import { SkillError } from "./skills.js";
import { UserSkills, MAX_SKILL_UPLOAD_BYTES } from "./skill-upload.js";
import { SkillStore, expandSkillPrompt, projectTemplateReminder } from "./skills.js";
import { SessionManager } from "./session-manager.js";
import { WorkbookSession } from "./session.js";
import { ToolGateway } from "./tool-gateway.js";
import { activeWorkbookPath, pickWorkspaceFolder } from "./windows-workspace.js";
import { MockAdapter, ProviderRegistry, createDefaultRegistry } from "../../providers/src/index.js";
import { ProviderLoginError, ProviderLoginManager } from "./provider-login.js";
import type { McpEndpoint } from "../../providers/src/types.js";

interface SocketState {
  session: WorkbookSession | null;
  helloDone: boolean;
  alive: boolean;
}

export interface BridgeOptions {
  registry?: ProviderRegistry;
  logger?: Logger;
  /** Override where skills are loaded from (tests). */
  skills?: SkillStore;
  /** Override the readable reference folders (tests). */
  references?: ReferenceFiles;
  /** Where user-added workspaces are saved; null disables persistence in tests. */
  workspaceConfigPath?: string | null;
  skillUploadDir?: string;
  usageFile?: string | null;
}

const MAX_BODY_BYTES = 8 * 1024 * 1024;

/** Wires the HTTP(S) + WebSocket surface to the session manager and providers. */
export class Bridge {
  private readonly usage: UsageStore;
  private readonly activeUsage = new Map<string, UsageRecord>();
  private readonly skills: SkillStore;
  readonly config: BridgeConfig;
  readonly logger: Logger;
  readonly sessions: SessionManager;
  readonly gateway: ToolGateway;
  readonly providers: ProviderRegistry;
  private readonly references: ReferenceFiles;
  private readonly workspaceConfigPath: string | null;

  private server: http.Server | https.Server | null = null;
  private wss: WebSocketServer | null = null;
  private heartbeat: NodeJS.Timeout | null = null;
  private reaper: NodeJS.Timeout | null = null;
  private readonly startedAt = Date.now();
  private readonly sockets = new Map<WebSocket, SocketState>();
  private readonly logins = new ProviderLoginManager();
  private readonly userSkills: UserSkills;

  constructor(config: BridgeConfig, options: BridgeOptions = {}) {
    this.config = config;
    this.usage = new UsageStore(options.usageFile === undefined
      ? process.env.BRIDGE_USAGE_FILE ?? path.join(PROJECT_ROOT, "run", "usage.jsonl") : options.usageFile);
    this.logger = options.logger ?? new Logger("bridge", config.logLevel);
    this.sessions = new SessionManager({
      ttlMs: config.sessionTtlMs,
      logger: this.logger.child("sessions"),
    });
    this.skills = options.skills ?? new SkillStore(config.skillDirs);
    this.userSkills = new UserSkills(options.skillUploadDir ?? path.join(os.homedir(), ".excel-partner", "skills"));
    this.references = options.references ?? new ReferenceFiles(config.referenceDirs);
    this.workspaceConfigPath = options.workspaceConfigPath === undefined
      ? options.references ? null : (process.env.BRIDGE_CONFIG ?? path.join(PROJECT_ROOT, "excel-partner.config.json"))
      : options.workspaceConfigPath;
    this.gateway = new ToolGateway(
      this.sessions,
      config.limits,
      config.toolTimeoutMs,
      this.logger.child("gateway"),
      config.allowOfficeJs,
      this.skills,
      this.references,
      this.userSkills,
    );

    const runtime = {
      bridgeUrl: config.publicBaseUrl,
      mcp: config.mcp,
      logger: this.logger.child("providers"),
    };
    this.providers = options.registry ?? createDefaultRegistry(runtime);

    const mockProvider = process.env.BRIDGE_MOCK_PROVIDER;
    if (mockProvider) {
      const parsed = ProviderIdSchema.safeParse(mockProvider);
      if (parsed.success) {
        this.providers.register(new MockAdapter({ id: parsed.data }));
        this.logger.warn("mock provider registered", { provider: parsed.data });
      } else {
        this.logger.warn("BRIDGE_MOCK_PROVIDER is not a valid provider id", { mockProvider });
      }
    }
  }

  async listen(): Promise<{ host: string; port: number; url: string }> {
    const requestHandler: http.RequestListener = (req, res) => {
      this.handleRequest(req, res).catch((error) => {
        this.logger.error("request handler failed", { error: String(error) });
        if (!res.headersSent) {
          res.writeHead(500, { "content-type": "application/json" });
        }
        res.end(JSON.stringify({ error: "internal_error" }));
      });
    };

    if (this.config.tls) {
      this.server = https.createServer(
        {
          cert: readFileSync(this.config.tls.certPath),
          key: readFileSync(this.config.tls.keyPath),
        },
        requestHandler,
      );
    } else {
      this.server = http.createServer(requestHandler);
    }

    this.wss = new WebSocketServer({ noServer: true });
    this.wss.on("connection", (ws) => this.handleConnection(ws));
    this.server.on("upgrade", (req, socket, head) => this.handleUpgrade(req, socket, head));

    await new Promise<void>((resolve, reject) => {
      this.server!.once("error", reject);
      this.server!.listen(this.config.port, this.config.host, () => resolve());
    });

    this.heartbeat = setInterval(() => this.pingAll(), this.config.heartbeatMs);
    this.heartbeat.unref?.();
    this.reaper = setInterval(() => this.sessions.reapIdle(), 60_000);
    this.reaper.unref?.();

    const address = this.server.address();
    const port = typeof address === "object" && address ? address.port : this.config.port;
    const scheme = this.config.tls ? "https" : "http";
    const url = `${scheme}://${this.config.host}:${port}`;
    this.logger.info("bridge listening", { url, tls: Boolean(this.config.tls) });
    return { host: this.config.host, port, url };
  }

  /* ----------------------------------------------------------------------- */
  /* HTTP                                                                    */
  /* ----------------------------------------------------------------------- */

  private async handleRequest(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const url = new URL(req.url ?? "/", "http://localhost");

    if (req.method === "GET" && url.pathname === "/health") {
      return this.sendJson(res, 200, {
        status: "ok",
        name: "excel-ai-addin-bridge",
        protocolVersion: PROTOCOL_VERSION,
        uptimeMs: Date.now() - this.startedAt,
        tls: Boolean(this.config.tls),
        sessions: this.sessions.list().map((session) => session.toJSON()),
      });
    }

    // Development only: lets scripts/tool-call.ts act as an independent MCP client against a
    // live task pane. Exposes session tokens, so it is off unless BRIDGE_DEV=1 (loopback bridge).
    if (req.method === "GET" && url.pathname === "/debug/sessions" && process.env.BRIDGE_DEV === "1") {
      return this.sendJson(
        res,
        200,
        this.sessions.list().map((session) => ({ ...session.toJSON(), token: session.sessionToken })),
      );
    }

    // Names and descriptions only: feeds the task pane's `/` menu.
    if (req.method === "GET" && url.pathname === "/skills") {
      return this.sendJson(res, 200, {
        skills: this.skills.list().map(({ name, description }) => ({ name, description })),
      });
    }

    if (req.method === "POST" && url.pathname === "/internal/tool") {
      return this.handleInternalTool(req, res);
    }

    if (req.method === "GET" && url.pathname === "/session/usage") {
      const session = this.sessions.getByToken(bearerToken(req.headers.authorization) ?? "");
      if (!session?.isConnected) return this.sendJson(res, 401, { error: "unauthorized" });
      const id = url.searchParams.get("sessionId");
      return this.sendJson(res, 200, { sessions: this.usage.summaries().filter((item) => !id || item.sessionId === id) });
    }
    if (req.method === "POST" && url.pathname === "/session/attachment") {
      return this.handleAttachment(req, res);
    }
    if (req.method === "GET" && url.pathname === "/session/models") return this.listProviderModels(req, res, url);
    if (url.pathname === "/session/skills") return this.handleUserSkills(req, res);
    if (url.pathname === "/session/skill-templates" && req.method === "GET") return this.listSkillTemplates(req, res, url);
    if (url.pathname === "/session/skill-templates" && req.method === "POST") return this.addSkillTemplate(req, res);
    if (url.pathname === "/session/skill-templates" && req.method === "DELETE") return this.deleteSkillTemplate(req, res);

    if (url.pathname === "/session/workspaces" && req.method === "GET") return this.listWorkspaces(req, res, url);
    if (url.pathname === "/session/workspaces" && req.method === "POST") return this.addWorkspace(req, res);
    if (url.pathname === "/session/workspaces" && req.method === "DELETE") return this.removeWorkspace(req, res);
    if (url.pathname === "/session/workspaces/pick" && req.method === "POST") return this.pickWorkspace(req, res);
    if (url.pathname === "/session/workspaces/entries" && req.method === "GET") return this.listWorkspaceEntries(req, res, url);
    if (url.pathname === "/session/workspaces/attach" && req.method === "POST") return this.attachWorkspaceFile(req, res);

    if (url.pathname === "/session/provider-login") return this.handleProviderLogin(req, res);

    if (req.method === "GET" && url.pathname.startsWith("/addin/")) {
      return this.serveAddin(url.pathname, res);
    }

    return this.sendJson(res, 404, { error: "not_found" });
  }

  /**
   * Serve the task pane assets over the same HTTPS origin as the bridge so a
   * single localhost certificate covers both the UI and the API.
   */
  private serveAddin(pathname: string, res: http.ServerResponse): void {
    if (pathname === "/addin/build-id") return this.sendJson(res, 200, { id: this.addinBuildId() });

    const root = path.join(PROJECT_ROOT, "addin");
    const relative = pathname.slice("/addin/".length);
    const target = path.resolve(root, relative);
    if (!target.startsWith(root + path.sep) && target !== root) {
      return this.sendJson(res, 403, { error: "forbidden" });
    }

    let filePath = target;
    try {
      if (statSync(filePath).isDirectory()) filePath = path.join(filePath, "index.html");
      statSync(filePath);
    } catch {
      return this.sendJson(res, 404, { error: "not_found" });
    }

    if (pathname === "/addin/taskpane/index.html") {
      const html = readFileSync(filePath, "utf8").replace("__ADDIN_BUILD_ID__", this.addinBuildId());
      res.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store" });
      res.end(html);
      return;
    }

    res.writeHead(200, {
      "content-type": contentType(filePath),
      "cache-control": "no-store",
    });
    createReadStream(filePath)
      .on("error", () => res.destroy())
      .pipe(res);
  }

  private async handleUserSkills(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!this.workspaceSession(req)) return this.sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    try {
      if (req.method === "GET") return this.sendJson(res, 200, {
        skills: this.skills.list().map(({ name, description, dir }) => ({ name, description, removable: this.userSkills.owns(name, dir) })),
      });
      if (req.method === "POST") {
        let filename: string;
        try { filename = decodeURIComponent(String(req.headers["x-file-name"] ?? "")); }
        catch { return this.sendJson(res, 400, { error: "Invalid file name." }); }
        const bytes = await readBinaryBody(req, MAX_SKILL_UPLOAD_BYTES);
        const skill = this.userSkills.install(filename, bytes, name => this.skills.get(name) !== null);
        return this.sendJson(res, 201, { skill });
      }
      if (req.method === "DELETE") {
        const body = JSON.parse(await readBody(req)) as { name?: unknown };
        if (typeof body.name !== "string") return this.sendJson(res, 400, { error: "Specify a skill name." });
        const skill = this.skills.get(body.name);
        if (!skill || !this.userSkills.owns(skill.name, skill.dir)) return this.sendJson(res, 403, { error: "Only uploaded skills can be removed here." });
        this.userSkills.remove(skill.name);
        return this.sendJson(res, 200, { removed: skill.name });
      }
      return this.sendJson(res, 405, { error: "method_not_allowed" });
    } catch (error) {
      if (error instanceof SkillError || error instanceof AttachmentError || error instanceof SyntaxError) return this.sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  private async handleProviderLogin(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = this.workspaceSession(req);
    if (!session) return this.sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    if (req.method === "GET") {
      const status = this.logins.status(session.id);
      if (status.state === "success") this.providers.invalidateModels("copilot");
      return this.sendJson(res, 200, status);
    }
    if (req.method === "DELETE") return this.sendJson(res, 200, this.logins.cancel(session.id));
    if (req.method !== "POST") return this.sendJson(res, 405, { error: "method_not_allowed" });
    if (session.isBusy) return this.sendJson(res, 409, { error: "Stop the current response before signing in." });
    try {
      const status = this.logins.start(session.id);
      return this.sendJson(res, 200, status);
    } catch (error) {
      if (error instanceof ProviderLoginError) return this.sendJson(res, 409, { error: error.message });
      throw error;
    }
  }

  private async handleInternalTool(
    req: http.IncomingMessage,
    res: http.ServerResponse,
  ): Promise<void> {
    const token = bearerToken(req.headers.authorization);
    if (!token) return this.sendJson(res, 401, { error: "unauthorized" });
    const session = this.sessions.getByToken(token);
    if (!session || !session.isConnected) {
      return this.sendJson(res, 401, { error: "unauthorized" });
    }

    let body: { sessionId?: string; tool?: string; args?: unknown };
    try {
      body = JSON.parse(await readBody(req));
    } catch {
      return this.sendJson(res, 400, { error: "invalid_json" });
    }

    // Defence in depth: the token already identifies the session, but a stray
    // session id must never route a tool to the wrong workbook.
    if (body.sessionId && body.sessionId !== session.id) {
      return this.sendJson(res, 403, { error: "session_mismatch" });
    }
    if (typeof body.tool !== "string") {
      return this.sendJson(res, 400, { error: "missing_tool" });
    }

    const result = await this.gateway.invoke({
      sessionId: session.id,
      tool: body.tool,
      args: body.args,
    });
    return this.sendJson(res, 200, result);
  }

  private async handleAttachment(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = this.sessions.getByToken(bearerToken(req.headers.authorization) ?? "");
    if (!session?.isConnected) return this.sendJson(res, 401, { error: "unauthorized" });
    let name: string;
    try { name = decodeURIComponent(String(req.headers["x-file-name"] ?? "")); }
    catch { return this.sendJson(res, 400, { error: "invalid_file_name" }); }
    try {
      const bytes = await readBinaryBody(req, MAX_ATTACHMENT_BYTES);
      const attachment = createAttachment(name, bytes);
      if (session.attachments.size >= 10) session.attachments.delete(session.attachments.keys().next().value!);
      session.attachments.set(attachment.id, attachment);
      session.touch();
      return this.sendJson(res, 200, { id: attachment.id, name: attachment.name, kind: attachment.kind, path: attachmentPath(attachment.id) });
    } catch (error) {
      if (error instanceof AttachmentError) return this.sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  private listSkillTemplates(req: http.IncomingMessage, res: http.ServerResponse, url: URL): void {
    if (!this.workspaceSession(req)) return this.sendJson(res, 401, { error: "unauthorized" });
    const skill = url.searchParams.get("skill");
    try {
      if (skill) return this.sendJson(res, 200, { templates: this.skills.listTemplates(skill) });
      return this.sendJson(res, 200, { skills: this.skills.list().map(({ name, description }) => ({ name, description })) });
    } catch (error) {
      if (error instanceof SkillError) return this.sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  private async listProviderModels(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (!this.workspaceSession(req)) return this.sendJson(res, 401, { error: "unauthorized" });
    const parsed = ProviderIdSchema.safeParse(url.searchParams.get("provider"));
    if (!parsed.success) return this.sendJson(res, 400, { error: "invalid_provider" });
    try {
      return this.sendJson(res, 200, await this.providers.listModels(parsed.data));
    } catch (error) {
      this.logger.warn("could not list provider models", { provider: parsed.data, error: String(error) });
      if (/not (?:authenticated|logged in)|(?:authentication|login|sign.?in) (?:required|failed)|unauthorized|401|please (?:log|sign) in/i.test(String(error))) {
        return this.sendJson(res, 401, { error: "Sign in to your AI account to load the available models.", code: "login_required" });
      }
      return this.sendJson(res, 503, { error: "Could not load models from the provider CLI." });
    }
  }

  private async addSkillTemplate(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!this.workspaceSession(req)) return this.sendJson(res, 401, { error: "unauthorized" });
    let skill: string, name: string;
    try {
      skill = decodeURIComponent(String(req.headers["x-skill-name"] ?? ""));
      name = decodeURIComponent(String(req.headers["x-file-name"] ?? ""));
    } catch { return this.sendJson(res, 400, { error: "invalid_name" }); }
    try {
      const bytes = await readBinaryBody(req, MAX_ATTACHMENT_BYTES);
      return this.sendJson(res, 200, { template: this.skills.addTemplate(skill, name, bytes) });
    } catch (error) {
      if (error instanceof SkillError || error instanceof AttachmentError) return this.sendJson(res, 400, { error: error.message });
      if ((error as NodeJS.ErrnoException).code === "EEXIST") return this.sendJson(res, 409, { error: "A template with this name already exists. Delete it first to replace it." });
      throw error;
    }
  }

  private async deleteSkillTemplate(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!this.workspaceSession(req)) return this.sendJson(res, 401, { error: "unauthorized" });
    let body: { skill?: unknown; name?: unknown };
    try { body = JSON.parse(await readBody(req)); }
    catch { return this.sendJson(res, 400, { error: "invalid_json" }); }
    if (typeof body.skill !== "string" || typeof body.name !== "string") return this.sendJson(res, 400, { error: "invalid_name" });
    try { this.skills.deleteTemplate(body.skill, body.name); return this.sendJson(res, 200, { deleted: true }); }
    catch (error) { if (error instanceof SkillError) return this.sendJson(res, 400, { error: error.message }); throw error; }
  }

  private workspaceSession(req: http.IncomingMessage): WorkbookSession | null {
    const session = this.sessions.getByToken(bearerToken(req.headers.authorization) ?? "");
    return session?.isConnected ? session : null;
  }

  private async listWorkspaces(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    const session = this.workspaceSession(req);
    if (!session) return this.sendJson(res, 401, { error: "unauthorized" });
    let defaultError: string | null = null;
    const documentUrl = url.searchParams.get("documentUrl") ?? "";
    const workbookName = url.searchParams.get("workbookName") ?? "";
    let folder: string | null = null;
    try { folder = workbookFolderFromUrl(documentUrl); }
    catch (error) { defaultError = error instanceof Error ? error.message : String(error); }
    if (!folder && workbookName && workbookName.length <= 255) {
      try {
        const active = await activeWorkbookPath(workbookName);
        if (active) { folder = workbookFolderFromUrl(active); defaultError = null; }
      } catch (error) {
        this.logger.warn("active workbook path lookup failed", { error: String(error) });
      }
    }
    if (folder) {
      if (session.defaultWorkspace && session.defaultWorkspace !== folder) session.referenceGrants.delete(session.defaultWorkspace);
      session.defaultWorkspace = folder;
      session.referenceGrants.add(folder);
    }
    const roots = [...new Set([...(session.defaultWorkspace ? [session.defaultWorkspace] : []), ...this.references.allowedFolders])];
    return this.sendJson(res, 200, { roots: roots.map((folder) => ({
      path: folder,
      name: path.basename(folder),
      isDefault: folder === session.defaultWorkspace,
      removable: folder !== session.defaultWorkspace && this.references.allowedFolders.includes(folder),
    })), defaultError });
  }

  private async addWorkspace(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!this.workspaceSession(req)) return this.sendJson(res, 401, { error: "unauthorized" });
    let body: { path?: unknown };
    try { body = JSON.parse(await readBody(req)); }
    catch { return this.sendJson(res, 400, { error: "invalid_json" }); }
    if (typeof body.path !== "string" || !body.path.trim() || body.path.length > 1024) return this.sendJson(res, 400, { error: "invalid_path" });
    try {
      const folder = this.references.addRoot(body.path);
      this.saveWorkspaceRoots();
      return this.sendJson(res, 200, { path: folder, name: path.basename(folder) });
    } catch (error) {
      if (error instanceof ReferenceError_) return this.sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  private async removeWorkspace(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = this.workspaceSession(req);
    if (!session) return this.sendJson(res, 401, { error: "unauthorized" });
    let body: { path?: unknown };
    try { body = JSON.parse(await readBody(req)); }
    catch { return this.sendJson(res, 400, { error: "invalid_json" }); }
    if (typeof body.path !== "string" || body.path === session.defaultWorkspace || !this.references.removeRoot(body.path)) {
      return this.sendJson(res, 400, { error: "This workspace cannot be removed." });
    }
    this.saveWorkspaceRoots();
    return this.sendJson(res, 200, { removed: true });
  }

  private async pickWorkspace(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!this.workspaceSession(req)) return this.sendJson(res, 401, { error: "unauthorized" });
    let selected: string | null;
    try { selected = await pickWorkspaceFolder(); }
    catch (error) {
      this.logger.warn("folder picker failed", { error: String(error) });
      return this.sendJson(res, 500, { error: "Could not open the Windows folder picker." });
    }
    if (!selected) return this.sendJson(res, 200, { cancelled: true });
    try {
      const folder = this.references.addRoot(selected);
      this.saveWorkspaceRoots();
      return this.sendJson(res, 200, { path: folder, name: path.basename(folder), cancelled: false });
    } catch (error) {
      if (error instanceof ReferenceError_) return this.sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  private saveWorkspaceRoots(): void {
    if (!this.workspaceConfigPath) return;
    let current: Record<string, unknown> = {};
    try { current = JSON.parse(readFileSync(this.workspaceConfigPath, "utf8")) as Record<string, unknown>; }
    catch { /* first workspace */ }
    const temporary = `${this.workspaceConfigPath}.tmp`;
    writeFileSync(temporary, JSON.stringify({ ...current, referenceDirs: this.references.allowedFolders }, null, 2) + "\n", { mode: 0o600 });
    renameSync(temporary, this.workspaceConfigPath);
  }

  private listWorkspaceEntries(req: http.IncomingMessage, res: http.ServerResponse, url: URL): void {
    const session = this.workspaceSession(req);
    if (!session) return this.sendJson(res, 401, { error: "unauthorized" });
    const folder = url.searchParams.get("path");
    if (!folder) return this.sendJson(res, 400, { error: "missing_path" });
    try {
      const listing = this.references.list(folder, session.referenceGrants);
      return this.sendJson(res, 200, {
        ...listing,
        entries: listing.entries.map((entry) => ({
          ...entry,
          path: path.join(listing.path, entry.name),
          supported: true,
        })),
      });
    } catch (error) {
      if (error instanceof ReferenceError_) return this.sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  private async attachWorkspaceFile(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = this.workspaceSession(req);
    if (!session) return this.sendJson(res, 401, { error: "unauthorized" });
    let body: { path?: unknown };
    try { body = JSON.parse(await readBody(req)); }
    catch { return this.sendJson(res, 400, { error: "invalid_json" }); }
    if (typeof body.path !== "string" || !body.path.trim()) return this.sendJson(res, 400, { error: "invalid_path" });
    try {
      const selected = this.references.selectedPath(body.path, session.referenceGrants);
      session.touch();
      return this.sendJson(res, 200, { name: path.basename(selected), kind: "file", path: selected });
    } catch (error) {
      if (error instanceof ReferenceError_ || error instanceof AttachmentError) return this.sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  private sendJson(res: http.ServerResponse, status: number, payload: unknown): void {
    const body = JSON.stringify(payload);
    res.writeHead(status, {
      "content-type": "application/json; charset=utf-8",
      "content-length": Buffer.byteLength(body),
      "access-control-allow-origin": "*",
      "cache-control": "no-store",
    });
    res.end(body);
  }

  /* ----------------------------------------------------------------------- */
  /* WebSocket                                                               */
  /* ----------------------------------------------------------------------- */

  private handleUpgrade(req: http.IncomingMessage, socket: Duplex, head: Buffer): void {
    const url = new URL(req.url ?? "/", "http://localhost");
    if (url.pathname !== "/ws") {
      socket.destroy();
      return;
    }
    this.wss!.handleUpgrade(req, socket, head, (ws) => this.wss!.emit("connection", ws, req));
  }

  private handleConnection(ws: WebSocket): void {
    const state: SocketState = { session: null, helloDone: false, alive: true };
    this.sockets.set(ws, state);
    this.logger.debug("socket connected");

    ws.on("message", (data) => {
      this.handleClientMessage(ws, state, data).catch((error) => {
        this.logger.error("client message handler failed", { error: String(error) });
      });
    });

    ws.on("pong", () => {
      state.alive = true;
    });

    ws.on("close", () => {
      this.sockets.delete(ws);
      if (state.session) {
        this.logins.forget(state.session.id);
        state.session.detachSocket(ws);
        state.session.abortTurn();
        this.closeProviderProcess(state.session, "socket_closed").catch(() => undefined);
      }
      this.logger.debug("socket closed");
    });

    ws.on("error", (error) => {
      this.logger.warn("socket error", { error: String(error) });
    });
  }

  private async handleClientMessage(
    ws: WebSocket,
    state: SocketState,
    data: RawData,
  ): Promise<void> {
    let message: ClientMessage;
    try {
      message = JSON.parse(data.toString()) as ClientMessage;
    } catch {
      return this.sendTo(ws, {
        type: "error",
        requestId: null,
        code: "bad_json",
        message: "message is not valid JSON",
      });
    }
    if (!message || typeof message.type !== "string") {
      return this.sendTo(ws, {
        type: "error",
        requestId: null,
        code: "bad_message",
        message: "message has no type",
      });
    }

    if (message.type === "hello") return this.onHello(ws, state, message);
    if (!state.helloDone || !state.session) {
      return this.sendTo(ws, {
        type: "hello_error",
        code: "rejected",
        message: "handshake required before any other message",
      });
    }

    const session = state.session;
    switch (message.type) {
      case "set_provider":
        return this.onSetProvider(session, message.provider);
      case "prompt":
        return this.onPrompt(session, message.requestId, message.prompt, message.selection, message.context, message.attachmentIds, message.model, message.referencePaths, message.effort, message.chatSessionId);
      case "cancel":
        session.abortTurn();
        return this.onCancel(session, message.requestId);
      case "tool_result":
        session.resolveTool(message.callId, message.result);
        return;
      case "pong":
        state.alive = true;
        return;
      default:
        return;
    }
  }

  private onHello(ws: WebSocket, state: SocketState, message: ClientMessage): void {
    if (message.type !== "hello") return;
    if (message.protocolVersion !== PROTOCOL_VERSION) {
      return this.sendTo(ws, {
        type: "hello_error",
        code: "protocol_mismatch",
        message: `expected protocol ${PROTOCOL_VERSION}, got ${message.protocolVersion}`,
      });
    }

    const session = this.sessions.create(
      message.workbookConnectionId,
      message.clientInfo ?? null,
    );
    session.attachSocket(ws);
    state.session = session;
    state.helloDone = true;

    this.sendTo(ws, {
      type: "hello_ok",
      protocolVersion: PROTOCOL_VERSION,
      sessionToken: session.sessionToken,
      workbookConnectionId: session.workbookConnectionId,
      serverTime: new Date().toISOString(),
      addinBuildId: this.addinBuildId(),
    });
  }

  /** Changes whenever the task pane bundle is rebuilt (its modification time). */
  private addinBuildId(): string {
    try {
      return String(Math.round(statSync(path.join(PROJECT_ROOT, "addin", "dist", "taskpane.js")).mtimeMs));
    } catch {
      return "none";
    }
  }

  private async onSetProvider(session: WorkbookSession, provider: ProviderId): Promise<void> {
    const adapter = this.providers.get(provider);
    if (!adapter) {
      return void session.send({
        type: "error",
        requestId: null,
        code: "unknown_provider",
        message: `provider '${provider}' is not available`,
      });
    }

    if (session.provider && session.provider !== provider) {
      await this.closeProviderProcess(session, "provider_change");
    }
    session.provider = provider;

    const detection = await adapter.detect(true);
    session.send({ type: "provider_status", provider, detection });
  }

  private async onPrompt(
    session: WorkbookSession,
    requestId: string,
    prompt: string,
    selection: SelectionSnapshot | null,
    context?: string,
    attachmentIds?: string[],
    model?: string | null,
    referencePaths?: string[],
    effort: "low" | "medium" | "high" = "medium",
    chatSessionId?: string,
  ): Promise<void> {
    return this.runPrompt(session, requestId, prompt, selection, context, attachmentIds, model, referencePaths, effort, chatSessionId);
  }

  private async runPrompt(
    session: WorkbookSession,
    requestId: string,
    prompt: string,
    selection: SelectionSnapshot | null,
    context?: string,
    attachmentIds?: string[],
    model?: string | null,
    referencePaths?: string[],
    effort: "low" | "medium" | "high" = "medium",
    chatSessionId?: string,
  ): Promise<void> {
    if (!session.provider) {
      return void session.send({
        type: "error",
        requestId,
        code: "no_provider",
        message: "select a provider before sending a prompt",
      });
    }
    if (session.isBusy) {
      return void session.send({
        type: "error",
        requestId,
        code: "busy",
        message: "another turn is already running for this workbook",
      });
    }

    if (attachmentIds !== undefined && (!Array.isArray(attachmentIds) || attachmentIds.length > 10 || attachmentIds.some((id) => typeof id !== "string" || !/^[0-9a-f-]{36}$/.test(id)))) {
      session.send({ type: "error", requestId, code: "invalid_attachment", message: "Invalid attached file reference." });
      session.send({ type: "turn_finished", requestId, reason: "error" });
      return;
    }
    if (referencePaths?.length) {
      try {
        const selected = referencePaths.map((file) => this.references.selectedPath(file, session.referenceGrants));
        prompt += "\n\n<user_attached_files>\n" + selected.map((file) =>
          JSON.stringify({ name: path.basename(file), path: file, kind: "reference" }),
        ).join("\n") + "\n</user_attached_files>";
      } catch (error) {
        session.send({ type: "error", requestId, code: "invalid_attachment", message: error instanceof Error ? error.message : String(error) });
        session.send({ type: "turn_finished", requestId, reason: "error" });
        return;
      }
    }
    if (attachmentIds?.length) {
      const attached = attachmentIds.map((id) => session.attachments.get(id));
      if (attached.some((file) => !file)) {
        session.send({ type: "error", requestId, code: "attachment_missing", message: "Attached file is unavailable; select it again." });
        session.send({ type: "turn_finished", requestId, reason: "error" });
        return;
      }
      prompt += "\n\n<user_attached_files>\n" + attached.map((file) =>
        JSON.stringify({ name: file!.name, path: attachmentPath(file!.id), kind: file!.kind }),
      ).join("\n") + "\n</user_attached_files>";
    }

    const selectedModel = model ?? null;
    if (session.providerSessionId && (session.providerModel !== selectedModel || session.providerEffort !== effort)) {
      await this.closeProviderProcess(session, "model_change");
    }
    const adapter = this.providers.require(session.provider);
    const controller = session.beginTurn(requestId);
    this.activeUsage.set(`${session.id}:${requestId}`, {
      sessionId: chatSessionId ?? session.id, requestId, provider: session.provider,
      model: selectedModel, startedAt: Date.now(), durationMs: 0, reason: "completed", usage: null,
    });
    // `/skill-name request` loads the skill's instructions into this turn.
    const expanded = expandSkillPrompt(prompt, this.skills);
    if (expanded.skill) this.logger.info("skill invoked", { skill: expanded.skill });
    prompt = [context, projectTemplateReminder(this.skills), expanded.prompt].filter(Boolean).join("\n\n");

    try {
      if (!session.providerSessionId) {
        session.providerSessionId = await adapter.startSession({
          workbookConnectionId: session.workbookConnectionId,
          sessionId: session.id,
          model: selectedModel,
          effort,
          mcp: this.mcpEndpointFor(session),
          signal: controller.signal,
        });
        session.providerModel = selectedModel;
        session.providerEffort = effort;
      }

      for await (const event of adapter.sendTurn({
        providerSessionId: session.providerSessionId,
        requestId,
        model: selectedModel,
        effort,
        prompt,
        selection,
        signal: controller.signal,
      })) {
        if (event.type === "turn_finished") {
          this.finishTurn(session, requestId, event.reason);
        } else if (!session.hasFinished(requestId)) {
          if (event.type === "usage_update") {
            const record = this.activeUsage.get(`${session.id}:${requestId}`);
            if (record) record.usage = event.usage;
          }
          session.send(event as ServerMessage);
        }
      }
      this.finishTurn(session, requestId, controller.signal.aborted ? "cancelled" : "completed");
    } catch (error) {
      const aborted = controller.signal.aborted;
      if (!session.hasFinished(requestId)) session.send({
        type: "error",
        requestId,
        code: aborted ? "cancelled" : "provider_error",
        message: aborted ? "turn cancelled" : String(error instanceof Error ? error.message : error),
      });
      this.finishTurn(session, requestId, aborted ? "cancelled" : "error");
    } finally {
      session.endTurn(requestId);
      session.touch();
    }
  }

  /** Emit `turn_finished` exactly once per request. */
  private finishTurn(
    session: WorkbookSession,
    requestId: string,
    reason: "completed" | "cancelled" | "error",
  ): void {
    if (!session.markFinished(requestId)) return;
    const key = `${session.id}:${requestId}`;
    const record = this.activeUsage.get(key);
    if (record) {
      this.activeUsage.delete(key);
      record.durationMs = Math.max(0, Date.now() - record.startedAt);
      record.reason = reason;
      try { this.usage.append(record); }
      catch (error) {
        this.logger.error("could not save usage", { error: String(error) });
        session.send({ type: "error", requestId, code: "usage_save_failed", message: "Usage statistics could not be saved to disk." });
      }
    }
    // Release the workbook before telling the pane, otherwise a client that
    // sends its next prompt on `turn_finished` races the adapter's cleanup
    // and gets a spurious `busy`.
    session.endTurn(requestId);
    session.send({ type: "turn_finished", requestId, reason });
  }

  private async onCancel(session: WorkbookSession, requestId: string): Promise<void> {
    if (session.activeRequestId !== requestId) return;
    if (session.provider && session.providerSessionId) {
      try {
        await this.providers.get(session.provider)?.cancelTurn(session.providerSessionId);
      } catch (error) {
        this.logger.warn("cancelTurn failed", { error: String(error) });
      }
    }
    // The adapter normally ends its stream right after cancel. If it hangs,
    // report the cancellation anyway and drop the provider process so the
    // next prompt starts clean instead of hitting `busy` forever.
    const watchdog = setTimeout(() => {
      if (session.activeRequestId !== requestId) return;
      this.logger.warn("provider ignored cancel; closing its process", { requestId });
      this.finishTurn(session, requestId, "cancelled");
      session.endTurn(requestId);
      void this.closeProviderProcess(session, "cancel_timeout");
    }, this.config.cancelGraceMs);
    watchdog.unref();
  }

  private async closeProviderProcess(session: WorkbookSession, reason: string): Promise<void> {
    session.abortTurn();
    const provider = session.provider;
    const providerSessionId = session.providerSessionId;
    session.providerSessionId = null;
    session.providerModel = null;
    session.activeRequestId = null;
    if (!provider || !providerSessionId) return;
    try {
      await this.providers.get(provider)?.closeSession(providerSessionId);
      this.logger.debug("provider session closed", { provider, reason });
    } catch (error) {
      this.logger.warn("closeSession failed", { error: String(error) });
    }
  }

  private mcpEndpointFor(session: WorkbookSession): McpEndpoint {
    return {
      bridgeUrl: this.config.publicBaseUrl,
      sessionId: session.id,
      token: session.sessionToken,
      command: this.config.mcp.command,
      args: this.config.mcp.args,
      env: {
        EXCEL_BRIDGE_URL: this.config.publicBaseUrl,
        EXCEL_SESSION_ID: session.id,
        EXCEL_SESSION_TOKEN: session.sessionToken,
        ...(this.config.allowOfficeJs ? { EXCEL_ALLOW_OFFICE_JS: "1" } : {}),
        ...(this.config.tls ? { NODE_EXTRA_CA_CERTS: this.config.tls.certPath } : {}),
      },
    };
  }

  private sendTo(ws: WebSocket, message: ServerMessage): void {
    if (ws.readyState === 1) ws.send(JSON.stringify(message));
  }

  private pingAll(): void {
    for (const [ws, state] of this.sockets) {
      if (!state.alive) {
        ws.terminate();
        continue;
      }
      state.alive = false;
      try {
        ws.ping();
      } catch {
        /* ignore */
      }
    }
  }

  /* ----------------------------------------------------------------------- */
  /* Lifecycle                                                               */
  /* ----------------------------------------------------------------------- */

  async close(): Promise<void> {
    this.logins.dispose();
    if (this.heartbeat) clearInterval(this.heartbeat);
    if (this.reaper) clearInterval(this.reaper);
    for (const [ws] of this.sockets) ws.close(1001, "server_shutdown");
    await this.providers.disposeAll();
    await new Promise<void>((resolve) => {
      if (!this.wss) return resolve();
      this.wss.close(() => resolve());
    });
    await new Promise<void>((resolve) => {
      if (!this.server) return resolve();
      this.server.close(() => resolve());
    });
  }
}

const CONTENT_TYPES: Record<string, string> = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".map": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon",
};

function contentType(filePath: string): string {
  return CONTENT_TYPES[path.extname(filePath).toLowerCase()] ?? "application/octet-stream";
}

function bearerToken(header: string | undefined): string | null {
  if (!header) return null;
  const match = /^Bearer\s+(.+)$/i.exec(header.trim());
  return match ? match[1]!.trim() : null;
}

function readBody(req: http.IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks: Buffer[] = [];
    req.on("data", (chunk: Buffer) => {
      size += chunk.length;
      if (size > MAX_BODY_BYTES) {
        reject(new Error("body too large"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

async function readBinaryBody(req: http.IncomingMessage, limit: number): Promise<Buffer> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    size += bytes.length;
    if (size > limit) {
      req.resume();
      throw new AttachmentError(`file exceeds the ${Math.floor(limit / 1024 / 1024)} MB limit`);
    }
    chunks.push(bytes);
  }
  return Buffer.concat(chunks);
}
