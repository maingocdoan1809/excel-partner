import http from "node:http";
import https from "node:https";
import { timingSafeEqual } from "node:crypto";
import os from "node:os";
import path from "node:path";
import { createReadStream, readFileSync, statSync } from "node:fs";
import { PROTOCOL_VERSION, ProviderIdSchema } from "../../protocol/src/index.js";
import { MockAdapter, ProviderRegistry, ProviderSessionManager, createDefaultRegistry } from "../../providers/src/index.js";
import { officeToolCatalog } from "../../runtime/src/catalog.js";
import { McpServers } from "../../runtime/src/mcp-servers.js";
import { OfficeRuntime } from "../../runtime/src/office-runtime.js";
import { PROJECT_ROOT, type BridgeConfig } from "./config.js";
import { Delegations } from "./delegations.js";
import { DocumentBackups } from "./document-backups.js";
import { bearerToken, connectedSession, contentType, readBody, sendJson } from "./http-util.js";
import { Logger } from "./logger.js";
import { PaneSockets } from "./pane-connection.js";
import { ProviderLoginManager } from "./provider-login.js";
import { ReferenceFiles } from "./reference-files.js";
import { FileRoutes } from "./routes/files.js";
import { McpRoutes } from "./routes/mcp.js";
import { ProviderRoutes } from "./routes/providers.js";
import { SkillRoutes } from "./routes/skills.js";
import { TestingRoutes } from "./routes/testing.js";
import { WorkspaceRoutes } from "./routes/workspaces.js";
import { SystemRoutes } from "./routes/system.js";
import { SavedReferences } from "./saved-references.js";
import { SessionManager } from "./session-manager.js";
import { discoverSessionMcp } from "./session-mcp.js";
import { WorkbookSession } from "./session.js";
import { SiteKnowledge } from "./site-knowledge.js";
import { UserSkills } from "./skill-upload.js";
import { SkillStore } from "./skills.js";
import { standalonePage } from "./standalone-page.js";
import { SystemPromptStore } from "./system-prompt.js";
import { loginSession } from "./test-login.js";
import { TestScripts } from "./test-scripts.js";
import { TestSessions } from "./test-sessions.js";
import { TestingAdmin } from "./testing-admin.js";
import { TurnRunner } from "./turns.js";
import { UpdateError, UpdateService } from "./updater.js";
import { UsageStore } from "./usage.js";
import { enrichWindowsSelection, type SelectionEnricher } from "./windows-selection.js";
import { activateWorkbook } from "./windows-workspace.js";

export interface BridgeOptions {
  onShutdown?: () => void;
  selectionEnricher?: SelectionEnricher;
  registry?: ProviderRegistry;
  logger?: Logger;
  /** Override where skills are loaded from (tests). */
  skills?: SkillStore;
  /** Override the readable reference folders (tests). */
  references?: ReferenceFiles;
  /** Where user-added workspaces are saved; null disables persistence in tests. */
  workspaceConfigPath?: string | null;
  skillUploadDir?: string;
  systemPromptFile?: string | null;
  usageFile?: string | null;
  savedReferencesFile?: string | null;
  /** Replaces the updater (tests); null turns updating from the task pane off. */
  updater?: UpdateService | null;
  /** Replaces bringing a workbook's window to the front (tests); it is Excel on Windows only. */
  activateWorkbook?: (workbookName: string) => Promise<unknown>;
}

/**
 * The bridge process: wires the stores, the Office tool runtime and the providers together, serves the task pane over
 * HTTPS, dispatches its HTTP requests to the route groups under `routes/`, and hands each pane's WebSocket to
 * `PaneSockets`, which runs turns through `TurnRunner`.
 */
export class Bridge {
  readonly config: BridgeConfig;
  readonly logger: Logger;
  readonly sessions: SessionManager;
  readonly gateway: OfficeRuntime;
  readonly providers: ProviderRegistry;

  private server: http.Server | https.Server | null = null;
  private heartbeat: NodeJS.Timeout | null = null;
  private reaper: NodeJS.Timeout | null = null;
  private readonly startedAt = Date.now();
  private maintenance = false;
  private readonly onShutdown?: () => void;
  private readonly skills: SkillStore;
  private readonly logins = new ProviderLoginManager();
  private readonly mcpServers: McpServers;
  private readonly updater: UpdateService | null;
  private readonly panes: PaneSockets;
  private readonly routes: {
    files: FileRoutes; mcp: McpRoutes; providers: ProviderRoutes; skills: SkillRoutes; testing: TestingRoutes; workspaces: WorkspaceRoutes; system: SystemRoutes;
  };

  constructor(config: BridgeConfig, options: BridgeOptions = {}) {
    this.onShutdown = options.onShutdown;
    this.config = config;
    const enrichSelection = options.selectionEnricher ?? enrichWindowsSelection;
    const usage = new UsageStore(options.usageFile === undefined
      ? process.env.BRIDGE_USAGE_FILE ?? path.join(PROJECT_ROOT, "run", "usage.jsonl") : options.usageFile);
    this.logger = options.logger ?? new Logger("bridge", config.logLevel);
    this.sessions = new SessionManager({
      ttlMs: config.sessionTtlMs,
      logger: this.logger.child("sessions"),
    });
    this.skills = options.skills ?? new SkillStore(config.skillDirs);
    const userSkills = new UserSkills(options.skillUploadDir ?? path.join(os.homedir(), ".excel-partner", "skills"));
    const systemPrompt = new SystemPromptStore(options.systemPromptFile === undefined
      ? process.env.BRIDGE_CONFIG ? path.join(path.dirname(process.env.BRIDGE_CONFIG), "system-prompt.txt") : path.join(PROJECT_ROOT, "run", "system-prompt.txt")
      : options.systemPromptFile);
    const references = options.references ?? new ReferenceFiles(config.referenceDirs);
    const savedReferences = new SavedReferences(options.savedReferencesFile === undefined
      ? path.join(os.homedir(), ".excel-partner", "references.json") : options.savedReferencesFile);
    references.setSavedPaths(() => savedReferences.paths());
    const workspaceConfigPath = options.workspaceConfigPath === undefined
      ? options.references ? null : (process.env.BRIDGE_CONFIG ?? path.join(PROJECT_ROOT, "excel-partner.config.json"))
      : options.workspaceConfigPath;
    const dataDir = path.dirname(config.configFile);
    const testSessions = new TestSessions(path.join(dataDir, "sessions"));
    const siteKnowledge = new SiteKnowledge(path.join(dataDir, "knowledge"));
    const testScripts = new TestScripts(path.join(dataDir, "scripts"));
    const testingAdmin = new TestingAdmin(testSessions, siteKnowledge, testScripts, login => loginSession({ ...login, toolsDir: path.join(dataDir, "tools") }));
    this.mcpServers = new McpServers(path.join(dataDir, "mcp-servers.json"));
    const backups = new DocumentBackups(path.join(dataDir, "backups"));
    this.updater = options.updater === undefined
      ? new UpdateService({
          packageRoot: PROJECT_ROOT,
          dataDir,
          gatewayDir: process.env.AGENT_GATEWAY_HOME ?? path.join(process.env.LOCALAPPDATA || path.join(os.homedir(), ".config"), "Agent Gateway"),
        })
      : options.updater;
    // `turns` is assigned a few lines below; a delegation can only be cancelled once a turn exists.
    const delegations = new Delegations((session, requestId) => void turns.cancel(session, requestId));
    this.gateway = new OfficeRuntime(
      this.sessions,
      config.limits,
      config.toolTimeoutMs,
      this.logger.child("gateway"),
      config.allowOfficeJs,
      this.skills,
      references,
      userSkills,
      enrichSelection,
      testSessions,
      siteKnowledge,
      testScripts,
      delegations,
    );
    this.gateway.mcpServers = this.mcpServers;

    this.providers = options.registry ?? createDefaultRegistry({
      bridgeUrl: config.publicBaseUrl,
      mcp: config.mcp,
      logger: this.logger.child("providers"),
    });
    const providerSessions = new ProviderSessionManager<WorkbookSession>(this.providers,
      session => this.gateway.externalTools.closeSession(session.id), this.logger);
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

    const discoverMcp = (session: WorkbookSession) => discoverSessionMcp(this.mcpServers, config.allowOfficeJs, session);
    const turns = new TurnRunner({
      isUpdating: () => this.maintenance,
      config, logger: this.logger, providers: this.providers, providerSessions, delegations, externalTools: this.gateway.externalTools,
      usage, skills: this.skills, systemPrompt, savedReferences, references, enrichSelection, discoverMcp,
    });
    this.panes = new PaneSockets({
      sessions: this.sessions, logger: this.logger, turns, delegations, logins: this.logins, mcpServers: this.mcpServers,
      externalTools: this.gateway.externalTools, skills: this.skills, references,
      isTrustedRequest: req => this.isTrustedRequest(req), addinBuildId: () => this.addinBuildId(),
    });
    this.routes = {
      system: new SystemRoutes(this.sessions, this.providers, this.updater, () => this.logins.isBusy),
      files: new FileRoutes({ sessions: this.sessions, backups, logger: this.logger, activateWorkbook: options.activateWorkbook ?? activateWorkbook }),
      mcp: new McpRoutes({ sessions: this.sessions, mcpServers: this.mcpServers, discover: discoverMcp }),
      providers: new ProviderRoutes({ sessions: this.sessions, providers: this.providers, logins: this.logins, usage, logger: this.logger }),
      skills: new SkillRoutes({ sessions: this.sessions, skills: this.skills, userSkills, systemPrompt }),
      testing: new TestingRoutes({ sessions: this.sessions, admin: testingAdmin }),
      workspaces: new WorkspaceRoutes({ sessions: this.sessions, references, savedReferences, logger: this.logger, workspaceConfigPath }),
    };
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
    this.panes.attach(this.server);

    await new Promise<void>((resolve, reject) => {
      this.server!.once("error", reject);
      this.server!.listen(this.config.port, this.config.host, () => resolve());
    });

    this.heartbeat = setInterval(() => this.panes.pingAll(), this.config.heartbeatMs);
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

  /**
   * The task pane is served by this bridge, so its requests are same-origin and its Host is a name we expect.
   * A web page the user happens to visit sends a foreign Origin (or, after DNS rebinding, a foreign Host), and
   * must never obtain a session token: tool approvals travel over the session's own socket.
   * Clients that are not browsers (the MCP process, CLIs) send no Origin and are allowed.
   */
  private isTrustedRequest(req: http.IncomingMessage): boolean {
    const host = req.headers.host;
    if (!host) return false;
    let hostname: string;
    try { hostname = new URL(`http://${host}`).hostname.toLowerCase().replace(/^\[(.*)\]$/, "$1"); }
    catch { return false; }
    if (!this.config.allowedHosts.includes(hostname)) return false;
    const origin = req.headers.origin;
    if (origin === undefined) return true;
    try { return new URL(origin).host.toLowerCase() === host.toLowerCase(); }
    catch { return false; }
  }

  private async handleRequest(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!this.isTrustedRequest(req)) {
      this.logger.warn("rejected request from untrusted origin or host", { origin: req.headers.origin, host: req.headers.host,
        hint: "to reach the bridge by another hostname, run: office-partner allow-host <name>" });
      res.writeHead(403, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
      res.end(JSON.stringify({ error: "forbidden_origin" }));
      return;
    }
    const url = new URL(req.url ?? "/", "http://localhost");
    const method = req.method;
    const { files, mcp, providers, skills, testing, workspaces } = this.routes;
    if (this.maintenance && url.pathname.startsWith("/session/") && method !== "GET") return sendJson(res, 503, { error: "Office Partner is switching releases. Retry after it reconnects." });
    if (url.pathname === "/session/system") return this.routes.system.handle(req, res, url);

    if (method === "POST" && ["/admin/maintenance", "/admin/resume", "/admin/shutdown"].includes(url.pathname)) {
      const configured = process.env.OFFICE_PARTNER_CONTROL_TOKEN, supplied = req.headers["x-office-partner-control"];
      if (!configured || typeof supplied !== "string" || Buffer.byteLength(configured) !== Buffer.byteLength(supplied)
        || !timingSafeEqual(Buffer.from(configured), Buffer.from(supplied))) return sendJson(res, 401, { error: "unauthorized" });
      const busy = this.logins.isBusy || this.sessions.list().some(session => session.isBusy);
      if (url.pathname === "/admin/resume") { this.maintenance = false; return sendJson(res, 200, { busy, maintenance: false }); }
      if (busy) return sendJson(res, 409, { busy: true });
      this.maintenance = true;
      if (url.pathname === "/admin/shutdown") {
        if (!this.onShutdown) return sendJson(res, 503, { error: "shutdown_unavailable" });
        sendJson(res, 200, { status: "stopping" }); setImmediate(this.onShutdown); return;
      }
      return sendJson(res, 200, { busy: false, maintenance: true });
    }

    if (url.pathname === "/session/mcp") return mcp.handle(req, res);
    if (method === "GET" && (url.pathname === "/" || url.pathname === "/index.html")) return this.serveStandalonePage(res);
    if (method === "GET" && url.pathname === "/health") {
      return sendJson(res, 200, {
        status: "ok",
        name: "excel-ai-addin-bridge",
        protocolVersion: PROTOCOL_VERSION,
        instance: process.env.OFFICE_PARTNER_BRIDGE_INSTANCE ?? null,
        controlSupported: true,
        busy: this.logins.isBusy || this.sessions.list().some(session => session.isBusy),
        maintenance: this.maintenance,
        uptimeMs: Date.now() - this.startedAt,
        tls: Boolean(this.config.tls),
        sessions: this.sessions.list().map((session) => session.toJSON()),
      });
    }
    // Development only: lets scripts/tool-call.ts act as an independent MCP client against a
    // live task pane. Exposes session tokens, so it is off unless BRIDGE_DEV=1 (loopback bridge).
    if (method === "GET" && url.pathname === "/debug/sessions" && process.env.BRIDGE_DEV === "1") {
      return sendJson(res, 200, this.sessions.list().map((session) => ({ ...session.toJSON(), token: session.sessionToken })));
    }
    // Names and descriptions only: feeds the task pane's `/` menu.
    if (method === "GET" && url.pathname === "/skills") {
      return sendJson(res, 200, { skills: this.skills.list().map(({ name, description }) => ({ name, description })) });
    }
    if (method === "GET" && url.pathname === "/internal/tools") {
      const session = connectedSession(this.sessions, req);
      if (!session) return sendJson(res, 401, { error: "unauthorized" });
      const external = await discoverSessionMcp(this.mcpServers, this.config.allowOfficeJs, session);
      return sendJson(res, 200, { tools: [...officeToolCatalog(session.host, this.config.allowOfficeJs, session.capabilities), ...[...external.values()].map(tool => tool.definition)] });
    }
    if (method === "POST" && url.pathname === "/internal/tool") return this.handleInternalTool(req, res);

    if (method === "GET" && url.pathname === "/session/usage") return providers.usage(req, res, url);
    if (method === "POST" && url.pathname === "/session/attachment") return files.attachment(req, res);
    if (method === "POST" && url.pathname === "/session/backup") return files.backup(req, res);
    if (method === "POST" && url.pathname === "/session/go-to") return files.goTo(req, res);
    if (url.pathname === "/session/update" && (method === "GET" || method === "POST")) return this.handleUpdate(req, res, url);
    if (method === "GET" && url.pathname === "/session/models") return providers.models(req, res, url);
    if (method === "GET" && url.pathname === "/session/commands") return providers.commands(req, res, url);
    if (url.pathname === "/session/references") return workspaces.savedReferences(req, res);
    if (url.pathname === "/session/skills") return skills.userSkills(req, res, url);
    if (url.pathname === "/session/system-prompt") return skills.systemPrompt(req, res);
    if (url.pathname === "/session/skill-templates" && method === "GET") return skills.listTemplates(req, res, url);
    if (url.pathname === "/session/skill-templates" && method === "POST") return skills.addTemplate(req, res);
    if (url.pathname === "/session/skill-templates" && method === "DELETE") return skills.deleteTemplate(req, res);

    if (url.pathname === "/session/workspaces" && method === "GET") return workspaces.list(req, res, url);
    if (url.pathname === "/session/workspaces" && method === "POST") return workspaces.add(req, res);
    if (url.pathname === "/session/workspaces" && method === "DELETE") return workspaces.remove(req, res);
    if (url.pathname === "/session/workspaces/pick" && method === "POST") return workspaces.pick(req, res);
    if (url.pathname === "/session/workspaces/entries" && method === "GET") return workspaces.entries(req, res, url);
    if (url.pathname === "/session/workspaces/attach" && method === "POST") return workspaces.attach(req, res);

    if (url.pathname === "/session/testing") return testing.handle(req, res, url);
    if (url.pathname === "/session/provider-login") return providers.login(req, res);

    if (method === "GET" && url.pathname.startsWith("/addin/")) return this.serveAddin(url.pathname, res);

    return sendJson(res, 404, { error: "not_found" });
  }

  /** The task pane for a browser tab: the add-in's page with a stand-in for Office.js (see standalone-page.ts). */
  private serveStandalonePage(res: http.ServerResponse): void {
    try {
      const html = readFileSync(path.join(PROJECT_ROOT, "addin", "taskpane", "index.html"), "utf8").replace("__ADDIN_BUILD_ID__", this.addinBuildId());
      res.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store" });
      res.end(standalonePage(html));
    } catch (error) {
      this.logger.error("could not serve the browser page", { error: String(error) });
      sendJson(res, 500, { error: "page_unavailable" });
    }
  }

  /**
   * Serve the task pane assets over the same HTTPS origin as the bridge so a
   * single localhost certificate covers both the UI and the API.
   */
  private serveAddin(pathname: string, res: http.ServerResponse): void {
    if (pathname === "/addin/build-id") return sendJson(res, 200, { id: this.addinBuildId() });

    const root = path.join(PROJECT_ROOT, "addin");
    const relative = pathname.slice("/addin/".length);
    const target = path.resolve(root, relative);
    if (!target.startsWith(root + path.sep) && target !== root) {
      return sendJson(res, 403, { error: "forbidden" });
    }

    let filePath = target;
    try {
      if (statSync(filePath).isDirectory()) filePath = path.join(filePath, "index.html");
      statSync(filePath);
    } catch {
      return sendJson(res, 404, { error: "not_found" });
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

  /** Changes whenever the task pane bundle is rebuilt (its modification time). */
  private addinBuildId(): string {
    try {
      // Version and size too: packages installed from npm carry one fixed file date, so the date alone never changes.
      const bundle = statSync(path.join(PROJECT_ROOT, "addin", "dist", "taskpane.js"));
      return `${this.updater?.version ?? "0"}-${Math.round(bundle.mtimeMs)}-${bundle.size}`;
    } catch {
      return "none";
    }
  }

  /** A tool call from the provider's MCP process, routed to the pane that owns the session token. */
  private async handleInternalTool(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const token = bearerToken(req.headers.authorization);
    if (!token) return sendJson(res, 401, { error: "unauthorized" });
    const session = this.sessions.getByToken(token);
    if (!session || !session.isConnected) {
      return sendJson(res, 401, { error: "unauthorized" });
    }

    let body: { sessionId?: string; tool?: string; args?: unknown };
    try {
      body = JSON.parse(await readBody(req));
    } catch {
      return sendJson(res, 400, { error: "invalid_json" });
    }

    // Defence in depth: the token already identifies the session, but a stray
    // session id must never route a tool to the wrong workbook.
    if (body.sessionId && body.sessionId !== session.id) {
      return sendJson(res, 403, { error: "session_mismatch" });
    }
    if (typeof body.tool !== "string") {
      return sendJson(res, 400, { error: "missing_tool" });
    }

    // A call that waits on another file can outlive the MCP process that made it; stop it when that process hangs up.
    const callerGone = new AbortController();
    res.on("close", () => { if (!res.writableFinished) callerGone.abort(); });
    const result = await this.gateway.invoke({
      sessionId: session.id,
      tool: body.tool,
      args: body.args,
      signal: callerGone.signal,
    });
    return sendJson(res, 200, result);
  }

  /** GET: is a newer version out, and how is an update going. POST: start one (the bridge is replaced while it runs). */
  private async handleUpdate(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (!connectedSession(this.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    if (!this.updater) return sendJson(res, 404, { error: "Updating from the task pane is turned off." });
    if (req.method === "GET") {
      return sendJson(res, 200, await this.updater.check(url.searchParams.get("refresh") === "1"));
    }
    try {
      const busy = this.sessions.list().some((other) => other.isBusy);
      const status = this.updater.start({ busy });
      this.logger.info("update started", { from: this.updater.version });
      return sendJson(res, 202, { status });
    } catch (error) {
      if (error instanceof UpdateError) return sendJson(res, error.code === "unsupported" ? 400 : error.code === "failed" ? 500 : 409, { error: error.message, code: error.code });
      throw error;
    }
  }

  /* ----------------------------------------------------------------------- */
  /* Lifecycle                                                               */
  /* ----------------------------------------------------------------------- */

  async close(): Promise<void> {
    this.logins.dispose();
    if (this.heartbeat) clearInterval(this.heartbeat);
    if (this.reaper) clearInterval(this.reaper);
    this.panes.closeSockets();
    await this.mcpServers.close();
    await this.gateway.externalTools.close();
    await this.providers.disposeAll();
    await this.panes.close();
    await new Promise<void>((resolve) => {
      if (!this.server) return resolve();
      this.server.close(() => resolve());
    });
  }
}
