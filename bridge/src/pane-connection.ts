import type http from "node:http";
import type https from "node:https";
import path from "node:path";
import type { Duplex } from "node:stream";
import { WebSocketServer, type RawData, type WebSocket } from "ws";
import { HostIdSchema, PROTOCOL_VERSION, WorkbookInfoMessageSchema, type ClientMessage, type ServerMessage } from "../../protocol/src/index.js";
import type { McpServers } from "../../shared/src/mcp-servers.js";
import type { OfficeRuntime } from "./office-runtime.js";
import { attachmentPath } from "./attachments.js";
import type { Delegations } from "./delegations.js";
import type { Logger } from "./logger.js";
import type { ProviderLoginManager } from "./provider-login.js";
import type { ReferenceFiles } from "./reference-files.js";
import { normalizePromptSelection } from "./selection.js";
import type { WorkbookSessions } from "./workbook-sessions.js";
import type { WorkbookSession } from "./session.js";
import { type SkillStore, expandSkillPrompt } from "./skills.js";
import type { TurnRunner } from "./turns.js";

interface SocketState {
  session: WorkbookSession | null;
  helloDone: boolean;
  alive: boolean;
  /** The pane's last provider change, which later messages for the new provider wait for. */
  providerChange: Promise<void>;
}

export interface PaneSocketDeps {
  sessions: WorkbookSessions;
  logger: Logger;
  turns: TurnRunner;
  delegations: Delegations;
  logins: ProviderLoginManager;
  mcpServers: McpServers;
  externalTools: OfficeRuntime["externalTools"];
  skills: SkillStore;
  references: ReferenceFiles;
  /** Same rule as for HTTP requests: a foreign Origin or Host never gets a socket. */
  isTrustedRequest: (req: http.IncomingMessage) => boolean;
  addinBuildId: () => string;
}

/** The WebSocket side of the bridge: one socket per task pane, the handshake, and the messages a pane sends. */
export class PaneSockets {
  private wss: WebSocketServer | null = null;
  private readonly sockets = new Map<WebSocket, SocketState>();
  private refreshingFiles: Promise<void> | null = null;

  constructor(private readonly deps: PaneSocketDeps) {}

  attach(server: http.Server | https.Server): void {
    this.wss = new WebSocketServer({ noServer: true });
    this.wss.on("connection", (ws) => this.handleConnection(ws));
    server.on("upgrade", (req, socket, head) => this.handleUpgrade(req, socket, head));
  }

  private handleUpgrade(req: http.IncomingMessage, socket: Duplex, head: Buffer): void {
    if (!this.deps.isTrustedRequest(req)) {
      this.deps.logger.warn("rejected websocket from untrusted origin or host", { origin: req.headers.origin, host: req.headers.host,
        hint: "to reach the bridge by another hostname, run: office-partner allow-host <name>" });
      socket.end("HTTP/1.1 403 Forbidden\r\nConnection: close\r\nContent-Length: 0\r\n\r\n");
      return;
    }
    const url = new URL(req.url ?? "/", "http://localhost");
    if (url.pathname !== "/ws") {
      socket.destroy();
      return;
    }
    this.wss!.handleUpgrade(req, socket, head, (ws) => this.wss!.emit("connection", ws, req));
  }

  private handleConnection(ws: WebSocket): void {
    const { logger, logins, externalTools, mcpServers, delegations, turns } = this.deps;
    const state: SocketState = { session: null, helloDone: false, alive: true, providerChange: Promise.resolve() };
    this.sockets.set(ws, state);
    logger.debug("socket connected");

    ws.on("message", (data) => {
      this.handleClientMessage(ws, state, data).catch((error) => {
        logger.error("client message handler failed", { error: String(error) });
      });
    });

    ws.on("pong", () => {
      state.alive = true;
    });

    ws.on("close", () => {
      this.sockets.delete(ws);
      if (state.session) {
        logins.forget(state.session.id);
        void externalTools.closeSession(state.session.id);
        void mcpServers.closeSession(state.session.id);
        state.session.detachSocket(ws);
        delegations.sessionClosed(state.session);
        state.session.abortTurn();
        turns.closeProviderProcess(state.session, "socket_closed").catch(() => undefined);
        this.broadcastConnectedFiles();
      }
      logger.debug("socket closed");
    });

    ws.on("error", (error) => {
      logger.warn("socket error", { error: String(error) });
    });
  }

  private async handleClientMessage(ws: WebSocket, state: SocketState, data: RawData): Promise<void> {
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
    const { turns, delegations } = this.deps;
    switch (message.type) {
      case "workbook_info": {
        const info = WorkbookInfoMessageSchema.safeParse(message);
        if (!info.success) return;
        session.workbookName = info.data.name;
        this.broadcastConnectedFiles();
        return;
      }
      case "refresh_connected_files":
        void this.refreshConnectedFiles();
        return;
      case "set_provider": {
        const change = state.providerChange.then(() => turns.setProvider(session, message.provider));
        state.providerChange = change.catch(() => undefined);
        return change;
      }
      case "prepare": {
        // Behind a provider change that is still running, so the runtime is opened for the provider the pane chose last.
        await state.providerChange;
        void turns.warm(session, { model: message.model ?? null, effort: message.effort ?? "medium", chatSessionId: message.chatSessionId,
          resumeSessionId: message.resumeSessionId, capabilities: message.capabilities ?? { webSearch: false, playwright: false } });
        return;
      }
      case "prompt": {
        const picked = normalizePromptSelection(message.selection, session.host);
        if (!picked.ok) return void session.send({ type: "error", requestId: message.requestId, code: "bad_message", message: picked.message });
        return turns.run(session, {
          requestId: message.requestId, prompt: message.prompt, selection: picked.selection, context: message.context,
          attachmentIds: message.attachmentIds, model: message.model, referencePaths: message.referencePaths, effort: message.effort,
          chatSessionId: message.chatSessionId, resumeSessionId: message.resumeSessionId, capabilities: message.capabilities,
          isCommand: message.command, delegationId: message.delegationId,
        });
      }
      case "delegation_declined":
        delegations.decline(session, message.delegationId, typeof message.reason === "string" ? message.reason.slice(0, 200) : "");
        return;
      case "followup": {
        const respond = (state: "pending" | "delivered" | "next_turn" | "rejected", detail?: string) => session.send({ type: "followup_status", requestId: message.requestId, id: message.id, state, message: detail });
        if (session.activeRequestId !== message.requestId || session.turnController?.signal.aborted) { respond("next_turn"); return; }
        try {
          let text = expandSkillPrompt(message.prompt, this.deps.skills).prompt;
          if (message.attachmentIds?.length) {
            const files = message.attachmentIds.map(id => session.attachments.get(id));
            if (files.some(file => !file)) throw new Error("Attached file is unavailable; select it again.");
            text += "\n\n<user_attached_files>\n" + files.map(file => JSON.stringify({ name: file!.name, path: attachmentPath(file!.id), kind: file!.kind })).join("\n") + "\n</user_attached_files>";
          }
          if (message.referencePaths?.length) {
            text += this.deps.references.mentionedFilesBlock(message.referencePaths, session.referenceGrants);
          }
          // The CLI takes a message sent during a turn itself (at its next step); the queue is for CLIs that cannot.
          if (await turns.steer(session, message.requestId, text)) { respond("delivered"); return; }
          session.queueFollowup(message.requestId, message.id, text);
          respond("pending");
        } catch (error) { respond("rejected", error instanceof Error ? error.message : String(error)); }
        return;
      }
      case "cancel":
        return turns.cancel(session, message.requestId);
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

    const host = HostIdSchema.safeParse(message.clientInfo?.host ?? "excel");
    if (!host.success) {
      return this.sendTo(ws, { type: "hello_error", code: "invalid_hello", message: "clientInfo.host must be excel, word or powerpoint" });
    }
    const session = this.deps.sessions.create(
      message.workbookConnectionId,
      { ...message.clientInfo, host: host.data },
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
      addinBuildId: this.deps.addinBuildId(),
    });
    this.broadcastConnectedFiles();
  }

  /** Read every pane's workbook name now, then tell every pane. Requests that arrive meanwhile share the same round. */
  private refreshConnectedFiles(): Promise<void> {
    this.refreshingFiles ??= this.deps.sessions.refreshNames()
      .then(() => this.broadcastConnectedFiles())
      .catch((error) => this.deps.logger.warn("could not refresh connected files", { error: String(error) }))
      .finally(() => { this.refreshingFiles = null; });
    return this.refreshingFiles;
  }

  /** Tell every connected task pane which workbooks are open, each with itself first. Called whenever the list or a name changes. */
  private broadcastConnectedFiles(): void {
    const connected = this.deps.sessions.connected();
    for (const session of connected) {
      session.send({ type: "connected_files", files: this.deps.sessions.connectedFiles(session) });
      this.deps.turns.refreshTools(session);
    }
  }

  private sendTo(ws: WebSocket, message: ServerMessage): void {
    if (ws.readyState === 1) ws.send(JSON.stringify(message));
  }

  /** Heartbeat: drop sockets that did not answer the previous ping. */
  pingAll(): void {
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

  /** Tell every pane the bridge is going away; the server itself is closed later with `close()`. */
  closeSockets(): void {
    for (const [ws] of this.sockets) ws.close(1001, "server_shutdown");
  }

  async close(): Promise<void> {
    await new Promise<void>((resolve) => {
      if (!this.wss) return resolve();
      this.wss.close(() => resolve());
    });
  }
}
