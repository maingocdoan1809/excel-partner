import type { ClientInfo } from "../../protocol/src/index.js";
import { WorkbookSession } from "./session.js";
import type { Logger } from "./logger.js";

export interface SessionManagerOptions {
  ttlMs: number;
  maxSessions?: number;
  logger?: Logger;
}

/**
 * Owns the lifetime of every workbook session and the lookup tables used to
 * authenticate the task pane (`sessionToken`), the MCP side (session id) and
 * reconnect attempts (`workbookConnectionId`).
 */
export class SessionManager {
  private readonly sessions = new Map<string, WorkbookSession>();
  private readonly byToken = new Map<string, WorkbookSession>();
  private readonly byConnection = new Map<string, WorkbookSession>();
  private readonly ttlMs: number;
  private readonly maxSessions: number;
  private readonly log?: Logger;

  constructor(options: SessionManagerOptions) {
    this.ttlMs = options.ttlMs;
    this.maxSessions = options.maxSessions ?? 8;
    this.log = options.logger;
  }

  get size(): number {
    return this.sessions.size;
  }

  create(workbookConnectionId: string, clientInfo: ClientInfo | null): WorkbookSession {
    const existing = this.byConnection.get(workbookConnectionId);
    if (existing) this.remove(existing, "reconnect");

    if (this.sessions.size >= this.maxSessions) {
      const oldest = [...this.sessions.values()].sort((a, b) => a.lastActivityAt - b.lastActivityAt)[0];
      if (oldest) this.remove(oldest, "evicted");
    }

    const session = new WorkbookSession(workbookConnectionId, clientInfo);
    this.sessions.set(session.id, session);
    this.byToken.set(session.sessionToken, session);
    this.byConnection.set(workbookConnectionId, session);
    this.log?.info("session created", { sessionId: session.id, workbookConnectionId });
    return session;
  }

  get(id: string): WorkbookSession | undefined {
    return this.sessions.get(id);
  }

  getByToken(token: string): WorkbookSession | undefined {
    return this.byToken.get(token);
  }

  getByConnection(workbookConnectionId: string): WorkbookSession | undefined {
    return this.byConnection.get(workbookConnectionId);
  }

  list(): WorkbookSession[] {
    return [...this.sessions.values()];
  }

  remove(session: WorkbookSession, reason: string): void {
    if (!this.sessions.delete(session.id)) return;
    session.abortTurn();
    this.byToken.delete(session.sessionToken);
    this.byConnection.delete(session.workbookConnectionId);
    session.rejectAllPending({ code: "disconnected", message: `session closed (${reason})` });
    try {
      session.socket?.close(1000, reason);
    } catch {
      /* ignore */
    }
    this.log?.info("session removed", { sessionId: session.id, reason });
  }

  removeById(id: string, reason: string): boolean {
    const session = this.sessions.get(id);
    if (!session) return false;
    this.remove(session, reason);
    return true;
  }

  /** Drop sessions that have been idle longer than the TTL. */
  reapIdle(now = Date.now()): number {
    let reaped = 0;
    for (const session of [...this.sessions.values()]) {
      if (now - session.lastActivityAt > this.ttlMs) {
        this.remove(session, "idle_timeout");
        reaped++;
      }
    }
    return reaped;
  }
}
