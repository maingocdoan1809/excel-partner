import type http from "node:http";
import { connectedSession, readBody, sendJson } from "../http-util.js";
import type { WorkbookSessions } from "../workbook-sessions.js";
import { type FolderKind, type TestingAdmin, TestingAdminError } from "../testing-admin.js";

export interface TestingRouteDeps {
  sessions: WorkbookSessions;
  admin: TestingAdmin;
}

/** Sign-ins, notes and regression scripts for the testing workflow: what the `office-partner login/knowledge/scripts` commands do, for buttons. */
export class TestingRoutes {
  constructor(private readonly deps: TestingRouteDeps) {}

  async handle(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    const admin = this.deps.admin;
    try {
      if (req.method === "GET") {
        const kind = url.searchParams.get("view");
        if (kind === "note") return sendJson(res, 200, admin.viewNote(url.searchParams.get("site") ?? "", url.searchParams.get("name") ?? ""));
        if (kind === "script") return sendJson(res, 200, { content: admin.viewScript(url.searchParams.get("site") ?? "", url.searchParams.get("caseId") ?? "") });
        return sendJson(res, 200, admin.overview());
      }
      if (req.method !== "POST" && req.method !== "DELETE") return sendJson(res, 405, { error: "method_not_allowed" });
      const body = JSON.parse(await readBody(req)) as Record<string, unknown>;
      const text = (key: string) => typeof body[key] === "string" ? body[key] as string : "";
      if (req.method === "DELETE") {
        const kind = text("kind");
        if (kind === "signin") return sendJson(res, 200, { removed: admin.deleteSignIn(text("name")) });
        if (kind === "note") return sendJson(res, 200, { removed: admin.deleteNote(text("site"), text("name")) });
        if (kind === "script") return sendJson(res, 200, { removed: admin.deleteScript(text("site"), text("caseId")) });
        throw new TestingAdminError("Say what to delete: signin, note or script.");
      }
      const action = text("action");
      if (action === "sign-in-start") {
        const extra = Array.isArray(body.extraOrigins) ? body.extraOrigins.filter((item): item is string => typeof item === "string").slice(0, 5) : [];
        return sendJson(res, 200, await admin.startSignIn({ name: text("name"), url: text("url"), extraOrigins: extra }));
      }
      if (action === "sign-in-finish") return sendJson(res, 200, await admin.finishSignIn());
      if (action === "sign-in-cancel") return sendJson(res, 200, { cancelled: admin.cancelSignIn() });
      if (action === "open-folder") {
        const kind = text("kind");
        if (kind !== "sessions" && kind !== "knowledge" && kind !== "scripts") throw new TestingAdminError("Choose sessions, knowledge or scripts.");
        return sendJson(res, 200, { folder: admin.folder(kind as FolderKind) });
      }
      throw new TestingAdminError("Unknown action.");
    } catch (error) {
      if (error instanceof TestingAdminError || error instanceof SyntaxError) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }
}
