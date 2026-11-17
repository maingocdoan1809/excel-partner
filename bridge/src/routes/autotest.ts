import { randomBytes, timingSafeEqual } from "node:crypto";
import fs from "node:fs";
import type http from "node:http";
import { connectedSession, readBinaryBody, readBody, sendJson } from "../http-util.js";
import type { WorkbookSessions } from "../workbook-sessions.js";
import { AutotestHomeError, type AutotestHome } from "../autotest-home.js";
import type { AutotestRuntime } from "../autotest-runtime.js";
import { dashboardPage } from "../autotest-dashboard.js";

export const MAX_WORKBOOK_BYTES = 50 * 1024 * 1024;

export interface AutotestRouteDeps { sessions: WorkbookSessions; home: AutotestHome; runtime: AutotestRuntime }

/**
 * The Test screen's side of the bridge (`/session/autotest`, for the task pane) and the read-only dashboard (`/tests`, for a browser tab).
 * The dashboard page carries a random token that only a page served by this bridge can read, so another site cannot call its API.
 */
export class AutotestRoutes {
  private readonly token = randomBytes(18).toString("hex");

  constructor(private readonly deps: AutotestRouteDeps) {}

  private tokenOk(url: URL): boolean {
    const given = Buffer.from(url.searchParams.get("t") ?? "");
    const want = Buffer.from(this.token);
    return given.length === want.length && timingSafeEqual(given, want);
  }

  /** `/tests` and `/tests/api/*`. */
  async dashboard(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (req.method !== "GET") return sendJson(res, 405, { error: "method_not_allowed" });
    if (url.pathname === "/tests" || url.pathname === "/tests/") {
      res.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store", "x-content-type-options": "nosniff" });
      res.end(dashboardPage(this.token));
      return;
    }
    if (!this.tokenOk(url)) return sendJson(res, 401, { error: "unauthorized" });
    const runs = this.deps.home.runs();
    const id = url.searchParams.get("id") ?? "";
    switch (url.pathname) {
      case "/tests/api/runs": return sendJson(res, 200, { runs: runs.list() });
      case "/tests/api/run": {
        const run = runs.state(id);
        return run ? sendJson(res, 200, { run }) : sendJson(res, 404, { error: "run_not_found" });
      }
      case "/tests/api/knowledge": return sendJson(res, 200, { lines: this.deps.home.knowledge() });
      case "/tests/api/image": {
        const file = runs.image(id, url.searchParams.get("path") ?? "");
        if (!file) return sendJson(res, 404, { error: "image_not_found" });
        res.writeHead(200, { "content-type": /\.png$/i.test(file) ? "image/png" : "image/jpeg", "cache-control": "private, max-age=600", "x-content-type-options": "nosniff" });
        fs.createReadStream(file).on("error", () => res.destroy()).pipe(res);
        return;
      }
      default: return sendJson(res, 404, { error: "not_found" });
    }
  }

  /** `/session/autotest`, for the connected task pane. */
  async pane(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    const { home, runtime } = this.deps;
    try {
      if (req.method === "GET") {
        const run = url.searchParams.get("run");
        if (run) {
          const state = home.runs().state(run);
          return state ? sendJson(res, 200, { run: state }) : sendJson(res, 404, { error: "run_not_found" });
        }
        const kitSync = home.syncKit();
        return sendJson(res, 200, {
          kitSync, home: home.dir, settings: home.settings(), kitInstalled: home.kitInstalled, kitVersion: home.kitVersion,
          resultValues: home.resultValues(), activeWorkbook: home.activeWorkbook(), runs: home.runs().list(20),
          knowledge: home.knowledge().length, runtime: await runtime.check(), dashboardPath: "/tests", signIns: home.signInList(),
        });
      }
      if (req.method === "PUT") {
        const name = url.searchParams.get("name") ?? "";
        if (!name) throw new AutotestHomeError("Say the workbook's file name.");
        const file = home.saveWorkbook(name, await readBinaryBody(req, MAX_WORKBOOK_BYTES));
        runtime.buildCatalog();
        return sendJson(res, 200, { saved: file });
      }
      if (req.method !== "POST") return sendJson(res, 405, { error: "method_not_allowed" });
      const body = JSON.parse(await readBody(req)) as Record<string, unknown>;
      switch (body.action) {
        case "save-settings": return sendJson(res, 200, home.save(body.settings));
        case "choose-home": home.choose(typeof body.dir === "string" ? body.dir : ""); return sendJson(res, 200, { home: home.dir });
        case "install-kit": return sendJson(res, 200, { installed: home.installKit() });
        case "install-runtime": runtime.install(); return sendJson(res, 200, { started: true });
        default: throw new AutotestHomeError("Unknown action.");
      }
    } catch (error) {
      if (error instanceof AutotestHomeError || error instanceof SyntaxError || (error instanceof Error && /Save the settings first|too large|exceeds/.test(error.message))) {
        return sendJson(res, 400, { error: (error as Error).message });
      }
      throw error;
    }
  }
}
