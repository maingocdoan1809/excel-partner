import type http from "node:http";
import path from "node:path";
import { readFileSync, renameSync, writeFileSync } from "node:fs";
import { AttachmentError } from "../attachments.js";
import { connectedSession, readBody, sendJson } from "../http-util.js";
import type { Logger } from "../logger.js";
import { ReferenceError_, type ReferenceFiles, workbookFolderFromUrl } from "../reference-files.js";
import type { SavedReferences } from "../saved-references.js";
import type { WorkbookSessions } from "../workbook-sessions.js";
import { activeWorkbookPath as findActiveWorkbookPath, pickWorkspaceFolder as pickFolderDialog } from "../windows-workspace.js";

export interface WorkspaceRouteDeps {
  sessions: WorkbookSessions;
  references: ReferenceFiles;
  savedReferences: SavedReferences;
  logger: Logger;
  /** Where user-added workspace roots are saved; null disables persistence (tests). */
  workspaceConfigPath: string | null;
  activeWorkbookPath?: (workbookName: string) => Promise<string | null>;
  pickWorkspaceFolder?: () => Promise<string | null>;
}

/** `/session/workspaces*` and `/session/references`: the folders and files the assistant may read. */
export class WorkspaceRoutes {
  constructor(private readonly deps: WorkspaceRouteDeps) {}

  async list(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    let defaultError: string | null = null;
    const documentUrl = url.searchParams.get("documentUrl") ?? "";
    const workbookName = url.searchParams.get("workbookName") ?? "";
    let folder: string | null = null;
    try { folder = workbookFolderFromUrl(documentUrl); }
    catch (error) { defaultError = error instanceof Error ? error.message : String(error); }
    if (!folder && workbookName && workbookName.length <= 255) {
      try {
        const active = await (this.deps.activeWorkbookPath ?? findActiveWorkbookPath)(workbookName);
        if (active) { folder = workbookFolderFromUrl(active); defaultError = null; }
      } catch (error) {
        this.deps.logger.warn("active workbook path lookup failed", { error: String(error) });
      }
    }
    if (folder) {
      if (session.defaultWorkspace && session.defaultWorkspace !== folder) session.referenceGrants.delete(session.defaultWorkspace);
      session.defaultWorkspace = folder;
      session.referenceGrants.add(folder);
    }
    const roots = [...new Set([...(session.defaultWorkspace ? [session.defaultWorkspace] : []), ...this.deps.references.allowedFolders])];
    return sendJson(res, 200, { roots: roots.map((folder) => ({
      path: folder,
      name: path.basename(folder),
      isDefault: folder === session.defaultWorkspace,
      removable: folder !== session.defaultWorkspace && this.deps.references.allowedFolders.includes(folder),
    })), defaultError });
  }

  async add(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    let body: { path?: unknown };
    try { body = JSON.parse(await readBody(req)); }
    catch { return sendJson(res, 400, { error: "invalid_json" }); }
    if (typeof body.path !== "string" || !body.path.trim() || body.path.length > 1024) return sendJson(res, 400, { error: "invalid_path" });
    try {
      const folder = this.deps.references.addRoot(body.path);
      this.saveRoots();
      return sendJson(res, 200, { path: folder, name: path.basename(folder) });
    } catch (error) {
      if (error instanceof ReferenceError_) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  async remove(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    let body: { path?: unknown };
    try { body = JSON.parse(await readBody(req)); }
    catch { return sendJson(res, 400, { error: "invalid_json" }); }
    if (typeof body.path !== "string" || body.path === session.defaultWorkspace || !this.deps.references.removeRoot(body.path)) {
      return sendJson(res, 400, { error: "This workspace cannot be removed." });
    }
    this.saveRoots();
    return sendJson(res, 200, { removed: true });
  }

  async pick(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    let selected: string | null;
    try { selected = await (this.deps.pickWorkspaceFolder ?? pickFolderDialog)(); }
    catch (error) {
      this.deps.logger.warn("folder picker failed", { error: String(error) });
      return sendJson(res, 500, { error: "Could not open the Windows folder picker." });
    }
    if (!selected) return sendJson(res, 200, { cancelled: true });
    try {
      const folder = this.deps.references.addRoot(selected);
      this.saveRoots();
      return sendJson(res, 200, { path: folder, name: path.basename(folder), cancelled: false });
    } catch (error) {
      if (error instanceof ReferenceError_) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  private saveRoots(): void {
    const file = this.deps.workspaceConfigPath;
    if (!file) return;
    let current: Record<string, unknown> = {};
    try { current = JSON.parse(readFileSync(file, "utf8")) as Record<string, unknown>; }
    catch { /* first workspace */ }
    const temporary = `${file}.tmp`;
    writeFileSync(temporary, JSON.stringify({ ...current, referenceDirs: this.deps.references.allowedFolders }, null, 2) + "\n", { mode: 0o600 });
    renameSync(temporary, file);
  }

  entries(req: http.IncomingMessage, res: http.ServerResponse, url: URL): void {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    const folder = url.searchParams.get("path");
    if (!folder) return sendJson(res, 400, { error: "missing_path" });
    try {
      const listing = this.deps.references.list(folder, session.referenceGrants);
      return sendJson(res, 200, {
        ...listing,
        entries: listing.entries.map((entry) => ({
          ...entry,
          path: path.join(listing.path, entry.name),
          supported: true,
        })),
      });
    } catch (error) {
      if (error instanceof ReferenceError_) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  async attach(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    let body: { path?: unknown };
    try { body = JSON.parse(await readBody(req)); }
    catch { return sendJson(res, 400, { error: "invalid_json" }); }
    if (typeof body.path !== "string" || !body.path.trim()) return sendJson(res, 400, { error: "invalid_path" });
    try {
      const selected = this.deps.references.selectedPath(body.path, session.referenceGrants);
      session.touch();
      return sendJson(res, 200, { name: path.basename(selected), kind: "file", path: selected });
    } catch (error) {
      if (error instanceof ReferenceError_ || error instanceof AttachmentError) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  /** Saved references: named files or folders the assistant is told about in every prompt. */
  async savedReferences(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    const saved = this.deps.savedReferences;
    try {
      if (req.method === "GET") return sendJson(res, 200, { references: saved.list() });
      if (req.method !== "POST" && req.method !== "DELETE") return sendJson(res, 405, { error: "method_not_allowed" });
      const body = JSON.parse(await readBody(req)) as Record<string, unknown>;
      if (!body || typeof body !== "object") throw new ReferenceError_("Invalid reference metadata.");
      if (req.method === "DELETE") {
        if (typeof body.id !== "string") throw new ReferenceError_("Specify a reference ID.");
        saved.remove(body.id);
        return sendJson(res, 200, { removed: true });
      }
      if (typeof body.name !== "string" || typeof body.description !== "string" || (body.id !== undefined && typeof body.id !== "string")) throw new ReferenceError_("Specify a name and description.");
      const existing = saved.list().find(e => e.id === body.id);
      let target = existing ? { path: existing.path, kind: existing.kind } : undefined;
      if (body.path !== undefined) {
        if (typeof body.path !== "string") throw new ReferenceError_("Enter a valid file or folder path.");
        // Explorer's "Copy as path" wraps the path in double quotes.
        const typed = body.path.trim().replace(/^"(.*)"$/, "$1").trim();
        if (!typed) throw new ReferenceError_("Enter the full path of the file or folder.");
        target = this.deps.references.referenceTarget(typed);
      }
      if (!target) throw new ReferenceError_("Enter the full path of the file or folder.");
      const reference = saved.save({ id: body.id as string | undefined, name: body.name, description: body.description, path: target.path, kind: target.kind });
      return sendJson(res, 200, { reference });
    } catch (error) {
      if (error instanceof ReferenceError_ || error instanceof SyntaxError) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }
}
