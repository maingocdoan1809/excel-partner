import type http from "node:http";
import { GO_TO_REQUEST, UNNAMED_WORKBOOK, sameWorkbookName } from "../../../protocol/src/index.js";
import { AttachmentError, MAX_ATTACHMENT_BYTES, attachmentPath, createAttachment } from "../attachments.js";
import { BackupError, type DocumentBackups, MAX_BACKUP_BYTES, backupExtensionFor } from "../document-backups.js";
import { connectedSession, decodedHeader, readBinaryBody, readBody, sendJson } from "../http-util.js";
import type { Logger } from "../logger.js";
import type { SessionManager } from "../session-manager.js";

export interface FileRouteDeps {
  sessions: SessionManager;
  backups: DocumentBackups;
  logger: Logger;
  /** Brings a workbook's window to the front; Excel on Windows only (replaced in tests). */
  activateWorkbook: (workbookName: string) => Promise<unknown>;
}

/** Files that travel between the pane and the bridge: attachments, backups and links into other open files. */
export class FileRoutes {
  constructor(private readonly deps: FileRouteDeps) {}

  async attachment(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    const name = decodedHeader(req, "x-file-name");
    if (name === null) return sendJson(res, 400, { error: "invalid_file_name" });
    try {
      const bytes = await readBinaryBody(req, MAX_ATTACHMENT_BYTES);
      const attachment = createAttachment(name, bytes);
      if (session.attachments.size >= 10) session.attachments.delete(session.attachments.keys().next().value!);
      session.attachments.set(attachment.id, attachment);
      session.touch();
      return sendJson(res, 200, { id: attachment.id, name: attachment.name, kind: attachment.kind, path: attachmentPath(attachment.id) });
    } catch (error) {
      if (error instanceof AttachmentError) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  /** A copy of the open file (workbook, document or presentation), taken by the task pane before something the undo journal cannot reverse. */
  async backup(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    const name = decodedHeader(req, "x-file-name");
    if (name === null) return sendJson(res, 400, { error: "invalid_file_name" });
    try {
      const bytes = await readBinaryBody(req, MAX_BACKUP_BYTES);
      session.touch();
      return sendJson(res, 200, { ...this.deps.backups.save(name, bytes, backupExtensionFor(session.host)), folder: this.deps.backups.dir });
    } catch (error) {
      if (error instanceof BackupError || error instanceof AttachmentError) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  /**
   * A link in one file's chat that points into another open file: ask that file's pane to show the range, then bring its
   * window forward. The file is found by name (the chat only knows what the text said); the answer is always 200 with `ok`,
   * and `error` says why not (not open, two files with that name, no such sheet) in words the person can read.
   */
  async goTo(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    let body: Record<string, unknown>;
    try { body = JSON.parse(await readBody(req, 16 * 1024)) as Record<string, unknown>; }
    catch { return sendJson(res, 400, { error: "invalid_json" }); }
    const file = typeof body.file === "string" ? body.file.trim() : "";
    const address = typeof body.address === "string" ? body.address.trim() : "";
    const sheet = typeof body.sheet === "string" ? body.sheet.trim() : "";
    if (!file || file.length > 255 || !/^[A-Za-z0-9$:]{1,64}$/.test(address) || sheet.length > 255 || (body.sheet !== undefined && typeof body.sheet !== "string")) {
      return sendJson(res, 400, { error: "invalid_target" });
    }
    session.touch();
    await this.deps.sessions.refreshNames(); // a rename (Save As) has no event: names are read when someone needs them
    const matches = this.deps.sessions.connected().filter((item) => item.host === "excel" && item.id !== session.id && sameWorkbookName(item.workbookName ?? "", file));
    if (matches.length === 0) return sendJson(res, 200, { ok: false, error: `"${file}" is not open in Excel with this add-in.` });
    if (matches.length > 1) return sendJson(res, 200, { ok: false, error: `More than one open file is called "${file}", so it is not clear which one to show.` });
    const target = matches[0]!;
    const result = await target.requestPeerRead(GO_TO_REQUEST, { ...(sheet ? { sheet } : {}), address }, 15_000);
    if (!result.ok) return sendJson(res, 200, { ok: false, error: `[${target.workbookName ?? UNNAMED_WORKBOOK}] ${result.error.message}` });
    try { await this.deps.activateWorkbook(target.workbookName ?? ""); }
    catch (error) { this.deps.logger.warn("could not bring the workbook window forward", { error: String(error) }); }
    return sendJson(res, 200, { ok: true });
  }
}
