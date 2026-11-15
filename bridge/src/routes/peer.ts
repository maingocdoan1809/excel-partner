import type http from "node:http";
import fs from "node:fs";
import { randomBytes, timingSafeEqual } from "node:crypto";
import { MAX_DELEGATED_TEXT, UNNAMED_FILE, type ToolResult } from "../../../protocol/src/index.js";
import { terminalAgents } from "../../../providers/src/registry.js";
import { terminalAgentLabel } from "../../../providers/src/terminal-agent.js";
import type { Delegations } from "../delegations.js";
import { toolError as fail } from "../tool-error.js";
import { readBody, sendJson } from "../http-util.js";
import type { WorkbookSessions } from "../workbook-sessions.js";

export interface PeerRouteDeps {
  sessions: WorkbookSessions;
  delegations: Delegations;
  /** Where the shared secret lives; created on first start, readable by this user only. */
  tokenFile: string;
}

const PEER_HEADER = "x-office-partner-peer";

/**
 * `/peer/*`: a program on this computer (a terminal agent) talking to the assistant of an open file. It rides on the
 * delegation that already lets one open file's assistant ask another's: the request becomes a normal turn in that file's
 * chat and the final reply comes back. The caller is a peer, not a pane, so it has no session token; it proves itself with
 * the secret in `peer-token.txt` (the bridge already refuses connections from other users' processes).
 */
export class PeerRoutes {
  private readonly token: string;

  constructor(private readonly deps: PeerRouteDeps) {
    this.token = readOrCreateToken(deps.tokenFile);
  }

  async handle(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (!this.authorized(req)) return sendJson(res, 401, { error: "unauthorized" });
    if (req.method === "GET" && url.pathname === "/peer/files") return sendJson(res, 200, { files: await this.files() });
    if (req.method === "POST" && url.pathname === "/peer/ask") return this.ask(req, res);
    return sendJson(res, 404, { error: "not_found" });
  }

  private authorized(req: http.IncomingMessage): boolean {
    const supplied = req.headers[PEER_HEADER];
    return typeof supplied === "string" && Buffer.byteLength(supplied) === Buffer.byteLength(this.token)
      && timingSafeEqual(Buffer.from(supplied), Buffer.from(this.token));
  }

  private async files() {
    const open = this.deps.sessions.connected().filter(session => session.clientInfo?.standalone !== true);
    await this.deps.sessions.refreshNames(open);
    return open.map(session => ({ id: session.fileId, name: session.workbookName ?? UNNAMED_FILE[session.host], host: session.host, ready: Boolean(session.provider), busy: session.isBusy }));
  }

  private async ask(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    let body: { file?: unknown; message?: unknown; client?: unknown };
    try { body = JSON.parse(await readBody(req, 64_000)); }
    catch { return sendJson(res, 400, { error: "invalid_json" }); }
    if (typeof body.file !== "string" || typeof body.message !== "string" || !body.message.trim()) return sendJson(res, 400, { error: "file and message are required" });
    const from = terminalAgentLabel(terminalAgents(), typeof body.client === "string" ? body.client : undefined);
    // A caller that hangs up must not leave the other file working for nobody.
    const callerGone = new AbortController();
    res.on("close", () => { if (!res.writableFinished) callerGone.abort(); });
    return sendJson(res, 200, await this.send(body.file, from, body.message, callerGone.signal));
  }

  private async send(file: string, from: string, message: string, signal: AbortSignal): Promise<ToolResult> {
    if (message.length > MAX_DELEGATED_TEXT) return fail("invalid_arguments", `The message is longer than ${MAX_DELEGATED_TEXT} characters; send a shorter, self-contained request.`);
    const target = this.deps.sessions.connected().find(session => session.fileId === file && session.clientInfo?.standalone !== true);
    if (!target) return fail("invalid_arguments", `No open file '${file}'. List the files again: a closed or reopened file gets a new id.`);
    await this.deps.sessions.refreshNames([target]);
    const name = target.workbookName ?? UNNAMED_FILE[target.host];
    if (!target.provider) return fail("unsupported", `"${name}" has no AI provider selected yet. Ask the user to choose one in that window.`);
    if (target.isBusy) return fail("unsupported", `"${name}" is busy with its own request. Try again when it has finished.`);
    // The person at that window sees who is asking and every character of the request, as for one file asking another.
    const answer = await target.requestTool("__authorize_connected_file", { fileId: target.fileId, name, host: target.host, action: "message", text: message, from }, 120_000);
    if (!answer.ok) return answer;
    if ((answer.data as { approved?: unknown })?.approved !== true) return fail("cancelled", `The user declined the request to "${name}".`);
    if (signal.aborted) return fail("cancelled", "The caller gave up before the request was sent.");
    if (!target.isConnected) return fail("disconnected", `"${name}" was closed. List the files again.`);
    if (target.isBusy) return fail("unsupported", `"${name}" became busy with its own request. Try again when it has finished.`);
    const result = await this.deps.delegations.run(target, { id: "terminal", name: from }, message, signal);
    return result.ok ? { ok: true, data: { file: { id: target.fileId, name, host: target.host }, ...(result.data as Record<string, unknown>) } }
      : { ...result, error: { ...result.error, message: `[${name}] ${result.error.message}` } };
  }
}

function readOrCreateToken(file: string): string {
  try { return fs.readFileSync(file, "utf8").trim(); } catch { /* first start */ }
  const token = randomBytes(32).toString("hex");
  try { fs.writeFileSync(file, token, { mode: 0o600, flag: "wx" }); return token; }
  catch { return fs.readFileSync(file, "utf8").trim(); } // another bridge process won the race
}
