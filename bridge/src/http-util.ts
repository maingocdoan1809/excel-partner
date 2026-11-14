import type http from "node:http";
import path from "node:path";
import { AttachmentError } from "./attachments.js";
import type { WorkbookSessions } from "./workbook-sessions.js";
import type { WorkbookSession } from "./session.js";

export const MAX_BODY_BYTES = 8 * 1024 * 1024;

export function sendJson(res: http.ServerResponse, status: number, payload: unknown): void {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "content-length": Buffer.byteLength(body),
    "cache-control": "no-store",
  });
  res.end(body);
}

export function bearerToken(header: string | undefined): string | null {
  if (!header) return null;
  const match = /^Bearer\s+(.+)$/i.exec(header.trim());
  return match ? match[1]!.trim() : null;
}

/** The connected task-pane session whose token the request carries, or null (the caller answers 401). */
export function connectedSession(sessions: WorkbookSessions, req: http.IncomingMessage): WorkbookSession | null {
  const session = sessions.getByToken(bearerToken(req.headers.authorization) ?? "");
  return session?.isConnected ? session : null;
}

export function readBody(req: http.IncomingMessage, limit = MAX_BODY_BYTES): Promise<string> {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks: Buffer[] = [];
    req.on("data", (chunk: Buffer) => {
      size += chunk.length;
      if (size > limit) {
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

export async function readBinaryBody(req: http.IncomingMessage, limit: number): Promise<Buffer> {
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

/** `x-file-name` style headers arrive URL-encoded; null when the encoding is broken (the caller answers 400). */
export function decodedHeader(req: http.IncomingMessage, name: string): string | null {
  try { return decodeURIComponent(String(req.headers[name] ?? "")); }
  catch { return null; }
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

export function contentType(filePath: string): string {
  return CONTENT_TYPES[path.extname(filePath).toLowerCase()] ?? "application/octet-stream";
}
