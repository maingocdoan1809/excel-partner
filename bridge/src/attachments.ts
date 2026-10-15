import path from "node:path";
import { randomUUID } from "node:crypto";

export const MAX_ATTACHMENT_BYTES = 30 * 1024 * 1024;

export interface SessionAttachment {
  id: string;
  name: string;
  kind: "workbook" | "image" | "file";
  mimeType: string;
  bytes: Buffer;
}

export class AttachmentError extends Error {}

/** Validate user-selected bytes before retaining them in a workbook session. */
export function createAttachment(name: string, bytes: Buffer): SessionAttachment {
  const safeName = path.basename(name.replace(/\\/g, "/")).trim();
  if (!safeName || safeName.length > 255 || /[\x00-\x1f<>]/.test(safeName) || bytes.length === 0) throw new AttachmentError("invalid or empty file");
  if (bytes.length > MAX_ATTACHMENT_BYTES) throw new AttachmentError("file exceeds the 30 MB limit");
  const ext = path.extname(safeName).toLowerCase();
  const zip = bytes.subarray(0, 2).toString() === "PK";
  if ([".xlsx", ".xlsm", ".xltx", ".xltm"].includes(ext) && zip) {
    return { id: randomUUID(), name: safeName, kind: "workbook", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", bytes };
  }
  const png = ext === ".png" && bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
  const jpeg = [".jpg", ".jpeg"].includes(ext) && bytes.subarray(0, 3).equals(Buffer.from([255, 216, 255]));
  const gif = ext === ".gif" && ["GIF87a", "GIF89a"].includes(bytes.subarray(0, 6).toString());
  const webp = ext === ".webp" && bytes.subarray(0, 4).toString() === "RIFF" && bytes.subarray(8, 12).toString() === "WEBP";
  if (png || jpeg || gif || webp) {
    if (bytes.length > 10 * 1024 * 1024) throw new AttachmentError("image exceeds the 10 MB limit");
    return { id: randomUUID(), name: safeName, kind: "image", mimeType: png ? "image/png" : jpeg ? "image/jpeg" : gif ? "image/gif" : "image/webp", bytes };
  }
  return { id: randomUUID(), name: safeName, kind: "file", mimeType: "application/octet-stream", bytes };
}

export function attachmentPath(id: string): string {
  return `attachment://${id}`;
}

export function attachmentId(pathOrId: string): string | null {
  const match = /^attachment:\/\/([0-9a-f-]{36})$/.exec(pathOrId);
  return match?.[1] ?? null;
}
