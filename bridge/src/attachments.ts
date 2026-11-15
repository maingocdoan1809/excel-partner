import path from "node:path";
import { randomUUID } from "node:crypto";
import { IMAGE_EXTENSIONS, detectImage } from "./image-info.js";

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
  const image = detectImage(bytes);
  if (image && IMAGE_EXTENSIONS[image.mimeType].includes(ext)) {
    if (bytes.length > 10 * 1024 * 1024) throw new AttachmentError("image exceeds the 10 MB limit");
    return { id: randomUUID(), name: safeName, kind: "image", mimeType: image.mimeType, bytes };
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
