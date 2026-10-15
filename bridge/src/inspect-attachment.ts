import { unzipSync } from "fflate";
import { AttachmentError } from "./attachments.js";

interface ReadOptions { entry?: string; offset: number; maxBytes: number }

/** Read bounded chunks; ZIP entries are listed first and decompressed only on request. */
export function inspectAttachment(file: { name: string; bytes: Buffer }, options: ReadOptions): unknown {
  let bytes: Uint8Array = file.bytes;
  if (file.bytes[0] === 0x50 && file.bytes[1] === 0x4b) {
    if (!options.entry) {
      const entries: Array<{ name: string; size: number }> = [];
      try {
        unzipSync(file.bytes, { filter: (item) => {
          if (entries.length < 1000) entries.push({ name: item.name, size: item.originalSize });
          return false;
        } });
      } catch { throw new AttachmentError("invalid ZIP archive"); }
      return { name: file.name, format: "zip", entries, truncated: entries.length === 1000 };
    }
    let found = false;
    let oversized = false;
    try {
      const extracted = unzipSync(file.bytes, { filter: (item) => {
        if (item.name !== options.entry) return false;
        found = true;
        oversized = item.originalSize > 2 * 1024 * 1024;
        return !oversized;
      } });
      if (oversized) throw new AttachmentError("ZIP entry exceeds the 2 MB read limit");
      const selected = extracted[options.entry];
      if (!found || !selected) throw new AttachmentError("ZIP entry was not found");
      bytes = selected;
    } catch (error) {
      if (error instanceof AttachmentError) throw error;
      throw new AttachmentError("invalid or unsupported ZIP entry");
    }
  } else if (options.entry) throw new AttachmentError("entry is only valid for ZIP files");

  const chunk = bytes.subarray(options.offset, options.offset + options.maxBytes);
  let text: string | null = null;
  if (!chunk.includes(0)) {
    try { text = new TextDecoder("utf-8", { fatal: true }).decode(chunk); }
    catch { /* binary or a chunk that ends inside a UTF-8 sequence */ }
  }
  return {
    name: file.name,
    ...(options.entry ? { entry: options.entry } : {}),
    size: bytes.length,
    offset: options.offset,
    nextOffset: Math.min(bytes.length, options.offset + chunk.length),
    truncated: options.offset + chunk.length < bytes.length,
    encoding: text === null ? "base64" : "utf-8",
    content: text ?? Buffer.from(chunk).toString("base64"),
  };
}
