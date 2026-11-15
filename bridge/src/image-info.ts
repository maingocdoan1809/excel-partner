export type ImageMime = "image/png" | "image/jpeg" | "image/gif" | "image/webp";

export interface ImageInfo {
  mimeType: ImageMime;
  /** Pixel size; absent when the header does not give it (WebP, truncated files). */
  width?: number;
  height?: number;
}

/** File extensions each type is stored under. */
export const IMAGE_EXTENSIONS: Record<ImageMime, readonly string[]> = {
  "image/png": [".png"],
  "image/jpeg": [".jpg", ".jpeg"],
  "image/gif": [".gif"],
  "image/webp": [".webp"],
};

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

/** JPEG frame headers that carry the picture size (every SOFn marker except DHT, JPG and DAC). */
function isJpegFrame(marker: number): boolean {
  return marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
}

function jpegSize(bytes: Uint8Array): { width: number; height: number } | null {
  const view = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let i = 2;
  while (i + 9 < view.length) {
    if (view[i] !== 0xff) { i++; continue; }
    const marker = view[i + 1]!;
    if (marker === 0xff) { i++; continue; }
    if (isJpegFrame(marker)) return { height: view.readUInt16BE(i + 5), width: view.readUInt16BE(i + 7) };
    // Markers without a length: stuffed byte, SOI, TEM, RSTn.
    if (marker === 0x00 || marker === 0x01 || marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7)) { i += 2; continue; }
    i += 2 + view.readUInt16BE(i + 2);
  }
  return null;
}

/** What a file is, from its first bytes. The name and the declared type are never trusted. */
export function detectImage(bytes: Uint8Array): ImageInfo | null {
  const view = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.length >= 8 && view.subarray(0, 8).equals(PNG_SIGNATURE)) {
    const header = view.length >= 24 && view.subarray(12, 16).toString("latin1") === "IHDR";
    return header ? { mimeType: "image/png", width: view.readUInt32BE(16), height: view.readUInt32BE(20) } : { mimeType: "image/png" };
  }
  if (view.length >= 3 && view[0] === 0xff && view[1] === 0xd8 && view[2] === 0xff) {
    const size = jpegSize(view);
    return size ? { mimeType: "image/jpeg", ...size } : { mimeType: "image/jpeg" };
  }
  if (view.length >= 10 && ["GIF87a", "GIF89a"].includes(view.subarray(0, 6).toString("latin1"))) {
    return { mimeType: "image/gif", width: view.readUInt16LE(6), height: view.readUInt16LE(8) };
  }
  if (view.length >= 12 && view.subarray(0, 4).toString("latin1") === "RIFF" && view.subarray(8, 12).toString("latin1") === "WEBP") {
    return { mimeType: "image/webp" };
  }
  return null;
}
