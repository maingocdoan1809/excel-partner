import { MAX_INSERT_IMAGE_BYTES, INSERT_IMAGE_MIME_TYPES, type InsertImageMimeType, type ResolvedImage } from "../../protocol/src/index.js";
import { detectImage } from "./image-info.js";

/** A source the assistant gave that cannot be turned into a picture. The message is for the assistant: it says what to change. */
export class ImageSourceError extends Error {}

/** What `insert_image` needs from the bridge to find a picture. Each member does its own permission asking. */
export interface ImageSources {
  /** A picture the user attached in this chat. */
  attachment(id: string): { name: string; kind: string; bytes: Buffer } | undefined;
  /** A picture saved with a skill; null when this bridge has no skills. */
  template(uri: string): { name: string; kind: string; bytes: Buffer } | null;
  /** A file on this computer, after the user allowed reading it. */
  localFile(path: string): Promise<{ name: string; bytes: Buffer }>;
  /** A picture on the internet, after the user allowed the download. */
  download(url: string): Promise<{ bytes: Buffer; finalUrl: string }>;
}

const ATTACHMENT = /^attachment:\/\/([0-9a-f-]{36})$/;
const DATA_URI = /^data:([^,;]*)((?:;[^,;]*)*),([\s\S]*)$/i;
const BASE64_TEXT = /^[A-Za-z0-9+/_-]+={0,2}$/;

/** Plain base64 text, as a model writes a picture it made itself: long, no dots, no folder separators. */
export function looksLikeBase64(source: string): boolean {
  const text = source.replace(/\s+/g, "");
  return text.length >= 32 && BASE64_TEXT.test(text);
}

function decodeBase64(text: string): Buffer {
  const clean = text.replace(/\s+/g, "").replace(/-/g, "+").replace(/_/g, "/");
  if (clean.length === 0) throw new ImageSourceError("the picture data is empty");
  const bytes = Buffer.from(clean, "base64");
  if (bytes.length === 0) throw new ImageSourceError("the picture data is not valid base64");
  return bytes;
}

/** The label shown in results: never the data, and for an address only its host and file name. */
function urlLabel(address: string): string {
  try {
    const url = new URL(address);
    const file = url.pathname.split("/").filter(Boolean).at(-1);
    return file ? `${url.hostname}/${decodeURIComponent(file)}` : url.hostname;
  } catch { return "web address"; }
}

function checkedImage(bytes: Buffer, label: string, hint: string): ResolvedImage {
  if (bytes.length > MAX_INSERT_IMAGE_BYTES) throw new ImageSourceError(`the picture is ${bytes.length} bytes, over the ${MAX_INSERT_IMAGE_BYTES} byte limit; use a smaller one`);
  const info = detectImage(bytes);
  if (!info) throw new ImageSourceError(`${hint} is not a PNG, JPEG or GIF picture (its content is not an image of a type Office takes)`);
  if (!(INSERT_IMAGE_MIME_TYPES as readonly string[]).includes(info.mimeType)) {
    throw new ImageSourceError(`${hint} is a ${info.mimeType.replace("image/", "").toUpperCase()} picture, which Word, Excel and PowerPoint cannot insert; convert it to PNG first (for example with run_python)`);
  }
  return {
    imageBase64: bytes.toString("base64"),
    mimeType: info.mimeType as InsertImageMimeType,
    ...(info.width && info.height ? { naturalWidth: info.width, naturalHeight: info.height } : {}),
    sourceLabel: label,
  };
}

/** Read the picture `source` points at. Throws `ImageSourceError` (or the source's own error) when it cannot. */
export async function resolveImage(source: string, sources: ImageSources): Promise<ResolvedImage> {
  const text = source.trim();
  const data = DATA_URI.exec(text);
  if (data) {
    if (!/;base64/i.test(data[2]!)) throw new ImageSourceError("a data: address must be base64 encoded (data:image/png;base64,...)");
    return checkedImage(decodeBase64(data[3]!), "pasted data", "the data");
  }
  const attachment = ATTACHMENT.exec(text);
  if (attachment) {
    const file = sources.attachment(attachment[1]!);
    if (file?.kind !== "image") throw new ImageSourceError("the attached picture is unavailable; ask the user to attach it again");
    return checkedImage(file.bytes, file.name, `'${file.name}'`);
  }
  if (/^skill-template:\/\//i.test(text)) {
    const template = sources.template(text);
    if (!template) throw new ImageSourceError("skills are not configured");
    if (template.kind !== "image") throw new ImageSourceError("that skill file is not a picture");
    return checkedImage(template.bytes, template.name, `'${template.name}'`);
  }
  if (/^https?:\/\//i.test(text)) {
    const { bytes, finalUrl } = await sources.download(text);
    return checkedImage(bytes, urlLabel(finalUrl), "what the address returned");
  }
  if (looksLikeBase64(text)) return checkedImage(decodeBase64(text), "pasted data", "the data");
  const file = await sources.localFile(text);
  return checkedImage(file.bytes, file.name, `'${file.name}'`);
}
