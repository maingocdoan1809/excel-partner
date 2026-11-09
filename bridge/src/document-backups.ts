import fs from "node:fs";
import path from "node:path";
import type { HostId } from "../../protocol/src/index.js";

export const MAX_BACKUP_BYTES = 100 * 1024 * 1024;
export const KEPT_BACKUPS = 10;

/** The file type each Office application saves, and what that file is called in messages. */
export const BACKUP_EXTENSIONS = [".xlsx", ".docx", ".pptx"] as const;
export type BackupExtension = (typeof BACKUP_EXTENSIONS)[number];
const BACKUP_NOUN: Record<BackupExtension, string> = { ".xlsx": "workbook", ".docx": "document", ".pptx": "presentation" };

/** The backup file type for an Office host. */
export function backupExtensionFor(host: HostId): BackupExtension {
  return host === "word" ? ".docx" : host === "powerpoint" ? ".pptx" : ".xlsx";
}

export interface BackupInfo {
  file: string;
  size: number;
  savedAt: string;
}

export class BackupError extends Error {}

/**
 * Whole-file copies (workbook, document or presentation) taken just before the assistant does something the undo
 * journal cannot reverse. They are ordinary .xlsx / .docx / .pptx files in one folder, so restoring is "open the
 * copy". Only the newest few are kept, across all three kinds.
 */
export class DocumentBackups {
  constructor(readonly dir: string, private readonly keep = KEPT_BACKUPS, private readonly now: () => Date = () => new Date()) {}

  save(name: string, bytes: Buffer, extension: BackupExtension = ".xlsx"): BackupInfo {
    const noun = BACKUP_NOUN[extension];
    if (bytes.length === 0) throw new BackupError(`the ${noun} copy is empty`);
    if (bytes.length > MAX_BACKUP_BYTES) throw new BackupError(`the ${noun} is larger than the ${MAX_BACKUP_BYTES / 1024 / 1024} MB backup limit`);
    // .xlsx, .docx and .pptx are zips: "PK". Anything else is not a copy of the file and would only fill the folder.
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new BackupError(`the data is not a ${noun} file`);
    fs.mkdirSync(this.dir, { recursive: true });
    const stamp = this.now().toISOString().replace(/[-:]/g, "").replace(/\..*$/, "").replace("T", "-");
    const base = safeBaseName(name, noun === "workbook" ? "workbook" : noun);
    let file = `${base}-${stamp}${extension}`;
    for (let n = 2; fs.existsSync(path.join(this.dir, file)); n++) file = `${base}-${stamp}-${n}${extension}`;
    fs.writeFileSync(path.join(this.dir, file), bytes, { flag: "wx" });
    this.prune();
    const stat = fs.statSync(path.join(this.dir, file));
    return { file, size: stat.size, savedAt: stat.mtime.toISOString() };
  }

  list(): BackupInfo[] {
    if (!fs.existsSync(this.dir)) return [];
    return fs
      .readdirSync(this.dir)
      .filter((name) => BACKUP_EXTENSIONS.some((extension) => name.toLowerCase().endsWith(extension)))
      .map((name) => {
        const stat = fs.statSync(path.join(this.dir, name));
        return { file: name, size: stat.size, savedAt: stat.mtime.toISOString() };
      })
      .sort((a, b) => b.savedAt.localeCompare(a.savedAt) || b.file.localeCompare(a.file));
  }

  private prune(): void {
    for (const old of this.list().slice(this.keep)) fs.rmSync(path.join(this.dir, old.file), { force: true });
  }
}

/** The file's own name made safe for a file name: no folders, no reserved characters, no extension. */
export function safeBaseName(name: string, fallback = "workbook"): string {
  const base = name.replace(/\.(xlsx|xlsm|xlsb|xls|xltx|csv|docx|docm|doc|dotx|dotm|pptx|pptm|ppt|potx|potm|ppsx)$/i, "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_").replace(/^\.+/, "").trim();
  return (base || fallback).slice(0, 60);
}
