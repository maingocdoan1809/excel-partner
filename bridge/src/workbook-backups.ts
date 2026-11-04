import fs from "node:fs";
import path from "node:path";

export const MAX_BACKUP_BYTES = 100 * 1024 * 1024;
export const KEPT_BACKUPS = 10;

export interface BackupInfo {
  file: string;
  size: number;
  savedAt: string;
}

export class BackupError extends Error {}

/**
 * Whole-workbook copies taken just before the assistant does something the undo journal cannot reverse. They are
 * ordinary .xlsx files in one folder, so restoring is "open the copy". Only the newest few are kept.
 */
export class WorkbookBackups {
  constructor(readonly dir: string, private readonly keep = KEPT_BACKUPS, private readonly now: () => Date = () => new Date()) {}

  save(workbookName: string, bytes: Buffer): BackupInfo {
    if (bytes.length === 0) throw new BackupError("the workbook copy is empty");
    if (bytes.length > MAX_BACKUP_BYTES) throw new BackupError(`the workbook is larger than the ${MAX_BACKUP_BYTES / 1024 / 1024} MB backup limit`);
    // An .xlsx is a zip: "PK". Anything else is not a workbook copy and would only fill the folder.
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new BackupError("the data is not a workbook file");
    fs.mkdirSync(this.dir, { recursive: true });
    const stamp = this.now().toISOString().replace(/[-:]/g, "").replace(/\..*$/, "").replace("T", "-");
    let file = `${safeBaseName(workbookName)}-${stamp}.xlsx`;
    for (let n = 2; fs.existsSync(path.join(this.dir, file)); n++) file = `${safeBaseName(workbookName)}-${stamp}-${n}.xlsx`;
    fs.writeFileSync(path.join(this.dir, file), bytes, { flag: "wx" });
    this.prune();
    const stat = fs.statSync(path.join(this.dir, file));
    return { file, size: stat.size, savedAt: stat.mtime.toISOString() };
  }

  list(): BackupInfo[] {
    if (!fs.existsSync(this.dir)) return [];
    return fs
      .readdirSync(this.dir)
      .filter((name) => name.toLowerCase().endsWith(".xlsx"))
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

/** The workbook's own name made safe for a file name: no folders, no reserved characters, no extension. */
export function safeBaseName(name: string): string {
  const base = name.replace(/\.(xlsx|xlsm|xlsb|xls|xltx|csv)$/i, "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_").replace(/^\.+/, "").trim();
  return (base || "workbook").slice(0, 60);
}
