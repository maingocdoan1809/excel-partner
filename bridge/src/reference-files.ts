import fs from "node:fs";
import path from "node:path";

export class ReferenceError_ extends Error {
  constructor(message: string, readonly reason: "outside" | "other" = "other") { super(message); }
}

export interface FileEntry {
  name: string;
  type: "file" | "folder";
  size: number | null;
  modified: string;
}

const WORKBOOK_EXT = new Set([".xlsx", ".xlsm", ".xltx", ".xltm"]);
const MAX_WORKBOOK_BYTES = 30 * 1024 * 1024;
const MAX_LISTING = 200;

/**
 * `C:\\Users\\me\\x.xlsx` or `\\\\wsl.localhost\\Ubuntu\\home\\me\\x.xlsx` -> the path the bridge (WSL / Linux) can open.
 * Other paths are returned unchanged.
 */
export function toLocalPath(input: string): string {
  const trimmed = input.trim().replace(/^"|"$/g, "");
  const drive = /^([A-Za-z]):[\\/](.*)$/.exec(trimmed);
  if (drive && process.platform !== "win32") return `/mnt/${drive[1]!.toLowerCase()}/${drive[2]!.replace(/\\/g, "/")}`;
  const wsl = /^\\\\(?:wsl\.localhost|wsl\$)\\[^\\]+\\(.*)$/i.exec(trimmed);
  if (wsl && process.platform !== "win32") return `/${wsl[1]!.replace(/\\/g, "/")}`;
  return trimmed;
}

/** Local folder containing the workbook currently open in Excel. Cloud URLs have no bridge-local folder. */
export function workbookFolderFromUrl(documentUrl: string): string {
  let input = documentUrl.trim();
  if (!input) throw new ReferenceError_("Save this workbook first, or add a workspace folder manually.");
  if (/^https?:\/\//i.test(input)) throw new ReferenceError_("This workbook is opened from a cloud URL. Add its local folder manually.");
  if (/^file:\/\//i.test(input)) {
    const url = new URL(input);
    if (url.hostname && url.hostname !== "localhost") throw new ReferenceError_("Network file URLs cannot be mapped to a local workspace. Add the folder manually.");
    input = decodeURIComponent(url.pathname).replace(/^\/(?=[A-Za-z]:\/)/, "");
  }
  let file: string;
  try { file = fs.realpathSync(path.resolve(toLocalPath(input))); }
  catch { throw new ReferenceError_("The open workbook path is unavailable to the bridge. Add its folder manually."); }
  if (!fs.statSync(file).isFile()) throw new ReferenceError_("The open workbook path is not a file.");
  return path.dirname(file);
}

/**
 * Read-only access to files the user explicitly opened up: only paths inside the configured folders can be
 * listed or read, symlinks cannot escape them, and only workbook files can be read.
 */
export class ReferenceFiles {
  private readonly roots: string[];
  private savedPaths: () => string[] = () => [];

  setSavedPaths(provider: () => string[]): void { this.savedPaths = provider; }

  constructor(dirs: string[]) {
    this.roots = dirs
      .map((d) => toLocalPath(d))
      .map((d) => {
        try {
          return fs.realpathSync(d);
        } catch {
          return null;
        }
      })
      .filter((d): d is string => d !== null);
  }

  get configured(): boolean {
    return this.roots.length > 0;
  }

  get allowedFolders(): string[] {
    return [...this.roots];
  }

  addRoot(input: string): string {
    let real: string;
    try { real = fs.realpathSync(path.resolve(toLocalPath(input))); }
    catch { throw new ReferenceError_(`'${input}' does not exist or is not readable`); }
    if (!fs.statSync(real).isDirectory()) throw new ReferenceError_(`'${input}' is not a folder`);
    if (!this.roots.includes(real)) this.roots.push(real);
    return real;
  }

  removeRoot(input: string): boolean {
    const root = this.roots.indexOf(input);
    if (root < 0) return false;
    this.roots.splice(root, 1);
    return true;
  }

  /** Resolve a user-supplied path to a real path inside an allowed folder, or throw with guidance. */
  resolve(input: string, grants: ReadonlySet<string> = new Set()): string {
    const candidate = path.resolve(toLocalPath(input));
    let real: string;
    try {
      real = fs.realpathSync(candidate);
    } catch {
      throw new ReferenceError_(`'${input}' does not exist or is not readable`);
    }
    if (!this.savedPaths().includes(real) && ![...this.roots, ...grants].some((root) => real === root || real.startsWith(root + path.sep))) {
      throw new ReferenceError_(`'${input}' is outside the folders the assistant may read: ${this.roots.join(", ") || "none"}`, "outside");
    }
    return real;
  }

  list(input?: string, grants: ReadonlySet<string> = new Set()): { path: string; entries: FileEntry[]; truncated: boolean } {
    if (!input) {
      return {
        path: "(allowed folders)",
        truncated: false,
        entries: [...this.roots, ...grants].filter((r) => fs.statSync(r).isDirectory()).map((r) => ({ name: r, type: "folder" as const, size: null, modified: fs.statSync(r).mtime.toISOString() })),
      };
    }
    const dir = this.resolve(input, grants);
    if (!fs.statSync(dir).isDirectory()) throw new ReferenceError_(`'${input}' is not a folder`);
    const all = fs.readdirSync(dir, { withFileTypes: true })
      .filter((e) => !e.name.startsWith(".") && !e.name.startsWith("~$") && !e.isSymbolicLink())
      .sort((a, b) => (a.isDirectory() === b.isDirectory() ? a.name.localeCompare(b.name) : a.isDirectory() ? -1 : 1));
    const entries = all.slice(0, MAX_LISTING).map((e): FileEntry => {
      const stat = fs.statSync(path.join(dir, e.name));
      return { name: e.name, type: stat.isDirectory() ? "folder" : "file", size: stat.isFile() ? stat.size : null, modified: stat.mtime.toISOString() };
    });
    entries.sort((a, b) => (a.type === b.type ? a.name.localeCompare(b.name) : a.type === "folder" ? -1 : 1));
    return { path: dir, entries, truncated: all.length > MAX_LISTING };
  }

  /** Bytes of an allowed workbook file. */
  readWorkbook(input: string, grants: ReadonlySet<string> = new Set()): { path: string; bytes: Uint8Array } {
    const file = this.resolve(input, grants);
    if (!fs.statSync(file).isFile()) throw new ReferenceError_(`'${input}' is not a file`);
    if (!WORKBOOK_EXT.has(path.extname(file).toLowerCase())) {
      throw new ReferenceError_("only Excel workbooks (.xlsx, .xlsm, .xltx, .xltm) can be read");
    }
    const size = fs.statSync(file).size;
    if (size > MAX_WORKBOOK_BYTES) throw new ReferenceError_(`the file is ${Math.round(size / 1e6)} MB; the limit is ${MAX_WORKBOOK_BYTES / 1e6} MB`);
    return { path: file, bytes: new Uint8Array(fs.readFileSync(file)) };
  }

  /** Read a selected workbook or image from a workspace, with the same limits as a picker upload. */
  readSelectedFile(input: string, grants: ReadonlySet<string> = new Set()): { path: string; bytes: Buffer } {
    const file = this.resolve(input, grants);
    const stat = fs.statSync(file);
    if (!stat.isFile()) throw new ReferenceError_(`'${input}' is not a file`);
    if (stat.size > MAX_WORKBOOK_BYTES) throw new ReferenceError_("file exceeds the 30 MB limit");
    return { path: file, bytes: fs.readFileSync(file) };
  }

  /** Check a selected file without copying its content into the session. */
  selectedPath(input: string, grants: ReadonlySet<string> = new Set()): string {
    const file = this.resolve(input, grants);
    if (!fs.statSync(file).isFile()) throw new ReferenceError_(`'${input}' is not a file`);
    return file;
  }

  /** Canonical, validated target shown to the user before a one-session grant. */
  permissionTarget(input: string): string {
    let real: string;
    try { real = fs.realpathSync(path.resolve(toLocalPath(input))); }
    catch { throw new ReferenceError_(`'${input}' does not exist or is not readable`); }
    const stat = fs.statSync(real);
    if (!stat.isFile() || !WORKBOOK_EXT.has(path.extname(real).toLowerCase())) throw new ReferenceError_("only Excel workbooks can be read");
    if (stat.size > MAX_WORKBOOK_BYTES) throw new ReferenceError_("workbook exceeds the 30 MB limit");
    return real;
  }

  /** Exact file selected through the native picker, regardless of extension. */
  permissionTargetAny(input: string): string {
    let real: string;
    try { real = fs.realpathSync(path.resolve(toLocalPath(input))); }
    catch { throw new ReferenceError_(`'${input}' does not exist or is not readable`); }
    if (!fs.statSync(real).isFile()) throw new ReferenceError_(`'${input}' is not a file`);
    return real;
  }
}
