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

/**
 * `\\server\share`, `//server/share`, `\\?\`, `\\.\` and WebDAV `\\host@port\...`: touching one makes Windows connect to that
 * host and offer the user's NTLM credentials, so such a path is refused by its text before any file system call.
 */
export function isNetworkPath(input: string): boolean {
  return /^[\\/]{2}/.test(input.trim());
}

/** `input` as an absolute local path; a network path only when it lies inside one of `networkRoots` (folders the user configured). */
export function localPath(input: string, networkRoots: readonly string[] = []): string {
  const candidate = path.resolve(toLocalPath(input));
  if (!isNetworkPath(candidate)) return candidate;
  const key = (value: string) => value.toLowerCase().replace(/[\\/]+$/, "");
  if (networkRoots.filter(isNetworkPath).some(root => key(candidate) === key(root) || key(candidate).startsWith(key(root) + path.sep))) return candidate;
  throw new ReferenceError_(`'${input}' is a network location. Add its folder as a workspace first to let the assistant read it.`);
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
  if (isNetworkPath(path.resolve(toLocalPath(input)))) throw new ReferenceError_("This workbook is opened from a network location. Add its folder manually.");
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

  /**
   * True for a saved reference file, and for anything inside a saved reference folder. `real` is already a real path, so a
   * symlink inside a folder cannot lead out of it; and a saved folder only counts while its path is its own real path,
   * so a folder swapped for a link to somewhere else opens nothing.
   */
  private isSavedReference(real: string): boolean {
    for (const saved of this.savedPaths()) {
      if (saved === real) return true;
      const prefix = saved.endsWith(path.sep) ? saved : saved + path.sep; // a drive root already ends with one
      if (!real.startsWith(prefix)) continue;
      try { if (fs.realpathSync(saved) === saved && fs.statSync(saved).isDirectory()) return true; }
      catch { /* moved or removed */ }
    }
    return false;
  }

  /** A model- or pane-supplied path; a network path is refused before any file system call unless it is inside a configured one. */
  private local(input: string): string {
    return localPath(input, [...this.roots, ...this.savedPaths()]);
  }

  /** Resolve a user-supplied path to a real path inside an allowed folder, or throw with guidance. */
  resolve(input: string, grants: ReadonlySet<string> = new Set()): string {
    const candidate = this.local(input);
    let real: string;
    try {
      real = fs.realpathSync(candidate);
    } catch {
      throw new ReferenceError_(`'${input}' does not exist or is not readable`);
    }
    if (!this.isSavedReference(real) && ![...this.roots, ...grants].some((root) => real === root || real.startsWith(root + path.sep))) {
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

  /**
   * What a person pointed at with @ and an allowed path: a file or, for a saved reference folder (or any folder the assistant
   * may read), a folder. Nothing is read or copied; the assistant is told where it is.
   */
  mentionedPath(input: string, grants: ReadonlySet<string> = new Set()): { path: string; folder: boolean } {
    const real = this.resolve(input, grants);
    const stat = fs.statSync(real);
    if (stat.isDirectory()) return { path: real, folder: true };
    if (!stat.isFile()) throw new ReferenceError_(`'${input}' is neither a file nor a folder`);
    return { path: real, folder: false };
  }

  /**
   * The `<user_attached_files>` block for what a person pointed at with @, shared by a message sent now and one queued while
   * a reply runs: each entry says whether it is a file or a folder, so a saved reference folder is never refused as "not a file".
   */
  mentionedFilesBlock(inputs: string[], grants: ReadonlySet<string> = new Set()): string {
    const entries = inputs.map((input) => this.mentionedPath(input, grants)).map(({ path: file, folder }) => JSON.stringify(folder
      ? { name: path.basename(file), path: file, kind: "reference_folder", hint: "A folder the user pointed at. Call list_reference_files with this path to see what is inside, then read the files you need." }
      : { name: path.basename(file), path: file, kind: "reference" }));
    return "\n\n<user_attached_files>\n" + entries.join("\n") + "\n</user_attached_files>";
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
    const local = this.local(input);
    try { real = fs.realpathSync(local); }
    catch { throw new ReferenceError_(`'${input}' does not exist or is not readable`); }
    const stat = fs.statSync(real);
    if (!stat.isFile() || !WORKBOOK_EXT.has(path.extname(real).toLowerCase())) throw new ReferenceError_("only Excel workbooks can be read");
    if (stat.size > MAX_WORKBOOK_BYTES) throw new ReferenceError_("workbook exceeds the 30 MB limit");
    return real;
  }

  /** What a person typed for a saved reference: the real path of a file or a folder, and which of the two it is. */
  referenceTarget(input: string): { path: string; kind: "file" | "folder" } {
    let real: string;
    const local = this.local(input);
    try { real = fs.realpathSync(local); }
    catch { throw new ReferenceError_(`'${input}' does not exist or is not readable`); }
    const stat = fs.statSync(real);
    if (stat.isDirectory()) return { path: real, kind: "folder" };
    if (stat.isFile()) return { path: real, kind: "file" };
    throw new ReferenceError_(`'${input}' is neither a file nor a folder`);
  }

  /** Exact file selected through the native picker, regardless of extension. */
  permissionTargetAny(input: string): string {
    let real: string;
    const local = this.local(input);
    try { real = fs.realpathSync(local); }
    catch { throw new ReferenceError_(`'${input}' does not exist or is not readable`); }
    if (!fs.statSync(real).isFile()) throw new ReferenceError_(`'${input}' is not a file`);
    return real;
  }
}
