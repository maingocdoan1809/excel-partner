// Pure helpers for the thin MSI payload (see scripts/build-msi-payload.mjs). No I/O here, so tests can run them without a build.
import path from "node:path";

/** Node.js ABI (process.versions.modules) of each major release the pinned prebuilt SQLite binary can be built for. */
export const NODE_ABI_BY_MAJOR = { 20: 115, 22: 127, 24: 137 };
export const ARCHS = ["x64", "arm64"];
const SEMVER = /^(\d+)\.(\d+)\.(\d+)$/;
const SHA256 = /^[0-9a-f]{64}$/;

/**
 * Reads the pins an MSI is built from out of `bin/runtime-manifest.json` and refuses anything loose.
 * Purpose: one place for every version and hash, so a build can never pick "whatever is latest".
 * @param {any} manifest Parsed runtime-manifest.json.
 * @param {"x64"|"arm64"} arch Windows architecture of the MSI.
 * @returns {{ node: { version: string, abi: number, sha256: string, file: string, url: string, shasumsUrl: string },
 *            sqlite: { version: string, sha256: string, file: string, url: string },
 *            copilot: { package: string, version: string, sdkVersion: string } }}
 */
export function readPins(manifest, arch) {
  if (!ARCHS.includes(arch)) throw new Error(`Unknown architecture '${arch}' (use ${ARCHS.join(" or ")}).`);
  const node = manifest?.node;
  const nodeMatch = SEMVER.exec(node?.version ?? "");
  if (!nodeMatch) throw new Error("runtime-manifest.json: node.version must be an exact x.y.z release.");
  const major = Number(nodeMatch[1]);
  if (NODE_ABI_BY_MAJOR[major] !== node.abi) throw new Error(`runtime-manifest.json: node ${node.version} has ABI ${NODE_ABI_BY_MAJOR[major] ?? "unknown"}, the manifest says ${node.abi}.`);
  const nodeSha = node.win32?.[arch]?.sha256;
  if (!SHA256.test(nodeSha ?? "")) throw new Error(`runtime-manifest.json: node.win32.${arch}.sha256 is missing or not a SHA-256.`);
  const sqlite = manifest?.native?.["better-sqlite3"];
  const sqliteSha = sqlite?.prebuilt?.[`win32-${arch}`]?.sha256;
  if (!SEMVER.test(sqlite?.version ?? "") || !SHA256.test(sqliteSha ?? "")) throw new Error(`runtime-manifest.json: native.better-sqlite3 needs an exact version and a prebuilt win32-${arch} SHA-256.`);
  const copilot = manifest?.providers?.copilot;
  if (!copilot?.package || !SEMVER.test(copilot.version ?? "") || !SEMVER.test(copilot.sdkVersion ?? "")) throw new Error("runtime-manifest.json: providers.copilot needs an exact version and sdkVersion.");
  const nodeFile = `node-v${node.version}-win-${arch}.zip`;
  const sqliteFile = `better-sqlite3-v${sqlite.version}-node-v${node.abi}-win32-${arch}.tar.gz`;
  return {
    node: { version: node.version, abi: node.abi, sha256: nodeSha, file: nodeFile,
      url: `https://nodejs.org/dist/v${node.version}/${nodeFile}`, shasumsUrl: `https://nodejs.org/dist/v${node.version}/SHASUMS256.txt` },
    sqlite: { version: sqlite.version, sha256: sqliteSha, file: sqliteFile,
      url: `https://github.com/WiseLibs/better-sqlite3/releases/download/v${sqlite.version}/${sqliteFile}` },
    copilot: { package: copilot.package, version: copilot.version, sdkVersion: copilot.sdkVersion },
  };
}

/** Folders (by name) that hold nothing a package needs at run time. */
const DEAD_FOLDERS = new Set(["test", "tests", "__tests__", "docs", "doc", "example", "examples", ".github", ".vscode"]);
/** Files that are documentation, typings or maps. Licence files are never in this list: their text must travel with the code. */
const DEAD_FILE = /(?:\.d\.[cm]?ts|\.map|\.markdown|\.tsbuildinfo)$|^(?:readme|changelog|changes|history|authors|contributing|security)(?:\.(?:md|markdown|txt|rst))?$/i;
const KEEP_FILE = /^(?:licen[sc]e|notice|copying)/i;
/** Whole packages the app never loads from `node_modules`, with the reason. */
const DEAD_PACKAGES = [
  // The Copilot SDK carries its own 200 MB runtime for in-process hosting. The adapter always starts the managed CLI by path (RuntimeConnection.forStdio).
  { prefix: "@github/copilot-sdk-", reason: "SDK in-process runtime; the adapter starts the managed CLI by path" },
];
/** Source and build files of a native module whose compiled binary is shipped instead. */
const NATIVE_SOURCE = { "better-sqlite3": new Set(["deps", "src", "binding.gyp"]) };

/**
 * Says whether a path inside the payload's `node_modules` can be left out of the MSI.
 * Purpose: shrink the installer without touching code that runs.
 * @param {string} relative Path below `node_modules`, with `/` separators (for example `zod/lib/index.d.ts`).
 * @param {boolean} isDirectory True for a folder.
 * @returns {string | null} The reason it can go, or null when it must stay.
 */
export function pruneReason(relative, isDirectory) {
  const all = relative.split("/").filter(Boolean);
  const parts = all.slice(all.lastIndexOf("node_modules") + 1); // judged inside its own package, however deep it is nested
  const name = parts[parts.length - 1] ?? "";
  const scoped = parts[0]?.startsWith("@");
  const packageName = scoped ? parts.slice(0, 2).join("/") : parts[0];
  for (const dead of DEAD_PACKAGES) if (isDirectory && parts.length === (scoped ? 2 : 1) && packageName?.startsWith(dead.prefix)) return dead.reason;
  const inside = parts.slice(scoped ? 2 : 1);
  const native = NATIVE_SOURCE[packageName ?? ""];
  if (native && inside.length === 1 && native.has(inside[0] ?? "")) return "native module source";
  // Only a folder directly in the package root: a deeper `doc` or `test` can be code (yaml keeps its document model in dist/doc).
  if (isDirectory) return inside.length === 1 && DEAD_FOLDERS.has(name.toLowerCase()) ? "documentation, tests or examples" : null;
  if (KEEP_FILE.test(name)) return null;
  return DEAD_FILE.test(name) ? "documentation, typings or source map" : null;
}

/**
 * The `office-partner.cmd` command the MSI puts next to the package, like the one `npm install -g` writes.
 * Purpose: run the pinned Node.js (the one the bundled SQLite binary was built for) when this installer put it there; otherwise the first `node` on PATH.
 * @returns {string} Batch file text with CRLF line endings.
 */
export function commandShim() {
  return [
    "@ECHO off",
    "GOTO start",
    ":find_dp0",
    "SET dp0=%~dp0",
    "EXIT /b",
    ":start",
    "SETLOCAL",
    "CALL :find_dp0",
    "",
    'IF EXIST "%LOCALAPPDATA%\\Programs\\nodejs\\node.exe" (',
    '  SET "_prog=%LOCALAPPDATA%\\Programs\\nodejs\\node.exe"',
    ") ELSE IF EXIST \"%LOCALAPPDATA%\\Programs\\office-partner-node\\node.exe\" (",
    '  SET "_prog=%LOCALAPPDATA%\\Programs\\office-partner-node\\node.exe"',
    ") ELSE (",
    '  SET "_prog=node"',
    "  SET PATHEXT=%PATHEXT:;.JS;=;%",
    ")",
    "",
    'endLocal & goto #_undefined_# 2>NUL || title %COMSPEC% & "%_prog%"  "%dp0%\\node_modules\\office-partner\\bin\\office-partner.mjs" %*',
    "",
  ].join("\r\n");
}

/**
 * Turns a package version into the four-part number Windows Installer accepts.
 * @param {string} version Package version such as `0.1.26` or `0.2.0-beta.1`.
 * @returns {string} `major.minor.patch` (a pre-release suffix is dropped).
 */
export function msiVersion(version) {
  const match = /^(\d+)\.(\d+)\.(\d+)/.exec(version);
  if (!match || Number(match[1]) > 255 || Number(match[2]) > 255 || Number(match[3]) > 65535) throw new Error(`'${version}' cannot be an MSI version (limits 255.255.65535).`);
  return `${match[1]}.${match[2]}.${match[3]}`;
}

/** Normalises a Windows or POSIX path to `/` separators for the rules above. */
export const slash = (value) => value.split(path.sep).join("/");
