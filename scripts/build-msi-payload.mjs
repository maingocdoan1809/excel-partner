// Builds the folder an MSI installs into %APPDATA%\npm: the same layout `npm install -g office-partner` leaves, so `office-partner update` keeps working.
//   npm run pack:release && node scripts/build-msi-payload.mjs [--arch x64|arm64] [--out <dir>] [--tgz <release tarball>]
// Everything is built from pins: the exact dependencies of npm-shrinkwrap.json, plus the SQLite binary named in bin/runtime-manifest.json
// (verified by SHA-256), so the installed machine downloads no native module and needs no compiler.
// Node.js and the Copilot CLI are not in the payload: installer/msi/bootstrap.ps1 downloads those pinned versions while installing.
import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";
import { commandShim, msiVersion, pruneReason, readPins, slash } from "./msi/payload.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const arg = (name, fallback) => { const at = process.argv.indexOf(name); return at >= 0 ? process.argv[at + 1] : fallback; };
const arch = arg("--arch", "x64");
const manifest = JSON.parse(fs.readFileSync(path.join(root, "bin", "runtime-manifest.json"), "utf8"));
const pins = readPins(manifest, arch);
const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
const out = path.resolve(arg("--out", path.join(root, "release", "msi", arch)));
const payload = path.join(out, "payload");
const target = path.join(payload, "node_modules", pkg.name);

/** Runs a program and fails the build when it fails. `program` is a path or a command name; arguments are never passed through a shell. */
function run(program, args, options = {}) {
  const result = spawnSync(program, args, { stdio: "inherit", ...options });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${path.basename(program)} ${args.slice(0, 2).join(" ")} failed (exit ${result.status}).`);
}
/** npm through the running Node, so Windows needs no shell. */
const npm = (args, options) => process.env.npm_execpath ? run(process.execPath, [process.env.npm_execpath, ...args], options) : run(process.platform === "win32" ? "npm.cmd" : "npm", args, { shell: process.platform === "win32", ...options });
const sha256 = (file) => crypto.createHash("sha256").update(fs.readFileSync(file)).digest("hex");

/**
 * Downloads `url` to `file` and refuses it unless it has the pinned SHA-256.
 * Purpose: a corrupted or replaced download can never end up inside the installer.
 */
async function download(url, file, expected) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Could not download ${url} (HTTP ${response.status}).`);
  fs.writeFileSync(file, Buffer.from(await response.arrayBuffer()));
  const actual = sha256(file);
  if (actual !== expected) { fs.rmSync(file, { force: true }); throw new Error(`${path.basename(file)} does not match its pinned SHA-256 (expected ${expected}, got ${actual}).`); }
}

/** Deletes what `pruneReason` allows below `dir`; returns how many paths and bytes went. */
function prune(dir) {
  const removed = { paths: 0, bytes: 0 };
  const size = (target) => { const stat = fs.lstatSync(target); return stat.isDirectory() ? fs.readdirSync(target).reduce((sum, name) => sum + size(path.join(target, name)), 0) : stat.size; };
  const walk = (current) => {
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const full = path.join(current, entry.name);
      const isDirectory = entry.isDirectory();
      if (pruneReason(slash(path.relative(dir, full)), isDirectory)) { removed.bytes += size(full); removed.paths++; fs.rmSync(full, { recursive: true, force: true }); }
      else if (isDirectory) walk(full);
    }
  };
  walk(dir);
  return removed;
}

function installedVersion(name) {
  return JSON.parse(fs.readFileSync(path.join(target, "node_modules", ...name.split("/"), "package.json"), "utf8")).version;
}
/** Every file below `dir` with its size. */
function listFiles(dir) {
  const files = [];
  const walk = (current) => { for (const entry of fs.readdirSync(current, { withFileTypes: true })) { const full = path.join(current, entry.name); if (entry.isDirectory()) walk(full); else files.push({ full, size: fs.statSync(full).size }); } };
  walk(dir);
  return files;
}
/** Compressed size of a folder as a .tar.gz: a close upper bound for what the MSI cabinet takes. */
function compressedSize(dir) {
  return new Promise((resolve, reject) => {
    const tar = spawnSync("tar", ["-cf", "-", "-C", dir, "."], { maxBuffer: 1 << 30 });
    if (tar.error || tar.status !== 0) return reject(tar.error ?? new Error("tar failed"));
    zlib.gzip(tar.stdout, { level: 9 }, (error, data) => error ? reject(error) : resolve(data.length));
  });
}

fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(target, { recursive: true });
const work = fs.mkdtempSync(path.join(out, ".work-"));
try {
  let tgz = arg("--tgz", path.join(root, "release", `${pkg.name}-${pkg.version}.tgz`));
  if (!fs.existsSync(tgz)) { console.log("No release tarball yet: running npm run pack:release"); npm(["run", "pack:release"], { cwd: root }); }
  if (!fs.existsSync(tgz)) throw new Error(`Release tarball not found: ${tgz}`);
  tgz = path.resolve(tgz);

  console.log(`==> Unpacking ${path.basename(tgz)}`);
  run("tar", ["-xzf", tgz, "-C", target, "--strip-components=1"]);

  console.log(`==> Installing the pinned dependencies for win32-${arch} (npm-shrinkwrap.json, no scripts)`);
  npm(["install", "--omit=dev", "--ignore-scripts", "--no-audit", "--no-fund", "--os=win32", `--cpu=${arch}`], { cwd: target });

  // The versions the manifest names must be the ones npm installed; a drift here is a build error, not a surprise on a user's PC.
  for (const [name, version] of [["better-sqlite3", pins.sqlite.version], ["@github/copilot-sdk", pins.copilot.sdkVersion], ["koffi", pkg.dependencies.koffi]]) {
    if (installedVersion(name) !== version) throw new Error(`${name} ${installedVersion(name)} was installed, the pins say ${version}.`);
  }
  if (!fs.existsSync(path.join(target, "node_modules", "@koromix", `koffi-win32-${arch}`))) throw new Error(`koffi's win32-${arch} binary was not installed.`);

  console.log(`==> SQLite binary ${pins.sqlite.file} (SHA-256 checked)`);
  const sqliteArchive = path.join(work, pins.sqlite.file);
  await download(pins.sqlite.url, sqliteArchive, pins.sqlite.sha256);
  run("tar", ["-xzf", sqliteArchive, "-C", path.join(target, "node_modules", "better-sqlite3")]);
  const binary = path.join(target, "node_modules", "better-sqlite3", "build", "Release", "better_sqlite3.node");
  if (!fs.existsSync(binary)) throw new Error("The SQLite archive did not contain build/Release/better_sqlite3.node.");

  console.log("==> Removing what is never loaded");
  const removed = prune(path.join(target, "node_modules"));
  console.log(`    ${removed.paths} paths, ${(removed.bytes / 1048576).toFixed(1)} MB`);

  for (const required of ["dist/bridge/src/index.js", "frontend/manifest.xml", "frontend/dist/taskpane.js", "scripts/package-windows.ps1", "bin/office-partner.mjs", "bin/runtime-manifest.json", "npm-shrinkwrap.json"]) {
    if (!fs.existsSync(path.join(target, required))) throw new Error(`The payload is missing ${required}.`);
  }
  if (!process.argv.includes("--no-smoke")) {
    console.log("==> Loading every dependency and compiled module from the pruned payload");
    const smoke = path.join(target, ".smoke.mjs");
    fs.copyFileSync(path.join(root, "scripts", "msi", "smoke-payload.mjs"), smoke);
    try { run(process.execPath, [smoke, ...(process.platform === "win32" && arch === process.arch ? [] : ["--no-native"])], { cwd: target }); } finally { fs.rmSync(smoke, { force: true }); }
  }
  fs.mkdirSync(path.join(target, "installer"), { recursive: true });
  fs.copyFileSync(path.join(root, "installer", "msi", "bootstrap.ps1"), path.join(target, "installer", "bootstrap.ps1"));
  fs.writeFileSync(path.join(payload, `${pkg.name}.cmd`), commandShim());

  const files = listFiles(payload);
  const bytes = files.reduce((sum, file) => sum + file.size, 0);
  const compressed = await compressedSize(payload);
  const info = { name: pkg.name, version: pkg.version, msiVersion: msiVersion(pkg.version), arch, files: files.length, bytes, compressedBytes: compressed, pins };
  fs.writeFileSync(path.join(out, "payload-info.json"), JSON.stringify(info, null, 2) + "\n");
  console.log(`\nPayload ${out}`);
  console.log(`  ${files.length} files, ${(bytes / 1048576).toFixed(1)} MB on disk, about ${(compressed / 1048576).toFixed(1)} MB compressed (MSI limit goal: 10 MB)`);
  console.log(`  Installed by the MSI: Node ${pins.node.version}, ${pins.copilot.package}@${pins.copilot.version} are downloaded at install time (pinned).`);
} finally {
  fs.rmSync(work, { recursive: true, force: true });
}
