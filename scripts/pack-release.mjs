// Builds the release tarball (add-in, bridge and gateway in one package) into ./release.
//   npm run pack:release
// Publish with: npm publish ./release/excel-partner-<version>.tgz   (or simply `npm publish`)
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const release = path.join(root, "release");
fs.mkdirSync(release, { recursive: true });

const args = ["pack", "--pack-destination", release];
// Running through npm exposes its CLI path; invoking it with node avoids shell quoting on Windows.
const result = process.env.npm_execpath
  ? spawnSync(process.execPath, [process.env.npm_execpath, ...args], { cwd: root, stdio: "inherit" })
  : spawnSync("npm", args, { cwd: root, stdio: "inherit", shell: process.platform === "win32" });
if (result.status !== 0) process.exit(result.status ?? 1);

console.log("\nRelease artifacts:");
for (const file of fs.readdirSync(release).filter(f => f.endsWith(".tgz"))) {
  console.log(`  release/${file}  ${(fs.statSync(path.join(release, file)).size / 1024).toFixed(0)} KB`);
}
