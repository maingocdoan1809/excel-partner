// Turns the payload folder of scripts/build-msi-payload.mjs into release/office-partner-<version>-<arch>.msi (needs the WiX 5 CLI on Windows).
//   dotnet tool install --global wix --version 5.0.2 && wix extension add --global WixToolset.Util.wixext/5.0.2
//   node scripts/build-msi-payload.mjs --arch x64 && node scripts/build-msi.mjs --arch x64
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

/** The WiX release the installer definition is written and tested for. WiX 6 and later change the licence terms, so this is pinned. */
const WIX_VERSION = "5.0.2";
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const arg = (name, fallback) => { const at = process.argv.indexOf(name); return at >= 0 ? process.argv[at + 1] : fallback; };
const arch = arg("--arch", "x64");
const outDir = path.resolve(arg("--payload", path.join(root, "release", "msi", arch)));
const info = JSON.parse(fs.readFileSync(path.join(outDir, "payload-info.json"), "utf8"));
if (info.arch !== arch) throw new Error(`The payload in ${outDir} is for ${info.arch}, not ${arch}.`);

const version = spawnSync("wix", ["--version"], { encoding: "utf8", shell: process.platform === "win32" });
if (version.error || version.status !== 0) throw new Error(`The WiX CLI was not found. Install it: dotnet tool install --global wix --version ${WIX_VERSION}`);
if (!version.stdout.trim().startsWith(WIX_VERSION)) throw new Error(`WiX ${version.stdout.trim()} is installed; this installer is built with ${WIX_VERSION} (dotnet tool update --global wix --version ${WIX_VERSION}).`);

const msi = path.join(root, "release", `${info.name}-${info.version}-${arch}.msi`);
// `wix build` does not run ICE validation (that is the separate `wix msi validate`), so the per-user-file ICE38/ICE64/ICE91 warnings never appear here.
const args = ["build", "-arch", arch, "-ext", "WixToolset.Util.wixext",
  "-d", `Version=${info.msiVersion}`, "-d", `PayloadDir=${path.join(outDir, "payload")}`, "-o", msi, path.join(root, "installer", "msi", "OfficePartner.wxs")];
const result = spawnSync("wix", args, { stdio: "inherit", shell: process.platform === "win32" });
if (result.error || result.status !== 0) throw new Error("wix build failed.");
console.log(`\n${msi}  ${(fs.statSync(msi).size / 1048576).toFixed(1)} MB`);
