import { spawnSync } from "node:child_process";
import { mkdirSync } from "node:fs";
import path from "node:path";

/** Run setup from the newly installed package, after npm has replaced this one. */
export async function updatePackage({ root, dataDir, nodePath, stop, wasRunning, run = spawnSync }) {
  mkdirSync(dataDir, { recursive: true });
  // npm.cmd needs the Windows command shell. All arguments are fixed literals.
  const npmOptions = { cwd: dataDir, shell: true, windowsHide: true };
  const location = run("npm.cmd", ["root", "--global"], { ...npmOptions, encoding: "utf8" });
  if (location.error || location.status !== 0 || !location.stdout?.trim()) {
    throw new Error("Could not locate the global npm installation. Check npm is available on PATH.");
  }
  const installedRoot = path.resolve(location.stdout.trim(), "excel-partner");
  if (installedRoot.toLowerCase() !== path.resolve(root).toLowerCase()) {
    throw new Error("Update supports the global npm installation. For a source checkout, run scripts/update.ps1 or scripts/update.sh; for npx, rerun npx with excel-partner@latest.");
  }
  console.log("Updating Excel Partner to the latest npm release…");
  await stop();
  const installed = run("npm.cmd", ["install", "--global", "excel-partner@latest"], { ...npmOptions, stdio: "inherit" });
  const cli = path.join(installedRoot, "bin", "excel-partner.mjs");
  if (installed.error || installed.status !== 0) {
    if (wasRunning) {
      const recovery = run(nodePath, [cli, "start"], { cwd: dataDir, stdio: "inherit", windowsHide: true });
      if (recovery.error || recovery.status !== 0) console.error("Could not restart the bridge. After fixing npm, run excel-partner setup.");
    }
    throw new Error("Update failed during npm install. Check the npm error above and retry excel-partner update.");
  }
  const setup = run(nodePath, [cli, "setup"], { cwd: dataDir, stdio: "inherit", windowsHide: true });
  if (setup.error || setup.status !== 0) throw new Error("Package updated, but setup failed. Run excel-partner setup to finish registration and restart the bridge.");
  console.log("Update complete. Reopen Excel to load the updated add-in.");
}
