#!/usr/bin/env node
// Runs an update that the bridge started from Settings. The bridge copies this file and update.mjs to a temporary
// folder first: the installed package is replaced while this runs, and the bridge itself is stopped on the way.
//
// Environment: EP_ROOT (installed package), EP_DATA (data folder), EP_SPEC (npm spec), EP_FROM (version before),
// EP_GATEWAY ("1" when the gateway is set up and must be started again).
import { spawnSync } from "node:child_process";
import { rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { setTimeout as sleep } from "node:timers/promises";
import { installedVersion, packageNameOf, stopPackageProcesses, updatePackage, writeUpdateStatus } from "./update.mjs";

const root = process.env.EP_ROOT;
const dataDir = process.env.EP_DATA;
if (!root || !dataDir) {
  console.error("EP_ROOT and EP_DATA are required.");
  process.exit(2);
}
const spec = process.env.EP_SPEC || `${packageNameOf(root)}@latest`;
const from = process.env.EP_FROM || null;
const startedAt = new Date().toISOString();
const status = (patch) => writeUpdateStatus(dataDir, { from, spec, startedAt, ...patch });

const STEP_TEXT = {
  stopping: "Stopping the bridge",
  installing: "Installing the new version",
  recovering: "The install failed; restarting the old version",
  starting: "Starting the new version",
};

try {
  status({ state: "running", step: "preparing", message: "Getting ready", pid: process.pid, error: null, to: null, finishedAt: null });
  // Give the bridge time to answer the request that started this, and the pane time to show "Updating".
  await sleep(1500);
  const result = await updatePackage({
    root,
    dataDir,
    nodePath: process.execPath,
    wasRunning: true,
    spec,
    report: (step) => status({ state: "running", step, message: STEP_TEXT[step] ?? step }),
    stop: async () => { stopPackageProcesses({ root }); },
  });
  if (process.env.EP_GATEWAY === "1") {
    spawnSync(process.execPath, [path.join(result.installedRoot, "bin", "office-partner.mjs"), "gateway", "start"], { cwd: dataDir, stdio: "inherit", windowsHide: true });
  }
  status({ state: "done", step: "done", message: "Updated", to: result.version ?? installedVersion(result.installedRoot), finishedAt: new Date().toISOString() });
  removeOwnFolder();
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  status({ state: "failed", step: "failed", message: "The update did not finish", error: error instanceof Error ? error.message : String(error), finishedAt: new Date().toISOString() });
  process.exitCode = 1;
  removeOwnFolder();
}

/** The bridge copied this runner into a temporary folder; remove it again (the files are already loaded). */
function removeOwnFolder() {
  const folder = path.dirname(fileURLToPath(import.meta.url));
  // Only the bridge's temporary copy: never a runner started from the package or a checkout.
  if (!path.basename(folder).startsWith("excel-partner-update-") || path.dirname(folder).toLowerCase() !== path.resolve(tmpdir()).toLowerCase()) return;
  try { rmSync(folder, { recursive: true, force: true }); } catch { /* a leftover temp folder is harmless */ }
}
