import { mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import path from "node:path";

export const UPDATE_STATUS_FILE = "update-status.json";

/** Read the progress record of the last (or running) update; null when there is none or it is unreadable. */
export function readUpdateStatus(dataDir) {
  try { return JSON.parse(readFileSync(path.join(dataDir, UPDATE_STATUS_FILE), "utf8")); } catch { return null; }
}

/** Merge `patch` into the progress record. Written through a temp file so a reader never sees half of it. */
export function writeUpdateStatus(dataDir, patch) {
  mkdirSync(dataDir, { recursive: true });
  const next = { ...(readUpdateStatus(dataDir) ?? {}), ...patch, updatedAt: new Date().toISOString() };
  const file = path.join(dataDir, UPDATE_STATUS_FILE);
  writeFileSync(`${file}.tmp`, JSON.stringify(next, null, 2));
  renameSync(`${file}.tmp`, file);
  return next;
}

/** What the installed package is called on npm: the name in its own package.json, so a rename needs no change here. */
export function packageNameOf(root) {
  try {
    const name = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8")).name;
    if (typeof name === "string" && /^[a-z0-9-~][a-z0-9-._~]*$/.test(name)) return name;
  } catch { /* fall through */ }
  return "office-partner";
}
