import { spawnSync, type SpawnSyncReturns } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

/** Pinned so every installation runs the same driver; bump it deliberately, together with the Testing checks. */
export const PLAYWRIGHT_CORE_VERSION = "1.64.1";

export type PlaywrightModule = typeof import("playwright-core");

export interface PlaywrightLoaderDeps {
  /** Replaces `spawnSync` (tests). */
  run?: (command: string, args: string[], options: { cwd: string; shell: boolean; windowsHide: boolean; encoding: "utf8" }) => Pick<SpawnSyncReturns<string>, "status" | "error" | "stderr">;
  exists?: (file: string) => boolean;
  platform?: NodeJS.Platform;
  /** Replaces `import()` (tests). */
  load?: (specifier: string) => Promise<unknown>;
}

export function defaultToolsDir(): string {
  return path.join(os.homedir(), ".excel-partner", "tools");
}

/** Where the downloaded copy lives: a small private npm project under `toolsDir`. */
export function installedPlaywrightEntry(toolsDir: string): string {
  return path.join(toolsDir, "node_modules", "playwright-core", "index.mjs");
}

/**
 * The add-in does not ship playwright-core: most people never use the Testing sign-in, and the driver is large. A source
 * checkout has it as a devDependency and uses that copy. An installed package downloads it once, on the first sign-in,
 * into `toolsDir` (beside the other Excel Partner data, so an update of the add-in does not remove it) and loads it
 * from there afterwards.
 */
export async function loadPlaywright(toolsDir: string, out: (line: string) => void, deps: PlaywrightLoaderDeps = {}): Promise<PlaywrightModule> {
  const load = deps.load ?? (specifier => import(specifier));
  const exists = deps.exists ?? fs.existsSync;
  try {
    return await load("playwright-core") as PlaywrightModule;
  } catch (error) {
    if (!isMissingModule(error)) throw error;
  }
  const entry = installedPlaywrightEntry(toolsDir);
  if (!exists(entry)) {
    out(`Downloading the browser driver (playwright-core ${PLAYWRIGHT_CORE_VERSION}) once into ${toolsDir}…`);
    installPlaywright(toolsDir, deps);
    if (!exists(entry)) throw new Error(`The browser driver did not install into ${toolsDir}. Check that npm works in a terminal, then try again.`);
  }
  return await load(pathToFileURL(entry).href) as PlaywrightModule;
}

function isMissingModule(error: unknown): boolean {
  const code = (error as NodeJS.ErrnoException | undefined)?.code;
  return code === "ERR_MODULE_NOT_FOUND" || code === "MODULE_NOT_FOUND";
}

function installPlaywright(toolsDir: string, deps: PlaywrightLoaderDeps): void {
  const platform = deps.platform ?? process.platform;
  const run = deps.run ?? spawnSync;
  fs.mkdirSync(toolsDir, { recursive: true });
  const manifest = path.join(toolsDir, "package.json");
  if (!fs.existsSync(manifest)) fs.writeFileSync(manifest, JSON.stringify({ name: "office-partner-tools", private: true }, null, 2) + "\n");
  // npm.cmd needs the Windows command shell; every argument is a fixed literal.
  const result = run(platform === "win32" ? "npm.cmd" : "npm",
    ["install", "--no-audit", "--no-fund", "--no-package-lock", "--omit=dev", `playwright-core@${PLAYWRIGHT_CORE_VERSION}`],
    { cwd: toolsDir, shell: platform === "win32", windowsHide: true, encoding: "utf8" });
  if (result.error) throw new Error(`Could not run npm to download the browser driver: ${result.error.message}`);
  if (result.status !== 0) {
    const detail = String(result.stderr ?? "").trim().split(/\r?\n/).filter(Boolean).at(-1);
    throw new Error(`npm could not download the browser driver (exit ${result.status}).${detail ? ` ${detail}` : ""}`);
  }
}
