import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * Excel Partner was renamed to Office Partner (npm: office-partner). This package stays on npm under the old name so that
 * everything that still says "excel-partner" keeps working: an installed copy updating itself from the old name, old
 * instructions, scripts. It holds no product code. Its only job is to make sure `office-partner` is installed globally and
 * to hand every command over to it.
 *
 * It never removes anything, itself included: the data folder, certificate, backups and settings are the new package's
 * too, and a copy that deleted its own folder while running would be a race on Windows.
 */

export const NEW_PACKAGE = "office-partner";

/** Where the global node_modules is, or null when npm cannot say. */
export function globalRoot(run = spawnSync, platform = process.platform) {
  const result = run(platform === "win32" ? "npm.cmd" : "npm", ["root", "--global"], { encoding: "utf8", shell: platform === "win32", windowsHide: true });
  const out = typeof result.stdout === "string" ? result.stdout.trim() : "";
  return result.error || result.status !== 0 || !out ? null : out;
}

/**
 * The global node_modules this copy sits in, when it sits in one. `npm root --global` is not always reliable: npm hides
 * anything in its output that looks like a secret (a folder named with a UUID, for one), so this is a second place to look.
 */
export function ownGlobalRoot(selfDir = path.dirname(fileURLToPath(import.meta.url))) {
  const packageRoot = path.resolve(selfDir, "..");
  const parent = path.dirname(packageRoot);
  return path.basename(parent) === "node_modules" ? parent : null;
}

/** The installed Office Partner command line, inside the global node_modules. */
export function newCliPath(root) {
  return path.join(root, NEW_PACKAGE, "bin", "office-partner.mjs");
}

/** What to `npm install`: the latest release, or what OFFICE_PARTNER_SPEC says (a version, a .tgz; for mirrors and testing). */
export function installSpec(env = process.env) {
  const spec = (env.OFFICE_PARTNER_SPEC || `${NEW_PACKAGE}@latest`).trim();
  // On Windows npm runs through the command shell, so only plain package specs and paths may get that far.
  if (!/^[\w@./:\\~+ -]+$/.test(spec)) throw new Error(`OFFICE_PARTNER_SPEC '${spec}' has characters that are not allowed.`);
  return spec;
}

/**
 * Make sure Office Partner is installed and return the path of its command line. Installs it (once) when it is missing.
 * Throws an Error whose message is fit to show to the person.
 */
export function ensureOfficePartner({ run = spawnSync, platform = process.platform, exists = existsSync, env = process.env, log = console.error, selfDir } = {}) {
  const npmRoot = globalRoot(run, platform);
  const own = ownGlobalRoot(selfDir);
  if (!npmRoot && !own) throw new Error("npm was not found. Install Node.js (version 20 or later), then run this command again.");
  const places = [...new Set([npmRoot, own].filter(Boolean))].map(newCliPath);
  const found = () => places.find((place) => exists(place));
  if (found()) return found();
  const cli = places[0];

  const spec = installSpec(env);
  log(`Excel Partner is now called Office Partner. Installing ${spec}...`);
  const install = run(platform === "win32" ? "npm.cmd" : "npm", ["install", "--global", /\s/.test(spec) ? `"${spec}"` : spec], {
    stdio: "inherit", shell: platform === "win32", windowsHide: true,
  });
  if (install.error || install.status !== 0) {
    throw new Error(`Could not install ${NEW_PACKAGE}. Fix the npm error above, then run:  npm install -g ${NEW_PACKAGE}`);
  }
  if (!found()) throw new Error(`${NEW_PACKAGE} was installed, but its command line was not found at ${cli}. Run:  npm install -g ${NEW_PACKAGE}`);
  return found();
}

/** Run the command with Office Partner and return its exit code. */
export function forward(argv, deps = {}) {
  const { run = spawnSync, nodePath = process.execPath, error = console.error } = deps;
  let cli;
  try {
    cli = ensureOfficePartner(deps);
  } catch (err) {
    error(err instanceof Error ? err.message : String(err));
    return 1;
  }
  const result = run(nodePath, [cli, ...argv], { stdio: "inherit", windowsHide: true });
  return result.error ? 1 : result.status ?? 1;
}
