import { spawnSync } from "node:child_process";
import { mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { randomUUID } from "node:crypto";
import path from "node:path";

const windowsRunSubkey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
export const windowsRunKey = `HKCU\\${windowsRunSubkey}`;
export const startupNames = { addin: "Excel Partner Bridge", gateway: "Agent Gateway" };

const launcherNames = { addin: "start-bridge.js", gateway: "start-gateway.js" };
const commandPrefix = "var command = ";
const quotedPaths = command => [...command.matchAll(/"([^"]+)"/g)].map(match => match[1]).slice(0, 2);

export function startupTargets(registryOutput, service) {
  const paths = quotedPaths(registryOutput);
  // Existing installations may still have a direct Node command in the Run value.
  if (paths.length !== 2 || path.win32.basename(paths[0]).toLowerCase() !== "wscript.exe") return paths;
  if (path.win32.basename(paths[1]) !== launcherNames[service]) return paths;
  try {
    const source = readFileSync(paths[1], "utf8");
    const declaration = source.split(/\r?\n/, 1)[0];
    if (declaration.startsWith(commandPrefix) && declaration.endsWith(";")) {
      const targets = quotedPaths(JSON.parse(declaration.slice(commandPrefix.length, -1)));
      if (targets.length === 2) return [...paths, ...targets];
    }
  } catch { /* The launcher is missing or unreadable. */ }
  return [...paths, null];
}

export function setWindowsStartup(service, { dataDir, cliPath, nodePath = process.execPath, enabled = true, run = spawnSync }) {
  const name = startupNames[service];
  const launcher = path.join(dataDir, launcherNames[service]);
  const runRegistry = args => run("reg.exe", [...args, "/f"], { stdio: "ignore", windowsHide: true });
  if (!enabled) {
    const result = runRegistry(["delete", windowsRunKey, "/v", name]);
    if (result.error) throw result.error;
    // reg.exe uses the same failure status for a missing value and a denied deletion.
    if (result.status !== 0) {
      const check = `$ErrorActionPreference = 'Stop'; $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('${windowsRunSubkey}'); if ($null -ne $key -and $key.GetValueNames() -contains '${name}') { exit 1 }`;
      const verified = run("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", check], { stdio: "ignore", windowsHide: true });
      if (verified.error) throw verified.error;
      if (verified.status !== 0) throw new Error(`Could not remove Windows sign-in startup for ${name}`);
    }
    rmSync(launcher, { force: true });
    return;
  }

  const host = path.win32.join(process.env.SystemRoot || "C:\\Windows", "System32", "wscript.exe");
  const startup = `"${host}" //B "${launcher}"`;
  const command = `"${nodePath}" "${cliPath}" start${service === "addin" ? " --at-logon" : ""}`;
  // WSH reads ASCII source. Escape Unicode in the string so paths survive any Windows code page.
  const literal = JSON.stringify(command).replace(/[^\x00-\x7f]/g, char => "\\u" + char.charCodeAt(0).toString(16).padStart(4, "0"));
  const source = `${commandPrefix}${literal};\r\nWScript.Quit(new ActiveXObject("WScript.Shell").Run(command, 0, true));\r\n`;
  mkdirSync(dataDir, { recursive: true });
  const temporary = `${launcher}.${randomUUID()}.tmp`;
  try {
    writeFileSync(temporary, source);
    renameSync(temporary, launcher);
  } finally {
    rmSync(temporary, { force: true });
  }
  const result = runRegistry(["add", windowsRunKey, "/v", name, "/t", "REG_SZ", "/d", startup]);
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`Could not register Windows sign-in startup for ${name}`);
}
