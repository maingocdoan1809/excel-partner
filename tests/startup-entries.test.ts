import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { describe, expect, it } from "vitest";
// @ts-expect-error The packaged Windows launcher is registered by the standalone CLI.
import { setWindowsStartup, startupTargets } from "../bin/windows-startup.mjs";

const read = (file: string) => fs.readFileSync(file, "utf8");

describe("Windows startup registration", () => {
  const script = read("scripts/package-windows.ps1");

  it.each([
    ["addin", "Excel Partner Bridge", "start-bridge.js", ["start", "--at-logon"]],
    ["gateway", "Agent Gateway", "start-gateway.js", ["start"]],
  ])("registers and removes the hidden %s launcher", (service, name, filename, expectedArgs) => {
    const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "office partner startup-"));
    const cliPath = path.join(dataDir, "CLI thử.mjs");
    const marker = path.join(dataDir, "args.json");
    const launcher = path.join(dataDir, filename);
    const registryCommands: string[][] = [];
    const run = (_command: string, args: string[]) => {
      registryCommands.push(args);
      return { status: 0 };
    };
    try {
      fs.writeFileSync(cliPath, `import fs from "node:fs"; fs.writeFileSync(${JSON.stringify(marker)}, JSON.stringify(process.argv.slice(2)));`);
      setWindowsStartup(service, { dataDir, cliPath, run });
      const register = registryCommands[0]!;
      expect(register).toContain(name);
      const command = register[register.indexOf("/d") + 1];
      const [host] = startupTargets(command, service);
      expect(command).toBe(`"${host}" //B "${launcher}"`);
      expect(startupTargets(command, service)).toEqual([host, launcher, process.execPath, cliPath]);
      if (process.platform === "win32") {
        const result = spawnSync(host, ["//B", launcher], { windowsHide: true, timeout: 10_000 });
        expect(result.error).toBeUndefined();
        expect(result.status).toBe(0);
        expect(JSON.parse(read(marker))).toEqual(expectedArgs);
      }
      setWindowsStartup(service, { dataDir, enabled: false, run });
      expect(registryCommands[1]).toEqual(["delete", register[1], "/v", name, "/f"]);
      expect(fs.existsSync(launcher)).toBe(false);
    } finally {
      fs.rmSync(launcher, { force: true });
      fs.rmSync(cliPath, { force: true });
      fs.rmSync(marker, { force: true });
      fs.rmdirSync(dataDir);
    }
  });

  it.each([true, false])("handles failed registry deletion when the startup value is absent: %s", absent => {
    const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "office partner startup-"));
    const launcher = path.join(dataDir, "start-gateway.js");
    const commands: string[] = [];
    const run = (command: string) => {
      commands.push(command);
      return { status: command === "powershell.exe" && absent ? 0 : 1 };
    };
    try {
      fs.writeFileSync(launcher, "launcher");
      const disable = () => setWindowsStartup("gateway", { dataDir, enabled: false, run });
      if (absent) {
        disable();
        expect(fs.existsSync(launcher)).toBe(false);
      } else {
        expect(disable).toThrow("Could not remove Windows sign-in startup for Agent Gateway");
        expect(read(launcher)).toBe("launcher");
      }
      expect(commands).toEqual(["reg.exe", "powershell.exe"]);
    } finally {
      fs.rmSync(launcher, { force: true });
      fs.rmdirSync(dataDir);
    }
  });

  it("D4 puts a Fix connection shortcut in the Start Menu and removes it again on uninstall", () => {
    expect(script).toMatch(/GetFolderPath\('Programs'\)\) 'Office Partner - Fix connection\.lnk'/);
    expect(script).toContain(`$link.Arguments = '"' + $CliPath + '" fix'`);
    expect(script).toContain("$link.TargetPath = $NodePath");
    expect(script).toContain("Remove-Item -LiteralPath $fixShortcut");
  });

  it("D4 offers the same repair from the command line", () => {
    expect(read("bin/office-partner.mjs")).toMatch(/\n  fix {8}Start the bridge and gateway again/);
  });
});
