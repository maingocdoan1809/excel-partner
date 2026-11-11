import fs from "node:fs";
import { describe, expect, it } from "vitest";

const read = (file: string) => fs.readFileSync(file, "utf8");

// The Windows sign-in entries cannot be run here, so check the commands and the shortcut they are built from.
describe("Windows startup registration", () => {
  const script = read("scripts/package-windows.ps1");

  it("D3 starts the bridge at sign-in in the mode that retries and records the outcome", () => {
    expect(script).toContain(`'" start --at-logon')`);
  });

  it("D4 puts a Fix connection shortcut in the Start Menu and removes it again on uninstall", () => {
    expect(script).toMatch(/GetFolderPath\('Programs'\)\) 'Office Partner - Fix connection\.lnk'/);
    expect(script).toContain(`$link.Arguments = '"' + $CliPath + '" fix'`);
    expect(script).toContain("$link.TargetPath = $NodePath");
    expect(script).toContain("Remove-Item -LiteralPath $fixShortcut");
  });

  it("D3 registers the gateway sign-in entry in the same mode", () => {
    expect(read("gateway/bin/agent-gateway.mjs")).toContain('start gateway --at-logon`');
  });

  it("D4 offers the same repair from the command line", () => {
    expect(read("bin/office-partner.mjs")).toMatch(/\n  fix {8}Start the bridge and gateway again/);
  });
});
