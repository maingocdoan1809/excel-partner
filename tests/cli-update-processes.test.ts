import { expect, it, vi } from "vitest";
import { spawnSync } from "node:child_process";
// @ts-expect-error The CLI ships independently of the TypeScript build.
import { stopPackageProcesses } from "../bin/update.mjs";

const powershell = spawnSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", "exit 0"], { stdio: "ignore", timeout: 5000 }).status === 0;

it.skipIf(!powershell)("runs the Windows selector against a process snapshot without terminating any processes", () => {
  const root = "C:\\npm\\node_modules\\excel-partner";
  const row = (ProcessId: number, CommandLine: string, ParentProcessId = 0, Cwd = "", Name = "node.exe") => ({ ProcessId, ParentProcessId, CommandLine, Cwd, Name });
  const processes = [
    row(100, `node ${root}\\bin\\excel-partner.mjs update`, 10),
    row(10, `powershell -Command node ${root}\\bin\\excel-partner.mjs update`),
    row(111, `node ${root}\\keep.js`),
    row(20, `node "${root}\\dist\\bridge.js"`),
    row(21, "node helper.js", 20),
    row(22, "node worker.js", 0, root + "\\providers"),
    row(23, "node worker.js", 0, root + "-helper"),
    row(24, `node ${root}-helper\\main.js`),
    row(25, `EXCEL.EXE ${root}\\test.xlsx`, 20, root, "EXCEL.EXE"),
    row(26, "node C:\\node\\npm-cli.js install --global --prefix C:\\npm excel-partner-helper"),
    row(27, "node C:\\node\\npm-cli.js install --global --prefix D:\\isolated excel-partner"),
    row(28, "node C:\\node\\npm-cli.js install --global excel-partner"),
    row(29, "node C:\\node\\npm-cli.js install --global --prefix C:\\npm excel-partner"),
    row(30, 'node "C:\\node\\npm-cli.js" install -g --prefix="C:\\npm" excel-partner@latest'),
    row(40, "node C:\\Temp\\excel-partner-update-same\\update-runner.mjs"),
    row(41, "node C:\\Temp\\excel-partner-update-other\\update-runner.mjs"),
    row(42, "node C:\\Temp\\excel-partner-update-unknown\\update-runner.mjs"),
  ];
  const run = vi.fn().mockReturnValue({ stdout: "0" });
  stopPackageProcesses({ root, keep: [111], platform: "win32", run });
  const original = Buffer.from(run.mock.calls[0]![1][3], "base64").toString("utf16le");
  // Substitute only OS observations and remove ALL termination commands before executing the actual selector.
  const selector = original.slice(original.indexOf("$root = "))
    .replace("$all = @(Get-CimInstance Win32_Process)", "$all = $reviewProcesses")
    .replace(`$cur = ${process.pid}`, "$cur = 100")
    .replace(`$spared = @(${process.pid},111)`, "$spared = @(100,111)")
    .replace("[EpCwd]::Get($id)", "(GetReviewCwd $id)")
    .replace("foreach ($id in $ids) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }", "")
    .replace("Start-Sleep -Milliseconds 800", "")
    .replace("Write-Output $ids.Count", "ConvertTo-Json -Compress @($ids | Sort-Object)");
  expect(selector).not.toContain("Stop-Process");
  expect(selector).not.toContain("Get-CimInstance");
  const setup = `
    $ErrorActionPreference = 'Stop'
    $reviewProcesses = ConvertFrom-Json '${JSON.stringify(processes).replace(/'/g, "''")}'
    function GetReviewCwd($id) { ($reviewProcesses | Where-Object { $_.ProcessId -eq $id } | Select-Object -First 1).Cwd }
    function Get-Content($LiteralPath, [switch]$Raw) {
      if ($LiteralPath -like '*update-same*') { return '{"root":"C:\\\\npm\\\\node_modules\\\\excel-partner"}' }
      if ($LiteralPath -like '*update-other*') { return '{"root":"D:\\\\other\\\\node_modules\\\\excel-partner"}' }
      throw 'No target metadata'
    }
  `;
  const result = spawnSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-EncodedCommand", Buffer.from(setup + selector, "utf16le").toString("base64")], {
    encoding: "utf8", windowsHide: true, timeout: 10000,
  });
  expect(result.status, result.stderr).toBe(0);
  expect(JSON.parse(result.stdout.trim())).toEqual([20, 21, 22, 29, 30, 40]);
});
