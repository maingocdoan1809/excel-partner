import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { expect, it } from "vitest";

it.skipIf(process.platform !== "win32")("runs the embedded MSI stop script before any bootstrap is installed", () => {
  const work = fs.mkdtempSync(path.join(os.tmpdir(), "office-partner-msi-launcher-"));
  try {
    const result = spawnSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
      "-File", path.resolve("tests/fixtures/msi-launcher.ps1"), "-Root", process.cwd(), "-Work", work], {
      encoding: "utf8", timeout: 30_000, windowsHide: true,
    });
    expect(result.error).toBeUndefined();
    expect(result.status, result.stdout + result.stderr).toBe(0);
    expect(result.stdout).toContain("PASS: embedded stop works");
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}, 35_000);
