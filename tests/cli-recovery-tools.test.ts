import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { afterEach, describe, expect, it } from "vitest";

const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
/** A data folder whose recovery launcher only reports how it was called. */
function dataFolder(withRecovery: boolean) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "cli-recovery-")); dirs.push(dir);
  if (withRecovery) {
    fs.mkdirSync(path.join(dir, "recovery"));
    fs.writeFileSync(path.join(dir, "recovery", "recovery.mjs"), 'console.log("RECOVERY " + JSON.stringify(process.argv.slice(2)) + " nested=" + process.env.OFFICE_PARTNER_RECOVERY_ACTIVE); process.exitCode = 7;');
  }
  return dir;
}
const cli = (dir: string, ...args: string[]) => spawnSync(process.execPath, [path.resolve("bin/office-partner.mjs"), ...args],
  { env: { ...process.env, OFFICE_PARTNER_DATA_DIR: dir, OFFICE_PARTNER_RECOVERY_ACTIVE: "" }, encoding: "utf8", timeout: 20000 });

describe("management commands use the recovery tools", () => {
  it("D6 sends update, rollback, repair, versions and fix to the always-current recovery copy and keeps its exit code", () => {
    const dir = dataFolder(true);
    for (const command of [["update"], ["rollback", "1.0.0"], ["repair"], ["versions"], ["fix"], ["providers", "list"], ["doctor", "--json"]]) {
      const result = cli(dir, ...command);
      expect(result.stdout, command.join(" ")).toContain(`RECOVERY ${JSON.stringify(command)} nested=1`);
      expect(result.status, command.join(" ")).toBe(7);
    }
  });

  it("passes the exact release given to update --from, and maps `gateway doctor` to the doctor command", () => {
    const dir = dataFolder(true);
    expect(cli(dir, "update", "--from", "office-partner@1.2.3").stdout).toContain('RECOVERY ["update","--from","office-partner@1.2.3"]');
    expect(cli(dir, "gateway", "doctor", "--json").stdout).toContain('RECOVERY ["doctor","gateway","--json"]');
  });

  it("does not hand over again when it is already running inside the recovery tools", () => {
    const dir = dataFolder(true);
    const result = spawnSync(process.execPath, [path.resolve("bin/office-partner.mjs"), "versions"],
      { env: { ...process.env, OFFICE_PARTNER_DATA_DIR: dir, OFFICE_PARTNER_RECOVERY_ACTIVE: "1" }, encoding: "utf8", timeout: 20000 });
    expect(result.stdout).not.toContain("RECOVERY");
    expect(JSON.parse(result.stdout)).toMatchObject({ active: null, providers: expect.any(Array) });
  });

  it("runs from the package itself when there is no recovery launcher yet", () => {
    const result = cli(dataFolder(false), "versions");
    expect(result.status, result.stderr).toBe(0);
    expect(JSON.parse(result.stdout)).toMatchObject({ active: null, providers: expect.any(Array) });
  });
});
