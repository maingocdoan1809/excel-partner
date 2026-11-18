import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { OdooTest, OdooTestError } from "../bridge/src/odoo-test.js";
import { SkillStore } from "../bridge/src/skills.js";

const folders: string[] = [];
function temp(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "odoo-test-"));
  folders.push(dir);
  return dir;
}
afterEach(() => { for (const dir of folders.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });

/** A skills folder holding an `odoo-test` skill whose entry script (run by node here, so no Python is needed) records how it was started. */
function install(marker: string) {
  const skills = temp();
  const dir = path.join(skills, "odoo-test");
  fs.mkdirSync(dir);
  fs.writeFileSync(path.join(dir, "SKILL.md"), ["---", "name: odoo-test", "description: Test Odoo from a workbook.", "hooks:", "  PreToolUse:", "    - matcher: \"Bash\"", "      hooks:", "        - type: command", "          command: \"python x\"", "---", "", "Run python \"${CLAUDE_SKILL_DIR}/oqa.py\" status"].join("\n"));
  fs.writeFileSync(path.join(dir, "oqa.py"), `require("node:fs").writeFileSync(${JSON.stringify(marker)}, process.argv.slice(2).join(" "));`);
  return { store: new SkillStore([skills], temp()), dir };
}

async function waitFor(file: string) {
  for (let i = 0; i < 100 && !fs.existsSync(file); i++) await new Promise(resolve => setTimeout(resolve, 50));
  return fs.existsSync(file) ? fs.readFileSync(file, "utf8") : null;
}

describe("the odoo-test skill in the add-in", () => {
  it("is listed, its folder is filled in where the skill names its scripts, and its hooks block does not confuse the reader", () => {
    const { store, dir } = install(path.join(temp(), "m"));
    expect(store.list().map(skill => skill.name)).toContain("odoo-test");
    const read = store.read("odoo-test");
    expect(read.description).toBe("Test Odoo from a workbook.");
    expect(read.content).toContain(`python "${dir.replaceAll("\\", "/")}/oqa.py" status`);
    expect(read.content).not.toContain("CLAUDE_SKILL_DIR");
  });

  it("starts the configuration form and the dashboard through the skill's own entry script", async () => {
    const marker = path.join(temp(), "started.txt");
    const { store } = install(marker);
    const odoo = new OdooTest(store, process.execPath);
    await expect(odoo.open("init")).resolves.toEqual({ started: true });
    expect(await waitFor(marker)).toBe("init");
    fs.rmSync(marker);
    await odoo.open("dashboard");
    expect(await waitFor(marker)).toBe("dashboard");
  });

  it("says plainly when the skill is missing or Python cannot start", async () => {
    await expect(new OdooTest(new SkillStore([temp()], temp())).open("init")).rejects.toThrow(/not installed/);
    const { store } = install(path.join(temp(), "m"));
    await expect(new OdooTest(store, path.join(temp(), "no-such-python")).open("init")).rejects.toBeInstanceOf(OdooTestError);
  });
});
