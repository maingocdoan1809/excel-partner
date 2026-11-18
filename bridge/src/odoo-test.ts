import { type ChildProcess, spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import type { SkillStore } from "./skills.js";

export class OdooTestError extends Error {}
export type OdooTestPage = "init" | "dashboard";

/**
 * The odoo-test skill's own pages: its configuration form (`oqa init`) and its progress dashboard (`oqa dashboard`). The skill is a folder in the
 * user's skills; the add-in only starts the page and the skill opens it in the browser. Nothing about the testing itself lives here.
 */
export class OdooTest {
  private readonly running = new Map<OdooTestPage, ChildProcess>();

  constructor(
    private readonly skills: SkillStore,
    private readonly python = process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3"),
  ) {}

  async open(page: OdooTestPage): Promise<{ started: true }> {
    const skill = this.skills.get("odoo-test");
    const entry = skill ? path.join(skill.dir, "oqa.py") : "";
    if (!skill || !fs.existsSync(entry)) {
      throw new OdooTestError("The odoo-test skill is not installed. Put its folder in your skills (for example: python oqa.py install --to \"%USERPROFILE%\\.excel-partner\\skills\").");
    }
    this.running.get(page)?.kill();
    const child = spawn(this.python, [entry, page], { detached: true, stdio: "ignore", windowsHide: true });
    await new Promise<void>((resolve, reject) => {
      child.once("spawn", resolve);
      child.once("error", error => reject(new OdooTestError(`Could not start Python (${this.python}): ${error.message}`)));
    });
    child.unref();
    this.running.set(page, child);
    return { started: true };       // the page outlives the bridge on purpose: it stops by itself (the dashboard after 4 hours, the form once saved)
  }
}
