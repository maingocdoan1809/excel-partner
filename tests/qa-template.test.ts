import { describe, expect, it } from "vitest";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { QA_TOOL_NAMES } from "../protocol/src/index.js";
import { DEFAULT_COLUMNS } from "../qa/src/config.js";
import { QA_SKILLS, installSkills, templateWorkbookPath, templatesDir } from "../qa/src/harness.js";
import { QA_HELP } from "../qa/src/cli.js";
import { parseFrontmatter, SkillStore } from "../bridge/src/skills.js";
import { makeProject } from "./qa-helpers.js";

const template = templateWorkbookPath();

function openpyxl(script: string): { ok: boolean; data?: any } {
  const result = spawnSync("python", ["-I", "-c", script, template], { encoding: "utf8" });
  if (result.error || result.status !== 0) return { ok: false };
  return { ok: true, data: JSON.parse(result.stdout) };
}

const probe = openpyxl("print('{}')");

describe.runIf(probe.ok)("TestPlan.xlsx template", () => {
  const info = openpyxl(`
import json, sys
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1])
tc = wb["TestCases"]
out = {
  "sheets": wb.sheetnames,
  "headers": [c.value for c in tc[1] if c.value],
  "rtm_row2": [c.value for c in wb["RTM"][2]],
  "dash": {r[0].value: r[1].value for r in wb["Dashboard"].iter_rows(min_row=3, max_row=11, max_col=2)},
  "req_validation": [d.formula1 for d in tc.data_validations.dataValidation],
  "req_ids": [r[0].value for r in wb["Requirements"].iter_rows(min_row=2, max_row=3)],
  "tc_ids": [r[0].value for r in tc.iter_rows(min_row=2, max_row=3)],
  "freeze": tc.freeze_panes,
}
print(json.dumps(out))
`).data;

  const letter = (header: string) => String.fromCharCode(65 + info.headers.indexOf(header));

  it("has the sheets the skills talk about", () => {
    expect(info.sheets).toEqual(["README", "Dashboard", "Requirements", "TestCases", "RTM", "Bugs"]);
  });

  it("uses exactly the column names the harness writes results to, in the order the tool-filled ones follow", () => {
    for (const header of Object.values(DEFAULT_COLUMNS)) expect(info.headers, header).toContain(header);
    expect(info.freeze).toBe("A2");
    expect(info.headers.indexOf("Script_Path")).toBeGreaterThan(info.headers.indexOf("Automation_Status") - 1);
  });

  it("points the RTM and dashboard formulas at the right TestCases columns", () => {
    const [, , perReq, automated, passed, failed] = info.rtm_row2;
    expect(perReq).toContain(`TestCases!$${letter("Req_ID")}$2:$${letter("Req_ID")}$1001`);
    expect(automated).toContain(`TestCases!$${letter("Automation_Status")}$2`);
    expect(automated).toContain('"Approved"');
    expect(passed).toContain(`TestCases!$${letter("Last_Result")}$2`);
    expect(failed).toContain('"Timed out"');
    const dashboard = JSON.stringify(info.dash);
    expect(dashboard).toContain(`TestCases!$${letter("TC_ID")}$2`);
    expect(dashboard).toContain(`TestCases!$${letter("Last_Result")}$2`);
    expect(dashboard).toContain(`TestCases!$${letter("Automation_Status")}$2`);
  });

  it("offers requirement ids from the Requirements sheet and ships two sample test cases for those requirements", () => {
    expect(info.req_validation.some((formula: string) => formula.includes("Requirements!$A$2"))).toBe(true);
    expect(info.req_ids).toEqual(["REQ_001", "REQ_002"]);
    expect(info.tc_ids).toEqual(["TC_001", "TC_002"]);
  });
});

describe("QA skills", () => {
  const dir = path.join(templatesDir(), "skills");

  it("are valid skills the bridge can list", () => {
    const store = new SkillStore([dir], fs.mkdtempSync(path.join(os.tmpdir(), "qa-skill-tpl-")));
    const names = store.list().map(skill => skill.name).sort();
    expect(names).toEqual([...QA_SKILLS].sort());
    for (const skill of store.list()) expect(skill.description.length).toBeGreaterThan(20);
  });

  it("only mention tools and commands that exist", () => {
    for (const name of QA_SKILLS) {
      const text = fs.readFileSync(path.join(dir, name, "SKILL.md"), "utf8");
      expect(parseFrontmatter(text).meta.name).toBe(name);
      for (const tool of text.match(/\bqa_[a-z_]+\b/g) ?? []) expect(QA_TOOL_NAMES as readonly string[], `${name} mentions ${tool}`).toContain(tool);
      for (const command of text.match(/excel-partner qa ([a-z]+)/g) ?? []) expect(QA_HELP, `${name}: ${command}`).toContain(command.replace("excel-partner qa ", ""));
    }
  });

  it("tell the model never to ask for credentials", () => {
    for (const name of ["qa-automate", "qa-heal", "qa-report"]) {
      expect(fs.readFileSync(path.join(dir, name, "SKILL.md"), "utf8"), name).toMatch(/mật khẩu/);
    }
    expect(fs.readFileSync(path.join(dir, "qa-automate", "SKILL.md"), "utf8")).toContain("Không bao giờ xin hoặc nhận mật khẩu");
  });

  it("installs without overwriting a skill the person edited", () => {
    const target = fs.mkdtempSync(path.join(os.tmpdir(), "qa-skills-"));
    expect(installSkills(false, target).map(item => item.action)).toEqual(["installed", "installed", "installed", "installed"]);
    expect(installSkills(false, target).every(item => item.action === "unchanged")).toBe(true);
    const mine = path.join(target, "qa-heal", "SKILL.md");
    fs.appendFileSync(mine, "\nMy own rule.\n");
    const again = installSkills(false, target);
    expect(again.find(item => item.name === "qa-heal")!.action).toBe("kept");
    expect(fs.readFileSync(mine, "utf8")).toContain("My own rule.");
    expect(installSkills(true, target).find(item => item.name === "qa-heal")!.action).toBe("updated");
    expect(fs.readFileSync(mine, "utf8")).not.toContain("My own rule.");
  });
});

describe("project scaffold", () => {
  it("includes the workbook template", () => {
    const { dir } = makeProject();
    expect(fs.statSync(path.join(dir, "TestPlan.xlsx")).size).toBeGreaterThan(5000);
  });
});
