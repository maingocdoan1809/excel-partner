import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { zipSync, strToU8 } from "fflate";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { UserSkills, MAX_SKILL_UPLOAD_BYTES } from "../bridge/src/skill-upload.js";
import { SkillStore } from "../bridge/src/skills.js";

let folder: string;
let uploads: UserSkills;
const manifest = "---\nname: monthly-report\ndescription: Format a monthly report\n---\nUse the user-selected workbook.\n";

beforeEach(() => {
  folder = fs.mkdtempSync(path.join(os.tmpdir(), "user-skills-"));
  uploads = new UserSkills(path.join(folder, "skills"));
});
afterEach(() => fs.rmSync(folder, { recursive: true, force: true }));

describe("user skill uploads", () => {
  it("installs Markdown and makes it available to the existing skill engine", () => {
    const result = uploads.install("SKILL.md", Buffer.from(manifest), () => false);
    expect(result).toEqual({ name: "monthly-report", description: "Format a monthly report" });
    const skills = new SkillStore([uploads.root]);
    expect(skills.list().map(skill => skill.name)).toEqual(["monthly-report"]);
    expect(skills.read("monthly-report").content).toContain("user-selected workbook");
  });

  it("uses a Markdown filename when frontmatter has no name", () => {
    expect(uploads.install("format-table.md", Buffer.from("Format the selected table."), () => false).name).toBe("format-table");
  });

  it("installs a ZIP containing a skill folder and text references", () => {
    const archive = zipSync({ "report/SKILL.md": strToU8(manifest), "report/references/style.md": strToU8("Use readable headings.") });
    uploads.install("report.zip", Buffer.from(archive), () => false);
    expect(new SkillStore([uploads.root]).read("monthly-report", "references/style.md").content).toBe("Use readable headings.");
  });

  it("rejects archive paths that could escape the upload directory", () => {
    for (const invalid of ["../escape.md", "/absolute.md", "C:/escape.md", "references\\escape.md", "references/CON.txt"]) {
      const archive = zipSync({ "SKILL.md": strToU8(manifest), [invalid]: strToU8("Invalid") });
      expect(() => uploads.install("report.zip", Buffer.from(archive), () => false)).toThrow(/paths|file names/);
    }
    expect(fs.existsSync(uploads.root)).toBe(false);
  });

  it("rejects ambiguous archives, binary files and excessive expansion", () => {
    expect(() => uploads.install("report.zip", Buffer.from(zipSync({ "README.md": strToU8("Missing skill") })), () => false)).toThrow(/exactly one SKILL/);
    expect(() => uploads.install("report.zip", Buffer.from(zipSync({ "SKILL.md": strToU8(manifest), "run.exe": new Uint8Array([0, 1]) })), () => false)).toThrow(/text reference/);
    expect(() => uploads.install("report.zip", Buffer.from(zipSync({ "SKILL.md": strToU8(manifest), "big.md": new Uint8Array(MAX_SKILL_UPLOAD_BYTES + 1) })), () => false)).toThrow(/size limits/);
  });

  it("never overwrites an existing skill", () => {
    expect(() => uploads.install("SKILL.md", Buffer.from(manifest), () => true)).toThrow(/already exists/);
    uploads.install("SKILL.md", Buffer.from(manifest), () => false);
    expect(() => uploads.install("SKILL.md", Buffer.from(manifest), () => false)).toThrow(/already exists/);
  });

  it("removes only the selected uploaded skill", () => {
    uploads.install("SKILL.md", Buffer.from(manifest), () => false);
    expect(uploads.owns("monthly-report", path.join(uploads.root, "monthly-report"))).toBe(true);
    expect(uploads.owns("monthly-report", path.join(folder, "outside"))).toBe(false);
    expect(() => uploads.remove("../outside")).toThrow(/Invalid/);
    uploads.remove("monthly-report");
    expect(new SkillStore([uploads.root]).list()).toEqual([]);
  });

  it("rejects empty instructions, unsupported formats and invalid UTF-8", () => {
    expect(() => uploads.install("SKILL.md", Buffer.from("---\nname: empty\n---\n"), () => false)).toThrow(/instructions/);
    expect(() => uploads.install("skill.exe", Buffer.from(manifest), () => false)).toThrow(/Markdown/);
    expect(() => uploads.install("skill.md", Buffer.from([0xc3, 0x28]), () => false)).toThrow(/UTF-8/);
  });
});
