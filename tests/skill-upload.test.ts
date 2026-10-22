import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { zipSync, strToU8 } from "fflate";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { UserSkills, MAX_SKILL_UPLOAD_BYTES } from "../bridge/src/skill-upload.js";
import { SkillStore } from "../bridge/src/skills.js";
import { MAX_SKILL_TEXT_BYTES } from "../protocol/src/skill-limits.js";

let folder: string;
let uploads: UserSkills;
const manifest = "---\nname: monthly-report\ndescription: Format a monthly report\n---\nUse the user-selected workbook.\n";

beforeEach(() => {
  folder = fs.mkdtempSync(path.join(os.tmpdir(), "user-skills-"));
  uploads = new UserSkills(path.join(folder, "skills"));
});
afterEach(() => fs.rmSync(folder, { recursive: true, force: true }));

describe("user skill uploads", () => {
  it("accepts the reported 300 KB Markdown and lets the skill engine read it", () => {
    const content = manifest + "a".repeat(300 * 1024);
    uploads.install("SKILL.md", Buffer.from(content), () => false);
    expect(new SkillStore([uploads.root]).read("monthly-report").content.length).toBeGreaterThan(300 * 1024);
  });

  it("accepts exactly 5 MB uploads and rejects one byte over the limit", () => {
    const bytes = Buffer.concat([Buffer.from(manifest), Buffer.alloc(MAX_SKILL_UPLOAD_BYTES - Buffer.byteLength(manifest), "a")]);
    uploads.install("SKILL.md", bytes, () => false);
    expect(new SkillStore([uploads.root]).read("monthly-report").content).toContain("user-selected workbook");
    expect(() => uploads.install("SKILL.md", Buffer.concat([bytes, Buffer.from("a")]), () => false)).toThrow(/5 MB/);
  });

  it("uploads and reads a full 5 MB text reference in a ZIP alongside SKILL.md", () => {
    const reference = new Uint8Array(MAX_SKILL_TEXT_BYTES).fill(97);
    const archive = zipSync({ "SKILL.md": strToU8(manifest), "references/large.txt": reference });
    uploads.install("reference.zip", Buffer.from(archive), () => false);
    expect(new SkillStore([uploads.root]).read("monthly-report", "references/large.txt").content.length).toBe(MAX_SKILL_TEXT_BYTES);
  });

  it("still bounds the total expanded ZIP size", () => {
    const reference = new Uint8Array(MAX_SKILL_TEXT_BYTES).fill(97);
    const archive = zipSync({ "SKILL.md": strToU8(manifest), "first.txt": reference, "second.txt": reference });
    expect(() => uploads.install("oversized.zip", Buffer.from(archive), () => false)).toThrow(/expanded total/);
  });

  it("accepts up to 1024 entries and reports the entry limit separately from size limits", () => {
    const files: Record<string, Uint8Array> = { "SKILL.md": strToU8(manifest) };
    for (let i = 0; i < 1023; i++) files[`references/section-${i}.md`] = strToU8("Reference");
    uploads.install("many-files.zip", Buffer.from(zipSync(files)), () => false);
    expect(fs.readFileSync(path.join(uploads.root, "monthly-report", "references", "section-1022.md"), "utf8")).toBe("Reference");
    files["references/one-too-many.md"] = strToU8("Reference");
    expect(() => uploads.install("too-many-files.zip", Buffer.from(zipSync(files)), () => false)).toThrow("maximum 1024 files and folders");
  });

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

  it("installs a packaged skill with Excel templates and inert Python/Shell scripts", () => {
    const archive = zipSync({
      "report/SKILL.md": strToU8(manifest),
      "report/references/style.md": strToU8("Use the packaged workbook."),
      "report/spec/tokens.json": strToU8('{"TITLE":"A1"}'),
      "report/templates/BD_Master.xlsx": strToU8("PKworkbook"),
      "report/tools/lint.py": strToU8('raise RuntimeError("Do not execute during upload")'),
      "report/install.sh": strToU8("exit 1"),
      "report/dist/bundle.zip": new Uint8Array([0, 1, 2]),
      "report/dist/templates/BD_Master.xlsx": strToU8("PKduplicate"),
    });
    uploads.install("report.zip", Buffer.from(archive), () => false);
    const store = new SkillStore([uploads.root], path.join(folder, "templates"));
    expect(store.read("monthly-report", "references/style.md").content).toBe("Use the packaged workbook.");
    expect(store.listTemplates("monthly-report")).toMatchObject([{ name: "BD_Master.xlsx", kind: "workbook" }]);
    expect(store.readTemplate("skill-template://monthly-report/BD_Master.xlsx").bytes.toString()).toBe("PKworkbook");
    expect(fs.readFileSync(path.join(uploads.root, "monthly-report", "tools", "lint.py"), "utf8")).toContain("Do not execute");
    expect(fs.statSync(path.join(uploads.root, "monthly-report", "install.sh")).mode & 0o111).toBe(0);
    expect(fs.existsSync(path.join(uploads.root, "monthly-report", "dist", "bundle.zip"))).toBe(false);
    expect(fs.readFileSync(path.join(uploads.root, "monthly-report", "dist", "templates", "BD_Master.xlsx"), "utf8")).toBe("PKduplicate");
  });

  it("skips nested ZIP files regardless of case or size, preserving neighboring files", () => {
    const archive = zipSync({
      "report/SKILL.md": strToU8(manifest),
      "report/backup.ZIP": zipSync({ "SKILL.md": strToU8("Another skill") }),
      "report/dist/large.Zip": new Uint8Array(MAX_SKILL_TEXT_BYTES + 1),
      "report/dist/notes.txt": strToU8("Keep this file"),
      "report/assets/zip/icon.png": new Uint8Array([0, 255, 1]),
    });
    uploads.install("report.zip", Buffer.from(archive), () => false);
    const target = path.join(uploads.root, "monthly-report");
    expect(fs.existsSync(path.join(target, "backup.ZIP"))).toBe(false);
    expect(fs.existsSync(path.join(target, "dist", "large.Zip"))).toBe(false);
    expect(fs.readFileSync(path.join(target, "dist", "notes.txt"), "utf8")).toBe("Keep this file");
    expect(fs.readFileSync(path.join(target, "assets", "zip", "icon.png"))).toEqual(Buffer.from([0, 255, 1]));
  });

  it("preserves arbitrary binary assets and unknown extensions without executing them", () => {
    const binary = new Uint8Array([0, 255, 128, 1]);
    const archive = zipSync({ "SKILL.md": strToU8(manifest), "assets/model.custom": binary,
      "diagram.pdf": binary, "run.exe": binary, "templates/file.xlsx": binary, "photo.png": binary });
    uploads.install("arbitrary.zip", Buffer.from(archive), () => false);
    for (const file of ["assets/model.custom", "diagram.pdf", "run.exe", "templates/file.xlsx", "photo.png"]) {
      expect(fs.readFileSync(path.join(uploads.root, "monthly-report", file))).toEqual(Buffer.from(binary));
    }
  });

  it("rejects archive paths that could escape the upload directory", () => {
    for (const invalid of ["../escape.md", "/absolute.md", "C:/escape.md", "references\\escape.md", "references/CON.txt"]) {
      const archive = zipSync({ "SKILL.md": strToU8(manifest), [invalid]: strToU8("Invalid") });
      expect(() => uploads.install("report.zip", Buffer.from(archive), () => false)).toThrow(`Invalid skill archive entry ${JSON.stringify(invalid)}`);
    }
    expect(fs.existsSync(uploads.root)).toBe(false);
  });

  it("accepts deeply nested skill folders and assets with valid file names", () => {
    const prefix = "package/export/skills/report/";
    const relative = "assets/templates/monthly/2026/report.xlsx";
    const archive = zipSync({
      [prefix.slice(0, -1)]: {},
      [`${prefix}SKILL.md`]: strToU8(manifest),
      [`${prefix}${relative}`]: strToU8("workbook"),
    });
    uploads.install("report.zip", Buffer.from(archive), () => false);
    expect(fs.readFileSync(path.join(uploads.root, "monthly-report", relative), "utf8")).toBe("workbook");
  });

  it("rejects ambiguous archives and excessive expansion", () => {
    expect(() => uploads.install("report.zip", Buffer.from(zipSync({ "README.md": strToU8("Missing skill") })), () => false)).toThrow(/exactly one SKILL/);
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
