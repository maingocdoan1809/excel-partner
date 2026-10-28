import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import os from "node:os";
import { AUTOMATION_DIR } from "./config.js";
import { loadManifest, saveManifest, automationDir } from "./manifest.js";
import { QaError, sha256Text, writeFileAtomic } from "./util.js";

export const HARNESS_VERSION = "1";
/** Files copied into every project, relative to automation/. They are trusted code, so changes are detected. */
export const HARNESS_FILES = [
  "playwright.config.js", "_harness/launch.js", "_harness/redact.js", "_harness/reporter.js",
  "_harness/fixtures.js", "_harness/odoo.js", "_harness/login.js", "_harness/page-state.js",
] as const;
export const PLAYWRIGHT_VERSION = "1.64.1";

function packageRoot(): string {
  let dir = path.dirname(fileURLToPath(import.meta.url));
  for (let i = 0; i < 8; i++) {
    if (fs.existsSync(path.join(dir, "qa", "templates"))) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  throw new QaError("The QA templates are missing from this installation. Reinstall excel-partner.", "missing_templates");
}

export const templatesDir = (): string => path.join(packageRoot(), "qa", "templates");
const templateText = (relative: string): string => fs.readFileSync(path.join(templatesDir(), "project", AUTOMATION_DIR, relative), "utf8");
export const projectTemplate = (relative: string): string => fs.readFileSync(path.join(templatesDir(), "project", relative), "utf8");
export const exampleTemplate = (name: string): string => fs.readFileSync(path.join(templatesDir(), "examples", name), "utf8");

export function packagedHashes(): Record<string, string> {
  return Object.fromEntries(HARNESS_FILES.map(file => [file, sha256Text(templateText(file))]));
}

/** Copy the harness into the project and record its hashes. Existing copies are replaced. */
export function installHarness(projectDir: string): void {
  for (const file of HARNESS_FILES) writeFileAtomic(path.join(automationDir(projectDir), file), templateText(file).replace(/\r\n/g, "\n"));
  const manifest = loadManifest(projectDir);
  manifest.harness = { version: HARNESS_VERSION, files: packagedHashes() };
  saveManifest(projectDir, manifest);
}

export interface HarnessCheck { ok: boolean; outdated: boolean; problems: string[] }

export function checkHarness(projectDir: string): HarnessCheck {
  const manifest = loadManifest(projectDir);
  const problems: string[] = [];
  if (!manifest.harness) return { ok: false, outdated: false, problems: ["The harness is not installed. Run: excel-partner qa upgrade"] };
  for (const file of HARNESS_FILES) {
    const expected = manifest.harness.files[file];
    const target = path.join(automationDir(projectDir), file);
    if (!fs.existsSync(target)) { problems.push(`${AUTOMATION_DIR}/${file} is missing`); continue; }
    if (!expected || sha256Text(fs.readFileSync(target, "utf8")) !== expected) problems.push(`${AUTOMATION_DIR}/${file} was modified`);
  }
  const packaged = packagedHashes();
  const outdated = manifest.harness.version !== HARNESS_VERSION || HARNESS_FILES.some(file => manifest.harness!.files[file] !== packaged[file]);
  return { ok: problems.length === 0, outdated, problems };
}

export const QA_SKILLS = ["qa-testcases", "qa-automate", "qa-heal", "qa-report"] as const;
export const userSkillsDir = (): string => process.env.QA_SKILLS_DIR || path.join(os.homedir(), ".excel-partner", "skills");

export interface SkillInstall { name: string; action: "installed" | "updated" | "kept" | "unchanged" }

/**
 * Copy the QA skills into the user's skill folder. Each installed copy is stamped with the hash of what was written,
 * so a later version can replace a copy nobody touched, while a skill the user has edited is kept unless `force` is set.
 */
export function installSkills(force = false, target = userSkillsDir()): SkillInstall[] {
  const normalise = (text: string) => text.split(String.fromCharCode(13, 10)).join(String.fromCharCode(10));
  return QA_SKILLS.map((name): SkillInstall => {
    const source = normalise(fs.readFileSync(path.join(templatesDir(), "skills", name, "SKILL.md"), "utf8"));
    const file = path.join(target, name, "SKILL.md");
    const stamp = path.join(target, name, ".installed-sha256");
    const existing = fs.existsSync(file) ? normalise(fs.readFileSync(file, "utf8")) : null;
    if (existing === source) {
      if (!fs.existsSync(stamp)) writeFileAtomic(stamp, sha256Text(source));
      return { name, action: "unchanged" };
    }
    const untouched = existing !== null && fs.existsSync(stamp) && fs.readFileSync(stamp, "utf8").trim() === sha256Text(existing);
    if (existing !== null && !force && !untouched) return { name, action: "kept" };
    writeFileAtomic(file, source);
    writeFileAtomic(stamp, sha256Text(source));
    return { name, action: existing === null ? "installed" : "updated" };
  });
}

export function templateWorkbookPath(): string { return path.join(templatesDir(), "TestPlan.xlsx"); }
