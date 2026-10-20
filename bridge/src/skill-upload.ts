import fs from "node:fs";
import path from "node:path";
import { unzipSync } from "fflate";
import { parseFrontmatter, SkillError } from "./skills.js";
import { MAX_SKILL_UPLOAD_BYTES, MAX_SKILL_TEXT_BYTES, MAX_SKILL_EXPANDED_BYTES } from "../../protocol/src/skill-limits.js";

export { MAX_SKILL_UPLOAD_BYTES } from "../../protocol/src/skill-limits.js";
const MAX_FILES = 64;
const MAX_TEXT_BYTES = MAX_SKILL_TEXT_BYTES;
const NAME = /^[a-z0-9][a-z0-9_-]{0,63}$/;

export class UserSkills {
  readonly root: string;
  constructor(root: string) { this.root = path.resolve(root); }

  install(filename: string, bytes: Buffer, exists: (name: string) => boolean): { name: string; description: string } {
    if (!bytes.length || bytes.length > MAX_SKILL_UPLOAD_BYTES) throw new SkillError("Upload a Markdown or ZIP file up to 5 MB.");
    let files: Record<string, Uint8Array>;
    let fallbackName = "";
    if (path.extname(filename).toLowerCase() === ".md") {
      files = { "SKILL.md": bytes };
      const stem = path.basename(filename, path.extname(filename));
      fallbackName = stem.toLowerCase() === "skill" ? "" : stem.toLowerCase();
    } else if (path.extname(filename).toLowerCase() === ".zip") {
      let expanded = 0;
      let entries = 0;
      const seen = new Set<string>();
      try {
        files = unzipSync(bytes, { filter: (file) => {
          this.validatePath(file.name.replace(/\/$/, ""));
          expanded += file.originalSize;
          if (++entries > MAX_FILES || expanded > MAX_SKILL_EXPANDED_BYTES || file.originalSize > MAX_TEXT_BYTES) throw new SkillError("The archive exceeds the skill file or size limits (5 MB per file, 10 MB expanded total).");
          const normalized = file.name.toLowerCase();
          if (seen.has(normalized)) throw new SkillError("The archive contains duplicate file paths.");
          seen.add(normalized);
          return !file.name.endsWith("/");
        } });
      } catch (error) {
        if (error instanceof SkillError) throw error;
        throw new SkillError("Could not read this ZIP file.");
      }
      const manifests = Object.keys(files).filter(name => name === "SKILL.md" || name.endsWith("/SKILL.md"));
      if (manifests.length !== 1) throw new SkillError("The archive must contain exactly one SKILL.md.");
      const manifest = manifests[0]!;
      const prefix = manifest.slice(0, -"SKILL.md".length);
      fallbackName = prefix ? prefix.slice(0, -1).split("/").at(-1)! : "";
      const normalized: Record<string, Uint8Array> = Object.create(null);
      for (const [name, content] of Object.entries(files)) {
        if (!name.startsWith(prefix)) throw new SkillError("Keep all skill files inside the folder containing SKILL.md.");
        normalized[name.slice(prefix.length)] = content;
      }
      files = normalized;
    } else throw new SkillError("Choose a Markdown (.md) or ZIP (.zip) skill file.");

    for (const [name, content] of Object.entries(files)) {
      this.validatePath(name);
      if (content.length > MAX_TEXT_BYTES) throw new SkillError("Skill files must be 5 MB or smaller.");
      if (name === "SKILL.md") {
        if (content.includes(0)) throw new SkillError("SKILL.md must contain text, not binary data.");
        try { new TextDecoder("utf-8", { fatal: true }).decode(content); }
        catch { throw new SkillError("SKILL.md must use UTF-8 text encoding."); }
      }
    }
    const manifest = Buffer.from(files["SKILL.md"] ?? []).toString("utf8");
    const { meta, body } = parseFrontmatter(manifest);
    const name = (meta.name ?? fallbackName).trim();
    if (!NAME.test(name) || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(name)) throw new SkillError("Add a valid name to SKILL.md frontmatter: 1–64 lowercase letters, numbers, hyphens or underscores.");
    if (!body.trim()) throw new SkillError("SKILL.md must include skill instructions.");
    if (exists(name)) throw new SkillError(`A skill named '${name}' already exists. Remove it before uploading a replacement.`);
    fs.mkdirSync(this.root, { recursive: true, mode: 0o700 });
    if (fs.lstatSync(this.root).isSymbolicLink()) throw new SkillError("The skill upload folder cannot be a symbolic link.");
    const target = path.join(this.root, name);
    if (fs.existsSync(target)) throw new SkillError("A skill folder with this name already exists.");
    const staging = fs.mkdtempSync(path.join(this.root, ".upload-"));
    try {
      for (const [relative, content] of Object.entries(files)) {
        const file = path.resolve(staging, relative);
        if (!file.startsWith(staging + path.sep)) throw new SkillError("Invalid skill file path.");
        fs.mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 });
        fs.writeFileSync(file, content, { flag: "wx", mode: 0o600 });
      }
      this.assertInside(staging);
      this.assertInside(target, false);
      fs.renameSync(staging, target);
    } finally {
      if (fs.existsSync(staging)) { this.assertInside(staging); fs.rmSync(staging, { recursive: true }); }
    }
    return { name, description: (meta.description ?? "").trim() };
  }

  owns(name: string, directory: string): boolean {
    return NAME.test(name) && path.resolve(directory) === path.resolve(this.root, name);
  }

  remove(name: string): void {
    if (!NAME.test(name)) throw new SkillError("Invalid skill name.");
    const target = path.resolve(this.root, name);
    this.assertInside(target);
    if (!fs.lstatSync(target).isDirectory()) throw new SkillError("Skill folder not found.");
    fs.rmSync(target, { recursive: true });
  }

  private assertInside(target: string, mustExist = true): void {
    const root = fs.realpathSync(this.root);
    const resolved = path.resolve(target);
    if (!resolved.startsWith(path.resolve(this.root) + path.sep)) throw new SkillError("Invalid skill folder.");
    if (mustExist && (fs.lstatSync(target).isSymbolicLink() || !fs.realpathSync(target).startsWith(root + path.sep))) throw new SkillError("The skill folder must remain inside the upload folder.");
  }

  private validatePath(name: string): void {
    const segments = name.split("/");
    if (!name || segments.length > 4 || segments.some(segment => !segment || segment === "." || segment === ".." || /[\x00-\x1f<>:"|?*\\]/.test(segment) || /[. ]$/.test(segment) || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(segment))) {
      throw new SkillError("Skill archives cannot contain absolute paths, traversal paths or unsupported file names.");
    }
  }
}
