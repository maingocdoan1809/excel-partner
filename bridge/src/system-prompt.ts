import fs from "node:fs";
import path from "node:path";
import { createHash, randomUUID } from "node:crypto";
import { SYSTEM_PROMPT } from "../../providers/src/prompt.js";

export class PromptSettingsError extends Error {}
export class PromptSettingsConflict extends PromptSettingsError {}
export class SystemPromptStore {
  private memory: string | null = null;
  constructor(private readonly file: string | null) {}

  read(): { content: string; revision: string; customized: boolean } {
    let custom = this.file ? null : this.memory;
    if (this.file) {
      try { custom = fs.readFileSync(this.file, "utf8"); }
      catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error; }
    }
    const content = custom ?? SYSTEM_PROMPT;
    return { content, revision: createHash("sha256").update(JSON.stringify([custom !== null, content])).digest("hex"), customized: custom !== null };
  }

  update(content: string | null, revision: string): ReturnType<SystemPromptStore["read"]> {
    if (this.read().revision !== revision) throw new PromptSettingsConflict("The system prompt changed elsewhere. Reopen the editor before saving.");
    if (content !== null && (Buffer.byteLength(content, "utf8") > 64 * 1024 || content.includes("\0"))) throw new PromptSettingsError("The system prompt must be at most 64 KB and contain no null characters. Leave it empty for no add-in instructions.");
    if (this.file) {
      if (content === null) {
        try { fs.unlinkSync(this.file); }
        catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error; }
      } else {
        fs.mkdirSync(path.dirname(this.file), { recursive: true, mode: 0o700 });
        const temporary = this.file + "." + randomUUID() + ".tmp";
        try { fs.writeFileSync(temporary, content, { flag: "wx", mode: 0o600 }); fs.renameSync(temporary, this.file); }
        finally { if (fs.existsSync(temporary)) fs.unlinkSync(temporary); }
      }
    }
    this.memory = content;
    return this.read();
  }
}
