import fs from "node:fs";
import path from "node:path";
import { createHash, randomUUID } from "node:crypto";
import { SYSTEM_PROMPT } from "../../providers/src/prompt.js";
import type { SessionHost } from "../../protocol/src/history.js";
import { hostSystemPrompt } from "../../protocol/src/system-prompts.js";

export class PromptSettingsError extends Error {}
export class PromptSettingsConflict extends PromptSettingsError {}
/**
 * The instructions a person saves in Settings. Each Office application has its own: Excel's are the file itself, Word's and PowerPoint's
 * sit beside it (`system-prompt-word.txt`), so a prompt written for one application never reaches a chat in another.
 */
export class SystemPromptStore {
  private memory: string | null = null;
  private readonly others = new Map<SessionHost, SystemPromptStore>();
  /** `file` is Excel's; the stores of the other applications are made from it by `forHost`. */
  constructor(private readonly file: string | null, private readonly host: SessionHost = "excel") {}

  /** The store of one Office application (call it on the Excel store, the one the bridge creates). */
  forHost(host: SessionHost): SystemPromptStore {
    if (host === "excel") return this;
    let store = this.others.get(host);
    if (!store) {
      const parsed = this.file ? path.parse(this.file) : null;
      store = new SystemPromptStore(parsed ? path.join(parsed.dir, `${parsed.name}-${host}${parsed.ext}`) : null, host);
      this.others.set(host, store);
    }
    return store;
  }

  read(): { content: string; revision: string; customized: boolean } {
    let custom = this.file ? null : this.memory;
    if (this.file) {
      try { custom = fs.readFileSync(this.file, "utf8"); }
      catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error; }
    }
    // Earlier versions kept one prompt for every application, so Word's or PowerPoint's recommended text may be saved as Excel's. An
    // unchanged copy of it is not something the person wrote for Excel: ignore it, so Excel does not tell the model it is in PowerPoint.
    if (this.host === "excel" && custom !== null && (custom === hostSystemPrompt("word", "") || custom === hostSystemPrompt("powerpoint", ""))) custom = null;
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
