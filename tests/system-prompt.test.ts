import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { SystemPromptStore } from "../bridge/src/system-prompt.js";
import { EXCEL_TOOL_NAMES, isExcelToolName } from "../protocol/src/index.js";
import { SYSTEM_PROMPT } from "../providers/src/prompt.js";
import { RECOMMENDED_SYSTEM_PROMPT, hostSystemPrompt, recommendedPromptFor } from "../protocol/src/system-prompts.js";

describe("system prompt settings", () => {
  it("persists a full override, detects conflicting edits and restores the shipped default", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "system-prompt-"));
    try {
      const file = path.join(dir, "system-prompt.txt");
      const store = new SystemPromptStore(file);
      const initial = store.read();
      expect(initial).toMatchObject({ content: SYSTEM_PROMPT, customized: false });
      const saved = store.update("User's complete system prompt", initial.revision);
      expect(new SystemPromptStore(file).read()).toEqual(saved);
      expect(() => store.update("Stale edit", initial.revision)).toThrow(/changed/);
      expect(() => store.update("a".repeat(65537), saved.revision)).toThrow(/64 KB/);
      expect(store.update(null, saved.revision)).toEqual(initial);
      expect(fs.existsSync(file)).toBe(false);
      const again = store.update("Custom", initial.revision);
      fs.unlinkSync(file);
      expect(store.read()).toEqual(initial);
      expect(again.customized).toBe(true);
    } finally { fs.rmSync(dir, { recursive: true, force: true }); }
  });

  it("keeps each Office application's instructions apart, so PowerPoint's never reach Excel", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "system-prompt-"));
    try {
      const file = path.join(dir, "system-prompt.txt");
      const excel = new SystemPromptStore(file);
      const powerpoint = excel.forHost("powerpoint");
      powerpoint.update("Only for decks", powerpoint.read().revision);
      expect(fs.existsSync(path.join(dir, "system-prompt-powerpoint.txt"))).toBe(true);
      expect(excel.read()).toMatchObject({ content: SYSTEM_PROMPT, customized: false });
      expect(new SystemPromptStore(file).forHost("powerpoint").read().content).toBe("Only for decks");
      expect(excel.forHost("word").read().customized).toBe(false);
      expect(excel.forHost("excel")).toBe(excel);
    } finally { fs.rmSync(dir, { recursive: true, force: true }); }
  });

  it("ignores a PowerPoint or Word preset that an earlier version saved as Excel's prompt", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "system-prompt-"));
    try {
      const file = path.join(dir, "system-prompt.txt");
      for (const host of ["powerpoint", "word"] as const) {
        fs.writeFileSync(file, hostSystemPrompt(host, ""));
        expect(new SystemPromptStore(file).read()).toMatchObject({ content: SYSTEM_PROMPT, customized: false });
      }
      fs.writeFileSync(file, hostSystemPrompt("powerpoint", "") + "\nMy own rule.");
      expect(new SystemPromptStore(file).read().customized).toBe(true);
    } finally { fs.rmSync(dir, { recursive: true, force: true }); }
  });

  it("offers Word and PowerPoint only their introduction as the preset, because their rules are added to whatever is saved", () => {
    for (const host of ["word", "powerpoint"] as const) {
      expect(recommendedPromptFor(host)).not.toContain("# Working in");
      expect(hostSystemPrompt(host, recommendedPromptFor(host))).toBe(hostSystemPrompt(host, ""));
    }
    expect(recommendedPromptFor("excel")).toBe(RECOMMENDED_SYSTEM_PROMPT);
  });

  it("defaults to no add-in prompt and persists an explicitly cleared override", () => {
    const store = new SystemPromptStore(null);
    const initial = store.read();
    expect(initial).toMatchObject({ content: "", customized: false });
    const custom = store.update("Custom instructions", initial.revision);
    const cleared = store.update("", custom.revision);
    expect(cleared).toMatchObject({ content: "", customized: true });
    expect(store.read()).toEqual(cleared);
    expect(store.update(null, cleared.revision)).toEqual(initial);
    const blank = store.update(" ", initial.revision);
    expect(blank.content).toBe(" ");
    expect(() => store.update("\0", blank.revision)).toThrow(/null characters/);
  });

  it("ships a recommended preset that fits the editor and names only real tools", () => {
    expect(SYSTEM_PROMPT).toBe("");
    expect(Buffer.byteLength(RECOMMENDED_SYSTEM_PROMPT, "utf8")).toBeLessThan(64 * 1024);
    expect(RECOMMENDED_SYSTEM_PROMPT).not.toContain("\0");
    const mentioned = new Set(RECOMMENDED_SYSTEM_PROMPT.match(/\b[a-z]+(?:_[a-z]+)+\b/g) ?? []);
    const toolLike = [...mentioned].filter((word) => EXCEL_TOOL_NAMES.some((name) => name.split("_")[0] === word.split("_")[0]));
    for (const word of toolLike) {
      if (/^(would_overwrite|formula_results)$/.test(word)) continue;
      expect(isExcelToolName(word), `${word} is not an Excel tool`).toBe(true);
    }
    for (const name of ["get_selection", "get_sheets", "read_range", "get_cell_ranges", "search_data", "read_range_image", "set_formula", "find_errors", "execute_office_js"]) {
      expect(RECOMMENDED_SYSTEM_PROMPT).toContain(name);
      expect(isExcelToolName(name)).toBe(true);
    }
  });

});
