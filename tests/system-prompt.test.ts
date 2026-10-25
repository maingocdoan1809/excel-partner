import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { SystemPromptStore } from "../bridge/src/system-prompt.js";
import { EXCEL_TOOL_NAMES, isExcelToolName } from "../protocol/src/index.js";
import { RECOMMENDED_SYSTEM_PROMPT, SYSTEM_PROMPT } from "../providers/src/prompt.js";

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
