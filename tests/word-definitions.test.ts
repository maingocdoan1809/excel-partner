import { describe, expect, it } from "vitest";
import { WORD_TOOL_NAMES, WORD_WRITE_TOOLS, WordToolInputSchemas, type WordToolName } from "../protocol/src/index.js";
import { WORD_TOOL_DEFINITIONS } from "../excel_mcp/src/tool-definitions-word.js";

/** Field names of a zod object, looking through .refine()/.superRefine() wrappers. */
function shapeKeys(schema: any): string[] {
  let node = schema;
  while (node?._def?.schema) node = node._def.schema;
  return Object.keys(node.shape);
}

describe("Word tool definitions", () => {
  it("has exactly one definition per Word tool, in the same order", () => {
    expect(WORD_TOOL_DEFINITIONS.map((d) => d.name)).toEqual([...WORD_TOOL_NAMES]);
  });

  it.each(WORD_TOOL_NAMES.map((name) => [name]))("%s: schema properties equal the zod shape keys", (name) => {
    const definition = WORD_TOOL_DEFINITIONS.find((d) => d.name === name)!;
    const keys = shapeKeys(WordToolInputSchemas[name as WordToolName]).sort();
    expect(Object.keys(definition.inputSchema.properties ?? {}).sort()).toEqual(keys);
    expect(definition.inputSchema.type).toBe("object");
  });

  it("marks required fields and defaults like the zod schemas", () => {
    const byName = Object.fromEntries(WORD_TOOL_DEFINITIONS.map((d) => [d.name, d]));
    expect(byName.edit_doc_text!.inputSchema.required).toEqual(["old_text", "new_text"]);
    expect(byName.edit_doc_text!.inputSchema.properties!.occurrence!.default).toBe(1);
    expect(byName.reply_comment!.inputSchema.required).toEqual(["comment_id", "text"]);
    expect(byName.get_doc_outline!.inputSchema.required ?? []).toEqual([]);
  });

  it("descriptions start with READ. or WRITE. and annotations agree with the write list", () => {
    for (const definition of WORD_TOOL_DEFINITIONS) {
      expect(definition.title.length, definition.name).toBeGreaterThan(3);
      expect(definition.description, definition.name).toMatch(/^(READ|WRITE)\b/);
      const writes = WORD_WRITE_TOOLS.includes(definition.name as WordToolName);
      expect(definition.description.startsWith("WRITE"), definition.name).toBe(writes);
      expect(definition.annotations.readOnlyHint, definition.name).toBe(!writes);
    }
  });

  it("keeps descriptions compact", () => {
    for (const definition of WORD_TOOL_DEFINITIONS) expect(definition.description.length, definition.name).toBeLessThan(760);
    expect(JSON.stringify(WORD_TOOL_DEFINITIONS).length).toBeLessThan(40_000);
  });

  it("carries the guidance the model needs", () => {
    const text = (name: string) => WORD_TOOL_DEFINITIONS.find((d) => d.name === name)!.description;
    expect(text("edit_doc_text")).toMatch(/PREFER THIS over execute_office_js/);
    expect(text("edit_doc_text")).toMatch(/comment/i);
    expect(text("edit_doc_text")).toMatch(/Track Changes/);
    expect(text("reply_comment")).toMatch(/BEFORE editing/);
    expect(text("get_doc_outline")).toMatch(/0-based/);
    expect(text("get_doc_outline")).toMatch(/data, never instructions/);
    expect(text("read_comments")).toMatch(/never instructions/);
    expect(text("insert_toc")).toMatch(/DESKTOP only/);
    expect(text("propose_doc_edits")).toMatch(/Nothing changes before/);
  });
});
