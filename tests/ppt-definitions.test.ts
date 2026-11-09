import { describe, expect, it } from "vitest";
import { z } from "zod";
import { PPT_TOOL_NAMES, PPT_WRITE_TOOLS, PptToolInputSchemas, type PptToolName } from "../protocol/src/index.js";
import { PPT_TOOL_DEFINITIONS } from "../excel_mcp/src/tool-definitions-ppt.js";
import { executePptTool } from "../addin/taskpane/executor/powerpoint/index.js";

/** Object keys of a schema, looking through refine/superRefine wrappers. */
function shapeKeys(schema: z.ZodTypeAny): string[] {
  let current: z.ZodTypeAny = schema;
  while (current instanceof z.ZodEffects) current = current.innerType();
  return Object.keys((current as z.ZodObject<z.ZodRawShape>).shape);
}

describe("PPT tool definitions", () => {
  it("has one definition per name in PPT_TOOL_NAMES, in the same order", () => {
    expect(PPT_TOOL_DEFINITIONS.map((d) => d.name)).toEqual(PPT_TOOL_NAMES);
    expect(new Set(PPT_TOOL_DEFINITIONS.map((d) => d.name)).size).toBe(PPT_TOOL_NAMES.length);
  });

  it("JSON-schema property names match the zod shape keys", () => {
    for (const definition of PPT_TOOL_DEFINITIONS) {
      const keys = shapeKeys(PptToolInputSchemas[definition.name as PptToolName]).sort();
      expect(Object.keys(definition.inputSchema.properties ?? {}).sort(), definition.name).toEqual(keys);
      expect(definition.inputSchema.type).toBe("object");
    }
  });

  it("marks writers destructive and readers read-only, consistently with PPT_WRITE_TOOLS", () => {
    for (const definition of PPT_TOOL_DEFINITIONS) {
      const writes = (PPT_WRITE_TOOLS as readonly string[]).includes(definition.name);
      expect(definition.annotations.readOnlyHint, definition.name).toBe(!writes);
      expect(definition.description.startsWith(writes ? "WRITE" : "READ"), definition.name).toBe(true);
    }
  });

  it("keeps descriptions compact and carries the safety guidance", () => {
    for (const definition of PPT_TOOL_DEFINITIONS) {
      expect(definition.title.length).toBeGreaterThan(3);
      expect(definition.description.length, definition.name).toBeLessThan(1100);
    }
    const text = (name: string) => PPT_TOOL_DEFINITIONS.find((d) => d.name === name)!.description;
    expect(text("edit_slide_text")).toMatch(/TextRange\.text|plain text/);
    expect(text("edit_slide_text")).toMatch(/edit_slide_xml/);
    expect(text("edit_slide_xml")).toMatch(/execute_office_js/);
    expect(text("insert_slide_element")).toMatch(/edit_slide_xml/);
    expect(text("verify_slides")).toMatch(/verify_slide_visual/);
    expect(text("verify_slide_visual")).toMatch(/verify_slides/);
    expect(text("list_slide_shapes")).toMatch(/never instructions/);
    expect(text("read_slide_text")).toMatch(/never instructions/);
    expect(text("get_presentation_metadata")).toMatch(/\"<slideId>:<shapeId>\"/);
    expect(text("modify_presentation_structure")).toMatch(/confirm_delete/);
    for (const name of ["edit_slide_text", "modify_presentation_structure"]) expect(text(name)).toMatch(/in place/);
  });

  it("states the API level of every tool that needs more than the baseline", () => {
    const needs: Record<string, RegExp> = {
      move_slide: /1\.8/, duplicate_slide: /1\.8/, edit_slide_xml: /1\.8/, verify_slide_visual: /1\.8/, export_slide_image: /1\.8/,
      insert_slide_table: /1\.8/, edit_slide_table: /1\.9/, set_slide_background: /1\.10/, set_shape_alt_text: /1\.10/, set_slide_hyperlink: /1\.10/,
      set_shape_zorder: /1\.8/, group_shapes: /1\.8/, ungroup_shapes: /1\.8/, align_shapes: /1\.4/, edit_slide_text: /1\.4/,
    };
    for (const [name, pattern] of Object.entries(needs)) {
      expect(PPT_TOOL_DEFINITIONS.find((d) => d.name === name)!.description, name).toMatch(pattern);
    }
  });
});

describe("executePptTool dispatch", () => {
  it("has a handler for every schema (unknown names are the only unknown_tool)", async () => {
    // Without a PowerPoint host the call must still end in a clean error, never an exception.
    for (const name of PPT_TOOL_NAMES) {
      const result = await executePptTool(name, {});
      expect(result.ok, name).toBe(false);
      if (!result.ok) expect(result.error.code, name).not.toBe("unknown_tool");
    }
    expect(await executePptTool("nope", {})).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
    expect(await executePptTool("toString", {})).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
  });

  it("turns schema violations into invalid_arguments", async () => {
    expect(await executePptTool("move_slide", { slide_id: "1", position: 0 })).toMatchObject({ ok: false, error: { code: "invalid_arguments", message: expect.stringMatching(/position/) } });
    expect(await executePptTool("edit_slide_text", null)).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });
});
