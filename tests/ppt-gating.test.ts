import { describe, expect, it } from "vitest";
import { executePptTool, getPptSelection, getPptSelectionLabel } from "../frontend/taskpane/executor/powerpoint/index.js";
import { buildDeck } from "./support/ppt-fixture.js";
import { installFakePowerPoint } from "./support/fake-powerpoint.js";

/** One representative valid call per tool (and per sub-operation), built from the fixture's ids. */
function calls(f: ReturnType<typeof buildDeck>): Array<[string, unknown]> {
  const { agenda, data, title } = f;
  return [
    ["get_presentation_metadata", {}],
    ["list_layouts", {}],
    ["add_slide", { layout_id: "l2#", position: 1 }],
    ["list_slide_shapes", { slide_id: agenda }],
    ["list_slide_shapes", {}],
    ["read_slide_text", { slide_id: agenda }],
    ["read_slide_text", { ref: `${agenda}:2` }],
    ["edit_slide_text", { ref: `${agenda}:2`, text: "x" }],
    ["edit_slide_text", { edits: [{ ref: `${agenda}:2`, text: "x" }, { ref: `${agenda}:3`, text: "y" }] }],
    ["edit_slide_xml", { slide_id: title, code: "return 1;" }],
    ["insert_slide_element", { slide_id: agenda, element: { type: "textbox", text: "t", left: 1, top: 1, width: 10, height: 10 } }],
    ["insert_slide_element", { slide_id: agenda, element: { type: "shape", shape: "Rectangle", left: 1, top: 1, width: 10, height: 10 } }],
    ["insert_slide_element", { slide_id: agenda, element: { type: "line", left: 1, top: 1, width: 10, height: 0 } }],
    ["insert_slide_element", { slide_id: agenda, element: { type: "table", values: [["a"]] } }],
    ["remove_slide_element", { slide_id: agenda, shape_ids: ["4"] }],
    ["duplicate_slide", { slide_id: agenda }],
    ["move_slide", { slide_id: data, position: 1 }],
    ["modify_presentation_structure", { operation: "add", layout_id: "l1#" }],
    ["modify_presentation_structure", { operation: "move", slide_id: data, position: 1 }],
    ["modify_presentation_structure", { operation: "delete", slide_id: data, confirm_delete: true }],
    ["verify_slides", {}],
    ["verify_slide_visual", { slide_id: agenda }],
    ["export_slide_image", { slide_id: agenda }],
    ["insert_slide_table", { slide_id: agenda, values: [["a", "b"]] }],
    ["edit_slide_table", { ref: `${data}:3`, operation: "read" }],
    ["edit_slide_table", { ref: `${data}:3`, operation: "set_cell", row: 0, column: 0, text: "z" }],
    ["edit_slide_table", { ref: `${data}:3`, operation: "add_rows", row: 1 }],
    ["edit_slide_table", { ref: `${data}:3`, operation: "delete_rows", row: 1 }],
    ["edit_slide_table", { ref: `${data}:3`, operation: "add_columns", column: 1 }],
    ["edit_slide_table", { ref: `${data}:3`, operation: "delete_columns", column: 1 }],
    ["set_slide_background", { slide_id: agenda, color: "#123456" }],
    ["set_shape_alt_text", { ref: `${agenda}:4`, description: "d" }],
    ["set_slide_hyperlink", { ref: `${agenda}:2`, url: "https://example.com" }],
    ["set_shape_zorder", { ref: `${agenda}:2`, action: "bring_to_front" }],
    ["group_shapes", { slide_id: agenda, shape_ids: ["2", "3"] }],
    ["ungroup_shapes", { ref: `${agenda}:2` }],
    ["align_shapes", { slide_id: agenda, shape_ids: ["2", "3"], align: "left" }],
    ["align_shapes", { slide_id: agenda, shape_ids: ["2", "3"], align: "center", relative_to: "slide" }],
  ];
}

describe("API-set gating sweep", () => {
  for (let api = 2; api <= 10; api++) {
    it(`never calls an API above PowerPointApi 1.${api} and never throws`, async () => {
      // slide ids are deterministic (256#, 257#, ...), so arguments built from one deck fit every fresh deck
      for (const [tool, args] of calls(buildDeck(api))) {
        // a fresh deck per call keeps one call's changes from hiding another's gate
        const fixture = buildDeck(api);
        fixture.deck.selection = { slideIds: [fixture.agenda], shapes: [[fixture.agenda, "2"]] };
        const installed = installFakePowerPoint(fixture.deck);
        try {
          const result = await executePptTool(tool, args);
          expect(typeof result.ok, tool).toBe("boolean");
          if (!result.ok) expect(["unsupported", "invalid_arguments", "confirmation_required"], `${tool}@1.${api}: ${result.error.message}`).toContain(result.error.code);
          expect(fixture.deck.violations, `${tool}@1.${api}`).toEqual([]);
        } finally {
          installed.restore();
        }
      }
    });
  }

  it("selection works at every level without gate violations", async () => {
    for (let api = 2; api <= 10; api++) {
      const f = buildDeck(api);
      f.deck.selection = { slideIds: [f.agenda], shapes: [[f.agenda, "2"]] };
      const installed = installFakePowerPoint(f.deck);
      try {
        const snapshot = await getPptSelection();
        expect(snapshot.slides[0]?.id).toBe(f.agenda);
        expect(snapshot.shapes.length).toBe(api >= 5 ? 1 : 0);
        expect((await getPptSelectionLabel()).label.length).toBeGreaterThan(0);
        expect(f.deck.violations).toEqual([]);
      } finally {
        installed.restore();
      }
    }
  });
});
