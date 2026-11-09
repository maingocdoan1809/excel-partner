import { describe, expect, it } from "vitest";
import { computeAlignment } from "../addin/taskpane/executor/powerpoint/shapes.js";
import { call, fails, ok, useFakeDeck } from "./support/ppt-fixture.js";
import { executeWithPermission } from "../addin/taskpane/permissions.js";

const fx = useFakeDeck();
const shapeOf = (slideId: string, shapeId: string) => fx().deck.slideById(slideId).shapes.find((s) => s.id === shapeId)!;

describe("add_slide / move_slide / modify_presentation_structure", () => {
  it("Ask cannot be bypassed by a model-supplied deletion confirmation", async () => {
    let asked = 0;
    const result = await executeWithPermission("modify_presentation_structure",
      { operation: "delete", slide_id: fx().agenda, confirm_delete: true }, "ask", call,
      async () => { asked++; expect(fx().deck.ids()).toHaveLength(3); return false; });
    expect(asked).toBe(1);
    expect(result).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(fx().deck.ids()).toHaveLength(3);
  });

  it("Ask deletes only after approval and Auto supplies the confirmation itself", async () => {
    const result = await executeWithPermission("modify_presentation_structure",
      { operation: "delete", slide_id: fx().agenda }, "ask", call,
      async () => { expect(fx().deck.ids()).toHaveLength(3); return true; });
    expect(result.ok).toBe(true);
    expect(fx().deck.ids()).toHaveLength(2);
    const auto = await executeWithPermission("modify_presentation_structure",
      { operation: "delete", slide_id: fx().data }, "auto", call,
      async () => { throw new Error("Auto must not ask"); });
    expect(auto.ok).toBe(true);
    expect(fx().deck.ids()).toHaveLength(1);
  });

  it("adds a slide from a layout and returns its placeholders", async () => {
    const data = await ok("add_slide", { layout_id: "l2#" });
    expect(data).toMatchObject({ position: 4, slide_count: 4, layout: "Title and Content" });
    expect(data.summary).toMatch(/Added slide 4/);
    expect(data.shapes.map((s: any) => [s.name, s.placeholder])).toEqual([["Title 1", "Title"], ["Content Placeholder 2", "Body"]]);
    expect(data.shapes[0].ref).toBe(`${data.slide_id}:2`);
  });

  it("places the slide at a position with 1.8, and at the first position", async () => {
    const second = await ok("add_slide", { layout_id: "l3#", position: 2 });
    expect(second.position).toBe(2);
    const first = await ok("add_slide", { layout_id: "l3#", position: 1 });
    expect(first.position).toBe(1);
    expect(fx().deck.ids()[0]).toBe(first.slide_id);
  });

  it("reports the real position below 1.8 instead of moving", async () => {
    fx().deck.api = 5;
    const data = await ok("add_slide", { layout_id: "l3#", position: 1 });
    expect(data.position).toBe(4);
    expect(data.warnings[0]).toMatch(/1\.8/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("names the layouts when layout_id is wrong", async () => {
    const result = await call("add_slide", { layout_id: "nope" });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toMatchObject({ code: "invalid_arguments", detail: { layouts: expect.arrayContaining([{ master_id: "m1#", id: "l1#", name: "Title Slide" }]) } });
  });

  it("moves a slide, clamps past the end and is unsupported below 1.8", async () => {
    const { title, agenda, data } = fx();
    const moved = await ok("move_slide", { slide_id: data, position: 1 });
    expect(moved).toMatchObject({ from: 3, position: 1 });
    expect(fx().deck.ids()).toEqual([data, title, agenda]);
    const clamped = await ok("move_slide", { slide_id: data, position: 99 });
    expect(clamped.position).toBe(3);
    expect(clamped.note).toMatch(/end/);
    fx().deck.api = 7;
    expect(await fails("move_slide", { slide_id: data, position: 1 }, "unsupported")).toMatch(/1\.8/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("delete needs confirmation, never removes every slide, and rejects unknown ids", async () => {
    const { title, agenda, data } = fx();
    expect(await fails("modify_presentation_structure", { operation: "delete", slide_id: agenda }, "confirmation_required")).toMatch(/approval/);
    expect(fx().deck.ids()).toHaveLength(3);
    expect(await fails("modify_presentation_structure", { operation: "delete", slide_ids: [title, agenda, data], confirm_delete: true }, "invalid_arguments")).toMatch(/every slide/);
    expect(await fails("modify_presentation_structure", { operation: "delete", slide_id: "zzz", confirm_delete: true }, "invalid_arguments")).toMatch(/does not exist/);
    const done = await ok("modify_presentation_structure", { operation: "delete", slide_id: agenda, confirm_delete: true });
    expect(done).toMatchObject({ deleted: [agenda], slide_count: 2 });
    expect(fx().deck.ids()).toEqual([title, data]);
  });

  it("add and move go through the same paths", async () => {
    const added = await ok("modify_presentation_structure", { operation: "add", layout_id: "l1#", position: 1 });
    expect(added).toMatchObject({ position: 1, layout: "Title Slide" });
    const moved = await ok("modify_presentation_structure", { operation: "move", slide_id: fx().data, position: 1 });
    expect(moved.position).toBe(1);
  });
});

describe("edit_slide_text", () => {
  it("replaces the text of one shape and reports what was there", async () => {
    const { agenda } = fx();
    const data = await ok("edit_slide_text", { ref: `${agenda}:2`, text: "Plan\nfor Q4" });
    expect(data).toMatchObject({ edited: 1, failed: 0, results: [{ ref: `${agenda}:2`, ok: true, before: "Agenda", characters: 11 }] });
    expect(data.summary).toMatch(/Edited the text of/);
    expect(shapeOf(agenda, "2").text).toBe("Plan\nfor Q4");
  });

  it("edits several shapes in one sync and reports the ones it skipped", async () => {
    const { agenda, data: dataSlide } = fx();
    const result = await ok("edit_slide_text", {
      edits: [
        { ref: `${agenda}:2`, text: "A" },
        { ref: `${agenda}:3`, text: "B" },
        { ref: `${agenda}:4`, text: "picture" },
        { ref: `${dataSlide}:3`, text: "table" },
        { ref: `${agenda}:99`, text: "missing" },
        { ref: `bad`, text: "x" },
      ],
    }).catch((e) => e);
    // the schema rejects the malformed ref up front; the tool call itself never starts
    expect(String(result)).toMatch(/invalid arguments/);
    const partial = await ok("edit_slide_text", {
      edits: [
        { ref: `${agenda}:2`, text: "A" },
        { ref: `${agenda}:3`, text: "B" },
        { ref: `${agenda}:4`, text: "picture" },
        { ref: `${dataSlide}:3`, text: "table" },
        { ref: `${agenda}:99`, text: "missing" },
        { ref: `nope:1`, text: "no slide" },
      ],
    });
    expect(partial).toMatchObject({ edited: 2, failed: 4 });
    expect(partial.results.map((r: any) => r.ok)).toEqual([true, true, false, false, false, false]);
    expect(partial.results[2].error).toMatch(/Image shape/);
    expect(partial.results[3].error).toMatch(/edit_slide_table/);
    expect(partial.results[4].error).toMatch(/does not exist/);
    expect(partial.results[5].error).toMatch(/nope:1 does not exist/);
    expect(shapeOf(agenda, "3").text).toBe("B");
  });

  it("fails when nothing could be edited", async () => {
    expect(await fails("edit_slide_text", { ref: `${fx().agenda}:4`, text: "x" }, "invalid_arguments")).toMatch(/Image shape/);
  });

  it("is unsupported below 1.4", async () => {
    fx().deck.api = 3;
    expect(await fails("edit_slide_text", { ref: `${fx().agenda}:2`, text: "x" }, "unsupported")).toMatch(/1\.4/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("recovers one by one when a placeholder has no text frame", async () => {
    const id = fx().deck.addSlideWith([{ name: "Pic", placeholder: "Body", textBroken: true }, { name: "Note", text: "old" }]);
    const result = await ok("edit_slide_text", { edits: [{ ref: `${id}:2`, text: "a" }, { ref: `${id}:3`, text: "new" }] });
    expect(result.results.map((r: any) => r.ok)).toEqual([false, true]);
    expect(shapeOf(id, "3").text).toBe("new");
  });
});

describe("insert_slide_element / remove_slide_element", () => {
  it("adds a styled text box", async () => {
    const data = await ok("insert_slide_element", {
      slide_id: fx().agenda,
      element: { type: "textbox", text: "Footnote", left: 40, top: 480, width: 400, height: 30, style: { size: 12, bold: true, color: "#333333", align: "Left" }, name: "Footnote" },
    });
    expect(data).toMatchObject({ name: "Footnote", type: "TextBox", left: 40, top: 480, width: 400, height: 30 });
    const stored = shapeOf(fx().agenda, data.shape_id);
    expect(stored).toMatchObject({ text: "Footnote", fontSize: 12, fontColor: "#333333" });
  });

  it("adds a filled shape with text, matching the type name case-insensitively", async () => {
    const data = await ok("insert_slide_element", {
      slide_id: fx().agenda,
      element: { type: "shape", shape: "roundrectangle", left: 10, top: 10, width: 100, height: 50, fill: "#FF0000", line: "#000000", line_weight: 2, text: "Go" },
    });
    expect(data.type).toBe("GeometricShape");
    expect(shapeOf(fx().agenda, data.shape_id)).toMatchObject({ shapeKind: "RoundRectangle", fillColor: "#FF0000", lineWeight: 2, text: "Go" });
  });

  it("rejects an unknown shape type with examples", async () => {
    expect(await fails("insert_slide_element", { slide_id: fx().agenda, element: { type: "shape", shape: "Blob", left: 0, top: 0, width: 10, height: 10 } }, "invalid_arguments")).toMatch(/Rectangle/);
  });

  it("adds a line and a table", async () => {
    const line = await ok("insert_slide_element", { slide_id: fx().agenda, element: { type: "line", left: 0, top: 100, width: 300, height: 0, line: "#999999" } });
    expect(line.type).toBe("Line");
    const table = await ok("insert_slide_element", { slide_id: fx().agenda, element: { type: "table", values: [["a", "b"], ["c", "d"]] } });
    expect(table).toMatchObject({ rows: 2, columns: 2 });
    expect(shapeOf(fx().agenda, table.shape_id).table).toEqual([["a", "b"], ["c", "d"]]);
  });

  it("gates shapes at 1.4 and tables at 1.8", async () => {
    fx().deck.api = 3;
    expect(await fails("insert_slide_element", { slide_id: fx().agenda, element: { type: "textbox", text: "x", left: 0, top: 0, width: 1, height: 1 } }, "unsupported")).toMatch(/1\.4/);
    fx().deck.api = 7;
    expect(await fails("insert_slide_element", { slide_id: fx().agenda, element: { type: "table", values: [["a"]] } }, "unsupported")).toMatch(/1\.8/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("removes shapes and reports the ones that were not there", async () => {
    const { agenda } = fx();
    const data = await ok("remove_slide_element", { slide_id: agenda, shape_ids: ["4", "77"] });
    expect(data).toMatchObject({ removed: [{ ref: `${agenda}:4`, name: "Logo" }], not_found: ["77"] });
    expect(fx().deck.slideById(agenda).shapes).toHaveLength(2);
    expect(await fails("remove_slide_element", { slide_id: agenda, shape_ids: ["77"] }, "invalid_arguments")).toMatch(/none of the shapes/);
  });
});

describe("tables", () => {
  const ref = () => `${fx().data}:3`;

  it("inserts a table with values and position", async () => {
    const data = await ok("insert_slide_table", { slide_id: fx().agenda, values: [["h1", "h2"], ["1", "2"]], left: 10, top: 20, width: 300, height: 80, name: "Results" });
    expect(data).toMatchObject({ rows: 2, columns: 2, name: "Results", left: 10, width: 300 });
    expect(data.ref).toBe(`${fx().agenda}:${data.shape_id}`);
  });

  it("reads and edits cells", async () => {
    expect(await ok("edit_slide_table", { ref: ref(), operation: "read" })).toMatchObject({ rows: 3, columns: 2, values: [["Region", "Sales"], ["EMEA", "10"], ["APAC", "12"]], truncated: false });
    await ok("edit_slide_table", { ref: ref(), operation: "set_cell", row: 1, column: 1, text: "11" });
    expect(shapeOf(fx().data, "3").table![1]).toEqual(["EMEA", "11"]);
    expect(await fails("edit_slide_table", { ref: ref(), operation: "set_cell", row: 9, column: 0, text: "x" }, "invalid_arguments")).toMatch(/outside the table/);
  });

  it("adds and deletes rows and columns", async () => {
    const added = await ok("edit_slide_table", { ref: ref(), operation: "add_rows", row: 1, count: 2 });
    expect(added).toMatchObject({ rows: 5, columns: 2, before: { rows: 3, columns: 2 } });
    expect(shapeOf(fx().data, "3").table![1]).toEqual(["", ""]);
    await ok("edit_slide_table", { ref: ref(), operation: "delete_rows", row: 1, count: 2 });
    expect(shapeOf(fx().data, "3").table).toEqual([["Region", "Sales"], ["EMEA", "10"], ["APAC", "12"]]);
    await ok("edit_slide_table", { ref: ref(), operation: "add_columns", column: 2 });
    expect(shapeOf(fx().data, "3").table![0]).toEqual(["Region", "Sales", ""]);
    await ok("edit_slide_table", { ref: ref(), operation: "delete_columns", column: 2 });
    expect(shapeOf(fx().data, "3").table![0]).toEqual(["Region", "Sales"]);
  });

  it("refuses to delete everything and non-tables", async () => {
    expect(await fails("edit_slide_table", { ref: ref(), operation: "delete_rows", row: 0, count: 3 }, "invalid_arguments")).toMatch(/every row/);
    expect(await fails("edit_slide_table", { ref: `${fx().data}:2`, operation: "read" }, "invalid_arguments")).toMatch(/not a table/);
  });

  it("gates reading at 1.8 and row/column changes at 1.9", async () => {
    fx().deck.api = 8;
    expect((await call("edit_slide_table", { ref: ref(), operation: "read" })).ok).toBe(true);
    expect(await fails("edit_slide_table", { ref: ref(), operation: "add_rows", row: 1 }, "unsupported")).toMatch(/1\.9/);
    fx().deck.api = 7;
    expect(await fails("edit_slide_table", { ref: ref(), operation: "read" }, "unsupported")).toMatch(/1\.8/);
    expect(await fails("insert_slide_table", { slide_id: fx().agenda, values: [["a"]] }, "unsupported")).toMatch(/1\.8/);
    expect(fx().deck.violations).toEqual([]);
  });
});

describe("z-order, grouping, alignment", () => {
  it("changes the stacking order and reports the layers", async () => {
    const { agenda } = fx();
    const data = await ok("set_shape_zorder", { ref: `${agenda}:2`, action: "bring_to_front" });
    expect(data).toMatchObject({ z_before: 1, z_after: 3 });
    expect(fx().deck.slideById(agenda).shapes.map((s) => s.id)).toEqual(["3", "4", "2"]);
    expect(await ok("set_shape_zorder", { ref: `${agenda}:2`, action: "send_backward" })).toMatchObject({ z_before: 3, z_after: 2 });
  });

  it("groups and ungroups", async () => {
    const { agenda } = fx();
    const grouped = await ok("group_shapes", { slide_id: agenda, shape_ids: ["2", "3"], name: "Header" });
    expect(grouped).toMatchObject({ name: "Header", grouped: ["2", "3"] });
    expect(fx().deck.slideById(agenda).shapes.map((s) => s.type)).toEqual(["Image", "Group"]);
    const released = await ok("ungroup_shapes", { ref: grouped.ref });
    expect(released.released.map((s: any) => s.id)).toEqual(["2", "3"]);
    expect(await fails("ungroup_shapes", { ref: `${agenda}:2` }, "invalid_arguments")).toMatch(/not a group/);
    expect(await fails("group_shapes", { slide_id: agenda, shape_ids: ["2", "99"] }, "invalid_arguments")).toMatch(/99/);
  });

  it("gates z-order and grouping at 1.8", async () => {
    fx().deck.api = 7;
    expect(await fails("set_shape_zorder", { ref: `${fx().agenda}:2`, action: "send_to_back" }, "unsupported")).toMatch(/1\.8/);
    expect(await fails("group_shapes", { slide_id: fx().agenda, shape_ids: ["2", "3"] }, "unsupported")).toMatch(/1\.8/);
    expect(await fails("ungroup_shapes", { ref: `${fx().agenda}:2` }, "unsupported")).toMatch(/1\.8/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("aligns shapes to each other and to the slide", async () => {
    const id = fx().deck.addSlideWith([
      { name: "A", left: 100, top: 50, width: 100, height: 40 },
      { name: "B", left: 300, top: 120, width: 200, height: 40 },
      { name: "C", left: 700, top: 300, width: 100, height: 40 },
    ]);
    const left = await ok("align_shapes", { slide_id: id, shape_ids: ["2", "3", "4"], align: "left" });
    expect(left.moved.map((m: any) => [m.name, m.left])).toEqual([["B", 100], ["C", 100]]);
    await ok("align_shapes", { slide_id: id, shape_ids: ["2", "3", "4"], align: "right", relative_to: "slide" });
    expect(["2", "3", "4"].map((sid) => shapeOf(id, sid).left)).toEqual([860, 760, 860]);
    const nothing = await ok("align_shapes", { slide_id: id, shape_ids: ["2"], align: "top", relative_to: "shapes" });
    expect(nothing.moved).toEqual([]);
  });

  it("distributes and rejects too few shapes", async () => {
    const id = fx().deck.addSlideWith([
      { name: "A", left: 0, top: 0, width: 100, height: 40 },
      { name: "B", left: 120, top: 0, width: 100, height: 40 },
      { name: "C", left: 700, top: 0, width: 100, height: 40 },
    ]);
    await ok("align_shapes", { slide_id: id, shape_ids: ["2", "3", "4"], distribute: "horizontal" });
    expect(["2", "3", "4"].map((sid) => shapeOf(id, sid).left)).toEqual([0, 350, 700]);
    expect(await fails("align_shapes", { slide_id: id, shape_ids: ["2", "3"], distribute: "vertical" }, "invalid_arguments")).toMatch(/at least 3/);
  });

  it("computeAlignment: vertical centring and distribution across the slide", () => {
    const boxes = [
      { id: "a", left: 0, top: 10, width: 50, height: 20 },
      { id: "b", left: 0, top: 200, width: 50, height: 40 },
    ];
    const middle = computeAlignment(boxes, { align: "middle", relativeTo: "shapes" });
    expect(middle.get("a")!.top).toBe(125 - 10);
    expect(middle.get("b")!.top).toBe(125 - 20);
    const spread = computeAlignment(boxes, { distribute: "vertical", relativeTo: "slide", slide: { width: 960, height: 540 } });
    expect(spread.get("a")!.top).toBe(0);
    expect(spread.get("b")!.top).toBe(500);
    expect(() => computeAlignment(boxes, { align: "left", relativeTo: "slide", slide: null })).toThrow(/slide size/);
  });
});

describe("alt text, hyperlink, background", () => {
  it("sets alt text through the API on 1.10", async () => {
    const { agenda } = fx();
    const data = await ok("set_shape_alt_text", { ref: `${agenda}:4`, title: "Logo", description: "Company logo" });
    expect(data.method).toBe("api");
    expect(shapeOf(agenda, "4")).toMatchObject({ altTitle: "Logo", altDescr: "Company logo" });
  });

  it("falls back to the slide XML on 1.8, with a new slide id", async () => {
    fx().deck.api = 8;
    const { agenda } = fx();
    const data = await ok("set_shape_alt_text", { ref: `${agenda}:4`, description: "Company logo" });
    expect(data.method).toBe("ooxml");
    expect(data.slide_id).not.toBe(agenda);
    expect(data.ref).toBe(`${data.slide_id}:4`);
    expect(shapeOf(data.slide_id, "4").altDescr).toBe("Company logo");
    expect(fx().deck.ids()).not.toContain(agenda);
    expect(fx().deck.violations).toEqual([]);
  });

  it("is unsupported below 1.8", async () => {
    fx().deck.api = 7;
    expect(await fails("set_shape_alt_text", { ref: `${fx().agenda}:4`, title: "x" }, "unsupported")).toMatch(/1\.10/);
    expect(await fails("set_slide_background", { slide_id: fx().agenda, color: "#112233" }, "unsupported")).toMatch(/1\.10/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("links a shape on 1.10 only", async () => {
    const { agenda } = fx();
    const data = await ok("set_slide_hyperlink", { ref: `${agenda}:2`, url: "https://example.com/a", tooltip: "Open" });
    expect(data.summary).toMatch(/example\.com/);
    expect(shapeOf(agenda, "2").hyperlink).toEqual({ address: "https://example.com/a", screenTip: "Open" });
    fx().deck.api = 9;
    expect(await fails("set_slide_hyperlink", { ref: `${agenda}:2`, url: "mailto:a@b.co" }, "unsupported")).toMatch(/1\.10/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("sets the background through the API on 1.10", async () => {
    const data = await ok("set_slide_background", { slide_id: fx().agenda, color: "#112233" });
    expect(data.method).toBe("api");
    expect(fx().deck.slideById(fx().agenda).background).toBe("#112233");
  });

  it("sets the background through the XML on 1.8, putting <p:bg> first", async () => {
    fx().deck.api = 9;
    const data = await ok("set_slide_background", { slide_id: fx().agenda, color: "#abcdef" });
    expect(data).toMatchObject({ method: "ooxml", replaced_slide_id: fx().agenda });
    expect(fx().deck.slideById(data.slide_id).background).toBe("#ABCDEF");
    expect(fx().deck.slideById(data.slide_id).shapes).toHaveLength(3);
  });
});
