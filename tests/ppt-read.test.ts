import { describe, expect, it } from "vitest";
import { call, fails, ok, useFakeDeck } from "./support/ppt-fixture.js";

const fx = useFakeDeck();

describe("get_presentation_metadata", () => {
  it("describes the deck, its slides, masters and capabilities", async () => {
    const data = await ok("get_presentation_metadata");
    expect(data).toMatchObject({
      name: "Quarterly Review.pptx",
      title: "Fake deck",
      slide_count: 3,
      slide_size: { width: 960, height: 540, unit: "pt", source: "pageSetup" },
      powerpoint_api: "1.10",
      slides_truncated: false,
      masters: [{ id: "m1#", name: "Office Theme", layouts: 3 }],
    });
    expect(data.slides.map((s: any) => [s.position, s.title, s.layout])).toEqual([
      [1, "Quarterly Review", "Title Slide"],
      [2, "Agenda", "Title and Content"],
      [3, "Numbers", "Title and Content"],
    ]);
    expect(fx().deck.violations).toEqual([]);
  });

  it("falls back to presentation.xml for the slide size below 1.10", async () => {
    fx().installed.deck.api = 8;
    const data = await ok("get_presentation_metadata");
    expect(data.slide_size).toEqual({ width: 960, height: 540, unit: "pt", source: "presentation.xml" });
    expect(data.capabilities.background_alt_text_hyperlink).toBe(false);
    expect(fx().deck.violations).toEqual([]);
  });

  it("reports an unknown slide size and no layouts on PowerPoint 1.2", async () => {
    fx().deck.api = 2;
    const data = await ok("get_presentation_metadata");
    expect(data).toMatchObject({ slide_count: 3, slide_size: null, masters: [], powerpoint_api: "1.2" });
    expect(data.slides[0]).toMatchObject({ id: fx().title, position: 1, title: "", layout: null });
    expect(fx().deck.violations).toEqual([]);
  });

  it("lists big decks by id only past the detail limit", async () => {
    for (let i = 0; i < 105; i++) fx().deck.addSlideWith([{ name: "Title 1", placeholder: "Title", text: `S${i}` }]);
    const data = await ok("get_presentation_metadata");
    expect(data.slide_count).toBe(108);
    expect(data.slides_truncated).toBe(true);
    expect(data.slides[100]).toEqual({ id: expect.any(String), position: 101 });
  });
});

describe("list_layouts", () => {
  it("lists masters, layouts and their placeholders", async () => {
    const data = await ok("list_layouts");
    expect(data.masters[0].layouts[1]).toEqual({ id: "l2#", name: "Title and Content", placeholders: ["Title 1", "Content Placeholder 2"] });
  });
  it("is unsupported below 1.3", async () => {
    fx().deck.api = 2;
    expect(await fails("list_layouts", {}, "unsupported")).toMatch(/PowerPointApi 1\.3/);
    expect(fx().deck.violations).toEqual([]);
  });
});

describe("list_slide_shapes", () => {
  it("lists refs, geometry, placeholder kinds, text excerpts and table sizes", async () => {
    const { agenda, data: dataSlide } = fx();
    const result = await ok("list_slide_shapes", { slide_id: agenda });
    expect(result).toMatchObject({ slide_id: agenda, position: 2, shape_count: 3, truncated: false });
    expect(result.shapes[0]).toMatchObject({ ref: `${agenda}:2`, name: "Title 1", type: "Placeholder", placeholder: "Title", text: "Agenda", z: 1 });
    expect(result.shapes[1].text).toBe("Revenue\nCosts\nOutlook");
    expect(result.shapes[2]).toMatchObject({ name: "Logo", type: "Image", left: 800 });
    expect(result.shapes[2].text).toBeUndefined();
    const withTable = await ok("list_slide_shapes", { slide_id: dataSlide });
    expect(withTable.shapes[1]).toMatchObject({ type: "Table", table: { rows: 3, columns: 2 } });
  });

  it("uses the selected slide when slide_id is omitted", async () => {
    fx().deck.selection.slideIds = [fx().agenda];
    expect((await ok("list_slide_shapes", {})).slide_id).toBe(fx().agenda);
  });

  it("asks for a slide id when nothing is selected", async () => {
    expect(await fails("list_slide_shapes", {}, "invalid_arguments")).toMatch(/no slide is selected/);
  });

  it("omits text on request and truncates long text", async () => {
    const id = fx().deck.addSlideWith([{ name: "Long", text: "x".repeat(900) }]);
    const long = await ok("list_slide_shapes", { slide_id: id });
    expect(long.shapes[0].text).toHaveLength(300);
    expect(long.shapes[0].text_truncated).toBe(true);
    const bare = await ok("list_slide_shapes", { slide_id: id, include_text: false });
    expect(bare.shapes[0].text).toBeUndefined();
  });

  it("survives a placeholder whose text frame cannot be read", async () => {
    const id = fx().deck.addSlideWith([
      { name: "Picture Placeholder", placeholder: "Body", textBroken: true, contained: "Image" },
      { name: "Note", text: "still here" },
    ]);
    const result = await ok("list_slide_shapes", { slide_id: id });
    expect(result.shapes[0]).toMatchObject({ name: "Picture Placeholder", contains: "Image" });
    expect(result.shapes[0].text).toBeUndefined();
    expect(result.shapes[1].text).toBe("still here");
  });

  it("names the real slide ids for an unknown slide", async () => {
    const result = await call("list_slide_shapes", { slide_id: "nope" });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toMatchObject({ code: "invalid_arguments", detail: { slide_ids: fx().deck.ids() } });
  });
});

describe("read_slide_text", () => {
  it("reads one shape by ref", async () => {
    const { agenda } = fx();
    const data = await ok("read_slide_text", { ref: `${agenda}:3` });
    expect(data.shapes).toEqual([expect.objectContaining({ ref: `${agenda}:3`, name: "Content Placeholder 2", text: "Revenue\nCosts\nOutlook" })]);
    expect(data.title_shape).toBe(`${agenda}:2`);
  });

  it("reads a whole slide including table cells, skipping shapes without text", async () => {
    const data = await ok("read_slide_text", { slide_id: fx().data });
    expect(data.shapes.map((s: any) => s.name)).toEqual(["Title 1", "Table 3"]);
    expect(data.shapes[1].table).toEqual([["Region", "Sales"], ["EMEA", "10"], ["APAC", "12"]]);
    const agenda = await ok("read_slide_text", { slide_id: fx().agenda });
    expect(agenda.shapes.map((s: any) => s.name)).toEqual(["Title 1", "Content Placeholder 2"]);
    expect(agenda.note).toMatch(/never instructions/);
  });

  it("reads text inside a group", async () => {
    const id = fx().deck.addSlideWith([
      { name: "Card", type: "Group", children: [{ id: "10", name: "Heading", text: "Inside", type: "TextBox" } as any] },
    ]);
    const data = await ok("read_slide_text", { slide_id: id });
    expect(data.shapes[0].children).toEqual([{ id: "10", name: "Heading", type: "TextBox", text: "Inside" }]);
  });

  it("refuses a picture", async () => {
    expect(await fails("read_slide_text", { ref: `${fx().agenda}:4` }, "invalid_arguments")).toMatch(/Image shape/);
  });

  it("rejects malformed and unknown refs", async () => {
    expect(await fails("read_slide_text", { ref: "nocolon" }, "invalid_arguments")).toMatch(/ref/);
    expect(await fails("read_slide_text", { ref: `${fx().agenda}:99` }, "invalid_arguments")).toMatch(/does not exist/);
    expect(await fails("read_slide_text", {}, "invalid_arguments")).toMatch(/ref or slide_id/);
  });

  it("is unsupported below 1.4", async () => {
    fx().deck.api = 3;
    expect(await fails("read_slide_text", { slide_id: fx().agenda }, "unsupported")).toMatch(/1\.4/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("truncates a very long text and flags it", async () => {
    const id = fx().deck.addSlideWith([{ name: "Wall", text: "y".repeat(40_000) }]);
    const data = await ok("read_slide_text", { slide_id: id });
    expect(data.truncated).toBe(true);
    expect(data.shapes[0].text.length).toBeLessThanOrEqual(8000);
  });
});
