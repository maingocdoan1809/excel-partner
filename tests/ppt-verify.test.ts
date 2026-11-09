import { describe, expect, it } from "vitest";
import { guardResultSize } from "../addin/taskpane/executor/powerpoint/index.js";
import type { ShapeInfo } from "../addin/taskpane/executor/powerpoint/model.js";
import { analyzeSlide, contrastRatio, estimateTextHeight } from "../addin/taskpane/executor/powerpoint/verify.js";
import { call, fails, ok, useFakeDeck } from "./support/ppt-fixture.js";

const fx = useFakeDeck();
const SIZE = { width: 960, height: 540, source: "pageSetup" as const };

function info(init: Partial<ShapeInfo> & { name: string }): ShapeInfo {
  return { id: init.id ?? init.name, type: "TextBox", z: 1, left: 100, top: 100, width: 300, height: 60, text: "Hello", fontSize: 18, shape: {} as never, ...init };
}
const codes = (shapes: ShapeInfo[], size = SIZE as typeof SIZE | null) => analyzeSlide("1", shapes, size).map((i) => `${i.code}:${i.shape_ref}`);

describe("analyzeSlide (geometry, text, colour)", () => {
  it("flags shapes outside the slide", () => {
    const issues = analyzeSlide("7", [
      info({ name: "Gone", left: 1000, top: 10 }),
      info({ name: "Hanging", left: 800, top: 100, width: 300 }),
      info({ name: "Fine", left: 10, top: 200 }),
      info({ name: "Edge", left: 660, top: 300, width: 300.5 }),
    ], SIZE);
    expect(issues.map((i) => [i.shape_ref, i.code, i.severity])).toEqual([
      ["7:Gone", "off_slide", "error"],
      ["7:Hanging", "off_slide", "warning"],
    ]);
    expect(issues[1]!.message).toMatch(/sticks out/);
    expect(issues[0]).toMatchObject({ slide_id: "7", message: expect.stringMatching(/completely outside/) });
  });

  it("skips the off-slide rule when the slide size is unknown", () => {
    expect(codes([info({ name: "Gone", left: 5000 })], null)).toEqual([]);
  });

  it("flags overlapping text shapes but not containers or full-bleed backgrounds", () => {
    expect(codes([info({ name: "A", left: 100, top: 100 }), info({ name: "B", left: 200, top: 120 })])).toEqual(["overlap:1:A"]);
    expect(codes([info({ name: "Card", left: 50, top: 50, width: 500, height: 300, text: "Card" }), info({ name: "Inner", left: 100, top: 100 })])).toEqual([]);
    expect(codes([info({ name: "Bg", left: 0, top: 0, width: 960, height: 540, text: "bg" }), info({ name: "Title", left: 100, top: 100 })])).toEqual([]);
    expect(codes([info({ name: "A", left: 100, top: 100 }), info({ name: "B", left: 100, top: 400 })])).toEqual([]);
    expect(codes([info({ name: "A" }), info({ name: "Pic", type: "Image", text: undefined })])).toEqual([]);
    expect(codes([info({ name: "A", left: 100, top: 100 }), info({ name: "Empty", left: 150, top: 110, text: "" })])).toEqual([]);
  });

  it("flags empty placeholders but not footers or filled ones", () => {
    const placeholder = (name: string, kind: string, text: string, contained: string | null = null) => info({ name, type: "Placeholder", placeholder: kind, text, contained });
    expect(codes([placeholder("Body", "Body", ""), placeholder("Footer", "Footer", ""), placeholder("Number", "SlideNumber", ""), placeholder("Pic", "Body", "", "Image"), placeholder("Title", "Title", "Hi")])).toEqual(["empty_placeholder:1:Body"]);
  });

  it("flags tiny fonts", () => {
    expect(codes([info({ name: "Small", fontSize: 8 }), info({ name: "Ten", fontSize: 10, top: 300 }), info({ name: "Mixed", fontSize: null, top: 400 })])).toEqual(["tiny_font:1:Small"]);
  });

  it("flags text that probably overflows, with 'maybe' wording, unless the shape resizes", () => {
    const long = "word ".repeat(60);
    const issues = analyzeSlide("1", [info({ name: "Tight", text: long, height: 40 }), info({ name: "Shrinks", text: long, height: 40, autoSize: "AutoSizeTextToFitShape", top: 200 }), info({ name: "Grows", text: long, height: 40, autoSize: "AutoSizeShapeToFitText", top: 300 }), info({ name: "Roomy", text: "short", height: 60, top: 400 })], SIZE);
    expect(issues.map((i) => [i.shape_ref, i.code])).toEqual([["1:Tight", "text_overflow_maybe"]]);
    expect(issues[0]!.message).toMatch(/may overflow/);
    expect(issues[0]!.severity).toBe("warning");
  });

  it("estimates line wrapping, CJK width and margins", () => {
    expect(estimateTextHeight("abc", 10, 300)).toBeCloseTo(12 + 7.2);
    expect(estimateTextHeight("a\nb\nc", 10, 300)).toBeCloseTo(36 + 7.2);
    expect(estimateTextHeight("あ".repeat(10), 10, 64.4)).toBeGreaterThan(estimateTextHeight("a".repeat(10), 10, 64.4));
    expect(estimateTextHeight("x".repeat(100), 10, 100, undefined, false)).toBeCloseTo(12 + 7.2);
  });

  it("flags low contrast between font and solid fill when both are known", () => {
    expect(codes([info({ name: "Pale", fillType: "Solid", fillColor: "#FFFF00", fontColor: "#FFFFFF" })])).toEqual(["low_contrast:1:Pale"]);
    expect(codes([info({ name: "Dark", fillType: "Solid", fillColor: "#000000", fontColor: "#FFFFFF" })])).toEqual([]);
    expect(codes([info({ name: "NoFill", fillType: "NoFill", fillColor: "#FFFF00", fontColor: "#FFFFFF" })])).toEqual([]);
    expect(codes([info({ name: "Unknown", fillType: "Solid", fillColor: "#FFFF00", fontColor: null })])).toEqual([]);
    // large text needs only 3:1
    expect(codes([info({ name: "Big", fontSize: 28, fillType: "Solid", fillColor: "#FFFFFF", fontColor: "#888888" })])).toEqual([]);
    expect(codes([info({ name: "Small", fontSize: 12, fillType: "Solid", fillColor: "#FFFFFF", fontColor: "#999999" })])).toEqual(["low_contrast:1:Small"]);
    expect(contrastRatio("#000000", "#FFFFFF")).toBeCloseTo(21, 0);
    expect(contrastRatio("red", "#FFFFFF")).toBeNull();
  });
});

describe("verify_slides", () => {
  it("checks every slide through the host objects and sorts errors first", async () => {
    const { deck } = fx();
    const bad = deck.addSlideWith([
      { name: "Off", left: 2000, top: 0, text: "far away" },
      { name: "Tiny", text: "tiny", fontSize: 6, top: 200, left: 20, width: 200 },
      { name: "Body", placeholder: "Body", text: "", top: 300 },
    ]);
    const data = await ok("verify_slides", {});
    expect(data.checked).toEqual({ slides: 4, shapes: 10 });
    const mine = data.issues.filter((i: any) => i.slide_id === bad);
    expect(mine.map((i: any) => [i.code, i.severity])).toEqual([["off_slide", "error"], ["tiny_font", "warning"], ["empty_placeholder", "warning"]]);
    expect(mine[0].shape_ref).toBe(`${bad}:2`);
    expect(data.slide_size).toEqual({ width: 960, height: 540, unit: "pt" });
    expect(data.summary).toMatch(/issue/);
  });

  it("reports a clean deck and limits to the given slides", async () => {
    const clean = await ok("verify_slides", { slide_ids: [fx().agenda] });
    expect(clean.issues).toEqual([]);
    expect(clean.checked.slides).toBe(1);
    expect(clean.summary).toMatch(/no issues/);
  });

  it("rejects unknown slides and is unsupported below 1.4", async () => {
    await fails("verify_slides", { slide_ids: ["ghost"] }, "invalid_arguments");
    fx().deck.api = 3;
    expect(await fails("verify_slides", {}, "unsupported")).toMatch(/1\.4/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("uses the exported presentation.xml for the slide size on 1.8 and notes an unknown size below", async () => {
    fx().deck.api = 8;
    fx().deck.addSlideWith([{ name: "Off", left: 3000 }]);
    const data = await ok("verify_slides", {});
    expect(data.issues.some((i: any) => i.code === "off_slide")).toBe(true);
    fx().deck.api = 4;
    const old = await ok("verify_slides", {});
    expect(old.slide_size).toBeNull();
    expect(old.issues.some((i: any) => i.code === "off_slide")).toBe(false);
    expect(old.notes.join(" ")).toMatch(/slide size is unknown/);
    expect(fx().deck.violations).toEqual([]);
  });
});

describe("verify_slide_visual / export_slide_image", () => {
  it("returns an image block payload plus the slide's issues", async () => {
    const id = fx().deck.addSlideWith([{ name: "Off", left: 2000, text: "x" }]);
    const data = await ok("verify_slide_visual", { slide_id: id, width: 640 });
    expect(data.image).toEqual({ mimeType: "image/png", data: fx().deck.imageData });
    expect(data.issues.map((i: any) => i.code)).toEqual(["off_slide"]);
    expect(data.width).toBe(640);
    expect(fx().deck.imageRequests).toEqual([640]);
  });

  it("defaults the width and exports without issues", async () => {
    const data = await ok("export_slide_image", { slide_id: fx().agenda });
    expect(data).toMatchObject({ width: 960, image: { mimeType: "image/png" } });
    expect(data.issues).toBeUndefined();
    expect(fx().deck.imageRequests).toEqual([960]);
  });

  it("gates on 1.8 without touching the host and refuses oversized images", async () => {
    fx().deck.imageData = "A".repeat(1_600_000);
    expect(await fails("export_slide_image", { slide_id: fx().agenda }, "size_limit_exceeded")).toMatch(/smaller width/);
    fx().deck.api = 7;
    expect(await fails("verify_slide_visual", { slide_id: fx().agenda }, "unsupported")).toMatch(/1\.8/);
    expect(await fails("export_slide_image", { slide_id: fx().agenda }, "unsupported")).toMatch(/1\.8/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("does not count the image against the 30,000 character cap", async () => {
    fx().deck.imageData = "B".repeat(900_000);
    const result = await call("export_slide_image", { slide_id: fx().agenda });
    expect(result.ok).toBe(true);
  });
});

describe("result size cap", () => {
  it("fails a read that is too big and keeps the summary of a write that already happened", () => {
    const big = { rows: "x".repeat(40_000) };
    expect(() => guardResultSize(big)).toThrow(/limit 30000/);
    expect(guardResultSize({ summary: "Edited", rows: "x".repeat(40_000) })).toEqual({ summary: "Edited", note: expect.stringMatching(/too large/) });
    expect(guardResultSize({ summary: "ok" })).toEqual({ summary: "ok" });
    expect(guardResultSize({ image: { data: "x".repeat(900_000) }, summary: "ok" })).toMatchObject({ summary: "ok" });
  });

  it("shortens, then drops, shape text instead of failing on a crowded slide", async () => {
    const shapes = Array.from({ length: 150 }, (_, i) => ({ name: `Shape ${i}`, text: "t".repeat(400), left: i, top: i }));
    const id = fx().deck.addSlideWith(shapes);
    const result = await call("list_slide_shapes", { slide_id: id });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const data = result.data as any;
      expect(JSON.stringify(data).length).toBeLessThan(30_000);
      expect(data.text_shortened).toMatch(/text/);
      expect(data.shapes).toHaveLength(150);
    }
  });

  it("caps the shape list and the whole-slide text read", async () => {
    const shapes = Array.from({ length: 200 }, (_, i) => ({ name: `S${i}`, text: "word ".repeat(100), left: i, top: i }));
    const id = fx().deck.addSlideWith(shapes);
    const list = await ok("list_slide_shapes", { slide_id: id, include_text: false });
    expect(list).toMatchObject({ shape_count: 200, truncated: true });
    expect(list.shapes).toHaveLength(150);
    const text = await ok("read_slide_text", { slide_id: id });
    expect(text.truncated).toBe(true);
    expect(JSON.stringify(text).length).toBeLessThan(30_000);
  });

  it("clips very long titles in the metadata", async () => {
    fx().deck.addSlideWith([{ name: "Title 1", placeholder: "Title", text: "T".repeat(5000) }]);
    const data = await ok("get_presentation_metadata");
    expect(data.slides.at(-1).title.length).toBeLessThanOrEqual(153);
  });
});
