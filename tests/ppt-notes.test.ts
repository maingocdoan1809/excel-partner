import { describe, expect, it } from "vitest";
import { resolvePart } from "../frontend/taskpane/executor/powerpoint/notes.js";
import { fails, ok, useFakeDeck } from "./support/ppt-fixture.js";

const fx = useFakeDeck();
const notesOf = (slideId: string) => fx().deck.slideById(slideId).notes;

describe("read_slide_notes", () => {
  it("reads the notes of every slide, and says which have no notes page", async () => {
    fx().deck.slideById(fx().agenda).notes = "Welcome everyone.\n\nThen walk through the agenda.";
    const data = await ok("read_slide_notes");
    expect(data.slides).toEqual([
      { slide_id: fx().title, position: 1, notes: "", has_notes_page: false },
      { slide_id: fx().agenda, position: 2, notes: "Welcome everyone.\n\nThen walk through the agenda.", has_notes_page: true },
      { slide_id: fx().data, position: 3, notes: "", has_notes_page: false },
    ]);
  });

  it("reads only the slides asked for and names an unknown one", async () => {
    fx().deck.slideById(fx().data).notes = "Numbers are preliminary.";
    expect((await ok("read_slide_notes", { slide_ids: [fx().data] })).slides).toEqual([{ slide_id: fx().data, position: 3, notes: "Numbers are preliminary.", has_notes_page: true }]);
    expect(await fails("read_slide_notes", { slide_ids: ["999#"] }, "invalid_arguments")).toMatch(/does not exist/);
  });

  it("reads a long deck in order, a few slides per export", async () => {
    const ids = Array.from({ length: 12 }, (_, i) => {
      const id = fx().deck.addSlideWith([{ name: "Title 1", placeholder: "Title", text: `S${i}`, top: 20, height: 80 }]);
      fx().deck.slideById(id).notes = `note ${i}`;
      return id;
    });
    const data = await ok("read_slide_notes", { slide_ids: ids });
    expect(data.slides.map((x: any) => x.notes)).toEqual(ids.map((_, i) => `note ${i}`));
    expect(data.slides.map((x: any) => x.slide_id)).toEqual(ids);
  });

  it("needs PowerPointApi 1.8", async () => {
    fx().deck.api = 7;
    expect(await fails("read_slide_notes", {}, "unsupported")).toMatch(/1\.8/);
  });
});

describe("set_slide_notes", () => {
  it("replaces existing notes; the slide comes back with a new id in the same place with its shapes", async () => {
    fx().deck.slideById(fx().agenda).notes = "Old notes.";
    const shapesBefore = fx().deck.slideById(fx().agenda).shapes.map((s) => s.name);
    const data = await ok("set_slide_notes", { slide_id: fx().agenda, text: "Line one.\nLine two." });
    expect(data).toMatchObject({ position: 2, replaced_slide_id: fx().agenda });
    expect(data.slide_id).not.toBe(fx().agenda);
    expect(fx().deck.ids()).toHaveLength(3);
    expect(fx().deck.ids()[1]).toBe(data.slide_id);
    expect(notesOf(data.slide_id)).toBe("Line one.\nLine two.");
    expect(fx().deck.slideById(data.slide_id).shapes.map((s) => s.name)).toEqual(shapesBefore);
    expect((await ok("read_slide_notes", { slide_ids: [data.slide_id] })).slides[0].notes).toBe("Line one.\nLine two.");
  });

  it("gives a slide its first notes page from the notes master", async () => {
    const data = await ok("set_slide_notes", { slide_id: fx().title, text: "Say hello." });
    expect(data).toMatchObject({ created_notes_page: true, position: 1 });
    expect(notesOf(data.slide_id)).toBe("Say hello.");
  });

  it("clears notes with empty text, and says why a first note is impossible without a notes master", async () => {
    fx().deck.slideById(fx().agenda).notes = "Remove me.";
    const cleared = await ok("set_slide_notes", { slide_id: fx().agenda, text: "" });
    expect(notesOf(cleared.slide_id)).toBe("");
    fx().deck.notesMaster = false;
    expect(await fails("set_slide_notes", { slide_id: fx().title, text: "x" }, "unsupported")).toMatch(/notes master/);
    expect(fx().deck.ids()[0]).toBe(fx().title);
  });

  it("escapes text that looks like markup", async () => {
    const data = await ok("set_slide_notes", { slide_id: fx().data, text: "Q3 < Q4 & <b>not bold</b>" });
    expect(notesOf(data.slide_id)).toBe("Q3 < Q4 & <b>not bold</b>");
  });
});

describe("notes package paths", () => {
  it("resolves relationship targets from the slide's folder", () => {
    expect(resolvePart("ppt/slides/slide1.xml", "../notesSlides/notesSlide1.xml")).toBe("ppt/notesSlides/notesSlide1.xml");
    expect(resolvePart("ppt/slides/slide1.xml", "/ppt/notesSlides/notesSlide2.xml")).toBe("ppt/notesSlides/notesSlide2.xml");
    expect(resolvePart("ppt/slides/slide1.xml", "./x.xml")).toBe("ppt/slides/x.xml");
  });
});

describe("update_slide_shape", () => {
  const logo = () => fx().deck.slideById(fx().agenda).shapes.find((s) => s.name === "Logo")!;

  it("resizes a picture in proportion and moves it", async () => {
    const data = await ok("update_slide_shape", { ref: `${fx().agenda}:${logo().id}`, width: 200, left: 700, top: 10 });
    expect(data).toMatchObject({ before: { left: 800, top: 20, width: 100, height: 50 }, left: 700, top: 10, width: 200, height: 100 });
    expect(logo()).toMatchObject({ left: 700, top: 10, width: 200, height: 100 });
  });

  it("takes an exact size when asked, and renames", async () => {
    await ok("update_slide_shape", { ref: `${fx().agenda}:${logo().id}`, height: 80, keep_proportions: false, name: "Company logo" });
    expect(fx().deck.slideById(fx().agenda).shapes.find((s) => s.name === "Company logo")).toMatchObject({ width: 100, height: 80 });
  });

  it("refuses an unknown shape and an empty change", async () => {
    expect(await fails("update_slide_shape", { ref: `${fx().agenda}:999`, left: 1 }, "invalid_arguments")).toMatch(/does not exist/);
    expect(await fails("update_slide_shape", { ref: `${fx().agenda}:${logo().id}` }, "invalid_arguments")).toMatch(/give left/);
  });
});
