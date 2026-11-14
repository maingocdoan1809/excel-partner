import { describe, expect, it } from "vitest";
import { getPptSelection, getPptSelectionLabel, readPptPresentationName } from "../frontend/taskpane/executor/powerpoint/index.js";
import { AnySelectionSnapshotSchema, PptSelectionSnapshotSchema } from "../protocol/src/index.js";
import { useFakeDeck } from "./support/ppt-fixture.js";

const fx = useFakeDeck();

describe("getPptSelection with PowerPointApi 1.5", () => {
  it("reports selected slides and shapes with refs, validated against the schema", async () => {
    const { deck, agenda } = fx();
    deck.selection = { slideIds: [agenda], shapes: [[agenda, "2"], [agenda, "3"]] };
    const snapshot = await getPptSelection();
    expect(PptSelectionSnapshotSchema.parse(snapshot)).toEqual(snapshot);
    expect(AnySelectionSnapshotSchema.parse(snapshot)).toEqual(snapshot);
    expect(snapshot).toEqual({
      host: "powerpoint",
      slides: [{ id: agenda, position: 2, title: "Agenda" }],
      shapes: [
        { ref: `${agenda}:2`, id: "2", name: "Title 1", type: "Placeholder" },
        { ref: `${agenda}:3`, id: "3", name: "Content Placeholder 2", type: "Placeholder" },
      ],
      truncated: false,
    });
    expect(deck.violations).toEqual([]);
  });

  it("reports an empty selection", async () => {
    const snapshot = await getPptSelection();
    expect(snapshot).toEqual({ host: "powerpoint", slides: [], shapes: [], truncated: false });
    expect(PptSelectionSnapshotSchema.safeParse(snapshot).success).toBe(true);
  });

  it("takes each shape's slide from the shape itself", async () => {
    const { deck, agenda, data } = fx();
    deck.selection = { slideIds: [agenda, data], shapes: [[data, "3"]] };
    const snapshot = await getPptSelection();
    expect(snapshot.slides.map((s) => s.position)).toEqual([2, 3]);
    expect(snapshot.shapes[0]).toMatchObject({ ref: `${data}:3`, type: "Table" });
  });

  it("caps at 50 slides and 100 shapes and says so", async () => {
    const { deck } = fx();
    const ids: string[] = [];
    for (let i = 0; i < 60; i++) ids.push(deck.addSlideWith(Array.from({ length: 2 }, (_, k) => ({ name: `S${k}`, text: "x" }))));
    deck.selection = { slideIds: ids, shapes: ids.flatMap((id) => [[id, "2"], [id, "3"]] as Array<[string, string]>) };
    const snapshot = await getPptSelection();
    expect(snapshot.slides).toHaveLength(50);
    expect(snapshot.shapes).toHaveLength(100);
    expect(snapshot.truncated).toBe(true);
    expect(PptSelectionSnapshotSchema.safeParse(snapshot).success).toBe(true);
  });
});

describe("getPptSelection without 1.5", () => {
  it("falls back to the active slide through the common API, without shapes", async () => {
    const { deck, agenda } = fx();
    deck.api = 4;
    deck.selection = { slideIds: [agenda], shapes: [[agenda, "2"]] };
    const snapshot = await getPptSelection();
    expect(snapshot).toEqual({ host: "powerpoint", slides: [{ id: agenda, position: 2, title: "" }], shapes: [], truncated: false });
    expect(PptSelectionSnapshotSchema.safeParse(snapshot).success).toBe(true);
    expect(deck.violations).toEqual([]);
  });

  it("matches the numeric id Office reports to the slide id string", async () => {
    const { deck, data } = fx();
    deck.api = 3;
    deck.selection = { slideIds: [data], shapes: [] };
    expect((await getPptSelection()).slides[0]!.id).toBe(data);
  });

  it("returns no slides when nothing is selected", async () => {
    fx().deck.api = 3;
    expect((await getPptSelection()).slides).toEqual([]);
  });
});

describe("getPptSelectionLabel", () => {
  it("labels slides, one shape, several shapes and nothing", async () => {
    const { deck, title, agenda } = fx();
    expect(await getPptSelectionLabel()).toEqual({ label: "No selection", title: expect.any(String) });
    deck.selection = { slideIds: [agenda], shapes: [] };
    expect(await getPptSelectionLabel()).toEqual({ label: "Slide 2", title: "Selected slide 2: Agenda" });
    deck.selection = { slideIds: [title, agenda], shapes: [] };
    expect((await getPptSelectionLabel()).label).toBe("2 slides");
    deck.selection = { slideIds: [agenda], shapes: [[agenda, "2"]] };
    expect((await getPptSelectionLabel()).label).toBe("Title 1 (slide 2)");
    deck.selection = { slideIds: [agenda], shapes: [[agenda, "2"], [agenda, "3"]] };
    expect(await getPptSelectionLabel()).toEqual({ label: "2 shapes on slide 2: Title 1, Content Placeholder 2", title: "Selected shapes on slide 2: Title 1, Content Placeholder 2" });
  });
});

describe("readPptPresentationName", () => {
  it("decodes the basename of the document url", async () => {
    expect(await readPptPresentationName()).toBe("Quarterly Review.pptx");
    fx().deck.url = "C:\\Users\\me\\Documents\\Plan 2025.pptx";
    expect(await readPptPresentationName()).toBe("Plan 2025.pptx");
    fx().deck.url = "https://host/x/Deck.pptx?web=1#frag";
    expect(await readPptPresentationName()).toBe("Deck.pptx");
  });
  it("is null for an unsaved presentation", async () => {
    fx().deck.url = "";
    expect(await readPptPresentationName()).toBeNull();
  });
});
