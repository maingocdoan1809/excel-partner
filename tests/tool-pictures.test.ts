// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { formatToolValue, paintTool, picturesIn, renderToolCard, splitPictures } from "../frontend/taskpane/ui/tool-view.js";
import type { StoredTool } from "../frontend/taskpane/ui/sessions.js";

const PNG = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 2, 8, 6, 0, 0, 0]).toString("base64");
const imageResult = { ok: true, data: { summary: "Rendered slide 256", slide_id: "256#", width: 960, image: { mimeType: "image/png", data: PNG } } };
const tool = (result?: string): StoredTool => ({ name: "export_slide_image", args: "", result, state: "ok", error: null });

describe("pictures on tool cards", () => {
  it("keeps a picture's base64 whole in the result text, even a long one", () => {
    const long = "A".repeat(50_000);
    const text = formatToolValue({ ok: true, data: { image: { mimeType: "image/png", data: long } } });
    expect(picturesIn(JSON.parse(text))).toEqual([{ mimeType: "image/png", data: long }]);
    expect(formatToolValue({ ok: true, data: { note: long } })).toContain("truncated");
  });

  it("shows the picture from the result first, and the rest of the result without its base64", () => {
    const card = renderToolCard(tool(formatToolValue(imageResult)));
    const body = card.querySelector(".tool-call-body")!;
    const image = body.querySelector<HTMLImageElement>(".tool-call-picture img")!;
    expect(body.firstElementChild?.classList.contains("tool-call-picture")).toBe(true);
    expect(image.src).toBe(`data:image/png;base64,${PNG}`);
    expect(body.textContent).toContain("Rendered slide 256");
    expect(body.textContent).not.toContain(PNG);
    image.click();
    expect(image.parentElement!.classList.contains("expanded")).toBe(true);
  });

  it("shows plain text for results without a picture, cut short or not JSON", () => {
    expect(splitPictures("not json")).toEqual({ pictures: [], text: "not json" });
    expect(splitPictures(formatToolValue(imageResult).slice(0, 60)).pictures).toEqual([]);
    const card = renderToolCard(tool(formatToolValue({ ok: true, data: { summary: "x" } })));
    expect(card.querySelector(".tool-call-picture")).toBeNull();
    paintTool(card, { ...tool(), state: "running" });
    expect(card.querySelector(".tool-call-picture")).toBeNull();
  });
});
