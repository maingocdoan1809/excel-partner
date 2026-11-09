import type { PptSelectionSnapshot } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { loadShapeInfos, readSlideTitle, tryEach } from "./model.js";
import { PPT_LAYOUTS, PPT_SELECTION, isSupported, loadSlideIds } from "./util.js";

const MAX_SLIDES = 50;
const MAX_SHAPES = 100;

interface SelectedSlideData {
  id: string;
  position: number;
  title?: string;
}

/** Selected slide through the common API (Selection 1.1): only the active slide, with no shapes. */
function readSlideRange(): Promise<Array<{ id: number; title: string; index: number }>> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve((result.value as Office.SlideRange | null)?.slides ?? []);
        else reject(new Error(result.error?.message ?? "could not read the selected slide"));
      });
    } catch (error) {
      reject(error);
    }
  });
}

/** Match the numeric id Office reports to the id string PowerPoint.Slide uses ("256" or "256#"). */
function matchSlideId(ids: string[], raw: number | string): string {
  const wanted = String(raw);
  return ids.find((id) => id === wanted) ?? ids.find((id) => id.replace(/#.*$/, "") === wanted.replace(/#.*$/, "")) ?? wanted;
}

async function selectedSlidesViaCommonApi(): Promise<SelectedSlideData[]> {
  const range = await readSlideRange();
  const first = range[0];
  if (!first) return [];
  let id = String(first.id);
  try {
    id = matchSlideId(await PowerPoint.run((context) => loadSlideIds(context)), first.id);
  } catch {
    /* keep the raw id */
  }
  return [{ id, position: first.index, title: first.title ?? "" }];
}

/** Id of the slide the person is looking at, for tools whose `slide_id` is optional. */
export async function selectedSlideId(context: PowerPoint.RequestContext): Promise<string> {
  if (isSupported(PPT_SELECTION)) {
    const selected = context.presentation.getSelectedSlides();
    selected.load("items/id");
    await context.sync();
    const first = selected.items[0];
    if (first) return first.id;
  } else {
    const fallback = await selectedSlidesViaCommonApi().catch(() => []);
    if (fallback[0]) return fallback[0].id;
  }
  throw new ToolFailure("invalid_arguments", "no slide is selected; pass slide_id (see get_presentation_metadata)");
}

/** The slides and shapes the person has selected, for the chat chip and the prompt. */
export async function getPptSelection(): Promise<PptSelectionSnapshot> {
  if (!isSupported(PPT_SELECTION)) {
    const slides = await selectedSlidesViaCommonApi().catch(() => []);
    return { host: "powerpoint", slides: slides.map((s) => ({ id: s.id, position: s.position, title: s.title ?? "" })), shapes: [], truncated: false };
  }
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const selectedSlides = presentation.getSelectedSlides();
    selectedSlides.load("items/id");
    const selectedShapes = presentation.getSelectedShapes();
    selectedShapes.load("items/id,items/name,items/type");
    presentation.slides.load("items/id");
    await context.sync();

    const allIds = presentation.slides.items.map((slide) => slide.id);
    const slideItems = selectedSlides.items;
    const shapeItems = selectedShapes.items;
    const keptSlides = slideItems.slice(0, MAX_SLIDES);
    const keptShapes = shapeItems.slice(0, MAX_SHAPES);

    // Shapes know their slide; the active slide is only the first selected slide.
    // getParentSlide() returns a fresh proxy on every call, so keep the one that was loaded.
    const parentSlides: PowerPoint.Slide[] = [];
    const parents = await tryEach(context, keptShapes, (shape) => {
      const parent = shape.getParentSlide();
      parent.load("id");
      parentSlides[keptShapes.indexOf(shape)] = parent;
    });

    const titles: string[] = [];
    for (const slide of keptSlides) {
      try {
        if (!isSupported(PPT_LAYOUTS)) throw new Error("no shapes");
        const infos = await loadShapeInfos(context, slide.shapes, { placeholders: true });
        titles.push(await readSlideTitle(context, infos));
      } catch {
        titles.push("");
      }
    }

    const activeId = keptSlides[0]?.id;
    return {
      host: "powerpoint" as const,
      slides: keptSlides.map((slide, i) => ({ id: slide.id, position: allIds.indexOf(slide.id) + 1, title: titles[i] ?? "" })),
      shapes: keptShapes.map((shape, i) => {
        const slideId = parents[i] && parentSlides[i] ? parentSlides[i]!.id : activeId ?? "";
        return { ref: `${slideId}:${shape.id}`, id: shape.id, name: shape.name, type: String(shape.type) };
      }),
      truncated: slideItems.length > MAX_SLIDES || shapeItems.length > MAX_SHAPES,
    };
  });
}

/** Small identity read for the composer: "Slide 3", "2 shapes on slide 3: Title 1, Content 2". */
export async function getPptSelectionLabel(): Promise<{ label: string; title: string }> {
  const selection = await getPptSelection();
  const slideFor = (ref: string) => selection.slides.find((slide) => ref.startsWith(`${slide.id}:`));
  if (selection.shapes.length > 0) {
    const names = selection.shapes.map((shape) => shape.name);
    const slide = slideFor(selection.shapes[0]!.ref) ?? selection.slides[0];
    const where = slide ? ` on slide ${slide.position}` : "";
    const list = names.slice(0, 4).join(", ") + (names.length > 4 ? ", ..." : "");
    const label = selection.shapes.length === 1 ? `${names[0]}${slide ? ` (slide ${slide.position})` : ""}` : `${selection.shapes.length} shapes${where}: ${list}`;
    return { label, title: `Selected ${selection.shapes.length === 1 ? "shape" : "shapes"}${where}: ${names.join(", ")}` };
  }
  if (selection.slides.length === 1) {
    const slide = selection.slides[0]!;
    return { label: `Slide ${slide.position}`, title: slide.title ? `Selected slide ${slide.position}: ${slide.title}` : `Selected slide ${slide.position}` };
  }
  if (selection.slides.length > 1) {
    const positions = selection.slides.map((slide) => slide.position);
    return { label: `${positions.length} slides`, title: `Selected slides: ${positions.join(", ")}` };
  }
  return { label: "No selection", title: "Nothing is selected in the presentation" };
}

/** File name of the open presentation (basename of Office.context.document.url, decoded); null when it is not saved. */
export async function readPptPresentationName(): Promise<string | null> {
  try {
    const url = Office.context.document.url;
    if (!url) return null;
    const path = url.replace(/[?#].*$/, "");
    const base = path.split(/[\\/]/).filter(Boolean).pop() ?? "";
    let name = base;
    try {
      name = decodeURIComponent(base);
    } catch {
      /* keep the raw segment */
    }
    return name || null;
  } catch {
    return null;
  }
}
