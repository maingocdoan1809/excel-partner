import type {
  AddSlideInput,
  DuplicateSlideInput,
  ModifyPresentationStructureInput,
  MoveSlideInput,
  SetSlideBackgroundInput,
} from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { loadShapeInfos, publicShape } from "./model.js";
import { NS, makeElementHelper, roundTripSlide } from "./ooxml.js";
import {
  PPT_EXPORT,
  PPT_LAYOUTS,
  PPT_MODERN,
  PPT_MOVE,
  isSupported,
  loadSlideIds,
  requireSlide,
  requireSupport,
  slideNotFound,
} from "./util.js";

/** Core of add_slide, shared with modify_presentation_structure. Runs inside an existing PowerPoint.run. */
async function addSlideIn(context: PowerPoint.RequestContext, input: { layout_id: string; master_id?: string; position?: number }): Promise<Record<string, unknown>> {
  requireSupport("add_slide", PPT_LAYOUTS);
  const before = await loadSlideIds(context);

  const masters = context.presentation.slideMasters;
  masters.load("items/id,items/name");
  await context.sync();
  for (const master of masters.items) master.layouts.load("items/id,items/name");
  await context.sync();
  let masterId = input.master_id;
  let layoutName = "";
  for (const master of masters.items) {
    if (input.master_id !== undefined && master.id !== input.master_id) continue;
    const layout = master.layouts.items.find((item) => item.id === input.layout_id);
    if (layout) {
      masterId = master.id;
      layoutName = layout.name;
      break;
    }
  }
  if (!masterId || !layoutName) {
    throw new ToolFailure("invalid_arguments", `layout '${input.layout_id}' was not found${input.master_id ? ` in master '${input.master_id}'` : ""}; list the layouts with list_layouts`, {
      layouts: masters.items.flatMap((master) => master.layouts.items.map((layout) => ({ master_id: master.id, id: layout.id, name: layout.name }))).slice(0, 60),
    });
  }

  context.presentation.slides.add({ slideMasterId: masterId, layoutId: input.layout_id });
  await context.sync();
  let ids = await loadSlideIds(context);
  const known = new Set(before);
  const added = ids.filter((id) => !known.has(id));
  if (added.length !== 1) throw new ToolFailure("office_error", `PowerPoint did not report exactly one new slide (found ${added.length}); check the deck with get_presentation_metadata`);
  const newId = added[0]!;

  const warnings: string[] = [];
  const wanted = Math.min(input.position ?? ids.length, ids.length);
  if (wanted !== ids.indexOf(newId) + 1) {
    if (isSupported(PPT_MOVE)) {
      context.presentation.slides.getItem(newId).moveTo(wanted - 1);
      await context.sync();
      ids = await loadSlideIds(context);
    } else {
      warnings.push(`position ${wanted} could not be honoured: moving slides needs PowerPointApi 1.8, so the slide stays at position ${ids.indexOf(newId) + 1}`);
    }
  }
  const position = ids.indexOf(newId) + 1;

  let placeholders: Array<Record<string, unknown>> = [];
  try {
    const infos = await loadShapeInfos(context, context.presentation.slides.getItem(newId).shapes, { geometry: true, placeholders: true });
    placeholders = infos.map((info) => publicShape(newId, info));
  } catch {
    /* the slide exists; its shapes can be listed later */
  }
  return {
    summary: `Added slide ${position} (id ${newId}) with layout "${layoutName}"`,
    slide_id: newId,
    position,
    slide_count: ids.length,
    layout: layoutName,
    shapes: placeholders,
    ...(warnings.length ? { warnings } : {}),
  };
}

export async function addSlide(input: AddSlideInput): Promise<unknown> {
  return PowerPoint.run((context) => addSlideIn(context, input));
}

export async function duplicateSlide(input: DuplicateSlideInput): Promise<unknown> {
  requireSupport("duplicate_slide", PPT_EXPORT, "PowerPoint has no duplicate API; it is done by exporting the slide and inserting it again, which needs 1.8.");
  return PowerPoint.run(async (context) => {
    const result = await roundTripSlide(context, { slideId: input.slide_id, mode: "copy", copyPosition: input.position });
    return {
      summary: `Duplicated slide ${input.slide_id} as slide ${result.position} (id ${result.newSlideId})`,
      source_slide_id: input.slide_id,
      slide_id: result.newSlideId,
      position: result.position,
      slide_count: result.slideCount,
      shape_count: result.shapeCount,
      note: "The copy has a new slide id; its shapes keep their ids, so refs are '<new slide id>:<shape id>'. Notes may not be carried over.",
      ...(result.warnings.length ? { warnings: result.warnings } : {}),
    };
  });
}

async function moveSlideIn(context: PowerPoint.RequestContext, slideId: string, position: number): Promise<Record<string, unknown>> {
  requireSupport("move_slide", PPT_MOVE, "Reordering slides is not available on this version of PowerPoint; ask the person to drag the slide in the thumbnail pane.");
  const ids = await loadSlideIds(context);
  const from = ids.indexOf(slideId);
  if (from === -1) throw slideNotFound(slideId, ids);
  const target = Math.min(position, ids.length);
  if (target !== from + 1) {
    context.presentation.slides.getItem(slideId).moveTo(target - 1);
    await context.sync();
  }
  const after = await loadSlideIds(context);
  const now = after.indexOf(slideId) + 1;
  return {
    summary: target === from + 1 ? `Slide ${slideId} is already at position ${now}` : `Moved slide ${slideId} from position ${from + 1} to ${now}`,
    slide_id: slideId,
    from: from + 1,
    position: now,
    slide_count: after.length,
    ...(position > ids.length ? { note: `position ${position} is past the last slide, so the slide went to the end` } : {}),
  };
}

export async function moveSlide(input: MoveSlideInput): Promise<unknown> {
  return PowerPoint.run((context) => moveSlideIn(context, input.slide_id, input.position));
}

export async function modifyPresentationStructure(input: ModifyPresentationStructureInput): Promise<unknown> {
  return PowerPoint.run(async (context) => {
    if (input.operation === "add") return addSlideIn(context, { layout_id: input.layout_id!, position: input.position });
    if (input.operation === "move") return moveSlideIn(context, input.slide_id!, input.position!);

    const wanted = [...new Set([...(input.slide_ids ?? []), ...(input.slide_id ? [input.slide_id] : [])])];
    const ids = await loadSlideIds(context);
    const missing = wanted.filter((id) => !ids.includes(id));
    if (missing.length) throw slideNotFound(missing[0]!, ids);
    if (!input.confirm_delete) {
      throw new ToolFailure(
        "confirmation_required",
        `Deleting ${wanted.length} slide(s) (${wanted.join(", ")}) needs the person's approval: ask them, then repeat with confirm_delete: true. Editing the slide in place is usually better than deleting and rebuilding it.`,
      );
    }
    if (wanted.length >= ids.length) {
      throw new ToolFailure("invalid_arguments", "refusing to delete every slide; a presentation needs at least one. Delete the others or edit this one in place.");
    }
    for (const id of wanted) context.presentation.slides.getItem(id).delete();
    await context.sync();
    const after = await loadSlideIds(context);
    return {
      summary: `Deleted ${wanted.length} slide(s): ${wanted.join(", ")}`,
      deleted: wanted,
      slide_count: after.length,
      remaining_slide_ids: after.slice(0, 200),
    };
  });
}

export async function setSlideBackground(input: SetSlideBackgroundInput): Promise<unknown> {
  return PowerPoint.run(async (context) => {
    const slide = await requireSlide(context, input.slide_id);
    if (isSupported(PPT_MODERN)) {
      slide.background.fill.setSolidFill({ color: input.color });
      await context.sync();
      return { summary: `Set the background of slide ${input.slide_id} to ${input.color}`, slide_id: input.slide_id, color: input.color, method: "api" };
    }
    requireSupport("set_slide_background", PPT_EXPORT, "Slide backgrounds need PowerPointApi 1.10, or 1.8 for the slower file-based fallback.");
    const result = await roundTripSlide(context, {
      slideId: input.slide_id,
      mode: "replace",
      mutate: (doc) => {
        const el = makeElementHelper(doc);
        const cSld = doc.getElementsByTagNameNS(NS.p, "cSld")[0];
        if (!cSld) throw new ToolFailure("office_error", "the slide XML has no <p:cSld>");
        for (const old of Array.from(cSld.children).filter((child) => child.localName === "bg")) cSld.removeChild(old);
        const bg = el("p:bg", null, el("p:bgPr", null, el("a:solidFill", null, el("a:srgbClr", { val: input.color.slice(1).toUpperCase() })), el("a:effectLst")));
        cSld.insertBefore(bg, cSld.firstChild);
      },
    });
    return {
      summary: `Set the background of slide ${result.position} to ${input.color} (slide rebuilt, new id ${result.newSlideId})`,
      slide_id: result.newSlideId,
      replaced_slide_id: input.slide_id,
      position: result.position,
      color: input.color,
      method: "ooxml",
      note: "The slide was rebuilt through its .pptx, so it has a new slide id.",
      ...(result.warnings.length ? { warnings: result.warnings } : {}),
    };
  });
}
