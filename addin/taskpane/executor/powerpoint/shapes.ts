import type {
  AlignShapesInput,
  GroupShapesInput,
  InsertSlideElementInput,
  RemoveSlideElementInput,
  SetShapeAltTextInput,
  SetShapeZOrderInput,
  SetSlideHyperlinkInput,
  UngroupShapesInput,
} from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { NS, getSlideSize, roundTripSlide } from "./ooxml.js";
import { addTable } from "./tables.js";
import {
  PPT_EXPORT,
  PPT_GEOMETRY,
  PPT_MODERN,
  PPT_ZORDER,
  isSupported,
  makeRef,
  parseShapeRef,
  requireShape,
  requireSlide,
  requireSupport,
  round,
} from "./util.js";

type Element = InsertSlideElementInput["element"];
type TextStyle = Extract<Element, { type: "textbox" }>["style"];

function applyTextStyle(shape: PowerPoint.Shape, style: TextStyle): void {
  if (!style) return;
  const range = shape.textFrame.textRange;
  if (style.font !== undefined) range.font.name = style.font;
  if (style.size !== undefined) range.font.size = style.size;
  if (style.bold !== undefined) range.font.bold = style.bold;
  if (style.italic !== undefined) range.font.italic = style.italic;
  if (style.color !== undefined) range.font.color = style.color;
  if (style.align !== undefined) range.paragraphFormat.horizontalAlignment = style.align;
}

/** Match a requested name against PowerPoint.GeometricShapeType, case-insensitively; unknown names get examples. */
function resolveGeometricShape(name: string): PowerPoint.GeometricShapeType {
  const known = typeof PowerPoint !== "undefined" && PowerPoint.GeometricShapeType ? (Object.values(PowerPoint.GeometricShapeType) as string[]) : null;
  if (!known) return name as PowerPoint.GeometricShapeType;
  const hit = known.find((value) => value.toLowerCase() === name.toLowerCase());
  if (!hit) {
    throw new ToolFailure("invalid_arguments", `'${name}' is not a PowerPoint shape type; examples: Rectangle, RoundRectangle, Ellipse, Triangle, Diamond, RightArrow, Chevron, Star5, Hexagon, Cloud`);
  }
  return hit as PowerPoint.GeometricShapeType;
}

export async function insertSlideElement(input: InsertSlideElementInput): Promise<unknown> {
  const element = input.element;
  if (element.type === "table") {
    requireSupport("insert_slide_element (table)", ["PowerPointApi", "1.8"], "Tables need PowerPointApi 1.8.");
    return PowerPoint.run((context) => addTable(context, input.slide_id, element));
  }
  requireSupport("insert_slide_element", PPT_GEOMETRY, "Adding shapes needs PowerPointApi 1.4.");
  return PowerPoint.run(async (context) => {
    const slide = await requireSlide(context, input.slide_id);
    const bounds = { left: element.left, top: element.top, width: element.width, height: element.height };
    let shape: PowerPoint.Shape;
    switch (element.type) {
      case "textbox":
        shape = slide.shapes.addTextBox(element.text, bounds);
        applyTextStyle(shape, element.style);
        break;
      case "shape": {
        shape = slide.shapes.addGeometricShape(resolveGeometricShape(element.shape), bounds);
        if (element.fill !== undefined) shape.fill.setSolidColor(element.fill);
        if (element.line !== undefined) shape.lineFormat.color = element.line;
        if (element.line_weight !== undefined) {
          if (element.line_weight === 0) shape.lineFormat.visible = false;
          else shape.lineFormat.weight = element.line_weight;
        }
        if (element.text !== undefined) {
          shape.textFrame.textRange.text = element.text;
          applyTextStyle(shape, element.style);
        }
        break;
      }
      case "line":
        shape = slide.shapes.addLine(element.connector, bounds);
        if (element.line !== undefined) shape.lineFormat.color = element.line;
        if (element.line_weight !== undefined) shape.lineFormat.weight = element.line_weight;
        break;
    }
    if (element.name) shape.name = element.name;
    shape.load("id,name,type,left,top,width,height");
    await context.sync();
    return {
      summary: `Added a ${element.type} to slide ${input.slide_id}${shape.name ? ` ("${shape.name}")` : ""}`,
      slide_id: input.slide_id,
      ref: makeRef(input.slide_id, shape.id),
      shape_id: shape.id,
      name: shape.name,
      type: String(shape.type),
      left: round(shape.left),
      top: round(shape.top),
      width: round(shape.width),
      height: round(shape.height),
    };
  });
}

export async function removeSlideElement(input: RemoveSlideElementInput): Promise<unknown> {
  requireSupport("remove_slide_element", PPT_GEOMETRY, "Shape names need PowerPointApi 1.4.");
  return PowerPoint.run(async (context) => {
    const slide = await requireSlide(context, input.slide_id);
    const wanted = [...new Set(input.shape_ids)];
    const shapes = wanted.map((id) => {
      const shape = slide.shapes.getItemOrNullObject(id);
      shape.load("isNullObject,name");
      return { id, shape };
    });
    await context.sync();
    const found = shapes.filter((item) => !item.shape.isNullObject);
    const missing = shapes.filter((item) => item.shape.isNullObject).map((item) => item.id);
    if (found.length === 0) {
      throw new ToolFailure("invalid_arguments", `none of the shapes (${wanted.join(", ")}) exist on slide ${input.slide_id}; list them with list_slide_shapes`);
    }
    const names = found.map((item) => item.shape.name);
    for (const item of found) item.shape.delete();
    await context.sync();
    return {
      summary: `Removed ${found.length} shape(s) from slide ${input.slide_id}: ${names.join(", ")}`,
      slide_id: input.slide_id,
      removed: found.map((item, i) => ({ ref: makeRef(input.slide_id, item.id), name: names[i] })),
      not_found: missing,
    };
  });
}

const ZORDER = { bring_to_front: "BringToFront", send_to_back: "SendToBack", bring_forward: "BringForward", send_backward: "SendBackward" } as const;

export async function setShapeZOrder(input: SetShapeZOrderInput): Promise<unknown> {
  requireSupport("set_shape_zorder", PPT_ZORDER, "Changing the stacking order needs PowerPointApi 1.8.");
  return PowerPoint.run(async (context) => {
    const { slideId, shapeId } = parseShapeRef(input.ref);
    const { shape } = await requireShape(context, slideId, shapeId);
    shape.load("name,zOrderPosition");
    await context.sync();
    const before = shape.zOrderPosition;
    shape.setZOrder(ZORDER[input.action]);
    await context.sync();
    shape.load("zOrderPosition");
    await context.sync();
    return {
      summary: `${input.action.replace(/_/g, " ")}: ${shape.name} is now at layer ${shape.zOrderPosition} (was ${before})`,
      ref: input.ref,
      z_before: before,
      z_after: shape.zOrderPosition,
    };
  });
}

export async function groupShapes(input: GroupShapesInput): Promise<unknown> {
  requireSupport("group_shapes", PPT_ZORDER, "Grouping needs PowerPointApi 1.8.");
  return PowerPoint.run(async (context) => {
    const slide = await requireSlide(context, input.slide_id);
    const ids = [...new Set(input.shape_ids)];
    const probes = ids.map((id) => {
      const shape = slide.shapes.getItemOrNullObject(id);
      shape.load("isNullObject");
      return { id, shape };
    });
    await context.sync();
    const missing = probes.filter((p) => p.shape.isNullObject).map((p) => p.id);
    if (missing.length) throw new ToolFailure("invalid_arguments", `shapes ${missing.join(", ")} do not exist on slide ${input.slide_id}; list them with list_slide_shapes`);
    if (ids.length < 2) throw new ToolFailure("invalid_arguments", "a group needs at least two different shapes");
    const group = slide.shapes.addGroup(ids);
    if (input.name) group.name = input.name;
    group.load("id,name");
    await context.sync();
    return {
      summary: `Grouped ${ids.length} shapes on slide ${input.slide_id} as "${group.name}"`,
      slide_id: input.slide_id,
      ref: makeRef(input.slide_id, group.id),
      shape_id: group.id,
      name: group.name,
      grouped: ids,
      note: "The grouped shapes now live inside the group; their old refs may stop working until the group is ungrouped.",
    };
  });
}

export async function ungroupShapes(input: UngroupShapesInput): Promise<unknown> {
  requireSupport("ungroup_shapes", PPT_ZORDER, "Ungrouping needs PowerPointApi 1.8.");
  return PowerPoint.run(async (context) => {
    const { slideId, shapeId } = parseShapeRef(input.ref);
    const { slide, shape } = await requireShape(context, slideId, shapeId);
    shape.load("type,name");
    await context.sync();
    if (String(shape.type) !== "Group") throw new ToolFailure("invalid_arguments", `${input.ref} is a ${String(shape.type)} shape, not a group`);
    slide.shapes.load("items/id");
    await context.sync();
    const before = new Set(slide.shapes.items.map((s) => s.id));
    shape.group.ungroup();
    await context.sync();
    slide.shapes.load("items/id,items/name");
    await context.sync();
    const released = slide.shapes.items.filter((s) => !before.has(s.id));
    return {
      summary: `Ungrouped "${shape.name}" into ${released.length} shape(s)`,
      slide_id: slideId,
      released: released.map((s) => ({ ref: makeRef(slideId, s.id), id: s.id, name: s.name })),
    };
  });
}

/* ----- align / distribute ------------------------------------------------------------------------------------ */

export interface Box {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface AlignmentPlan {
  align?: AlignShapesInput["align"];
  distribute?: AlignShapesInput["distribute"];
  relativeTo: "shapes" | "slide";
  slide?: { width: number; height: number } | null;
}

/** New left/top for every box; pure, so it can be tested without PowerPoint. */
export function computeAlignment(boxes: Box[], plan: AlignmentPlan): Map<string, { left: number; top: number }> {
  const pos = new Map(boxes.map((box) => [box.id, { left: box.left, top: box.top }]));
  const size = new Map(boxes.map((box) => [box.id, box]));
  if (plan.relativeTo === "slide" && !plan.slide) {
    throw new ToolFailure("unsupported", "the slide size is unknown on this version of PowerPoint (needs PowerPointApi 1.8 or 1.10); use relative_to \"shapes\"");
  }
  const x0 = plan.relativeTo === "slide" ? 0 : Math.min(...boxes.map((b) => b.left));
  const x1 = plan.relativeTo === "slide" ? plan.slide!.width : Math.max(...boxes.map((b) => b.left + b.width));
  const y0 = plan.relativeTo === "slide" ? 0 : Math.min(...boxes.map((b) => b.top));
  const y1 = plan.relativeTo === "slide" ? plan.slide!.height : Math.max(...boxes.map((b) => b.top + b.height));

  for (const box of boxes) {
    const p = pos.get(box.id)!;
    switch (plan.align) {
      case "left": p.left = x0; break;
      case "right": p.left = x1 - box.width; break;
      case "center": p.left = (x0 + x1) / 2 - box.width / 2; break;
      case "top": p.top = y0; break;
      case "bottom": p.top = y1 - box.height; break;
      case "middle": p.top = (y0 + y1) / 2 - box.height / 2; break;
    }
  }

  if (plan.distribute) {
    const horizontal = plan.distribute === "horizontal";
    const minimum = plan.relativeTo === "slide" ? 2 : 3;
    if (boxes.length < minimum) {
      throw new ToolFailure("invalid_arguments", `distributing ${plan.relativeTo === "slide" ? "across the slide" : "between shapes"} needs at least ${minimum} shapes`);
    }
    const key = horizontal ? "left" : "top";
    const extent = (id: string) => (horizontal ? size.get(id)!.width : size.get(id)!.height);
    const ordered = [...boxes].sort((a, b) => pos.get(a.id)![key] - pos.get(b.id)![key]);
    const lo = plan.relativeTo === "slide" ? 0 : Math.min(...ordered.map((b) => pos.get(b.id)![key]));
    const hi =
      plan.relativeTo === "slide"
        ? horizontal ? plan.slide!.width : plan.slide!.height
        : Math.max(...ordered.map((b) => pos.get(b.id)![key] + extent(b.id)));
    const gap = (hi - lo - ordered.reduce((sum, b) => sum + extent(b.id), 0)) / (ordered.length - 1);
    let cursor = lo;
    for (const box of ordered) {
      pos.get(box.id)![key] = cursor;
      cursor += extent(box.id) + gap;
    }
  }
  for (const p of pos.values()) {
    p.left = Math.round(p.left * 100) / 100;
    p.top = Math.round(p.top * 100) / 100;
  }
  return pos;
}

export async function alignShapes(input: AlignShapesInput): Promise<unknown> {
  requireSupport("align_shapes", PPT_GEOMETRY, "Moving shapes needs PowerPointApi 1.4.");
  return PowerPoint.run(async (context) => {
    const slide = await requireSlide(context, input.slide_id);
    const ids = [...new Set(input.shape_ids)];
    const probes = ids.map((id) => {
      const shape = slide.shapes.getItemOrNullObject(id);
      shape.load("isNullObject,name,left,top,width,height");
      return { id, shape };
    });
    await context.sync();
    const missing = probes.filter((p) => p.shape.isNullObject).map((p) => p.id);
    if (missing.length) throw new ToolFailure("invalid_arguments", `shapes ${missing.join(", ")} do not exist on slide ${input.slide_id}; list them with list_slide_shapes`);

    const relativeTo = input.relative_to ?? "shapes";
    const slideSize = relativeTo === "slide" ? await getSlideSize(context, input.slide_id) : null;
    const boxes: Box[] = probes.map((p) => ({ id: p.id, left: p.shape.left, top: p.shape.top, width: p.shape.width, height: p.shape.height }));
    const plan = computeAlignment(boxes, { align: input.align, distribute: input.distribute, relativeTo, slide: slideSize });

    const moved: Array<{ ref: string; name: string; left: number; top: number }> = [];
    for (const probe of probes) {
      const next = plan.get(probe.id)!;
      if (Math.abs(next.left - probe.shape.left) < 0.01 && Math.abs(next.top - probe.shape.top) < 0.01) continue;
      probe.shape.left = next.left;
      probe.shape.top = next.top;
      moved.push({ ref: makeRef(input.slide_id, probe.id), name: probe.shape.name, left: next.left, top: next.top });
    }
    await context.sync();
    const what = [input.align && `aligned ${input.align}`, input.distribute && `distributed ${input.distribute}ly`].filter(Boolean).join(" and ");
    return {
      summary: moved.length ? `${what[0]!.toUpperCase()}${what.slice(1)}: moved ${moved.length} of ${ids.length} shapes on slide ${input.slide_id}` : `Nothing to move: the shapes were already ${what}`,
      slide_id: input.slide_id,
      relative_to: relativeTo,
      moved,
    };
  });
}

/* ----- alt text and hyperlink -------------------------------------------------------------------------------- */

export async function setShapeAltText(input: SetShapeAltTextInput): Promise<unknown> {
  return PowerPoint.run(async (context) => {
    const { slideId, shapeId } = parseShapeRef(input.ref);
    const { shape } = await requireShape(context, slideId, shapeId);
    if (isSupported(PPT_MODERN)) {
      if (input.title !== undefined) shape.altTextTitle = input.title;
      if (input.description !== undefined) shape.altTextDescription = input.description;
      await context.sync();
      return { summary: `Set the alt text of ${input.ref}`, ref: input.ref, title: input.title, description: input.description, method: "api" };
    }
    requireSupport("set_shape_alt_text", PPT_EXPORT, "Alt text needs PowerPointApi 1.10, or 1.8 for the slower file-based fallback.");
    const result = await roundTripSlide(context, {
      slideId,
      mode: "replace",
      mutate: (doc) => {
        const target = Array.from(doc.getElementsByTagNameNS(NS.p, "cNvPr")).find((node) => node.getAttribute("id") === shapeId);
        if (!target) throw new ToolFailure("invalid_arguments", `shape id ${shapeId} was not found in the slide XML`);
        if (input.title !== undefined) target.setAttribute("title", input.title);
        if (input.description !== undefined) target.setAttribute("descr", input.description);
      },
    });
    return {
      summary: `Set the alt text of shape ${shapeId} (slide rebuilt, new id ${result.newSlideId})`,
      ref: makeRef(result.newSlideId, shapeId),
      slide_id: result.newSlideId,
      replaced_slide_id: slideId,
      method: "ooxml",
      note: "The slide was rebuilt through its .pptx, so it has a new slide id.",
      ...(result.warnings.length ? { warnings: result.warnings } : {}),
    };
  });
}

export async function setSlideHyperlink(input: SetSlideHyperlinkInput): Promise<unknown> {
  requireSupport("set_slide_hyperlink", PPT_MODERN, "Hyperlinks on shapes need PowerPointApi 1.10; ask the person to add the link by hand (Insert > Link).");
  return PowerPoint.run(async (context) => {
    const { slideId, shapeId } = parseShapeRef(input.ref);
    const { shape } = await requireShape(context, slideId, shapeId);
    shape.setHyperlink({ address: input.url, ...(input.tooltip !== undefined ? { screenTip: input.tooltip } : {}) });
    await context.sync();
    return {
      summary: `Linked ${input.ref} to ${input.url}`,
      ref: input.ref,
      url: input.url,
      note: "A slide with an external link cannot be rebuilt by edit_slide_xml or duplicate_slide; remove the link first if you need those.",
    };
  });
}
