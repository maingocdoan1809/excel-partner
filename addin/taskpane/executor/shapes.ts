import {
  DASH_STYLES,
  EIGHT_SITE_SHAPES,
  FOUR_SITE_SHAPES,
  type AddConnectorInput,
  type AddLineInput,
  type AddShapeInput,
  type ConnectorSide,
  type DeleteShapesInput,
  type ListShapesInput,
  type ShapeInfo,
  type UpdateShapeInput,
} from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet } from "./common.js";

/** Cells are inset a little so shapes drawn "over" a range do not touch the gridlines. */
const ANCHOR_INSET_PT = 2;
/** Properties every shape has. `geometricShapeType` is separate: loading it on a line/connector raises InvalidOperation. */
const BASE_PROPS = "id,name,type,left,top,width,height,rotation";

type Side = Exclude<ConnectorSide, "auto">;
type Placement = Pick<AddShapeInput, "anchor" | "left" | "top" | "width" | "height">;
type Style = Pick<
  AddShapeInput,
  "fillColor" | "lineColor" | "lineWeight" | "fontColor" | "fontSize" | "fontName" | "bold" | "horizontalAlignment"
>;

function requireShapeApi(): void {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    throw new ToolFailure("unsupported", "this Excel version cannot draw shapes (needs ExcelApi 1.9)");
  }
}

/** Connection-site index of a side, or null when the shape's site layout is not known. */
export function siteFor(shapeType: string, side: Side): number | null {
  if ((FOUR_SITE_SHAPES as readonly string[]).includes(shapeType)) {
    return { top: 0, left: 1, bottom: 2, right: 3 }[side];
  }
  if ((EIGHT_SITE_SHAPES as readonly string[]).includes(shapeType)) {
    return { top: 0, left: 2, bottom: 4, right: 6 }[side];
  }
  return null;
}

/** The side of `a` that faces `b`; vertical wins ties so flowcharts read top-down. */
export function facingSide(a: Box, b: Box): Side {
  const dx = b.left + b.width / 2 - (a.left + a.width / 2);
  const dy = b.top + b.height / 2 - (a.top + a.height / 2);
  if (Math.abs(dy) >= Math.abs(dx)) return dy >= 0 ? "bottom" : "top";
  return dx >= 0 ? "right" : "left";
}

interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

function edgePoint(box: Box, side: Side): { x: number; y: number } {
  const cx = box.left + box.width / 2;
  const cy = box.top + box.height / 2;
  return {
    top: { x: cx, y: box.top },
    bottom: { x: cx, y: box.top + box.height },
    left: { x: box.left, y: cy },
    right: { x: box.left + box.width, y: cy },
  }[side];
}

/** Resolve `anchor` (cell range) or explicit points into a box; partial updates keep the rest. */
async function resolveBox(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  place: Placement,
  current?: Box,
): Promise<Box | null> {
  if (place.anchor) {
    const range = sheet.getRange(place.anchor);
    range.load("left,top,width,height");
    await context.sync();
    const inset = Math.min(ANCHOR_INSET_PT, range.width / 4, range.height / 4);
    return {
      left: range.left + inset,
      top: range.top + inset,
      width: Math.max(1, range.width - 2 * inset),
      height: Math.max(1, range.height - 2 * inset),
    };
  }
  const given = [place.left, place.top, place.width, place.height];
  if (given.every((v) => v === undefined)) return null;
  if (!current && given.some((v) => v === undefined)) {
    throw new ToolFailure("invalid_arguments", "give all of left/top/width/height, or an `anchor` range");
  }
  return {
    left: place.left ?? current!.left,
    top: place.top ?? current!.top,
    width: place.width ?? current!.width,
    height: place.height ?? current!.height,
  };
}

/**
 * Excel's shape-outline API rejects fractional point sizes (verified on a real
 * install: 1, 3 and 6 are accepted; 0.5, 1.25, 1.5, 2.5 fail with InvalidArgument), so
 * weights are rounded to whole points, at least 1. Returns a note when it changed the value.
 */
export function wholePoints(weight: number): { value: number; note: string | null } {
  const value = Math.max(1, Math.round(weight));
  return { value, note: value === weight ? null : `line weight ${weight} rounded to ${value} pt (Excel needs whole points)` };
}

function applyStyle(shape: Excel.Shape, style: Style, text?: string, notes: string[] = []): void {
  if (style.fillColor) shape.fill.setSolidColor(style.fillColor);
  if (style.lineColor === "none") shape.lineFormat.visible = false;
  else if (style.lineColor) {
    shape.lineFormat.visible = true;
    shape.lineFormat.color = style.lineColor;
  }
  if (style.lineWeight !== undefined) {
    const { value, note } = wholePoints(style.lineWeight);
    shape.lineFormat.weight = value;
    if (note) notes.push(note);
  }

  // Touch `textFrame` only when text/font is being changed: on lines and connectors even *reading* the
  // property raises InvalidOperation, which used to break plain lineColor / rename updates.
  const wantsText =
    text !== undefined || style.fontColor || style.fontSize !== undefined || style.fontName || style.bold !== undefined || style.horizontalAlignment;
  if (!wantsText) return;
  const frame = shape.textFrame;
  if (text !== undefined) frame.textRange.text = text;
  if (style.fontColor) frame.textRange.font.color = style.fontColor;
  if (style.fontSize !== undefined) frame.textRange.font.size = style.fontSize;
  if (style.fontName) frame.textRange.font.name = style.fontName;
  if (style.bold !== undefined) frame.textRange.font.bold = style.bold;
  if (style.horizontalAlignment) frame.horizontalAlignment = style.horizontalAlignment;
}

function setBox(shape: Excel.Shape, box: Box): void {
  shape.left = box.left;
  shape.top = box.top;
  shape.width = box.width;
  shape.height = box.height;
}

/** Load a shape's properties, adding `geometricShapeType` only where Excel allows it. */
async function loadShape(context: Excel.RequestContext, shape: Excel.Shape): Promise<void> {
  shape.load(BASE_PROPS);
  await context.sync();
  if (shape.type === "GeometricShape") {
    shape.load("geometricShapeType");
    await context.sync();
  }
}

/** `FlowChartProcess` for geometric shapes, otherwise the shape type (`TextBox`, `Line`, ...). */
function shapeKind(shape: Excel.Shape): string {
  return shape.type === "GeometricShape" ? String(shape.geometricShapeType) : String(shape.type);
}

async function findShape(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  name: string,
  role = "shape",
): Promise<Excel.Shape> {
  sheet.shapes.load("items/name,items/type");
  await context.sync();
  let level = sheet.shapes.items.filter(item => item.type === "Group");
  const matches: Excel.Shape[] = sheet.shapes.items.filter(item => item.name === name);
  for (let depth = 0; level.length > 0; depth++) {
    if (depth >= 32) throw new ToolFailure("invalid_arguments", "Shape groups exceed the supported nesting depth (32).");
    for (const group of level) group.group.shapes.load("items/name,items/type");
    await context.sync();
    const children = level.flatMap(group => group.group.shapes.items);
    matches.push(...children.filter(child => child.name === name));
    level = children.filter(child => child.type === "Group");
  }
  if (matches.length > 1) throw new ToolFailure("invalid_arguments", `${role} '${name}' matches multiple shapes; identify its parent group with list_shapes before editing.`, { matches: matches.length });
  if (matches.length === 0) throw new ToolFailure("invalid_arguments", `${role} '${name}' not found on this sheet or in its groups`, {
    available: sheet.shapes.items.map((s) => s.name).slice(0, 100),
  });
  const shape = matches[0]!;
  await loadShape(context, shape);
  return shape;
}

function info(shape: Excel.Shape, text: string | null, warnings: string[] = []): ShapeInfo & { warnings?: string[] } {
  return {
    ...(warnings.length > 0 ? { warnings } : {}),
    id: shape.id,
    rotation: shape.rotation,
    name: shape.name,
    type: shapeKind(shape),
    text,
    left: round(shape.left),
    top: round(shape.top),
    width: round(shape.width),
    height: round(shape.height),
  };
}

const round = (n: number) => Math.round(n * 100) / 100;

type TextRead = Pick<ShapeInfo, "text" | "textReadStatus" | "readWarnings">;

function unreadableText(error: unknown): TextRead {
  const detail = error as { code?: string; message?: string };
  return {
    text: null,
    textReadStatus: "unavailable",
    readWarnings: [{
      property: "text", code: detail?.code ?? "read_failed",
      message: `Text could not be read; the shape and its geometry were retained. ${detail?.message ?? String(error)}`.slice(0, 600),
    }],
  };
}

/**
 * Optional text reads must not invalidate shape/connector inspection. Try a batch
 * first; if Excel rejects it during sync, retry each read to identify the offender.
 * Store the loaded proxy rather than navigating textRange again during serialization.
 */
async function readShapeTexts(context: Excel.RequestContext, shapes: Excel.Shape[]): Promise<Map<Excel.Shape, TextRead>> {
  const result = new Map<Excel.Shape, TextRead>();
  const candidates = shapes.filter(shape => shape.type === "GeometricShape" || (shape.type as string) === "TextBox");
  for (const shape of shapes) result.set(shape, { text: null, textReadStatus: "not_applicable" });

  async function readBatch<T>(items: Excel.Shape[], prepare: (shape: Excel.Shape) => () => T): Promise<Map<Excel.Shape, T>> {
    const values = new Map<Excel.Shape, T>();
    const pending: { shape: Excel.Shape; read: () => T }[] = [];
    for (const shape of items) {
      try { pending.push({ shape, read: prepare(shape) }); }
      catch (error) { result.set(shape, unreadableText(error)); }
    }
    let synced = false;
    try { await context.sync(); synced = true; } catch { /* Retry separately below; the rejected batch has been consumed. */ }
    for (const item of pending) {
      try {
        let read = item.read;
        if (!synced) { read = prepare(item.shape); await context.sync(); }
        values.set(item.shape, read());
      } catch (error) { result.set(item.shape, unreadableText(error)); }
    }
    return values;
  }

  const hasText = await readBatch(candidates, shape => {
    const frame = shape.textFrame;
    frame.load("hasText");
    return () => frame.hasText;
  });
  const withText: Excel.Shape[] = [];
  for (const [shape, present] of hasText) {
    if (present) withText.push(shape);
    else result.set(shape, { text: "", textReadStatus: "empty" });
  }
  if (withText.length) {
    const texts = await readBatch(withText, shape => {
      const range = shape.textFrame.textRange;
      range.load("text");
      return () => range.text;
    });
    for (const [shape, text] of texts) result.set(shape, { text, textReadStatus: text ? "read" : "empty" });
  }
  return result;
}

/** Read type-specific properties only after the base type has been synced. */
export async function describeShapes(context: Excel.RequestContext, shapes: Excel.Shape[]): Promise<ShapeInfo[]> {
  for (const shape of shapes) {
    if (shape.type === "GeometricShape") shape.load("geometricShapeType");
    if (shape.type === "Line") shape.line.load("isBeginConnected,isEndConnected,beginConnectedSite,endConnectedSite,connectorType,beginArrowheadStyle,endArrowheadStyle");
  }
  await context.sync();
  for (const shape of shapes) {
    if (shape.type !== "Line") continue;
    if (shape.line.isBeginConnected) shape.line.beginConnectedShape.load("id,name");
    if (shape.line.isEndConnected) shape.line.endConnectedShape.load("id,name");
  }
  await context.sync();
  const texts = await readShapeTexts(context, shapes);
  return shapes.map(shape => ({
    ...info(shape, texts.get(shape)!.text),
    ...texts.get(shape),
    ...(shape.type === "Line" ? {
      connection: {
        from: shape.line.isBeginConnected ? { id: shape.line.beginConnectedShape.id, name: shape.line.beginConnectedShape.name, site: shape.line.beginConnectedSite } : null,
        to: shape.line.isEndConnected ? { id: shape.line.endConnectedShape.id, name: shape.line.endConnectedShape.name, site: shape.line.endConnectedSite } : null,
        connectorType: String(shape.line.connectorType),
        beginArrow: String(shape.line.beginArrowheadStyle),
        endArrow: String(shape.line.endArrowheadStyle),
      },
    } : {}),
  }));
}

export async function readActiveShape(context: Excel.RequestContext): Promise<{ activeShape?: ShapeInfo; shapeSelectionStatus: "active" | "none" | "unavailable" }> {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.19")) return { shapeSelectionStatus: "unavailable" };
  const shape = context.workbook.getActiveShapeOrNullObject();
  shape.load("isNullObject");
  await context.sync();
  if (shape.isNullObject) return { shapeSelectionStatus: "none" };
  await loadShape(context, shape);
  return { activeShape: (await describeShapes(context, [shape]))[0]!, shapeSelectionStatus: "active" };
}

export async function listShapes(input: ListShapesInput): Promise<unknown> {
  requireShapeApi();
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    sheet.shapes.load(`items/${BASE_PROPS.split(",").join(",items/")}`);
    await context.sync();

    // Flatten groups in preorder; keep parent IDs so a whole flow can be reconstructed.
    const all: { shape: Excel.Shape; parentGroupId?: string }[] = [];
    const scan = async (shapes: Excel.Shape[], parentGroupId?: string, depth = 0): Promise<void> => {
      if (depth > 32) throw new ToolFailure("invalid_arguments", "Shape groups exceed the supported nesting depth (32).");
      for (const shape of shapes) {
        all.push({ shape, parentGroupId });
        if ((shape.type as string) === "Group") {
          shape.group.shapes.load(`items/${BASE_PROPS.split(",").join(",items/")}`);
          await context.sync();
          await scan(shape.group.shapes.items, shape.id, depth + 1);
        }
      }
    };
    await scan(sheet.shapes.items);
    const offset = input.offset ?? 0;
    const page = all.slice(offset, offset + (input.limit ?? 200));
    const shapes = (await describeShapes(context, page.map(item => item.shape))).map((shape, i) => ({
      ...shape, ...(page[i]!.parentGroupId ? { parentGroupId: page[i]!.parentGroupId } : {}),
    }));
    const hasMore = offset + shapes.length < all.length;
    return {
      sheet: sheet.name,
      count: all.length,
      returned: shapes.length,
      hasMore,
      nextOffset: hasMore ? offset + shapes.length : null,
      shapes,
      readWarnings: shapes.flatMap(shape => (shape.readWarnings ?? []).map(warning => ({ id: shape.id, name: shape.name, ...warning }))),
      textReadComplete: !shapes.some(shape => shape.textReadStatus === "unavailable"),
      validation: validateFlow(shapes, offset > 0 || hasMore),
    };
  });
}

/** Structural evidence only: business correctness still needs the user's flow requirements. */
export function validateFlow(shapes: ShapeInfo[], partial: boolean) {
  const nodes = shapes.filter(s => s.type !== "Line" && s.type !== "Group" && s.type !== "TextBox" && s.type !== "Image");
  const edges = shapes.filter(s => s.connection);
  const linked = new Set(edges.flatMap(s => [s.connection!.from?.id, s.connection!.to?.id]).filter(Boolean));
  const issues: { code: string; shapes: string[] }[] = [];
  for (const edge of edges) {
    const c = edge.connection!;
    if (!c.from || !c.to) issues.push({ code: "unattached_line", shapes: [edge.name] });
    if (c.from && c.to && c.from.id === c.to.id) issues.push({ code: "self_connection", shapes: [edge.name, c.from.name] });
  }
  const adjacency = new Map(nodes.map(node => [node.id!, new Set<string>()]));
  for (const edge of edges) {
    const { from, to } = edge.connection!;
    if (from && to && adjacency.has(from.id) && adjacency.has(to.id)) {
      adjacency.get(from.id)!.add(to.id);
      adjacency.get(to.id)!.add(from.id);
    }
  }
  const components: string[][] = [];
  const visited = new Set<string>();
  for (const node of nodes) {
    if (visited.has(node.id!)) continue;
    const component: string[] = [];
    const queue = [node.id!];
    while (queue.length) {
      const id = queue.pop()!;
      if (visited.has(id)) continue;
      visited.add(id);
      component.push(id);
      queue.push(...adjacency.get(id)!);
    }
    components.push(component);
  }
  if (!partial && nodes.length > 1) {
    for (const node of nodes) if (!linked.has(node.id)) issues.push({ code: "isolated_shape", shapes: [node.name] });
  }
  return {
    scope: partial ? "partial_page" : "whole_sheet",
    complete: !partial,
    nodeCount: nodes.length,
    lineCount: edges.length,
    components,
    issues,
    note: "Structural checks only. Unattached lines may be intentional decorations; isolated shapes may belong to separate flows. Read every page before judging connectivity. Check business logic against the user's requirements; use an image for visual layout.",
  };
}

export async function addShape(input: AddShapeInput): Promise<unknown> {
  requireShapeApi();
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const box = (await resolveBox(context, sheet, input))!;

    const name = input.name ?? `${input.type}_${Math.random().toString(36).slice(2, 7)}`;
    const existing = sheet.shapes.getItemOrNullObject(name);
    existing.load("isNullObject");
    await context.sync();
    if (!existing.isNullObject) {
      throw new ToolFailure("invalid_arguments", `a shape named '${name}' already exists; pick another name or use update_shape`);
    }

    const shape =
      input.type === "TextBox"
        ? sheet.shapes.addTextBox(input.text ?? "")
        : sheet.shapes.addGeometricShape(input.type as Excel.GeometricShapeType);
    shape.name = name;
    setBox(shape, box);
    shape.textFrame.verticalAlignment = "Middle";
    shape.textFrame.horizontalAlignment = input.horizontalAlignment ?? "Center";
    const warnings: string[] = [];
    applyStyle(shape, input, input.type === "TextBox" ? undefined : input.text, warnings);
    await context.sync();
    await loadShape(context, shape);
    const created = info(shape, input.text ?? null, warnings);
    // Some Excel builds report a shape from addTextBox as a plain Rectangle. Say what was created.
    if (input.type === "TextBox" && created.type !== "TextBox") {
      return { ...created, type: "TextBox", hostReportedType: created.type };
    }
    return created;
  });
}

/** Current text of a shape, so results never claim "no text" for a shape that has some. */
async function currentText(context: Excel.RequestContext, shape: Excel.Shape): Promise<TextRead> {
  return (await readShapeTexts(context, [shape])).get(shape)!;
}

/** Properties that only make sense on shapes with a body; lines and connectors have an outline only. */
const BODY_ONLY_KEYS = ["text", "fillColor", "fontColor", "fontSize", "fontName", "bold", "horizontalAlignment"] as const;

export async function updateShape(input: UpdateShapeInput): Promise<unknown> {
  requireShapeApi();
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const shape = await findShape(context, sheet, input.name);
    const isLine = shape.type === "Line";

    if (isLine) {
      // Lines/connectors reject body operations with an opaque InvalidOperation; say what is possible instead.
      const unsupported = BODY_ONLY_KEYS.filter((key) => input[key] !== undefined);
      if (unsupported.length > 0) {
        throw new ToolFailure(
          "invalid_arguments",
          `'${input.name}' is a line/connector, so ${unsupported.join(", ")} cannot be changed. Only lineColor, lineWeight and newName apply; ` +
            "to change anything else delete it with delete_shapes and add it again.",
        );
      }
    }

    const box = await resolveBox(context, sheet, input, shape);
    if (box) setBox(shape, box);
    const warnings: string[] = [];
    applyStyle(shape, input, input.text, warnings);
    if (input.newName && input.newName !== input.name) {
      const clash = sheet.shapes.getItemOrNullObject(input.newName);
      clash.load("isNullObject");
      await context.sync();
      if (!clash.isNullObject) {
        throw new ToolFailure("invalid_arguments", `a shape named '${input.newName}' already exists`);
      }
      shape.name = input.newName;
    }
    await context.sync();
    await loadShape(context, shape);
    const text = await currentText(context, shape);
    return { ...info(shape, text.text, warnings), ...text };
  });
}

export async function deleteShapes(input: DeleteShapesInput): Promise<unknown> {
  requireShapeApi();
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const found = input.names.map((name) => {
      const shape = sheet.shapes.getItemOrNullObject(name);
      shape.load("isNullObject");
      return { name, shape };
    });
    await context.sync();

    const deleted: string[] = [];
    const notFound: string[] = [];
    for (const { name, shape } of found) {
      if (shape.isNullObject) notFound.push(name);
      else {
        shape.delete();
        deleted.push(name);
      }
    }
    await context.sync();
    return { sheet: sheet.name, deleted, notFound };
  });
}

/** Dash style and arrowheads shared by connectors and free lines. */
function styleLine(line: Excel.Shape, input: { arrow: "end" | "both" | "none" | "diamond"; dash: keyof typeof DASH_STYLES }): void {
  if (input.dash !== "solid") line.lineFormat.dashStyle = DASH_STYLES[input.dash] as Excel.ShapeLineDashStyle;
  if (input.arrow === "diamond") line.line.endArrowheadStyle = "Diamond";
  else if (input.arrow !== "none") line.line.endArrowheadStyle = "Triangle";
  if (input.arrow === "both") line.line.beginArrowheadStyle = "Triangle";
}

export async function addLine(input: AddLineInput): Promise<unknown> {
  requireShapeApi();
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const start = sheet.getRange(input.from);
    const end = sheet.getRange(input.to);
    start.load("left,top,width,height");
    end.load("left,top,width,height");
    await context.sync();

    const name = input.name ?? `Line_${Math.random().toString(36).slice(2, 7)}`;
    const clash = sheet.shapes.getItemOrNullObject(name);
    clash.load("isNullObject");
    await context.sync();
    if (!clash.isNullObject) throw new ToolFailure("invalid_arguments", `a shape named '${name}' already exists`);

    const centre = (r: Excel.Range) => ({ x: r.left + r.width / 2, y: r.top + r.height / 2 });
    const a = centre(start);
    const b = centre(end);
    const line = sheet.shapes.addLine(a.x, a.y, b.x, b.y, Excel.ConnectorType.straight);
    line.name = name;
    styleLine(line, input);
    if (input.color) line.lineFormat.color = input.color;
    const warnings: string[] = [];
    if (input.weight !== undefined) {
      const { value, note } = wholePoints(input.weight);
      line.lineFormat.weight = value;
      if (note) warnings.push(note);
    }
    await context.sync();
    return { sheet: sheet.name, name, from: input.from, to: input.to, ...(warnings.length > 0 ? { warnings } : {}) };
  });
}

export async function addConnector(input: AddConnectorInput): Promise<unknown> {
  requireShapeApi();
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const from = await findShape(context, sheet, input.from, "source shape");
    const to = await findShape(context, sheet, input.to, "target shape");

    const fromSide: Side = input.fromSide === "auto" ? facingSide(from, to) : input.fromSide;
    const toSide: Side = input.toSide === "auto" ? facingSide(to, from) : input.toSide;
    const fromSite = input.fromSite ?? siteFor(shapeKind(from), fromSide);
    const toSite = input.toSite ?? siteFor(shapeKind(to), toSide);
    if (fromSite === null || toSite === null) {
      throw new ToolFailure(
        "invalid_arguments",
        `connection sites are only known for rectangle/flowchart/ellipse shapes; pass fromSite/toSite for ${shapeKind(from)} / ${shapeKind(to)}`,
      );
    }

    const start = edgePoint(from, fromSide);
    const end = edgePoint(to, toSide);
    const name = input.name ?? `Connector_${Math.random().toString(36).slice(2, 7)}`;
    const warnings: string[] = [];
    const line = sheet.shapes.addLine(start.x, start.y, end.x, end.y, input.connectorType);
    line.name = name;
    try {
      line.line.connectBeginShape(from, fromSite);
      line.line.connectEndShape(to, toSite);
      styleLine(line, input);
      if (input.color) line.lineFormat.color = input.color;
      if (input.weight !== undefined) {
        const { value, note } = wholePoints(input.weight);
        line.lineFormat.weight = value;
        if (note) warnings.push(note);
      }
      await context.sync();
    } catch (error) {
      // A rejected connection site leaves an unattached line behind; remove it.
      const leftover = sheet.shapes.getItemOrNullObject(name);
      leftover.load("isNullObject");
      await context.sync().catch(() => undefined);
      if (!leftover.isNullObject) {
        leftover.delete();
        await context.sync().catch(() => undefined);
      }
      throw new ToolFailure(
        "invalid_arguments",
        `could not attach connector (${error instanceof Error ? error.message : String(error)}); check fromSite/toSite`,
      );
    }
    return {
      sheet: sheet.name,
      name,
      from: input.from,
      to: input.to,
      fromSide,
      toSide,
      fromSite,
      toSite,
      ...(warnings.length > 0 ? { warnings } : {}),
    };
  });
}
