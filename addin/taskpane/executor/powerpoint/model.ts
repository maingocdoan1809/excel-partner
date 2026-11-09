import { PPT_GEOMETRY, PPT_ZORDER, isSupported, normalizeText } from "./util.js";

/** Shape types that carry a text frame. Pictures, lines, groups, tables and charts do not. */
export const TEXT_TYPES: ReadonlySet<string> = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"]);

export interface ShapeInfo {
  id: string;
  name: string;
  type: string;
  /** Position in the slide's z-order, 1 = back. */
  z: number;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  /** Placeholder kind (Title, Body, ...) when known. */
  placeholder?: string | null;
  /** What a placeholder holds (Image, Table, ...); null when empty. */
  contained?: string | null;
  /** Text, or undefined when the shape has none or its text could not be read. */
  text?: string;
  hasText?: boolean;
  fontSize?: number | null;
  fontColor?: string | null;
  fillType?: string;
  fillColor?: string;
  autoSize?: string;
  wordWrap?: boolean;
  margins?: { left: number; top: number; right: number; bottom: number };
  /** The live proxy, for follow-up reads and writes. Never serialised. */
  shape: PowerPoint.Shape;
}

export interface LoadShapeOptions {
  geometry?: boolean;
  text?: boolean;
  /** Font size / colour, fill, auto-size and margins: what verify_slides needs. */
  style?: boolean;
  placeholders?: boolean;
}

/**
 * Queue `queue(item)` for every item and sync once. When that batch fails (one shape without a text frame is enough),
 * retry item by item so one bad shape does not hide the others. Returns which items loaded.
 */
export async function tryEach<T>(context: PowerPoint.RequestContext, items: T[], queue: (item: T) => void): Promise<boolean[]> {
  if (items.length === 0) return [];
  try {
    for (const item of items) queue(item);
    await context.sync();
    return items.map(() => true);
  } catch {
    const loaded: boolean[] = [];
    for (const item of items) {
      try {
        queue(item);
        await context.sync();
        loaded.push(true);
      } catch {
        loaded.push(false);
      }
    }
    return loaded;
  }
}

/** Read the shapes of a collection (identity first, then text / style / placeholder facts, each failure-isolated). */
export async function loadShapeInfos(
  context: PowerPoint.RequestContext,
  collection: PowerPoint.ShapeCollection | PowerPoint.ShapeScopedCollection,
  options: LoadShapeOptions = {},
): Promise<ShapeInfo[]> {
  // name and type arrived with PowerPointApi 1.4, together with geometry and text frames; 1.3 can only give ids.
  const modern = isSupported(PPT_GEOMETRY);
  const geometry = options.geometry === true && modern;
  const props = ["id", ...(modern ? ["name", "type"] : []), ...(geometry ? ["left", "top", "width", "height"] : [])];
  collection.load(props.map((p) => `items/${p}`).join(","));
  await context.sync();

  const infos: ShapeInfo[] = collection.items.map((shape, index) => {
    const info: ShapeInfo = { id: shape.id, name: modern ? shape.name : "", type: modern ? String(shape.type) : "Unknown", z: index + 1, shape };
    if (geometry) Object.assign(info, { left: shape.left, top: shape.top, width: shape.width, height: shape.height });
    return info;
  });

  if ((options.text || options.style) && isSupported(PPT_GEOMETRY)) {
    const textual = infos.filter((info) => TEXT_TYPES.has(info.type));
    const style = options.style === true;
    const loaded = await tryEach(context, textual, (info) => {
      const frame = info.shape.textFrame;
      frame.load(style ? "hasText,autoSizeSetting,wordWrap,leftMargin,topMargin,rightMargin,bottomMargin" : "hasText");
      frame.textRange.load("text");
      if (style) frame.textRange.font.load("size,color");
    });
    textual.forEach((info, i) => {
      if (!loaded[i]) return;
      const frame = info.shape.textFrame;
      info.hasText = frame.hasText;
      info.text = normalizeText(frame.textRange.text);
      if (style) {
        info.fontSize = frame.textRange.font.size;
        info.fontColor = frame.textRange.font.color;
        info.autoSize = String(frame.autoSizeSetting);
        info.wordWrap = frame.wordWrap;
        info.margins = { left: frame.leftMargin, top: frame.topMargin, right: frame.rightMargin, bottom: frame.bottomMargin };
      }
    });
    if (style) {
      const withText = textual.filter((_, i) => loaded[i]);
      const filled = await tryEach(context, withText, (info) => info.shape.fill.load("foregroundColor,type"));
      withText.forEach((info, i) => {
        if (!filled[i]) return;
        info.fillType = String(info.shape.fill.type);
        info.fillColor = info.shape.fill.foregroundColor;
      });
    }
  }

  if (options.placeholders && isSupported(PPT_ZORDER)) {
    const placeholders = infos.filter((info) => info.type === "Placeholder");
    const loaded = await tryEach(context, placeholders, (info) => info.shape.placeholderFormat.load("type,containedType"));
    placeholders.forEach((info, i) => {
      if (!loaded[i]) return;
      info.placeholder = String(info.shape.placeholderFormat.type);
      const contained = info.shape.placeholderFormat.containedType;
      info.contained = contained === null || contained === undefined ? null : String(contained);
    });
  }
  return infos;
}

const TITLE_PLACEHOLDERS = new Set(["Title", "CenterTitle", "VerticalTitle"]);

/** The shape that is most likely the slide's title: a title placeholder, else a shape named like one. */
export function findTitleShape(infos: ShapeInfo[]): ShapeInfo | undefined {
  return (
    infos.find((info) => info.placeholder && TITLE_PLACEHOLDERS.has(info.placeholder)) ??
    infos.find((info) => info.type === "Placeholder" && /(^|\b)(title|ctr ?title|center title|tiêu đề|タイトル|titel|titre|título)/i.test(info.name))
  );
}

/** Title text of one slide's shapes (first non-empty line), "" when it has no title shape. */
export async function readSlideTitle(context: PowerPoint.RequestContext, infos: ShapeInfo[]): Promise<string> {
  const title = findTitleShape(infos);
  if (!title || !isSupported(PPT_GEOMETRY)) return "";
  const [ok] = await tryEach(context, [title], (info) => info.shape.textFrame.textRange.load("text"));
  if (!ok) return "";
  const line = normalizeText(title.shape.textFrame.textRange.text).split("\n").find((l) => l.trim() !== "")?.trim() ?? "";
  return line.length > 150 ? `${line.slice(0, 150)}...` : line;
}

/** The fields the model sees for a shape. */
export function publicShape(slideId: string, info: ShapeInfo, extra: Record<string, unknown> = {}): Record<string, unknown> {
  const name = info.name.length > 200 ? `${info.name.slice(0, 200)}...` : info.name;
  const out: Record<string, unknown> = { ref: `${slideId}:${info.id}`, id: info.id, name, type: info.type, z: info.z };
  if (info.placeholder) out.placeholder = info.placeholder;
  if (info.contained) out.contains = info.contained;
  if (info.left !== undefined) {
    Object.assign(out, { left: round1(info.left), top: round1(info.top!), width: round1(info.width!), height: round1(info.height!) });
  }
  return Object.assign(out, extra);
}

function round1(value: number): number {
  return Math.round(value * 10) / 10;
}
