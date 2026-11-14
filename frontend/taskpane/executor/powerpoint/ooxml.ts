/// <reference lib="dom" />
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { ToolFailure } from "../common.js";
import { PPT_EXPORT, PPT_LAYOUTS, PPT_MODERN, isSupported, loadSlideIds, requireSlide, requireSupport, slideNotFound } from "./util.js";

/**
 * PowerPoint has no typed API for everything (notes, images, backgrounds before 1.10, duplicating a slide), so some tools
 * export one slide as a .pptx (Slide.exportAsBase64, PowerPointApi 1.8), change its XML, and insert it back
 * (Presentation.insertSlidesFromBase64, 1.2). The inserted slide gets a NEW id; the old one is deleted only after the
 * new one is confirmed present.
 *
 * Formatting choice: KeepSourceFormatting. The exported file carries the slide's own layout, master and theme, and
 * "keep source" makes the copy look exactly like the original. UseDestinationTheme would re-map the slide onto the
 * deck's layouts by name, which silently changes the look when two layouts share a name. The cost: PowerPoint may add a
 * second copy of the master when it does not recognise the exported one as identical; round trips detect that and
 * report it in `warnings` (the extra master can be deleted in Slide Master view).
 */
export const INSERT_FORMATTING = "KeepSourceFormatting" as const;

export const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
} as const;

const EMU_PER_POINT = 12700;

/* ----- bytes <-> base64 -------------------------------------------------------------------------------------- */

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64.replace(/\s+/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk) as unknown as number[]);
  }
  return btoa(binary);
}

/* ----- XML ----------------------------------------------------------------------------------------------------- */

export function parseXml(text: string, label: string): Document {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) throw new ToolFailure("office_error", `${label} is not well-formed XML: ${(error.textContent ?? "").slice(0, 200)}`);
  return doc;
}

export function serializeXml(doc: Document): string {
  const text = new XMLSerializer().serializeToString(doc);
  return text.startsWith("<?xml") ? text : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${text}`;
}

/** The parts of an exported one-slide .pptx that the round trip works with. */
export interface SlideParts {
  files: Record<string, Uint8Array>;
  slidePath: string;
  relsPath: string;
}

export function unpackSlideFile(base64: string): SlideParts {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(base64ToBytes(base64));
  } catch (error) {
    throw new ToolFailure("office_error", `the exported slide is not a readable .pptx: ${error instanceof Error ? error.message : String(error)}`);
  }
  const slides = Object.keys(files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(/\d+/.exec(a)![0]) - Number(/\d+/.exec(b)![0]));
  const slidePath = slides[0];
  if (!slidePath) throw new ToolFailure("office_error", "the exported file has no ppt/slides/slide*.xml part");
  const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
  return { files, slidePath, relsPath };
}

export function packSlideFile(files: Record<string, Uint8Array>): string {
  return bytesToBase64(zipSync(files, { level: 6 }));
}

/** Slide size in points, read from ppt/presentation.xml (`<p:sldSz cx cy>` in EMU). */
export function slideSizeFromParts(parts: SlideParts): SlideSize | null {
  const bytes = parts.files["ppt/presentation.xml"];
  if (!bytes) return null;
  try {
    const size = parseXml(strFromU8(bytes), "ppt/presentation.xml").getElementsByTagNameNS(NS.p, "sldSz")[0];
    const cx = Number(size?.getAttribute("cx"));
    const cy = Number(size?.getAttribute("cy"));
    return cx > 0 && cy > 0 ? { width: cx / EMU_PER_POINT, height: cy / EMU_PER_POINT, source: "presentation.xml" } : null;
  } catch {
    return null;
  }
}

export interface SlideSize {
  width: number;
  height: number;
  source: "pageSetup" | "presentation.xml";
}

/** Slide size from presentation.pageSetup (1.10), else from an exported slide's presentation.xml (1.8), else null. */
export async function getSlideSize(context: PowerPoint.RequestContext, firstSlideId?: string): Promise<SlideSize | null> {
  if (isSupported(PPT_MODERN)) {
    try {
      const setup = context.presentation.pageSetup;
      setup.load("slideWidth,slideHeight");
      await context.sync();
      if (setup.slideWidth > 0 && setup.slideHeight > 0) return { width: setup.slideWidth, height: setup.slideHeight, source: "pageSetup" };
    } catch {
      /* fall through to the export fallback */
    }
  }
  if (isSupported(PPT_EXPORT)) {
    try {
      const id = firstSlideId ?? (await loadSlideIds(context))[0];
      if (id) {
        const exported = context.presentation.slides.getItem(id).exportAsBase64();
        await context.sync();
        return slideSizeFromParts(unpackSlideFile(exported.value));
      }
    } catch {
      /* size stays unknown */
    }
  }
  return null;
}

/** Relationships of the slide that point outside the file (hyperlinks, linked media). insertSlidesFromBase64 refuses those. */
export function findExternalRelationships(parts: SlideParts): Array<{ type: string; target: string }> {
  const rels = parts.files[parts.relsPath];
  if (!rels) return [];
  const doc = parseXml(strFromU8(rels), parts.relsPath);
  const found: Array<{ type: string; target: string }> = [];
  for (const rel of Array.from(doc.getElementsByTagName("Relationship"))) {
    if ((rel.getAttribute("TargetMode") ?? "").toLowerCase() === "external") {
      found.push({ type: (rel.getAttribute("Type") ?? "").split("/").pop() ?? "", target: (rel.getAttribute("Target") ?? "").slice(0, 120) });
    }
  }
  return found;
}

/* ----- helpers handed to edit_slide_xml code --------------------------------------------------------------- */

type ElChild = Node | string | number | null | undefined;

function namespaceFor(doc: Document, prefix: string): string | null {
  return (NS as Record<string, string>)[prefix] ?? doc.documentElement?.lookupNamespaceURI(prefix) ?? null;
}

/** `el("p:sp", { id: 3 }, child, "text")`: build a namespaced element; "r:id" style attributes get the right namespace. */
export function makeElementHelper(doc: Document) {
  return function el(name: string, attrs?: Record<string, string | number | boolean | null | undefined> | null, ...children: ElChild[]): Element {
    const colon = name.indexOf(":");
    const prefix = colon === -1 ? null : name.slice(0, colon);
    const uri = prefix ? namespaceFor(doc, prefix) : doc.documentElement?.namespaceURI ?? null;
    if (prefix && !uri) throw new Error(`el(): unknown namespace prefix '${prefix}' (use a, p or r, or a prefix declared on the slide root)`);
    const element = doc.createElementNS(uri, name);
    for (const [key, value] of Object.entries(attrs ?? {})) {
      if (value === null || value === undefined) continue;
      const attrColon = key.indexOf(":");
      const attrUri = attrColon === -1 ? null : namespaceFor(doc, key.slice(0, attrColon));
      if (attrColon !== -1 && attrUri) element.setAttributeNS(attrUri, key, String(value));
      else element.setAttribute(key, String(value));
    }
    for (const child of children) {
      if (child === null || child === undefined) continue;
      if (typeof child === "string" || typeof child === "number") element.appendChild(doc.createTextNode(String(child)));
      else element.appendChild(child.ownerDocument === doc ? child : doc.importNode(child, true));
    }
    return element;
  };
}

interface Step {
  prefix: string | null;
  local: string;
  attr?: { name: string; value: string };
  childOnly: boolean;
}

function parseSelector(selector: string): Step[] {
  const steps: Step[] = [];
  const tokens = /\s*(>)\s*|\s+|([^\s>]+)/g;
  let childOnly = false;
  let match: RegExpExecArray | null;
  while ((match = tokens.exec(selector.trim())) !== null) {
    if (match[1]) {
      childOnly = true;
      continue;
    }
    const raw = match[2];
    if (!raw) continue;
    const parsed = /^(?:([A-Za-z_*][\w.-]*|\*):)?([\w.-]+|\*)(?:\[([\w:.-]+)=["']?([^\]"']*)["']?\])?$/.exec(raw);
    if (!parsed) throw new Error(`q(): cannot parse selector part '${raw}' (use prefix:name, name, * and [attr=value])`);
    steps.push({ prefix: parsed[1] ?? null, local: parsed[2]!, attr: parsed[3] ? { name: parsed[3], value: parsed[4] ?? "" } : undefined, childOnly });
    childOnly = false;
  }
  return steps;
}

/**
 * `q(parent, "p:sp > p:nvSpPr p:cNvPr[id=4]")`: all matching elements, in document order. Names are prefix:local
 * (a, p, r or a prefix declared on the root) or bare local names; `>` means direct child, a space means descendant.
 */
export function makeQueryHelper(doc: Document) {
  return function q(parent: Node, selector: string): Element[] {
    let current: Element[] = [parent as Element];
    for (const step of parseSelector(selector)) {
      const uri = step.prefix && step.prefix !== "*" ? namespaceFor(doc, step.prefix) : null;
      const next = new Set<Element>();
      for (const base of current) {
        const candidates = step.childOnly ? Array.from(base.children) : Array.from(base.getElementsByTagName("*"));
        for (const candidate of candidates) {
          if (step.local !== "*" && candidate.localName !== step.local) continue;
          if (step.prefix && step.prefix !== "*" && candidate.namespaceURI !== uri) continue;
          if (step.attr && candidate.getAttribute(step.attr.name) !== step.attr.value) continue;
          next.add(candidate);
        }
      }
      current = [...next];
    }
    return current;
  };
}

/* ----- structural validation before insert ---------------------------------------------------------------- */

const SHAPE_CHILDREN = ["sp", "pic", "graphicFrame", "grpSp", "cxnSp", "contentPart"];
/** Child order the schema requires, per element; names not listed are ignored (extLst always last). */
const CHILD_ORDER: Record<string, string[][]> = {
  "p:sld": [["cSld"], ["clrMapOvr"], ["transition"], ["timing"], ["extLst"]],
  "p:cSld": [["bg"], ["spTree"], ["custDataLst"], ["controls"], ["extLst"]],
  "p:spTree": [["nvGrpSpPr"], ["grpSpPr"], SHAPE_CHILDREN, ["extLst"]],
  "p:grpSp": [["nvGrpSpPr"], ["grpSpPr"], SHAPE_CHILDREN, ["extLst"]],
  "p:sp": [["nvSpPr"], ["spPr"], ["style"], ["txBody"], ["extLst"]],
  "p:pic": [["nvPicPr"], ["blipFill"], ["spPr"], ["style"], ["extLst"]],
  "p:cxnSp": [["nvCxnSpPr"], ["spPr"], ["style"], ["extLst"]],
  "p:graphicFrame": [["nvGraphicFramePr"], ["xfrm"], ["graphic"], ["extLst"]],
  "p:txBody": [["bodyPr"], ["lstStyle"], ["p"]],
  "a:p": [["pPr"], ["r", "br", "fld"], ["endParaRPr"]],
  "a:r": [["rPr"], ["t"]],
};

/** Reject XML PowerPoint would repair or refuse: wrong root, schema child order, duplicate shape ids. */
export function validateSlideXml(doc: Document): string[] {
  const problems: string[] = [];
  const root = doc.documentElement;
  if (!root || root.localName !== "sld" || root.namespaceURI !== NS.p) {
    return [`the document root must be <p:sld> in the presentationml namespace, not <${root?.nodeName ?? "nothing"}>`];
  }
  const all = [root, ...Array.from(root.getElementsByTagName("*"))];
  for (const element of all) {
    const prefix = element.namespaceURI === NS.p ? "p" : element.namespaceURI === NS.a ? "a" : null;
    const order = prefix ? CHILD_ORDER[`${prefix}:${element.localName}`] : undefined;
    if (!order) continue;
    let last = -1;
    for (const child of Array.from(element.children)) {
      const rank = order.findIndex((names) => names.includes(child.localName));
      if (rank === -1) continue;
      if (rank < last) {
        problems.push(`<${element.nodeName}> has <${child.nodeName}> out of schema order (expected ${order.map((n) => n.join("|")).join(", ")})`);
        break;
      }
      last = rank;
    }
  }
  const seen = new Set<string>();
  const duplicates = new Set<string>();
  for (const cNvPr of Array.from(root.getElementsByTagNameNS(NS.p, "cNvPr"))) {
    const id = cNvPr.getAttribute("id");
    if (id === null) continue;
    if (seen.has(id)) duplicates.add(id);
    seen.add(id);
  }
  if (duplicates.size) problems.push(`duplicate shape ids ${[...duplicates].join(", ")}; every <p:cNvPr id> must be unique on the slide`);
  return problems.slice(0, 8);
}

/** Number of top-level shapes in a slide's spTree. */
export function countTopLevelShapes(doc: Document): number {
  const tree = doc.getElementsByTagNameNS(NS.p, "spTree")[0];
  return tree ? Array.from(tree.children).filter((child) => child.namespaceURI === NS.p && SHAPE_CHILDREN.includes(child.localName)).length : 0;
}

/* ----- the round trip --------------------------------------------------------------------------------------- */

export interface RoundTripOptions {
  slideId: string;
  /**
   * "replace": the new slide takes the original's position and the original is deleted once the new one is confirmed.
   * "copy": the original stays; the copy lands at `copyPosition` (1-based, final position) or right after the original.
   */
  mode: "replace" | "copy";
  copyPosition?: number;
  /** Change the slide here (parsed XML plus the unpacked file). Return value is passed back to the caller. */
  mutate?: (slideXml: Document, parts: SlideParts) => Promise<unknown> | unknown;
}

export interface RoundTripResult {
  newSlideId: string;
  position: number;
  slideCount: number;
  replacedSlideId?: string;
  shapeCount: number;
  warnings: string[];
  returned?: unknown;
}

/**
 * Export -> (mutate) -> insert -> verify -> (delete original). Never deletes the original unless the replacement is
 * present in the deck; when insertion or verification fails the original stays and the error says so.
 */
export async function roundTripSlide(context: PowerPoint.RequestContext, options: RoundTripOptions): Promise<RoundTripResult> {
  requireSupport("Rebuilding or copying a slide through its .pptx", PPT_EXPORT, "There is no fallback for this on older PowerPoint; ask the person to update Office.");
  const warnings: string[] = [];
  const beforeIds = await loadSlideIds(context);
  const index = beforeIds.indexOf(options.slideId);
  if (index === -1) throw slideNotFound(options.slideId, beforeIds);
  const original = await requireSlide(context, options.slideId);

  const masterIdsBefore = await loadMasterIds(context);

  const exported = original.exportAsBase64();
  await context.sync();
  const parts = unpackSlideFile(exported.value);
  const slideXml = parseXml(strFromU8(parts.files[parts.slidePath]!), parts.slidePath);

  let returned: unknown;
  if (options.mutate) {
    returned = await options.mutate(slideXml, parts);
    const problems = validateSlideXml(slideXml);
    if (problems.length) throw new ToolFailure("invalid_arguments", `the changed slide XML is not valid, nothing was inserted and the slide is unchanged: ${problems.join("; ")}`);
    parts.files[parts.slidePath] = strToU8(serializeXml(slideXml));
  }

  const external = findExternalRelationships(parts);
  if (external.length) {
    throw new ToolFailure(
      "unsupported",
      "this slide has external relationships (hyperlinks or linked media), which PowerPoint refuses when a slide is inserted from a file. " +
        "The slide is unchanged. Remove those links first, or make the change with execute_office_js.",
      { external },
    );
  }

  const expectedShapes = countTopLevelShapes(slideXml);
  let targetSlideId: string | undefined;
  if (options.mode === "replace") {
    targetSlideId = index > 0 ? beforeIds[index - 1] : undefined;
  } else {
    const wanted = options.copyPosition === undefined ? index + 2 : Math.min(options.copyPosition, beforeIds.length + 1);
    targetSlideId = wanted >= 2 ? beforeIds[wanted - 2] : undefined;
  }

  context.presentation.insertSlidesFromBase64(packSlideFile(parts.files), {
    formatting: INSERT_FORMATTING,
    ...(targetSlideId ? { targetSlideId } : {}),
  });
  try {
    await context.sync();
  } catch (error) {
    throw new ToolFailure("office_error", `PowerPoint refused the rebuilt slide, so the original is unchanged: ${error instanceof Error ? error.message : String(error)}`);
  }

  const afterIds = await loadSlideIds(context);
  const known = new Set(beforeIds);
  const added = afterIds.filter((id) => !known.has(id));
  if (added.length !== 1) {
    throw new ToolFailure(
      "office_error",
      added.length === 0
        ? "PowerPoint reported success but no new slide appeared; the original is unchanged."
        : `PowerPoint added ${added.length} slides (${added.join(", ")}) instead of one; the original is unchanged, check the deck and delete the extras.`,
      { added },
    );
  }
  const newSlideId = added[0]!;

  const created = context.presentation.slides.getItem(newSlideId);
  created.shapes.load("items/id");
  await context.sync();
  const shapeCount = created.shapes.items.length;
  if (expectedShapes > 0 && shapeCount === 0) {
    created.delete();
    await context.sync().catch(() => undefined);
    throw new ToolFailure("office_error", "the inserted slide came back empty, so it was removed again and the original is unchanged.");
  }
  if (shapeCount !== expectedShapes) warnings.push(`the new slide has ${shapeCount} shapes where the XML had ${expectedShapes}; check it with list_slide_shapes`);

  if (isSupported(PPT_LAYOUTS)) {
    const masterIdsAfter = await loadMasterIds(context);
    if (masterIdsAfter.length > masterIdsBefore.length) {
      warnings.push(`PowerPoint added ${masterIdsAfter.length - masterIdsBefore.length} extra slide master(s) while inserting the slide; they can be deleted in Slide Master view`);
    }
  }

  let finalIds = afterIds;
  if (options.mode === "replace") {
    original.delete();
    try {
      await context.sync();
    } catch (error) {
      throw new ToolFailure(
        "office_error",
        `the new slide ${newSlideId} was inserted but the original ${options.slideId} could not be deleted (${error instanceof Error ? error.message : String(error)}); delete one of them.`,
        { newSlideId, originalSlideId: options.slideId },
      );
    }
    finalIds = await loadSlideIds(context);
    if (finalIds.includes(options.slideId)) warnings.push(`the original slide ${options.slideId} is still in the deck`);
  }

  const result: RoundTripResult = {
    newSlideId,
    position: finalIds.indexOf(newSlideId) + 1,
    slideCount: finalIds.length,
    shapeCount,
    warnings,
  };
  if (options.mode === "replace") result.replacedSlideId = options.slideId;
  if (options.mutate) result.returned = returned;
  return result;
}

async function loadMasterIds(context: PowerPoint.RequestContext): Promise<string[]> {
  if (!isSupported(PPT_LAYOUTS)) return [];
  const masters = context.presentation.slideMasters;
  masters.load("items/id");
  await context.sync();
  return masters.items.map((master) => master.id);
}
