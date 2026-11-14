import type { ExportSlideImageInput, VerifySlideVisualInput, VerifySlidesInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { TEXT_TYPES, loadShapeInfos, type ShapeInfo } from "./model.js";
import { getSlideSize, type SlideSize } from "./ooxml.js";
import { MAX_IMAGE_BASE64, PPT_EXPORT, PPT_GEOMETRY, isSupported, loadSlideIds, requireSlide, requireSupport, round, slideNotFound } from "./util.js";

export interface VerifyIssue {
  slide_id: string;
  severity: "error" | "warning" | "info";
  code: string;
  message: string;
  shape_ref?: string;
}

/** Slides checked when verify_slides gets no slide_ids. */
const MAX_SLIDES = 100;
const MAX_ISSUES = 200;
const TOLERANCE = 1; // points of slack before a shape counts as outside the slide
const MIN_FONT = 10;
const SKIPPED_PLACEHOLDERS = new Set(["Date", "SlideNumber", "Footer", "Header"]);

/* ----- pure analysis (testable without PowerPoint) ------------------------------------------------------------ */

function hexToRgb(hex: string | null | undefined): [number, number, number] | null {
  const match = /^#?([0-9a-f]{6})$/i.exec((hex ?? "").trim());
  if (!match) return null;
  const n = parseInt(match[1]!, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function luminance([r, g, b]: [number, number, number]): number {
  const channel = (v: number) => {
    const s = v / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}

/** WCAG contrast ratio of two #RRGGBB colours, or null when either is unknown. */
export function contrastRatio(a: string | null | undefined, b: string | null | undefined): number | null {
  const x = hexToRgb(a);
  const y = hexToRgb(b);
  if (!x || !y) return null;
  const [hi, lo] = [luminance(x), luminance(y)].sort((p, q) => q - p) as [number, number];
  return (hi + 0.05) / (lo + 0.05);
}

const isCjk = (ch: string) => /[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/.test(ch);

/** Rough height (points) the text needs: width per character ~0.5 em (1 em for CJK), 1.2 line height, plus margins. */
export function estimateTextHeight(text: string, fontSize: number, boxWidth: number, margins = { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 }, wrap = true): number {
  const usable = Math.max(1, boxWidth - margins.left - margins.right);
  let lines = 0;
  for (const paragraph of text.split("\n")) {
    if (!wrap) {
      lines += 1;
      continue;
    }
    let width = 0;
    for (const ch of paragraph) width += isCjk(ch) ? fontSize : fontSize * 0.5;
    lines += Math.max(1, Math.ceil(width / usable));
  }
  return lines * fontSize * 1.2 + margins.top + margins.bottom;
}

interface Rect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

const rectOf = (info: ShapeInfo): Rect | null =>
  info.left === undefined ? null : { left: info.left, top: info.top!, right: info.left + info.width!, bottom: info.top! + info.height! };

const area = (r: Rect) => Math.max(0, r.right - r.left) * Math.max(0, r.bottom - r.top);

function intersection(a: Rect, b: Rect): number {
  return area({ left: Math.max(a.left, b.left), top: Math.max(a.top, b.top), right: Math.min(a.right, b.right), bottom: Math.min(a.bottom, b.bottom) });
}

const contains = (outer: Rect, inner: Rect, slack = 2) =>
  outer.left <= inner.left + slack && outer.top <= inner.top + slack && outer.right >= inner.right - slack && outer.bottom >= inner.bottom - slack;

/** Geometry / text checks for one slide. `size` null skips the off-slide and background rules. */
export function analyzeSlide(slideId: string, shapes: ShapeInfo[], size: SlideSize | null): VerifyIssue[] {
  const issues: VerifyIssue[] = [];
  const add = (severity: VerifyIssue["severity"], code: string, message: string, info?: ShapeInfo) =>
    issues.push({ slide_id: slideId, severity, code, message, ...(info ? { shape_ref: `${slideId}:${info.id}` } : {}) });
  const slideArea = size ? size.width * size.height : 0;
  const label = (info: ShapeInfo) => `"${info.name}"`;

  for (const info of shapes) {
    const rect = rectOf(info);

    // Off the slide.
    if (rect && size) {
      const outside = area(rect) - intersection(rect, { left: 0, top: 0, right: size.width, bottom: size.height });
      const overshoot = Math.max(-rect.left, -rect.top, rect.right - size.width, rect.bottom - size.height);
      if (area(rect) > 0 && intersection(rect, { left: 0, top: 0, right: size.width, bottom: size.height }) === 0) {
        add("error", "off_slide", `${label(info)} is completely outside the slide`, info);
      } else if (overshoot > TOLERANCE) {
        const share = area(rect) > 0 ? outside / area(rect) : 0;
        add(share > 0.5 ? "error" : "warning", "off_slide", `${label(info)} sticks out of the slide by ${round(overshoot)} pt (${Math.round(share * 100)}% of it is outside)`, info);
      }
    }

    // Empty placeholders leave "Click to add text" boxes on a presented slide.
    if (info.type === "Placeholder" && !SKIPPED_PLACEHOLDERS.has(info.placeholder ?? "")) {
      const holdsContent = info.contained !== undefined && info.contained !== null;
      if (!holdsContent && info.text !== undefined && info.text.trim() === "") {
        add("warning", "empty_placeholder", `${label(info)} is an empty ${info.placeholder ?? "placeholder"} placeholder; fill it or delete it`, info);
      }
    }

    const hasText = TEXT_TYPES.has(info.type) && info.text !== undefined && info.text.trim() !== "";
    if (!hasText) continue;

    // Small fonts.
    if (typeof info.fontSize === "number" && info.fontSize > 0 && info.fontSize < MIN_FONT) {
      add("warning", "tiny_font", `${label(info)} uses ${info.fontSize} pt text, below ${MIN_FONT} pt`, info);
    }

    // Text that probably does not fit. Shapes that resize or shrink their text are fine.
    if (rect && typeof info.fontSize === "number" && info.fontSize > 0 && info.autoSize !== "AutoSizeShapeToFitText" && info.autoSize !== "AutoSizeTextToFitShape") {
      const need = estimateTextHeight(info.text!, info.fontSize, info.width!, info.margins, info.wordWrap !== false);
      if (need > info.height! * 1.05) {
        add("warning", "text_overflow_maybe", `${label(info)} may overflow: the text needs about ${round(need, 0)} pt of height at ${info.fontSize} pt but the shape is ${round(info.height!, 0)} pt tall (estimate)`, info);
      }
    }

    // Contrast between the shape's solid fill and its text colour, when both are known.
    if (info.fillType === "Solid") {
      const ratio = contrastRatio(info.fontColor, info.fillColor);
      const large = typeof info.fontSize === "number" && info.fontSize >= 18;
      if (ratio !== null && ratio < (large ? 3 : 4.5)) {
        add("warning", "low_contrast", `${label(info)} has low contrast (${round(ratio, 1)}:1) between text ${info.fontColor} and fill ${info.fillColor}`, info);
      }
    }
  }

  // Overlapping text shapes. Full-bleed backgrounds and containers (one shape lying inside the other) are intentional.
  const texty = shapes.filter((info) => TEXT_TYPES.has(info.type) && info.text !== undefined && info.text.trim() !== "" && rectOf(info));
  for (let i = 0; i < texty.length; i++) {
    for (let j = i + 1; j < texty.length; j++) {
      const a = texty[i]!;
      const b = texty[j]!;
      const ra = rectOf(a)!;
      const rb = rectOf(b)!;
      if (slideArea > 0 && (area(ra) >= slideArea * 0.9 || area(rb) >= slideArea * 0.9)) continue;
      if (contains(ra, rb) || contains(rb, ra)) continue;
      const shared = intersection(ra, rb);
      const smaller = Math.min(area(ra), area(rb));
      if (smaller > 0 && shared / smaller > 0.1) {
        add("warning", "overlap", `${label(a)} and ${label(b)} overlap (${Math.round((shared / smaller) * 100)}% of the smaller one)`, a);
      }
    }
  }
  return issues;
}

/* ----- tools ---------------------------------------------------------------------------------------------------- */

async function analyzeSlides(context: PowerPoint.RequestContext, slideIds: string[], size: SlideSize | null): Promise<{ issues: VerifyIssue[]; shapes: number }> {
  const issues: VerifyIssue[] = [];
  let shapes = 0;
  for (const slideId of slideIds) {
    const slide = context.presentation.slides.getItem(slideId);
    const infos = await loadShapeInfos(context, slide.shapes, { geometry: true, text: true, style: true, placeholders: true });
    shapes += infos.length;
    issues.push(...analyzeSlide(slideId, infos, size));
  }
  return { issues, shapes };
}

const order = { error: 0, warning: 1, info: 2 } as const;

export async function verifySlides(input: VerifySlidesInput): Promise<unknown> {
  requireSupport("verify_slides", PPT_GEOMETRY, "Checking shapes needs PowerPointApi 1.4.");
  return PowerPoint.run(async (context) => {
    const all = await loadSlideIds(context);
    const wanted = input.slide_ids ? [...new Set(input.slide_ids)] : all.slice(0, MAX_SLIDES);
    for (const id of wanted) if (!all.includes(id)) throw slideNotFound(id, all);
    const size = await getSlideSize(context, all[0]);
    const { issues, shapes } = await analyzeSlides(context, wanted, size);
    issues.sort((a, b) => order[a.severity] - order[b.severity]);
    const notes: string[] = [];
    if (!size) notes.push("The slide size is unknown on this PowerPoint version, so off-slide and background checks were skipped.");
    if (!input.slide_ids && all.length > MAX_SLIDES) notes.push(`Only the first ${MAX_SLIDES} of ${all.length} slides were checked; pass slide_ids for the others.`);
    notes.push("Geometry checks are estimates; use verify_slide_visual to look at the rendered slide.");
    const shown = issues.slice(0, MAX_ISSUES);
    return {
      summary: issues.length === 0 ? `Checked ${wanted.length} slide(s): no issues found` : `Checked ${wanted.length} slide(s): ${issues.length} issue(s)`,
      issues: shown,
      issues_truncated: issues.length > shown.length,
      checked: { slides: wanted.length, shapes },
      slide_size: size ? { width: round(size.width), height: round(size.height), unit: "pt" } : null,
      notes,
    };
  });
}

async function renderSlide(context: PowerPoint.RequestContext, slideId: string, width: number): Promise<{ mimeType: "image/png"; data: string }> {
  requireSupport("Rendering a slide image", PPT_EXPORT, "Slide images need PowerPointApi 1.8; without it ask the person to look at the slide.");
  await requireSlide(context, slideId);
  const image = context.presentation.slides.getItem(slideId).getImageAsBase64({ width });
  await context.sync();
  if (!image.value) throw new ToolFailure("office_error", "PowerPoint returned an empty image");
  if (image.value.length > MAX_IMAGE_BASE64) {
    throw new ToolFailure("size_limit_exceeded", `the rendered image is too large (${image.value.length} base64 characters); ask for a smaller width`);
  }
  return { mimeType: "image/png", data: image.value };
}

export async function exportSlideImage(input: ExportSlideImageInput): Promise<unknown> {
  return PowerPoint.run(async (context) => {
    const width = input.width ?? 960;
    const image = await renderSlide(context, input.slide_id, width);
    return { summary: `Rendered slide ${input.slide_id} at ${width}px wide`, slide_id: input.slide_id, width, image };
  });
}

/** Rendered image plus the geometry issues of the same slide: the "look at it" half of verification. */
export async function verifySlideVisual(input: VerifySlideVisualInput): Promise<unknown> {
  requireSupport("verify_slide_visual", PPT_EXPORT, "Slide images need PowerPointApi 1.8; use verify_slides for the geometry checks.");
  return PowerPoint.run(async (context) => {
    const width = input.width ?? 960;
    const image = await renderSlide(context, input.slide_id, width);
    let issues: VerifyIssue[] = [];
    let note = "Look at the image for overlapping, clipped or unbalanced content; the issues list is a geometry estimate.";
    if (isSupported(PPT_GEOMETRY)) {
      const size = await getSlideSize(context, input.slide_id);
      issues = (await analyzeSlides(context, [input.slide_id], size)).issues.sort((a, b) => order[a.severity] - order[b.severity]).slice(0, MAX_ISSUES);
    } else note = "Shape geometry needs PowerPointApi 1.4, so only the image is returned.";
    return {
      summary: `Rendered slide ${input.slide_id}: ${issues.length} geometry issue(s)`,
      slide_id: input.slide_id,
      width,
      issues,
      note,
      image,
    };
  });
}
