import { type ShapeInfo, type ValidateFlowchartInput } from "../../../protocol/src/index.js";
import { listShapes } from "./shapes.js";

type Point = { x: number; y: number };
type Segment = { shape: ShapeInfo; a: Point; b: Point };
const EPS = 0.01;
const cross = (a: Point, b: Point) => a.x * b.y - a.y * b.x;
const sub = (a: Point, b: Point): Point => ({ x: a.x - b.x, y: a.y - b.y });
const at = (a: Point, d: Point, t: number): Point => ({ x: a.x + d.x * t, y: a.y + d.y * t });
const length = (p: Point) => Math.hypot(p.x, p.y);

/** Ignore shared endpoints, but retain collinear segments sharing actual length. */
export function segmentConflict(a: Segment, b: Segment): { code: string; point: Point } | null {
  const r = sub(a.b, a.a), s = sub(b.b, b.a), delta = sub(b.a, a.a);
  const lr = length(r), ls = length(s);
  if (lr < EPS || ls < EPS) return null;
  const determinant = cross(r, s);
  if (Math.abs(determinant) < EPS * Math.min(lr, ls)) {
    if (Math.abs(cross(delta, r)) > EPS * lr) return null;
    const dot = (v: Point) => (v.x * r.x + v.y * r.y) / (lr * lr);
    const t0 = dot(delta), t1 = dot(sub(b.b, a.a));
    const lo = Math.max(0, Math.min(t0, t1)), hi = Math.min(1, Math.max(t0, t1));
    return (hi - lo) * lr > EPS ? { code: "connector_overlap", point: at(a.a, r, (lo + hi) / 2) } : null;
  }
  const t = cross(delta, s) / determinant, u = cross(delta, r) / determinant;
  if (t < -EPS / lr || t > 1 + EPS / lr || u < -EPS / ls || u > 1 + EPS / ls) return null;
  const endA = t * lr <= EPS || (1 - t) * lr <= EPS;
  const endB = u * ls <= EPS || (1 - u) * ls <= EPS;
  if (endA && endB) return null;
  return { code: "connector_crossing", point: at(a.a, r, t) };
}

// Cardinal connection sites with known positions. Other sites are explicitly unverified.
function connectionPoint(shape: ShapeInfo, site: number): Point | null {
  if ((shape.rotation ?? 0) % 360 !== 0) return null;
  const four = ["Rectangle", "RoundRectangle", "Diamond", "FlowChartProcess", "FlowChartDecision", "FlowChartAlternateProcess", "FlowChartPredefinedProcess", "FlowChartTerminator"];
  const eight = ["Ellipse", "FlowChartConnector", "FlowChartOr", "FlowChartSummingJunction"];
  const index = four.includes(shape.type) ? site : eight.includes(shape.type) && site % 2 === 0 ? site / 2 : -1;
  return [
    { x: shape.left + shape.width / 2, y: shape.top },
    { x: shape.left, y: shape.top + shape.height / 2 },
    { x: shape.left + shape.width / 2, y: shape.top + shape.height },
    { x: shape.left + shape.width, y: shape.top + shape.height / 2 },
  ][index] ?? null;
}

function boundsOverlap(a: ShapeInfo, b: ShapeInfo): boolean {
  return Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left) > EPS &&
    Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top) > EPS;
}

/** Clip a segment against the interior of a node's bounding rectangle. */
function crossesBounds(segment: Segment, node: ShapeInfo): boolean {
  let lo = 0, hi = 1;
  for (const [origin, delta, min, max] of [
    [segment.a.x, segment.b.x - segment.a.x, node.left + EPS, node.left + node.width - EPS],
    [segment.a.y, segment.b.y - segment.a.y, node.top + EPS, node.top + node.height - EPS],
  ] as [number, number, number, number][]) {
    if (min > max) return false;
    if (Math.abs(delta) < EPS) { if (origin <= min || origin >= max) return false; }
    else {
      const t0 = (min - origin) / delta, t1 = (max - origin) / delta;
      lo = Math.max(lo, Math.min(t0, t1)); hi = Math.min(hi, Math.max(t0, t1));
      if (lo >= hi) return false;
    }
  }
  return hi > lo;
}

export function inspectFlowchartGeometry(shapes: ShapeInfo[]) {
  const byId = new Map(shapes.map(s => [s.id, s]));
  const nodes = shapes.filter(s => s.type !== "Line" && s.type !== "Group" && s.type !== "TextBox" && s.type !== "Image");
  const segments: Segment[] = [];
  const unverified: { id?: string; name: string; reason: string }[] = [];
  const issues: { code: string; shapes: { id?: string; name: string }[]; point?: Point }[] = [];
  const ref = (shape: ShapeInfo) => ({ id: shape.id, name: shape.name });
  let issueCount = 0;
  const record = (issue: typeof issues[number]) => { issueCount++; if (issues.length < 200) issues.push(issue); };
  for (const shape of shapes.filter(s => s.type === "Line")) {
    const c = shape.connection;
    const from = c?.from ? byId.get(c.from.id) : undefined;
    const to = c?.to ? byId.get(c.to.id) : undefined;
    const a = from && c?.from ? connectionPoint(from, c.from.site) : null;
    const b = to && c?.to ? connectionPoint(to, c.to.site) : null;
    if (!c || c.connectorType !== "Straight" || !a || !b || (shape.rotation ?? 0) % 360 !== 0 || shape.parentGroupId || from?.parentGroupId || to?.parentGroupId) {
      unverified.push({ ...ref(shape), reason: "Full route unavailable: requires an unrotated straight connector with known attached sites outside groups. Review the rendered diagram." });
    } else if (length(sub(b, a)) < EPS) {
      record({ code: "zero_length_connector", shapes: [ref(shape)] });
    } else segments.push({ shape, a, b });
  }
  for (let i = 0; i < nodes.length; i++) {
    for (const b of nodes.slice(i + 1)) {
      const a = nodes[i]!;
      if (!a.parentGroupId && !b.parentGroupId && !(a.rotation ?? 0) && !(b.rotation ?? 0) && boundsOverlap(a, b)) {
        record({ code: "node_bounds_overlap", shapes: [ref(a), ref(b)] });
      }
    }
  }
  for (let i = 0; i < segments.length; i++) {
    const a = segments[i]!;
    for (const b of segments.slice(i + 1)) {
      const conflict = segmentConflict(a, b);
      if (conflict) record({ ...conflict, shapes: [ref(a.shape), ref(b.shape)] });
    }
    for (const node of nodes) {
      if (node.id === a.shape.connection!.from!.id || node.id === a.shape.connection!.to!.id || node.parentGroupId || (node.rotation ?? 0)) continue;
      if (crossesBounds(a, node)) record({ code: "connector_through_node_bounds", shapes: [ref(a.shape), ref(node)] });
    }
  }
  return {
    model: "Straight segments reconstructed from attached cardinal connection sites; node checks use bounding rectangles, not shape outlines.",
    checkedConnectors: segments.length,
    unverified,
    unverifiedNodes: nodes.filter(n => n.parentGroupId || (n.rotation ?? 0)).map(n => ({
      ...ref(n), reason: "Grouped or rotated node bounds require visual review.",
    })),
    complete: unverified.length === 0 && !nodes.some(n => n.parentGroupId || (n.rotation ?? 0)),
    issueCount,
    issuesTruncated: issueCount > issues.length,
    issues,
  };
}

export async function validateFlowchart(input: ValidateFlowchartInput): Promise<unknown> {
  const result = await listShapes({ sheet: input.sheet, offset: 0, limit: input.maxShapes ?? 1000 }) as {
    sheet: string; count: number; hasMore: boolean; nextOffset: number | null;
    readWarnings: unknown[]; textReadComplete: boolean;
    shapes: ShapeInfo[]; validation: { complete: boolean; issues: unknown[] };
  };
  const geometry = inspectFlowchartGeometry(result.shapes);
  return {
    sheet: result.sheet,
    shapeCount: result.count,
    scanned: result.shapes.length,
    scanTruncated: result.hasMore,
    nextOffset: result.nextOffset,
    status: geometry.issueCount || result.validation.issues.length ? "issues_found"
      : result.hasMore || !geometry.complete || !result.shapes.length ? "needs_visual_review" : "no_issues_in_checked_geometry",
    structural: result.validation,
    geometry,
    readWarnings: result.readWarnings,
    textReadComplete: result.textReadComplete,
    note: "Read-only inspection; no automatic rerouting. Text read failures do not prevent connection/geometry checks; do not delete shapes to work around read warnings. Review unverified routes visually. A clean geometric result does not validate business logic. After moving shapes or recreating connectors, run this tool again.",
  };
}
