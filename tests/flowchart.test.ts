import { describe, expect, it } from "vitest";
import type { ShapeInfo } from "../protocol/src/index.js";
import { inspectFlowchartGeometry } from "../addin/taskpane/executor/flowchart.js";

function node(id: string, left: number, top: number, width = 10, height = 10): ShapeInfo {
  return { id, name: id, type: "Rectangle", text: id, left, top, width, height, rotation: 0 };
}
function line(id: string, from: ShapeInfo, fromSite: number, to: ShapeInfo, toSite: number, kind = "Straight"): ShapeInfo {
  return { ...node(id, 0, 0), type: "Line", text: null, connection: {
    from: { id: from.id!, name: from.name, site: fromSite }, to: { id: to.id!, name: to.name, site: toSite },
    connectorType: kind, beginArrow: "None", endArrow: "Triangle",
  } };
}
const a = node("a", 0, 40, 10, 20), b = node("b", 90, 40, 10, 20);
const c = node("c", 40, 0, 20, 10), d = node("d", 40, 90, 20, 10);
const horizontal = line("horizontal", a, 3, b, 1), vertical = line("vertical", c, 2, d, 0);

describe("flowchart geometry", () => {
  it("reports a crossing with connector names and worksheet point coordinates", () => {
    expect(inspectFlowchartGeometry([a, b, c, d, horizontal, vertical])).toMatchObject({
      complete: true, checkedConnectors: 2, issues: [{ code: "connector_crossing", shapes: [{ name: "horizontal" }, { name: "vertical" }], point: { x: 50, y: 50 } }],
    });
  });

  it("detects collinear shared length even with opposite direction", () => {
    const reversed = line("reversed", b, 1, a, 3);
    expect(inspectFlowchartGeometry([a, b, horizontal, reversed]).issues).toMatchObject([{ code: "connector_overlap" }]);
  });

  it("allows connectors to share a node attachment point", () => {
    const above = node("above", 0, 0, 10, 10);
    const diagonal = line("diagonal", a, 3, above, 3);
    expect(inspectFlowchartGeometry([a, b, above, horizontal, diagonal]).issues).toEqual([]);
  });

  it("distinguishes parallel lines from collinear overlaps", () => {
    const e = node("e", 0, 60, 10, 20), f = node("f", 90, 60, 10, 20);
    expect(inspectFlowchartGeometry([a, b, e, f, horizontal, line("parallel", e, 3, f, 1)]).issues).toEqual([]);
  });

  it("detects a connector through an unrelated node but excludes its own endpoints", () => {
    const obstruction = node("obstruction", 45, 45);
    expect(inspectFlowchartGeometry([a, b, horizontal, obstruction]).issues).toEqual([
      { code: "connector_through_node_bounds", shapes: [{ id: "horizontal", name: "horizontal" }, { id: "obstruction", name: "obstruction" }] },
    ]);
  });

  it("detects overlapping node bounds without treating mere touching as overlap", () => {
    expect(inspectFlowchartGeometry([node("x", 0, 0), node("y", 5, 5)]).issues[0]?.code).toBe("node_bounds_overlap");
    expect(inspectFlowchartGeometry([node("x", 0, 0), node("y", 10, 0)]).issues).toEqual([]);
  });

  it.each(["Elbow", "Curve"])("marks %s routes as unverified instead of inventing bends", kind => {
    expect(inspectFlowchartGeometry([a, b, line("routed", a, 3, b, 1, kind)])).toMatchObject({ complete: false, checkedConnectors: 0, issues: [], unverified: [{ name: "routed" }] });
  });

  it("does not assert geometry for rotated/grouped nodes or unknown sites", () => {
    for (const source of [{ ...a, rotation: 30 }, { ...a, parentGroupId: "group" }, { ...a, type: "Hexagon" }]) {
      expect(inspectFlowchartGeometry([source, b, horizontal])).toMatchObject({ complete: false, checkedConnectors: 0 });
    }
    expect(inspectFlowchartGeometry([a, b, line("bad-site", a, 99, b, 1)])).toMatchObject({ complete: false });
  });

  it("bounds the report while preserving the actual issue count", () => {
    const nodes = Array.from({ length: 25 }, (_, i) => node(`n${i}`, 0, 0));
    expect(inspectFlowchartGeometry(nodes)).toMatchObject({ issueCount: 300, issuesTruncated: true, issues: expect.any(Array) });
    expect(inspectFlowchartGeometry(nodes).issues).toHaveLength(200);
  });
});
