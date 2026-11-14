import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  base64ToBytes,
  bytesToBase64,
  findExternalRelationships,
  makeElementHelper,
  makeQueryHelper,
  parseXml,
  unpackSlideFile,
  validateSlideXml,
} from "../frontend/taskpane/executor/powerpoint/ooxml.js";
import { call, fails, ok, useFakeDeck } from "./support/ppt-fixture.js";

const fx = useFakeDeck();
const unchanged = (before: string[]) => expect(fx().deck.ids()).toEqual(before);

describe("duplicate_slide", () => {
  it("inserts the copy right after the original and keeps the original", async () => {
    const { deck, title, agenda, data } = fx();
    const result = await ok("duplicate_slide", { slide_id: agenda });
    expect(result).toMatchObject({ source_slide_id: agenda, position: 3, slide_count: 4, shape_count: 3 });
    expect(result.slide_id).not.toBe(agenda);
    expect(deck.ids()).toEqual([title, agenda, result.slide_id, data]);
    const copy = deck.slideById(result.slide_id);
    expect(copy.shapes.map((s) => [s.id, s.name, s.type])).toEqual([["2", "Title 1", "Placeholder"], ["3", "Content Placeholder 2", "Placeholder"], ["4", "Logo", "Image"]]);
    expect(copy.shapes[1]!.text).toBe("Revenue\nCosts\nOutlook");
    expect(copy.layoutId).toBe("l2#");
    expect(deck.log.filter((l) => l.startsWith("slide.delete"))).toEqual([]);
    expect(deck.violations).toEqual([]);
  });

  it("copies the first slide to the very start and to explicit positions", async () => {
    const { deck, title, agenda, data } = fx();
    const first = await ok("duplicate_slide", { slide_id: agenda, position: 1 });
    expect(first.position).toBe(1);
    expect(deck.ids()).toEqual([first.slide_id, title, agenda, data]);
    const last = await ok("duplicate_slide", { slide_id: title, position: 99 });
    expect(last.position).toBe(5);
    expect(deck.ids()[4]).toBe(last.slide_id);
    const second = await ok("duplicate_slide", { slide_id: data, position: 2 });
    expect(second.position).toBe(2);
    const afterFirst = await ok("duplicate_slide", { slide_id: deck.ids()[0]! });
    expect(afterFirst.position).toBe(2);
  });

  it("refuses slides with external relationships before inserting anything", async () => {
    const { deck } = fx();
    const linked = deck.addSlideWith([{ name: "Link", text: "see site" }], { externalLinks: ["https://example.com/page"] });
    const before = deck.ids();
    expect(await fails("duplicate_slide", { slide_id: linked }, "unsupported")).toMatch(/external relationships/);
    expect(deck.insertCalls).toBe(0);
    unchanged(before);
  });

  it("warns when PowerPoint adds an extra master", async () => {
    fx().deck.duplicateMasterOnInsert = true;
    const result = await ok("duplicate_slide", { slide_id: fx().agenda });
    expect(result.warnings[0]).toMatch(/extra slide master/);
  });

  it("is unsupported below 1.8 and for unknown slides", async () => {
    expect(await fails("duplicate_slide", { slide_id: "nope" }, "invalid_arguments")).toMatch(/does not exist/);
    fx().deck.api = 7;
    expect(await fails("duplicate_slide", { slide_id: fx().agenda }, "unsupported")).toMatch(/1\.8/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("leaves the deck alone when PowerPoint refuses the insert", async () => {
    const before = fx().deck.ids();
    fx().deck.failInsert = true;
    expect(await fails("duplicate_slide", { slide_id: fx().agenda }, "office_error")).toMatch(/original is unchanged/);
    unchanged(before);
  });
});

describe("edit_slide_xml", () => {
  const code = (body: string) => ({ slide_id: fx().agenda, code: body });

  it("rebuilds the slide in place: inserted first, original deleted last, same position", async () => {
    const { deck, title, agenda, data } = fx();
    const result = await ok("edit_slide_xml", {
      slide_id: agenda,
      description: "retitle",
      code: `
        const t = helpers.q(xml.documentElement, "p:sp a:t")[0];
        t.textContent = "New agenda";
        return { files: files.length, hasSlide: files.includes("ppt/slides/slide1.xml"), ns: Object.keys(ns) };`,
    });
    expect(result).toMatchObject({ replaced_slide_id: agenda, position: 2, slide_count: 3, shape_count: 3, result: { hasSlide: true, ns: ["a", "p", "r"] } });
    expect(result.summary).toMatch(/retitle/);
    expect(deck.ids()).toEqual([title, result.slide_id, data]);
    const log = deck.log.filter((l) => /^insert|slide\.delete/.test(l));
    expect(log).toHaveLength(2);
    expect(log[0]).toMatch(/^insert/);
    expect(log[1]).toBe(`slide.delete ${agenda}`);
    const edited = deck.slideById(result.slide_id);
    expect(edited.shapes[0]!.text).toBe("New agenda");
    expect(edited.shapes[2]).toMatchObject({ name: "Logo", type: "Image" });
    expect(deck.violations).toEqual([]);
  });

  it("works for the first and the last slide", async () => {
    const { deck, title, agenda, data } = fx();
    const first = await ok("edit_slide_xml", { slide_id: title, code: `helpers.q(xml.documentElement, "a:t")[0].textContent = "Annual Review";` });
    expect(first.position).toBe(1);
    expect(deck.slideById(first.slide_id).shapes[0]!.text).toBe("Annual Review");
    expect(deck.ids()).toEqual([first.slide_id, agenda, data]);
    const last = await ok("edit_slide_xml", { slide_id: data, code: `return "noop";` });
    expect(last.position).toBe(3);
    expect(last.result).toBe("noop");
    expect(deck.ids()).toEqual([first.slide_id, agenda, last.slide_id]);
  });

  it("keeps run formatting that the typed text edit would lose", async () => {
    const { deck, title } = fx();
    const result = await ok("edit_slide_xml", { slide_id: title, code: `helpers.q(xml.documentElement, "a:t")[0].textContent = "Kept";` });
    expect(deck.slideById(result.slide_id).shapes[0]).toMatchObject({ text: "Kept", fontSize: 40 });
  });

  it("builds new elements with helpers.el in schema order", async () => {
    const { deck, agenda } = fx();
    const result = await ok("edit_slide_xml", {
      slide_id: agenda,
      code: `
        const tree = helpers.q(xml.documentElement, "p:spTree")[0];
        const E = helpers.el;
        tree.appendChild(E("p:sp", {}, 
          E("p:nvSpPr", {}, E("p:cNvPr", { id: 20, name: "Badge" }), E("p:cNvSpPr"), E("p:nvPr")),
          E("p:spPr", {}, E("a:xfrm", {}, E("a:off", { x: 635000, y: 635000 }), E("a:ext", { cx: 1270000, cy: 635000 }))),
          E("p:txBody", {}, E("a:bodyPr"), E("a:lstStyle"), E("a:p", {}, E("a:r", {}, E("a:t", {}, "NEW"))))));
        return "added";`,
    });
    expect(result.shape_count).toBe(4);
    expect(deck.slideById(result.slide_id).shapes.find((s) => s.name === "Badge")).toMatchObject({ id: "20", text: "NEW", left: 50, top: 50, width: 100, height: 50 });
  });

  it("rejects XML with children out of schema order and leaves the original", async () => {
    const { deck } = fx();
    const before = deck.ids();
    const message = await fails("edit_slide_xml", code(`
      const sp = helpers.q(xml.documentElement, "p:sp")[0];
      const spPr = helpers.q(sp, "p:spPr")[0];
      sp.removeChild(spPr);
      sp.appendChild(spPr);`), "invalid_arguments");
    expect(message).toMatch(/out of schema order/);
    expect(deck.insertCalls).toBe(0);
    unchanged(before);
  });

  it("rejects duplicate shape ids", async () => {
    const before = fx().deck.ids();
    const message = await fails("edit_slide_xml", code(`
      const ids = helpers.q(xml.documentElement, "p:cNvPr");
      ids[2].setAttribute("id", ids[1].getAttribute("id"));`), "invalid_arguments");
    expect(message).toMatch(/duplicate shape ids/);
    unchanged(before);
  });

  it("leaves the original when the code throws or does not compile", async () => {
    const before = fx().deck.ids();
    expect(await fails("edit_slide_xml", code(`throw new Error("boom");`), "invalid_arguments")).toMatch(/boom.*unchanged/);
    expect(await fails("edit_slide_xml", code(`this is not js`), "invalid_arguments")).toMatch(/does not compile/);
    expect(fx().deck.insertCalls).toBe(0);
    unchanged(before);
  });

  it("refuses a slide with external links and says so", async () => {
    const linked = fx().deck.addSlideWith([{ name: "Link", text: "x" }], { externalLinks: ["https://example.com"] });
    const before = fx().deck.ids();
    const message = await fails("edit_slide_xml", { slide_id: linked, code: "return 1;" }, "unsupported");
    expect(message).toMatch(/external relationships/);
    unchanged(before);
  });

  it("keeps the original when PowerPoint refuses the insert or returns an empty slide", async () => {
    const { deck } = fx();
    const before = deck.ids();
    deck.failInsert = true;
    expect(await fails("edit_slide_xml", code("return 1;"), "office_error")).toMatch(/original is unchanged/);
    unchanged(before);
    deck.failInsert = false;
    deck.dropShapesOnInsert = true;
    expect(await fails("edit_slide_xml", code("return 1;"), "office_error")).toMatch(/came back empty/);
    unchanged(before);
    expect(deck.log.filter((l) => l.startsWith("slide.delete")).length).toBe(1);
    expect(deck.log.at(-1)).not.toBe(`slide.delete ${fx().agenda}`);
  });

  it("reports both slides when the original cannot be deleted afterwards", async () => {
    const { deck, agenda } = fx();
    deck.failSlideDelete = true;
    const message = await fails("edit_slide_xml", code("return 1;"), "office_error");
    expect(message).toMatch(new RegExp(`original ${agenda}`));
    expect(deck.ids()).toHaveLength(4);
  });

  it("caps what the code returns", async () => {
    const result = await ok("edit_slide_xml", code(`return "z".repeat(40000);`));
    expect(result.result).toMatchObject({ truncated: true, characters: 40002 });
    expect(JSON.stringify(result).length).toBeLessThan(30_000);
  });

  it("is unsupported below 1.8", async () => {
    fx().deck.api = 7;
    expect(await fails("edit_slide_xml", code("return 1;"), "unsupported")).toMatch(/1\.8/);
    expect(fx().deck.violations).toEqual([]);
  });

  it("returns an error for a slide that is gone", async () => {
    await fails("edit_slide_xml", { slide_id: "ghost", code: "return 1;" }, "invalid_arguments");
  });
});

describe("OOXML helpers", () => {
  const doc = () =>
    parseXml(
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="2" name="A"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>one</a:t></a:r></a:p></p:txBody></p:sp>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="3" name="B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>` +
        `</p:spTree></p:cSld></p:sld>`,
      "test",
    );

  it("q: descendant, child, attribute and wildcard selectors", () => {
    const d = doc();
    const q = makeQueryHelper(d);
    expect(q(d.documentElement, "p:sp")).toHaveLength(2);
    expect(q(d.documentElement, "p:sp > p:nvSpPr > p:cNvPr").map((e) => e.getAttribute("name"))).toEqual(["A", "B"]);
    expect(q(d.documentElement, "p:cNvPr[id=3]")[0]!.getAttribute("name")).toBe("B");
    expect(q(d.documentElement, "a:t")[0]!.textContent).toBe("one");
    expect(q(d.documentElement, "p:spTree > p:cNvPr")).toHaveLength(0);
    expect(q(d.documentElement, "*:t")).toHaveLength(1);
    expect(() => q(d.documentElement, "p:sp[")).toThrow(/cannot parse/);
  });

  it("el: namespaced elements, attributes, text and imported children", () => {
    const d = doc();
    const el = makeElementHelper(d);
    const node = el("a:r", { lang: "en-US", n: 3, skip: null }, el("a:t", null, "hi"), "tail");
    expect(node.namespaceURI).toBe("http://schemas.openxmlformats.org/drawingml/2006/main");
    expect(node.getAttribute("n")).toBe("3");
    expect(node.hasAttribute("skip")).toBe(false);
    expect(node.textContent).toBe("hitail");
    expect(() => el("zz:x")).toThrow(/unknown namespace prefix/);
  });

  it("validateSlideXml: wrong root, order and duplicates", () => {
    expect(validateSlideXml(doc())).toEqual([]);
    expect(validateSlideXml(parseXml("<foo/>", "x"))[0]).toMatch(/root must be <p:sld>/);
    const bad = doc();
    const sp = bad.getElementsByTagName("p:sp")[0]!;
    sp.insertBefore(sp.lastElementChild!, sp.firstElementChild);
    expect(validateSlideXml(bad)[0]).toMatch(/out of schema order/);
  });

  it("findExternalRelationships reads the slide rels", () => {
    const rels = `<Relationships><Relationship Id="1" Type="x/hyperlink" Target="https://a.b/c" TargetMode="External"/><Relationship Id="2" Type="x/slideLayout" Target="../l.xml"/></Relationships>`;
    const parts = { files: { "r.rels": strToU8(rels) } as Record<string, Uint8Array>, slidePath: "s.xml", relsPath: "r.rels" };
    expect(findExternalRelationships(parts)).toEqual([{ type: "hyperlink", target: "https://a.b/c" }]);
    expect(findExternalRelationships({ ...parts, relsPath: "missing" })).toEqual([]);
  });

  it("base64 round trip and unpackSlideFile", () => {
    const bytes = new Uint8Array(70_000).map((_, i) => i % 251);
    expect(bytesToBase64(bytes)).toBe(Buffer.from(bytes).toString("base64"));
    expect([...base64ToBytes(bytesToBase64(bytes))]).toEqual([...bytes]);
    const zip = Buffer.from(zipSync({ "ppt/slides/slide2.xml": strToU8("<a/>"), "ppt/slides/slide1.xml": strToU8("<b/>") })).toString("base64");
    expect(unpackSlideFile(zip)).toMatchObject({ slidePath: "ppt/slides/slide1.xml", relsPath: "ppt/slides/_rels/slide1.xml.rels" });
    expect(strFromU8(unzipSync(base64ToBytes(zip))["ppt/slides/slide1.xml"]!)).toBe("<b/>");
    expect(() => unpackSlideFile(Buffer.from(zipSync({ "x.txt": strToU8("x") })).toString("base64"))).toThrow(/no ppt\/slides/);
    expect(() => unpackSlideFile("bm90IGEgemlw")).toThrow(/not a readable/);
  });

  it("the whole tool surface answers unknown_tool for other names", async () => {
    const result = await call("read_range", {});
    expect(result).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
  });
});
