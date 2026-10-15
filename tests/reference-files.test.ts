import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { strToU8, zipSync } from "fflate";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { ReferenceError_, ReferenceFiles, toLocalPath, workbookFolderFromUrl } from "../bridge/src/reference-files.js";
import { readSheetDetails, readWorkbookOverview } from "../bridge/src/xlsx-reader.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { executeTool } from "../addin/taskpane/executor/index.js";
import { parseToolInput } from "../protocol/src/index.js";
import { MockWorkbook, installOfficeMock } from "./support/office-mock.js";

/** A tiny but real workbook: 2 sheets, shared strings, a merge, a bold grey header and a drawn rectangle. */
function buildXlsx(): Uint8Array {
  const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  return zipSync({
    "[Content_Types].xml": strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
    "xl/workbook.xml": strToU8(
      `<workbook ${ns} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
        '<sheet name="Main" sheetId="1" r:id="rId1"/><sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/></sheets>' +
        '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Main!$A$1:$D$5</definedName></definedNames></workbook>',
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="x/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
    ),
    "xl/sharedStrings.xml": strToU8(`<sst ${ns}><si><t>Item</t></si><si><r><t>Qty</t></r><rPh><t>フリガナ</t></rPh></si></sst>`),
    "xl/styles.xml": strToU8(
      `<styleSheet ${ns}><fonts><font><sz val="9"/><name val="Arial"/></font><font><b/><sz val="9"/><color rgb="FFFF0000"/><name val="Arial"/></font></fonts>` +
        '<fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD8D8D8"/></patternFill></fill></fills>' +
        '<borders><border/><border><left style="thin"/><bottom style="double"/></border></borders>' +
        '<cellXfs><xf fontId="0" fillId="0" borderId="0"/><xf fontId="1" fillId="2" borderId="1"><alignment horizontal="center"/></xf></cellXfs></styleSheet>',
    ),
    "xl/worksheets/sheet1.xml": strToU8(
      `<worksheet ${ns}><dimension ref="A1:C3"/><sheetViews><sheetView showGridLines="0" zoomScale="80"/></sheetViews>` +
        '<sheetFormatPr defaultColWidth="2.625" defaultRowHeight="12.95"/><sheetData>' +
        '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c><c r="C1" s="1"/></row>' +
        '<row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2"><v>3</v></c><c r="C2"><f>B2*2</f><v>6</v></c></row></sheetData>' +
        '<mergeCells count="1"><mergeCell ref="A3:C3"/></mergeCells><pageSetup orientation="landscape"/></worksheet>',
    ),
    "xl/worksheets/sheet2.xml": strToU8(`<worksheet ${ns}><dimension ref="A1"/><sheetData/></worksheet>`),
  });
}

let root: string;
beforeEach(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), "ref-"));
  fs.mkdirSync(path.join(root, "templates"));
  fs.writeFileSync(path.join(root, "templates", "t.xlsx"), buildXlsx());
  fs.writeFileSync(path.join(root, "templates", "notes.txt"), "hello");
});
afterEach(() => fs.rmSync(root, { recursive: true, force: true }));

describe("toLocalPath", () => {
  it("maps Windows and WSL UNC paths to Linux paths", () => {
    expect(toLocalPath("C:\\Users\\me\\a b\\t.xlsx")).toBe("/mnt/c/Users/me/a b/t.xlsx");
    expect(toLocalPath('"d:/data/x.xlsx"')).toBe("/mnt/d/data/x.xlsx");
    expect(toLocalPath("\\\\wsl.localhost\\Ubuntu\\home\\me\\x.xlsx")).toBe("/home/me/x.xlsx");
    expect(toLocalPath("/already/linux")).toBe("/already/linux");
  });
});

describe("ReferenceFiles", () => {
  it("uses the saved workbook folder as a default workspace and accepts added folders", () => {
    const file = path.join(root, "templates", "t.xlsx");
    expect(workbookFolderFromUrl(`file://${file}`)).toBe(path.join(root, "templates"));
    expect(workbookFolderFromUrl(file)).toBe(path.join(root, "templates"));
    expect(() => workbookFolderFromUrl("https://example.com/book.xlsx")).toThrow(/cloud URL/);
    const refs = new ReferenceFiles([]);
    expect(refs.addRoot(path.join(root, "templates"))).toBe(path.join(root, "templates"));
    expect(refs.list(path.join(root, "templates")).entries.some((entry) => entry.name === "t.xlsx")).toBe(true);
    expect(refs.readSelectedFile(file).bytes.subarray(0, 2).toString()).toBe("PK");
    expect(refs.readSelectedFile(path.join(root, "templates", "notes.txt")).bytes.toString()).toBe("hello");
  });
  it("refuses everything until an exact file is approved", () => {
    expect(() => new ReferenceFiles([]).resolve(path.join(root, "templates", "t.xlsx"))).toThrow(/outside/);
  });

  it("allows paths inside the folder and rejects outside ones and traversal", () => {
    const refs = new ReferenceFiles([root]);
    expect(refs.resolve(path.join(root, "templates", "t.xlsx"))).toContain("t.xlsx");
    expect(() => refs.resolve("/etc/passwd")).toThrow(/outside/);
    expect(() => refs.resolve(path.join(root, "..", "x"))).toThrow(ReferenceError_);
    expect(() => refs.resolve(path.join(root, "templates", "..", "..", "etc"))).toThrow(ReferenceError_);
  });

  it("does not follow a symlink out of the folder", () => {
    const outside = fs.mkdtempSync(path.join(os.tmpdir(), "outside-"));
    fs.writeFileSync(path.join(outside, "secret.xlsx"), buildXlsx());
    fs.symlinkSync(path.join(outside, "secret.xlsx"), path.join(root, "templates", "link.xlsx"));
    expect(() => new ReferenceFiles([root]).readWorkbook(path.join(root, "templates", "link.xlsx"))).toThrow(/outside/);
    fs.rmSync(outside, { recursive: true, force: true });
  });

  it("lists folders first, hides temp files and only reads workbooks", () => {
    fs.writeFileSync(path.join(root, "templates", "~$lock.xlsx"), "x");
    const refs = new ReferenceFiles([root]);
    expect(refs.list().entries[0]!.type).toBe("folder");
    expect(refs.list(path.join(root, "templates")).entries.map((e) => e.name)).toEqual(["notes.txt", "t.xlsx"]);
    expect(() => refs.readWorkbook(path.join(root, "templates", "notes.txt"))).toThrow(/only Excel workbooks/);
  });
});

describe("xlsx reader", () => {
  const bytes = buildXlsx();

  it("summarises sheets, state, print area, zoom and gridlines", () => {
    const o = readWorkbookOverview(bytes);
    expect(o.sheets[0]).toMatchObject({ name: "Main", state: "visible", dimension: "A1:C3", zoom: 80, gridlines: false, printArea: "$A$1:$D$5", orientation: "landscape", mergeCount: 1 });
    expect(o.sheets[1]).toMatchObject({ name: "Hidden", state: "hidden" });
  });

  it("reads values (without furigana), formulas, merges, sizes and styled empty cells", () => {
    const d = readSheetDetails(bytes, "Main");
    expect(d.cells).toEqual({ A1: "Item", B1: "Qty", A2: "Item", B2: 3, C2: [6, "=B2*2"] });
    expect(d.merges).toEqual(["A3:C3"]);
    expect(d.defaultColumnWidth).toBe(2.625);
    expect(d.styles).toMatchObject({ "A1:C1": { bold: true, fontColor: "#FF0000", backgroundColor: "#D8D8D8", horizontalAlignment: "center" } });
    expect(d.borders!["A1:C1"]).toMatchObject({ left: "thin solid #000000", bottom: "thin double #000000" });
  });

  it("limits to a range, matches sheet names loosely and explains unknown sheets", () => {
    expect(Object.keys(readSheetDetails(bytes, " main ", { range: "A2:B2" }).cells)).toEqual(["A2", "B2"]);
    expect(() => readSheetDetails(bytes, "Nope")).toThrow(/Sheets: Main \| Hidden/);
    expect(() => readWorkbookOverview(new Uint8Array([1, 2, 3]))).toThrow(/not a valid .xlsx/);
  });
});

describe("gateway file tools", () => {
  function setup() {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("0123456789abcdef", null);
    const limits = { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 };
    const gateway = new ToolGateway(sessions, limits, 1000, undefined, false, null, new ReferenceFiles([root]));
    return { session, gateway };
  }
  const file = () => path.join(root, "templates", "t.xlsx");

  it("lists the allowed folders, reads an overview and a sheet", async () => {
    const { session, gateway } = setup();
    const call = (tool: string, args: unknown) => gateway.invoke({ sessionId: session.id, tool, args });
    expect(await call("list_reference_files", {})).toMatchObject({ ok: true, data: { entries: [{ type: "folder" }] } });
    expect(await call("read_workbook_file", { path: file() })).toMatchObject({ ok: true, data: { sheets: [{ name: "Main" }, { name: "Hidden" }] } });
    expect(await call("read_workbook_file", { path: file(), sheet: "Main", range: "A1:C2" })).toMatchObject({ ok: true, data: { sheet: "Main", cells: { A1: "Item" } } });
    expect(await call("read_workbook_file", { path: "/etc/passwd" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });

  it("attaches the file bytes to insert_sheets_from_file and validates sheet names", async () => {
    const { session, gateway } = setup();
    let forwarded: any = null;
    session.requestTool = async (tool: string, args: unknown) => {
      forwarded = { tool, args };
      return { ok: true, data: { inserted: ["Main"], count: 1 } };
    };
    const ok = await gateway.invoke({ sessionId: session.id, tool: "insert_sheets_from_file", args: { path: file(), sheets: ["Main"] } });
    expect(ok.ok).toBe(true);
    expect(forwarded.tool).toBe("insert_sheets_from_file");
    expect(Buffer.from(forwarded.args.fileBase64, "base64").subarray(0, 2).toString()).toBe("PK"); // a zip
    expect(forwarded.args.sheets).toEqual(["Main"]);

    forwarded = null;
    const bad = await gateway.invoke({ sessionId: session.id, tool: "insert_sheets_from_file", args: { path: file(), sheets: ["Nope"] } });
    expect(bad).toMatchObject({ ok: false, error: { code: "invalid_arguments", detail: { available: ["Main", "Hidden"] } } });
    expect(forwarded).toBeNull();
  });

  it("asks for one file, reads it after approval, and never grants its folder", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("0123456789abcdef", null);
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000, undefined, false, null, new ReferenceFiles([]));
    const requests: unknown[] = [];
    session.requestTool = async (tool, args) => {
      requests.push({ tool, args });
      return { ok: true, data: { approved: true } };
    };
    Object.defineProperty(session, "isConnected", { get: () => true });
    const result = await gateway.invoke({ sessionId: session.id, tool: "list_reference_files", args: {} });
    expect(result).toMatchObject({ ok: true });
    const read = await gateway.invoke({ sessionId: session.id, tool: "read_workbook_file", args: { path: file() } });
    expect(read.ok).toBe(true);
    if (read.ok) expect((read.data as { sheets: { name: string }[] }).sheets[0]!.name).toBe("Main");
    expect(requests).toEqual([{ tool: "__authorize_reference_path", args: { path: file(), resolvedPath: file() } }]);
    expect((await gateway.invoke({ sessionId: session.id, tool: "read_workbook_file", args: { path: file() } })).ok).toBe(true);
    expect(requests).toHaveLength(1);
    expect(await gateway.invoke({ sessionId: session.id, tool: "list_reference_files", args: { path: path.dirname(file()) } })).toMatchObject({ ok: false });
  });

  it("does not read an outside file when the user denies access", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("0123456789abcdef", null);
    session.requestTool = async () => ({ ok: true, data: { approved: false } });
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000, undefined, false, null, new ReferenceFiles([]));
    const read = await gateway.invoke({ sessionId: session.id, tool: "read_workbook_file", args: { path: file() } });
    expect(read).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(session.referenceGrants.size).toBe(0);
  });
});

describe("pane: insert_sheets_from_file", () => {
  let book: MockWorkbook;
  beforeEach(() => {
    book = new MockWorkbook();
    book.addSheet("Sheet1");
    installOfficeMock(book);
  });
  const call = (args: unknown) => executeTool("insert_sheets_from_file", { ...parseToolInput("insert_sheets_from_file", args), fileBase64: "UEsDBA==" });

  it("inserts the requested sheets and reports their names", async () => {
    const result = await call({ path: "x", sheets: ["表紙", "改版履歴_History"] });
    expect(result).toMatchObject({ ok: true, data: { inserted: ["表紙", "改版履歴_History"], count: 2 } });
    expect(book.inserted[0]).toMatchObject({ options: { positionType: "End", sheetNamesToInsert: ["表紙", "改版履歴_History"] } });
  });

  it("supports inserting after a sheet and validates the anchor", async () => {
    expect((await call({ path: "x", position: "after", afterSheet: "Sheet1" })).ok).toBe(true);
    expect(await call({ path: "x", position: "after", afterSheet: "Missing" })).toMatchObject({ ok: false, error: { code: "sheet_not_found" } });
    expect(await call({ path: "x", position: "after" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });

  it("explains unsupported Excel versions and a missing payload", async () => {
    book.supportsInsertFromFile = false;
    expect(await call({ path: "x" })).toMatchObject({ ok: false, error: { code: "unsupported" } });
    book.supportsInsertFromFile = true;
    expect(await executeTool("insert_sheets_from_file", parseToolInput("insert_sheets_from_file", { path: "x" }))).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });
});
