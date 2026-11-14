// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { filesFromClipboard, filesFromDrop, isFileDrag } from "../frontend/taskpane/ui/transfer-files.js";

function clipboard(text: string, files: File[], withItems = true) {
  return {
    files: files as unknown as FileList,
    items: (withItems ? files.map((file) => ({ kind: "file", getAsFile: () => file })) : []) as unknown as DataTransferItemList,
    getData: (type: string) => (type === "text/plain" ? text : ""),
  };
}

describe("filesFromClipboard", () => {
  it("returns a pasted screenshot with a unique name instead of image.png", () => {
    const shot = new File(["x"], "image.png", { type: "image/png" });
    const [file] = filesFromClipboard(clipboard("", [shot]), 1700000000000);
    expect(file!.name).toBe("pasted-image-1700000000000.png");
    expect(file!.type).toBe("image/png");
  });

  it("keeps the real name of a copied file and numbers several screenshots", () => {
    const report = new File(["x"], "report.xlsx");
    expect(filesFromClipboard(clipboard("", [report]))[0]).toBe(report);
    const names = filesFromClipboard(clipboard("", [
      new File(["a"], "image.png", { type: "image/png" }),
      new File(["b"], "image.png", { type: "image/png" }),
    ]), 5).map((file) => file.name);
    expect(names).toEqual(["pasted-image-5.png", "pasted-image-5-2.png"]);
  });

  it("falls back to clipboard files when items are not provided", () => {
    const report = new File(["x"], "notes.md");
    expect(filesFromClipboard(clipboard("", [report], false))).toEqual([report]);
  });

  it("leaves a normal text paste alone, even when the clipboard also carries a picture", () => {
    const picture = new File(["x"], "image.png", { type: "image/png" });
    expect(filesFromClipboard(clipboard("A1\tB1", [picture]))).toEqual([]);
    expect(filesFromClipboard(clipboard("hello", []))).toEqual([]);
    expect(filesFromClipboard(null)).toEqual([]);
  });
});

describe("file drag and drop", () => {
  const entry = (isDirectory: boolean) => ({ isDirectory }) as unknown as FileSystemEntry;
  const dropped = (rows: Array<{ file?: File; folder?: boolean }>) => ({
    files: rows.flatMap((row) => row.file ?? []) as unknown as FileList,
    items: rows.map((row) => ({
      kind: "file",
      getAsFile: () => row.file ?? null,
      webkitGetAsEntry: () => entry(!!row.folder),
    })) as unknown as DataTransferItemList,
  });

  it("only treats a drag that carries files as a file drag", () => {
    expect(isFileDrag({ types: ["Files"] as unknown as readonly string[] })).toBe(true);
    expect(isFileDrag({ types: ["text/plain", "Files"] as unknown as readonly string[] })).toBe(true);
    expect(isFileDrag({ types: ["text/plain"] as unknown as readonly string[] })).toBe(false);
    expect(isFileDrag(null)).toBe(false);
  });

  it("returns dropped files with their real names", () => {
    const a = new File(["a"], "sales.xlsx");
    const b = new File(["b"], "notes.md");
    expect(filesFromDrop(dropped([{ file: a }, { file: b }]))).toEqual({ files: [a, b], folders: 0 });
  });

  it("skips dropped folders and counts them", () => {
    const a = new File(["a"], "sales.xlsx");
    const result = filesFromDrop(dropped([{ file: a }, { file: new File([], "Reports"), folder: true }]));
    expect(result.files).toEqual([a]);
    expect(result.folders).toBe(1);
  });

  it("falls back to the file list when the browser gives no items", () => {
    const a = new File(["a"], "sales.xlsx");
    expect(filesFromDrop({ files: [a] as unknown as FileList })).toEqual({ files: [a], folders: 0 });
    expect(filesFromDrop(null)).toEqual({ files: [], folders: 0 });
  });
});
