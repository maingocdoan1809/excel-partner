import { zipSync, strToU8 } from "fflate";
import { describe, expect, it } from "vitest";
import { readWorkbookOverview, XlsxError } from "../bridge/src/xlsx-reader.js";

/** A tiny zip whose central directory claims each part inflates to `declared` bytes. */
function bomb(declared: number, parts = 1): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  for (let i = 0; i < parts; i++) files[`xl/worksheets/sheet${i}.xml`] = strToU8("<worksheet/>");
  const zip = zipSync(files);
  const view = new DataView(zip.buffer);
  for (let at = 0; at < zip.length - 4; at++) if (view.getUint32(at, true) === 0x02014b50) view.setUint32(at + 24, declared, true);
  return zip;
}

describe("workbook zip size guard", () => {
  it("refuses a part that declares more than the per-part limit before inflating it", () => {
    expect(() => readWorkbookOverview(bomb(0x7fffffff))).toThrow(XlsxError);
    expect(() => readWorkbookOverview(bomb(0x7fffffff))).toThrow(/expands to more than/);
  });

  it("refuses parts that together declare more than the total limit", () => {
    expect(() => readWorkbookOverview(bomb(60 * 1024 * 1024, 5))).toThrow(/expands to more than/);
  });
});
