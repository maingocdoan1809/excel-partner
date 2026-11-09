import type { HostId } from "../../../protocol/src/index.js";
import { readOpenFile } from "../host/file.js";

/** The open file (xlsx, docx or pptx, by host) as zip bytes, read through Office. */
export function readDocumentFile(host: HostId = "excel"): Promise<Uint8Array> {
  return readOpenFile(host);
}

/** The open workbook as an .xlsx file (zip bytes), read through Office. */
export function readWorkbookFile(): Promise<Uint8Array> {
  return readDocumentFile("excel");
}
