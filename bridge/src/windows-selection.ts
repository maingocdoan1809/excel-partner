import { execFile } from "node:child_process";
import { z } from "zod";
import { SelectionSnapshotSchema, type SelectionSnapshot } from "../../protocol/src/index.js";
import { WINDOWS_SELECTION_SCRIPT } from "./windows-selection-script.js";

const NativeReply = z.discriminatedUnion("status", [
  z.object({ status: z.literal("ok"), workbookName: z.string(), sheet: z.string(),
    shapes: SelectionSnapshotSchema.shape.selectedShapes.unwrap().min(1).max(1000) }),
  z.object({ status: z.literal("not_shapes") }),
  z.object({ status: z.literal("unavailable"), reason: z.string() }),
]);
export type NativeSelectionReply = z.infer<typeof NativeReply>;
export type SelectionEnricher = (selection: SelectionSnapshot, signal?: AbortSignal) => Promise<SelectionSnapshot>;

function workbookPath(url?: string): string | undefined {
  if (!url) return undefined;
  try {
    const decoded = decodeURIComponent(url);
    if (/^file:\/\//i.test(decoded)) {
      const local = decoded.replace(/^file:\/\//i, "");
      return (/^\/[a-z]:/i.test(local) ? local.slice(1) : local.startsWith("/") ? local : `//${local}`).replaceAll("/", "\\");
    }
    return decoded.replaceAll("/", "\\");
  } catch { return url.replaceAll("/", "\\"); }
}

/** A bounded, read-only COM call. No shell interpolation or workbook changes. */
export function readWindowsSelection(selection: SelectionSnapshot, signal?: AbortSignal): Promise<NativeSelectionReply> {
  return new Promise((resolve, reject) => {
    const child = execFile("powershell.exe", ["-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand",
      Buffer.from(WINDOWS_SELECTION_SCRIPT, "utf16le").toString("base64")],
    { encoding: "utf8", windowsHide: true, timeout: 15000, maxBuffer: 512 * 1024, signal }, (error, stdout) => {
      if (error) { reject(error); return; }
      try { resolve(NativeReply.parse(JSON.parse(stdout.trim().replace(/^\uFEFF/, "")))); }
      catch (error) { reject(error); }
    });
    child.stdin?.on("error", () => undefined);
    child.stdin?.end(JSON.stringify({ workbookName: selection.workbookName, sheet: selection.sheet,
      workbookPath: workbookPath(selection.workbookUrl) }) + "\n");
  });
}

export function mergeWindowsSelection(selection: SelectionSnapshot, reply: NativeSelectionReply): SelectionSnapshot {
  const { selectedShapes: _old, selectedShapesSource: _source, selectedShapesStatus: _status,
    selectedShapesNote: _note, ...base } = selection;
  const unavailable = (reason: string): SelectionSnapshot => ({ ...base, selectedShapesStatus: "unavailable", selectedShapesNote: reason });
  if (reply.status === "unavailable") return unavailable(reply.reason);
  if (reply.status === "not_shapes") {
    return base.activeShape ? unavailable("Selection changed while reading; call get_selection again.") : base;
  }
  if (reply.workbookName.toLowerCase() !== base.workbookName.toLowerCase() || reply.sheet.toLowerCase() !== base.sheet.toLowerCase()) {
    return unavailable("Native selection does not match the task pane workbook and worksheet.");
  }
  const activeShape = base.activeShape && reply.shapes.some(shape => shape.name === base.activeShape!.name) ? base.activeShape : undefined;
  return { ...base, activeShape, address: "", rowCount: 0, columnCount: 0, values: [], headers: null, truncated: false,
    shapeSelectionStatus: "active", selectedShapes: reply.shapes, selectedShapesSource: "windows_com", selectedShapesStatus: "complete",
    note: "Selected shape identities read from Excel Windows. Use names with list_shapes for text, geometry and connections; nativeId is a COM ID and may differ from an Office.js ID." };
}

export const enrichWindowsSelection: SelectionEnricher = async (selection, signal) => {
  if (process.platform !== "win32" || signal?.aborted) return selection;
  if (selection.address && !selection.activeShape) return selection;
  try { return mergeWindowsSelection(selection, await readWindowsSelection(selection, signal)); }
  catch { return mergeWindowsSelection(selection, { status: "unavailable", reason: "Native multi-selection is unavailable (Excel busy, closed, or read timed out). Active-shape information may be incomplete." }); }
};
