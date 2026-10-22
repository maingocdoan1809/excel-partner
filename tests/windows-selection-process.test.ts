import { beforeEach, describe, expect, it, vi } from "vitest";
import { SelectionSnapshotSchema } from "../protocol/src/index.js";

const processMock = vi.hoisted(() => ({ exec: vi.fn(), input: vi.fn() }));
vi.mock("node:child_process", () => ({ execFile: processMock.exec }));
import { readWindowsSelection } from "../bridge/src/windows-selection.js";

const selection = SelectionSnapshotSchema.parse({ workbookName: "Flow.xlsx", workbookUrl: "file:///C:/My%20Files/Flow.xlsx",
  sheet: "Flow", sheetId: null, address: "A1", rowCount: 1, columnCount: 1, values: [[1]], headers: null, truncated: false });

function reply(stdout: string, error: Error | null = null) {
  processMock.exec.mockImplementation((_command, _args, _options, callback) => {
    queueMicrotask(() => callback(error, stdout));
    return { stdin: { on: vi.fn(), end: processMock.input } };
  });
}
beforeEach(() => { vi.clearAllMocks(); });

describe("native selection process boundary", () => {
  it("passes workbook identity as data and normalizes local file URLs", async () => {
    reply('{"status":"not_shapes"}');
    const workbookName = 'Book\"; $(Remove-Item anything).xlsx';
    expect(await readWindowsSelection({ ...selection, workbookName })).toEqual({ status: "not_shapes" });
    const [command, args, options] = processMock.exec.mock.calls[0]!;
    expect(command).toBe("powershell.exe");
    expect(args).toContain("-EncodedCommand");
    expect(Buffer.from(args.at(-1), "base64").toString("utf16le")).not.toContain(workbookName);
    expect(options).toMatchObject({ timeout: 15000, windowsHide: true, maxBuffer: 512 * 1024 });
    expect(JSON.parse(processMock.input.mock.calls[0]![0])).toMatchObject({ workbookName, workbookPath: "C:\\My Files\\Flow.xlsx", sheet: "Flow" });
  });

  it("normalizes UNC workbook paths", async () => {
    reply('{"status":"not_shapes"}');
    await readWindowsSelection({ ...selection, workbookUrl: "file://server/share/Flow.xlsx" });
    expect(JSON.parse(processMock.input.mock.calls[0]![0]).workbookPath).toBe("\\\\server\\share\\Flow.xlsx");
  });

  it("rejects malformed or empty native shape collections", async () => {
    for (const output of ['not json', '{"status":"ok","workbookName":"Flow.xlsx","sheet":"Flow","shapes":[]}']) {
      reply(output);
      await expect(readWindowsSelection(selection)).rejects.toThrow();
    }
  });

  it("propagates cancellation and subprocess failures to the fallback layer", async () => {
    reply("", new Error("Excel read timed out"));
    const controller = new AbortController();
    await expect(readWindowsSelection(selection, controller.signal)).rejects.toThrow("timed out");
    expect(processMock.exec.mock.calls[0]![2].signal).toBe(controller.signal);
  });
});
