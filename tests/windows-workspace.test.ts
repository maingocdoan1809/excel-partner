import { afterEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({ execFile: vi.fn(), run: vi.fn(), spawn: vi.fn() }));
vi.mock("node:child_process", () => {
  Object.assign(mocks.execFile, { [Symbol.for("nodejs.util.promisify.custom")]: mocks.run });
  return { execFile: mocks.execFile, spawn: mocks.spawn };
});

import { EventEmitter } from "node:events";
import { activeWorkbookPath, pickReferenceFile, stopWorkspaceHelper } from "../bridge/src/windows-workspace.js";

afterEach(() => {
  stopWorkspaceHelper();
  vi.resetAllMocks();
});

function mockCommands(selected: string | null): void {
  mocks.run.mockImplementation(async (command) => {
    if (command === "wslpath") return { stdout: "C:\\excel-partner\\scripts\\workspace-windows.ps1\n" };
    return { stdout: selected ? selected + "\r\n" : "" };
  });
}

describe("Windows reference picker", () => {
  it("returns the chosen path or cancellation from a dedicated STA process", async () => {
    mockCommands("C:\\files\\rules.xlsx");
    expect(await pickReferenceFile()).toBe("C:\\files\\rules.xlsx");
    expect(mocks.run).toHaveBeenLastCalledWith(expect.any(String), expect.arrayContaining([
      "-STA", "-Mode", "file",
    ]), expect.objectContaining({ timeout: 600_000 }));
    expect(mocks.spawn).not.toHaveBeenCalled();
    mockCommands(null);
    expect(await pickReferenceFile()).toBeNull();
  });

  it("opens while the shared helper is still waiting for Excel COM", async () => {
    mockCommands("C:\\files\\rules.xlsx");
    const child = Object.assign(new EventEmitter(), {
      stdout: Object.assign(new EventEmitter(), { setEncoding: vi.fn() }),
      stderr: new EventEmitter(),
      stdin: { write: vi.fn(), end: vi.fn() },
      kill: vi.fn(),
    });
    mocks.spawn.mockImplementation(() => {
      queueMicrotask(() => child.stdout.emit("data", '{"ready":true}\n'));
      return child;
    });
    const workbook = activeWorkbookPath("open.xlsx");
    await vi.waitFor(() => expect(child.stdin.write).toHaveBeenCalledOnce());
    expect(await pickReferenceFile()).toBe("C:\\files\\rules.xlsx");
    expect(child.stdin.write).toHaveBeenCalledOnce();
    const request = JSON.parse(child.stdin.write.mock.calls[0]![0] as string);
    child.stdout.emit("data", JSON.stringify({ id: request.id, path: null }) + "\n");
    expect(await workbook).toBeNull();
  });
});
