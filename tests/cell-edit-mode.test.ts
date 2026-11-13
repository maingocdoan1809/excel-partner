import { afterEach, describe, expect, it, vi } from "vitest";
import { runExcel } from "../addin/taskpane/executor/common.js";

const editing = Object.assign(new Error("in edit mode"), { code: "InvalidOperationInCellEditMode" });

afterEach(() => { vi.useRealTimers(); delete (globalThis as { Excel?: unknown }).Excel; delete (globalThis as { Office?: unknown }).Office; });

function install(run: (...args: unknown[]) => Promise<unknown>) {
  (globalThis as { Excel?: unknown }).Excel = { run };
  (globalThis as { Office?: unknown }).Office = { context: { platform: "PC" }, PlatformType: { OfficeOnline: "OfficeOnline" } };
}

describe("runExcel while the person is typing in a cell", () => {
  it("runs normally when Excel accepts the call", async () => {
    install(async (batch) => (batch as (c: unknown) => unknown)({ sync: async () => undefined }));
    expect(await runExcel(async () => 7)).toBe(7);
  });

  it("waits with delayForCellEdit and then runs the batch", async () => {
    const run = vi.fn(async (first: unknown, second?: unknown) => {
      if (typeof first === "function") throw editing;
      expect(first).toEqual({ delayForCellEdit: true });
      return (second as (c: unknown) => unknown)({ sync: async () => undefined });
    });
    install(run);
    expect(await runExcel(async () => "done")).toBe("done");
    expect(run).toHaveBeenCalledTimes(2);
  });

  it("gives up with a clear message and never runs the batch late", async () => {
    vi.useFakeTimers();
    let release!: () => void;
    const batch = vi.fn(async () => "late");
    install(async (first: unknown, second?: unknown) => {
      if (typeof first === "function") throw editing;
      await new Promise<void>((resolve) => { release = resolve; });
      return (second as (c: unknown) => unknown)({ sync: async () => undefined });
    });
    const pending = runExcel(batch).catch((error) => error);
    await vi.advanceTimersByTimeAsync(30_000);
    const error = await pending;
    expect(error).toMatchObject({ code: "timeout", message: expect.stringContaining("press Enter or Escape") });
    release();
    await vi.advanceTimersByTimeAsync(0);
    expect(batch).not.toHaveBeenCalled();
  });

  it("does not run a batch again once part of it has already been applied", async () => {
    const run = vi.fn(async (first: unknown) => {
      if (typeof first !== "function") throw new Error("must not wait and retry");
      return (first as (c: { sync(): Promise<void> }) => unknown)({ sync: async () => undefined });
    });
    install(run);
    const batch = vi.fn(async (context: { sync(): Promise<void> }) => {
      await context.sync();
      throw editing;
    });
    await expect(runExcel(batch as never)).rejects.toBe(editing);
    expect(batch).toHaveBeenCalledTimes(1);
    expect(run).toHaveBeenCalledTimes(1);
  });

  it("does not hide other errors", async () => {
    install(async () => { throw new Error("boom"); });
    await expect(runExcel(async () => 1)).rejects.toThrow("boom");
  });
});
