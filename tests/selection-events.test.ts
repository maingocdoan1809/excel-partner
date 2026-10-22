import { beforeEach, describe, expect, it, vi } from "vitest";
import { watchShapeSelectionEvents } from "../addin/taskpane/executor/selection-events.js";
import { installOfficeMock, MockShape, MockWorkbook } from "./support/office-mock.js";

let book: MockWorkbook;
beforeEach(() => { book = new MockWorkbook(); book.addSheet("Sheet1"); installOfficeMock(book); });
const callbacks = () => ({ onSelection: vi.fn(), onError: vi.fn() });

describe("native shape selection subscriptions", () => {
  it("subscribes to both native events and disposes them without a timer", async () => {
    vi.useFakeTimers();
    try {
      const shape = book.sheets[0]!.shapes.addLine(0, 0, 100, 100, "Straight");
      const options = callbacks();
      const watcher = watchShapeSelectionEvents(options);
      await watcher.ready;
      expect(options.onError).not.toHaveBeenCalled();
      expect(shape.onActivated.handlers.size).toBe(1);
      expect(shape.onDeactivated.handlers.size).toBe(1);
      expect(book.trackedObjects.size).toBe(1);
      expect(vi.getTimerCount()).toBe(0);
      await shape.onActivated.emit();
      expect(options.onSelection).toHaveBeenCalledTimes(1);
      await shape.onDeactivated.emit();
      await watcher.refresh();
      expect(options.onSelection).toHaveBeenCalledTimes(2);
      await watcher.stop();
      expect(shape.onActivated.handlers.size).toBe(0);
      expect(shape.onDeactivated.handlers.size).toBe(0);
      expect(book.worksheetActivated.handlers.size).toBe(0);
      expect(book.trackedObjects.size).toBe(0);
      await shape.onActivated.emit();
      expect(options.onSelection).toHaveBeenCalledTimes(2);
    } finally { vi.useRealTimers(); }
  });

  it("discovers added shapes, avoids duplicate handlers and drops deleted objects", async () => {
    const sheet = book.sheets[0]!;
    const first = sheet.shapes.addLine(0, 0, 100, 100, "Straight");
    const watcher = watchShapeSelectionEvents(callbacks());
    await watcher.ready;
    const added = sheet.shapes.addLine(0, 0, 100, 100, "Straight");
    await watcher.refresh();
    await watcher.refresh();
    expect(first.onActivated.handlers.size).toBe(1);
    expect(added.onActivated.handlers.size).toBe(1);
    first.delete();
    await watcher.refresh();
    expect(first.onActivated.handlers.size).toBe(0);
    expect(added.onActivated.handlers.size).toBe(1);
    await watcher.stop();
  });

  it("rebinds when the active worksheet changes", async () => {
    const first = book.sheets[0]!.shapes.addLine(0, 0, 100, 100, "Straight");
    const second = book.addSheet("Sheet2").shapes.addLine(0, 0, 100, 100, "Straight");
    const options = callbacks();
    const watcher = watchShapeSelectionEvents(options);
    await watcher.ready;
    book.active = 1;
    await book.worksheetActivated.emit();
    await watcher.refresh();
    expect(first.onActivated.handlers.size).toBe(0);
    expect(second.onActivated.handlers.size).toBe(1);
    await second.onActivated.emit();
    expect(options.onSelection).toHaveBeenCalledTimes(2);
    await watcher.stop();
  });

  it("includes nested group children without reading their text", async () => {
    const child = new MockShape("GeometricShape", "Rectangle");
    child.textReadFailure = "frame";
    const inner = new MockShape("Group", ""); inner.children = [child];
    const outer = new MockShape("Group", ""); outer.children = [inner];
    book.sheets[0]!.shapeList.push(outer);
    const options = callbacks();
    const watcher = watchShapeSelectionEvents(options);
    await watcher.ready;
    expect(options.onError).not.toHaveBeenCalled();
    expect(child.onActivated.handlers.size).toBe(1);
    expect(inner.onActivated.handlers.size).toBe(1);
    expect(outer.onActivated.handlers.size).toBe(1);
    await watcher.stop();
  });

  it("uses native activation identity and ignores a late deactivation from the previous shape", async () => {
    const sheet = book.sheets[0]!;
    const a = sheet.shapes.addLine(0, 0, 100, 100, "Straight"); a.name = "A";
    const b = sheet.shapes.addLine(0, 0, 100, 100, "Straight"); b.name = "B";
    const options = { ...callbacks(), onActivated: vi.fn() };
    const watcher = watchShapeSelectionEvents(options);
    await watcher.ready;
    await a.onActivated.emit();
    await b.onActivated.emit();
    await a.onDeactivated.emit();
    expect(options.onActivated.mock.calls.map(call => call[0].label)).toEqual(["A", "B"]);
    expect(options.onSelection).not.toHaveBeenCalled();
    await b.onDeactivated.emit();
    expect(options.onSelection).toHaveBeenCalledTimes(1);
    await watcher.stop();
  });

  it("does not register unsupported APIs on older hosts", async () => {
    book.supportsShapes = false;
    const watcher = watchShapeSelectionEvents(callbacks());
    await watcher.ready;
    await watcher.refresh();
    expect(book.worksheetActivated.handlers.size).toBe(0);
    await watcher.stop();
  });
});
