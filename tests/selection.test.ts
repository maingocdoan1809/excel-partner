import { afterEach, describe, expect, it, vi } from "vitest";
import { monitorSelection, type SelectionLabel } from "../addin/taskpane/ui/selection.js";

const label = (name: string): SelectionLabel => ({ label: name, title: `Selected ${name}` });
afterEach(() => { vi.clearAllTimers(); vi.useRealTimers(); });

describe("composer selection monitor", () => {
  it("does not read on a timer and updates when a native selection event requests it", async () => {
    vi.useFakeTimers();
    let selected = "Connector_A";
    const render = vi.fn();
    const read = vi.fn(async () => label(selected));
    const monitor = monitorSelection({ read, render, onError: vi.fn(), isVisible: () => true });
    await vi.advanceTimersByTimeAsync(0);
    expect(render).toHaveBeenLastCalledWith(label("Connector_A"));
    selected = "Connector_B";
    await vi.advanceTimersByTimeAsync(10_000);
    expect(read).toHaveBeenCalledTimes(1);
    expect(vi.getTimerCount()).toBe(0);
    monitor.refresh();
    await vi.advanceTimersByTimeAsync(0);
    expect(render).toHaveBeenLastCalledWith(label("Connector_B"));
    monitor.refresh();
    await vi.advanceTimersByTimeAsync(0);
    expect(render).toHaveBeenCalledTimes(2); // unchanged events do not repeat announcements
    monitor.stop();
  });

  it("serializes slow reads and coalesces simultaneous selection events", async () => {
    vi.useFakeTimers();
    let finish!: (value: SelectionLabel) => void;
    const read = vi.fn(() => new Promise<SelectionLabel>(resolve => { finish = resolve; }));
    const monitor = monitorSelection({ read, render: vi.fn(), onError: vi.fn(), isVisible: () => true });
    monitor.refresh();
    monitor.refresh();
    await vi.advanceTimersByTimeAsync(3000);
    expect(read).toHaveBeenCalledTimes(1);
    finish(label("Connector_A"));
    await vi.advanceTimersByTimeAsync(0);
    expect(read).toHaveBeenCalledTimes(2);
    monitor.stop();
  });

  it("discards an old result when a newer selection event arrives mid-read", async () => {
    vi.useFakeTimers();
    let finish!: (value: SelectionLabel) => void;
    const read = vi.fn().mockImplementationOnce(() => new Promise<SelectionLabel>(resolve => { finish = resolve; }))
      .mockResolvedValue(label("Connector_B"));
    const render = vi.fn();
    const monitor = monitorSelection({ read, render, onError: vi.fn(), isVisible: () => true });
    monitor.refresh();
    monitor.refresh();
    finish(label("Connector_A"));
    await vi.advanceTimersByTimeAsync(0);
    expect(render.mock.calls).toEqual([[label("Connector_B")]]);
    expect(read).toHaveBeenCalledTimes(2);
    monitor.stop();
  });

  it("recovers from a transient Office read error on the next selection event", async () => {
    vi.useFakeTimers();
    const error = new Error("Excel is busy");
    const read = vi.fn().mockRejectedValueOnce(error).mockResolvedValue(label("Connector_B"));
    const render = vi.fn(), onError = vi.fn();
    const monitor = monitorSelection({ read, render, onError, isVisible: () => true });
    await vi.advanceTimersByTimeAsync(0);
    expect(onError).toHaveBeenCalledWith(error);
    monitor.refresh();
    await vi.advanceTimersByTimeAsync(0);
    expect(render).toHaveBeenLastCalledWith(label("Connector_B"));
    monitor.stop();
  });

  it("pauses when hidden and refreshes immediately when made visible", async () => {
    vi.useFakeTimers();
    let visible = false;
    const read = vi.fn().mockResolvedValue(label("Connector_B"));
    const monitor = monitorSelection({ read, render: vi.fn(), onError: vi.fn(), isVisible: () => visible });
    await vi.advanceTimersByTimeAsync(1500);
    expect(read).not.toHaveBeenCalled();
    visible = true;
    monitor.refresh();
    await vi.advanceTimersByTimeAsync(0);
    expect(read).toHaveBeenCalledTimes(1);
    monitor.stop();
    await vi.advanceTimersByTimeAsync(1500);
    expect(read).toHaveBeenCalledTimes(1);
  });

  it("keeps native activation identity when an older cell read finishes later", async () => {
    vi.useFakeTimers();
    let finish!: (value: SelectionLabel) => void;
    const render = vi.fn();
    const monitor = monitorSelection({ read: () => new Promise<SelectionLabel>(resolve => { finish = resolve; }), render, onError: vi.fn(), isVisible: () => true });
    monitor.show(label("Connector_B"));
    finish(label("Sheet1!C5"));
    await vi.advanceTimersByTimeAsync(0);
    expect(render.mock.calls).toEqual([[label("Connector_B")]]);
    monitor.stop();
  });

  it("does not render a pending read after the page is disposed", async () => {
    vi.useFakeTimers();
    let finish!: (value: SelectionLabel) => void;
    const render = vi.fn();
    const monitor = monitorSelection({ read: () => new Promise<SelectionLabel>(resolve => { finish = resolve; }), render, onError: vi.fn(), isVisible: () => true });
    monitor.stop();
    finish(label("Connector_A"));
    await vi.advanceTimersByTimeAsync(0);
    expect(render).not.toHaveBeenCalled();
  });
});
