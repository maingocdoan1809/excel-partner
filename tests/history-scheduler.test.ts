import { afterEach, expect, it, vi } from "vitest";
import { createHistoryScheduler } from "../addin/taskpane/ui/history-scheduler.js";
afterEach(() => vi.useRealTimers());
it("H14 checkpoints continuous streaming within 500ms and flushes once on close", () => {
  vi.useFakeTimers(); const persist = vi.fn(); const scheduler = createHistoryScheduler(persist);
  for (let i = 0; i < 30; i++) { scheduler.schedule(); vi.advanceTimersByTime(20); }
  expect(persist).toHaveBeenCalledTimes(1);
  scheduler.flush(); expect(persist).toHaveBeenCalledTimes(2);
  vi.advanceTimersByTime(1000); expect(persist).toHaveBeenCalledTimes(2);
});
