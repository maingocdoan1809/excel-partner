import { noteUserActivity } from "../executor/bulk.js";
import { watchShapeSelectionEvents } from "../executor/selection-events.js";
import { activeHost } from "../host/index.js";
import { monitorSelection } from "./selection.js";

/**
 * Keeps the chip beside the composer showing what is selected in the document. `refresh` reads it again at once, for example when
 * a tool has just changed shapes or the pane gets focus; it does nothing once the pane is closing.
 */
export function setupSelectionChip(chip: HTMLElement): { refresh(): void } {
  const monitor = monitorSelection({
    read: () => activeHost().readSelectionLabel(),
    isVisible: () => document.visibilityState !== "hidden",
    render: value => {
      chip.textContent = value.label;
      chip.setAttribute("aria-label", `Current selection: ${value.label}`);
      chip.title = value.title;
    },
    onError: error => {
      chip.textContent = (error as { code?: string })?.code === "InvalidSelection" ? "Object selected (shape / chart)" : "No selection";
      chip.setAttribute("aria-label", `Current selection: ${chip.textContent}`);
      chip.title = "Current selection could not be identified. Select an object again or focus this pane to retry.";
    },
  });
  // Let Excel finish its native activation/deactivation transition before reading
  // the active object. This is a one-shot debounce, not a polling timer.
  let selectionTimer: number | undefined;
  let closing = false;
  const selectionChanged = () => {
    // A big write in the background gives way while the person is moving around the workbook.
    noteUserActivity();
    window.clearTimeout(selectionTimer);
    selectionTimer = window.setTimeout(() => monitor.refresh(), 80);
  };
  // Shape selection events exist only for Excel; Word and PowerPoint rely on DocumentSelectionChanged, focus and the send-time read.
  const events = activeHost().hasSheets ? watchShapeSelectionEvents({
    onActivated: value => {
      window.clearTimeout(selectionTimer);
      monitor.show(value);
    },
    onSelection: selectionChanged,
    onError: error => { console.warn("Shape selection event registration failed", error); },
  }) : { refresh: async () => undefined, stop: async () => undefined };
  const refresh = () => { if (closing) return; monitor.refresh(); void events.refresh(); };
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, selectionChanged);
  window.addEventListener("focus", refresh);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState !== "hidden") refresh();
  });
  window.addEventListener("pagehide", () => {
    closing = true;
    window.clearTimeout(selectionTimer);
    monitor.stop();
    void events.stop();
  }, { once: true });
  return { refresh };
}
