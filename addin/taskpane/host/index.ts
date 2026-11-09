import type { HostId } from "../../../protocol/src/index.js";
import { detectHost } from "./detect.js";
import { excelHost } from "./excel.js";
import { powerpointHost } from "./powerpoint.js";
import type { HostAdapter } from "./types.js";
import { wordHost } from "./word.js";

export { detectHost } from "./detect.js";
export type { HostAdapter } from "./types.js";

export function hostAdapterFor(id: HostId): HostAdapter {
  return id === "word" ? wordHost : id === "powerpoint" ? powerpointHost : excelHost;
}

let current: HostAdapter | null = null;

/** Fix the host from what `Office.onReady` reported (the page is ready then, so the answer cannot change). */
export function initHost(reported?: string | null): HostAdapter {
  const id: HostId | null = reported === "Excel" ? "excel" : reported === "Word" ? "word" : reported === "PowerPoint" ? "powerpoint" : null;
  current = hostAdapterFor(id ?? detectHost());
  return current;
}

/** The adapter of the Office application this pane runs in. Before `initHost` it is worked out each time, never remembered. */
export function activeHost(): HostAdapter {
  return current ?? hostAdapterFor(detectHost());
}
