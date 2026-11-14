import type { HostId } from "../../../protocol/src/index.js";

/**
 * Which Office application this task pane runs in. `Office.context.host` is the source; the requirement sets are only
 * a fallback for when it is missing. A page opened in a plain browser tab has neither: it is treated as Excel.
 */
export function detectHost(): HostId {
  if (typeof Office === "undefined") return "excel";
  const reported = (Office.context as { host?: unknown } | undefined)?.host as string | undefined;
  if (reported === "Excel") return "excel";
  if (reported === "Word") return "word";
  if (reported === "PowerPoint") return "powerpoint";
  try {
    const requirements = Office.context?.requirements;
    if (requirements?.isSetSupported("ExcelApi", "1.1")) return "excel";
    if (requirements?.isSetSupported("WordApi", "1.1")) return "word";
    if (requirements?.isSetSupported("PowerPointApi", "1.1")) return "powerpoint";
  } catch { /* no Office behind this page */ }
  return "excel";
}
