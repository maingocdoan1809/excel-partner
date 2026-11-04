/**
 * Helpers for the live view of the assistant's browser. They read what Playwright MCP returns (which varies a
 * little between releases) defensively: anything unexpected simply means "no label" or "no frame", never an error
 * for the assistant's own tool call.
 */

/** Tools whose result can change what the page looks like. Reading tools (snapshot, console, network) are left out. */
export const FRAME_AFTER = new Set([
  "browser_navigate", "browser_navigate_back", "browser_click", "browser_type", "browser_press_key", "browser_fill_form",
  "browser_select_option", "browser_hover", "browser_drag", "browser_drop", "browser_file_upload", "browser_handle_dialog",
  "browser_wait_for", "browser_tabs", "browser_resize", "browser_evaluate", "browser_run_code_unsafe", "browser_emulate_media", "browser_find",
]);

/** Shown the first time the assistant only reads the page, so the panel is not empty. */
export const FRAME_ON_FIRST_READ = new Set(["browser_snapshot"]);

export interface ToolContent { type: string; text?: string; data?: string; mimeType?: string }
export interface ToolOutput { isError?: boolean; content?: ToolContent[] }

/** `https://host/path?token=x#y` -> `host/path`. Query strings and fragments can carry tokens, so they are never shown. */
export function displayUrl(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  try {
    const url = new URL(raw);
    if (url.protocol === "about:" || url.protocol === "data:") return url.protocol + (url.protocol === "about:" ? url.pathname : "");
    return `${url.host}${url.pathname === "/" ? "" : url.pathname}`.slice(0, 500);
  } catch { return undefined; }
}

const quote = (text: unknown, max = 60): string => {
  const clean = String(text ?? "").replace(/\s+/g, " ").trim();
  return clean ? `"${clean.length > max ? `${clean.slice(0, max - 1)}…` : clean}"` : "";
};
/** Only named keys are shown; a single character could be part of a password. */
const NAMED_KEY = /^(?:(?:Control|ControlOrMeta|Alt|Shift|Meta)\+)*(?:Enter|Tab|Escape|Backspace|Delete|Space|Home|End|PageUp|PageDown|Arrow(?:Up|Down|Left|Right)|F\d{1,2})$/;
const target = (args: Record<string, unknown>) => quote(args.element ?? args.description);

/**
 * One line saying what the assistant is doing. Typed text, key sequences for passwords, evaluated code and
 * form values are deliberately left out: the label is shown on screen and must be safe to show.
 */
export function describeBrowserAction(tool: string, args: Record<string, unknown> = {}): string {
  switch (tool) {
    case "browser_navigate": return `Opening ${displayUrl(String(args.url ?? "")) ?? "a page"}`;
    case "browser_navigate_back": return "Going back";
    case "browser_click": return `Clicking ${target(args) || "an element"}`;
    case "browser_type": return `Typing into ${target(args) || "a field"}`;
    case "browser_press_key": return NAMED_KEY.test(String(args.key ?? "")) ? `Pressing ${args.key}` : "Pressing a key";
    case "browser_fill_form": return `Filling ${Array.isArray(args.fields) ? args.fields.length : "several"} field(s)`;
    case "browser_select_option": return `Choosing an option in ${target(args) || "a list"}`;
    case "browser_hover": return `Pointing at ${target(args) || "an element"}`;
    case "browser_drag": return "Dragging an element";
    case "browser_drop": return "Dropping onto an element";
    case "browser_file_upload": return "Choosing a file";
    case "browser_handle_dialog": return "Answering a dialog";
    case "browser_wait_for": return "Waiting for the page";
    case "browser_tabs": return "Switching tabs";
    case "browser_resize": return "Resizing the window";
    case "browser_take_screenshot": return "Taking a screenshot";
    case "browser_snapshot": return "Reading the page";
    case "browser_find": return `Looking for ${quote(args.query ?? args.text) || "an element"}`;
    case "browser_evaluate": case "browser_run_code_unsafe": return "Running a script on the page";
    default: return tool.replace(/^browser_/, "").replace(/_/g, " ");
  }
}

const text = (output: ToolOutput | undefined): string =>
  (output?.content ?? []).filter(item => item.type === "text" && typeof item.text === "string").map(item => item.text).join("\n");

/** `### Page` / `- Page URL:` / `- Page Title:` lines of a Playwright MCP result. */
export function parsePageInfo(output: ToolOutput | undefined): { url?: string; title?: string } {
  const body = text(output);
  const url = /^- Page URL: (.+)$/m.exec(body)?.[1]?.trim();
  const title = /^- Page Title: (.+)$/m.exec(body)?.[1]?.trim();
  return { url: displayUrl(url), title: title ? title.slice(0, 200) : undefined };
}

export function extractImage(output: ToolOutput | undefined): { mimeType: "image/jpeg" | "image/png"; data: string } | null {
  for (const item of output?.content ?? []) {
    if (item.type === "image" && typeof item.data === "string" && item.data && (item.mimeType === "image/jpeg" || item.mimeType === "image/png")) {
      return { mimeType: item.mimeType, data: item.data };
    }
  }
  return null;
}
