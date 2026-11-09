import { OTHER_FILE_NAME_FALLBACK } from "./delegation.js";
import { toolChevron } from "./tool-view.js";

/**
 * A request another open file's assistant sent to this chat. It is shown as what it is, not as something the person
 * typed: a collapsible "Received message" with who it came from and the text quoted.
 */
const APPS = { excel: "Excel", word: "Word", powerpoint: "PowerPoint" } as const;

export function renderReceivedMessage(fromName: string, text: string, host: keyof typeof APPS = "excel"): HTMLDetailsElement {
  const name = fromName.replace(/\s+/g, " ").trim() || OTHER_FILE_NAME_FALLBACK;
  const block = document.createElement("details");
  block.className = "received";
  block.open = true;

  const summary = document.createElement("summary");
  const label = document.createElement("span");
  label.textContent = "Received message";
  const chevron = document.createElement("span");
  chevron.className = "received-chevron";
  chevron.append(toolChevron());
  summary.append(label, chevron);

  const body = document.createElement("div");
  body.className = "received-body";
  const from = document.createElement("div");
  from.className = "received-from";
  from.title = name;
  const arrow = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  arrow.setAttribute("viewBox", "0 0 20 20");
  arrow.setAttribute("aria-hidden", "true");
  arrow.classList.add("received-arrow");
  arrow.innerHTML = '<path d="M14.5 5.5 6 14m0 0V8.5M6 14h5.5" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/>';
  const source = document.createElement("span");
  source.className = "received-source";
  source.textContent = `From ${APPS[host] ?? APPS.excel} · ${name}`;
  from.append(arrow, source);
  const quote = document.createElement("blockquote");
  quote.className = "received-text";
  quote.textContent = text;
  body.append(from, quote);

  block.append(summary, body);
  return block;
}
