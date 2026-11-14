import { exposedToolNamesForHost } from "../../../protocol/src/index.js";
import { activeHost } from "../host/index.js";
import { mentionsFile, mentionsSheet } from "../ui/cell-links.js";
import type { StoredMessage, StoredTool } from "../ui/sessions.js";
import { renderReceivedMessage } from "../ui/received-message.js";
import { leadingCommand } from "../ui/slash.js";
import { renderMarkdown } from "../ui/markdown.js";
import { renderToolCard, toolChevron } from "../ui/tool-view.js";
import { renderArtifactCard } from "../ui/artifact.js";
import { showToast } from "../ui/toast.js";
import { randomId } from "../bridge-client.js";
import { fileMentionRanges } from "../ui/file-mentions.js";
import { follow, preview, thinking } from "./services.js";
import { els } from "../ui/elements.js";
import { schedulePersist } from "./chats.js";
import { type PendingPrompt, app, pendingPrompts } from "./state.js";
import { submit } from "./turn.js";

// Tools the bridge answers by itself reach the pane only as provider events, never as a tool request, so they must keep
// their own card: the person sees which other file was read or sent a request.
export const bridgeAnsweredTools: ReadonlySet<string> = new Set(["list_connected_files", "read_connected_file", "send_to_connected_file"]);
let workbookToolNamesCache: Set<string> | null = null;
/** Tools of the open file's application that reach the pane as tool requests (their card comes from the executed result). */
export const isDocumentTool = (name: string): boolean => (workbookToolNamesCache ??= new Set<string>(exposedToolNamesForHost(activeHost().id, true, { playwright: true, webSearch: true }).filter((tool) => !bridgeAnsweredTools.has(tool)))).has(name);
export const cellLinkOptions = (text: string, aboutOtherFiles = false) => ({ plain: !aboutOtherFiles && !mentionsFile(text, app.otherFiles) && !mentionsSheet(text, app.sheetNames) });
export function renderMessage(container: HTMLElement, message: StoredMessage): HTMLElement {
  // Older saved chats may contain tool-only assistant records with no text.
  // Keep them out of the transcript as well as newly streamed empty messages.
  if (message.kind === "assistant" && !message.text.trim()) return document.createElement("div");
  container.querySelector(".welcome")?.remove();
  let element: HTMLElement;
  switch (message.kind) {
    case "user": {
      if (message.from !== undefined) { element = renderReceivedMessage(message.from, message.text, message.fromHost); break; }
      element = bubble("user");
      const lead = leadingCommand(message.text);
      if (lead) {
        const token = document.createElement("span");
        token.className = "slash-token";
        token.textContent = lead.command;
        element.append(token);
        appendFileMentions(element, lead.rest, message.attachments?.map((item) => item.name) ?? []);
      } else appendFileMentions(element, message.text, message.attachments?.map((item) => item.name) ?? []);
      if (message.delivery) renderDelivery(element, message);
      break;
    }
    case "assistant":
      element = bubble("assistant");
      element.innerHTML = renderMarkdown(message.text, { cellLinks: cellLinkOptions(message.text, message.aboutOtherFiles) });
      break;
    case "error":
      element = bubble("error");
      element.textContent = message.text;
      break;
    case "note":
      element = document.createElement("div");
      element.className = "note";
      element.textContent = message.text;
      break;
    case "tool":
      element = renderToolCard(message.tool!);
      break;
    case "artifact":
      element = renderArtifactCard(message.artifact!, artifact => preview.open(artifact));
      break;
    case "tool_group": {
      const group = document.createElement("details");
      group.className = "tool-group";
      const summary = document.createElement("summary");
      const label = document.createElement("span");
      label.textContent = toolGroupLabel(message.tools?.length ?? 0);
      const chevron = document.createElement("span");
      chevron.className = "tool-group-chevron";
      chevron.append(toolChevron());
      summary.append(label, chevron);
      const list = document.createElement("div");
      list.className = "tool-list";
      for (const tool of message.tools ?? []) list.append(renderToolCard(tool));
      group.append(summary, list);
      element = group;
      break;
    }
  }
  if (container === els.messages && thinking.element) container.insertBefore(element, thinking.element);
  else container.append(element);
  return element;
}


export function toolGroupLabel(count: number): string {
  return `${count} tool${count === 1 ? "" : "s"} called`;
}

/** One disclosure per provider turn, shared by native and workbook tool calls. */
export function appendTurnTool(tool: StoredTool): HTMLElement {
  if (!app.liveToolGroup) {
    const message: StoredMessage = { kind: "tool_group", text: "", tools: [] };
    app.liveToolGroup = { element: append(message) as HTMLDetailsElement, message };
  }
  app.liveToolGroup.message.tools!.push(tool);
  app.liveToolGroup.element.querySelector("summary > span")!.textContent = toolGroupLabel(app.liveToolGroup.message.tools!.length);
  const card = renderToolCard(tool);
  app.liveToolGroup.element.querySelector(".tool-list")!.append(card);
  schedulePersist();
  return card;
}

export function bubble(kind: "user" | "assistant" | "error"): HTMLDivElement {
  const element = document.createElement("div");
  element.className = `bubble ${kind}`;
  return element;
}


/** Add a message to the current chat, show it and save. */
export function append(message: StoredMessage): HTMLElement {
  app.session.messages.push(message);
  const element = renderMessage(els.messages, message);
  scrollToEnd(message.kind === "user");
  schedulePersist();
  return element;
}

export function renderWelcome(): void {
  const box = document.createElement("div");
  box.className = "welcome";
  const illustration = document.createElement("img");
  illustration.className = "welcome-sheet";
  illustration.src = "../assets/spreadsheet-chalk.svg";
  illustration.alt = "Hand-drawn spreadsheet";
  illustration.draggable = false;
  box.append(illustration);
  els.messages.append(box);
}

export function renderDelivery(element: HTMLElement, message: StoredMessage): void {
  element.querySelector(".message-delivery")?.remove();
  if (!message.delivery) return;
  const label = document.createElement("span");
  label.className = "message-delivery";
  const pending = message.queueId ? pendingPrompts.get(message.queueId) : undefined;
  label.textContent = message.delivery === "pending" && pending?.state !== "held" && pending ? "Pending" : "Not sent";
  if (pending?.reason) label.title = pending.reason;
  if (!pending || pending.state === "held") {
    const send = document.createElement("button");
    send.type = "button";
    send.textContent = "Send";
    send.addEventListener("click", () => {
      // A held message goes out as the next reply, so it cannot be sent while one is still running.
      if (app.currentRequestId) { showToast("A reply is still running. Wait for it to finish (or press Stop), then press Send."); return; }
      if (!app.client.connected) { showToast("Not connected to the bridge. Try again after it reconnects."); return; }
      void submit(pending ?? { id: message.queueId ?? randomId(), parentRequestId: "", prompt: message.text, attachments: [], message, state: "held" });
    });
    label.append(send);
  }
  element.append(label);
  if (message.queueId) element.dataset.queueId = message.queueId;
}

export function updateDelivery(item: PendingPrompt): void {
  item.message.delivery = item.state === "held" ? "held" : "pending";
  for (const element of els.messages.querySelectorAll<HTMLElement>(".bubble.user[data-queue-id]")) {
    if (element.dataset.queueId === item.id) renderDelivery(element, item.message);
  }
  schedulePersist();
}

export function appendFileMentions(container: HTMLElement, text: string, names: readonly string[]): void {
  let cursor = 0;
  for (const range of fileMentionRanges(text, names)) {
    container.append(document.createTextNode(text.slice(cursor, range.start)));
    const badge = document.createElement("span");
    badge.className = "file-mention";
    badge.textContent = text.slice(range.start, range.end);
    container.append(badge);
    cursor = range.end;
  }
  container.append(document.createTextNode(text.slice(cursor)));
}

export function startLive(kind: "assistant", text: string): { element: HTMLElement; message: StoredMessage } {
  const message: StoredMessage = { kind, text, ...(app.turnTouchedOtherFiles ? { aboutOtherFiles: true } : {}) };
  app.session.messages.push(message);
  return { element: renderMessage(els.messages, message), message };
}

/** Draw the current chat's messages again, where they were. Only between requests: a running one has live parts on screen. */
export function redrawMessages(): void {
  const top = els.messages.scrollTop;
  els.messages.replaceChildren();
  for (const message of app.session.messages) renderMessage(els.messages, message);
  els.messages.scrollTop = top;
}

/** Streaming text follows the end of the chat only while the person is reading the end; `force` for what they asked for. */
export function scrollToEnd(force = false): void {
  follow.scroll(force);
}

export function summarize(value: unknown): string {
  try {
    const text = JSON.stringify(value);
    return !text || text === "{}" ? "" : text.length > 70 ? `${text.slice(0, 67)}...` : text;
  } catch {
    return "";
  }
}
