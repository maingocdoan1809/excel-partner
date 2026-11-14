import { slashBadgeRange } from "../ui/slash.js";
import { fileMentionRanges } from "../ui/file-mentions.js";
import type { PermissionMode } from "../permissions.js";
import { byId } from "../ui/dom.js";
import { els } from "../ui/elements.js";
import { appendFileMentions } from "./messages.js";
import { PERMISSION_MODE_KEY, app, pendingAttachments } from "./state.js";

export function renderComposerPreview(): void {
  updateSendControl();
  els.promptDisplay.replaceChildren();
  const text = els.input.value;
  const command = slashBadgeRange(text, app.selectedSlashToken);
  if (command) {
    const badge = document.createElement("span");
    badge.className = "slash-token";
    badge.textContent = text.slice(command.start, command.end);
    els.promptDisplay.append(badge);
    appendFileMentions(els.promptDisplay, text.slice(command.end), pendingAttachments.map((item) => item.name));
  } else appendFileMentions(els.promptDisplay, text, pendingAttachments.map((item) => item.name));
  // A textarea keeps a caret line after a trailing newline; mirror that line too.
  if (text.endsWith("\n")) els.promptDisplay.append(document.createTextNode("\u200b"));
  els.promptDisplay.scrollTop = els.input.scrollTop;
}

export function syncSelectedSlashToken(): void {
  if (app.selectedSlashToken && !(
    els.input.value === app.selectedSlashToken ||
    els.input.value.startsWith(`${app.selectedSlashToken} `) ||
    els.input.value.startsWith(`${app.selectedSlashToken}\n`)
  )) app.selectedSlashToken = null;
}

export function syncAttachmentsWithInput(): void {
  const ranges = fileMentionRanges(els.input.value, pendingAttachments.map((item) => item.name));
  const counts = new Map<string, number>();
  for (const range of ranges) counts.set(range.name, (counts.get(range.name) ?? 0) + 1);
  for (let i = pendingAttachments.length - 1; i >= 0; i--) {
    const name = pendingAttachments[i]!.name;
    const left = counts.get(name) ?? 0;
    if (left > 0) counts.set(name, left - 1);
    else pendingAttachments.splice(i, 1);
  }
}

export function clearAttachment(): void {
  pendingAttachments.length = 0;
  app.selectedSlashToken = null;
  renderComposerPreview();
}

export function selectPermissionMode(mode: PermissionMode): void {
  if (mode !== "ask" && mode !== "auto") return;
  app.permissionMode = mode;
  localStorage.setItem(PERMISSION_MODE_KEY, mode);
  app.client.setPermissionMode(mode);
  els.permissionMode.value = mode;
  app.permissionPicker?.refresh();
}

export function updateSendControl(): void {
  const running = app.currentRequestId !== null;
  const hasPrompt = !!els.input.value.trim() || pendingAttachments.length > 0;
  const stop = running && !hasPrompt;
  els.send.dataset.mode = stop ? "stop" : "send";
  const label = stop ? "Stop response" : running ? "Queue message" : "Send message";
  els.sendLabel.textContent = stop ? "Stop" : running ? "Queue" : "Send";
  els.send.setAttribute("aria-label", label);
  els.send.title = label;
  byId("stop-button").hidden = !running || !hasPrompt;
}
