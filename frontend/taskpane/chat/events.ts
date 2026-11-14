import type { ProviderEvent } from "../../../protocol/src/index.js";
import { showToast } from "../ui/toast.js";
import { bindProviderSession, type StoredTool } from "../ui/sessions.js";
import { renderMarkdown } from "../ui/markdown.js";
import { formatToolValue, paintTool } from "../ui/tool-view.js";
import { activeHost } from "../host/index.js";
import { app, liveTools, pendingPrompts } from "./state.js";
import { els } from "../ui/elements.js";
import { append, appendTurnTool, cellLinkOptions, isDocumentTool, renderDelivery, renderMessage, scrollToEnd, startLive, summarize, updateDelivery } from "./messages.js";
import { persistNow, schedulePersist } from "./chats.js";
import { holdPendingPrompts, sendNextPending, setBusy } from "./turn.js";
import { connectedFiles, providerOptions, providerPicker, refreshUsage, renderChanges, renderUsage, slash, store, thinking, thoughts, updater, usagePanel, userSkills } from "./services.js";
import { refreshSheetNames } from "./workbook.js";

export const CONNECTED_FILE_TOOLS: ReadonlySet<string> = new Set(["list_connected_files", "read_connected_file", "send_to_connected_file"]);
/** A tool's name as words for the status line: `read_range` and `excel-ai-addin_read_range` both become "read range". */
const toolLabel = (name: string): string => name.replace(/^mcp__[^_]+__/, "").replace(/^excel-ai-addin[_-]/, "").replace(/[_-]+/g, " ").slice(0, 40);

export function handleEvent(event: ProviderEvent): void {
  if (event.type !== "followup_status" && event.requestId !== null && event.requestId !== app.currentRequestId) return;
  switch (event.type) {
    case "followup_status": {
      const item = pendingPrompts.get(event.id);
      if (!item || item.parentRequestId !== event.requestId) return;
      if (item.state === "held" && event.state !== "delivered") return;
      if (event.state === "delivered") {
        item.message.delivery = undefined;
        pendingPrompts.delete(item.id);
        if (app.session.messages.includes(item.message)) {
          for (const element of els.messages.querySelectorAll<HTMLElement>(".bubble.user[data-queue-id]")) {
            if (element.dataset.queueId === item.id) renderDelivery(element, item.message);
          }
        }
        schedulePersist();
      } else {
        item.state = event.state === "rejected" ? "held" : event.state;
        if (!app.currentRequestId && !app.autoQueueAllowed) item.state = "held";
        if (event.state === "rejected" && event.message) item.reason = event.message;
        updateDelivery(item);
        if (event.message) showToast(event.message);
        if (event.state === "next_turn" && app.autoQueueAllowed) sendNextPending();
      }
      return;
    }
    case "session_info": {
      if (event.requestId !== app.currentRequestId || app.session !== app.requestSession) return;
      const oldId = app.session.id;
      const bound = bindProviderSession(app.session, event, app.requestMessageIndex, app.requestChangeIndex, app.requestStartedAt);
      if (bound.previous) {
        store.save(bound.previous); // Removes a provisional empty history item, preserves older sessions.
        app.requestMessageIndex = 0;
        app.requestChangeIndex = 0;
      }
      app.session = bound.current;
      app.requestSession = app.session;
      if (oldId !== app.session.id) {
        usagePanel.rebind(event.requestId);
        // Metadata precedes output, so redraw only when a request moves to another native session.
        thinking.clear();
        els.messages.replaceChildren();
        for (const message of app.session.messages) renderMessage(els.messages, message);
        thinking.show();
        renderChanges();
      }
      persistNow();
      return;
    }
    case "session_starting":
      thinking.setLabel("Starting session…");
      return;
    case "provider_options":
      if (event.requestId === app.currentRequestId) providerOptions.render(event.options);
      return;
    case "provider_progress":
      if (event.requestId === app.currentRequestId) thinking.setLabel(event.message);
      return;
    case "turn_started":
      if (event.requestId !== app.currentRequestId) return;
      void refreshSheetNames();
      thinking.setLabel("Working…");
      usagePanel.start(event.requestId);
      return;
    case "usage_update":
      usagePanel.update(event.requestId, event.usage);
      return;
    case "thinking_delta":
      thinking.setLabel("Thinking…");
      thoughts.add(event.text);
      return;
    case "message_delta":
      if (!event.text || (!app.live && !event.text.trim())) return;
      thinking.setLabel("Writing…");
      thoughts.end();
      if (!app.live) app.live = startLive("assistant", event.text);
      else app.live.message.text += event.text;
      app.live.element.innerHTML = renderMarkdown(app.live.message.text, { cellLinks: cellLinkOptions(app.live.message.text, app.live.message.aboutOtherFiles) });
      scrollToEnd();
      schedulePersist();
      return;
    case "message_completed":
      // Providers can finish an assistant message that only calls a tool.
      // It has no chat text and must not create (or erase) a visible bubble.
      if (!event.text.trim()) { app.live = null; return; }
      thoughts.end();
      if (!app.live) app.live = startLive("assistant", event.text);
      app.live.message.text = event.text;
      app.live.element.innerHTML = renderMarkdown(event.text, { cellLinks: cellLinkOptions(event.text, app.live.message.aboutOtherFiles) });
      {
        // The reply may name a sheet the assistant just made: read the names again and draw it once more if they changed.
        const { element, message } = app.live;
        void refreshSheetNames(false).then((changed) => {
          if (changed && element.isConnected) element.innerHTML = renderMarkdown(message.text, { cellLinks: cellLinkOptions(message.text, message.aboutOtherFiles) });
        });
      }
      app.live = null; // the next message (after a tool call) gets its own bubble
      scrollToEnd();
      schedulePersist();
      return;
    case "tool_started": {
      thoughts.end();
      app.live = null;
      thinking.setLabel(`Using ${toolLabel(event.tool)}…`);
      if (CONNECTED_FILE_TOOLS.has(event.tool)) app.turnTouchedOtherFiles = true;
      // Workbook calls get their exact result from onToolExecuted. Provider-local
      // tools have a separate native call ID and must not be paired by name.
      if (isDocumentTool(event.tool)) return;
      const tool: StoredTool = { name: event.tool, args: summarize(event.args), input: formatToolValue(event.args), state: "running", error: null };
      const card = appendTurnTool(tool);
      liveTools.set(event.callId, { card, tool });
      schedulePersist();
      return;
    }
    case "tool_finished": {
      thinking.setLabel("Working…");
      const entry = liveTools.get(event.callId);
      if (!entry) return;
      entry.tool.state = event.ok ? "ok" : "error";
      entry.tool.error = event.error;
      paintTool(entry.card, entry.tool);
      liveTools.delete(event.callId);
      schedulePersist();
      return;
    }
    case "turn_finished":
      if (event.requestId !== app.currentRequestId) return;
      app.autoQueueAllowed = event.reason === "completed";
      usagePanel.finish(event.requestId);
      void refreshUsage();
      slash.reload();
      void userSkills.refresh();
      thinking.clear();
      app.currentRequestId = null;
      app.live = null;
      for (const { card, tool } of liveTools.values()) {
        tool.state = "error";
        tool.error = "no result";
        paintTool(card, tool);
      }
      liveTools.clear();
      if (app.liveToolGroup) app.liveToolGroup.element.open = false;
      app.liveToolGroup = null;
      setBusy(false);
      if (event.reason === "cancelled") append({ kind: "note", text: "Stopped." });
      if (event.reason !== "completed") holdPendingPrompts();
      persistNow();
      void store.flush().catch(() => {});
      sendNextPending();
      updater.reloadIfIdle();
      return;
    case "error":
      thinking.clear();
      if (event.code === "login_required") {
        providerPicker.requireLogin(app.provider);
      }
      showToast(event.message);
      return;
    default:
      return;
  }
}

export function renderState(state: string): void {
  connectedFiles.setConnected(state === "open");
  els.statusLabel.textContent = ({ open: `Connected to ${activeHost().app}`, connecting: "Connecting…", closed: "Reconnecting…", idle: `Waiting for ${activeHost().app}` } as Record<string, string>)[state] ?? state;
  els.status.dataset.state = state;
  els.send.disabled = state !== "open" || els.send.dataset.stopping === "true";
  if (state === "closed" && app.currentRequestId) {
    app.autoQueueAllowed = false;
    holdPendingPrompts();
    // The bridge dropped mid-turn; the running request can no longer finish, so free the UI.
    usagePanel.reset();
    renderUsage();
    app.currentRequestId = null;
    app.live = null;
    setBusy(false);
    showToast("Connection to the bridge was lost; the running request was stopped. Reconnecting…");
  }
}
