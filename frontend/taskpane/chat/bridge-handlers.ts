import { byId } from "../ui/dom.js";
import { showUserInput } from "../ui/user-input.js";
import { renderMarkdown } from "../ui/markdown.js";
import { showToast } from "../ui/toast.js";
import { formatToolValue } from "../ui/tool-view.js";
import { BridgeClient } from "../bridge-client.js";
import { browserView, connectedFiles, ensureWorkspaces, historyNotice, mcpSettings, navigation, providerPicker, recordChange, refreshModelCatalog, refreshUsage, resetWorkspaces, slash, store, thinking, updatePanel, updater } from "./services.js";
import { append, appendTurnTool, cellLinkOptions, redrawMessages, scrollToEnd, summarize } from "./messages.js";
import { handleEvent, renderState } from "./events.js";
import { loadSession } from "./chats.js";
import { publishWorkbookName, readWorkbookName, refreshSheetNames } from "./workbook.js";
import { submit } from "./turn.js";
import { prepareRuntime } from "./prepare.js";
import { app } from "./state.js";

export function createBridgeClient(): BridgeClient {
  return new BridgeClient({
    onUserInput: async (request, signal, requestId) => {
      if (requestId !== app.currentRequestId || signal.aborted) return { cancelled: true, answers: {} };
      thinking.clear();
      app.live = null;
      const questionText = request.questions.map(q => q.question).join("\n\n");
      const questionBubble = append({ kind: "assistant", text: questionText });
      const owner = app.session;
      const dock = byId("question-dock");
      dock.hidden = false;
      document.body.classList.add("question-active");
      const pending = showUserInput(dock, request, signal);
      scrollToEnd(true);
      const response = await pending;
      dock.hidden = true;
      document.body.classList.remove("question-active");
      if (app.session === owner) questionBubble.innerHTML = renderMarkdown(questionText, { cellLinks: cellLinkOptions(questionText) });
      if (app.session === owner && requestId === app.currentRequestId) {
        append({ kind: response.cancelled ? "note" : "user", text: response.cancelled ? "Question skipped or cancelled." : request.questions.map(q =>
          `${q.header || q.question}: ${q.isSecret ? "[Hidden answer]" : response.answers[q.id]!.join(", ")}`).join("\n") });
        if (!signal.aborted) thinking.show();
      }
      return response;
    },
    onState: renderState,
    onHello: ({ addinBuildId }) => {
      void store.connect().then(async () => {
        if (app.restoreChatId && !app.currentRequestId) {
          const id = app.restoreChatId; app.restoreChatId = null;
          const restored = await store.loadActive(id); if (restored && !app.currentRequestId) loadSession(restored);
        }
      }).catch(error => historyNotice(`History is unavailable: ${error.message}. Trying again automatically.`));
      void mcpSettings.restore();
      void publishWorkbookName();
      void refreshSheetNames();
      slash.reload();
      app.client.setProvider(app.provider);
      void refreshUsage();
      void refreshModelCatalog(app.provider);
      if (app.session.messages.length) app.needsContext = true;
      app.activeModel = undefined;
      resetWorkspaces();
      void ensureWorkspaces();
      updater.check(addinBuildId);
      void updatePanel.afterReload();
    },
    onProviderStatus: ({ provider: id, detection }) => {
      providerPicker.onStatus(id, detection);
      // The bridge answers every provider choice (connecting, switching, signing in) with its status: once the CLI is usable, open it.
      if (id === app.provider && detection.status === "ready") prepareRuntime();
    },
    onProviderEvent: handleEvent,
    onBrowserFrame: (frame) => { if (frame.requestId === app.currentRequestId) browserView.push(frame); },
    readWorkbookName,
    onConnectedFiles: (files) => {
      // Only workbooks count here: this list decides which cell references in a reply can be linked to another file.
      const names = files.filter((file) => !file.current && file.host === "excel").map((file) => file.name).filter(Boolean);
      const changed = names.join("\n") !== app.otherFiles.join("\n");
      app.otherFiles = names;
      connectedFiles.render(files);
      // Whether a reply's bare cells are linked depends on which other files are open (a chat opened before this list arrived, for one).
      if (changed && !app.currentRequestId && app.session.messages.length) redrawMessages();
    },
    // Another file's chat asked this workbook to show a cell (a link there): only the selection moves.
    goTo: target => navigation.selectRange(target),
    onDelegatedPrompt: (request) => {
      // The other assistant's request is a turn like any other here, so it waits its turn: refuse while this chat is busy.
      if (app.currentRequestId || app.cancelRequested || app.resolvingCommand) {
        app.client.declineDelegation(request.delegationId, "That file is busy with its own request.");
        return;
      }
      void submit(undefined, { delegationId: request.delegationId, from: request.from.name, fromHost: request.from.host ?? "excel", text: request.text });
    },
    onError: (message) => showToast(message),
    onToolExecuted: (tool, args, result, context) => {
      if (context.requestId === app.currentRequestId) thinking.setLabel("Working…");
      if (tool === "create_artifact") {
        // The card is the visible result; a second raw tool entry would only repeat the file contents.
        const file = args as { filename?: unknown; content?: unknown };
        if (result.ok && context.requestId === app.currentRequestId && app.session === app.requestSession && typeof file.filename === "string" && typeof file.content === "string") {
          app.liveToolGroup = null;
          append({ kind: "artifact", text: "", artifact: { name: file.filename, content: file.content } });
          scrollToEnd();
        }
        return;
      }
      if (context.requestId === app.currentRequestId && app.session === app.requestSession) appendTurnTool({ name: tool, args: summarize(args), input: formatToolValue(args),
        result: formatToolValue(result), state: result.ok ? "ok" : "error", error: result.ok ? null : result.error.message });
      recordChange(tool, args, result, context);
      if (result.ok && ["add_shape", "add_connector", "add_line", "update_shape", "delete_shapes", "execute_office_js", "manage_sheet", "insert_sheets_from_file"].includes(tool)) {
        app.selectionChip?.refresh();
      }
    },
  });
}
