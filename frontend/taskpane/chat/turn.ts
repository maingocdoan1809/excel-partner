import { parseNativeCommand, type ProviderCommand } from "../../../protocol/src/commands.js";
import { byId } from "../ui/dom.js";
import { PROVIDER_NAMES } from "../ui/provider-picker.js";
import { delegatedPrompt } from "../ui/delegation.js";
import { showToast } from "../ui/toast.js";
import { randomId } from "../bridge-client.js";
import { buildContextPreface, type StoredMessage } from "../ui/sessions.js";
import { activeHost } from "../host/index.js";
import { type PendingPrompt, app, pendingAttachments, pendingPrompts } from "./state.js";
import { append, renderDelivery, updateDelivery } from "./messages.js";
import { persistNow, startNewChat } from "./chats.js";
import { browserView, mcpSettings, providerOptions, openMenu, refreshModelCatalog, renderModelPicker, saveModelChoice, showPanel, slash, thinking, toolSettings, updateMentionMenu, updater, usagePanel } from "./services.js";
import { els } from "../ui/elements.js";
import { clearAttachment, renderComposerPreview, selectPermissionMode, updateSendControl } from "./composer.js";

export function holdPendingPrompts(): void {
  for (const item of pendingPrompts.values()) { item.state = "held"; updateDelivery(item); }
}

export function sendNextPending(): void {
  if (app.currentRequestId || !app.client.connected) return;
  const next = [...pendingPrompts.values()].find(item => item.state === "next_turn" && app.session.messages.includes(item.message));
  if (next) void submit(next);
}

export async function runClientCommand(command: ProviderCommand, args: string): Promise<void> {
  switch (command.action) {
    case "new": startNewChat(); break;
    case "history": showPanel("history"); break;
    case "model":
      if (args) {
        await refreshModelCatalog(app.provider);
        if (!Array.from(els.modelSelect.options).some(option => option.value === args)) throw new Error(`Unknown model: ${args}`);
        els.modelSelect.value = args; saveModelChoice(); renderModelPicker();
      } else byId<HTMLButtonElement>("model-button").click();
      break;
    case "effort": byId<HTMLButtonElement>("effort-button").click(); break;
    case "permissions":
      if (args === "ask" || args === "auto") selectPermissionMode(args);
      else if (args) throw new Error("Use /permissions ask or /permissions auto.");
      else byId<HTMLButtonElement>("permission-button").click();
      break;
    case "usage": byId<HTMLButtonElement>("usage-button").click(); break;
    case "login": showPanel("login"); break;
    case "theme": {
      const select = byId<HTMLSelectElement>("theme-select");
      if (args) {
        if (!["light", "dark", "auto"].includes(args)) throw new Error("Use /theme light, dark, or auto.");
        select.value = args; select.dispatchEvent(new Event("change"));
      } else { openMenu(); select.focus(); }
      break;
    }
    case "copy": {
      const answer = [...app.session.messages].reverse().find(message => message.kind === "assistant");
      if (!answer) throw new Error("There is no assistant response to copy yet.");
      await navigator.clipboard.writeText(answer.text);
      append({ kind: "note", text: "Copied the last response." });
      break;
    }
    case "export": {
      const markdown = app.session.messages.filter(message => message.kind !== "tool_group").map(message => `## ${message.kind}\n\n${message.text}`).join("\n\n");
      const url = URL.createObjectURL(new Blob([markdown], { type: "text/markdown;charset=utf-8" }));
      const link = document.createElement("a"); link.href = url; link.download = "excel-partner-chat.md";
      document.body.append(link); link.click(); link.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      break;
    }
    case "help": append({ kind: "note", text: slash.commands.map(item => `/${item.name} — ${item.description}${item.command?.execution === "unsupported" ? " (CLI only)" : ""}`).join("\n") }); break;
    case "status": append({ kind: "note", text: `Provider: ${PROVIDER_NAMES[app.provider]}\nModel: ${els.modelSelect.value || "Default"}\nSession: ${app.session.providerSessionId || "New chat"}\nPermissions: ${app.permissionMode}` }); break;
  }
}

/** `delegated`: another open workbook's assistant sent this request; it runs as a normal turn here and its reply goes back. */
export async function submit(queued?: PendingPrompt, delegated?: { delegationId: string; from: string; fromHost?: "excel" | "word" | "powerpoint"; text: string }): Promise<void> {
  // Whoever sent a delegated request waits for an answer: every way out of here that does not start the turn must say so.
  const decline = (reason: string) => { if (delegated) app.client.declineDelegation(delegated.delegationId, reason); };
  const prompt = delegated ? delegatedPrompt(delegated.from, delegated.text)
    : queued?.prompt ?? (els.input.value.trim() || (pendingAttachments.length ? "Please inspect the attached files." : ""));
  if (!prompt || !app.client.connected) { decline("That file is not connected right now."); return; }
  if (prompt.length > 32_000) { decline("The request is too long for that file."); showToast("Messages must be 32,000 characters or fewer."); return; }
  let nativeCommand = false;
  if (parseNativeCommand(prompt)) {
    if (app.resolvingCommand) return;
    const owner = app.session, selectedProvider = app.provider;
    app.resolvingCommand = true;
    try {
      const item = await slash.resolve(prompt);
      if (app.session !== owner || app.provider !== selectedProvider || (!queued && els.input.value.trim() !== prompt)) return;
      if (!item) throw new Error("Unknown command. Open the / menu to see available commands and skills.");
      if (item.command?.execution === "unsupported") throw new Error(item.command.reason ?? "This command requires the CLI terminal.");
      if (item.command?.execution === "client" || item.kind === "builtin") {
        await runClientCommand(item.command ?? { name: item.name, description: item.description, source: "builtin", execution: "client", action: "new" }, parseNativeCommand(prompt)!.args);
        els.input.value = ""; slash.close(); renderComposerPreview(); return;
      }
      nativeCommand = item.kind === "provider";
      if (nativeCommand && app.currentRequestId) throw new Error("Wait for the current response to finish before running a provider command.");
      if (nativeCommand && pendingAttachments.length) throw new Error("Send attached files separately from a provider command.");
    } catch (error) { showToast(error instanceof Error ? error.message : String(error)); return; }
    finally { app.resolvingCommand = false; }
  }
  const turnAttachments = delegated ? [] as typeof pendingAttachments : queued?.attachments ?? pendingAttachments.slice();
  if (!delegated && !turnAttachments.length && /(?:^|\s)@[^\n]*\.(?:xlsx|xlsm|xltx|xltm|docx|docm|pptx|pptm|png|jpe?g|gif|webp|md|txt|json|csv|ya?ml|zip)(?=\s|$)/i.test(prompt)) {
    showToast("The @ file is only text. Choose it from a workspace or the local file picker so I can read it.");
    updateMentionMenu();
    return;
  }
  if (app.currentRequestId) {
    if (delegated) { decline("That file is busy with its own request."); return; }
    if (queued || app.cancelRequested) return;
    if (pendingPrompts.size >= 20) { showToast("Wait for pending messages to be sent before adding more."); return; }
    const id = randomId();
    const attachments = pendingAttachments.slice();
    const message: StoredMessage = { kind: "user", text: prompt, delivery: "pending", queueId: id, attachments: attachments.map(item => ({ name: item.name })) };
    const item: PendingPrompt = { id, parentRequestId: app.currentRequestId, prompt, attachments, message, state: "pending" };
    pendingPrompts.set(id, item);
    try {
      app.client.queuePrompt(app.currentRequestId, id, prompt,
        attachments.flatMap(item => item.reference.id ? [] : [item.reference.path]),
        attachments.flatMap(item => item.reference.id ? [item.reference.id] : []));
      append(message);
      els.input.value = "";
      clearAttachment();
      persistNow();
    } catch (error) { pendingPrompts.delete(id); showToast(error instanceof Error ? error.message : String(error)); }
    return;
  }
  if (!app.session.identityVersion && !app.session.providerSessionId && app.session.messages.some(message => message.kind === "user")) {
    decline("The chat open in that file is an older chat that cannot take requests. Ask the user to start a new chat in that window.");
    showToast("This legacy chat has no saved CLI session ID, so it cannot be resumed reliably. Its history is preserved. Start a new chat to create a native session.");
    return;
  }
  const model = els.modelSelect.value || null;
  const effort = els.effortSelect.value as "low" | "medium" | "high";
  saveModelChoice();
  slash.close();
  if (prompt === "/clear") {
    els.input.value = "";
    startNewChat();
    return;
  }

  const requestId = randomId();
  if (app.liveToolGroup) app.liveToolGroup.element.open = false;
  app.liveToolGroup = null;
  app.currentRequestId = requestId;
  app.requestSession = app.session;
  app.turnTouchedOtherFiles = false;
  browserView.reset();
  setBusy(true);
  if (!queued && !delegated) els.input.value = "";
  renderComposerPreview();
  showPanel(null);
  const context = !nativeCommand && (app.needsContext || (app.activeModel !== undefined && app.activeModel !== model) || (app.activeEffort !== undefined && app.activeEffort !== effort))
    ? buildContextPreface(app.session) : undefined;
  if (!nativeCommand) app.needsContext = false;
  app.requestMessageIndex = queued ? Math.max(0, app.session.messages.indexOf(queued.message)) : app.session.messages.length;
  app.requestChangeIndex = app.session.changes.length;
  app.requestStartedAt = Date.now();
  if (queued) {
    queued.message.delivery = undefined;
    pendingPrompts.delete(queued.id);
    for (const element of els.messages.querySelectorAll<HTMLElement>(".bubble.user[data-queue-id]")) {
      if (element.dataset.queueId === queued.id) renderDelivery(element, queued.message);
    }
  } else append(delegated ? { kind: "user", text: delegated.text, from: delegated.from, fromHost: delegated.fromHost }
    : { kind: "user", text: prompt, attachments: turnAttachments.map((item) => ({ name: item.name })) });
  if (!queued && !delegated) clearAttachment();
  thinking.show();
  persistNow();

  const result = nativeCommand ? null : await activeHost().execute("get_selection", {});
  if (app.currentRequestId !== requestId || app.requestSession !== app.session) { decline("The chat open in that file changed before the request could start."); return; }
  const selection = result?.ok ? (result.data as never) : null;
  if (app.cancelRequested) return finishLocalCancel(requestId, delegated?.delegationId);
  try {
    await mcpSettings.ready();
    const attachmentIds = turnAttachments.flatMap((item) => item.reference.id ? [item.reference.id] : []);
    const referencePaths = turnAttachments.flatMap((item) => item.reference.id ? [] : [item.reference.path]);
    if (app.cancelRequested) return finishLocalCancel(requestId, delegated?.delegationId);
    app.client.sendPrompt(requestId, prompt, selection, context, attachmentIds, model, referencePaths, effort, app.session.id,
      app.session.provider === app.provider ? app.session.providerSessionId : undefined, toolSettings.value(), nativeCommand || undefined, delegated?.delegationId);
    app.promptDispatched = true;
    app.activeModel = model;
    app.activeEffort = effort;
  } catch (error) {
    if (app.cancelRequested) return finishLocalCancel(requestId, delegated?.delegationId);
    decline("The request could not be sent in that file.");
    thinking.clear();
    app.currentRequestId = null;
    setBusy(false);
    if (queued) { queued.state = "held"; pendingPrompts.set(queued.id, queued); updateDelivery(queued); }
    else if (!delegated) els.input.value = prompt;
    renderComposerPreview();
    showToast(error instanceof Error ? error.message : String(error));
  }
}

export function stopCurrentTurn(): void {
  if (!app.currentRequestId || app.cancelRequested) return;
  app.cancelRequested = true;
  app.autoQueueAllowed = false;
  els.send.dataset.stopping = "true";
  els.send.disabled = true;
  els.send.setAttribute("aria-label", "Stopping response");
  els.send.title = "Stopping response";
  if (app.promptDispatched) app.client.cancel(app.currentRequestId);
}

/** `delegationId`: the stopped turn answered another file's request, which is told it will not be done. */
export function finishLocalCancel(requestId: string, delegationId?: string): void {
  if (app.currentRequestId !== requestId) return;
  if (delegationId) app.client.declineDelegation(delegationId, "The user stopped the request in that file before it started.");
  thinking.clear();
  app.currentRequestId = null;
  holdPendingPrompts();
  const last = app.session.messages[app.session.messages.length - 1];
  if (last?.kind === "user" && pendingPrompts.size === 0) {
    app.session.messages.pop();
    if (last.from !== undefined) {
      // Another assistant's request is not the person's own text: drop its block and leave the composer alone.
      const received = els.messages.querySelectorAll(".received");
      received[received.length - 1]?.remove();
    } else {
      const bubbles = els.messages.querySelectorAll(".bubble.user");
      bubbles[bubbles.length - 1]?.remove();
      els.input.value = last.text;
      renderComposerPreview();
    }
  }
  app.needsContext = app.session.messages.length > 0;
  setBusy(false);
  append({ kind: "note", text: "Stopped." });
  persistNow();
  updater.reloadIfIdle();
}

/** The send control becomes a red Stop button while a turn is running. */
export function setBusy(busy: boolean): void {
  if (busy) app.autoQueueAllowed = true;
  if (!busy) thinking.clear();
  app.cancelRequested = false;
  app.promptDispatched = false;
  els.send.dataset.mode = busy ? "stop" : "send";
  els.send.dataset.stopping = "false";
  els.send.disabled = els.status.dataset.state !== "open";
  els.sendLabel.textContent = busy ? "Stop" : "Send";
  els.send.setAttribute("aria-label", busy ? "Stop response" : "Send message");
  els.send.title = busy ? "Stop response" : "Send message";
  els.status.dataset.busy = String(busy);
  providerOptions.sync();
  updateSendControl();
}

export function abandonRequest(): void {
  app.autoQueueAllowed = false;
  if (app.currentRequestId) app.client.cancel(app.currentRequestId);
  holdPendingPrompts();
  pendingPrompts.clear();
  app.currentRequestId = null;
  app.requestSession = null;
  usagePanel.reset();
  setBusy(false);
}
