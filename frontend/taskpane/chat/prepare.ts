import { els } from "../ui/elements.js";
import { chosenModel, toolSettings } from "./services.js";
import { app } from "./state.js";

/** A burst of changes (a new chat, then a model, then a tool) is one message to the bridge. */
const SETTLE_MS = 250;
let timer: ReturnType<typeof setTimeout> | undefined;

/**
 * Tell the bridge what the next message will run with, so the provider's CLI is already open when it is sent.
 * Call it whenever the connection, the provider, the chat, the model, the effort or a tool changes; it costs nothing
 * when the CLI is already open that way.
 */
export function prepareRuntime(): void {
  clearTimeout(timer);
  timer = setTimeout(() => {
    if (!app.client?.connected || app.currentRequestId) return;
    app.client.prepare({
      model: chosenModel(),
      effort: els.effortSelect.value as "low" | "medium" | "high",
      chatSessionId: app.session.id,
      resumeSessionId: app.session.provider === app.provider ? app.session.providerSessionId : undefined,
      capabilities: toolSettings.value(),
    });
  }, SETTLE_MS);
}
