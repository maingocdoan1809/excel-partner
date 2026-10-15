import type { BridgeClient } from "../bridge-client.js";
import type { ProviderLoginStatus } from "../../../protocol/src/provider-login.js";

export function setupProviderLogin(options: {
  client: () => BridgeClient;
  onSuccess: () => void;
  isBusy: () => boolean;
}): { open(): void; close(): void } {
  const status = document.getElementById("login-status")!;
  const code = document.getElementById("login-code")!;
  const codeCard = document.getElementById("login-code-card")!;
  const timer = document.getElementById("login-timer")!;
  const start = document.getElementById("login-start") as HTMLButtonElement;
  const copy = document.getElementById("login-copy") as HTMLButtonElement;
  const browser = document.getElementById("login-browser") as HTMLAnchorElement;
  const cancel = document.getElementById("login-cancel") as HTMLButtonElement;
  const proceed = document.getElementById("login-continue") as HTMLButtonElement;
  let poll: number | undefined;
  let generation = 0;
  let active = false;
  let pending = false;

  const render = (value: ProviderLoginStatus) => {
    status.textContent = value.message;
    status.dataset.state = value.state;
    const waiting = value.state === "waiting";
    const running = waiting || value.state === "starting";
    codeCard.hidden = !waiting;
    code.textContent = value.userCode ?? "";
    browser.hidden = !waiting;
    browser.removeAttribute("href");
    if (value.verificationUri === "https://github.com/login/device") browser.href = value.verificationUri;
    copy.textContent = "Copy code";
    copy.disabled = !waiting;
    start.hidden = running || value.state === "success";
    start.textContent = value.state === "idle" ? "Get device code →" : "Get a new code →";
    cancel.hidden = !running;
    proceed.hidden = value.state !== "success";
    timer.textContent = value.expiresAt && running
      ? `Session timeout in up to ${Math.max(0, Math.ceil((value.expiresAt - Date.now()) / 60_000))} min` : "";
    return running;
  };

  const request = async (method: "GET" | "POST" | "DELETE" = "GET") => {
    if (pending) return;
    const currentGeneration = generation;
    pending = true;
    window.clearTimeout(poll);
    start.disabled = true;
    cancel.disabled = true;
    try {
      const value = await options.client().providerLogin(method);
      if (!active || currentGeneration !== generation) return;
      if (render(value)) poll = window.setTimeout(() => void request(), 1500);
    } catch (error) {
      if (!active || currentGeneration !== generation) return;
      render({ state: "error", message: error instanceof Error ? error.message : "Could not connect to the bridge.", userCode: null, verificationUri: null, expiresAt: null });
    } finally {
      pending = false;
      if (active && currentGeneration !== generation) void request();
      start.disabled = false;
      cancel.disabled = false;
    }
  };

  start.addEventListener("click", () => {
    if (options.isBusy()) { status.textContent = "Stop the current response before signing in."; return; }
    void request("POST");
  });
  cancel.addEventListener("click", () => void request("DELETE"));
  copy.addEventListener("click", async () => {
    const text = code.textContent;
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      copy.textContent = "Copied ✓";
    } catch {
      const range = document.createRange();
      range.selectNodeContents(code);
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(range);
      copy.textContent = "Press Ctrl+C to copy";
    }
  });
  proceed.addEventListener("click", options.onSuccess);
  browser.addEventListener("click", (event) => {
    event.preventDefault();
    if (browser.getAttribute("href") !== "https://github.com/login/device") return;
    Office.context.ui.openBrowserWindow(browser.href);
  });

  return {
    open() { active = true; generation++; void request(); },
    close() { active = false; generation++; window.clearTimeout(poll); },
  };
}
