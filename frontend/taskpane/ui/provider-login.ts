import type { BridgeClient } from "../bridge-client.js";
import { confirmAction } from "./confirm-action.js";
import type { ProviderId } from "../../../protocol/src/index.js";
import type { ProviderLoginKind } from "../../../protocol/src/providers.js";
import type { LoginAccount, LoginMethod, LoginPrompt, ProviderLoginStatus } from "../../../protocol/src/provider-login.js";

export function setupProviderLogin(options: {
  client: () => BridgeClient;
  provider: () => ProviderId;
  loginKind: () => ProviderLoginKind | undefined;
  /** The provider's CLI can sign out. */
  canLogout: () => boolean;
  onSuccess: (provider: ProviderId) => void;
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
  const pasteCard = document.getElementById("login-paste-card")!;
  const paste = document.getElementById("login-paste") as HTMLInputElement;
  const pasteSubmit = document.getElementById("login-paste-submit") as HTMLButtonElement;
  const el = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const chooser = el("login-accounts");
  const loadingNote = el("login-account-loading");
  const retry = el<HTMLButtonElement>("login-account-retry");
  const form = el("login-account-form");
  const search = el<HTMLInputElement>("login-account-search");
  const accountList = el<HTMLSelectElement>("login-account-provider");
  const methodLabel = el("login-account-method-label");
  const methodList = el<HTMLSelectElement>("login-account-method");
  const promptBox = el("login-account-prompts");
  const keyRow = el("login-account-key-row");
  const keyInput = el<HTMLInputElement>("login-account-key");
  const go = el<HTMLButtonElement>("login-account-go");
  const logout = el<HTMLButtonElement>("login-account-logout");
  let accounts: LoginAccount[] = [];
  let accountsLoad = 0;
  let poll: number | undefined;
  let generation = 0;
  let active = false;
  let pending = false;
  /** Set when the screen opens: its first status request must not replay an earlier sign-in's result. */
  let fresher = false;

  let shown: ProviderId = options.provider();
  const safeLink = (uri: string | null) => {
    try { return uri && new URL(uri).protocol === "https:" ? uri : null; } catch { return null; }
  };

  const render = (value: ProviderLoginStatus) => {
    status.textContent = value.message;
    status.dataset.state = value.state;
    const waiting = value.state === "waiting";
    const running = waiting || value.state === "starting";
    const link = waiting ? safeLink(value.verificationUri) : null;
    codeCard.hidden = !waiting || !value.userCode;
    code.textContent = value.userCode ?? "";
    browser.hidden = !link;
    browser.removeAttribute("href");
    if (link) browser.href = link;
    pasteCard.hidden = !(waiting && value.needsCode);
    if (pasteCard.hidden) paste.value = "";
    copy.textContent = "Copy code";
    copy.disabled = !waiting;
    // A provider that lists accounts has a chooser instead of one Sign in button.
    const choosing = options.loginKind() === "account";
    chooser.hidden = !(choosing && !running && value.state !== "success");
    start.hidden = running || value.state === "success" || choosing;
    start.textContent = value.state === "idle" ? "Sign in →" : "Try again →";
    cancel.hidden = !running;
    proceed.hidden = value.state !== "success";
    timer.textContent = value.expiresAt && running
      ? `Session timeout in up to ${Math.max(0, Math.ceil((value.expiresAt - Date.now()) / 60_000))} min` : "";
    return running;
  };

  const providerNow = () => accounts.find(account => account.id === accountList.value);
  const methodNow = (): LoginMethod | undefined => providerNow()?.methods.find(method => String(method.index) === methodList.value);
  const answers = new Map<string, string>();
  const promptRows = new Map<string, { row: HTMLElement; prompt: LoginPrompt }>();

  const refreshPromptVisibility = () => {
    for (const { row, prompt } of promptRows.values()) {
      const when = prompt.when;
      row.hidden = !!when && ((answers.get(when.key) ?? "") === when.value) !== (when.op === "eq");
    }
  };

  const fillPrompts = () => {
    promptBox.replaceChildren();
    promptRows.clear();
    answers.clear();
    const method = methodNow();
    keyInput.value = "";
    keyRow.hidden = method?.type !== "api";
    go.textContent = method?.type === "api" ? "Save key →" : "Continue →";
    for (const [index, prompt] of (method?.prompts ?? []).entries()) {
      const row = document.createElement("div");
      const id = `login-account-prompt-${index}`;
      const label = document.createElement("label");
      label.htmlFor = id;
      label.textContent = prompt.message;
      let input: HTMLInputElement | HTMLSelectElement;
      if (prompt.type === "select") {
        input = document.createElement("select");
        for (const option of prompt.options ?? []) {
          const item = document.createElement("option");
          item.value = option.value;
          item.textContent = option.hint ? `${option.label} · ${option.hint}` : option.label;
          input.append(item);
        }
      } else {
        input = document.createElement("input");
        input.type = "text";
        input.autocomplete = "off";
        input.spellcheck = false;
        if (prompt.placeholder) input.placeholder = prompt.placeholder;
      }
      input.id = id;
      answers.set(prompt.key, input.value);
      const changed = () => { answers.set(prompt.key, input.value); refreshPromptVisibility(); };
      input.addEventListener("input", changed);
      input.addEventListener("change", changed);
      row.append(label, input);
      promptBox.append(row);
      promptRows.set(prompt.key, { row, prompt });
    }
    refreshPromptVisibility();
  };

  const fillMethods = () => {
    const account = providerNow();
    methodList.replaceChildren(...(account?.methods ?? []).map(method => {
      const item = document.createElement("option");
      item.value = String(method.index);
      item.textContent = method.label;
      return item;
    }));
    const single = (account?.methods.length ?? 0) < 2;
    methodList.hidden = single;
    methodLabel.hidden = single;
    logout.hidden = !(options.canLogout() && account?.connected);
    fillPrompts();
  };

  const fillAccounts = () => {
    const needle = search.value.trim().toLowerCase();
    const found = accounts.filter(account => !needle || account.name.toLowerCase().includes(needle) || account.id.includes(needle));
    const keep = accountList.value;
    accountList.replaceChildren(...found.map(account => {
      const item = document.createElement("option");
      item.value = account.id;
      item.textContent = account.connected ? `${account.name} ✓` : account.name;
      return item;
    }));
    accountList.value = found.some(account => account.id === keep) ? keep : found[0]?.id ?? "";
    go.disabled = found.length === 0;
    fillMethods();
  };

  const loadAccounts = async () => {
    const load = ++accountsLoad;
    loadingNote.hidden = false;
    loadingNote.textContent = "Loading accounts… the first time takes a few seconds.";
    retry.hidden = true;
    form.hidden = true;
    try {
      const list = await options.client().loginAccounts(shown);
      if (!active || load !== accountsLoad) return;
      accounts = list;
      loadingNote.hidden = true;
      form.hidden = false;
      fillAccounts();
    } catch (error) {
      if (!active || load !== accountsLoad) return;
      loadingNote.textContent = error instanceof Error ? error.message : "Could not load the accounts.";
      retry.hidden = false;
    }
  };

  const startWithAccount = async () => {
    const account = providerNow();
    const method = methodNow();
    if (!account || !method) return;
    if (options.isBusy()) { status.textContent = "Stop the current response before signing in."; return; }
    const inputs: Record<string, string> = {};
    for (const [key, { row }] of promptRows) { const value = answers.get(key) ?? ""; if (!row.hidden && value) inputs[key] = value; }
    const key = method.type === "api" ? keyInput.value.trim() : undefined;
    if (method.type === "api" && !key) { status.textContent = "Paste the API key first."; return; }
    go.disabled = true;
    window.clearTimeout(poll);
    try {
      const value = await options.client().providerLogin("POST", shown, { choice: { accountId: account.id, method: method.index, ...(Object.keys(inputs).length ? { inputs } : {}), ...(key ? { key } : {}) } });
      keyInput.value = "";
      if (active && render(value)) poll = window.setTimeout(() => void request(), 1500);
    } catch (error) {
      if (active) status.textContent = error instanceof Error ? error.message : "Could not start the sign-in.";
    } finally {
      go.disabled = false;
    }
  };

  search.addEventListener("input", fillAccounts);
  accountList.addEventListener("change", fillMethods);
  methodList.addEventListener("change", fillPrompts);
  go.addEventListener("click", () => void startWithAccount());
  logout.addEventListener("click", async () => {
    const account = providerNow();
    if (!account || options.isBusy()) return;
    if (!await confirmAction(form, `Sign out of ${account.name}?`, "Sign out")) return;
    try {
      await options.client().providerLogout(shown, account.id);
      await loadAccounts();
    } catch (error) {
      status.textContent = error instanceof Error ? error.message : "Could not sign out.";
    }
  });
  retry.addEventListener("click", () => void loadAccounts());

  const request = async (method: "GET" | "POST" | "DELETE" = "GET") => {
    if (pending) return;
    const currentGeneration = generation;
    const fresh = fresher;
    fresher = false;
    pending = true;
    window.clearTimeout(poll);
    start.disabled = true;
    cancel.disabled = true;
    try {
      const value = await options.client().providerLogin(method, shown, { fresh });
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
  pasteSubmit.addEventListener("click", async () => {
    const text = paste.value.trim();
    if (!text || pending) return;
    pasteSubmit.disabled = true;
    window.clearTimeout(poll);
    try {
      render(await options.client().submitProviderLoginCode(shown, text));
    } catch (error) {
      status.textContent = error instanceof Error ? error.message : "Could not send the code.";
    } finally {
      pasteSubmit.disabled = false;
      if (active) poll = window.setTimeout(() => void request(), 1500);
    }
  });
  proceed.addEventListener("click", () => options.onSuccess(shown));
  browser.addEventListener("click", (event) => {
    event.preventDefault();
    const href = browser.getAttribute("href");
    if (!href || !safeLink(href)) return;
    Office.context.ui.openBrowserWindow(href);
  });

  return {
    open() {
      active = true; shown = options.provider(); generation++; fresher = true; void request();
      if (options.loginKind() === "account") void loadAccounts();
    },
    close() { active = false; generation++; accountsLoad++; window.clearTimeout(poll); },
  };
}
