import type { BridgeClient, TestingOverview } from "../bridge-client.js";

/** Sign-ins, notes and regression scripts for the testing workflow: the `excel-partner login/knowledge/scripts` commands as buttons. */
export function setupTestingPanel(client: () => BridgeClient): { refresh(): Promise<void> } {
  const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const status = get<HTMLElement>("testing-status");
  const signIns = get<HTMLElement>("testing-signins");
  const notes = get<HTMLElement>("testing-notes");
  const scripts = get<HTMLElement>("testing-scripts");
  const form = get<HTMLFormElement>("testing-signin-form");
  const nameInput = get<HTMLInputElement>("testing-signin-name");
  const urlInput = get<HTMLInputElement>("testing-signin-url");
  const startButton = get<HTMLButtonElement>("testing-signin-start");
  const wait = get<HTMLElement>("testing-signin-wait");
  const waitText = get<HTMLElement>("testing-signin-wait-text");
  const doneButton = get<HTMLButtonElement>("testing-signin-done");
  const cancelButton = get<HTMLButtonElement>("testing-signin-cancel");
  let request = 0;
  let busy = false;
  const errorText = (error: unknown) => error instanceof Error ? error.message : "Something went wrong.";

  function button(label: string, onClick: () => void | Promise<void>, secondary = true): HTMLButtonElement {
    const element = document.createElement("button");
    element.type = "button";
    if (secondary) element.className = "secondary";
    element.textContent = label;
    element.addEventListener("click", () => { void onClick(); });
    return element;
  }

  /** Deleting needs a second click on the same button; a browser confirm() dialog is not reliable inside the task pane. */
  function deleteButton(label: string, action: () => Promise<void>): HTMLButtonElement {
    let armed = false;
    let timer: number | undefined;
    const element = button(label, async () => {
      if (!armed) {
        armed = true;
        element.textContent = "Click again to delete";
        timer = window.setTimeout(() => { armed = false; element.textContent = label; }, 4000);
        return;
      }
      window.clearTimeout(timer);
      element.disabled = true;
      try { await action(); await refresh(); } catch (error) { status.textContent = errorText(error); element.disabled = false; armed = false; element.textContent = label; }
    });
    return element;
  }

  function row(title: string, meta: string, detail?: string): { row: HTMLElement; actions: HTMLElement } {
    const element = document.createElement("div");
    element.className = "reference-row testing-row";
    const head = document.createElement("div");
    head.className = "testing-row-title";
    const strong = document.createElement("strong");
    strong.textContent = title;
    const small = document.createElement("span");
    small.textContent = meta;
    head.append(strong, small);
    element.append(head);
    if (detail) { const p = document.createElement("p"); p.textContent = detail; element.append(p); }
    const actions = document.createElement("div");
    actions.className = "reference-actions";
    element.append(actions);
    return { row: element, actions };
  }

  /** A View button that shows or hides the file's text under the row. */
  function viewButton(element: HTMLElement, load: () => Promise<string>): HTMLButtonElement {
    let shown: HTMLElement | null = null;
    const toggle = button("View", async () => {
      if (shown) { shown.remove(); shown = null; toggle.textContent = "View"; return; }
      toggle.disabled = true;
      try {
        const pre = document.createElement("pre");
        pre.textContent = await load();
        element.append(pre);
        shown = pre;
        toggle.textContent = "Hide";
      } catch (error) { status.textContent = errorText(error); }
      finally { toggle.disabled = false; }
    });
    return toggle;
  }

  function showWaiting(pending: { name: string; url: string } | null): void {
    wait.hidden = !pending;
    form.hidden = Boolean(pending);
    if (pending) waitText.textContent = `A browser window is open at ${pending.url}. Sign in there as usual (including any 2-step check). When you can see the home page, come back here and press "I'm signed in" to save it as '${pending.name}'.`;
  }

  function render(data: TestingOverview): void {
    showWaiting(data.pendingSignIn);
    signIns.replaceChildren();
    if (!data.signIns.length) signIns.textContent = "No saved sign-ins yet.";
    for (const item of data.signIns) {
      const { row: element, actions } = row(item.name, `saved ${item.savedAt.slice(0, 16).replace("T", " ")}`, `Sites: ${item.origins.join(", ")}`);
      actions.append(
        button("Sign in again", () => begin(item.name, item.url)),
        deleteButton("Delete", () => client().testingDelete({ kind: "signin", name: item.name })),
      );
      signIns.append(element);
    }
    notes.replaceChildren();
    if (!data.notes.length) notes.textContent = "Nothing saved yet. The assistant saves notes after it has learned something about a system.";
    for (const site of data.notes) {
      const heading = document.createElement("p");
      heading.className = "template-help";
      heading.textContent = site.site;
      notes.append(heading);
      for (const note of site.notes) {
        const { row: element, actions } = row(note.name, `${note.kind} · ${note.updated}`, note.summary);
        actions.append(
          viewButton(element, async () => (await client().testingView("note", site.site, note.name))),
          deleteButton("Delete", () => client().testingDelete({ kind: "note", site: site.site, name: note.name })),
        );
        notes.append(element);
      }
    }
    scripts.replaceChildren();
    if (!data.scripts.length) scripts.textContent = "No scripts yet. The assistant saves one for each test case that passes or fails because of the application.";
    for (const site of data.scripts) {
      const heading = document.createElement("p");
      heading.className = "template-help";
      heading.textContent = site.site;
      scripts.append(heading);
      for (const script of site.scripts) {
        const { row: element, actions } = row(script.caseId, `${script.outcome === "pass" ? "passed" : "application defect"} · ${script.tested}`, script.title);
        actions.append(
          viewButton(element, async () => (await client().testingView("script", site.site, script.caseId))),
          deleteButton("Delete", () => client().testingDelete({ kind: "script", site: site.site, caseId: script.caseId })),
        );
        scripts.append(element);
      }
    }
  }

  async function refresh(): Promise<void> {
    const current = ++request;
    try {
      const data = await client().testingOverview();
      if (current === request) render(data);
    } catch (error) { if (current === request) status.textContent = errorText(error); }
  }

  async function begin(name: string, url: string): Promise<void> {
    if (busy) return;
    busy = true;
    startButton.disabled = true;
    status.textContent = "Opening the browser…";
    try {
      await client().testingAction({ action: "sign-in-start", name, url });
      status.textContent = "";
      showWaiting({ name, url });
    } catch (error) { status.textContent = errorText(error); }
    finally { busy = false; startButton.disabled = false; }
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const name = nameInput.value.trim().toLowerCase();
    const url = urlInput.value.trim();
    if (!name || !url) { status.textContent = "Enter a name and the system's address."; return; }
    void begin(name, url).then(() => { if (wait.hidden === false) { nameInput.value = ""; urlInput.value = ""; } });
  });

  doneButton.addEventListener("click", async () => {
    doneButton.disabled = cancelButton.disabled = true;
    status.textContent = "Saving the sign-in…";
    try {
      const saved = await client().testingAction({ action: "sign-in-finish" }) as { name?: string; origins?: string[] };
      status.textContent = `Saved sign-in '${saved.name}'. The assistant can use it on: ${(saved.origins ?? []).join(", ")}.`;
    } catch (error) { status.textContent = errorText(error); }
    finally { doneButton.disabled = cancelButton.disabled = false; await refresh(); }
  });

  cancelButton.addEventListener("click", async () => {
    try { await client().testingAction({ action: "sign-in-cancel" }); status.textContent = "Cancelled."; }
    catch (error) { status.textContent = errorText(error); }
    await refresh();
  });

  document.querySelectorAll<HTMLButtonElement>("[data-testing-folder]").forEach(element => element.addEventListener("click", async () => {
    try { await client().testingAction({ action: "open-folder", kind: element.dataset.testingFolder }); status.textContent = ""; }
    catch (error) { status.textContent = errorText(error); }
  }));

  return { refresh };
}
