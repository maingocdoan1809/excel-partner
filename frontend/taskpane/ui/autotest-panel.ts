import { SYSTEM_TEST_COMMAND, type AutotestEvent, type AutotestRunState, type AutotestSettings } from "../../../protocol/src/autotest.js";
import { reduceRun } from "../../../protocol/src/autotest-state.js";
import { exportWorkbook, ProgressWriter } from "../executor/autotest-sheet.js";
import type { AutotestOverview, BridgeClient } from "../bridge-client.js";
import { errorText as sharedErrorText } from "./dom.js";

const ORDER = ["PASS", "FAIL", "RUNNING", "PENDING", "BLOCKED", "PLANNED", "NOT EXECUTED", "NOT AUTOMATED"] as const;
const COLOR: Record<string, string> = { PASS: "#18794e", FAIL: "#c0262d", RUNNING: "#1f5fbf", PENDING: "#a15c00", BLOCKED: "#a15c00", PLANNED: "#c9ced6", "NOT EXECUTED": "#6b7482", "NOT AUTOMATED": "#6b7482" };
const MAX_LIST = 12;

/** The Test run, environment and kit sections of the Testing screen. */
export function setupAutotestPanel(client: () => BridgeClient, workbookName: () => Promise<string | null>): {
  refresh(): Promise<void>; onEvent(run: string, event: AutotestEvent): void; sync(): Promise<boolean>;
  /** Before a message about testing is sent: give the kit the workbook as it is now. Never blocks or fails the message. */
  syncForPrompt(prompt: string): Promise<void>;
} {
  const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const progress = get<HTMLElement>("at-progress");
  const form = get<HTMLFormElement>("at-form");
  const note = get<HTMLElement>("at-save-note");
  const syncNote = get<HTMLElement>("at-sync-note");
  const kit = get<HTMLElement>("at-kit");
  const installButton = get<HTMLButtonElement>("at-install");
  const installLog = get<HTMLElement>("at-install-log");
  const auth = get<HTMLSelectElement>("at-auth");
  const errorText = (error: unknown) => sharedErrorText(error, "Something went wrong.");
  const field = (id: string) => get<HTMLInputElement>(id);
  // The live view is the SAME reduction the bridge and the dashboard use, over the events received so far (so the three can never disagree).
  let live: AutotestRunState | null = null;
  let liveRun = "";
  let liveEvents: AutotestEvent[] = [];
  let redraw: number | undefined;
  let resultValues: Record<string, string> = {};
  let installTimer: number | undefined;
  const writer = new ProgressWriter(() => resultValues, message => { syncNote.textContent = `Could not write to the sheet: ${message}`; });

  const showAuth = () => {
    get<HTMLElement>("at-form-fields").hidden = auth.value !== "form";
    get<HTMLElement>("at-session-fields").hidden = auth.value !== "session";
  };
  auth.addEventListener("change", showAuth);

  function fill(settings: AutotestSettings | null): void {
    if (!settings) { showAuth(); return; }
    field("at-project").value = settings.project; field("at-url").value = settings.url; field("at-account").value = settings.testAccount;
    auth.value = settings.authType; field("at-login-url").value = settings.loginUrl; field("at-user-env").value = settings.usernameEnv;
    field("at-pass-env").value = settings.passwordEnv; field("at-prefix").value = settings.testPrefix;
    get<HTMLSelectElement>("at-lang").value = settings.talkLanguage; field("at-repeat").value = String(settings.maxRepeat);
    field("at-blackbox").checked = settings.blackBox;
    field("at-ask-data").checked = settings.askBeforeTestData;
    get<HTMLSelectElement>("at-signin").value = settings.signIn;
    showAuth();
  }

  function read(): Record<string, unknown> {
    const repeat = Number(field("at-repeat").value);
    return {
      project: field("at-project").value.trim(), url: field("at-url").value.trim(), testAccount: field("at-account").value.trim(), authType: auth.value,
      // The variable names only matter when the kit fills the sign-in form itself; otherwise they are not sent at all.
      ...(auth.value === "form" ? { loginUrl: field("at-login-url").value.trim(), usernameEnv: field("at-user-env").value.trim(), passwordEnv: field("at-pass-env").value.trim() } : {}),
      testPrefix: field("at-prefix").value.trim() || "ZZ_AUTOTEST_", talkLanguage: get<HTMLSelectElement>("at-lang").value,
      maxRepeat: Number.isInteger(repeat) && repeat >= 2 && repeat <= 10 ? repeat : 3, blackBox: field("at-blackbox").checked,
      signIn: auth.value === "session" ? get<HTMLSelectElement>("at-signin").value : "", askBeforeTestData: field("at-ask-data").checked,
    };
  }

  function drawProgress(): void {
    progress.replaceChildren();
    if (!live) { progress.textContent = "No run yet."; return; }
    const head = document.createElement("p");
    const label = { preparing: "Preparing", planned: "Planned, waiting to start", running: "Running", paused: "Paused to fix a case", finished: "Finished", stopped: "Stopped", idle: "Idle" }[live.phase];
    head.textContent = `${label}: ${live.id}`;
    progress.append(head);
    if (live.currentPhase) { const note = document.createElement("p"); note.className = "template-help"; note.textContent = `${live.currentPhase.phase}: ${live.currentPhase.text}`; progress.append(note); }
    if (live.stopReason) { const stop = document.createElement("p"); stop.className = "at-FAIL"; stop.textContent = live.phase === "paused" ? `Paused so the case can be fixed right away: ${live.stopReason}. The cases not reached are waiting.` : `The run stopped by itself: ${live.stopReason}. Cases not run are marked NOT EXECUTED.`; progress.append(stop); }
    const total = live.cases.length || 1;
    const bar = document.createElement("div");
    bar.className = "at-bar";
    for (const status of ORDER) if (live.counts[status]) { const part = document.createElement("i"); part.style.width = `${(100 * live.counts[status]!) / total}%`; part.style.background = COLOR[status]!; part.title = `${status}: ${live.counts[status]}`; bar.append(part); }
    progress.append(bar);
    const summary = document.createElement("p");
    summary.textContent = ORDER.filter(status => live!.counts[status]).map(status => `${status} ${live!.counts[status]}`).join(" · ");
    progress.append(summary);
    const items = [...live.cases].sort((a, b) => Number(b.status === "RUNNING") - Number(a.status === "RUNNING"));
    for (const item of items.slice(0, MAX_LIST)) {
      const row = document.createElement("div");
      row.className = "at-case";
      const id = document.createElement("b"); id.textContent = item.id;
      const lastStep = item.steps[item.steps.length - 1];
      const detail = document.createElement("span"); detail.textContent = item.status === "RUNNING" ? `${item.step}${item.total ? `/${item.total}` : ""} ${lastStep?.caption ?? ""}`.trim() : item.remark || item.title;
      const status = document.createElement("span"); status.className = `at-${item.status.replace(/ /g, "_")}`; status.textContent = item.status;
      row.append(id, detail, status);
      progress.append(row);
    }
    if (items.length > MAX_LIST) { const more = document.createElement("p"); more.textContent = `…and ${items.length - MAX_LIST} more. The dashboard shows all of them.`; progress.append(more); }
  }

  function drawKit(data: AutotestOverview): void {
    kit.replaceChildren();
    const lines = [
      `Folder: ${data.home}`,
      data.kitInstalled ? `Kit ${data.kitVersion} installed` : "Kit not installed yet: save the settings above.",
      data.runtime.python ? `Python ${data.runtime.python}${data.runtime.pythonOk ? "" : " is too old: the kit needs Python 3.9 or newer. Install a newer Python, then reopen this screen."}` : "Python was not found. Install Python 3.9 or newer, then reopen this screen.",
      data.runtime.python ? (data.runtime.missing.length ? `Missing libraries: ${data.runtime.missing.join(", ")}` : "Libraries ready") : "",
      `${data.knowledge} fact${data.knowledge === 1 ? "" : "s"} learned so far`,
      data.kitSync.updated.length ? `Kit and skills refreshed to this release (${data.kitSync.updated.length} files).${data.kitSync.backup ? ` The old copies are in ${data.kitSync.backup}.` : ""}` : "",
    ].filter(Boolean);
    for (const text of lines) { const row = document.createElement("div"); row.textContent = text; kit.append(row); }
    const state = data.runtime.install.state;
    installButton.disabled = !data.runtime.pythonOk || state === "running" || !data.kitInstalled;
    installButton.textContent = state === "running" ? "Installing…" : data.runtime.ready ? "Reinstall libraries and browser" : "Install Python libraries and browser";
    installLog.hidden = !data.runtime.install.log;
    installLog.textContent = data.runtime.install.log;
    window.clearTimeout(installTimer);
    if (state === "running") installTimer = window.setTimeout(() => { void refresh(); }, 2500);
  }

  /** Send the open workbook to the kit. The kit reads this copy; results are written back into the open workbook. */
  async function sync(): Promise<boolean> {
    try {
      const name = (await workbookName()) || "workbook.xlsx";
      writer.file = await client().autotestUploadWorkbook(name, await exportWorkbook());
      syncNote.textContent = `Sent “${name}” to the test kit (${new Date().toLocaleTimeString()}). Results will be written into this sheet.`;
      return true;
    } catch (error) { syncNote.textContent = `Could not send the workbook: ${errorText(error)}`; return false; }
  }

  async function refresh(): Promise<void> {
    try {
      const data = await client().autotestOverview();
      resultValues = data.resultValues;
      const picker = get<HTMLSelectElement>("at-signin");
      const chosen = picker.value;
      picker.replaceChildren(...(data.signIns.length ? data.signIns : [{ name: "", url: "", savedAt: "" }]).map(item => {
        const option = document.createElement("option");
        option.value = item.name;
        option.textContent = item.name ? `${item.name} (${item.url})` : "No saved sign-in yet: create one below";
        return option;
      }));
      if (chosen && data.signIns.some(item => item.name === chosen)) picker.value = chosen;
      // First time here with a saved sign-in available: that is the easy way in.
      if (!data.settings && data.signIns.length && auth.value !== "session") { auth.value = "session"; showAuth(); }
      if (document.activeElement && form.contains(document.activeElement)) { /* keep what the person is typing */ } else fill(data.settings);
      drawKit(data);
      if (document.activeElement !== field("at-home")) field("at-home").value = data.home;
      get<HTMLAnchorElement>("at-dashboard").href = `${location.origin}${data.dashboardPath}`;
      if (!live && data.runs[0]) progress.textContent = `Latest run: ${data.runs[0].id} (${data.runs[0].phase}). The dashboard has the details.`;
    } catch (error) { note.textContent = errorText(error); }
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    note.textContent = "Saving…";
    void client().autotestAction({ action: "save-settings", settings: read() })
      .then(() => { note.textContent = "Saved."; return refresh(); })
      .catch((error) => { note.textContent = errorText(error); });
  });
  installButton.addEventListener("click", () => {
    void client().autotestAction({ action: "install-runtime" }).then(() => refresh()).catch((error) => { note.textContent = errorText(error); });
  });
  get<HTMLButtonElement>("at-home-save").addEventListener("click", () => {
    note.textContent = "Changing folder…";
    void client().autotestAction({ action: "choose-home", dir: field("at-home").value.trim() })
      .then(() => { note.textContent = "Folder changed. Save the settings to install the kit there."; return refresh(); })
      .catch((error) => { note.textContent = errorText(error); });
  });
  get<HTMLButtonElement>("at-sync").addEventListener("click", () => { void sync(); });

  return {
    refresh,
    sync,
    async syncForPrompt(prompt) {
      if (!SYSTEM_TEST_COMMAND.test(prompt) || Office.context.host !== Office.HostType.Excel) return;
      try {
        const data = await client().autotestOverview();
        resultValues = data.resultValues;
        if (data.kitInstalled) await sync();
      } catch { /* the message is sent anyway */ }
    },
    onEvent(run, event) {
      if (run !== liveRun) { liveRun = run; liveEvents = []; }
      liveEvents.push(event);
      // a fast run sends many events: fold and draw at most a few times a second
      if (redraw === undefined) redraw = window.setTimeout(() => { redraw = undefined; live = reduceRun(liveRun, liveEvents); drawProgress(); }, 250);
      writer.push(event);
      if (event.kind === "run_finished" || event.kind === "run_stopped") void writer.flush();
    },
  };
}
