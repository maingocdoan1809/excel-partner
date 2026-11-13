import type { ToolResult } from "../../../protocol/src/index.js";
import { activeHost } from "../host/index.js";
import { describeChange } from "./changes.js";
import { byId } from "./dom.js";
import type { StoredChange, StoredSession } from "./sessions.js";

export interface ChangesPanelDeps {
  getSession(): StoredSession;
  schedulePersist(): void;
  goTo(target: { sheet?: string; address: string }): Promise<void>;
  elements: {
    changesCount: HTMLElement;
    changesList: HTMLElement;
    menuDot: HTMLElement;
    panels: { changes: HTMLElement };
  };
}

/** The current chat's change list. */
export function setupChangesPanel(deps: ChangesPanelDeps) {
  const els = deps.elements;

  function recordChange(tool: string, args: unknown, result: ToolResult, context: { callId: string; requestId: string }): void {
    const change = describeChange(tool, args, result, activeHost().writeTools);
    if (!change) return;
    deps.getSession().changes.push({ at: Date.now(), tool: change.tool, ok: change.ok, summary: change.summary, target: change.target, callId: context.callId, requestId: context.requestId });
    renderChanges();
    deps.schedulePersist();
  }

  function renderChanges(): void {
    const total = deps.getSession().changes.length;
    els.changesCount.hidden = total === 0;
    els.changesCount.textContent = String(total);
    const tabCount = byId("tab-changes-count");
    tabCount.hidden = total === 0;
    tabCount.textContent = String(total);
    els.menuDot.hidden = total === 0;
    if (els.panels.changes.hidden) return;

    if (total === 0) {
      els.changesList.innerHTML = '<div class="empty">No changes yet. Edits made to the workbook in this chat will be listed here.</div>';
      return;
    }
    els.changesList.replaceChildren(...[...deps.getSession().changes].reverse().map((entry) => changeRow(entry)));
  }

  function changeRow(entry: StoredChange): HTMLElement {
    const row = document.createElement("div");
    row.className = "change";
    row.dataset.ok = String(entry.ok);
    const time = document.createElement("span");
    time.className = "time";
    time.textContent = new Date(entry.at).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
    const summary = document.createElement("span");
    summary.className = "summary";
    summary.textContent = entry.summary;
    row.append(time, summary);
    if (entry.target) {
      const target = entry.target;
      const go = document.createElement("button");
      go.className = "secondary";
      go.textContent = "Go to";
      go.addEventListener("click", () => void deps.goTo(target));
      row.append(go);
    }
    return row;
  }

  return { recordChange, renderChanges };
}
