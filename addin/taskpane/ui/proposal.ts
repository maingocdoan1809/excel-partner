import type { ProposalItem, ProposalRequest, ProposalResponse } from "../executor/word/proposal.js";

/** Most tokens compared word by word; longer edits are shown as plain before / after text. */
const DIFF_TOKEN_LIMIT = 1500;
/** Longest text shown in full in the card before it scrolls. */
const SHOWN_CHARS = 6000;

type Part = { text: string; changed: boolean };

function tokens(text: string): string[] {
  return text.match(/\s+|[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_]/gu) ?? [];
}

/** Word-level difference: the parts of `before` that were removed and the parts of `after` that were added. */
export function diffParts(before: string, after: string): { before: Part[]; after: Part[] } {
  const a = tokens(before);
  const b = tokens(after);
  if (a.length > DIFF_TOKEN_LIMIT || b.length > DIFF_TOKEN_LIMIT) {
    return { before: [{ text: before, changed: false }], after: [{ text: after, changed: false }] };
  }
  const width = b.length + 1;
  const table = new Uint16Array((a.length + 1) * width);
  for (let i = a.length - 1; i >= 0; i--) {
    for (let j = b.length - 1; j >= 0; j--) {
      table[i * width + j] = a[i] === b[j] ? table[(i + 1) * width + j + 1]! + 1 : Math.max(table[(i + 1) * width + j]!, table[i * width + j + 1]!);
    }
  }
  const left: Part[] = [];
  const right: Part[] = [];
  const push = (list: Part[], text: string, changed: boolean) => {
    const last = list[list.length - 1];
    if (last && last.changed === changed) last.text += text;
    else list.push({ text, changed });
  };
  let i = 0;
  let j = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) { push(left, a[i]!, false); push(right, b[j]!, false); i++; j++; }
    else if (table[(i + 1) * width + j]! >= table[i * width + j + 1]!) push(left, a[i++]!, true);
    else push(right, b[j++]!, true);
  }
  while (i < a.length) push(left, a[i++]!, true);
  while (j < b.length) push(right, b[j++]!, true);
  return { before: left, after: right };
}

function fill(target: HTMLElement, parts: Part[], changedClass: string): void {
  target.replaceChildren();
  let shown = 0;
  for (const part of parts) {
    if (shown >= SHOWN_CHARS) { target.append("…"); break; }
    const text = part.text.length + shown > SHOWN_CHARS ? `${part.text.slice(0, SHOWN_CHARS - shown)}…` : part.text;
    shown += text.length;
    if (part.changed) {
      const mark = document.createElement("mark");
      mark.className = changedClass;
      mark.textContent = text;
      target.append(mark);
    } else target.append(text);
  }
  if (target.childNodes.length === 0) {
    const empty = document.createElement("em");
    empty.textContent = "(empty)";
    target.append(empty);
  }
}

/**
 * A card with the edits the assistant proposes, before anything is written. Every piece of text comes from the document
 * or the model, so it is only ever put in with textContent. Resolves when the person applies a selection, rejects all
 * (nothing accepted), cancels with Escape, or the request is aborted (cancelled).
 */
export function showProposalReview(container: HTMLElement, request: ProposalRequest, signal: AbortSignal): Promise<ProposalResponse> {
  const cancelled: ProposalResponse = { cancelled: true, accepted: [] };
  if (signal.aborted) return Promise.resolve(cancelled);
  return new Promise((resolve) => {
    const card = document.createElement("form");
    card.className = "proposal-card";
    card.setAttribute("aria-label", "Review proposed edits");
    card.noValidate = true;

    const heading = document.createElement("h3");
    heading.className = "proposal-title";
    heading.textContent = request.title?.trim() || "Review proposed edits";
    const hint = document.createElement("p");
    hint.className = "proposal-hint";
    hint.textContent = "Nothing is changed in the document until you apply.";
    card.append(heading, hint);

    const rows = request.items.map((item) => buildRow(item));
    const list = document.createElement("div");
    list.className = "proposal-list";
    for (const row of rows) list.append(row.element);
    card.append(list);

    const actions = document.createElement("div");
    actions.className = "proposal-actions";
    const reject = document.createElement("button");
    reject.type = "button";
    reject.textContent = "Reject all";
    const apply = document.createElement("button");
    apply.type = "submit";
    actions.append(reject, apply);
    card.append(actions);

    const refresh = () => {
      const count = rows.filter((r) => r.checkbox.checked && !r.checkbox.disabled).length;
      apply.textContent = count > 0 ? `Apply selected (${count})` : "Apply selected";
      apply.disabled = count === 0;
    };
    for (const row of rows) row.checkbox.addEventListener("change", refresh);
    refresh();

    let finished = false;
    const finish = (response: ProposalResponse) => {
      if (finished) return;
      finished = true;
      signal.removeEventListener("abort", cancel);
      document.removeEventListener("keydown", onKey, true);
      card.remove();
      resolve(response);
    };
    const cancel = () => finish(cancelled);
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape" && card.isConnected) { event.preventDefault(); cancel(); }
    };
    reject.addEventListener("click", () => finish({ cancelled: false, accepted: [] }));
    card.addEventListener("submit", (event) => {
      event.preventDefault();
      const accepted = rows
        .filter((r) => r.checkbox.checked && !r.checkbox.disabled)
        .map((r) => ({ index: r.item.index, text: r.chosen() }));
      finish({ cancelled: false, accepted });
    });
    signal.addEventListener("abort", cancel, { once: true });
    document.addEventListener("keydown", onKey, true);
    container.append(card);
    (rows.find((r) => !r.checkbox.disabled)?.checkbox ?? apply).focus();
  });
}

interface Row {
  item: ProposalItem;
  element: HTMLElement;
  checkbox: HTMLInputElement;
  chosen: () => string;
}

function buildRow(item: ProposalItem): Row {
  const element = document.createElement("article");
  element.className = "proposal-item";
  element.dataset.found = String(item.found);
  element.dataset.ambiguous = String(item.ambiguous);

  const head = document.createElement("label");
  head.className = "proposal-item-head";
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.checked = item.found && !item.ambiguous;
  checkbox.disabled = !item.found;
  const name = document.createElement("strong");
  name.textContent = `Edit ${item.index + 1}`;
  head.append(checkbox, name);
  if (!item.found || item.ambiguous) {
    const badge = document.createElement("span");
    badge.className = "proposal-badge";
    badge.textContent = !item.found ? "Not found in the document" : "Appears more than once; the first match is used";
    head.append(badge);
  }
  element.append(head);

  if (item.reason) {
    const reason = document.createElement("p");
    reason.className = "proposal-reason";
    reason.textContent = item.reason;
    element.append(reason);
  }

  const before = document.createElement("div");
  before.className = "proposal-before";
  before.setAttribute("aria-label", "Current text");
  const after = document.createElement("div");
  after.className = "proposal-after";
  after.setAttribute("aria-label", "New text");
  element.append(before, after);

  const choices = [item.newText, ...(item.alternatives ?? [])];
  let selected = 0;
  const render = () => {
    const parts = diffParts(item.oldText, choices[selected]!);
    fill(before, parts.before, "proposal-del");
    fill(after, parts.after, "proposal-ins");
  };
  render();

  if (choices.length > 1) {
    const group = document.createElement("fieldset");
    group.className = "proposal-choices";
    const legend = document.createElement("legend");
    legend.textContent = "Wording";
    group.append(legend);
    const radioName = `proposal-${item.index}-${Math.random().toString(36).slice(2, 8)}`;
    choices.forEach((text, i) => {
      const label = document.createElement("label");
      label.className = "proposal-choice";
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = radioName;
      radio.checked = i === 0;
      radio.addEventListener("change", () => { if (radio.checked) { selected = i; render(); } });
      const span = document.createElement("span");
      span.textContent = text.length > 160 ? `${text.slice(0, 160)}…` : text || "(empty)";
      label.append(radio, span);
      group.append(label);
    });
    element.append(group);
  }
  return { item, element, checkbox, chosen: () => choices[selected]! };
}
