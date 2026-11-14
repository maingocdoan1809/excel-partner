import { setProposalReviewer } from "../executor/word/index.js";
import { activeHost } from "../host/index.js";
import { byId } from "./dom.js";
import { showProposalReview } from "./proposal.js";

/** Names and wording that depend on the host: the brand, the page title and the review card Word edits are approved in. */
export function applyHostBranding(): void {
  const host = activeHost();
  document.title = "Office Partner";
  const brand = document.querySelector(".brand-text strong");
  if (brand) brand.textContent = "Office Partner";
  document.body.dataset.host = host.id;
  // Static wording in index.html names Excel; say the application this pane is really in.
  for (const el of document.querySelectorAll<HTMLElement>(".host-app")) el.textContent = host.app;
  for (const el of document.querySelectorAll<HTMLElement>("[data-host-label]")) el.textContent = (el.dataset.hostLabel ?? "").replace("{app}", host.app);
  if (!host.hasSheets) {
    // The testing tools work from a workbook only. Connected files work from every Office app.
    byId("menu-testing").hidden = true;
  }
  if (host.id === "word") {
    // propose_doc_edits stages edits; the card below is where the person applies them.
    setProposalReviewer((request, signal) => {
      const dock = byId("question-dock");
      dock.hidden = false;
      document.body.classList.add("question-active");
      const outcome = showProposalReview(dock, request, signal ?? new AbortController().signal);
      const close = () => { dock.hidden = true; document.body.classList.remove("question-active"); };
      void outcome.then(close, close);
      return outcome;
    });
  }
}
