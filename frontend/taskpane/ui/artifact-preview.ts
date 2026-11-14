import { buildPreviewDocument, type StoredArtifact } from "./artifact.js";
import { byId } from "./dom.js";

/** The full-page view of an artifact a reply produced: its rendered page, its source, and a button to copy it. */
export function setupArtifactPreview(showPreviewPanel: () => void) {
  const frame = byId<HTMLIFrameElement>("preview-frame");
  const source = byId<HTMLElement>("preview-source");
  const toggle = byId<HTMLButtonElement>("preview-toggle");
  const copy = byId<HTMLButtonElement>("preview-copy");
  let previewing: StoredArtifact | null = null;

  function showSource(on: boolean): void {
    toggle.setAttribute("aria-pressed", String(on));
    toggle.textContent = on ? "Preview" : "Source";
    frame.hidden = on;
    source.hidden = !on;
  }

  toggle.addEventListener("click", () => showSource(toggle.getAttribute("aria-pressed") !== "true"));
  copy.addEventListener("click", () => {
    if (!previewing) return;
    navigator.clipboard.writeText(previewing.content).then(
      () => { copy.textContent = "Copied"; },
      () => { copy.textContent = "Copy failed"; },
    ).finally(() => window.setTimeout(() => { copy.textContent = "Copy"; }, 1500));
  });

  return {
    open(artifact: StoredArtifact): void {
      previewing = artifact;
      byId("preview-title").textContent = artifact.name;
      source.textContent = artifact.content;
      frame.srcdoc = buildPreviewDocument(artifact);
      showSource(false);
      showPreviewPanel();
    },
    /** The panel was left: free the page. */
    close(): void {
      frame.srcdoc = "";
      previewing = null;
    },
  };
}
