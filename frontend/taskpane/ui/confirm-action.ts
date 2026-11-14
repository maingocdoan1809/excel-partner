/** Office task panes may reject window.confirm. Keep confirmation in the pane. */
export function confirmAction(container: HTMLElement, message: string, label: string): Promise<boolean> {
 const panel = container.closest<HTMLElement>(".panel");
 if (!container.isConnected || panel?.hidden) return Promise.resolve(false);
 return new Promise(resolve => {
   const previous = document.activeElement as HTMLElement | null;
   const card = document.createElement("section"); card.className = "inline-confirmation";
   card.setAttribute("role", "group"); card.setAttribute("aria-label", "Confirm action");
   const text = document.createElement("p"); text.textContent = message;
   const actions = document.createElement("div"); actions.className = "settings-actions";
   const cancel = document.createElement("button"); cancel.type = "button"; cancel.className = "secondary";
   cancel.textContent = "Cancel"; cancel.dataset.confirmCancel = "";
   const accept = document.createElement("button"); accept.type = "button"; accept.className = "secondary settings-danger";
   accept.textContent = label; accept.dataset.confirmAccept = "";
   let settled = false;
   const finish = (value: boolean) => {
     if (settled) return; settled = true; observer.disconnect(); card.remove();
     if (previous?.isConnected && !panel?.hidden) previous.focus({ preventScroll: true });
     resolve(value);
   };
   const observer = new MutationObserver(() => { if (panel?.hidden || !card.isConnected) finish(false); });
   if (panel) observer.observe(panel, { attributes: true, attributeFilter: ["hidden"] });
   cancel.addEventListener("click", () => finish(false)); accept.addEventListener("click", () => finish(true));
   card.addEventListener("keydown", event => { if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); finish(false); } });
   actions.append(cancel, accept); card.append(text, actions); container.append(card); cancel.focus();
 });
}
