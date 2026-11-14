const INPUT_HEIGHT_KEY = "excelPartner.inputHeight";

/** Drag handle above the message box: drag up to make it taller (keyboard: Arrow Up/Down). Height is remembered. */
export function setupInputResizer(handle: HTMLElement, input: HTMLTextAreaElement): void {
  const apply = (height: number) => {
    const max = Math.round(window.innerHeight * 0.5);
    document.documentElement.style.setProperty("--input-height", `${Math.round(Math.min(max, Math.max(40, height)))}px`);
  };
  const remember = () => localStorage.setItem(INPUT_HEIGHT_KEY, String(input.offsetHeight));

  const saved = Number(localStorage.getItem(INPUT_HEIGHT_KEY));
  if (saved) apply(saved);

  handle.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    handle.setPointerCapture(event.pointerId);
    handle.classList.add("dragging");
    const startY = event.clientY;
    const startHeight = input.offsetHeight;
    const move = (e: PointerEvent) => apply(startHeight - (e.clientY - startY));
    const end = () => {
      handle.classList.remove("dragging");
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", end);
      handle.removeEventListener("pointercancel", end);
      remember();
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", end);
  });

  handle.addEventListener("keydown", (event) => {
    const delta = event.key === "ArrowUp" ? 16 : event.key === "ArrowDown" ? -16 : 0;
    if (!delta) return;
    event.preventDefault();
    apply(input.offsetHeight + delta);
    remember();
  });
}
