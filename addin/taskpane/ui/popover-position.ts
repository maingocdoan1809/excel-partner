/** Anchor to the viewport, so toolbar reordering cannot push a menu off-pane. */
export function positionComposerPopover(button: HTMLElement, popover: HTMLElement): void {
  const gutter = 12;
  const anchor = button.getBoundingClientRect();
  popover.style.position = "fixed";
  popover.style.minWidth = "0";
  popover.style.maxWidth = `calc(100vw - ${gutter * 2}px)`;
  popover.style.right = "auto";
  const width = popover.getBoundingClientRect().width;
  const left = Math.max(gutter, Math.min(anchor.left, window.innerWidth - width - gutter));
  popover.style.left = `${left}px`;
  const above = anchor.top - gutter - 8;
  const below = window.innerHeight - anchor.bottom - gutter - 8;
  if (above >= below) {
    popover.style.top = "auto";
    popover.style.bottom = `${Math.max(gutter, window.innerHeight - anchor.top + 8)}px`;
    popover.style.maxHeight = `${Math.max(0, Math.min(window.innerHeight * .6, above))}px`;
  } else {
    popover.style.top = `${anchor.bottom + 8}px`;
    popover.style.bottom = "auto";
    popover.style.maxHeight = `${Math.max(0, Math.min(window.innerHeight * .6, below))}px`;
  }
}
