/// <reference lib="dom" />

/**
 * When does the pane bring back the last chat?
 *
 * Only when the pane reloaded itself (a new build arrived, or an update finished) in the middle of a conversation.
 * Closing the pane and opening it again starts a new chat; the old one stays in the history list.
 *
 * The reload is marked in sessionStorage, which belongs to this one pane window and is gone once it is closed.
 * The mark is removed the moment it is read, so it can only ever apply to the very next start.
 */

const RESTORE_KEY = "excelPartner.restoreChat";

type Marker = Pick<Storage, "getItem" | "setItem" | "removeItem">;

function pageStorage(): Marker | null {
  try {
    return window.sessionStorage;
  } catch {
    return null; // blocked storage: every start is a new chat, which is the safe default
  }
}

/** Reload the pane and keep the current chat open across the reload. */
export function reloadKeepingChat(reload: () => void = () => window.location.reload(), storage: Marker | null = pageStorage()): void {
  try {
    storage?.setItem(RESTORE_KEY, "1");
  } catch {
    // The reload still happens; the chat is then found again in the history list.
  }
  reload();
}

/** True once after `reloadKeepingChat`; false on any other start (first open, reopened pane, a manual refresh). */
export function takeChatRestore(storage: Marker | null = pageStorage()): boolean {
  try {
    if (!storage || storage.getItem(RESTORE_KEY) !== "1") return false;
    storage.removeItem(RESTORE_KEY);
    return true;
  } catch {
    return false;
  }
}
