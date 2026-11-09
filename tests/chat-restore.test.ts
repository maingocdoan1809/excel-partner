import { describe, expect, it, vi } from "vitest";
import { reloadKeepingChat, takeChatRestore } from "../addin/taskpane/ui/chat-restore.js";

function memory() {
  const data = new Map<string, string>();
  return { data, getItem: (k: string) => data.get(k) ?? null, setItem: (k: string, v: string) => void data.set(k, v), removeItem: (k: string) => void data.delete(k) };
}

describe("when the pane brings the last chat back", () => {
  it("does not on a normal start: first open, or the pane closed and opened again", () => {
    expect(takeChatRestore(memory())).toBe(false);
  });

  it("does once after the pane reloaded itself, and not on the start after that", () => {
    const storage = memory();
    const reload = vi.fn();
    reloadKeepingChat(reload, storage);
    expect(reload).toHaveBeenCalledOnce();
    expect(takeChatRestore(storage)).toBe(true);
    expect(takeChatRestore(storage)).toBe(false);
  });

  it("starts a new chat when storage is blocked, and still reloads", () => {
    const blocked = {
      getItem: () => { throw new Error("blocked"); },
      setItem: () => { throw new Error("blocked"); },
      removeItem: () => { throw new Error("blocked"); },
    };
    const reload = vi.fn();
    expect(() => reloadKeepingChat(reload, blocked)).not.toThrow();
    expect(reload).toHaveBeenCalledOnce();
    expect(takeChatRestore(blocked)).toBe(false);
    expect(takeChatRestore(null)).toBe(false);
  });
});
