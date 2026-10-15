import { describe, expect, it } from "vitest";
import { SessionManager } from "../bridge/src/session-manager.js";

describe("SessionManager", () => {
  it("creates a session with a unique token", () => {
    const manager = new SessionManager({ ttlMs: 60_000 });
    const session = manager.create("conn-1", null);
    expect(session.sessionToken.length).toBeGreaterThanOrEqual(16);
    expect(manager.getByToken(session.sessionToken)?.id).toBe(session.id);
    expect(manager.getByConnection("conn-1")?.id).toBe(session.id);
  });

  it("replaces the previous session when the same workbook reconnects", () => {
    const manager = new SessionManager({ ttlMs: 60_000 });
    const first = manager.create("conn-1", null);
    const second = manager.create("conn-1", null);
    expect(second.id).not.toBe(first.id);
    expect(manager.get(first.id)).toBeUndefined();
    expect(manager.getByToken(first.sessionToken)).toBeUndefined();
    expect(manager.size).toBe(1);
  });

  it("evicts the oldest session when over the limit", () => {
    const manager = new SessionManager({ ttlMs: 60_000, maxSessions: 2 });
    const a = manager.create("a", null);
    const b = manager.create("b", null);
    const c = manager.create("c", null);
    expect(manager.get(a.id)).toBeUndefined();
    expect(manager.get(b.id)).toBeDefined();
    expect(manager.get(c.id)).toBeDefined();
  });

  it("reaps idle sessions", () => {
    const manager = new SessionManager({ ttlMs: 10 });
    manager.create("conn-1", null);
    expect(manager.reapIdle(Date.now() + 1000)).toBe(1);
    expect(manager.size).toBe(0);
  });
});
