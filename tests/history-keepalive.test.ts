import { describe, expect, it, vi } from "vitest";
import { BridgeClient } from "../addin/taskpane/bridge-client.js";
import type { HistorySave, StoredSession } from "../protocol/src/history.js";

const save = (text: string): HistorySave => ({ baseRevision: 0, generation: 0, operationId: "op",
  session: { id: "c", title: "t", provider: "claude", startedAt: 1, updatedAt: 2, messages: [{ kind: "user", text }], changes: [] } satisfies StoredSession });

describe("history save while the pane closes", () => {
  // Browsers reject keepalive requests above 64 KiB, so a long chat must wait in the outbox instead.
  it("asks the browser to finish a small save after the pane closes, but not a large one", async () => {
    const client = Object.create(BridgeClient.prototype) as { sessionFetch: ReturnType<typeof vi.fn>; historySave(input: HistorySave, keepalive?: boolean): Promise<unknown> };
    client.sessionFetch = vi.fn(async () => ({ revision: 1, generation: 0 }));
    await client.historySave(save("short"), true);
    expect(client.sessionFetch).toHaveBeenLastCalledWith("/session/history", expect.objectContaining({ keepalive: true }));
    await client.historySave(save("x".repeat(30_000)), true);
    expect(client.sessionFetch).toHaveBeenLastCalledWith("/session/history", expect.objectContaining({ keepalive: false }));
    await client.historySave(save("short"), false);
    expect(client.sessionFetch).toHaveBeenLastCalledWith("/session/history", expect.objectContaining({ keepalive: false }));
  });
});
