import { afterEach, describe, expect, it, vi } from "vitest";
import { HttpToolForwarder } from "../excel_mcp/src/forwarder.js";
import { SEND_TO_CONNECTED_FILE_MAX_MS } from "../protocol/src/index.js";

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

/** A bridge that never answers; the call only ends when the forwarder gives up. */
function silentBridge() {
  vi.useFakeTimers();
  vi.stubGlobal("fetch", (_url: string, init: { signal: AbortSignal }) => new Promise((_resolve, reject) => {
    init.signal.addEventListener("abort", () => reject(new Error("aborted")));
  }));
  return new HttpToolForwarder({ bridgeUrl: "http://bridge", sessionId: "s", token: "t" });
}

describe("how long the MCP process waits for the bridge", () => {
  it("gives up on an ordinary tool after about two minutes", async () => {
    const call = silentBridge().call("read_range", {});
    await vi.advanceTimersByTimeAsync(131_000);
    expect(await call).toMatchObject({ ok: false, error: { code: "disconnected" } });
  });

  it("keeps waiting for another file's assistant, which can take ten minutes plus the approval", async () => {
    const forwarder = silentBridge();
    let settled = false;
    const call = forwarder.call("send_to_connected_file", { file: "f2", message: "x" }).then((result) => { settled = true; return result; });
    await vi.advanceTimersByTimeAsync(12.5 * 60_000);
    expect(settled).toBe(false);
    await vi.advanceTimersByTimeAsync(SEND_TO_CONNECTED_FILE_MAX_MS);
    expect(await call).toMatchObject({ ok: false, error: { code: "disconnected" } });
  });

  it("covers the longest the bridge itself can wait: approval, start and run", () => {
    expect(SEND_TO_CONNECTED_FILE_MAX_MS).toBeGreaterThan(120_000 + 20_000 + 10 * 60_000);
  });
});
