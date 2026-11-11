import { describe, expect, it, vi } from "vitest";
// @ts-expect-error Ships as plain JavaScript beside the gateway CLI.
import { decideRecorded, processOwner } from "../gateway/bin/process-record.mjs";

const lookup = (stdout: string, status = 0) => vi.fn(() => ({ status, stdout, stderr: "" }));
const windows = { platform: "win32" as const };
const isBridge = (line: string) => /dist[\\/]bridge[\\/]src[\\/]index\.js/.test(line);

describe("who a recorded process number belongs to", () => {
  it("recognises this product's Node process", () => {
    const run = lookup(JSON.stringify({ Name: "node.exe", CommandLine: '"C:\\Program Files\\nodejs\\node.exe" "C:\\data\\releases\\ab\\dist\\bridge\\src\\index.js"' }));
    expect(processOwner(4321, isBridge, { ...windows, run })).toBe("ours");
    expect(run).toHaveBeenCalledWith("powershell.exe", expect.arrayContaining(["-NoProfile", "-NonInteractive"]), expect.objectContaining({ windowsHide: true }));
  });

  it("D1 sees that a number reused after a restart now belongs to another program", () => {
    expect(processOwner(4321, isBridge, { ...windows, run: lookup(JSON.stringify({ Name: "OneDrive.exe", CommandLine: '"C:\\Program Files\\OneDrive\\OneDrive.exe" /background' })) })).toBe("other");
    // Node, but not this product's file.
    expect(processOwner(4321, isBridge, { ...windows, run: lookup(JSON.stringify({ Name: "node.exe", CommandLine: "node C:\\tools\\other-server.js" })) })).toBe("other");
    // The process ended between the check and the lookup.
    expect(processOwner(4321, isBridge, { ...windows, run: lookup("") })).toBe("other");
  });

  it("says unknown when the lookup cannot run, so nothing is assumed", () => {
    expect(processOwner(4321, isBridge, { ...windows, run: lookup("", 1) })).toBe("unknown");
    expect(processOwner(4321, isBridge, { ...windows, run: lookup("not json") })).toBe("unknown");
    expect(processOwner(4321, isBridge, { ...windows, run: vi.fn(() => ({ error: new Error("spawn powershell.exe ENOENT"), status: null, stdout: "" })) })).toBe("unknown");
    expect(processOwner(4321, isBridge, { platform: "linux", run: lookup("{}") })).toBe("unknown");
    expect(processOwner(-1, isBridge, { ...windows, run: lookup("{}") })).toBe("unknown");
  });
});

describe("what to do with a leftover process record", () => {
  const waiting = (healthy: boolean) => vi.fn(async () => healthy);

  it("D1 starts at once when the number belongs to another program, without waiting", async () => {
    const waitHealthy = waiting(false);
    await expect(decideRecorded("other", { waitHealthy })).resolves.toBe("discard");
    expect(waitHealthy).not.toHaveBeenCalled();
  });

  it("accepts a service that is only slow to answer", async () => {
    for (const owner of ["ours", "unknown"]) await expect(decideRecorded(owner, { waitHealthy: waiting(true) })).resolves.toBe("ready");
  });

  it("reports our own stuck process, and replaces it only when a repair was asked for", async () => {
    await expect(decideRecorded("ours", { waitHealthy: waiting(false) })).resolves.toBe("fail");
    await expect(decideRecorded("ours", { repair: true, waitHealthy: waiting(false) })).resolves.toBe("kill");
  });

  it("treats an unidentified program that never answers as a stale record", async () => {
    await expect(decideRecorded("unknown", { waitHealthy: waiting(false) })).resolves.toBe("discard");
    await expect(decideRecorded("unknown", { repair: true, waitHealthy: waiting(false) })).resolves.toBe("discard");
  });
});
