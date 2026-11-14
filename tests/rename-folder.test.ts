import fs from "node:fs";
import { createRequire } from "node:module";
import { afterEach, describe, expect, it, vi } from "vitest";

const store = createRequire(import.meta.url)("../bin/runtime-store.cjs") as { renameFolder(from: string, to: string): void };
const denied = () => Object.assign(new Error("EPERM: operation not permitted, rename"), { code: "EPERM" });

afterEach(() => vi.restoreAllMocks());

describe("renaming a folder that was just written", () => {
  it.runIf(process.platform === "win32")("tries again while the Windows virus scanner holds the new files, then succeeds", () => {
    const rename = vi.spyOn(fs, "renameSync").mockImplementationOnce(() => { throw denied(); }).mockImplementationOnce(() => { throw denied(); }).mockImplementation(() => {});
    store.renameFolder("from", "to");
    expect(rename).toHaveBeenCalledTimes(3);
  });

  it("lets any other error through at once", () => {
    const rename = vi.spyOn(fs, "renameSync").mockImplementation(() => { throw Object.assign(new Error("ENOENT"), { code: "ENOENT" }); });
    expect(() => store.renameFolder("from", "to")).toThrow("ENOENT");
    expect(rename).toHaveBeenCalledTimes(1);
  });
});
