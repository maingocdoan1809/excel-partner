import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ReferenceFiles, isNetworkPath, localPath, workbookFolderFromUrl } from "../bridge/src/reference-files.js";

afterEach(() => { vi.restoreAllMocks(); });

const NETWORK = ["\\\\attacker\\share\\x.xlsx", "//attacker/share/x.xlsx", "\\\\?\\UNC\\attacker\\share\\x.xlsx", "\\\\.\\pipe\\x",
  "\\\\attacker@SSL@443\\DavWWWRoot\\x.xlsx", "/\\attacker\\share\\x.xlsx", "  \\\\attacker\\share\\x.xlsx"];

describe("network paths never reach the file system", () => {
  it("recognises UNC, device, long-path and WebDAV forms by their text", () => {
    for (const input of NETWORK) expect(isNetworkPath(input), input).toBe(true);
    for (const input of ["C:\\Users\\me\\x.xlsx", "/home/me/x.xlsx", "relative\\x.xlsx"]) expect(isNetworkPath(input), input).toBe(false);
  });

  it("refuses a model-supplied network path before any fs call, in every entry point", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "network-paths-"));
    try {
      const references = new ReferenceFiles([dir]);
      const realpath = vi.spyOn(fs, "realpathSync");
      const stat = vi.spyOn(fs, "statSync");
      for (const input of NETWORK) {
        expect(() => references.resolve(input)).toThrow(/network location/);
        expect(() => references.permissionTarget(input)).toThrow(/network location/);
        expect(() => references.permissionTargetAny(input)).toThrow(/network location/);
        expect(() => references.referenceTarget(input)).toThrow(/network location/);
        expect(() => references.readWorkbook(input)).toThrow(/network location/);
        expect(() => workbookFolderFromUrl(input)).toThrow(/network location/);
      }
      expect(() => workbookFolderFromUrl("file:////attacker/share/x.xlsx")).toThrow();
      expect(realpath).not.toHaveBeenCalled();
      expect(stat).not.toHaveBeenCalled();
    } finally { fs.rmSync(dir, { recursive: true, force: true }); }
  });

  it("allows a network path inside a network folder the user configured", () => {
    expect(localPath("\\\\files\\team\\book.xlsx", ["\\\\files\\team"])).toBe("\\\\files\\team\\book.xlsx");
    expect(() => localPath("\\\\files\\teamless\\book.xlsx", ["\\\\files\\team"])).toThrow(/network location/);
    expect(() => localPath("\\\\other\\team\\book.xlsx", ["C:\\work"])).toThrow(/network location/);
  });
});
