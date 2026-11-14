import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
// @ts-expect-error bin/ is plain JavaScript without declarations
import { importFromRelease } from "../bin/doctor-probes.mjs";

describe("doctor dependencies", () => {
  it("loads them from the release only, never from a node_modules folder above it", async () => {
    const outer = fs.mkdtempSync(path.join(os.tmpdir(), "doctor-imports-"));
    try {
      const planted = path.join(outer, "node_modules", "cross-spawn");
      fs.mkdirSync(planted, { recursive: true });
      fs.writeFileSync(path.join(planted, "package.json"), JSON.stringify({ name: "cross-spawn", main: "index.js" }));
      fs.writeFileSync(path.join(planted, "index.js"), "module.exports = 'planted';");
      const release = path.join(outer, "release");
      fs.mkdirSync(release);
      fs.writeFileSync(path.join(release, "package.json"), "{}");
      await expect(importFromRelease(release, "cross-spawn")).rejects.toThrow(/DEPENDENCY_MISSING/);
      expect(typeof (await importFromRelease(process.cwd(), "cross-spawn")).default).toBe("function");
    } finally { fs.rmSync(outer, { recursive: true, force: true }); }
  });
});
