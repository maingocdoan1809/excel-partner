import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { createRequire } from "node:module";
import { afterEach, describe, expect, it } from "vitest";
// @ts-expect-error plain ESM helper without types
import { prepareRelease, verifyRelease } from "../bin/deployment.mjs";

const store = createRequire(import.meta.url)(path.resolve("bin/runtime-store.cjs"));
const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });

const put = (root: string, file: string, value: string) => { const target = path.join(root, file); fs.mkdirSync(path.dirname(target), { recursive: true }); fs.writeFileSync(target, value); };
/** The files a release needs, as the package would hold them. */
function product(root: string, version: string, shrinkwrap: string) {
  put(root, "package.json", JSON.stringify({ name: "office-partner", version, type: "module" }));
  for (const file of ["bin/office-partner.mjs", "dist/bridge/src/index.js", "gateway/dist/server.mjs", "frontend/dist/taskpane.js"]) put(root, file, `// ${version}`);
  put(root, "bin/runtime-store.cjs", "module.exports = {};");
  put(root, "bin/runtime-manifest.json", "{}");
  put(root, "npm-shrinkwrap.json", shrinkwrap);
}
function tarball(base: string, version: string, shrinkwrap: string) {
  const source = path.join(base, "pack-" + version);
  product(path.join(source, "package"), version, shrinkwrap);
  const file = path.join(base, `office-partner-${version}.tgz`);
  const tar = process.platform === "win32" ? path.join(process.env.SystemRoot ?? "C:\\Windows", "System32", "tar.exe") : "tar";
  execFileSync(tar, ["-czf", file, "-C", source, "package"]);
  return file;
}
/** A running release laid out the way npm installs one: the product beside its dependencies in node_modules. */
async function runningRelease() {
  const base = fs.mkdtempSync(path.join(os.tmpdir(), "release-reuse-")); dirs.push(base);
  const dataDir = path.join(base, "owned");
  const install = async (prefix: string) => {
    product(path.join(prefix, "node_modules", "office-partner"), "1.0.0", '{"lock":1}');
    put(prefix, "node_modules/dependency/index.js", "module.exports = 'dependency';");
    put(prefix, "node_modules/@scope/other/index.js", "module.exports = 'other';");
  };
  const first = await prepareRelease("office-partner@1.0.0", { dataDir, install });
  const deployment = store.currentDeployment(dataDir); deployment.active = first.id; store.atomic(path.join(dataDir, "deployment.json"), deployment);
  return { base, dataDir, first };
}
const sameFile = (a: string, b: string) => fs.statSync(a, { bigint: true }).ino === fs.statSync(b, { bigint: true }).ino;

describe("an update that keeps the same dependencies", () => {
  it("hard-links the dependency files of the running release instead of installing them", async () => {
    const { base, dataDir, first } = await runningRelease();
    const next = await prepareRelease(tarball(base, "1.0.1", '{"lock":1}'), { dataDir });
    expect(next.version).toBe("1.0.1");
    for (const file of ["node_modules/dependency/index.js", "node_modules/@scope/other/index.js"]) {
      expect(fs.readFileSync(path.join(next.directory, file), "utf8")).toContain("module.exports");
      expect(sameFile(path.join(first.directory, file), path.join(next.directory, file)), file).toBe(true);
    }
    expect(fs.readFileSync(path.join(next.root, "package.json"), "utf8")).toContain("1.0.1");
    expect(fs.readFileSync(path.join(next.root, "bin", "office-partner.mjs"), "utf8")).toContain("1.0.1");
    expect(() => verifyRelease(next, true)).not.toThrow();
    expect(fs.readdirSync(path.join(dataDir, "releases")).filter(name => name.startsWith(".install-"))).toEqual([]);
  });

  it("keeps the new release when the old one is deleted", async () => {
    const { base, dataDir, first } = await runningRelease();
    const next = await prepareRelease(tarball(base, "1.0.1", '{"lock":1}'), { dataDir });
    fs.rmSync(first.directory, { recursive: true, force: true });
    expect(() => verifyRelease(next, true)).not.toThrow();
    expect(fs.readFileSync(path.join(next.directory, "node_modules/dependency/index.js"), "utf8")).toContain("dependency");
  });

  it("installs the dependencies again when the lock file changed", async () => {
    const { base, dataDir, first } = await runningRelease();
    let installed = 0;
    const install = async (prefix: string) => { installed++; product(path.join(prefix, "node_modules", "office-partner"), "1.0.1", '{"lock":2}'); put(prefix, "node_modules/dependency/index.js", "module.exports = 'newer';"); };
    const next = await prepareRelease("office-partner@1.0.1", { dataDir, install });
    expect(installed).toBe(1);
    expect(sameFile(path.join(first.directory, "node_modules/dependency/index.js"), path.join(next.directory, "node_modules/dependency/index.js"))).toBe(false);
    void base;
  });
});
