import fs from "node:fs";
import path from "node:path";
import { describe, expect, it, vi } from "vitest";
// @ts-expect-error The bridge package ships as plain JavaScript, independently of the TS build.
import { NEW_PACKAGE, ensureOfficePartner, forward, globalRoot, installSpec, newCliPath, ownGlobalRoot } from "../legacy/excel-partner/bin/migrate.mjs";

const root = path.join("/npm", "global");
const cli = newCliPath(root);
const modules = path.resolve("/npm", "lib", "node_modules");

/** A fake `spawnSync`: `npm root` answers `root`; an install makes the CLI appear; anything else is recorded. */
function world(options: { installed?: boolean; installStatus?: number; npmRoot?: string | null; cliStatus?: number } = {}) {
  let installed = options.installed ?? false;
  const calls: Array<{ command: string; args: string[] }> = [];
  const run = vi.fn((command: string, args: string[]) => {
    calls.push({ command, args });
    if (args[0] === "root") return options.npmRoot === null ? { status: 1, stdout: "" } : { status: 0, stdout: `${options.npmRoot ?? root}\n` };
    if (args[0] === "install") {
      if ((options.installStatus ?? 0) === 0) installed = true;
      return { status: options.installStatus ?? 0 };
    }
    return { status: options.cliStatus ?? 0 };
  });
  return { run, calls, deps: { run: run as never, exists: () => installed, platform: "linux" as const, env: {}, log: vi.fn(), error: vi.fn(), nodePath: "node" } };
}

describe("the excel-partner package that stays on npm after the rename", () => {
  it("is a package of its own: the old name, a newer version than any release before, nothing but the forwarder", () => {
    const manifest = JSON.parse(fs.readFileSync("legacy/excel-partner/package.json", "utf8"));
    const main = JSON.parse(fs.readFileSync("package.json", "utf8"));
    expect(manifest.name).toBe("excel-partner");
    expect(main.name).toBe(NEW_PACKAGE);
    expect(manifest.bin).toEqual({ "excel-partner": "bin/excel-partner.mjs" });
    expect(manifest.dependencies).toBeUndefined();
    expect(manifest.files).toEqual(["bin/", "README.md"]);
    const [a, b, c] = (manifest.version as string).split(".").map(Number);
    expect(a! > 0 || b! > 1 || c! > 21).toBe(true); // newer than 0.1.21, the last release under the old name
    expect(Object.keys(main.bin)).toEqual(["office-partner", "office-partner-doctor", "agent-gateway"]); // no `excel-partner` here: that would clash with this package
  });

  it("finds the global folder and the new command line in it", () => {
    expect(globalRoot(world().run as never, "linux")).toBe(root);
    expect(globalRoot(world({ npmRoot: null }).run as never, "linux")).toBeNull();
    expect(cli).toBe(path.join(root, "office-partner", "bin", "office-partner.mjs"));
  });

  it("installs Office Partner once when it is missing, then hands the command over", () => {
    const { deps, calls } = world();
    const code = forward(["setup", "all"], deps);
    expect(code).toBe(0);
    expect(calls.map((c) => c.args.slice(0, 2).join(" "))).toEqual(["root --global", "install --global", `${cli} setup`]);
    expect(calls[1]!.args).toEqual(["install", "--global", "office-partner@latest"]);
    expect(calls[2]).toEqual({ command: "node", args: [cli, "setup", "all"] });
    expect(deps.log).toHaveBeenCalledWith(expect.stringContaining("now called Office Partner"));
  });

  it("does not install again once it is there, and passes the exit code through", () => {
    const { deps, calls } = world({ installed: true, cliStatus: 3 });
    expect(forward(["status"], deps)).toBe(3);
    expect(calls.some((c) => c.args[0] === "install")).toBe(false);
    expect(deps.log).not.toHaveBeenCalled();
  });

  it("explains what to do when npm is missing or the install fails, and runs nothing else", () => {
    const missing = world({ npmRoot: null });
    expect(forward(["start"], missing.deps)).toBe(1);
    expect(missing.deps.error).toHaveBeenCalledWith(expect.stringContaining("npm was not found"));
    const failing = world({ installStatus: 1 });
    expect(forward(["start"], failing.deps)).toBe(1);
    expect(failing.deps.error).toHaveBeenCalledWith(expect.stringContaining("npm install -g office-partner"));
    expect(failing.calls.some((c) => c.command === "node")).toBe(false);
    expect(() => ensureOfficePartner({ ...failing.deps, exists: () => false })).toThrow(/Could not install/);
  });

  it("takes a version or a tarball from OFFICE_PARTNER_SPEC, and refuses anything the Windows shell could act on", () => {
    expect(installSpec({})).toBe("office-partner@latest");
    expect(installSpec({ OFFICE_PARTNER_SPEC: "office-partner@0.2.0" })).toBe("office-partner@0.2.0");
    expect(installSpec({ OFFICE_PARTNER_SPEC: "C:\\Users\\me\\My Files\\office-partner-0.2.0.tgz" })).toContain("My Files");
    for (const bad of ["office-partner & calc", "a|b", "x;y", "$(whoami)", "`id`", "a>b"]) expect(() => installSpec({ OFFICE_PARTNER_SPEC: bad }), bad).toThrow(/not allowed/);
    const spaced = world();
    ensureOfficePartner({ ...spaced.deps, env: { OFFICE_PARTNER_SPEC: "/tmp/my files/op.tgz" } });
    expect(spaced.calls[1]!.args).toEqual(["install", "--global", '"/tmp/my files/op.tgz"']);
  });

  it("also looks where this copy itself sits, because npm hides what looks like a secret in the path it prints", () => {
    expect(ownGlobalRoot(path.join(modules, "excel-partner", "bin"))).toBe(path.resolve(modules));
    expect(ownGlobalRoot(path.join("/work", "legacy", "excel-partner", "bin"))).toBeNull(); // a source checkout, not a global install
    const ownCli = newCliPath(modules);
    let installed = false;
    const calls: string[] = [];
    const run = vi.fn((_command: string, args: string[]) => {
      calls.push(args[0]!);
      if (args[0] === "root") return { status: 0, stdout: "/home/me/***/prefix/lib/node_modules\n" }; // what npm printed
      if (args[0] === "install") installed = true;
      return { status: 0 };
    });
    const code = forward(["help"], { run: run as never, exists: (file: string) => installed && file === ownCli, platform: "linux", env: {}, log: () => {}, error: () => {}, nodePath: "node", selfDir: path.join(modules, "excel-partner", "bin") });
    expect(code).toBe(0);
    expect(calls).toEqual(["root", "install", ownCli]);
    expect(run).toHaveBeenLastCalledWith("node", [ownCli, "help"], expect.anything());
  });

  it("uses npm.cmd through the shell on Windows", () => {
    const { run } = world();
    globalRoot(run as never, "win32");
    expect(run.mock.calls[0]).toEqual(["npm.cmd", ["root", "--global"], expect.objectContaining({ shell: true })]);
  });
});
