import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
// @ts-expect-error plain ESM module without type declarations
import { updateTarget } from "../bin/manage.mjs";
// @ts-expect-error CommonJS module without type declarations
import store from "../bin/runtime-store.cjs";

const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); delete process.env.npm_config_registry; });
function tempDir() { const dir = fs.mkdtempSync(path.join(os.tmpdir(), "op-update-")); dirs.push(dir); return dir; }
function release(version: string) { const root = tempDir(); fs.writeFileSync(path.join(root, "package.json"), JSON.stringify({ name: "office-partner", version })); return root; }
function registry(latest: string, deprecated: string[] = []) {
  const asked: string[] = [];
  return { asked, view: async (spec: string, field: string) => { asked.push(`${spec} ${field}`); return field === "version" ? latest : deprecated.includes(spec) ? "broken release" : undefined; } };
}
const open = { autoUpdate: true };

describe("choosing what an update installs", () => {
  it("installs the latest release when it is newer", async () => {
    expect(await updateTarget(release("0.1.23"), undefined, { ...registry("0.1.24"), policy: open })).toBe("office-partner@0.1.24");
  });

  it("refuses an older release unless asked to go back, and refuses a deprecated one", async () => {
    await expect(updateTarget(release("0.1.23"), undefined, { ...registry("0.1.22"), policy: open })).rejects.toThrow("UPDATE_DOWNGRADE_REFUSED");
    expect(await updateTarget(release("0.1.23"), "office-partner@0.1.22", { ...registry("0.1.24"), allowDowngrade: true, policy: open })).toBe("office-partner@0.1.22");
    await expect(updateTarget(release("0.1.23"), undefined, { ...registry("0.1.24", ["office-partner@0.1.24"]), policy: open })).rejects.toThrow("UPDATE_VERSION_DEPRECATED");
  });

  it("follows the admin policy: off, pinned (also backwards) and its registry", async () => {
    await expect(updateTarget(release("0.1.23"), undefined, { ...registry("0.1.24"), policy: { autoUpdate: false } })).rejects.toThrow("UPDATE_DISABLED_BY_POLICY");
    await expect(updateTarget(release("0.1.23"), undefined, { ...registry("0.1.24"), policy: { autoUpdate: false, error: "POLICY_INVALID" } })).rejects.toThrow("UPDATE_DISABLED_POLICY_INVALID");
    const pinned = registry("0.1.24");
    expect(await updateTarget(release("0.1.23"), undefined, { ...pinned, policy: { autoUpdate: true, pinVersion: "0.1.21", registry: "https://npm.example.test/" } })).toBe("office-partner@0.1.21");
    expect(pinned.asked).toEqual([]);
    expect(process.env.npm_config_registry).toBe("https://npm.example.test/");
    await expect(updateTarget(release("0.1.23"), "office-partner@0.1.24", { ...pinned, policy: { autoUpdate: true, pinVersion: "0.1.21" } })).rejects.toThrow("UPDATE_PINNED");
  });
});

describe("reading the update policy", () => {
  it("allows updates without a file and turns them off for a broken one", () => {
    const dir = tempDir(), file = path.join(dir, "policy.json");
    expect(store.readPolicy({ OFFICE_PARTNER_POLICY_FILE: file })).toEqual({ autoUpdate: true });
    fs.writeFileSync(file, "{ not json");
    expect(store.readPolicy({ OFFICE_PARTNER_POLICY_FILE: file })).toMatchObject({ autoUpdate: false, error: "POLICY_INVALID" });
    fs.writeFileSync(file, "﻿" + JSON.stringify({ pinVersion: "0.1.23", registry: "https://npm.example.test/" }));
    expect(store.readPolicy({ OFFICE_PARTNER_POLICY_FILE: file })).toEqual({ autoUpdate: true, pinVersion: "0.1.23", registry: "https://npm.example.test/" });
  });
});

describe("locks left behind", () => {
  it("are taken over after a restart, when too old, or when empty; a live one is kept", () => {
    const dir = tempDir(), file = path.join(dir, "x.lock"), now = Date.now();
    fs.writeFileSync(file, JSON.stringify({ pid: process.pid, created: now, boot: Math.round((now - os.uptime() * 1000) / 60000) }));
    expect(store.staleLock(file, now)).toBe(false);
    fs.writeFileSync(file, JSON.stringify({ pid: process.pid, created: now, boot: 1 }));
    expect(store.staleLock(file, now)).toBe(true);
    fs.writeFileSync(file, JSON.stringify({ pid: process.pid, created: now - 3 * 60 * 60_000 }));
    expect(store.staleLock(file, now)).toBe(true);
    fs.writeFileSync(file, "");
    expect(store.staleLock(file, now)).toBe(false);
    expect(store.staleLock(file, now + 3 * 60_000)).toBe(true);
  });
});
