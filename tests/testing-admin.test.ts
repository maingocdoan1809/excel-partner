import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it, vi } from "vitest";
import { SiteKnowledge, originOfKey, siteKey } from "../bridge/src/site-knowledge.js";
import { TestSessions } from "../bridge/src/test-sessions.js";
import { TestScripts } from "../bridge/src/test-scripts.js";
import { TestingAdmin, TestingAdminError } from "../bridge/src/testing-admin.js";
import type { LoginOptions } from "../bridge/src/test-login.js";

const script = `import { test, expect } from '@playwright/test';
test('a', async ({ page }) => {
  await page.goto('/x');
  await expect(page).toHaveURL(/x/);
});`;

function setup(login?: (options: LoginOptions) => Promise<{ origins: string[] }>, timeoutMs?: number) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ep-admin-"));
  const sessions = new TestSessions(path.join(dir, "sessions"));
  const knowledge = new SiteKnowledge(path.join(dir, "knowledge"));
  const scripts = new TestScripts(path.join(dir, "scripts"));
  const open = vi.fn();
  const admin = new TestingAdmin(sessions, knowledge, scripts, login, open, timeoutMs);
  return { dir, sessions, knowledge, scripts, admin, open };
}

/** A login that behaves like the real one: opens, waits for the person, then saves. */
const fakeLogin = (calls: LoginOptions[] = []) => async (options: LoginOptions) => {
  calls.push(options);
  options.out("opened");
  await options.waitForPerson();
  options.store.save(options.name, { cookies: [] }, { origins: [options.url], savedAt: "2026-10-30T09:00:00.000Z", url: options.url });
  return { origins: [options.url] };
};

describe("site folders", () => {
  it("turn back into the address they came from", () => {
    expect(originOfKey("staging.example.com")).toBe("http://staging.example.com");
    expect(originOfKey("localhost_8081")).toBe("http://localhost:8081");
    expect(siteKey(originOfKey("localhost_8081")!)).toBe("localhost_8081");
    expect(originOfKey("../etc")).toBeNull();
    expect(originOfKey("a/b")).toBeNull();
    for (const bad of ["..", ".", "a..b", ".hidden", "-x", ""]) expect(originOfKey(bad), bad).toBeNull();
  });
});

describe("the testing screen", () => {
  it("lists sign-ins, notes and scripts by system, and where their folders are", () => {
    const { admin, sessions, knowledge, scripts } = setup();
    const origin = "https://staging.example.com";
    sessions.save("admin", {}, { origins: [origin], savedAt: "2026-10-29T10:00:00.000Z", url: origin });
    knowledge.save(origin, { name: "login", kind: "login", summary: "Use the saved session", content: "Open /web/login." });
    scripts.save(origin, { caseId: "TC-001", title: "Create", sheet: "S", row: 4, outcome: "pass", evidence: "Saw the banner appear after saving.", script });
    const overview = admin.overview();
    expect(overview.signIns).toEqual([expect.objectContaining({ name: "admin" })]);
    expect(overview.notes).toEqual([{ site: "staging.example.com", notes: [expect.objectContaining({ name: "login", summary: "Use the saved session" })] }]);
    expect(overview.scripts).toEqual([{ site: "staging.example.com", scripts: [expect.objectContaining({ caseId: "TC-001", outcome: "pass" })] }]);
    expect(overview.pendingSignIn).toBeNull();
    expect(overview.folders.scripts).toBe(scripts.dir);
    expect(overview.signIns.map(item => Object.keys(item).sort())).toEqual([["name", "origins", "savedAt", "url"]]);
  });

  it("views and deletes a note and a script, and refuses anything that is not a known site folder", () => {
    const { admin, knowledge, scripts } = setup();
    const origin = "https://staging.example.com";
    knowledge.save(origin, { name: "login", kind: "login", summary: "s", content: "Open /web/login." });
    scripts.save(origin, { caseId: "TC-001", title: "Create", sheet: "S", row: 4, outcome: "pass", evidence: "Saw the banner appear after saving.", script });
    expect(admin.viewNote("staging.example.com", "login").content).toContain("Open /web/login.");
    expect(admin.viewScript("staging.example.com", "TC-001")).toContain("// case: TC-001");
    expect(admin.deleteNote("staging.example.com", "login")).toBe(true);
    expect(admin.deleteScript("staging.example.com", "TC-001")).toBe(true);
    expect(admin.overview().notes).toEqual([]);
    for (const bad of ["../x", "a/b", ".."]) {
      expect(() => admin.viewNote(bad, "login")).toThrow(TestingAdminError);
      expect(() => admin.deleteScript(bad, "TC-001")).toThrow(TestingAdminError);
    }
    expect(() => admin.viewNote("staging.example.com", "../../x")).toThrow(TestingAdminError);
  });

  it("deletes a saved sign-in by a valid name only", () => {
    const { admin, sessions } = setup();
    sessions.save("admin", {}, { origins: ["https://a.test"], savedAt: "x", url: "https://a.test" });
    expect(() => admin.deleteSignIn("../admin")).toThrow(TestingAdminError);
    expect(admin.deleteSignIn("admin")).toBe(true);
    expect(admin.deleteSignIn("admin")).toBe(false);
  });
});

describe("signing in from a button", () => {
  it("opens a browser, waits for 'I'm signed in', then saves the session", async () => {
    const calls: LoginOptions[] = [];
    const { admin } = setup(fakeLogin(calls));
    expect(await admin.startSignIn({ name: "admin", url: "https://staging.example.com/web/login" })).toEqual({ name: "admin", url: "https://staging.example.com" });
    expect(admin.overview().pendingSignIn).toEqual({ name: "admin", url: "https://staging.example.com" });
    expect(calls[0]).toMatchObject({ name: "admin", url: "https://staging.example.com" });
    expect(await admin.finishSignIn()).toEqual({ name: "admin", origins: ["https://staging.example.com"] });
    expect(admin.overview()).toMatchObject({ pendingSignIn: null, signIns: [{ name: "admin" }] });
  });

  it("allows only one open window, and a wrong name or address never opens one", async () => {
    const calls: LoginOptions[] = [];
    const { admin } = setup(fakeLogin(calls));
    await expect(admin.startSignIn({ name: "Bad Name", url: "https://a.test" })).rejects.toThrow(/lowercase/);
    await expect(admin.startSignIn({ name: "admin", url: "file:///etc/passwd" })).rejects.toThrow(/address/);
    expect(calls).toHaveLength(0);
    await admin.startSignIn({ name: "admin", url: "https://a.test" });
    await expect(admin.startSignIn({ name: "other", url: "https://b.test" })).rejects.toThrow(/already open/);
    expect(admin.cancelSignIn()).toBe(true);
    await vi.waitFor(() => expect(admin.overview().pendingSignIn).toBeNull());
    expect(admin.overview().signIns).toEqual([]);
    expect(admin.cancelSignIn()).toBe(false);
  });

  it("reports a browser that could not start and leaves nothing pending", async () => {
    const { admin } = setup(async () => { throw new Error("Could not start Chrome or Edge."); });
    await expect(admin.startSignIn({ name: "admin", url: "https://a.test" })).rejects.toThrow(/Could not start Chrome/);
    expect(admin.overview().pendingSignIn).toBeNull();
    await expect(admin.finishSignIn()).rejects.toThrow(/No sign-in window/);
  });

  it("closes a window that was left open too long", async () => {
    const { admin } = setup(fakeLogin(), 20);
    await admin.startSignIn({ name: "admin", url: "https://a.test" });
    await vi.waitFor(() => expect(admin.overview().pendingSignIn).toBeNull());
    expect(admin.overview().signIns).toEqual([]);
  });

  it("opens a folder only on Windows, and only the three known ones", () => {
    const { admin, open, sessions } = setup();
    if (process.platform === "win32") { expect(admin.folder("sessions")).toBe(sessions.dir); expect(open).toHaveBeenCalledWith(sessions.dir); }
    else { expect(() => admin.folder("sessions")).toThrow(sessions.dir); expect(open).not.toHaveBeenCalled(); }
  });
});
