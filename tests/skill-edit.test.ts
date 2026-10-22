import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { strToU8, zipSync } from "fflate";
import { UserSkills } from "../bridge/src/skill-upload.js";
import { SkillStore } from "../bridge/src/skills.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";

let root: string;
let uploads: UserSkills;
const manifest = "---\nname: report\ndescription: Original\n---\nKeep formulas.\n";
beforeEach(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), "skill-edit-"));
  uploads = new UserSkills(path.join(root, "skills"));
  uploads.install("report.zip", Buffer.from(zipSync({
    "SKILL.md": strToU8(manifest), "references/guide.md": strToU8("Old guide"),
    "tools/check.py": strToU8("print('old')"), "templates/example.xlsx": new Uint8Array([80, 75, 0, 255]),
  })), () => false);
});
afterEach(() => { vi.restoreAllMocks(); fs.rmSync(root, { recursive: true, force: true }); });

describe("installed skill editing", () => {
  it("edits instructions, references and scripts without altering binary templates", () => {
    expect(uploads.textFiles("report")).toEqual(["SKILL.md", "references/guide.md", "tools/check.py"]);
    const before = fs.readFileSync(path.join(uploads.root, "report/templates/example.xlsx"));
    for (const file of uploads.textFiles("report")) {
      const document = uploads.readTextFile("report", file);
      const content = file === "SKILL.md" ? manifest.replace("Original", "Updated") : "Updated text";
      const saved = uploads.updateTextFile("report", file, content, document.revision);
      expect(saved.content).toBe(content);
      expect(saved.revision).not.toBe(document.revision);
    }
    expect(new SkillStore([uploads.root]).get("report")?.description).toBe("Updated");
    expect(fs.readFileSync(path.join(uploads.root, "report/templates/example.xlsx"))).toEqual(before);
  });

  it("rejects stale revisions, renamed or empty manifests, and binary content without damaging files", () => {
    const document = uploads.readTextFile("report", "SKILL.md");
    for (const content of [manifest.replace("name: report", "name: other"), "---\nname: report\n---\n", "text\0binary"]) {
      expect(() => uploads.updateTextFile("report", "SKILL.md", content, document.revision)).toThrow();
      expect(uploads.readTextFile("report", "SKILL.md")).toEqual(document);
    }
    uploads.updateTextFile("report", "SKILL.md", manifest + "New rule", document.revision);
    expect(() => uploads.updateTextFile("report", "SKILL.md", manifest, document.revision)).toThrow(/changed/);
    expect(() => uploads.readTextFile("report", "templates/example.xlsx")).toThrow(/text/);
    expect(() => uploads.readTextFile("report", "../outside.md")).toThrow();
  });

  it("rejects linked files and folders that escape the installed skill", () => {
    const outside = path.join(root, "outside");
    fs.mkdirSync(outside);
    fs.writeFileSync(path.join(outside, "secret.md"), "outside");
    fs.symlinkSync(outside, path.join(uploads.root, "report", "linked"), process.platform === "win32" ? "junction" : "dir");
    expect(() => uploads.readTextFile("report", "linked/secret.md")).toThrow();
    expect(uploads.textFiles("report")).not.toContain("linked/secret.md");
  });

  it("lets the AI read exact file revisions and saves only after permission, rechecking for concurrent changes", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("workbook", null);
    session.beginTurn("request");
    const gateway = new ToolGateway(sessions, { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 }, 1000,
      undefined, false, new SkillStore([uploads.root]), null, uploads);
    const read = await gateway.invoke({ sessionId: session.id, tool: "read_skill", args: { name: "report", forEditing: true } });
    expect(read).toMatchObject({ ok: true, data: { file: "SKILL.md", content: manifest, revision: expect.any(String) } });
    if (!read.ok) throw new Error("read failed");
    const revision = (read.data as { revision: string }).revision;
    const args = { name: "report", file: "SKILL.md", content: manifest + "AI rule", revision };
    const approve = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: false } });
    expect(await gateway.invoke({ sessionId: session.id, tool: "update_skill", args })).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(uploads.readTextFile("report", "SKILL.md").revision).toBe(revision);
    approve.mockImplementation(async () => {
      uploads.updateTextFile("report", "SKILL.md", manifest + "Human rule", revision);
      return { ok: true, data: { approved: true } };
    });
    expect(await gateway.invoke({ sessionId: session.id, tool: "update_skill", args })).toMatchObject({ ok: false, error: { message: expect.stringContaining("changed") } });
    expect(uploads.readTextFile("report", "SKILL.md").content).toContain("Human rule");
    approve.mockResolvedValue({ ok: true, data: { approved: true } });
    args.revision = uploads.readTextFile("report", "SKILL.md").revision;
    expect(await gateway.invoke({ sessionId: session.id, tool: "update_skill", args })).toMatchObject({ ok: true, data: { saved: true } });
    expect(uploads.readTextFile("report", "SKILL.md").content).toBe(args.content);
    args.revision = uploads.readTextFile("report", "SKILL.md").revision;
    args.content = manifest + "Must not be saved after cancellation";
    approve.mockImplementation(async () => {
      session.abortTurn();
      session.endTurn("request");
      return { ok: true, data: { approved: true } };
    });
    expect(await gateway.invoke({ sessionId: session.id, tool: "update_skill", args })).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(uploads.readTextFile("report", "SKILL.md").content).not.toContain("Must not be saved");
  });
});
