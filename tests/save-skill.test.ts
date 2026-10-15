import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { SkillStore } from "../bridge/src/skills.js";
import { UserSkills } from "../bridge/src/skill-upload.js";
import { ALL_TOOL_DEFINITIONS } from "../excel_mcp/src/tool-definitions.js";
import { SERVER_TOOLS } from "../protocol/src/index.js";

let folder: string;
let sessions: SessionManager;
let gateway: ToolGateway;
let sessionId: string;
let uploads: UserSkills;
const args = { name: "monthly-report", description: "Monthly report: \"format\"\nUse when requested.", instructions: "# Monthly report\nFormat the selected table." };

beforeEach(() => {
  folder = fs.mkdtempSync(path.join(os.tmpdir(), "save-skill-"));
  uploads = new UserSkills(path.join(folder, "skills"));
  sessions = new SessionManager({ ttlMs: 60_000 });
  sessionId = sessions.create("test-connection", null).id;
  gateway = new ToolGateway(sessions, { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 }, 1000,
    undefined, false, new SkillStore([uploads.root]), null, uploads);
});
afterEach(() => {
  fs.rmSync(folder, { recursive: true, force: true });
});

describe("save_skill MCP tool", () => {
  it("is exposed as a non-destructive server-side write tool", () => {
    expect(SERVER_TOOLS).toContain("save_skill");
    expect(ALL_TOOL_DEFINITIONS.find(tool => tool.name === "save_skill")).toMatchObject({
      annotations: { readOnlyHint: false, destructiveHint: false },
      inputSchema: { required: expect.arrayContaining(["name", "instructions"]) },
    });
  });

  it("saves in the managed folder without a pane and persists across store instances", async () => {
    const file = path.join(uploads.root, args.name, "SKILL.md");
    expect(await gateway.invoke({ sessionId, tool: "save_skill", args })).toEqual({
      ok: true, data: { name: args.name, description: args.description, path: file, command: "/monthly-report", saved: true },
    });
    expect(fs.readFileSync(file, "utf8")).toContain("name: monthly-report");
    expect(new SkillStore([uploads.root]).read(args.name)).toMatchObject({
      name: args.name, description: args.description, content: args.instructions,
    });
    expect(await gateway.invoke({ sessionId, tool: "list_skills", args: {} })).toMatchObject({
      ok: true, data: { skills: [{ name: args.name, description: args.description }] },
    });
    expect(await gateway.invoke({ sessionId, tool: "read_skill", args: { name: args.name } })).toMatchObject({
      ok: true, data: { content: args.instructions },
    });
    expect(uploads.owns(args.name, new SkillStore([uploads.root]).get(args.name)!.dir)).toBe(true);
    uploads.remove(args.name);
    expect(new SkillStore([uploads.root]).list()).toEqual([]);
  });

  it("never overwrites an existing skill", async () => {
    await gateway.invoke({ sessionId, tool: "save_skill", args });
    expect(await gateway.invoke({ sessionId, tool: "save_skill", args: { ...args, instructions: "Replacement" } })).toMatchObject({
      ok: false, error: { code: "invalid_arguments", message: expect.stringContaining("already exists") },
    });
    expect(new SkillStore([uploads.root]).read(args.name).content).toContain(args.instructions);
  });

  it("rejects arbitrary paths, reserved names, empty and oversized instructions", async () => {
    for (const invalid of [
      { ...args, name: "../escape" }, { ...args, name: "C:/escape" }, { ...args, name: "con" },
      { ...args, path: "elsewhere/SKILL.md" }, { ...args, instructions: "  " },
      { ...args, instructions: "X".repeat(200_001) }, { ...args, instructions: "界".repeat(80_000) },
    ]) {
      expect(await gateway.invoke({ sessionId, tool: "save_skill", args: invalid })).toMatchObject({
        ok: false, error: { code: "invalid_arguments" },
      });
    }
    expect(fs.existsSync(uploads.root)).toBe(false);
  });

  it("requires an authenticated session before writing", async () => {
    expect(await gateway.invoke({ sessionId: "missing", tool: "save_skill", args })).toMatchObject({
      ok: false, error: { code: "disconnected" },
    });
    expect(fs.existsSync(uploads.root)).toBe(false);
  });
});
