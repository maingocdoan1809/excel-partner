import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { SkillError, SkillStore, expandSkillPrompt, parseFrontmatter, projectTemplateReminder, skillTemplateUri } from "../bridge/src/skills.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { PROJECT_ROOT } from "../bridge/src/config.js";

let root: string;
const write = (rel: string, text: string) => {
  const file = path.join(root, rel);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, text);
};

beforeEach(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), "skills-"));
  write("alpha/SKILL.md", '---\nname: alpha\ndescription: "Alpha rules for \\"design\\" docs"\n---\n# Alpha\nDo the thing.\n');
  write("alpha/references/style.md", "# Style\nblack only");
  write("alpha/references/data.json", "{}");
  write("beta/SKILL.md", "---\nname: beta\ndescription: >\n  Folded\n  description\n---\nBeta body");
  write("nofront/SKILL.md", "Body without frontmatter");
  write("broken/readme.txt", "no SKILL.md here");
});
afterEach(() => fs.rmSync(root, { recursive: true, force: true }));

describe("parseFrontmatter", () => {
  it("reads quoted values with escapes, block scalars and strips the header", () => {
    const { meta, body } = parseFrontmatter('---\nname: x\ndescription: "a \\"b\\""\nnotes: |\n  l1\n  l2\n---\nBODY');
    expect(meta).toMatchObject({ name: "x", description: 'a "b"', notes: "l1\nl2" });
    expect(body).toBe("BODY");
  });

  it("returns the whole text when there is no frontmatter", () => {
    expect(parseFrontmatter("just text")).toEqual({ meta: {}, body: "just text" });
  });
});

describe("SkillStore", () => {
  it("lists skills with descriptions and falls back to the folder name", () => {
    const store = new SkillStore([root]);
    expect(store.list().map((s) => [s.name, s.description])).toEqual([
      ["alpha", 'Alpha rules for "design" docs'],
      ["beta", "Folded description"],
      ["nofront", ""],
    ]);
  });

  it("lets earlier folders win on a name clash and ignores missing folders", () => {
    const other = fs.mkdtempSync(path.join(os.tmpdir(), "skills2-"));
    fs.mkdirSync(path.join(other, "alpha"));
    fs.writeFileSync(path.join(other, "alpha", "SKILL.md"), "---\nname: alpha\ndescription: other\n---\nx");
    const store = new SkillStore([root, other, "/nonexistent-dir"]);
    expect(store.get("alpha")!.description).toMatch(/Alpha rules/);
    fs.rmSync(other, { recursive: true, force: true });
  });

  it("reads the body, then reference files, and lists them", () => {
    const store = new SkillStore([root]);
    const main = store.read("alpha");
    expect(main.content).toBe("# Alpha\nDo the thing.");
    expect(main.files).toEqual(["references/data.json", "references/style.md"]);
    expect(store.read("alpha", "references/style.md").content).toContain("black only");
  });

  it("rejects traversal, unknown files, unreadable extensions and unknown skills", () => {
    write("secret.md", "top secret");
    const store = new SkillStore([root]);
    expect(() => store.read("alpha", "../secret.md")).toThrow(SkillError);
    expect(() => store.read("alpha", "/etc/passwd")).toThrow(SkillError);
    expect(() => store.read("alpha", "references/missing.md")).toThrow(/not found/);
    write("alpha/run.exe", "binary");
    expect(() => store.read("alpha", "run.exe")).toThrow(/not a readable file/);
    expect(() => store.read("ghost")).toThrow(/unknown skill 'ghost'. Available: alpha/);
  });

  it("refuses symlinks that leave the skill folder", () => {
    write("outside/secret.md", "nope");
    fs.symlinkSync(path.join(root, "outside"), path.join(root, "alpha", "references", "linked"), process.platform === "win32" ? "junction" : "dir");
    expect(() => new SkillStore([root]).read("alpha", "references/linked/secret.md")).toThrow(SkillError);
  });

  it("lists and reads Python and Shell scripts without executing them", () => {
    write("alpha/tools/layout.py", 'raise RuntimeError("read only")');
    write("alpha/install.sh", "exit 7");
    const store = new SkillStore([root]);
    expect(store.read("alpha").files).toContain("tools/layout.py");
    expect(store.read("alpha").files).toContain("install.sh");
    expect(store.read("alpha", "tools/layout.py").content).toContain("read only");
    expect(store.read("alpha", "install.sh").content).toBe("exit 7");
    expect(expandSkillPrompt("/alpha draw", store).prompt).toContain("tools/layout.py");
  });

  it("adds and deletes templates by skill without exposing arbitrary paths", () => {
    const templates = path.join(root, "user-templates");
    const store = new SkillStore([root], templates);
    const item = store.addTemplate("alpha", "example.xlsx", Buffer.from("PKexample"));
    expect(item).toMatchObject({ name: "example.xlsx", kind: "workbook", path: skillTemplateUri("alpha", "example.xlsx") });
    expect(store.read("alpha").content).toContain(item.path);
    expect(store.readTemplate(item.path).bytes.toString()).toBe("PKexample");
    expect(() => store.addTemplate("alpha", "example.xlsx", Buffer.from("PKagain"))).toThrow();
    expect(() => store.addTemplate("alpha", "../escape.xlsx", Buffer.from("PKexample"))).toThrow(SkillError);
    expect(() => store.readTemplate("skill-template://alpha/..%2Fescape.xlsx")).toThrow(SkillError);
    store.deleteTemplate("alpha", "example.xlsx");
    expect(store.listTemplates("alpha")).toEqual([]);
  });

  it("makes text templates readable through read_skill", () => {
    const store = new SkillStore([root], path.join(root, "user-templates"));
    store.addTemplate("alpha", "sample.md", Buffer.from("A model layout"));
    expect(store.read("alpha", "templates/sample.md").content).toBe("A model layout");
    expect(expandSkillPrompt("/alpha make one", store).prompt).toContain("sample.md");
  });
});

describe("expandSkillPrompt", () => {
  it("wraps the skill and keeps the request", () => {
    const { prompt, skill } = expandSkillPrompt("/alpha  vẽ flow cho màn A", new SkillStore([root]));
    expect(skill).toBe("alpha");
    expect(prompt).toMatch(/^<skill name="alpha">\n# Alpha\nDo the thing\./);
    expect(prompt).toContain("references/style.md");
    expect(prompt.endsWith("</skill>\n\nvẽ flow cho màn A")).toBe(true);
  });

  it("asks what to do when no request follows", () => {
    expect(expandSkillPrompt("/alpha", new SkillStore([root])).prompt).toContain("invoked this skill without a request");
  });

  it("leaves unknown slashes and normal prompts alone", () => {
    const store = new SkillStore([root]);
    expect(expandSkillPrompt("/nope do it", store)).toEqual({ prompt: "/nope do it", skill: null });
    expect(expandSkillPrompt("hello /alpha", store).skill).toBeNull();
  });
});

describe("optional user-installed skills", () => {
  it("ships no default project skills or template inventory", () => {
    const bundled = new SkillStore([path.join(PROJECT_ROOT, "skills")]);
    expect(bundled.list()).toEqual([]);
    expect(projectTemplateReminder(bundled)).toBe("");
  });

  it("exposes templates only for user-installed skills", () => {
    const custom = new SkillStore([root], path.join(root, "user-templates"));
    custom.addTemplate("alpha", "layout.xlsx", Buffer.from("PKexample"));
    expect(custom.readTemplate("skill-template://alpha/layout.xlsx").bytes.toString()).toBe("PKexample");
    expect(custom.read("alpha").templates).toMatchObject([{ name: "layout.xlsx", kind: "workbook" }]);
    expect(projectTemplateReminder(custom)).toContain('"path":"skill-template://alpha/layout.xlsx"');
    expect(projectTemplateReminder(custom)).toContain("read_workbook_file");
  });
});

describe("gateway skill tools", () => {
  it("answers list_skills and read_skill without touching the task pane", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("0123456789abcdef", null);
    const skillStore = new SkillStore([root], path.join(root, "user-templates"));
    skillStore.addTemplate("alpha", "sample.md", Buffer.from("Example style"));
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000, undefined, false, skillStore);

    const list = await gateway.invoke({ sessionId: session.id, tool: "list_skills", args: {} });
    expect(list).toMatchObject({ ok: true, data: { skills: [{ name: "alpha" }, { name: "beta" }, { name: "nofront" }] } });

    const read = await gateway.invoke({ sessionId: session.id, tool: "read_skill", args: { name: "alpha", file: "references/style.md" } });
    expect(read).toMatchObject({ ok: true, data: { name: "alpha", file: "references/style.md", templates: [{ name: "sample.md", path: "skill-template://alpha/sample.md" }] } });

    const bad = await gateway.invoke({ sessionId: session.id, tool: "read_skill", args: { name: "alpha", file: "../x.md" } });
    expect(bad).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });
});
