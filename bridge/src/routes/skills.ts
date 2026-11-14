import type http from "node:http";
import { AttachmentError, MAX_ATTACHMENT_BYTES } from "../attachments.js";
import { connectedSession, decodedHeader, readBinaryBody, readBody, sendJson } from "../http-util.js";
import type { WorkbookSessions } from "../workbook-sessions.js";
import { MAX_SKILL_UPLOAD_BYTES, SkillEditConflict, type UserSkills } from "../skill-upload.js";
import { SkillError, type SkillStore } from "../skills.js";
import { PromptSettingsConflict, PromptSettingsError, type SystemPromptStore } from "../system-prompt.js";

export interface SkillRouteDeps {
  sessions: WorkbookSessions;
  skills: SkillStore;
  userSkills: UserSkills;
  systemPrompt: SystemPromptStore;
}

/** What Settings edits: installed skills and their files, skill templates, and the user's own instructions. */
export class SkillRoutes {
  constructor(private readonly deps: SkillRouteDeps) {}

  async userSkills(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    const { skills, userSkills } = this.deps;
    try {
      if (req.method === "PATCH") {
        const body = JSON.parse(await readBody(req)) as { name?: unknown; enabled?: unknown };
        if (typeof body.name !== "string" || typeof body.enabled !== "boolean") throw new SkillError("Specify a skill name and enabled state.");
        skills.setEnabled(body.name, body.enabled);
        return sendJson(res, 200, { name: body.name, enabled: body.enabled });
      }
      if ((req.method === "GET" && url.searchParams.has("name")) || req.method === "PUT") {
        const body = req.method === "PUT" ? JSON.parse(await readBody(req, MAX_SKILL_UPLOAD_BYTES * 6 + 4096)) as Record<string, unknown> : null;
        const name = body?.name ?? url.searchParams.get("name");
        if (typeof name !== "string") throw new SkillError("Specify a skill name.");
        const skill = skills.get(name);
        if (!skill || !userSkills.owns(skill.name, skill.dir)) return sendJson(res, 403, { error: "Only your installed skills can be edited here." });
        const file = body?.file ?? url.searchParams.get("file");
        if (req.method === "GET") return sendJson(res, 200, file === null
          ? { files: userSkills.textFiles(name) }
          : { document: userSkills.readTextFile(name, String(file)) });
        if (typeof file !== "string" || typeof body?.content !== "string" || typeof body.revision !== "string") throw new SkillError("Specify the file, text content and revision.");
        return sendJson(res, 200, { document: userSkills.updateTextFile(name, file, body.content, body.revision) });
      }
      if (req.method === "GET") return sendJson(res, 200, {
        skills: skills.list(true).map(({ name, description, dir }) => ({ name, description, enabled: skills.isEnabled(name), removable: userSkills.owns(name, dir) })),
      });
      if (req.method === "POST") {
        const filename = decodedHeader(req, "x-file-name");
        if (filename === null) return sendJson(res, 400, { error: "Invalid file name." });
        const bytes = await readBinaryBody(req, MAX_SKILL_UPLOAD_BYTES);
        const skill = userSkills.install(filename, bytes, name => skills.get(name) !== null);
        return sendJson(res, 201, { skill });
      }
      if (req.method === "DELETE") {
        const body = JSON.parse(await readBody(req)) as { name?: unknown };
        if (typeof body.name !== "string") return sendJson(res, 400, { error: "Specify a skill name." });
        const skill = skills.get(body.name);
        if (!skill || !userSkills.owns(skill.name, skill.dir)) return sendJson(res, 403, { error: "Only uploaded skills can be removed here." });
        userSkills.remove(skill.name);
        return sendJson(res, 200, { removed: skill.name });
      }
      return sendJson(res, 405, { error: "method_not_allowed" });
    } catch (error) {
      if (error instanceof SkillEditConflict) return sendJson(res, 409, { error: error.message });
      if (error instanceof SkillError || error instanceof AttachmentError || error instanceof SyntaxError) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  listTemplates(req: http.IncomingMessage, res: http.ServerResponse, url: URL): void {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    const skill = url.searchParams.get("skill");
    try {
      if (skill) return sendJson(res, 200, { templates: this.deps.skills.listTemplates(skill) });
      return sendJson(res, 200, { skills: this.deps.skills.list().map(({ name, description }) => ({ name, description })) });
    } catch (error) {
      if (error instanceof SkillError) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }

  async addTemplate(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    const skill = decodedHeader(req, "x-skill-name");
    const name = decodedHeader(req, "x-file-name");
    if (skill === null || name === null) return sendJson(res, 400, { error: "invalid_name" });
    try {
      const bytes = await readBinaryBody(req, MAX_ATTACHMENT_BYTES);
      return sendJson(res, 200, { template: this.deps.skills.addTemplate(skill, name, bytes) });
    } catch (error) {
      if (error instanceof SkillError || error instanceof AttachmentError) return sendJson(res, 400, { error: error.message });
      if ((error as NodeJS.ErrnoException).code === "EEXIST") return sendJson(res, 409, { error: "A template with this name already exists. Delete it first to replace it." });
      throw error;
    }
  }

  async deleteTemplate(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    let body: { skill?: unknown; name?: unknown };
    try { body = JSON.parse(await readBody(req)); }
    catch { return sendJson(res, 400, { error: "invalid_json" }); }
    if (typeof body.skill !== "string" || typeof body.name !== "string") return sendJson(res, 400, { error: "invalid_name" });
    try { this.deps.skills.deleteTemplate(body.skill, body.name); return sendJson(res, 200, { deleted: true }); }
    catch (error) { if (error instanceof SkillError) return sendJson(res, 400, { error: error.message }); throw error; }
  }

  async systemPrompt(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    // Each Office application keeps its own instructions, so the pane in Word never reads or changes the ones Excel uses.
    const store = this.deps.systemPrompt.forHost(session.host);
    try {
      if (req.method === "GET") return sendJson(res, 200, store.read());
      if (req.method !== "PUT") return sendJson(res, 405, { error: "method_not_allowed" });
      const body = JSON.parse(await readBody(req)) as { content?: unknown; revision?: unknown };
      if ((typeof body.content !== "string" && body.content !== null) || typeof body.revision !== "string") throw new PromptSettingsError("Specify prompt content and revision.");
      return sendJson(res, 200, store.update(body.content, body.revision));
    } catch (error) {
      if (error instanceof PromptSettingsConflict) return sendJson(res, 409, { error: error.message });
      if (error instanceof PromptSettingsError || error instanceof SyntaxError) return sendJson(res, 400, { error: error.message });
      throw error;
    }
  }
}
