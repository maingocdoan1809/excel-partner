import { ZodError } from "zod";
import {
  exposedToolNames,
  isExcelToolName,
  parseToolInput,
  rangeCellCount,
  SelectionSnapshotSchema,
  type ExcelToolName,
  type ToolError,
  type ToolResult,
} from "../../protocol/src/index.js";
import { ReferenceError_, type ReferenceFiles } from "./reference-files.js";
import { SkillError, parseSkillTemplateUri, type SkillStore } from "./skills.js";
import type { UserSkills } from "./skill-upload.js";
import type { SaveSkillInput, UpdateSkillInput } from "../../protocol/src/index.js";
import type { RunSkillScriptInput } from "../../protocol/src/index.js";
import { resolveSkillScript, runSkillScript, skillFile } from "./skill-runner.js";
import { runPython, PythonError, PYTHON_FILE_BYTES } from "./python-runner.js";
import type { RunPythonInput } from "../../protocol/src/index.js";
import { XlsxError, listSheetNames, readSheetDetails, readWorkbookOverview } from "./xlsx-reader.js";
import { AttachmentError, attachmentId, attachmentPath, createAttachment, MAX_ATTACHMENT_BYTES } from "./attachments.js";
import { inspectAttachment } from "./inspect-attachment.js";
import fs from "node:fs";
import pathModule from "node:path";
import type { SessionManager } from "./session-manager.js";
import type { Logger } from "./logger.js";
import { enrichWindowsSelection, type SelectionEnricher } from "./windows-selection.js";

export interface GatewayLimits {
  maxReadCells: number;
  maxWriteCells: number;
  maxSheets: number;
  maxSelectionCells: number;
}

interface InsertArgs {
  path: string;
  sheets?: string[];
  position: string;
  afterSheet?: string;
  fileBase64?: string;
}

export interface InternalToolRequest {
  sessionId: string;
  tool: string;
  args: unknown;
}

/**
 * Validates tool calls arriving from the MCP process and routes the accepted
 * ones to the workbook session that owns the task pane.
 */
export class ToolGateway {
  constructor(
    private readonly sessions: SessionManager,
    private readonly limits: GatewayLimits,
    private readonly timeoutMs: number,
    private readonly log?: Logger,
    private readonly allowOptInTools = false,
    private readonly skills: SkillStore | null = null,
    private readonly references: ReferenceFiles | null = null,
    private readonly userSkills: UserSkills | null = null,
    private readonly enrichSelection: SelectionEnricher = enrichWindowsSelection,
  ) {}

  async invoke(request: InternalToolRequest): Promise<ToolResult> {
    const session = this.sessions.get(request.sessionId);
    const requestId = session?.activeRequestId;
    const result = await this.invokeTool(request);
    if (!session || !requestId || session.activeRequestId !== requestId || session.turnController?.signal.aborted) return result;
    const followups = session.takeFollowups(requestId, "delivered");
    return followups.length ? { ...result, additional_user_messages: followups } : result;
  }

  private async invokeTool(request: InternalToolRequest): Promise<ToolResult> {
    const session = this.sessions.get(request.sessionId);
    if (!session) {
      return error("disconnected", `unknown or expired session '${request.sessionId}'`);
    }
    if (!isExcelToolName(request.tool)) {
      return error("unknown_tool", `unknown tool '${request.tool}'`);
    }

    if (!exposedToolNames(this.allowOptInTools).includes(request.tool)) {
      return error(
        "unsupported",
        `tool '${request.tool}' is disabled; start the bridge with BRIDGE_ALLOW_OFFICE_JS=1 to enable it`,
      );
    }

    let parsed: unknown;
    try {
      parsed = parseToolInput(request.tool, request.args);
    } catch (err) {
      if (err instanceof ZodError) {
        return error("invalid_arguments", `invalid arguments for ${request.tool}`, err.issues);
      }
      return error("invalid_arguments", String(err));
    }

    if (request.tool === "list_skills" || request.tool === "read_skill" || request.tool === "save_skill") {
      return this.answerSkillTool(request.tool, parsed);
    }
    if (request.tool === "update_skill") {
      const input = parsed as UpdateSkillInput;
      const skill = this.skills?.get(input.name);
      if (!skill || !this.userSkills?.owns(skill.name, skill.dir)) return error("invalid_arguments", "Only user-installed skills can be updated.");
      try {
        const current = this.userSkills.readTextFile(input.name, input.file);
        if (current.revision !== input.revision) return error("invalid_arguments", "This skill file changed. Read it again with forEditing=true before updating.");
        const requestId = session.activeRequestId;
        const signal = session.turnController?.signal;
        const allowed = await session.requestTool("__authorize_skill_edit", { name: input.name, file: input.file, content: input.content }, this.timeoutMs);
        if (!allowed.ok) return allowed;
        if (!(allowed.data as { approved?: boolean })?.approved) return error("cancelled", "The user declined this skill update.");
        if (signal?.aborted || session.activeRequestId !== requestId) return error("cancelled", "The request was cancelled before saving the skill.");
        const updated = this.userSkills.updateTextFile(input.name, input.file, input.content, input.revision);
        return { ok: true, data: { name: input.name, file: updated.file, revision: updated.revision, saved: true } };
      } catch (err) { if (err instanceof SkillError) return error("invalid_arguments", err.message); throw err; }
    }
    if (request.tool === "run_skill_script") {
      return this.answerScriptTool(session, parsed as RunSkillScriptInput);
    }
    if (request.tool === "run_python") return this.answerPythonTool(session, parsed as RunPythonInput);
    if (request.tool === "list_reference_files" || request.tool === "read_workbook_file" || request.tool === "read_attached_image" || request.tool === "read_attached_file") {
      return this.answerFileTool(session, request.tool, parsed);
    }
    if (request.tool === "insert_sheets_from_file") {
      const enriched = await this.attachWorkbookFile(session, parsed as InsertArgs);
      if ("ok" in enriched) return enriched;
      parsed = enriched;
    }

    const limitError = this.enforceLimits(request.tool, parsed);
    if (limitError) return limitError;

    session.touch();
    this.log?.debug("tool -> pane", { sessionId: session.id, tool: request.tool });
    const started = Date.now();
    let result = await session.requestTool(request.tool, parsed, this.timeoutMs);
    if (request.tool === "get_selection" && result.ok) {
      const snapshot = SelectionSnapshotSchema.safeParse(result.data);
      if (snapshot.success) result = { ok: true, data: await this.enrichSelection(snapshot.data, session.turnController?.signal) };
    }
    this.logOutcome(request.tool, parsed, result, Date.now() - started);
    return result;
  }

  /** One line per tool call; failures carry the error and the arguments so "check the log" is enough to diagnose them. */
  private logOutcome(tool: string, args: unknown, result: ToolResult, ms: number): void {
    if (result.ok) {
      this.log?.info("tool ok", { tool, ms });
      return;
    }
    this.log?.warn("tool failed", {
      tool,
      ms,
      code: result.error.code,
      message: result.error.message.slice(0, 300),
      detail: result.error.detail === undefined ? undefined : JSON.stringify(result.error.detail).slice(0, 300),
      args: summariseArgs(args),
    });
  }

  private async answerPythonTool(session: import("./session.js").WorkbookSession, args: RunPythonInput): Promise<ToolResult> {
    const controller = new AbortController();
    session.scriptControllers.add(controller);
    try {
      const seen = new Set<string>();
      let bytes = 0;
      const files = args.files.map(file => {
        if (seen.has(file.path)) throw new PythonError("Duplicate Python input file path.");
        seen.add(file.path);
        const attachment = session.attachments.get(attachmentId(file.attachment)!);
        if (!attachment) throw new PythonError("Python input attachment is unavailable; select it again.");
        bytes += attachment.bytes.length;
        return { path: file.path, bytes: attachment.bytes };
      });
      if (bytes > PYTHON_FILE_BYTES) throw new PythonError("Python inputs exceed 10 MB combined.");
      const signal = session.turnController ? AbortSignal.any([controller.signal, session.turnController.signal]) : controller.signal;
      if (signal.aborted) return error("cancelled", "Python execution cancelled.");
      const approval = await session.requestTool("__authorize_python", args, 45_000);
      if (!approval.ok) return approval;
      if ((approval.data as { approved?: unknown })?.approved !== true) return error("cancelled", "The user declined Python execution.");
      if (!session.isConnected) return error("disconnected", "task pane disconnected while requesting Python execution");
      const started = Date.now();
      session.touch();
      const data = await runPython(args, files, signal);
      const { outputFiles: generated, ...status } = data;
      if (data.exitCode !== 0 || data.timedOut || data.cancelled || data.outputTruncated) {
        const result = error(data.cancelled ? "cancelled" : data.timedOut ? "timeout" : "internal_error", data.cancelled ? "Python execution cancelled." : data.timedOut ? "Python timed out." : data.outputTruncated ? "Python exceeded its output limit." : `Python exited with code ${data.exitCode}.`, status);
        this.logOutcome("run_python", {}, result, Date.now() - started);
        return result;
      }
      const attachments = generated.map(file => createAttachment(pathModule.posix.basename(file.name), Buffer.from(file.base64, "base64")));
      const outputFiles = attachments.map(attachment => {
        if (session.attachments.size >= 10) session.attachments.delete(session.attachments.keys().next().value!);
        session.attachments.set(attachment.id, attachment);
        return { name: attachment.name, path: attachmentPath(attachment.id), size: attachment.bytes.length, kind: attachment.kind };
      });
      const result: ToolResult = { ok: true, data: { ...status, execution: "local", outputFiles } };
      this.logOutcome("run_python", {}, result, Date.now() - started);
      return result;
    } catch (err) {
      if (err instanceof PythonError || err instanceof AttachmentError) return error(controller.signal.aborted || session.turnController?.signal.aborted ? "cancelled" : "invalid_arguments", err.message);
      throw err;
    } finally { session.scriptControllers.delete(controller); }
  }

  private async answerScriptTool(session: import("./session.js").WorkbookSession, args: RunSkillScriptInput): Promise<ToolResult> {
    if (!this.skills) return error("unsupported", "skills are not configured on this bridge");
    const controller = new AbortController();
    session.scriptControllers.add(controller);
    try {
      resolveSkillScript(this.skills, args);
      const signal = session.turnController
        ? AbortSignal.any([controller.signal, session.turnController.signal]) : controller.signal;
      if (signal?.aborted) return error("cancelled", "Script execution cancelled.");
      const approval = await session.requestTool("__authorize_skill_script", args, 45_000);
      if (!approval.ok) return approval;
      if ((approval.data as { approved?: unknown })?.approved !== true) return error("cancelled", "The user declined script execution.");
      if (!session.isConnected) return error("disconnected", "task pane disconnected while requesting script execution");
      if (signal?.aborted) return error("cancelled", "Script execution cancelled.");
      session.touch();
      const started = Date.now();
      // Resolve again after approval so a removed/replaced skill cannot bypass path validation.
      const data = await runSkillScript(this.skills, args, signal);
      const failed = data.exitCode !== 0 || data.timedOut || data.cancelled || data.outputTruncated;
      const outputFiles: Array<{ name: string; path: string; size: number; kind: string }> = [];
      if (!failed) {
        let total = 0;
        const attachments = args.outputFiles.map(relative => {
          const selected = skillFile(this.skills!, args.name, relative);
          total += fs.statSync(selected.file).size;
          if (total > MAX_ATTACHMENT_BYTES) throw new SkillError("Generated files exceed the combined 30 MB limit.");
          return createAttachment(pathModule.basename(selected.file), fs.readFileSync(selected.file));
        });
        for (const attachment of attachments) {
          if (session.attachments.size >= 10) session.attachments.delete(session.attachments.keys().next().value!);
          session.attachments.set(attachment.id, attachment);
          outputFiles.push({ name: attachment.name, path: attachmentPath(attachment.id), size: attachment.bytes.length, kind: attachment.kind });
        }
      }
      const result: ToolResult = failed
        ? error(data.cancelled ? "cancelled" : data.timedOut ? "timeout" : "internal_error", data.cancelled ? "Script cancelled." : data.timedOut ? "Script timed out." : data.outputTruncated ? "Script exceeded the output limit." : `Script exited with code ${data.exitCode}.`, data)
        : { ok: true, data: { ...data, outputFiles } };
      this.logOutcome("run_skill_script", { name: args.name, file: args.file }, result, Date.now() - started);
      return result;
    } catch (err) {
      if (err instanceof SkillError || err instanceof AttachmentError) return error("invalid_arguments", err.message);
      throw err;
    } finally {
      session.scriptControllers.delete(controller);
    }
  }

  /** Read-only file tools: answered here because the files live on the bridge's machine. */
  private async answerFileTool(session: import("./session.js").WorkbookSession, tool: "list_reference_files" | "read_workbook_file" | "read_attached_image" | "read_attached_file", args: unknown): Promise<ToolResult> {
    try {
      if (tool === "read_attached_file") {
        const { path, entry, offset, maxBytes } = args as { path: string; entry?: string; offset: number; maxBytes: number };
        const id = attachmentId(path);
        const file = id ? session.attachments.get(id) : null;
        if (file) return { ok: true, data: inspectAttachment(file, { entry, offset, maxBytes }) };
        if (!this.references) return error("unsupported", "reference files are not configured");
        const selected = this.references.selectedPath(path, session.referenceGrants);
        const stat = fs.statSync(selected);
        if (stat.size > 30 * 1024 * 1024) return error("invalid_arguments", "file exceeds the 30 MB read limit");
        return { ok: true, data: inspectAttachment({ name: pathModule.basename(selected), bytes: fs.readFileSync(selected) }, { entry, offset, maxBytes }) };
      }
      if (tool === "read_attached_image") {
        const input = (args as { path: string }).path;
        if (parseSkillTemplateUri(input)) {
          if (!this.skills) return error("unsupported", "skills are not configured");
          const { template, bytes } = this.skills.readTemplate(input);
          if (template.kind !== "image") return error("invalid_arguments", "template is not an image");
          const ext = template.name.split(".").at(-1)?.toLowerCase();
          return { ok: true, data: { name: template.name, image: { mimeType: ext === "png" ? "image/png" : ext === "gif" ? "image/gif" : ext === "webp" ? "image/webp" : "image/jpeg", data: bytes.toString("base64") } } };
        }
        const id = attachmentId(input);
        const file = id ? session.attachments.get(id) : null;
        if (file?.kind === "image") return { ok: true, data: { name: file.name, image: { mimeType: file.mimeType, data: file.bytes.toString("base64") } } };
        if (!this.references) return error("unsupported", "reference files are not configured");
        const selected = this.references.selectedPath(input, session.referenceGrants);
        const image = createAttachment(pathModule.basename(selected), fs.readFileSync(selected));
        if (image.kind !== "image") return error("invalid_arguments", "selected file is not a supported image");
        return { ok: true, data: { name: image.name, image: { mimeType: image.mimeType, data: image.bytes.toString("base64") } } };
      }
      if (tool === "list_reference_files") {
        if (!this.references) return error("unsupported", "reference files are not configured on this bridge");
        const { path } = args as { path?: string };
        return { ok: true, data: this.references.list(path, session.referenceGrants) };
      }
      const a = args as { path: string; sheet?: string; range?: string; includeStyles: boolean; includeShapes: boolean; maxCells: number };
      const source = await this.workbookSource(session, a.path);
      if ("ok" in source) return source;
      const { path, bytes } = source;
      if (!a.sheet) return { ok: true, data: { path, ...readWorkbookOverview(bytes) } };
      return {
        ok: true,
        data: readSheetDetails(bytes, a.sheet, { range: a.range, includeStyles: a.includeStyles, includeShapes: a.includeShapes, maxCells: a.maxCells }),
      };
    } catch (err) {
      if (err instanceof ReferenceError_ || err instanceof XlsxError || err instanceof SkillError || err instanceof AttachmentError) return error("invalid_arguments", err.message);
      throw err;
    }
  }

  /** The pane cannot read the bridge's disk, so the workbook bytes travel with the request. */
  private async attachWorkbookFile(session: import("./session.js").WorkbookSession, args: InsertArgs): Promise<InsertArgs | ToolResult> {
    try {
      const source = await this.workbookSource(session, args.path);
      if ("ok" in source) return source;
      const { path, bytes } = source;
      const available = listSheetNames(bytes);
      const missing = (args.sheets ?? []).filter((name) => !available.includes(name));
      if (missing.length > 0) {
        return error("invalid_arguments", `sheet(s) not found in the file: ${missing.join(", ")}`, { available });
      }
      return { ...args, path, fileBase64: Buffer.from(bytes).toString("base64") };
    } catch (err) {
      if (err instanceof ReferenceError_ || err instanceof XlsxError || err instanceof SkillError) return error("invalid_arguments", err.message);
      throw err;
    }
  }

  private async workbookSource(session: import("./session.js").WorkbookSession, input: string): Promise<{ path: string; bytes: Uint8Array } | ToolResult> {
    if (parseSkillTemplateUri(input)) {
      if (!this.skills) return error("unsupported", "skills are not configured");
      const { template, bytes } = this.skills.readTemplate(input);
      return template.kind === "workbook" ? { path: input, bytes } : error("invalid_arguments", "template is not a workbook");
    }
    const id = attachmentId(input);
    if (id) {
      const file = session.attachments.get(id);
      return file?.kind === "workbook"
        ? { path: input, bytes: file.bytes }
        : error("invalid_arguments", "attached workbook is unavailable; ask the user to select it again");
    }
    if (!this.references) return error("unsupported", "reference files are not configured on this bridge");
    const denied = await this.authorizePath(session, input);
    if (denied) return denied;
    return this.references.readWorkbook(input, session.referenceGrants);
  }

  private async authorizePath(session: import("./session.js").WorkbookSession, input: string): Promise<ToolResult | null> {
    if (!this.references) return error("unsupported", "reference files are not configured");
    try {
      this.references.resolve(input, session.referenceGrants);
      return null;
    } catch (err) {
      if (!(err instanceof ReferenceError_) || err.reason !== "outside") throw err;
    }
    const target = this.references.permissionTarget(input);
    // The pane asks the person using this workbook. Approval applies to this exact file and session only.
    const response = await session.requestTool("__authorize_reference_path", { path: input, resolvedPath: target }, 120_000);
    if (!response.ok) return response;
    if ((response.data as { approved?: unknown })?.approved !== true) {
      return error("cancelled", `The user declined read access to '${target}'`);
    }
    if (!session.isConnected) return error("disconnected", "task pane disconnected while requesting file access");
    session.referenceGrants.add(target);
    return null;
  }

  /** Skills live on the bridge's disk, so these two tools never reach the task pane. */
  private answerSkillTool(tool: "list_skills" | "read_skill" | "save_skill", args: unknown): ToolResult {
    if (!this.skills) return error("unsupported", "skills are not configured on this bridge");
    try {
      if (tool === "save_skill") {
        if (!this.userSkills) return error("unsupported", "skill saving is not configured on this bridge");
        const { name, description, instructions } = args as SaveSkillInput;
        if (!instructions.trim()) return error("invalid_arguments", "Skill instructions cannot be empty.");
        const content = `---\nname: ${name}\ndescription: ${JSON.stringify(description)}\n---\n${instructions}\n`;
        const skill = this.userSkills.install("SKILL.md", Buffer.from(content, "utf8"), candidate => this.skills!.get(candidate) !== null);
        return { ok: true, data: { ...skill, path: pathModule.join(this.userSkills.root, skill.name, "SKILL.md"), command: `/${skill.name}`, saved: true } };
      }
      if (tool === "list_skills") {
        return { ok: true, data: { skills: this.skills.list().map(({ name, description }) => ({ name, description })) } };
      }
      const { name, file, forEditing } = args as { name: string; file?: string; forEditing?: boolean };
      if (forEditing) {
        const skill = this.skills.get(name);
        if (!skill || !this.userSkills?.owns(skill.name, skill.dir)) return error("invalid_arguments", "Only user-installed skills can be edited.");
        return { ok: true, data: { name, ...this.userSkills.readTextFile(name, file ?? "SKILL.md"), files: this.userSkills.textFiles(name) } };
      }
      return { ok: true, data: this.skills.read(name, file) };
    } catch (err) {
      if (err instanceof SkillError) return error("invalid_arguments", err.message);
      throw err;
    }
  }

  private enforceLimits(tool: ExcelToolName, args: unknown): ToolResult | null {
    switch (tool) {
      case "read_range": {
        const address = (args as { address: string }).address;
        const count = rangeCellCount(address);
        if (count === null) return error("invalid_arguments", `cannot parse address '${address}'`);
        if (count > this.limits.maxReadCells) {
          return error(
            "size_limit_exceeded",
            `range ${address} has ${count} cells (limit ${this.limits.maxReadCells})`,
          );
        }
        return null;
      }
      case "write_range": {
        const values = (args as { values: unknown[][] }).values;
        const rows = values.length;
        const cols = values.reduce((max, row) => Math.max(max, row.length), 0);
        const total = rows * cols;
        if (total > this.limits.maxWriteCells) {
          return error(
            "size_limit_exceeded",
            `write of ${total} cells exceeds limit ${this.limits.maxWriteCells}`,
          );
        }
        return null;
      }
      case "copy_to":
      case "format_range":
      case "set_formula":
      case "clear_range": {
        const address = tool === "copy_to" ? (args as { destination: string }).destination : (args as { address: string }).address;
        const count = rangeCellCount(address);
        // Whole row/column references (`D:D`, `3:3`) have no countable size; Excel bounds them.
        if (count !== null && count > this.limits.maxWriteCells) {
          return error(
            "size_limit_exceeded",
            `${tool} range ${address} has ${count} cells (limit ${this.limits.maxWriteCells})`,
          );
        }
        return null;
      }
      default:
        return null;
    }
  }
}

function error(code: ToolError["code"], message: string, detail?: unknown): ToolResult {
  return { ok: false, error: detail === undefined ? { code, message } : { code, message, detail } };
}

/** Arguments for the log: bulky payloads (a workbook's bytes) are replaced and the rest is truncated. */
function summariseArgs(args: unknown): string {
  try {
    return JSON.stringify(args, (key, value) => (key === "fileBase64" ? `<${String(value).length} chars>` : value)).slice(0, 400);
  } catch {
    return "<unserialisable>";
  }
}
