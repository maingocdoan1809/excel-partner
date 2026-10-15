import { ZodError } from "zod";
import {
  exposedToolNames,
  isExcelToolName,
  parseToolInput,
  rangeCellCount,
  type ExcelToolName,
  type ToolError,
  type ToolResult,
} from "../../protocol/src/index.js";
import { ReferenceError_, type ReferenceFiles } from "./reference-files.js";
import { SkillError, parseSkillTemplateUri, type SkillStore } from "./skills.js";
import type { UserSkills } from "./skill-upload.js";
import type { SaveSkillInput } from "../../protocol/src/index.js";
import { XlsxError, listSheetNames, readSheetDetails, readWorkbookOverview } from "./xlsx-reader.js";
import { AttachmentError, attachmentId, createAttachment } from "./attachments.js";
import { inspectAttachment } from "./inspect-attachment.js";
import fs from "node:fs";
import pathModule from "node:path";
import type { SessionManager } from "./session-manager.js";
import type { Logger } from "./logger.js";

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
  ) {}

  async invoke(request: InternalToolRequest): Promise<ToolResult> {
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
    const result = await session.requestTool(request.tool, parsed, this.timeoutMs);
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
      const { name, file } = args as { name: string; file?: string };
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
