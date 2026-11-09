import type { HostId, ToolResult } from "../../../protocol/src/index.js";

/** What the task pane needs from the Office application it runs in. Everything host-specific sits behind this. */
export interface HostAdapter {
  readonly id: HostId;
  /** Name of the application, for the interface ("Connected to Excel"). */
  readonly app: string;
  /** What the open file is called in the interface: workbook, document, presentation. */
  readonly noun: string;
  /** File extension a backup copy gets. */
  readonly extension: ".xlsx" | ".docx" | ".pptx";
  /** Tools that change the file. The Changes list shows their results and a backup is taken before the first one of a reply. */
  readonly writeTools: ReadonlySet<string>;
  /** Run one tool request in the open file. Arguments were validated by the bridge. */
  execute(tool: string, args: unknown, signal?: AbortSignal): Promise<ToolResult>;
  /** The name the open file has now, or null when Office cannot say. */
  readName(): Promise<string | null>;
  /** Short text for the selection chip. Throws when the selection cannot be identified. */
  readSelectionLabel(): Promise<{ label: string; title: string }>;
  /** The open file as zip bytes (xlsx/docx/pptx), for backups. */
  readFile(): Promise<Uint8Array>;
  /** Excel only: sheets, cell links, shape-selection events and cross-file features. */
  readonly hasSheets: boolean;
}
