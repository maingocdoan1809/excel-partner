import type { ToolResult } from "../../../protocol/src/index.js";

export interface ChangeEntry {
  id: number;
  at: Date;
  tool: string;
  ok: boolean;
  summary: string;
  /** Range the user can jump to, when the change has one. */
  target: { sheet: string; address: string } | null;
}

const dim = (rows: unknown, cols: unknown) => `${rows}×${cols}`;

/**
 * Turn an executed write tool into a one-line history entry. Returns null for
 * read-only tools, which are not "changes".
 */
export function describeChange(
  tool: string,
  args: unknown,
  result: ToolResult,
  writeTools: ReadonlySet<string>,
): Omit<ChangeEntry, "id" | "at"> | null {
  if (!writeTools.has(tool)) return null;
  const a0 = (args ?? {}) as Record<string, any>;
  // Some tools double as readers: listing/inspecting is not a change.
  if (tool === "find_replace" && !a0.apply) return null;
  if ((tool === "conditional_format" && a0.operation === "list") || (tool === "data_validation" && a0.operation === "get") || (tool === "manage_names" && a0.operation === "list")) {
    return null;
  }

  if (!result.ok) {
    return { tool, ok: false, summary: `${tool} failed: ${result.error.message}`, target: null };
  }

  const d = (result.data ?? {}) as Record<string, any>;
  const a = (args ?? {}) as Record<string, any>;
  const where = d.sheet && d.address ? `${d.sheet}!${d.address}` : "";
  const target = d.sheet && d.address ? { sheet: String(d.sheet), address: String(d.address) } : null;

  switch (tool) {
    case "write_range":
      return { tool, ok: true, target, summary: `Wrote ${dim(d.rowsWritten, d.columnsWritten)} at ${where}` };
    case "set_formula":
      return { tool, ok: true, target, summary: `Formula ${d.formula} in ${where}` };
    case "find_replace":
      return { tool, ok: true, target: null, summary: `Replaced "${a.find}" with "${a.replace}" in ${d.replaced} cell(s)` };
    case "format_range":
      return { tool, ok: true, target, summary: `Formatted ${where} (${(d.applied ?? []).join(", ")})` };
    case "clear_range":
      return { tool, ok: true, target, summary: `Cleared ${d.clearType} of ${where}` };
    case "insert_rows":
      return { tool, ok: true, target: null, summary: `Inserted ${d.count} row(s) at ${d.shiftedRange} on ${d.sheet}` };
    case "modify_structure":
      return { tool, ok: true, target: null, summary: `${cap(d.operation)} ${d.target} on ${d.sheet}` };
    case "add_shape":
      return { tool, ok: true, target: null, summary: `Added ${d.type} "${d.name}"` };
    case "add_connector":
      return { tool, ok: true, target: null, summary: `Connected "${d.from}" → "${d.to}"` };
    case "add_line":
      return { tool, ok: true, target: null, summary: `Drew line ${d.from} → ${d.to}` };
    case "update_shape":
      return { tool, ok: true, target: null, summary: `Updated shape "${d.name ?? a.name}"` };
    case "delete_shapes":
      return { tool, ok: true, target: null, summary: `Deleted ${(d.deleted ?? []).length} shape(s) on ${d.sheet}` };
    case "set_cells":
      return { tool, ok: true, target: null, summary: `Set ${d.cellsWritten} cell(s) on ${d.sheet}${d.copiedTo ? `, copied to ${d.copiedTo}` : ""}` };
    case "copy_to":
      return { tool, ok: true, target, summary: `Copied ${d.source} → ${where}` };
    case "resize_range":
      return { tool, ok: true, target: null, summary: `Resized ${d.address} on ${d.sheet}` };
    case "merge_range":
      return { tool, ok: true, target, summary: `${cap(d.operation)}d ${where}` };
    case "sort_range":
      return { tool, ok: true, target, summary: `Sorted ${where} by ${(d.sortedBy ?? []).join(", ")}` };
    case "filter_range":
      return { tool, ok: true, target: null, summary: `Filter ${d.operation} on ${d.sheet}` };
    case "conditional_format":
      return { tool, ok: true, target, summary: d.added ? `Added ${d.added} rule to ${where}` : `Cleared conditional formats of ${where}` };
    case "data_validation":
      return { tool, ok: true, target, summary: d.cleared ? `Cleared validation of ${where}` : `Set ${d.type} validation on ${where}` };
    case "manage_sheet":
      return { tool, ok: true, target: null, summary: `${cap(d.operation)} sheet "${d.duplicate ?? d.sheet}"` };
    case "insert_sheets_from_file":
      return { tool, ok: true, target: null, summary: `Inserted ${d.count} sheet(s) from file: ${(d.inserted ?? []).join(", ")}` };
    case "manage_names":
      return { tool, ok: true, target: null, summary: `${cap(d.operation)} name "${d.name}"` };
    case "page_layout":
      return { tool, ok: true, target: null, summary: `Page layout on ${d.sheet}: ${(d.applied ?? []).join(", ")}` };
    case "modify_chart":
      return { tool, ok: true, target: null, summary: `${cap(d.operation)}d chart "${d.name}"` };
    case "modify_pivot_table":
      return { tool, ok: true, target: null, summary: `${cap(d.operation)}d pivot table "${d.name}"` };
    case "modify_table":
      return { tool, ok: true, target: null, summary: `${cap(d.operation)}d table "${d.name}"` };
    case "execute_office_js":
      return { tool, ok: true, target: null, summary: `Ran Office.js${a.description ? `: ${a.description}` : ""}` };
    default:
      return { tool, ok: true, target, summary: tool };
  }
}

const cap = (text: unknown) => String(text ?? "").replace(/^./, (c) => c.toUpperCase());
