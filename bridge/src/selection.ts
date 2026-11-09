import {
  AnySelectionSnapshotSchema,
  type AnySelectionSnapshot,
  type ExcelSelectionSnapshot,
  type HostId,
} from "../../protocol/src/index.js";
import type { SelectionEnricher } from "./windows-selection.js";

export type NormalizedSelection = { ok: true; selection: AnySelectionSnapshot | null } | { ok: false; message: string };

/**
 * The selection a pane sent with a prompt, tagged with its host (a snapshot without `host` is Excel's). It must belong
 * to the host the pane said it runs in. An Excel snapshot that does not match the schema is kept as it was sent, as
 * before hosts existed; a Word or PowerPoint snapshot is validated strictly because nothing else has ever vouched for it.
 */
export function normalizePromptSelection(raw: unknown, host: HostId): NormalizedSelection {
  if (raw === null || raw === undefined) return { ok: true, selection: null };
  const parsed = AnySelectionSnapshotSchema.safeParse(raw);
  if (parsed.success) {
    return parsed.data.host === host
      ? { ok: true, selection: parsed.data }
      : { ok: false, message: `the selection is from ${parsed.data.host}, but this pane runs in ${host}` };
  }
  if (host === "excel" && typeof raw === "object" && !Array.isArray(raw) && !("host" in raw && (raw as { host: unknown }).host !== "excel")) {
    return { ok: true, selection: { ...(raw as object), host: "excel" } as ExcelSelectionSnapshot };
  }
  return { ok: false, message: "the selection is not a valid snapshot" };
}

/** Add what the machine knows about the selection. Only Excel on Windows has anything to add (selected shapes through COM). */
export async function enrichSelectionForHost(host: HostId, selection: AnySelectionSnapshot, enrich: SelectionEnricher, signal?: AbortSignal): Promise<AnySelectionSnapshot> {
  if (host !== "excel" || selection.host !== "excel") return selection;
  return { ...(await enrich(selection, signal)), host: "excel" };
}
