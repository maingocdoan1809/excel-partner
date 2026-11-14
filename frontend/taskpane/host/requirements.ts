import type { HostId } from "../../../protocol/src/index.js";
import { ToolFailure } from "../executor/common.js";

/** An API set and the first version that has what a feature needs. */
export type Requirement = readonly [set: string, version: string];

/** What every tool of a host needs at minimum. Below this the pane says so instead of calling the host. */
export const HOST_BASELINE: Record<HostId, Requirement> = {
  excel: ["ExcelApi", "1.9"],
  word: ["WordApi", "1.1"],
  powerpoint: ["PowerPointApi", "1.2"],
};

export const WORD_COMMENTS: Requirement = ["WordApi", "1.4"];
export const WORD_FIELDS: Requirement = ["WordApi", "1.5"];
export const WORD_TRACK: Requirement = ["WordApi", "1.6"];
export const WORD_BOOKMARK: Requirement = ["WordApi", "1.4"];
export const WORD_TOC: Requirement = ["WordApiDesktop", "1.4"];
export const WORD_LISTS: Requirement = ["WordApi", "1.3"];
export const PPT_SELECTION: Requirement = ["PowerPointApi", "1.5"];
export const PPT_EXPORT: Requirement = ["PowerPointApi", "1.8"];
export const PPT_MODERN: Requirement = ["PowerPointApi", "1.10"];

export function isSupported(...requirements: Requirement[]): boolean {
  try {
    return requirements.every(([set, version]) => Office.context.requirements.isSetSupported(set, version));
  } catch {
    return false;
  }
}

/**
 * Stop with `unsupported` (never a crash) when the host is too old for a feature. The message says what to do,
 * because the model relays it to the person.
 */
export function requireSupport(feature: string, requirement: Requirement, hint?: string): void {
  if (isSupported(requirement)) return;
  const [set, version] = requirement;
  throw new ToolFailure("unsupported", `${feature} needs ${set} ${version}, which this version of Office does not have.${hint ? ` ${hint}` : ""}`.trim());
}
