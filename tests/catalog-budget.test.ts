import { describe, expect, it } from "vitest";
import { toolCatalogChars } from "../shared/src/catalog.js";
import type { HostId } from "../protocol/src/index.js";

/**
 * Every request pays for the tool definitions before the first word, so the catalog has a budget per host. Measured in
 * characters (about 4 per token). If a change pushes a host over, shrink the definition, or make the tool lazy; do not
 * just raise the number. Measured 2026-11: excel 66.1k / 87.4k (analyze_range added 2.6k), word 35.1k / 40.1k, powerpoint 33.8k / 40.7k.
 */
const BUDGET: Record<HostId, { standard: number; everything: number }> = {
  excel: { standard: 68_000, everything: 90_000 },
  word: { standard: 36_000, everything: 41_500 },
  powerpoint: { standard: 35_000, everything: 42_000 },
};
const EVERYTHING = { webSearch: true, playwright: true, connectedFiles: true };
/** The same, when the bridge offers most tools by summary (`toolSchemas: "lazy"`). Measured: excel 19.3k / 22.5k, word 14.9k / 16.1k, powerpoint 11.9k / 13.4k. */
const LAZY_BUDGET: Record<HostId, { standard: number; everything: number }> = {
  excel: { standard: 20_000, everything: 23_500 },
  word: { standard: 15_500, everything: 17_000 },
  powerpoint: { standard: 12_500, everything: 14_000 },
};

describe("tool catalog budget", () => {
  for (const host of Object.keys(BUDGET) as HostId[]) {
    it(`${host}: stays within its budget, with and without every optional tool group`, () => {
      expect(toolCatalogChars(host, false, {})).toBeLessThanOrEqual(BUDGET[host].standard);
      expect(toolCatalogChars(host, true, EVERYTHING)).toBeLessThanOrEqual(BUDGET[host].everything);
    });
  }

  for (const host of Object.keys(LAZY_BUDGET) as HostId[]) {
    it(`${host}: stays within its budget when tools are offered by summary`, () => {
      expect(toolCatalogChars(host, false, { lazySchemas: true })).toBeLessThanOrEqual(LAZY_BUDGET[host].standard);
      expect(toolCatalogChars(host, true, { ...EVERYTHING, lazySchemas: true })).toBeLessThanOrEqual(LAZY_BUDGET[host].everything);
    });
  }

  it("counts what the model is offered: more tools, more characters", () => {
    expect(toolCatalogChars("excel", true, EVERYTHING)).toBeGreaterThan(toolCatalogChars("excel", false, {}));
  });
});
