import { describe, expect, it } from "vitest";
import { toolCatalogChars } from "../shared/src/catalog.js";
import type { HostId } from "../protocol/src/index.js";

/**
 * Every request pays for the tool definitions before the first word, so the catalog has a budget per host. Measured in
 * characters (about 4 per token). If a change pushes a host over, shrink the definition, or make the tool lazy; do not
 * just raise the number. Measured 2026-11: excel 63.5k / 84.8k, word 35.1k / 40.1k, powerpoint 33.8k / 40.6k.
 */
const BUDGET: Record<HostId, { standard: number; everything: number }> = {
  excel: { standard: 65_000, everything: 87_000 },
  word: { standard: 36_000, everything: 41_500 },
  powerpoint: { standard: 35_000, everything: 42_000 },
};
const EVERYTHING = { webSearch: true, playwright: true, connectedFiles: true };

describe("tool catalog budget", () => {
  for (const host of Object.keys(BUDGET) as HostId[]) {
    it(`${host}: stays within its budget, with and without every optional tool group`, () => {
      expect(toolCatalogChars(host, false, {})).toBeLessThanOrEqual(BUDGET[host].standard);
      expect(toolCatalogChars(host, true, EVERYTHING)).toBeLessThanOrEqual(BUDGET[host].everything);
    });
  }

  it("counts what the model is offered: more tools, more characters", () => {
    expect(toolCatalogChars("excel", true, EVERYTHING)).toBeGreaterThan(toolCatalogChars("excel", false, {}));
  });
});
