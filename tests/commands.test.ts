import { describe, expect, it, vi } from "vitest";
import { cachedCommands, normalizeCommands } from "../providers/src/commands.js";
import { CODEX_COMMANDS } from "../providers/src/codex-commands.js";
import { parseNativeCommand } from "../protocol/src/commands.js";

describe("provider command catalogs", () => {
  it("preserves native skills, argument hints and client-owned controls", () => {
    expect(normalizeCommands([
      { name: "/skills", description: "Skills", kind: "builtin", input: { hint: "list | info" } },
      { name: "report", description: "Report", builtin: false, argumentHint: "file" },
      { name: "model", description: "Pick model" }, { name: "skills", description: "Duplicate" },
      { name: "invalid name", description: "Invalid" },
    ])).toEqual([
      expect.objectContaining({ name: "skills", execution: "native", inputHint: "list | info", source: "builtin" }),
      expect.objectContaining({ name: "report", execution: "native", inputHint: "file", source: "skill" }),
      expect.objectContaining({ name: "model", execution: "client", action: "model" }),
    ]);
  });
  it("includes the Codex catalog and distinguishes terminal-only controls", () => {
    expect(CODEX_COMMANDS.length).toBeGreaterThan(40);
    expect(CODEX_COMMANDS.find(command => command.name === "skills")).toMatchObject({ execution: "native" });
    expect(CODEX_COMMANDS.find(command => command.name === "theme")).toMatchObject({ execution: "client", action: "theme" });
    expect(CODEX_COMMANDS.find(command => command.name === "vim")).toMatchObject({ execution: "unsupported" });
    expect(new Set(CODEX_COMMANDS.map(command => command.name)).size).toBe(CODEX_COMMANDS.length);
  });
  it("parses command arguments without turning ordinary messages into commands", () => {
    expect(parseNativeCommand("/skills report first line\nsecond line")).toEqual({ name: "skills", args: "report first line\nsecond line" });
    expect(parseNativeCommand("/plugin:report task")).toEqual({ name: "plugin:report", args: "task" });
    expect(parseNativeCommand("Please run /skills")).toBeNull();
  });
  it("coalesces discovery and permits retry after a failed launch", async () => {
    const load = vi.fn().mockRejectedValueOnce(new Error("offline")).mockResolvedValue([]);
    const catalog = cachedCommands(load);
    await expect(catalog()).rejects.toThrow("offline");
    await Promise.all([catalog(), catalog()]);
    expect(load).toHaveBeenCalledTimes(2);
  });
});
