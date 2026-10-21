import { describe, expect, it } from "vitest";
import { parseClaudeModels } from "../providers/src/claude.js";

describe("Claude model list", () => {
  it("maps the CLI model list, keeps exact versions and reads the default", () => {
    const catalog = parseClaudeModels([
      { value: "default", resolvedModel: "claude-sonnet-5-5", displayName: "Default (recommended)" },
      { value: "sonnet", displayName: "Sonnet 5.5" },
      { value: "claude-sonnet-5", displayName: "Sonnet 5" },
      { value: "sonnet", displayName: "duplicate" },
      { value: "claude-haiku-4-5-20251001" },
    ]);
    expect(catalog).toEqual({
      defaultModelId: "claude-sonnet-5-5",
      models: [
        { id: "sonnet", name: "Sonnet 5.5" },
        { id: "claude-sonnet-5", name: "Sonnet 5" },
        { id: "claude-haiku-4-5-20251001", name: "claude-haiku-4-5-20251001" },
      ],
    });
  });

  it("takes the version from the description when displayName has none", () => {
    expect(parseClaudeModels([
      { value: "sonnet", displayName: "Sonnet", description: "Sonnet 5.5 · Efficient for routine tasks · $2/$10 per Mtok" },
      { value: "claude-fable-5-1[1m]", displayName: "Fable", description: "Fable 5.1 · Most capable for your hardest tasks" },
      { value: "opus", displayName: "Opus", description: "Best for everyday, complex tasks" },
    ])?.models.map((m) => m.name)).toEqual(["Sonnet 5.5", "Fable 5.1", "Opus"]);
  });

  it("returns null when the CLI gives no usable list", () => {
    expect(parseClaudeModels(undefined)).toBeNull();
    expect(parseClaudeModels([{ value: "default" }])).toBeNull();
  });
});
