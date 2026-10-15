import { describe, expect, it } from "vitest";
import { effortConfig } from "../providers/src/effort.js";
import { PromptMessageSchema } from "../protocol/src/messages.js";

describe("reasoning effort", () => {
  it("validates explicit levels while accepting older clients", () => {
    const prompt = { type: "prompt", requestId: "test", prompt: "Hello", selection: null };
    expect(PromptMessageSchema.parse(prompt).effort).toBeUndefined();
    for (const effort of ["low", "medium", "high"]) expect(PromptMessageSchema.parse({ ...prompt, effort }).effort).toBe(effort);
    expect(() => PromptMessageSchema.parse({ ...prompt, effort: "invalid" })).toThrow();
  });

  it("uses the advertised ACP config ID and supported effort", () => {
    expect(effortConfig([{ id: "reasoning_effort", type: "select", currentValue: "low", options: [{ value: "low" }, { value: "medium" }, { value: "high" }] }], "high"))
      .toEqual({ configId: "reasoning_effort", value: "high" });
  });

  it("recognizes thought-level categories and grouped values", () => {
    expect(effortConfig([{ id: "thinking", category: "thought_level", type: "select", options: [{ options: [{ value: "medium" }, { value: "high" }] }] }], "low"))
      .toEqual({ configId: "thinking", value: "medium" });
  });

  it("preserves defaults when effort is unsupported and avoids redundant writes", () => {
    expect(effortConfig(undefined, "medium")).toBeNull();
    expect(effortConfig([{ id: "model", type: "select", options: [{ value: "medium" }] }], "medium")).toBeNull();
    expect(effortConfig([{ id: "effort", type: "select", currentValue: "medium", options: [{ value: "medium" }] }], "medium")).toBeNull();
    expect(effortConfig([{ id: "effort", type: "boolean" }], "high")).toBeNull();
    expect(effortConfig([{ id: "effort", type: "select", options: [{ value: "default" }] }], "high")).toBeNull();
  });
});
