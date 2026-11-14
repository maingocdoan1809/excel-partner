import { describe, expect, it } from "vitest";
import { hostSystemPrompt, recommendedPromptFor, RECOMMENDED_SYSTEM_PROMPT } from "../protocol/src/system-prompts.js";
import { ALL_TOOL_NAMES, toolsForHost } from "../protocol/src/index.js";
import { definitionForHost } from "../shared/src/host-overrides.js";
import { ALL_TOOL_DEFINITIONS } from "../shared/src/tool-definitions.js";

describe("host system prompts", () => {
  it("leave Excel's prompt exactly as the user has it", () => {
    expect(hostSystemPrompt("excel", "")).toBe("");
    expect(hostSystemPrompt("excel", "my rules")).toBe("my rules");
    expect(recommendedPromptFor("excel")).toBe(RECOMMENDED_SYSTEM_PROMPT);
  });

  it("give Word and PowerPoint a self-contained prompt when the user has none", () => {
    expect(hostSystemPrompt("word", "")).toMatch(/^You are an AI assistant working inside the user's open Word document/);
    expect(hostSystemPrompt("powerpoint", "  ")).toMatch(/^You are an AI assistant working inside the user's open PowerPoint presentation/);
    // The preset is the introduction only: the rules are added when the prompt is sent, so saving it changes nothing.
    expect(hostSystemPrompt("word", recommendedPromptFor("word"))).toBe(hostSystemPrompt("word", ""));
  });

  it("append the host rules to the user's own prompt", () => {
    const word = hostSystemPrompt("word", "Answer in Vietnamese.\n");
    expect(word.startsWith("Answer in Vietnamese.\n\n# Working in Word")).toBe(true);
    expect(hostSystemPrompt("powerpoint", "x")).toContain("# Working in PowerPoint");
  });

  it("mention only tools that exist for that host", () => {
    const names = new Set<string>(ALL_TOOL_NAMES);
    for (const host of ["word", "powerpoint"] as const) {
      const allowed = new Set(toolsForHost(host, { allowOptIn: true, playwright: true }));
      const mentioned = (hostSystemPrompt(host, "").match(/\b[a-z]+(?:_[a-z]+)+\b/g) ?? []).filter((word) => names.has(word));
      expect(mentioned.length).toBeGreaterThan(3);
      for (const tool of mentioned) expect(allowed.has(tool), `${host}: ${tool}`).toBe(true);
    }
  });
});

describe("tool definitions per host", () => {
  const get = (name: string) => ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name)!;

  it("keep Excel's wording for Excel and swap it for Word and PowerPoint", () => {
    expect(definitionForHost("excel", get("get_selection"))).toBe(get("get_selection"));
    expect(definitionForHost("excel", get("execute_office_js"))).toBe(get("execute_office_js"));
    const word = definitionForHost("word", get("get_selection"));
    expect(word.description).toContain("Word");
    expect(word.description).not.toMatch(/cell|sheet|workbook/i);
    const ppt = definitionForHost("powerpoint", get("get_selection"));
    expect(ppt.description).toContain("slides");
    expect(ppt.description).not.toMatch(/cell|sheet|workbook/i);
  });

  it("describe execute_office_js with the host's own namespace and keep its input schema", () => {
    const original = get("execute_office_js");
    for (const [host, name] of [["word", "Word"], ["powerpoint", "PowerPoint"]] as const) {
      const adapted = definitionForHost(host, original);
      expect(adapted.inputSchema.properties?.code?.description).toContain(`${name}.RequestContext`);
      expect(adapted.inputSchema.properties?.code?.description).not.toContain("Excel");
      expect(Object.keys(adapted.inputSchema.properties ?? {})).toEqual(Object.keys(original.inputSchema.properties ?? {}));
      expect(adapted.inputSchema.required).toEqual(original.inputSchema.required);
      expect(adapted.annotations.destructiveHint).toBe(true);
    }
    expect(original.inputSchema.properties?.code?.description).toContain("Excel");
  });

  it("leave other tools alone", () => {
    expect(definitionForHost("word", get("edit_doc_text"))).toBe(get("edit_doc_text"));
    expect(definitionForHost("powerpoint", get("run_python"))).toBe(get("run_python"));
  });
});
