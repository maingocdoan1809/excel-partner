import { describe, expect, it } from "vitest";
import { ALL_TOOL_DEFINITIONS as TOOL_DEFINITIONS } from "../excel_mcp/src/tool-definitions.js";
import {
  EXCEL_TOOL_NAMES,
  ToolInputSchemas,
  ClientMessageSchema,
  ServerMessageSchema,
  columnToNumber,
  numberToColumn,
  parseRange,
  rangeCellCount,
  parseToolInput,
} from "../protocol/src/index.js";

describe("address helpers", () => {
  it("round-trips column numbers", () => {
    expect(columnToNumber("A")).toBe(1);
    expect(columnToNumber("Z")).toBe(26);
    expect(columnToNumber("AA")).toBe(27);
    expect(numberToColumn(27)).toBe("AA");
    expect(numberToColumn(1)).toBe("A");
  });

  it("parses ranges and counts cells", () => {
    expect(parseRange("A1")).toEqual({ start: { col: 1, row: 1 }, end: { col: 1, row: 1 } });
    expect(parseRange("$B$2:$D$10")).toEqual({
      start: { col: 2, row: 2 },
      end: { col: 4, row: 10 },
    });
    expect(rangeCellCount("A1:C3")).toBe(9);
    expect(rangeCellCount("not-a-range")).toBeNull();
  });
});

describe("tool input validation", () => {
  it("applies defaults for write_range", () => {
    const parsed = parseToolInput("write_range", {
      address: "A1",
      values: [[1, 2], [3, 4]],
    });
    expect(parsed.mode).toBe("values");
  });

  it("rejects a malformed address", () => {
    expect(() => parseToolInput("read_range", { address: "1A" })).toThrow();
  });

  it("validates insert_rows row reference", () => {
    expect(() => parseToolInput("insert_rows", { at: "A", count: 1 })).toThrow();
    expect(parseToolInput("insert_rows", { at: "A5", count: 2 }).position).toBe("above");
  });
});

describe("wire messages", () => {
  it("accepts a valid hello", () => {
    const message = ClientMessageSchema.parse({
      type: "hello",
      protocolVersion: 1,
      workbookConnectionId: "0123456789abcdef",
    });
    expect(message.type).toBe("hello");
  });

  it("rejects an unknown client message type", () => {
    expect(() => ClientMessageSchema.parse({ type: "nope" })).toThrow();
  });

  it("accepts a provider event from the bridge", () => {
    const message = ServerMessageSchema.parse({
      type: "message_delta",
      requestId: "r1",
      text: "hi",
    });
    expect(message.type).toBe("message_delta");
  });
});

describe("MCP tool definitions stay in sync with the zod schemas", () => {
  const unwrap = (schema: any): any => (schema._def?.schema ? unwrap(schema._def.schema) : schema);

  it("advertises exactly the tools the protocol knows", () => {
    expect(TOOL_DEFINITIONS.map((t) => t.name).sort()).toEqual([...EXCEL_TOOL_NAMES].sort());
  });

  it.each(EXCEL_TOOL_NAMES)("%s: JSON schema properties match the zod input", (name) => {
    const definition = TOOL_DEFINITIONS.find((t) => t.name === name)!;
    const zodKeys = Object.keys(unwrap(ToolInputSchemas[name]).shape).sort();
    expect(Object.keys(definition.inputSchema.properties ?? {}).sort()).toEqual(zodKeys);

    // Every required property must really be required by the validator.
    for (const key of definition.inputSchema.required ?? []) {
      const shape = unwrap(ToolInputSchemas[name]).shape;
      expect(shape[key].isOptional(), `${name}.${key} is required in JSON schema but optional in zod`).toBe(false);
    }
  });
});
