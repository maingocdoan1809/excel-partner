import { zodToJsonSchema } from "zod-to-json-schema";
import type { z } from "zod";
import { ToolInputSchemas, type ToolName } from "../../protocol/src/index.js";
import type { JsonSchema } from "./tool-definitions.js";

export const READ = { readOnlyHint: true, destructiveHint: false, idempotentHint: true } as const;
export const WRITE = { readOnlyHint: false, destructiveHint: true, idempotentHint: false } as const;

const HEX_PATTERN = "^#[0-9A-Fa-f]{6}$";

/**
 * Trim JSON Schema noise that costs model tokens on every turn without helping the model: length bounds,
 * closed-object markers and the colour regex (colours become `format: "#RRGGBB"`). The zod schema still
 * enforces all of it and returns a precise error when the model gets one wrong.
 */
function compact(node: unknown): unknown {
  if (Array.isArray(node)) return node.map(compact);
  if (node && typeof node === "object") {
    const out: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(node)) {
      if (key === "minLength" || key === "maxLength" || key === "$schema") continue;
      if (key === "additionalProperties" && value === false) continue;
      if (key === "pattern") {
        if (value === HEX_PATTERN) out.format = "#RRGGBB";
        continue;
      }
      out[key] = compact(value);
    }
    return out;
  }
  return node;
}

/** JSON Schema generated from the zod input, so MCP clients see what the validator accepts. */
export function schemaOf(name: ToolName): JsonSchema {
  const json = zodToJsonSchema(ToolInputSchemas[name] as z.ZodTypeAny, { target: "jsonSchema7", $refStrategy: "none" });
  return compact(json) as JsonSchema;
}

