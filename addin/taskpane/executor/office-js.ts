import type { ExecuteOfficeJsInput } from "../../../protocol/src/index.js";
import { ToolFailure } from "./common.js";

/** Largest JSON result (characters) sent back to the model. */
const MAX_RESULT_CHARS = 100_000;

const AsyncFunction = Object.getPrototypeOf(async function () {
  /* constructor access only */
}).constructor as new (...args: string[]) => (...values: unknown[]) => Promise<unknown>;

/**
 * Opt-in escape hatch. The bridge only forwards this tool when it was started with
 * BRIDGE_ALLOW_OFFICE_JS=1, and the MCP server only lists it then. The code runs with the
 * task pane's Office access and can change or delete anything in the workbook.
 */
export async function executeOfficeJs(input: ExecuteOfficeJsInput): Promise<unknown> {
  let fn: (...values: unknown[]) => Promise<unknown>;
  try {
    fn = new AsyncFunction("context", "Excel", "Office", input.code);
  } catch (error) {
    throw new ToolFailure("invalid_arguments", `the code does not compile: ${error instanceof Error ? error.message : String(error)}`);
  }

  return Excel.run(async (context) => {
    const returned = await fn(context, Excel, Office);
    await context.sync();
    if (returned === undefined) return { result: null };
    let json: string;
    try {
      json = JSON.stringify(returned);
    } catch {
      throw new ToolFailure("invalid_arguments", "the returned value is not JSON-serialisable; return plain objects, numbers or strings");
    }
    if (json.length > MAX_RESULT_CHARS) {
      throw new ToolFailure("size_limit_exceeded", `the result is ${json.length} characters (limit ${MAX_RESULT_CHARS}); return less data`);
    }
    return { result: JSON.parse(json) as unknown };
  });
}
