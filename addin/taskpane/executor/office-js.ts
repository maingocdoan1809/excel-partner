import type { ExecuteOfficeJsInput, HostId } from "../../../protocol/src/index.js";
import { ToolFailure } from "./common.js";

/** Largest JSON result (characters) sent back to the model. */
const MAX_RESULT_CHARS = 30_000;

const AsyncFunction = Object.getPrototypeOf(async function () {
  /* constructor access only */
}).constructor as new (...args: string[]) => (...values: unknown[]) => Promise<unknown>;

type SyncContext = { sync(): Promise<void> };

function hostRunner(host: HostId): (batch: (context: SyncContext) => Promise<unknown>) => Promise<unknown> {
  const available = host === "word" ? typeof Word !== "undefined" : host === "powerpoint" ? typeof PowerPoint !== "undefined" : typeof Excel !== "undefined";
  if (!available) throw new ToolFailure("unsupported", `${namespaceOf(host)} is not available in this task pane`);
  if (host === "word") return (batch) => Word.run(batch as never);
  if (host === "powerpoint") return (batch) => PowerPoint.run(batch as never);
  return (batch) => Excel.run(batch as never);
}

function namespaceOf(host: HostId): string {
  return host === "word" ? "Word" : host === "powerpoint" ? "PowerPoint" : "Excel";
}

function hostNamespace(host: HostId): unknown {
  return host === "word" ? Word : host === "powerpoint" ? PowerPoint : Excel;
}

/**
 * Office.js escape hatch, enabled by default. BRIDGE_ALLOW_OFFICE_JS=0 disables
 * forwarding and MCP discovery. The code runs with the
 * task pane's Office access and can change or delete anything in the workbook.
 */
export async function executeOfficeJs(input: ExecuteOfficeJsInput, host: HostId = "excel"): Promise<unknown> {
  // The code sees `context`, the host's namespace under its own name (Excel, Word or PowerPoint) and `Office`.
  const namespace = host === "word" ? "Word" : host === "powerpoint" ? "PowerPoint" : "Excel";
  let fn: (...values: unknown[]) => Promise<unknown>;
  try {
    fn = new AsyncFunction("context", namespace, "Office", input.code);
  } catch (error) {
    throw new ToolFailure("invalid_arguments", `the code does not compile: ${error instanceof Error ? error.message : String(error)}`);
  }

  const runner = hostRunner(host);
  return runner(async (context) => {
    const returned = await fn(context, hostNamespace(host), Office);
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
