// Optional probes. Importing the doctor itself needs only Node built-ins.
import { spawn } from "node:child_process";
import fs from "node:fs";
import https from "node:https";
import http from "node:http";
import path from "node:path";
import { createHash, randomUUID } from "node:crypto";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";

const RELEASE_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/**
 * A dependency from `<root>/node_modules` only. A bare import would keep looking in parent folders up to C:
ode_modules,
 * which any user can create, so a copy of bin/ without node_modules could load a planted module.
 */
export async function importFromRelease(root, specifier) {
  let resolved, modules;
  try { modules = fs.realpathSync(path.join(root, "node_modules")); resolved = createRequire(path.join(root, "package.json")).resolve(specifier); }
  catch { throw new Error(`DEPENDENCY_MISSING: ${specifier} is not installed in ${root}`); }
  if (!resolved.startsWith(modules + path.sep)) throw new Error(`DEPENDENCY_MISSING: ${specifier} is not installed in ${root}`);
  return import(pathToFileURL(resolved).href);
}

export async function runCommand(command, args, { env, timeoutMs = 8000, root = RELEASE_ROOT } = {}) {
  let nodeScript = /\.[cm]?js$/i.test(command);
  if (!nodeScript && path.isAbsolute(command)) {
    let handle;
    try { handle = fs.openSync(command, "r"); const header = Buffer.alloc(256); const size = fs.readSync(handle, header, 0, header.length, 0); nodeScript = /^#![^\r\n]*\bnode(?:\s|$)/.test(header.subarray(0, size).toString()); }
    catch { /* normal process error will be classified below */ }
    finally { if (handle !== undefined) fs.closeSync(handle); }
  }
  if (nodeScript) { args = [command, ...args]; command = process.execPath; }
  let launch = spawn;
  if (/\.(cmd|bat|m?js)$/i.test(command) || process.platform === "win32" && !/\.(exe|com)$/i.test(command)) {
    try { launch = (await importFromRelease(root, "cross-spawn")).default; }
    catch { return { code: null, errorCode: "RUNNER_MISSING", stdout: "", stderr: "" }; }
  }
  return new Promise(resolve => {
    let child, timer, settled = false, stdout = "", stderr = "";
    const finish = result => {
      if (settled) return;
      settled = true; clearTimeout(timer); resolve(result);
    };
    const kill = () => {
      if (!child?.pid) return;
      if (process.platform === "win32") {
        const killer = spawn("taskkill.exe", ["/pid", String(child.pid), "/t", "/f"], { windowsHide: true, stdio: "ignore" });
        killer.on("error", () => child.kill());
      } else child.kill("SIGKILL");
    };
    try { child = launch(command, args, { env, windowsHide: true, stdio: ["ignore", "pipe", "pipe"] }); }
    catch (error) { finish({ code: null, errorCode: error.code, stdout, stderr }); return; }
    for (const [stream, append] of [[child.stdout, text => { stdout += text; }], [child.stderr, text => { stderr += text; }]]) {
      stream?.on("data", chunk => {
        if (settled) return;
        if (stdout.length + stderr.length + chunk.length > 64 * 1024) {
          finish({ code: null, errorCode: "OUTPUT_LIMIT", stdout: "", stderr: "" }); kill(); return;
        }
        append(chunk.toString());
      });
    }
    timer = setTimeout(() => { finish({ code: null, errorCode: "ETIMEDOUT", stdout: "", stderr: "" }); kill(); }, timeoutMs);
    child.once("error", error => finish({ code: null, errorCode: error.code, stdout, stderr }));
    child.once("close", code => finish({ code, stdout, stderr }));
  });
}

// Never return the raw body (bridge health can contain document/session metadata).
export function requestHealth({ port, ca, timeoutMs = 2000 }) {
  return new Promise(resolve => {
    let settled = false, body = "", timer;
    const finish = value => { if (!settled) { settled = true; clearTimeout(timer); resolve(value); } };
    const req = (ca ? https : http).get({ hostname: "127.0.0.1", port, path: "/health", ...(ca ? { ca } : {}) }, res => {
      res.on("error", () => finish({ errorCode: "RESPONSE_ERROR" }));
      res.on("data", chunk => {
        body += chunk.toString();
        if (body.length > 256 * 1024) { finish({ errorCode: "OUTPUT_LIMIT" }); req.destroy(); }
      });
      res.on("end", () => {
        if (res.statusCode !== 200) return finish({ errorCode: "HTTP_STATUS" });
        try {
          const data = JSON.parse(body);
          finish({ bridge: data?.name === "excel-ai-addin-bridge", gateway: data?.service === "agent-gateway",
            ready: data?.ready === true, protocolVersion: Number.isInteger(data?.protocolVersion) ? data.protocolVersion : null });
        } catch { finish({ errorCode: "INVALID_JSON" }); }
      });
    });
    req.on("error", error => finish({ errorCode: error.code }));
    timer = setTimeout(() => { finish({ errorCode: "ETIMEDOUT" }); req.destroy(); }, timeoutMs);
  });
}

export async function probeMcp({ root, host, timeoutMs = 8000 }) {
  const entry = path.join(root, "dist", "excel_mcp", "src", "cli.js");
  if (!fs.existsSync(entry)) throw new Error("MCP_ENTRYPOINT_MISSING");
  const { Client } = await importFromRelease(root, "@modelcontextprotocol/sdk/client/index.js");
  const { StdioClientTransport } = await importFromRelease(root, "@modelcontextprotocol/sdk/client/stdio.js");
  const env = {};
  for (const key of ["PATH", "Path", "SystemRoot", "SYSTEMROOT", "TEMP", "TMP"]) {
    if (process.env[key]) env[key] = process.env[key];
  }
  Object.assign(env, { EXCEL_BRIDGE_URL: "http://127.0.0.1:1", EXCEL_SESSION_ID: "doctor",
    EXCEL_SESSION_TOKEN: randomUUID(), OFFICE_HOST: host, EXCEL_ALLOW_OFFICE_JS: "1" });
  const transport = new StdioClientTransport({ command: process.execPath,
    args: [entry], env, stderr: "pipe", maxBufferSize: 4 * 1024 * 1024 });
  const client = new Client({ name: "office-partner-doctor", version: "1" });
  transport.stderr?.on("data", () => {});
  const controller = new AbortController();
  let timer;
  const operation = async () => {
    // The stream is available on SDK transport construction with stderr: pipe.
    await client.connect(transport);
    const tools = [], cursors = new Set();
    let cursor;
    for (let page = 0; page < 20; page++) {
      controller.signal.throwIfAborted();
      const result = await client.listTools(cursor ? { cursor } : {}, { signal: controller.signal, timeout: timeoutMs });
      tools.push(...result.tools);
      if (tools.length > 2000) throw new Error("CATALOG_LIMIT");
      cursor = result.nextCursor;
      if (!cursor) break;
      if (cursors.has(cursor) || page === 19) throw new Error("CATALOG_PAGINATION");
      cursors.add(cursor);
    }
    const names = new Set();
    for (const tool of tools) {
      if (!tool.name || names.has(tool.name) || tool.inputSchema?.type !== "object") throw new Error("CATALOG_INVALID");
      names.add(tool.name);
    }
    if (!tools.length) throw new Error("CATALOG_EMPTY");
    const hostTool = { excel: "read_range", word: "get_doc_outline", powerpoint: "get_presentation_metadata" }[host];
    if (!hostTool || !names.has(hostTool)) throw new Error("CATALOG_HOST_MISMATCH");
    // Hash the complete definitions without exporting tool descriptions or schemas.
    const sorted = tools.sort((a, b) => a.name.localeCompare(b.name));
    return { count: tools.length, catalogHash: createHash("sha256").update(JSON.stringify(sorted)).digest("hex") };
  };
  try {
    return await Promise.race([operation(), new Promise((_, reject) => {
      timer = setTimeout(() => { controller.abort(); reject(new Error("ETIMEDOUT")); }, timeoutMs);
    })]);
  } finally {
    clearTimeout(timer); controller.abort();
    await client.close().catch(() => {});
    await transport.close().catch(() => {});
  }
}

export function readJson(file) {
  if (fs.statSync(file).size > 1024 * 1024) throw new Error("FILE_LIMIT");
  return JSON.parse(fs.readFileSync(file, "utf8"));
}

/** The newest start attempt recorded in startup.log (written by startup-log.mjs) for "addin" or "gateway", or null. */
export function readLastStartup(dataDir, service) {
  let lines;
  try { lines = fs.readFileSync(path.join(dataDir, "startup.log"), "utf8").split("\n").filter(Boolean); } catch { return null; }
  for (let i = lines.length - 1; i >= 0; i--) {
    try { const entry = JSON.parse(lines[i]); if (entry && entry.service === service) return entry; } catch { /* Skip a damaged line. */ }
  }
  return null;
}
