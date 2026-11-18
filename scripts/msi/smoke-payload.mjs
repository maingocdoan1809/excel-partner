// Run from INSIDE the payload package folder (build-msi-payload.mjs copies it there and removes it again).
// Checks that nothing pruned from node_modules is still needed: every dependency and every compiled module must load.
// On Windows it also opens SQLite and loads koffi, which proves the bundled native binaries match the Node.js running this.
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

const failures = [];
const attempt = async (label, work) => { try { await work(); } catch (error) { failures.push(`${label}: ${String(error?.message ?? error).split("\n")[0]}`); } };

for (const specifier of ["@modelcontextprotocol/sdk/server/index.js", "@modelcontextprotocol/sdk/server/stdio.js", "@modelcontextprotocol/sdk/client/index.js",
  "@github/copilot-sdk", "zod", "ws", "yaml", "marked", "fflate", "fast-xml-parser", "selfsigned", "cross-spawn", "dompurify", "zod-to-json-schema"]) {
  await attempt(specifier, () => import(specifier));
}
await attempt("copilot-sdk client", async () => {
  const sdk = await import("@github/copilot-sdk");
  new sdk.CopilotClient({ connection: sdk.RuntimeConnection.forStdio({ path: "unused", args: [] }), logLevel: "error" });
});

// Entry points start servers or read the command line when imported; everything else must load on its own.
const entryPoint = /(?:^|[\\/])(index|cli|peer-cli|server|relay)\.m?js$/;
let modules = 0;
const walk = async (dir) => {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) await walk(file);
    else if (file.endsWith(".js") && !entryPoint.test(file)) { modules++; await attempt(path.relative(process.cwd(), file), () => import(pathToFileURL(file).href)); }
  }
};
for (const dir of ["dist/protocol", "dist/shared", "dist/providers", "dist/excel_mcp", "dist/bridge"]) await walk(dir);

const checkNative = process.platform === "win32" && !process.argv.includes("--no-native"); // an arm64 payload cannot load on an x64 machine
if (checkNative) {
  await attempt("better-sqlite3", async () => {
    const { createRequire } = await import("node:module");
    const Database = createRequire(import.meta.url)("better-sqlite3");
    const db = new Database(":memory:");
    db.exec("create table t(x); insert into t values (1)");
    if (db.prepare("select x from t").get().x !== 1) throw new Error("query returned the wrong row");
    db.close();
  });
  await attempt("koffi", async () => { const { createRequire } = await import("node:module"); if (typeof createRequire(import.meta.url)("koffi").load !== "function") throw new Error("koffi has no load()"); });
}

console.log(`smoke: ${modules} compiled modules, ${failures.length} failures${checkNative ? " (native modules checked)" : " (native modules not loaded: they are for another OS or architecture)"}`);
for (const failure of failures) console.log(`  FAIL ${failure}`);
process.exit(failures.length ? 1 : 0);
