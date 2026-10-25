import { build } from "esbuild";
import path from "node:path";
import { fileURLToPath } from "node:url";
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const result = await build({ absWorkingDir: root,
  entryPoints: { server: "gateway/src/index.ts", relay: "gateway/src/relay.ts", config: "gateway/src/config.ts" },
  outdir: "gateway/dist", outExtension: { ".js": ".mjs" },
  platform: "node", format: "esm", target: "node20", bundle: true, packages: "external", sourcemap: true, metafile: true,
});
console.log(`Built standalone gateway (${Object.keys(result.metafile.outputs).length} artifacts).`);
