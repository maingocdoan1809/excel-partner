import fs from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { PROVIDER_IDS } from "../protocol/src/providers.js";

/**
 * Which folders each folder may import from. Anything else is a wrong-way dependency.
 * bridge and gateway may import providers directly (registry, adapter types, prompt, login): shared/ holds only what both of them already
 * reuse (provider session manager, MCP and Office tool runtime), and re-exporting providers through it would add a layer with no logic.
 */
const MAY_IMPORT: Record<string, string[]> = {
  protocol: [],
  providers: ["protocol"],
  shared: ["protocol", "providers"],
  bridge: ["protocol", "providers", "shared"],
  gateway: ["protocol", "providers", "shared"],
  excel_mcp: ["protocol", "shared"],
  frontend: ["protocol"],
};

function sources(folder: string): string[] {
  const found: string[] = [];
  const walk = (directory: string) => {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      if (["node_modules", "dist"].includes(entry.name)) continue;
      const full = path.join(directory, entry.name);
      if (entry.isDirectory()) walk(full); else if (/\.ts$/.test(entry.name)) found.push(full);
    }
  };
  if (fs.existsSync(folder)) walk(folder);
  return found;
}

describe("the folders depend on each other in one direction only", () => {
  it.each(Object.entries(MAY_IMPORT))("%s imports only from %j", (folder, allowed) => {
    const wrong: string[] = [];
    for (const file of sources(folder)) {
      for (const match of fs.readFileSync(file, "utf8").matchAll(/(?:from|import)\s*\(?\s*["'](\.[^"']+)["']/g)) {
        const target = path.normalize(path.join(path.dirname(file), match[1]!)).split(path.sep)[0]!;
        if (target !== folder && target in MAY_IMPORT && !allowed.includes(target)) wrong.push(`${file.split(path.sep).join("/")} imports ${target}`);
      }
    }
    expect(wrong).toEqual([]);
  });

  it("the install and update tools only import each other, so a copy of bin/ runs on its own", () => {
    const wrong: string[] = [];
    for (const name of fs.readdirSync("bin")) {
      if (!/\.(?:mjs|cjs)$/.test(name)) continue;
      for (const match of fs.readFileSync(path.join("bin", name), "utf8").matchAll(/(?:from|require\()\s*["'](\.[^"']+)["']/g)) {
        // The recovery launcher is written out as text and names a folder, not a module.
        if (match[1]!.endsWith("/")) continue;
        const target = path.normalize(path.join("bin", match[1]!));
        if (path.dirname(target) !== "bin" || !fs.existsSync(target)) wrong.push(`bin/${name} imports ${match[1]}`);
      }
    }
    expect(wrong).toEqual([]);
  });

  it("the runtime manifest names every provider once and never requires a CLI older than the one it installs", () => {
    const manifest = JSON.parse(fs.readFileSync("bin/runtime-manifest.json", "utf8")) as { providers: Record<string, { version: string; minVersion?: string }> };
    expect(Object.keys(manifest.providers).sort()).toEqual([...PROVIDER_IDS].sort());
    const parts = (version: string) => version.split(".").map(Number);
    for (const [id, spec] of Object.entries(manifest.providers)) {
      expect(spec.minVersion, `${id} minVersion`).toBeTruthy();
      const [min, pinned] = [parts(spec.minVersion!), parts(spec.version)];
      const firstDifference = min.findIndex((value, index) => value !== pinned[index]);
      const exceeds = firstDifference !== -1 && min[firstDifference]! > pinned[firstDifference]!;
      expect(exceeds, `${id}: minVersion ${spec.minVersion} must not exceed version ${spec.version}`).toBe(false);
    }
  });

  it("the providers folder reaches the install tools in bin/ through profiled.ts only", () => {
    // Known, deliberate exception: bin/ must run on its own when the updater copies it, so the install code stays there and providers calls it by path.
    const reaching = sources("providers").filter(file => /["'`][^"'`]*\bbin["'`/]/.test(fs.readFileSync(file, "utf8")) || /runtime-store/.test(fs.readFileSync(file, "utf8")));
    expect(reaching.map(file => file.split(path.sep).join("/"))).toEqual(["providers/src/profiled.ts"]);
  });

  it("only the providers folder decides anything by the name of a provider", () => {
    // A provider name used as a value to compare, switch on, look up or test with: `=== "x"`, `case "x":`, `"x" ?`, `includes("x")`, `map["x"]`.
    const name = String.raw`["'](?:claude|copilot|codex|opencode)["']`;
    const named = new RegExp([String.raw`[!=]==\s*${name}`, String.raw`case\s+${name}\s*:`, String.raw`${name}\s*(?:\?|&&)`, String.raw`(?:includes|startsWith|indexOf)\(\s*${name}`, String.raw`\[${name}\]`].join("|"));
    const found: string[] = [];
    for (const folder of ["shared", "bridge", "gateway", "frontend", "protocol"]) {
      for (const file of sources(folder)) {
        fs.readFileSync(file, "utf8").split(/\r?\n/).forEach((line, index) => { if (named.test(line)) found.push(`${file.split(path.sep).join("/")}:${index + 1}`); });
      }
    }
    expect(found).toEqual([]);
  });
});
