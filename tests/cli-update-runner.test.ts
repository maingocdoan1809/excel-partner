import { afterEach, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { spawnSync } from "node:child_process";

const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });

/** Exercise the real detached runner, substituting only npm/process termination and the gateway executable. */
function run(options: { installError?: string; gatewayExit?: number; gateway?: boolean } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ep-runner-test-")); dirs.push(dir);
  const root = path.join(dir, "package"), data = path.join(dir, "data");
  fs.mkdirSync(path.join(root, "bin"), { recursive: true }); fs.mkdirSync(data);
  const calls = path.join(data, "calls.jsonl");
  fs.writeFileSync(path.join(root, "bin", "excel-partner.mjs"),
    `import fs from 'node:fs';fs.appendFileSync(${JSON.stringify(calls)},JSON.stringify(process.argv.slice(2))+'\\n');process.exit(process.argv[2] === 'stop' ? 0 : ${options.gatewayExit ?? 0});`);
  fs.copyFileSync("bin/update-runner.mjs", path.join(dir, "update-runner.mjs"));
  fs.writeFileSync(path.join(dir, "update.mjs"), `
    export { writeUpdateStatus, installedVersion } from ${JSON.stringify(pathToFileURL(path.resolve("bin/update.mjs")).href)};
    import fs from 'node:fs';
    export function stopPackageProcesses() { fs.appendFileSync(${JSON.stringify(calls)},JSON.stringify(['sweep'])+'\\n'); return { stopped: 0 }; }
    export async function updatePackage(options) {
      await options.stop();
      ${options.installError ? `throw new Error(${JSON.stringify(options.installError)});` : `return { installedRoot:${JSON.stringify(root)},version:'0.1.23' };`}
    }
  `);
  const result = spawnSync(process.execPath, [path.join(dir, "update-runner.mjs")], {
    env: { ...process.env, EP_ROOT: root, EP_DATA: data, EP_GATEWAY: options.gateway === false ? "0" : "1", EP_FROM: "0.1.22" },
    encoding: "utf8", timeout: 10000,
  });
  return { result, status: JSON.parse(fs.readFileSync(path.join(data, "update-status.json"), "utf8")),
    calls: fs.existsSync(calls) ? fs.readFileSync(calls, "utf8").trim().split("\n").map(line => JSON.parse(line)) : [] };
}

it("reports success only after the configured gateway has restarted", () => {
  const actual = run();
  expect(actual.result.status).toBe(0);
  expect(actual.status).toMatchObject({ state: "done", to: "0.1.23" });
  expect(actual.calls).toEqual([["stop"], ["sweep"], ["gateway", "start"]]);
});
it("reports gateway startup failure and does not retry it through the install recovery branch", () => {
  const actual = run({ gatewayExit: 1 });
  expect(actual.result.status).toBe(1);
  expect(actual.status).toMatchObject({ state: "failed", error: expect.stringContaining("Could not restart Agent Gateway") });
  expect(actual.calls).toEqual([["stop"], ["sweep"], ["gateway", "start"]]);
});
it.each(["npm install failed", "Package updated, but setup failed"])("recovers the gateway after %s", installError => {
  const actual = run({ installError });
  expect(actual.result.status).toBe(1);
  expect(actual.status).toMatchObject({ state: "failed", error: installError });
  expect(actual.calls).toEqual([["stop"], ["sweep"], ["gateway", "start"]]);
});
it("preserves the installation error as well as a gateway recovery failure", () => {
  const actual = run({ installError: "npm install failed", gatewayExit: 1 });
  expect(actual.status.error).toContain("npm install failed");
  expect(actual.status.error).toContain("Could not restart Agent Gateway");
});
it("does not start an unconfigured gateway during failure recovery", () => {
  const actual = run({ installError: "npm install failed", gateway: false });
  expect(actual.status.state).toBe("failed");
  expect(actual.calls).toEqual([["stop"], ["sweep"]]);
});
