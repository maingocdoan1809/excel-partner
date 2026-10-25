import { describe, expect, it } from "vitest";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const cli = path.resolve("bin/excel-partner.mjs");
const run = (args: string[], home = fs.mkdtempSync(path.join(os.tmpdir(), "agw-cli-"))) =>
  spawnSync(process.execPath, [cli, ...args], { encoding: "utf8", env: { ...process.env, AGENT_GATEWAY_HOME: home } });

describe("excel-partner command targets", () => {
  it("lists the gateway targets in help", () => {
    const result = run(["help"]);
    expect(result.status).toBe(0);
    expect(result.stdout).toContain("[addin|gateway|all]");
    expect(result.stdout).toContain("setup gateway");
  });

  it("rejects an unknown target", () => {
    const result = run(["status", "banana"]);
    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain("Unknown target: banana");
  });

  it("explains how to set up a gateway that does not exist yet", () => {
    expect(run(["status", "gateway"]).stdout).toContain("excel-partner setup gateway");
    expect(run(["info", "gateway"]).stdout).toContain("excel-partner setup gateway");
  });

  it("passes gateway commands through to the gateway CLI", () => {
    const result = run(["gateway", "help"]);
    expect(result.status).toBe(0);
    expect(result.stdout).toContain("agent-gateway setup");
  });

  it("reports a missing gateway setup in doctor", () => {
    const result = run(["doctor", "gateway"]);
    expect(result.stdout).toContain("Not set up");
  });
});
