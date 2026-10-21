import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { SkillStore } from "../bridge/src/skills.js";
import { createExcelMcpServer } from "../excel_mcp/src/server.js";
import { HttpToolForwarder } from "../excel_mcp/src/forwarder.js";
import { PaneClient, getFreePort } from "./helpers.js";

describe("skill scripts through MCP, HTTP bridge and pane permission", () => {
  it("exposes the tool, respects denial, executes after approval and reads generated output", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "skill-script-e2e-"));
    fs.mkdirSync(path.join(root, "report", "tools"), { recursive: true });
    fs.writeFileSync(path.join(root, "report", "SKILL.md"), "---\nname: report\n---\nRead tools/layout.py and run it.");
    fs.writeFileSync(path.join(root, "report", "tools", "layout.py"), 'import json, sys\nresult = {"calls": [{"tool": "add_shape", "args": {"text": sys.argv[1]}}]}\nopen("result.json", "w").write(json.dumps(result))\nprint(json.dumps(result))');
    const port = await getFreePort();
    const bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" }), {
      skills: new SkillStore([root]), usageFile: null,
    });
    let pane: PaneClient | undefined;
    const client = new Client({ name: "skill-script-test", version: "1.0" });
    let mcp: ReturnType<typeof createExcelMcpServer> | undefined;
    try {
      const { url } = await bridge.listen();
      pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
      pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "skill-script-e2e" });
      const hello = await pane.waitFor(message => message.type === "hello_ok");
      if (hello.type !== "hello_ok") throw new Error("Missing handshake");
      let approved = false;
      const requests: string[] = [];
      pane.onMessage = message => {
        if (message.type === "tool_request") {
          requests.push(message.tool);
          pane!.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: { approved } } });
        }
      };
      const session = bridge.sessions.getByToken(hello.sessionToken)!;
      mcp = createExcelMcpServer(new HttpToolForwarder({ bridgeUrl: url, sessionId: session.id, token: hello.sessionToken }));
      const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
      await mcp.connect(serverTransport);
      await client.connect(clientTransport);
      expect((await client.listTools()).tools.map(tool => tool.name)).toContain("run_skill_script");
      const read = await client.callTool({ name: "read_skill", arguments: { name: "report" } });
      expect(JSON.stringify(read)).toContain("tools/layout.py");
      const args = { name: "report", file: "tools/layout.py", args: ["literal; $(not a command)"], outputFiles: ["result.json"] };
      const denied = await client.callTool({ name: "run_skill_script", arguments: args });
      expect(denied.isError).toBe(true);
      expect(fs.existsSync(path.join(root, "report", "result.json"))).toBe(false);
      approved = true;
      const result = await client.callTool({ name: "run_skill_script", arguments: args });
      expect(result.isError).not.toBe(true);
      const content = (result.content as Array<{ type: string; text: string }>)[0]!;
      const data = JSON.parse(content.text);
      expect(JSON.parse(data.stdout).calls[0].args.text).toBe("literal; $(not a command)");
      expect(data.exitCode).toBe(0);
      expect(data.outputFiles[0].path).toMatch(/^attachment:\/\//);
      const output = await client.callTool({ name: "read_attached_file", arguments: { path: data.outputFiles[0].path } });
      expect(output.isError).not.toBe(true);
      expect(JSON.stringify(output)).toContain("add_shape");
      expect(requests).toEqual(["__authorize_skill_script", "__authorize_skill_script"]);
    } finally {
      await client.close();
      await mcp?.close();
      await pane?.close();
      await bridge.close();
      fs.rmSync(root, { recursive: true, force: true });
    }
  });
});
