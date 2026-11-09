import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { expect, it } from 'vitest';
import { Harness } from './support/harness.js';
import { CopilotNativeSessions } from '../providers/src/copilot-native.js';

let installed = true;
try { execFileSync(process.env.COPILOT_BIN ?? 'copilot', ['--version'], { stdio: 'ignore', timeout: 5000 }); }
catch { installed = false; }

// Real CLI/SDK/MCP transport; initializes schemas without making a model request.
it.skipIf(!installed)('builds current Copilot native schemas on/off/on without a model prompt', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'copilot-catalog-'));
  const h = await Harness.start('copilot', { configFile: path.join(dir, 'config.json') });
  const manager = new CopilotNativeSessions();
  try {
    const pane = h.bridge.sessions.list()[0]!;
    const headers = { authorization: `Bearer ${pane.sessionToken}`, 'content-type': 'application/json' };
    const post = async (body: object) => {
      const response = await fetch(h.url + '/session/mcp', { method: 'POST', headers, body: JSON.stringify(body) });
      expect(response.status).toBe(200);
    };
    await post({ action: 'save', config: { id: 'fixture', name: 'Fixture', transport: 'stdio', target: process.execPath, args: [path.resolve('tests/support/mcp-fixture.mjs')] } });
    for (const enabled of [true, false, true]) {
      await post({ action: 'select', id: 'fixture', enabled });
      const id = await manager.start({ workbookConnectionId: pane.workbookConnectionId, sessionId: pane.id,
        model: null, signal: new AbortController().signal,
        mcp: (h.bridge as any).mcpEndpointFor(pane), requestUserInput: async () => ({ cancelled: true, answers: {} }),
      }, process.env.COPILOT_BIN ?? 'copilot', process.env);
      const session = (manager as any).sessions.get(id).session;
      await session.rpc.tools.initializeAndValidate();
      const { tools } = await session.rpc.tools.getCurrentMetadata();
      expect(tools?.length).toBeGreaterThan(0);
      const external = tools.filter((tool: any) => tool.mcpToolName?.startsWith('ext_'));
      expect(external.map((tool: any) => tool.mcpToolName.split('__').at(-1)).sort()).toEqual(enabled ? ['echo', 'env', 'fail'] : []);
      for (const tool of external) {
        expect(tool.deferLoading).not.toBe(true);
        if (tool.mcpToolName.endsWith('__echo')) expect(tool.input_schema).toMatchObject({ required: ['text'], properties: { text: { type: 'string' } } });
      }
      await manager.close(id);
    }
  } finally {
    await manager.dispose(); await h.close(); fs.rmSync(dir, { recursive: true, force: true });
  }
}, 120_000);
