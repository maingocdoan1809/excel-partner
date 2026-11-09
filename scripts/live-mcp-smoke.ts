/** Opt-in paid smoke: real provider + real stdio MCP, isolated from user settings.
 * npx tsx scripts/live-mcp-smoke.ts copilot gpt-4.1
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { Harness } from '../tests/support/harness.js';
import type { ProviderId } from '../protocol/src/index.js';
const provider = process.argv[2] as ProviderId;
const model = process.argv[3];
assert.ok(['copilot', 'opencode', 'claude', 'codex'].includes(provider) && model, 'Specify provider and model');
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-live-mcp-'));
let h: Harness | undefined;
try {
  h = await Harness.start(provider, { configFile: path.join(dir, 'config.json'), allowOfficeJs: false });
  const token = h.bridge.sessions.list()[0]!.sessionToken;
  const headers = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };
  const post = async (body: object) => {
    const response = await fetch(h!.url + '/session/mcp', { method: 'POST', headers, body: JSON.stringify(body) });
    assert.ok(response.ok, `MCP configuration HTTP ${response.status}`);
  };
  await post({ action: 'save', config: { id: 'live-fixture', name: 'Live smoke fixture', transport: 'stdio', target: process.execPath, args: [path.resolve('tests/support/mcp-fixture.mjs')] } });
  const calls: { name: string; args: any }[] = [];
  const manager = (h.bridge as any).mcpServers;
  const originalCall = manager.call.bind(manager);
  manager.call = async (session: string, id: string, name: string, args: any, signal: AbortSignal) => {
    const result = await originalCall(session, id, name, args, signal);
    calls.push({ name, args });
    return result;
  };
  const normal = h.pane.onMessage;
  h.pane.onMessage = message => {
    if (message.type === 'tool_request' && message.tool === '__authorize_mcp') h!.pane.send({ type: 'tool_result', callId: message.callId, result: { ok: true, data: { approved: true } } });
    else normal?.(message);
  };
  const paneSend = h.pane.send.bind(h.pane);
  h.pane.send = message => paneSend(message.type === 'prompt' ? { ...message, model, effort: 'low' } : message);
  let native: string | undefined;
  for (const enabled of [true, false, true]) {
    await post({ action: 'select', id: 'live-fixture', enabled });
    const catalog = await (await fetch(h.url + '/internal/tools', { headers })).json();
    const tools = catalog.tools.filter((t: any) => t.name.startsWith('ext_'));
    assert.equal(tools.length, enabled ? 3 : 0);
    const nonce = randomUUID();
    const before = calls.length;
    const prompt = enabled
      ? `The external MCP server is enabled now. Call the registered external echo tool exactly once with text "${nonce}". Do not call other tools. Reply with just its result.`
      : `The external MCP server is now disabled. If an external echo tool is available, call it with text "${nonce}"; otherwise reply only UNAVAILABLE. Do not call other tools.`;
    const result = await h.turn(prompt, { timeoutMs: 180_000 });
    console.log(JSON.stringify({ provider, model, nativeSessionId: result.events.find(e => e.type === 'session_info')?.providerSessionId, enabled, reason: result.finished.reason, text: result.text, calls: calls.slice(before), events: result.events.filter(e => ['error', 'message_delta', 'message_completed', 'tool_started', 'tool_finished'].includes(e.type)) }));
    assert.equal(result.finished.reason, 'completed');
    const info = result.events.find(e => e.type === 'session_info');
    assert.ok(info && 'providerSessionId' in info);
    if (native) assert.equal(info.providerSessionId, native, 'native conversation must be preserved');
    native = info.providerSessionId;
    if (enabled) {
      assert.equal(calls.length - before, 1, 'must execute exactly one real MCP call');
      assert.deepEqual(calls.at(-1), { name: 'echo', args: { text: nonce } });
      assert.ok(result.text.includes(nonce));
    } else {
      assert.equal(calls.length, before, 'disabled server must not execute');
      assert.ok(result.text.includes('UNAVAILABLE'));
    }
  }
  console.log(`PASS ${provider}/${model}: on/off/on, actual stdio calls, same native conversation`);
} finally {
  await h?.close();
  fs.rmSync(dir, { recursive: true, force: true });
}
