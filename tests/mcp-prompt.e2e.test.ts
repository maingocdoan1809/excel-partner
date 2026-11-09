import { expect, it, vi } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Bridge } from '../bridge/src/server.js';
import { loadConfig } from '../bridge/src/config.js';
import { MockAdapter } from '../providers/src/mock.js';
import { ProviderRegistry } from '../providers/src/registry.js';
import { PaneClient } from './helpers.js';
import { PROTOCOL_VERSION } from '../protocol/src/index.js';

it('refreshes native tool registration in the same conversation without injecting a discovery prompt', async () => {
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'mcp-turn-context-'));
 const adapter=new MockAdapter({id:'codex'});const turns=vi.spyOn(adapter,'sendTurn');const starts=vi.spyOn(adapter,'startSession');
 const bridge=new Bridge(loadConfig({port:0,tls:null,configFile:path.join(dir,'config.json'),skillDirs:[],referenceDirs:[],logLevel:'error'}),{registry:new ProviderRegistry().register(adapter),usageFile:null,systemPromptFile:null,savedReferencesFile:null,updater:null,workspaceConfigPath:null});
 let pane:PaneClient|undefined;
 try {
  const {url}=await bridge.listen();bridge.config.publicBaseUrl=url;
  pane=await PaneClient.connect(url.replace('http:','ws:')+'/ws');
  pane.send({type:'hello',protocolVersion:PROTOCOL_VERSION,workbookConnectionId:'mcp-discovery-test',clientInfo:{host:'excel',standalone:true}});
  const hello=await pane.waitFor(m=>m.type==='hello_ok');if(hello.type!=='hello_ok')throw Error();
  const post=(body:object)=>fetch(url+'/session/mcp',{method:'POST',headers:{authorization:`Bearer ${hello.sessionToken}`,'content-type':'application/json'},body:JSON.stringify(body)});
  await post({action:'save',config:{id:'demo',name:'Fixture MCP',transport:'stdio',target:process.execPath,args:[path.resolve('tests/support/mcp-fixture.mjs')]}});
  pane.send({type:'set_provider',provider:'codex'});await pane.waitFor(m=>m.type==='provider_status');
  pane.send({type:'prompt',requestId:'first',prompt:'Use Fixture MCP',selection:null,chatSessionId:'same-chat'});
  await pane.waitFor(m=>m.type==='turn_finished'&&m.requestId==='first');
  expect(turns.mock.calls[0]![0].prompt).not.toContain('mcp_list_servers');
  await post({action:'select',id:'demo',enabled:true});
  pane.send({type:'prompt',requestId:'second',prompt:'Use Fixture MCP for 10, 20, 30',selection:null,chatSessionId:'same-chat'});
  await pane.waitFor(m=>m.type==='turn_finished'&&m.requestId==='second');
  expect(turns.mock.calls[1]![0].prompt).not.toContain('enabledServers');
  expect(turns.mock.calls[1]![0].prompt).not.toContain('mcp_list_tools');
  expect(starts).toHaveBeenCalledTimes(2);
  expect(starts.mock.calls[1]![0].mcp.env.EXCEL_DYNAMIC_TOOLS).toBe("1");
  const token = hello.sessionToken;
  let catalog = await (await fetch(url + "/internal/tools", { headers: { authorization: `Bearer ${token}` } })).json();
  const tool = catalog.tools.find((tool: any) => tool.name.endsWith("__echo"));
  expect(tool.inputSchema.properties).toHaveProperty("text");
  await post({ action: "select", id: "demo", enabled: false });
  catalog = await (await fetch(url + "/internal/tools", { headers: { authorization: `Bearer ${token}` } })).json();
  expect(catalog.tools.some((tool: any) => tool.name.startsWith("ext_"))).toBe(false);
  expect(catalog.tools.some((tool: any) => tool.name === "mcp_call_tool")).toBe(false);
  pane.send({type:'prompt',requestId:'third',prompt:'Continue',selection:null,chatSessionId:'same-chat'});
  await pane.waitFor(m=>m.type==='turn_finished'&&m.requestId==='third');
  expect(starts).toHaveBeenCalledTimes(3);
  expect(starts.mock.calls[2]![0].mcp.env.EXCEL_DYNAMIC_TOOLS).toBeUndefined();
 }finally{await pane?.close();await bridge.close();fs.rmSync(dir,{recursive:true,force:true});}
});
