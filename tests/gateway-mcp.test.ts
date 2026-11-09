import { afterEach, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { GatewayServer } from "../gateway/src/server.js";
import { ConfigSchema } from "../gateway/src/config.js";
import { discoverMcpCatalog } from "../runtime/src/mcp-catalog.js";
import { McpServers } from "../runtime/src/mcp-servers.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { MockAdapter } from "../providers/src/mock.js";
import type { StartSessionContext, SendTurnInput } from "../providers/src/types.js";
import type { ProviderEvent } from "../protocol/src/index.js";

const cleanups:Array<()=>Promise<unknown>>=[];
afterEach(async()=>{for(const cleanup of cleanups.splice(0).reverse())await cleanup();});
class ApiAdapter extends MockAdapter {
 contexts:StartSessionContext[]=[]; external:any; official:any;
 constructor(){super({id:'copilot'});}
 override async startSession(ctx:StartSessionContext){this.contexts.push(ctx);return 'api-cli-session';}
 override async *sendTurn(input:SendTurnInput):AsyncIterable<ProviderEvent>{
  const ctx=this.contexts.at(-1)!;
  const call=async(name:string,args:object)=>{
   const response=await fetch(`${ctx.mcp.env.AGENT_GATEWAY_RELAY_URL}/internal/tool`,{method:'POST',headers:{authorization:`Bearer ${ctx.mcp.env.AGENT_GATEWAY_RELAY_TOKEN}`,'content-type':'application/json'},body:JSON.stringify({session:ctx.sessionId,name,arguments:args}),signal:input.signal});
   if(!response.ok)throw new Error('Relay failed');return response.json();
  };
  this.external=await call(ctx.mcp.toolNames!.find(name=>name.endsWith('__echo'))!,{text:'MCP_RESULT'});
  this.official=await call('official_office_read',{});
  yield {type:'message_delta',requestId:input.requestId,text:'OFFICIAL_CLIENT_DONE'};
  yield {type:'turn_finished',requestId:input.requestId,reason:'completed'};
 }
}
async function start(mode:'auto'|'deny'='auto'){
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'gateway-mcp-'));cleanups.push(async()=>fs.rmSync(dir,{recursive:true,force:true}));
 const file=path.join(dir,'mcp.json');const manager=new McpServers(file);
 await manager.save({id:'fixture',name:'Fixture',transport:'stdio',target:process.execPath,args:[path.resolve('tests/support/mcp-fixture.mjs')]});
 const catalog = await discoverMcpCatalog(manager, "test-catalog", new Set(["fixture"]));
 const registeredName = [...catalog.keys()].find(name => name.endsWith("__echo"))!; await manager.close();
 const adapter=new ApiAdapter();const gateway=new GatewayServer({config:{...ConfigSchema.parse({mcpConfigFile:file,enabledMcpServers:['fixture'],mcpApprovalMode:mode,warmupOnStart:false}),port:0},token:'official-client-test-token',tls:null,registry:new ProviderRegistry().register(adapter),conversationFile:path.join(dir,'conversations.json')});
 const url=await gateway.listen();cleanups.push(()=>gateway.close());
 const post=(messages:unknown[],host='excel',tools:unknown[]=[{name:'official_office_read',description:'Read from the official Office add-in',input_schema:{type:'object'}}])=>fetch(url+'/v1/messages',{method:'POST',headers:{'x-api-key':'official-client-test-token','content-type':'application/json'},body:JSON.stringify({model:'copilot/default',max_tokens:100,messages,metadata:{office_host:host},tools})});
 return{adapter,post,registeredName};
}
it.each(['excel','word','powerpoint'])('official Claude %s tools stay client-owned while gateway MCP runs server-side, without Office Partner',async host=>{
 const h=await start();const messages:any[]=[{role:'user',content:'Use external data, then read my Office file'}];
 const first=await h.post(messages,host);expect(first.status).toBe(200);const response=await first.json();
 expect(response.stop_reason).toBe('tool_use');expect(response.content.filter((b:any)=>b.type==='tool_use').map((b:any)=>b.name)).toEqual(['official_office_read']);
 expect(h.adapter.external).toMatchObject({content:[{text:'MCP_RESULT'}],structuredContent:{text:'MCP_RESULT'}});
 expect(h.adapter.contexts[0]!.mcp.toolNames).toEqual(expect.arrayContaining(['official_office_read',h.registeredName]));
 expect(h.adapter.contexts[0]!.mcp.toolNames).not.toContain('get_sheets');
 expect(h.adapter.contexts[0]!.mcp.toolNames).not.toContain('edit_doc_text');
 const call=response.content.find((b:any)=>b.type==='tool_use');
 const second=await h.post([...messages,{role:'assistant',content:response.content},{role:'user',content:[{type:'tool_result',tool_use_id:call.id,content:'OFFICIAL_OFFICE_RESULT'}]}],host);
 expect(second.status).toBe(200);expect(await second.json()).toMatchObject({stop_reason:'end_turn',content:[{text:'OFFICIAL_CLIENT_DONE'}]});
 expect(JSON.stringify(h.adapter.official)).toContain('OFFICIAL_OFFICE_RESULT');
 if(host!=='excel')expect(h.adapter.contexts[0]!.systemPrompt).not.toContain('edit_doc_text');
});
it('gateway denies external MCP calls unless explicitly authorized by gateway policy',async()=>{
 const h=await start('deny');const response=await(await h.post([{role:'user',content:'Use the tools'}])).json();
 expect(h.adapter.external.isError).toBe(true);expect(JSON.stringify(h.adapter.external)).toContain('not approved');
 expect(response.stop_reason).toBe('tool_use');expect(response.content.find((b:any)=>b.type==='tool_use').name).toBe('official_office_read');
});
it('rejects collisions without hijacking official-client tools',async()=>{
 const h=await start();const response=await h.post([{role:'user',content:'x'}],'excel',[{name:h.registeredName,input_schema:{type:'object'}}]);
 expect(response.status).toBe(400);expect(h.adapter.contexts).toHaveLength(0);
});
