import {afterEach,expect,it} from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {McpServers} from '../runtime/src/mcp-servers.js';
const dirs: string[]=[];const managers:McpServers[]=[];
function setup(){const dir=fs.mkdtempSync(path.join(os.tmpdir(),'mcp-test-'));dirs.push(dir);const file=path.join(dir,'servers.json');const manager=new McpServers(file);managers.push(manager);return {manager,file};}
const config={id:'fixture',name:'Fixture',transport:'stdio' as const,target:process.execPath,args:[path.resolve('tests/support/mcp-fixture.mjs')]};
afterEach(async()=>{await Promise.all(managers.splice(0).map(m=>m.close()));for(const dir of dirs.splice(0))fs.rmSync(dir,{recursive:true,force:true});});
it('persists configurations without launching servers and rejects bad configurations',async()=>{const {manager,file}=setup();await manager.save(config);expect(new McpServers(file).list()).toEqual([config]);await expect(manager.save({...config,transport:'http',target:'file:///etc/passwd'})).rejects.toThrow();expect(manager.list()).toHaveLength(1);await manager.remove('fixture');expect(manager.list()).toEqual([]);});
it('discovers real stdio tools, validates arguments, preserves content and MCP errors',async()=>{const {manager}=setup();await manager.save(config);expect((await manager.tools('pane','fixture')).map(t=>t.name)).toEqual(['echo','fail','env']);await expect(manager.call('pane','fixture','echo',{text:42})).rejects.toThrow('Invalid MCP arguments');await expect(manager.call('pane','fixture','missing',{})).rejects.toThrow('not found');const result=await manager.call('pane','fixture','echo',{text:'hello'});expect(result).toMatchObject({structuredContent:{text:'hello'},content:[{type:'text',text:'hello'}]});expect(await manager.call('pane','fixture','fail',{})).toMatchObject({isError:true});await manager.closeSession('pane');expect(await manager.call('pane2','fixture','echo',{text:'new pane'})).toMatchObject({structuredContent:{text:'new pane'}});});
it('does not pass bridge or provider credentials into MCP child processes',async()=>{const {manager}=setup();await manager.save(config);const saved={bridge:process.env.EXCEL_SESSION_TOKEN,provider:process.env.ANTHROPIC_API_KEY};process.env.EXCEL_SESSION_TOKEN='test-secret';process.env.ANTHROPIC_API_KEY='provider-secret';try{expect(await manager.call('pane','fixture','env',{})).toMatchObject({content:[{text:'{}'}]});}finally{if(saved.bridge===undefined)delete process.env.EXCEL_SESSION_TOKEN;else process.env.EXCEL_SESSION_TOKEN=saved.bridge;if(saved.provider===undefined)delete process.env.ANTHROPIC_API_KEY;else process.env.ANTHROPIC_API_KEY=saved.provider;}});
it('rejects cancelled operations and removed servers',async()=>{const {manager}=setup();await manager.save(config);const controller=new AbortController();controller.abort();await expect(manager.tools('pane','fixture',controller.signal)).rejects.toThrow();await manager.tools('pane','fixture');await manager.remove('fixture');await expect(manager.call('pane','fixture','echo',{text:'no'})).rejects.toThrow('not found');});

it('connects to Streamable HTTP servers and discovers every page',async()=>{
 const http=await import('node:http');
 const {Server}=await import('@modelcontextprotocol/sdk/server/index.js');
 const {StreamableHTTPServerTransport}=await import('@modelcontextprotocol/sdk/server/streamableHttp.js');
 const {ListToolsRequestSchema,CallToolRequestSchema}=await import('@modelcontextprotocol/sdk/types.js');
 const server=http.createServer((req,res)=>{void(async()=>{
   const mcp=new Server({name:'http-fixture',version:'1'},{capabilities:{tools:{}}});
   mcp.setRequestHandler(ListToolsRequestSchema,async request=>({tools:[{name:request.params?.cursor?'second':'first',inputSchema:{type:'object'}}],...(request.params?.cursor?{}:{nextCursor:'next'})}));
   mcp.setRequestHandler(CallToolRequestSchema,async()=>({content:[{type:'text',text:'HTTP_RESULT'}]}));
   const transport=new StreamableHTTPServerTransport({sessionIdGenerator:undefined});
   await mcp.connect(transport);res.on('close',()=>{void mcp.close();});
   await transport.handleRequest(req,res);
 })().catch(()=>{if(!res.headersSent)res.writeHead(500);res.end();});});
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const {manager}=setup();
 try{await manager.save({id:'remote',name:'Remote',transport:'http',target:`http://127.0.0.1:${(server.address() as {port:number}).port}/mcp`});
 expect((await manager.tools('pane','remote')).map(tool=>tool.name)).toEqual(['first','second']);
 expect(await manager.call('pane','remote','second',{})).toMatchObject({content:[{text:'HTTP_RESULT'}]});
 }finally{await manager.close();await new Promise<void>(resolve=>server.close(()=>resolve()));}
});
