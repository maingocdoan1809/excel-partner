import {afterEach,expect,it,vi} from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {McpServers} from '../runtime/src/mcp-servers.js';
import {discoverMcpCatalog,invokeRegisteredMcpTool} from '../runtime/src/mcp-catalog.js';
const cleanup:Array<()=>Promise<unknown>>=[];
afterEach(async()=>{for(const fn of cleanup.splice(0).reverse())await fn();});
async function setup(callTimeoutMs=60000){
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'mcp-hardening-'));cleanup.push(async()=>fs.rmSync(dir,{recursive:true,force:true}));
 const manager=new McpServers(path.join(dir,'config.json'),{callTimeoutMs});cleanup.push(()=>manager.close());
 const config={id:'fixture',name:'Fixture',transport:'stdio',target:process.execPath,args:[path.resolve('tests/support/mcp-fixture.mjs'),'--hardening']};
 await manager.save(config);return{manager,config};
}
it('identical tools from different servers have distinct stable names and actual schemas',async()=>{
 const{manager,config}=await setup();await manager.save({...config,id:'second',name:'Second'});
 const both=await discoverMcpCatalog(manager,'pane',new Set(['fixture','second']));
 const echoes=[...both.values()].filter(tool=>tool.toolName==='echo');expect(echoes).toHaveLength(2);expect(echoes[0]!.definition.name).not.toBe(echoes[1]!.definition.name);
 expect(echoes[0]!.definition.inputSchema.properties).toHaveProperty('text');
 const again=await discoverMcpCatalog(manager,'pane',new Set(['second','fixture']));expect([...again.keys()]).toEqual([...both.keys()]);
});
it('disabling a server while approval is pending prevents execution',async()=>{
 const{manager}=await setup();const catalog=await discoverMcpCatalog(manager,'pane',new Set(['fixture']));const tool=[...catalog.values()].find(t=>t.toolName==='hold')!;
 const enabled=new Set(['fixture']);let release!:(value:boolean)=>void;const approval=new Promise<boolean>(resolve=>{release=resolve;});
 const pending=invokeRegisteredMcpTool(manager,{sessionId:'pane',enabled,isActive:()=>true,authorize:()=>approval},tool,{});
 enabled.clear();release(true);expect(await pending).toMatchObject({ok:false,error:{code:'cancelled'}});
 expect(await manager.call('pane','fixture','state',{})).toMatchObject({structuredContent:{started:0}});
});
it('cancels a real in-flight request without replay and keeps other sessions isolated',async()=>{
 const{manager}=await setup();await manager.tools('one','fixture');await manager.tools('two','fixture');
 const controller=new AbortController();const pending=manager.call('one','fixture','hold',{},controller.signal);const rejection=expect(pending).rejects.toThrow();
 await vi.waitFor(async()=>expect(await manager.call('one','fixture','state',{})).toMatchObject({structuredContent:{started:1}}));
 controller.abort();await rejection;
 expect(await manager.call('one','fixture','state',{})).toMatchObject({structuredContent:{started:1}});
 expect(await manager.call('two','fixture','state',{})).toMatchObject({structuredContent:{started:0}});
});
it('bounds a real tool timeout without automatically repeating the action',async()=>{
 const{manager}=await setup(80);await manager.tools('pane','fixture');
 await expect(manager.call('pane','fixture','hold',{})).rejects.toThrow(/timeout|timed out/i);
 expect(await manager.call('pane','fixture','state',{})).toMatchObject({structuredContent:{started:1}});
});

it('pins validation and execution to one transport instead of reconnecting midway through a call',async()=>{
 const{manager}=await setup();
 const changed=vi.fn(async()=>({content:[{type:'text',text:'wrong process'}]}));
 const old={listTools:async()=>({tools:[{name:'echo',inputSchema:{type:'object'}}]}),callTool:vi.fn(async()=>{throw new Error('Connection closed after configuration changed');})};
 const client=vi.spyOn(manager as any,'client').mockResolvedValueOnce(old).mockResolvedValue({callTool:changed});
 await expect(manager.call('pane','fixture','echo',{})).rejects.toThrow('Connection closed');
 expect(changed).not.toHaveBeenCalled();expect(client).toHaveBeenCalledTimes(1);
});
