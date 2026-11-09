import { expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { ListToolsResultSchema } from "@modelcontextprotocol/sdk/types.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createExcelMcpServer } from "../excel_mcp/src/server.js";
import { McpServers } from "../runtime/src/mcp-servers.js";
import { discoverMcpCatalog } from "../runtime/src/mcp-catalog.js";
import { officeToolCatalog } from "../runtime/src/catalog.js";

it("a normal MCP client sees and calls actual server tools, then loses them when disabled", async () => {
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'direct-mcp-'));
 const manager=new McpServers(path.join(dir,'servers.json'));
 await manager.save({id:'demo',name:'Fixture MCP',transport:'stdio',target:process.execPath,args:[path.resolve('tests/support/mcp-fixture.mjs')]});
 const catalog=await discoverMcpCatalog(manager,'pane',new Set(['demo']));let enabled=true;
 const server=createExcelMcpServer({
  listTools:async()=>ListToolsResultSchema.parse({tools:[...officeToolCatalog('excel',false,{}),...(enabled?[...catalog.values()].map(tool=>tool.definition):[])]}).tools,
  call:async(name,args)=>{const entry=catalog.get(name)!;return{ok:true,data:{mcpResult:await manager.call('pane',entry.serverId,entry.toolName,args as Record<string,unknown>)}};},
 },{dynamicTools:true});
 const client=new Client({name:'ordinary-provider',version:'1'});const [a,b]=InMemoryTransport.createLinkedPair();
 try {
  await server.connect(a);await client.connect(b);
  let tools=(await client.listTools()).tools;
  const summary=tools.find(tool=>tool.name.endsWith('__echo'))!;
  expect(summary.inputSchema.properties).toHaveProperty('text');
  expect(tools.some(tool=>['mcp_list_servers','mcp_list_tools','mcp_call_tool'].includes(tool.name))).toBe(false);
  expect(await client.callTool({name:summary.name,arguments:{text:"hello"}})).toMatchObject({structuredContent:{text:"hello"}});
  enabled=false;tools=(await client.listTools()).tools;
  expect(tools.some(tool=>tool.name.startsWith('ext_'))).toBe(false);
  expect(await client.callTool({name:summary.name,arguments:{text:"hello"}})).toMatchObject({isError:true});
 } finally {await client.close();await server.close();await manager.close();fs.rmSync(dir,{recursive:true,force:true});}
});
