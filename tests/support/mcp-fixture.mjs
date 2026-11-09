import {Server} from '@modelcontextprotocol/sdk/server/index.js';
import {StdioServerTransport} from '@modelcontextprotocol/sdk/server/stdio.js';
import {ListToolsRequestSchema,CallToolRequestSchema} from '@modelcontextprotocol/sdk/types.js';
const hardening=process.argv.includes('--hardening');
let started=0;
const server=new Server({name:'fixture',version:'1'},{capabilities:{tools:{}}});
server.setRequestHandler(ListToolsRequestSchema,async()=>({tools:[{name:'echo',description:'Echo text',inputSchema:{type:'object',properties:{text:{type:'string'}},required:['text'],additionalProperties:false}},{name:'fail',inputSchema:{type:'object'}},{name:'env',inputSchema:{type:'object'}},...(hardening?[{name:'hold',inputSchema:{type:'object'}},{name:'state',inputSchema:{type:'object'}}]:[])]}));
server.setRequestHandler(CallToolRequestSchema,async req=>{
 if(req.params.name==='state')return {content:[{type:'text',text:JSON.stringify({started})}],structuredContent:{started}};
 if(req.params.name==='hold'){started++;await new Promise(resolve=>setTimeout(resolve,2000));return {content:[{type:'text',text:'finished'}]};}
 if(req.params.name==='fail')return {isError:true,content:[{type:'text',text:'Fixture failure'}]};
 if(req.params.name==='env')return {content:[{type:'text',text:JSON.stringify({bridge:process.env.EXCEL_SESSION_TOKEN,provider:process.env.ANTHROPIC_API_KEY})}]};
 return {content:[{type:'text',text:req.params.arguments.text}],structuredContent:{text:req.params.arguments.text}};
});
await server.connect(new StdioServerTransport());
