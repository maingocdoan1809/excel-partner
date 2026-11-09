import {expect,it} from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {Bridge} from '../bridge/src/server.js';
import {loadConfig} from '../bridge/src/config.js';
import {PaneClient} from './helpers.js';
import {PROTOCOL_VERSION} from '../protocol/src/index.js';

it('isolates per-pane selection, rolls back failed enable, and revokes stale tools after editing or removal',async()=>{
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'mcp-pane-isolation-'));
 const bridge=new Bridge(loadConfig({port:0,tls:null,configFile:path.join(dir,'config.json'),logLevel:'error',referenceDirs:[],skillDirs:[]}),{usageFile:null,systemPromptFile:null,savedReferencesFile:null,workspaceConfigPath:null,updater:null});
 const panes:PaneClient[]=[];
 try{
  const {url}=await bridge.listen();
  const connect=async(name:string)=>{
   const pane=await PaneClient.connect(url.replace('http:','ws:')+'/ws');panes.push(pane);
   pane.send({type:'hello',protocolVersion:PROTOCOL_VERSION,workbookConnectionId:name,clientInfo:{host:'excel',standalone:true}});
   const hello=await pane.waitFor(m=>m.type==='hello_ok');if(hello.type!=='hello_ok')throw Error();
   const headers={authorization:`Bearer ${hello.sessionToken}`,'content-type':'application/json'};
   return{token:hello.sessionToken,post:(body:object,method='POST')=>fetch(url+'/session/mcp',{method,headers,body:JSON.stringify(body)}),catalog:async()=>((await(await fetch(url+'/internal/tools',{headers})).json()).tools as Array<{name:string}>),settings:async()=>await(await fetch(url+'/session/mcp',{headers})).json()};
  };
  const one=await connect('isolated-pane-one'),two=await connect('isolated-pane-two');
  const config={id:'fixture',name:'Fixture',transport:'stdio',target:process.execPath,args:[path.resolve('tests/support/mcp-fixture.mjs')]};
  expect((await one.post({action:'save',config})).ok).toBe(true);
  expect((await one.post({action:'select',id:'fixture',enabled:true})).ok).toBe(true);
  const first=await one.catalog();const echo=first.find(tool=>tool.name.endsWith('__echo'))!;expect(echo).toBeDefined();
  expect((await two.catalog()).some(tool=>tool.name.startsWith('ext_'))).toBe(false);
  await one.post({action:'save',config:{...config,id:'broken',target:path.join(dir,'missing-executable')}});
  expect((await one.post({action:'select',id:'broken',enabled:true})).status).toBe(400);
  expect((await one.settings()).enabled).toEqual(['fixture']);
  expect((await one.catalog()).some(tool=>tool.name===echo.name)).toBe(true);
  await two.post({action:'select',id:'fixture',enabled:true});
  await one.post({action:'save',config:{...config,name:'Edited fixture'}});
  expect((await one.settings()).enabled).toEqual([]);expect((await two.settings()).enabled).toEqual([]);
  const session=bridge.sessions.getByToken(one.token)!;
  expect(await bridge.gateway.invoke({sessionId:session.id,tool:echo.name,args:{text:'must not run'}})).toMatchObject({ok:false,error:{code:'unknown_tool'}});
  await one.post({action:'select',id:'fixture',enabled:true});
  expect((await one.catalog()).some(tool=>tool.name.endsWith('__echo'))).toBe(true);
  await one.post({id:'fixture'},'DELETE');
  expect((await one.catalog()).some(tool=>tool.name.startsWith('ext_'))).toBe(false);
  expect((await two.settings()).servers.some((server:any)=>server.id==='fixture')).toBe(false);
  expect(await bridge.gateway.invoke({sessionId:session.id,tool:echo.name,args:{text:'must not run'}})).toMatchObject({ok:false,error:{code:'unknown_tool'}});
 }finally{await Promise.all(panes.map(p=>p.close()));await bridge.close();fs.rmSync(dir,{recursive:true,force:true});}
});
