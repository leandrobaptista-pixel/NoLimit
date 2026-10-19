import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';
import {stripTypeScriptTypes} from 'node:module';
import {onRequestPost,onRequestGet,dimensions} from '../functions/api/site-metrics.js';
const origin='https://nolimitcontractor.net';
const request=(body,extra={})=>({request:new Request(origin+'/api/site-metrics',{method:'POST',headers:{Origin:origin,'Content-Type':'application/json',...extra},body:JSON.stringify(body)})});
test('proxy accepts aggregate events and never forwards network or identifying data',async()=>{
 const before=globalThis.fetch;let forwarded;
 try{globalThis.fetch=async(url,init)=>{forwarded=init;return new Response(null,{status:204});};
 assert.equal((await onRequestPost(request({event:'gallery_view',dimension:'stairs'},{'X-Forwarded-For':'192.0.2.1'}))).status,204);
 assert.deepEqual(JSON.parse(forwarded.body),{event:'gallery_view',dimension:'stairs'});
 assert.equal(forwarded.headers['X-Forwarded-For'],undefined);
 }finally{globalThis.fetch=before;}
});
test('unexpected fields, dimensions, oversized streaming bodies and foreign origins are rejected',async()=>{
 assert.equal((await onRequestPost(request({event:'page_view',dimension:'home',email:'private@example.invalid'}))).status,400);
 assert.equal((await onRequestPost(request({event:'page_view',dimension:'https://example.invalid?email=private'}))).status,400);
 assert.equal((await onRequestPost(request({event:'page_view',dimension:'home'},{Origin:'https://other.invalid'}))).status,403);
 assert.equal((await onRequestPost(request({event:'page_view',dimension:'x'.repeat(500)}))).status,413);
 assert.equal(onRequestGet().status,405);
});
test('privacy headers stop collection before any downstream call',async()=>{
 assert.equal((await onRequestPost(request({event:'page_view',dimension:'home'},{'Sec-GPC':'1'}))).status,204);
 assert.equal((await onRequestPost(request({event:'page_view',dimension:'home'},{DNT:'1'}))).status,204);
});
function frontend(overrides={}){
 const calls=[],flags=new Map();const listeners={};
 const document={visibilityState:'visible',referrer:'https://www.google.com/search?q=private',addEventListener:(name,fn)=>{listeners[name]=fn;},removeEventListener:()=>{},getElementById:()=>null};
 const window={};
 const context={location:{hostname:'nolimitcontractor.net',pathname:'/',search:'?cat=Ceiling&utm_campaign=private@example.invalid'},document,window,navigator:{},URL,URLSearchParams,Set,localStorage:{getItem:()=>null},sessionStorage:{getItem:key=>flags.get(key),setItem:(key,v)=>flags.set(key,v)},fetch:(url,init)=>{calls.push({url,init,body:JSON.parse(init.body)});return Promise.resolve({});},...overrides};
 vm.runInNewContext(readFileSync(new URL('../assets/site-metrics.js',import.meta.url),'utf8'),context);return {calls,flags,listeners};
}
test('frontend emits no URLs, campaign values, identities or visitor IDs; normalizes gallery names',()=>{
 const h=frontend();assert.ok(h.calls.some(x=>x.body.event==='gallery_view'&&x.body.dimension==='crown-molding-ceiling-coffered-ceiling'));
 assert.ok(h.calls.some(x=>x.body.event==='source'&&x.body.dimension==='google'));
 for(const {body,init} of h.calls){assert.deepEqual(Object.keys(body).sort(),['dimension','event']);assert.ok(dimensions[body.event].includes(body.dimension));assert.equal(init.credentials,'omit');assert.equal(init.referrerPolicy,'no-referrer');}
 assert.equal(JSON.stringify(h.calls).includes('private@example.invalid'),false);
 assert.equal([...h.flags.values()][0],'1');
});
test('GPC, DNT, opt-out, hidden tabs and administrative pages are not measured',()=>{
 for(const navigator of [{globalPrivacyControl:true},{doNotTrack:'1'}])assert.equal(frontend({navigator}).calls.length,0);
 assert.equal(frontend({localStorage:{getItem:()=> '1'}}).calls.length,0);
 assert.equal(frontend({document:{visibilityState:'hidden',addEventListener:()=>{},getElementById:()=>null}}).calls.length,0);
 assert.equal(frontend({location:{hostname:'localhost'}}).calls.length,0);
 assert.equal(frontend({location:{hostname:'nolimitcontractor.net',pathname:'/no-limit-admin-beta/'}}).calls.length,0);
});
test('Edge Function also rejects arbitrary fields and reports storage failure without success',async()=>{
 let handler;const source=readFileSync(new URL('../supabase/functions/record-site-metrics/index.ts',import.meta.url),'utf8').replace(/^import .*\n/,'');
 let fail=false;const calls=[];
 vm.runInNewContext(stripTypeScriptTypes(source),{Response,Uint8Array,TextDecoder,createClient:()=>({rpc:async(name,args)=>{calls.push({name,args});return {error:fail};}}),Deno:{env:{get:()=> 'private'},serve:fn=>handler=fn}});
 const run=body=>handler(new Request(origin,{method:'POST',body:JSON.stringify(body)}));
 assert.equal((await run({event:'page_view',dimension:'home',visitorId:'not-allowed'})).status,400);
 assert.equal(calls.length,0);
 assert.equal((await run({event:'page_view',dimension:'home'})).status,204);
 fail=true;assert.equal((await run({event:'page_view',dimension:'home'})).status,503);
});

test('the default gallery counts when it becomes visible, without requiring a category click',()=>{
 class Observer {constructor(callback){this.callback=callback;}observe(target){this.callback([{target,isIntersecting:true}]);}unobserve(){}}
 const document={visibilityState:'visible',referrer:'',addEventListener:()=>{},removeEventListener:()=>{},getElementById:id=>id==='gallery'?{id}:null,querySelector:()=>({dataset:{cat:'Trim'}})};
 const h=frontend({document,location:{hostname:'nolimitcontractor.net',pathname:'/',search:''},window:{IntersectionObserver:Observer},IntersectionObserver:Observer});
 assert.equal(h.calls.filter(x=>x.body.event==='gallery_view'&&x.body.dimension==='trim').length,1);
});
