import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {stripTypeScriptTypes} from 'node:module';
import vm from 'node:vm';
import {webcrypto} from 'node:crypto';
import {onRequestPost} from '../functions/api/visit-request.js';

function submissionHandler(databaseError=null) {
 let handler, stored;
 const source=readFileSync(new URL('../supabase/functions/submit-visit-request/index.ts',import.meta.url),'utf8').replace(/^import .*\n/,'');
 const context={Response,Request,crypto:webcrypto,createClient:()=>({from:()=>({insert:async record=>{stored=record;return {error:databaseError};}})}),Deno:{env:{get:()=> 'test-server-secret'},serve:fn=>{handler=fn;}}};
 vm.runInNewContext(stripTypeScriptTypes(source),context);
 return {run:fields=>handler(new Request('https://example.invalid',{method:'POST',body:JSON.stringify(fields)})),stored:()=>stored};
}
test('public submissions go directly into the private Admin intake and expose no saved details',async()=>{
 const h=submissionHandler();
 const response=await h.run({name:'Example Customer',email:'test@example.invalid',type:'Stairs',date:'2026-10-15',details:'Site visit'});
 assert.equal(response.status,201);
 assert.equal(h.stored().full_name,'Example Customer');
 assert.equal(h.stored().project_type,'Stairs');
 assert.equal(h.stored().preferred_date,'2026-10-15');
 assert.equal(h.stored().status,undefined);
 assert.equal(h.stored().internal_note,undefined);
 const result=await response.json();
 assert.deepEqual(Object.keys(result).sort(),['id','ok']);
 assert.match(result.id,/^[0-9a-f-]{36}$/);
});
test('invalid public submissions are rejected and storage failures never claim success',async()=>{
 const invalid=submissionHandler();
 assert.equal((await invalid.run({name:'Example',email:'not-an-email'})).status,400);
 assert.equal(invalid.stored(),undefined);
 const failed=submissionHandler(new Error('private failure'));
 const response=await failed.run({name:'Example',email:'test@example.invalid'});
 assert.equal(response.status,502);
 assert.ok(!(await response.text()).includes('private failure'));
});
test('website proxy forwards the same fields to the active submission service',async()=>{
 const original=globalThis.fetch;
 let url, body;
 try {
  globalThis.fetch=async (target,init)=>{url=target;body=JSON.parse(init.body);return new Response('{"ok":true}',{status:201});};
  const fields={name:'Example',email:'test@example.invalid',type:'Stairs',details:'Site visit'};
  const response=await onRequestPost({request:new Request('https://nolimitcontractor.net/api/visit-request',{method:'POST',body:JSON.stringify(fields)})});
  assert.equal(response.status,201);
  assert.equal(url,'https://powxzhcscqkhtaicxkyt.supabase.co/functions/v1/submit-visit-request');
  assert.deepEqual(body,fields);
 } finally {globalThis.fetch=original;}
});
