import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {stripTypeScriptTypes} from 'node:module';
import vm from 'node:vm';
import {webcrypto} from 'node:crypto';
import {onRequestPost} from '../functions/api/visit-request.js';

function submissionHandler(databaseError=null, uploadError=null) {
 let handler, stored; const uploads=[],removed=[];
 const source=readFileSync(new URL('../supabase/functions/submit-visit-request/index.ts',import.meta.url),'utf8').replace(/^import .*\n/,'');
 const context={Response,Request,FormData,TextDecoder,Uint8Array,crypto:webcrypto,createClient:()=>({storage:{from:()=>({upload:async(path,file,options)=>{uploads.push({path,file,options});return {error:uploadError};},remove:async paths=>{removed.push(...paths);return {error:null};}})},from:()=>({insert:async record=>{stored=record;return {error:databaseError};}})}),Deno:{env:{get:()=> 'test-server-secret'},serve:fn=>{handler=fn;}}};
 vm.runInNewContext(stripTypeScriptTypes(source),context);
 return {run:fields=>handler(new Request('https://example.invalid',{method:'POST',body:fields instanceof FormData ? fields : JSON.stringify(fields)})),stored:()=>stored,uploads,removed};
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

function attachmentForm(content='%PDF-1.7\nExample drawing',name='drawing.pdf') {
 const form=new FormData(); form.append('name','Upload Test'); form.append('email','test@example.invalid');
 form.append('attachments',new Blob([content],{type:'application/pdf'}),name); return form;
}
test('multipart attachments are validated, privately uploaded and linked to the request',async()=>{
 const h=submissionHandler(); const response=await h.run(attachmentForm());
 assert.equal(response.status,201); assert.equal(h.uploads.length,1);
 const file=h.stored().attachments[0];
 assert.equal(file.bucket,'visit-request-attachments'); assert.equal(file.name,'drawing.pdf');
 assert.ok(file.path.startsWith(h.stored().source_record_id+'/')); assert.equal(file.url,undefined);
 assert.equal(h.uploads[0].options.contentType,'application/pdf');
});
test('disguised files and excessive attachment counts are rejected before storage',async()=>{
 const h=submissionHandler(); assert.equal((await h.run(attachmentForm('<html>bad</html>'))).status,415);
 const form=attachmentForm(); for(let i=0;i<5;i++)form.append('attachments',new Blob(['%PDF-1.7']),`drawing${i}.pdf`);
 assert.equal((await h.run(form)).status,413); assert.equal(h.uploads.length,0); assert.equal(h.stored(),undefined);
});
test('a failed request insert removes uploaded files and never returns success',async()=>{
 const h=submissionHandler(new Error('private database failure')); assert.equal((await h.run(attachmentForm())).status,502);
 assert.equal(h.removed.length,1); assert.equal(h.removed[0],h.uploads[0].path);
 const failed=submissionHandler(null,new Error('upload failed')); assert.equal((await failed.run(attachmentForm())).status,502);
 assert.equal(failed.stored(),undefined);
});
test('website proxy forwards multipart files without replacing the multipart boundary',async()=>{
 const original=globalThis.fetch;
 try { globalThis.fetch=async(target,init)=>{assert.ok(init.body instanceof FormData); assert.equal(init.headers['Content-Type'],undefined); assert.equal(init.body.get('attachments').name,'drawing.pdf');return new Response('{}',{status:201});};
 assert.equal((await onRequestPost({request:new Request('https://example.invalid',{method:'POST',body:attachmentForm()})})).status,201);
 } finally {globalThis.fetch=original;}
});
