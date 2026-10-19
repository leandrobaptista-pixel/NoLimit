import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {stripTypeScriptTypes} from 'node:module';
import vm from 'node:vm';
function helpers() {
 const source=readFileSync(new URL('../supabase/functions/manage-visit-requests/index.ts',import.meta.url),'utf8').replace(/^import .*\n/gm,'');
 let handler;
 const calls=[];
 const admin={storage:{from:bucket=>({createSignedUrl:async(path,seconds)=>{calls.push({bucket,path,seconds});return {data:{signedUrl:'https://example.invalid/private-signed-link'},error:null};}})},auth:{getUser:async()=>({data:{user:{id:'staff-id'}}})},from:()=>({select:()=>({eq(){return this;},maybeSingle:async()=>({data:{role:'client',status:'active'}})})})};
 const ctx={Response,Request,console,createClient:()=>admin,allowedStaffRoles:new Set(['admin','manager']),bearerToken:value=>value?.replace(/^Bearer /,''),Deno:{env:{get:()=> 'configured'},serve:fn=>{handler=fn;}}};
 vm.createContext(ctx); vm.runInContext(stripTypeScriptTypes(source),ctx);
 return {ctx,admin,calls,handler};
}
test('private attachment links are signed only for files belonging to the request',async()=>{
 const h=helpers(); h.ctx.admin=h.admin;
 h.ctx.item={source_record_id:'request-123',attachments:[{bucket:'visit-request-attachments',path:'request-123/file.pdf',name:'drawing.pdf'},{bucket:'visit-request-attachments',path:'different-request/file.pdf'}]};
 const item=await vm.runInContext('withAttachmentLinks(admin,item)',h.ctx);
 assert.equal(h.calls.length,1); assert.equal(h.calls[0].seconds,3600);
 assert.equal(item.attachments[0].url,'https://example.invalid/private-signed-link');
 assert.equal(item.attachments[1].url,null); assert.equal(item.attachments[1].path,null);
});
test('anonymous and non-staff callers cannot list requests or create attachment links',async()=>{
 const h=helpers();
 const request=auth=>new Request('https://example.invalid',{method:'POST',headers:auth?{Authorization:'Bearer token'}:{},body:JSON.stringify({action:'sync_and_list'})});
 assert.equal((await h.handler(request(false))).status,401);
 assert.equal((await h.handler(request(true))).status,403);
 assert.equal(h.calls.length,0);
});
