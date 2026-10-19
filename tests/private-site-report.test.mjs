import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {stripTypeScriptTypes} from 'node:module';
import vm from 'node:vm';
function endpoint({role='owner',status='active',expired=false,fail=false}={}){
 let handler;const calls=[];
 const chain={eq(){return this;},maybeSingle:async()=>({data:{role,status},error:null})};
 const admin={auth:{getUser:async()=>({data:{user:expired?null:{id:'verified-user'}},error:expired})},from:()=>({select:()=>chain}),rpc:async(name,args)=>{calls.push({name,args});return {data:{traffic:[],requests_received:2},error:fail};}};
 const source=readFileSync(new URL('../supabase/functions/get-site-report/index.ts',import.meta.url),'utf8').replace(/^import .*\n/,'');
 vm.runInNewContext(stripTypeScriptTypes(source),{Response,createClient:()=>admin,Deno:{env:{get:()=> 'server-organization'},serve:fn=>handler=fn}});
 return {calls,run:(body={period:'today'},auth=true)=>handler(new Request('https://example.invalid',{method:'POST',headers:auth?{Authorization:'Bearer signed-token'}:{},body:JSON.stringify(body)}))};
}
test('anonymous, expired, inactive and ordinary staff accounts cannot read private reports',async()=>{
 const anonymous=endpoint();assert.equal((await anonymous.run({},false)).status,401);assert.equal(anonymous.calls.length,0);
 assert.equal((await endpoint({expired:true}).run()).status,401);
 for(const role of ['admin','manager','office','viewer','vendor','subcontractor']){
  const h=endpoint({role});const r=await h.run({period:'today',role:'owner'});assert.equal(r.status,403);assert.equal(h.calls.length,0);assert.equal((await r.json()).report,undefined);
 }
 assert.equal((await endpoint({role:'owner',status:'inactive'}).run()).status,403);
});
test('Owner and Developer report requests derive identity and organization on the server',async()=>{
 for(const role of ['owner','developer']){
  const h=endpoint({role});const r=await h.run({period:'week'});assert.equal(r.status,200);assert.equal(r.headers.get('cache-control'),'no-store');
  assert.equal(h.calls[0].args.actor_id,'verified-user');assert.equal(h.calls[0].args.organization,'server-organization');assert.equal(h.calls[0].args.report_period,'week');
  assert.equal((await r.json()).report.requests_received,2);
 }
});
test('caller-selected identities, invalid periods and database failures never produce a report',async()=>{
 const h=endpoint();assert.equal((await h.run({period:'today',actor_id:'another-user'})).status,400);
 assert.equal((await h.run({period:'all-history'})).status,400);assert.equal(h.calls.length,0);
 const failed=await endpoint({fail:true}).run();assert.equal(failed.status,503);assert.equal((await failed.json()).report,undefined);
});
const context={window:{},Intl,Date};vm.runInNewContext(readFileSync(new URL('../no-limit-admin-beta/site-insights.js',import.meta.url),'utf8'),context);
const ui=context.window.NoLimitSiteInsights;
test('the icon access rule follows membership role, not the broad Admin interface label',()=>{
 for(const membershipRole of ['admin','manager','office','viewer','vendor','team_member'])assert.equal(ui.canView({role:'admin',membershipRole}),false);
 assert.equal(ui.canView(null),false);assert.equal(ui.canView({membershipRole:'owner'}),true);assert.equal(ui.canView({membershipRole:'developer'}),true);
 const source=readFileSync(new URL('../no-limit-admin-beta/app.js',import.meta.url),'utf8');
 const fn=source.slice(source.indexOf('function routesForBetaRole'),source.indexOf('function betaCanCreate'));
 const ctx={isLocalPreview:false,currentBetaUser:{membershipRole:'admin'},routes:{overview:{},'website-insights':{}}};
 vm.createContext(ctx);vm.runInContext(fn,ctx);
 assert.equal(vm.runInContext("routesForBetaRole('admin').includes('website-insights')",ctx),false);
 ctx.currentBetaUser.membershipRole='owner';assert.equal(vm.runInContext("routesForBetaRole('admin').includes('website-insights')",ctx),true);
});
test('report output escapes account activity and distinguishes missing data and incomplete safeguards',()=>{
 const unloaded=ui.render({period:'today',data:null,loading:true,error:''});assert.ok(unloaded.includes('Reading the protected report'));assert.ok(unloaded.includes('—'));
 const data={period_start:'2026-10-19T04:00:00Z',period_end:'2026-10-19T23:00:00Z',measurement_started_at:'2026-10-19T12:00:00Z',traffic:[],staff_events:[{actor:'<script>bad()</script>',action:'workspace.records_removed',entity_type:'workspace',created_at:'2026-10-19T20:00:00Z',removed_counts:{media:1}}],security_controls:{mfa_enrolled:false},requests_received:0};
 const html=ui.render({period:'today',data,loading:false,error:''});assert.ok(html.includes('&lt;script&gt;'));assert.equal(html.includes('<script>bad()'),false);
 assert.ok(html.includes('Partial traffic coverage'));assert.ok(html.includes('Not enrolled'));assert.ok(html.includes('Not verified'));assert.ok(html.includes('Not configured here'));
 assert.ok(ui.render({period:'today',data:null},{preview:true}).includes('no live reports'));
});
