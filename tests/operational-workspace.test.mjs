import {test} from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const source=readFileSync(new URL('../no-limit-admin-beta/app.js',import.meta.url),'utf8');
const section=(a,b)=>source.slice(source.indexOf(a),source.indexOf(b,source.indexOf(a)));
test('missing cloud collections never fall back to sample records',()=>{
 const context=vm.createContext({defaultState:{people:[{name:'Demo'}],projects:[{name:'Demo project'}]},isLocalPreview:false});
 vm.runInContext('function cloneDefaultState(){return structuredClone(defaultState)};'+section('function emptyOperationalState()', '\nlet state =')+section('function normalizedSharedState(', '\nfunction betaRequestHeaders('),context);
 const result=vm.runInContext('normalizedSharedState({projects:[{id:"real"}]})',context);
 assert.equal(result.people.length,0);assert.equal(result.projects[0].id,'real');
});
test('partner routes exclude finance, requests and account administration',()=>{
 const context=vm.createContext({routes:{overview:1,security:1,financial:1,requests:1},currentBetaUser:{role:'vendor'},isLocalPreview:false});
 vm.runInContext(section('function routesForBetaRole(', '\nfunction canAccessAuditControls(')+section('function betaCanCreate(', '\nfunction applyBetaAccess('),context);
 for(const role of ['vendor','subcontractor','collaborator','viewer']){
  context.role=role;const routes=vm.runInContext('routesForBetaRole(role)',context);
  for(const denied of ['financial','security','requests']) assert.ok(!routes.includes(denied));
 }
 assert.equal(vm.runInContext('betaCanCreate("person")',context),false);
});
test('attachment uploads real bytes to a private bucket and rejects storage errors',async()=>{
 const calls=[];let failure=null;
 const client={storage:{from:bucket=>({upload:async(path,file,options)=>{
   calls.push({bucket,path,bytes:await file.text(),options});return {error:failure};
 }})}};
 const context=vm.createContext({File,crypto:globalThis.crypto,currentAuthSession:{user:{id:'staff'}},betaConfig:()=>({organizationId:'org'}),safeFileName:n=>n,window:{noLimitSupabaseClient:client}});
 vm.runInContext(section('async function uploadWorkspaceFile(', '\nfunction privateFileControl('),context);
 context.file=new File(['test contents'],'document.pdf',{type:'application/pdf'});
 const result=await vm.runInContext('uploadWorkspaceFile(file,"compliance-files","people/person")',context);
 assert.ok(result.path.startsWith('org/people/person/'));assert.equal(calls[0].bytes,'test contents');assert.equal(calls[0].options.upsert,false);
 failure={message:'Access denied'};
 await assert.rejects(vm.runInContext('uploadWorkspaceFile(file,"compliance-files","people/person")',context),/Access denied/);
 context.file=new File(['<script>'],'document.html');
 await assert.rejects(vm.runInContext('uploadWorkspaceFile(file,"project-files","receipts/p")',context),/Choose a PDF/);
});
