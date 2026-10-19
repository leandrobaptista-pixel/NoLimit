import {test} from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const context=vm.createContext({window:{},Intl,Date,FormData,structuredClone,crypto:globalThis.crypto});
vm.runInContext(readFileSync(new URL('../no-limit-admin-beta/project-controls.js',import.meta.url),'utf8'),context);
const C=context.window.NoLimitProjectControls;
const base=()=>({projects:[{id:'p',name:'House',progress:50,contractValue:10000}],people:[{id:'v',name:'Carpenter',type:'Team Member',projectIds:[]}],projectEngagements:[],schedule:[],providerMilestones:[],providerPayments:[],payments:[],invoices:[],transactions:[]});
const terms={projectId:'p',personId:'v',scope:'Trim',basis:'daily',dailyRate:300,plannedDays:10,progressMode:'project',status:'active',advanceAllowance:0};
const add=(s,type,fields,id='')=>C.apply(s,type,fields,id,'tester','2026-10-23');
function agreement(s=base(),fields={}){return add(s,'engagement',{...terms,...fields});}
function work(s,e,fields={}){return add(s,'schedule',{engagementId:e.id,workDate:'2026-10-19',shift:'7–3',instructions:'Trim',status:'completed',units:1,dueDate:'2026-10-23',...fields});}
function payment(s,e,fields={}){return add(s,'providerPayment',{engagementId:e.id,date:'2026-10-23',amount:300,method:'ACH',reference:'TX-1',...fields});}
test('daily work links any registered provider to project and snapshots its rate',()=>{
 let {state:s,record:e}=agreement();assert.equal(s.people[0].projectIds[0],'p');
 s=work(s,e).state;s=add(s,'engagement',{...terms,dailyRate:400},e.id).state;
 assert.equal(C.engagementSummary(s,s.projectEngagements[0],'2026-10-23').earned,30000);
 assert.equal(C.engagementSummary(s,s.projectEngagements[0],'2026-10-23').budget,400000);
 const w=work(s,s.projectEngagements[0],{workDate:'2026-10-20',units:.5});assert.equal(C.workValue(w.record,s.projectEngagements[0]),20000);
});
test('confirmed attendance is planned, not incurred cost or a payable',()=>{
 const {state:s,record:e}=agreement();const x=work(s,e,{status:'confirmed'}).state;const summary=C.engagementSummary(x,e,'2026-10-23');assert.equal(summary.earned,0);assert.equal(summary.due,0);assert.equal(summary.plannedWork,30000);
});
test('partial weekly payments allocate once across completed daily charges',()=>{
 let {state:s,record:e}=agreement();s=work(s,e,{dueDate:'2026-10-19'}).state;s=work(s,e,{workDate:'2026-10-20',dueDate:'2026-10-20'}).state;
 s=payment(s,e,{amount:450}).state;let x=C.engagementSummary(s,e,'2026-10-23');assert.equal(x.paid,45000);assert.equal(x.due,15000);assert.equal(x.overdue,15000);assert.equal(x.earned,60000);
 s=payment(s,e,{amount:200}).state;x=C.engagementSummary(s,e,'2026-10-23');assert.equal(x.due,0);assert.equal(x.advanceBalance,5000);assert.equal(x.unapproved,5000);
});
test('fixed daily allocations never repeat the contract budget or daily payable',()=>{
 let {state:s,record:e}=agreement(base(),{basis:'fixed',contractAmount:2000});s=work(s,e,{workValue:400}).state;s=work(s,e,{workDate:'2026-10-20',workValue:500}).state;
 let x=C.engagementSummary(s,e,'2026-10-23');assert.equal(x.budget,200000);assert.equal(x.recorded,90000);assert.equal(x.earned,100000);assert.equal(x.due,0);assert.equal(x.paymentTermsMissing,true);
 s=add(s,'providerMilestone',{engagementId:e.id,label:'Week 1',amount:600,dueDate:'2026-10-22',status:'active'}).state;
 s=payment(s,e,{amount:250}).state;x=C.engagementSummary(s,e,'2026-10-23');assert.equal(x.due,35000);assert.equal(x.paid,25000);
 assert.throws(()=>add(s,'providerMilestone',{engagementId:e.id,label:'Extra',amount:1800,dueDate:'2026-10-22'}),/exceed/);
});
test('approved advances are explained, while excess and cash gaps stay visible',()=>{
 let {state:s,record:e}=agreement(base(),{advanceAllowance:300,advanceReason:'Mobilization'});s=payment(s,e).state;let x=C.projectSummary(s,s.projects[0],'2026-10-23');assert.ok(x.alerts.some(a=>a.code==='approved-advance'));assert.ok(!x.alerts.some(a=>a.code==='provider-advance'));assert.ok(x.alerts.some(a=>a.code==='client-progress'));assert.ok(x.alerts.some(a=>a.code==='cash-gap'));
 s=payment(s,e,{amount:100}).state;x=C.projectSummary(s,s.projects[0],'2026-10-23');assert.ok(x.alerts.some(a=>a.code==='provider-advance'&&a.amount===10000));
 assert.throws(()=>agreement(base(),{advanceAllowance:100}),/reason/);
});
test('weekly client cash uses dated receipts and avoids duplicate legacy transactions',()=>{
 let {state:s,record:e}=agreement();s.invoices=[{id:'i',projectId:'p',paid:2000}];s.payments=[{id:'a',invoiceId:'i',date:'2026-10-19',amount:500,status:'received'},{id:'b',invoiceId:'i',date:'2026-10-12',amount:1000,status:'cleared'},{id:'c',invoiceId:'i',date:'2026-10-20',amount:200,status:'pending'}];s.transactions=[{projectId:'p',type:'receivable',status:'paid',date:'2026-10-19',amount:500}];s=payment(s,e,{amount:200}).state;s=work(s,e).state;
 const r=C.weekly(s,'2026-10-23').projects[0];assert.equal(r.received,50000);assert.equal(r.paid,20000);assert.equal(r.net,30000);assert.equal(r.recordedWork,30000);assert.equal(r.undated,50000);assert.equal(C.clientCash(s,'p').total,200000);
 assert.equal(C.bounds('2026-10-25').start,'2026-10-19');assert.equal(C.bounds('2026-10-26').start,'2026-10-26');
});
test('paid daily charges cannot disappear; cash corrections require an auditable void',()=>{
 let {state:s,record:e}=agreement();let w=work(s,e);s=w.state;const p=payment(s,e,{chargeId:`day:${w.record.id}`});s=p.state;
 assert.throws(()=>add(s,'schedule',{...w.record,status:'cancelled'},w.record.id),/payments/);
 assert.throws(()=>add(s,'providerPayment',{...p.record,amount:900},p.record.id),/immutable/);
 s=add(s,'providerPayment',{engagementId:e.id,status:'void',notes:'Wrong bank reference'},p.record.id).state;
 assert.equal(s.providerPayments[0].amount,300);assert.equal(s.providerPayments[0].voidReason,'Wrong bank reference');assert.equal(C.engagementSummary(s,e,'2026-10-23').paid,0);
});
test('invalid dates, negative money and future actual entries cannot enter ledgers',()=>{
 const {state:s,record:e}=agreement();assert.equal(C.day('2026-02-30'),'');
 for(const amount of [-1,'bad',''])assert.throws(()=>payment(s,e,{amount}),/valid/);
 assert.throws(()=>payment(s,e,{date:'2026-10-24'}),/already made/);
 assert.throws(()=>work(s,e,{workDate:'2026-10-24'}),/Future work/);
 const fixed=agreement(base(),{basis:'fixed',contractAmount:100});assert.throws(()=>work(fixed.state,fixed.record),/daily contract work/);
});
test('CSV includes project/provider cash detail and neutralizes spreadsheet formula strings',()=>{
 const s=base();s.projects[0].name='=WEBSERVICE("secret")';const csv=C.csv(C.weekly(s,'2026-10-19'));assert.ok(csv.includes("'=WEBSERVICE"));assert.ok(csv.includes('Provider-only net cash'));assert.ok(csv.includes('2026-10-25'));
});
test('standalone receipts preserve legacy cash and remain additive to invoice balances',()=>{
 const s=base();s.transactions=[{id:'legacy',projectId:'p',type:'receivable',status:'paid',date:'2026-10-01',amount:1000}];s.payments=[{id:'new',projectId:'p',invoiceId:'',source:'client-receipt',date:'2026-10-19',status:'received',amount:200}];
 assert.equal(C.clientCash(s,'p').total,120000);assert.equal(C.clientCash(s,'p','2026-10-19','2026-10-25').total,20000);
 s.invoices=[{id:'invoice',projectId:'p',paid:500}];s.payments.push({id:'invoice-pay',invoiceId:'invoice',projectId:'p',date:'2026-10-19',amount:500,status:'cleared'});
 assert.equal(C.clientCash(s,'p').total,70000);assert.equal(C.weekly(s,'2026-10-19').projects[0].received,70000);
});
test('a failed shared save rolls back all project ledgers and does not claim success',async()=>{
 const app=readFileSync(new URL('../no-limit-admin-beta/app.js',import.meta.url),'utf8');
 const start=app.indexOf('async function saveDataEntry('),end=app.indexOf('\nfunction pageHead(',start);
 const ctx=vm.createContext({state:base(),window:{NoLimitProjectControls:C},ProjectControlsUI:{handles:()=>true},betaCanCreate:()=>true,pendingRecordType:'engagement',pendingRecordId:'',currentBetaUser:{id:'office'},savePreviewState:async()=>false});
 vm.runInContext(app.slice(start,end),ctx);ctx.input=terms;
 await assert.rejects(vm.runInContext('saveDataEntry(input)',ctx),/not saved/);
 assert.equal(ctx.state.projectEngagements.length,0);assert.equal(ctx.state.people[0].projectIds.length,0);
 ctx.savePreviewState=async()=>true;await vm.runInContext('saveDataEntry(input)',ctx);assert.equal(ctx.state.projectEngagements.length,1);assert.equal(ctx.pendingRecordId,ctx.state.projectEngagements[0].id);
});
