import {test} from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const ctx=vm.createContext({window:{},Intl,Date,FormData,structuredClone,crypto:globalThis.crypto});
for(const name of ['project-controls','invoice-controls'])vm.runInContext(readFileSync(new URL(`../no-limit-admin-beta/${name}.js`,import.meta.url),'utf8'),ctx);
const C=ctx.window.NoLimitProjectControls,I=ctx.window.NoLimitInvoiceControls;
function input(overrides={}){const f=new FormData();for(const [k,v] of Object.entries({clientId:'c',projectId:'p',issueDate:'2026-11-19',dueDate:'2026-12-19',reference:'QB-1144',...overrides}))f.append(k,v);for(const [title,price] of [['Basement',16086.10],['First Floor',26295.67],['Second Floor',21160.32],['Third Floor',5209.99]]){f.append('itemTitle',title);f.append('itemDescription','Labor only; owner supplies materials.');f.append('itemQuantity','1');f.append('itemPrice',price);}return f;}
const base=()=>({clients:[{id:'c',name:'Ingalls Builders'}],projects:[{id:'p',clientId:'c',name:'207 Washington',contractValue:68752.08,cost:0}],invoices:[],payments:[],transactions:[],expenseReceipts:[]});
test('direct project invoice preserves scope, exact total, and zero deposit without creating expenses or contracts',()=>{
 const s=base(),r=I.create(s,input(),'i','office');
 assert.equal(r.record.total,68752.08);assert.equal(r.record.balance,68752.08);assert.equal(r.record.items.length,4);assert.equal(r.record.paid,0);assert.equal(r.record.schedule.length,0);assert.equal(r.record.status,'draft');assert.equal(r.record.projectId,'p');assert.equal(r.state.payments.length,0);assert.equal(r.state.transactions.length,0);assert.equal(s.invoices.length,0);assert.equal(C.costBreakdown(r.state,s.projects[0]).total,0);
 assert.throws(()=>I.create(r.state,input(),'duplicate'),/reference already/);
});
test('invalid client/project, dates and line amounts cannot create an invoice',()=>{
 for(const o of [{clientId:'wrong'},{projectId:'wrong'},{dueDate:'2026-11-18'},{issueDate:'2026-02-30'}])assert.throws(()=>I.create(base(),input(o),'i'));
 for(const price of ['','bad','-1','Infinity']){const f=input();f.delete('itemPrice');for(let i=0;i<4;i++)f.append('itemPrice',price);assert.throws(()=>I.create(base(),f,'i'));}
});
test('expenses, wages, subcontractor earned costs and client billing reconcile independently',()=>{
 const s=base();s.people=[{id:'g',type:'Team Member',name:'Guy'},{id:'sub',type:'Subcontractor',name:'Sub'}];
 s.projectEngagements=[{id:'e',projectId:'p',personId:'g',basis:'daily',dailyRate:300,plannedDays:10},{id:'f',projectId:'p',personId:'sub',basis:'fixed',contractAmount:2000,progressMode:'scope',scopeProgress:50}];
 s.schedule=[{id:'w',projectId:'p',engagementId:'e',dailyRate:300,units:.5,status:'completed',workDate:'2026-11-18'},{id:'future',projectId:'p',engagementId:'e',dailyRate:300,units:1,status:'confirmed',workDate:'2026-11-20'}];
 s.providerPayments=[{id:'pay',engagementId:'e',projectId:'p',amount:150,date:'2026-11-19',status:'paid'}];
 s.expenseReceipts=[{id:'r',projectId:'p',category:'Materials',amount:120,purchaseDate:'2026-11-18'},{id:'other-project',projectId:'q',category:'Materials',amount:900}];
 s.transactions=[{id:'cost',source:'project-cost-entry',projectId:'p',type:'cost',category:'Equipment rental',amount:50,status:'paid'},{id:'cash',projectId:'p',type:'receivable',amount:1000,status:'paid'},{id:'duplicate-receipt',source:'project-cost-entry',projectId:'p',type:'cost',expenseReceiptId:'r',category:'Materials',amount:120}];
 s.materials=[{projectId:'p',status:'quoted',amount:10000}];s.projects[0].cost=10;
 const r=I.create(s,input(),'i').state,c=C.costBreakdown(r,s.projects[0],'2026-11-19');
 assert.equal(c.employees,15000);assert.equal(c.subcontractors,100000);assert.equal(c.materials,12000);assert.equal(c.other,5000);assert.equal(c.baseline,1000);assert.equal(c.total,133000);assert.equal(c.billed,6875208);
 assert.equal(C.costBreakdown({...r,providerPayments:[]},s.projects[0],'2026-11-19').total,c.total);
});
test('duplicate workday cannot double charge a provider',()=>{
 const s=base();s.people=[{id:'g',name:'Guy',type:'Team Member'}];s.projectEngagements=[{id:'e',projectId:'p',personId:'g',basis:'daily',dailyRate:300,plannedDays:10,status:'active'}];
 const fields={engagementId:'e',workDate:'2026-11-18',shift:'7-3',instructions:'Trim',status:'completed',units:1,dueDate:'2026-11-19'};
 const one=C.apply(s,'schedule',fields,'','office','2026-11-19');
 assert.throws(()=>C.apply(one.state,'schedule',fields,'','office','2026-11-19'),/already has work/);
 const edit=C.apply(one.state,'schedule',{...fields,units:.5},one.record.id,'office','2026-11-19');assert.equal(C.costBreakdown(edit.state,s.projects[0],'2026-11-19').employees,15000);
});
test('invoice persistence failure rolls back and retains no financial record',async()=>{
 const app=readFileSync(new URL('../no-limit-admin-beta/app.js',import.meta.url),'utf8'),start=app.indexOf('async function saveDataEntry('),end=app.indexOf('\nfunction pageHead(',start);
 const context=vm.createContext({state:base(),window:{NoLimitInvoiceControls:I},NoLimitAccess:{allows:()=>true},ProjectControlsUI:{handles:()=>false},pendingRecordType:'invoice',pendingRecordId:'',currentBetaUser:{id:'office'},nextId:()=> 'i',savePreviewState:async()=>false});
 vm.runInContext(app.slice(start,end),context);context.input=input();
 await assert.rejects(vm.runInContext('saveDataEntry(input)',context),/not saved/);assert.equal(context.state.invoices.length,0);
 context.savePreviewState=async()=>true;await vm.runInContext('saveDataEntry(input)',context);assert.equal(context.state.invoices.length,1);assert.equal(context.pendingRecordId,'i');
});
