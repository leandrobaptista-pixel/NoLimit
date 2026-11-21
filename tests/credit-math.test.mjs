import {test} from 'node:test';
import assert from 'node:assert/strict';
import '../no-limit-admin-beta/credit-math.js';
const M=globalThis.NoLimitCreditMath;
const debt={id:'a',name:'Card',kind:'card',in_plan:true,opening_balance:1000,apr:0.24,minimum:50,due_date:'2026-11-15'};
test('monthly payments and statement payments use separate dates; voids are excluded',()=>{
  const data={settings:{budget:200},debts:[debt],entries:[
    {debt_id:'a',posted_on:'2026-10-30',payment:30,interest:0,charges:0,invoice_due:'2026-11-15'},
    {debt_id:'a',posted_on:'2026-11-01',payment:20,interest:10,charges:5,invoice_due:'2026-11-15'},
    {debt_id:'a',posted_on:'2026-11-02',payment:900,voided_at:'2026-11-03'},
    {debt_id:'a',posted_on:'2026-12-01',payment:900},
  ]};
  const s=M.summary(data,'2026-11','2026-11-21');
  assert.equal(s.rows[0].balance,965);assert.equal(s.paid,20);assert.equal(s.rows[0].invoicePaid,50);
  assert.equal(s.rows[0].reduction,5);assert.equal(s.reserve,0);assert.equal(s.rows[0].status,'Minimum covered');
});
test('missing and zero minimum remain distinct; past-due statement is exposed',()=>{
  const f=d=>M.summary({settings:{budget:0},debts:[d],entries:[]},'2026-11','2026-11-21');
  assert.equal(f({...debt,minimum:null}).reserve,null);
  assert.equal(f({...debt,minimum:0}).reserve,0);
  assert.equal(f(debt).rows[0].status,'Past due');assert.equal(f(debt).shortfall,50);
  assert.equal(f({...debt,opening_balance:0}).reserve,0);
});
test('projection pays all minima before highest APR and conserves cents',()=>{
  const p=M.project([{balance:100,apr:0,minimum:10},{balance:100,apr:0.12,minimum:10}],100);
  assert.equal(p.rows[0].interest,1);assert.equal(p.rows[0].payment,100);assert.equal(p.rows[0].balance,101);
  assert.equal(p.paidOff,true);assert.equal(p.rows.at(-1).balance,0);
  assert.ok(p.rows.every(r=>r.payment<=100));
});
test('insufficient budget stops instead of inventing payments',()=>{
  const p=M.project([{balance:1000,apr:0.24,minimum:50}],49.99);
  assert.equal(p.blockedMonth,1);assert.equal(p.shortfall,0.01);assert.equal(p.rows.length,0);
});
test('missing minimum uses workbook estimate; fixed zero is preserved',()=>{
  assert.equal(M.project([{balance:1000,apr:0,minimum:null}],34.99).shortfall,0.01);
  assert.equal(M.project([{balance:1000,apr:0,minimum:0}],0).blockedMonth,null);
  assert.equal(M.project([],0).paidOff,true);
});
test('financing is tracked separately from the selected card budget',()=>{
  const s=M.summary({settings:{budget:100},debts:[debt,{...debt,id:'b',kind:'auto',in_plan:false,minimum:900}],entries:[{debt_id:'b',posted_on:'2026-11-01',payment:800}]},'2026-11','2026-11-21');
  assert.equal(s.paid,800);assert.equal(s.available,100);assert.equal(s.planReserve,50);assert.equal(s.shortfall,0);
});
