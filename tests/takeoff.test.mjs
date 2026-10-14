import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
const ctx=vm.createContext({});
vm.runInContext(fs.readFileSync(new URL('../no-limit-admin-beta/takeoff.js',import.meta.url),'utf8'),ctx);
const api=ctx.NoLimitTakeoff;
const make=(values)=>({...api.template(),floors:values.map(([name,doorCasing,windowCasing,extension,base,doubleDoors,pocketDoors,singleDoors])=>({...api.blankFloor(name),doorCasing,windowCasing,extension,base,doubleDoors,pocketDoors,singleDoors}))});
const lincoln=()=>make([['2nd Floor',1000,550,640,575,6,5,13],['1st Floor',180,400,600,325,2,0,2],['Basement',550,0,0,555,3,0,11]]);
test('Lincoln quantities use reduced window casing and round base waste up',()=>{
 const t=api.totals(lincoln());assert.equal(t.net.casing,2680);assert.equal(t.net.base,1455);assert.equal(t.doors,42);
 assert.deepEqual(JSON.parse(JSON.stringify(t.order)),{casing:2948,extension:1364,base:1601});
});
test('Park quantities match the revised PDF and support zero waste',()=>{
 const model=make([['2nd',700,450,450,500,3,3,11],['1st',400,450,500,375,1,3,6],['Basement',450,0,0,490,2,0,9]]);
 assert.equal(api.totals(model).order.casing,2695);assert.equal(api.totals(model).doors,38);
 model.wastePercent=0;assert.equal(api.totals(model).order.casing,2450);
 model.floors[0].base=.1;model.floors[1].base=.2;model.floors[2].base=0;model.wastePercent=10;assert.equal(api.totals(model).order.base,1);
});
test('malformed numbers, negative measurements and fractional door counts are rejected',()=>{
 for(const bad of [-1,'garbage',Infinity]){const model=lincoln();model.floors[0].base=bad;assert.throws(()=>api.normalize(model));}
 const model=lincoln();model.floors[0].singleDoors=1.5;assert.throws(()=>api.normalize(model),/whole numbers/);
 model.floors[0].singleDoors=13;model.wastePercent=101;assert.throws(()=>api.normalize(model),/100%/);
});
test('template is blank, each draft is independent, and printable notes are escaped',()=>{
 const first=api.template(),second=api.template();first.floors[0].base=100;assert.equal(second.floors[0].base,0);
 first.notes='<script>bad</script>';const html=api.documentHTML(first,{name:'Project <test>'});assert.ok(!html.includes('<script>'));assert.ok(html.includes('&lt;script&gt;'));assert.ok(html.includes('takeoff-watermark'));
});
test('saving edits preserves other takeoffs and rolls back a rejected cloud save',async()=>{
 const list=[{...lincoln(),id:'keep'}];const other={...lincoln(),id:'new'};
 await api.persist(list,other,async()=>true);assert.equal(list.length,2);
 const before=JSON.stringify(list);await assert.rejects(api.persist(list,{...other,title:'Bad update'},async()=>false),/could not be saved/);assert.equal(JSON.stringify(list),before);
 await assert.rejects(api.persist(list,{...other,id:'bad'},async()=>{throw Error('offline')}),/offline/);assert.equal(JSON.stringify(list),before);
});
test('proposal attachments preserve a snapshot and rejected writes restore the revision',async()=>{
 const record={id:'EST-1',projectId:'p',revision:2,items:[{quantity:1,unitPrice:200}],total:200};const takeoff={...lincoln(),id:'TK-1'};
 await api.attach(record,takeoff,async()=>true);assert.equal(record.revision,3);takeoff.floors[0].base=999;assert.equal(record.materialTakeoff.floors[0].base,575);assert.equal(record.total,200);
 const before=JSON.stringify(record);await assert.rejects(api.attach(record,null,async()=>false),/could not be saved/);assert.equal(JSON.stringify(record),before);
 await api.attach(record,null,async()=>true);assert.equal(record.materialTakeoff,undefined);assert.equal(record.revision,4);
});
