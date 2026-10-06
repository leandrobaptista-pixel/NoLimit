import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
const context = vm.createContext({});
vm.runInContext(fs.readFileSync(new URL('../no-limit-admin-beta/service-catalog.js', import.meta.url), 'utf8'), context);
const api = context.NoLimitServiceCatalog;
const headers = Array.from(api.columns);
const rows = [headers, ['Door', 'Door install', 'Service', 'Quoted "door", with lock\nand trim', 'unit', 0], [null, 'Floor 1', 'Location', null, null, null], [null, null, null, null, null, null]];
test('QuickBooks columns preserve zero, blanks, units, types and original source', () => {
  const items = api.fromRows(rows, 'original.xlsx');
  assert.equal(items.length, 2);
  assert.equal(items[0].unitPrice, 0); assert.equal(items[0].unit, 'unit');
  assert.equal(items[1].unitPrice, null); assert.equal(items[1].unit, '');
  assert.equal(items[0].product, 'Door'); assert.equal(items[1].type, 'Location');
  assert.equal(items[1].importSource.values.Quantity, null);
  assert.equal(items[0].importSource.row, 2);
  assert.equal(api.billable(items[0]), true); assert.equal(api.billable(items[1]), false);
});
test('CSV handles quotes, commas, newlines and BOM', () => {
  const csv = '\uFEFF' + rows.slice(0, 3).map(row => row.map(cell => '"' + String(cell ?? '').replaceAll('"', '""') + '"').join(',')).join('\r\n');
  const items = api.parse(csv, 'services.csv');
  assert.equal(items[0].description, rows[1][3]); assert.equal(items[0].unitPrice, 0);
  assert.equal(items[1].unitPrice, null);
});
test('invalid headers, missing names and invalid prices fail before any import', () => {
  assert.throws(() => api.fromRows([['Name'], ['bad']]), /Expected columns/);
  assert.throws(() => api.fromRows([headers, ['Door', '', 'Service', '', '', 20]]), /name is missing/);
  for (const value of [-2, 'abc', Infinity, '2O.00']) assert.throws(() => api.price(value), /Invalid sales price/);
  assert.throws(() => api.parse(headers.join(',') + '\n"bad', 'test.csv'), /unclosed/);
});
test('repeat imports skip active and archived names without modifying existing records or invoices', () => {
  const existing = [{id:'keep', title:' DOOR INSTALL ', archived:true, unitPrice:999}];
  const invoices = [{items:[{title:'Door install',unitPrice:42}]}];
  const before = JSON.stringify({existing, invoices});
  const items = api.fromRows(rows);
  const plan = api.plan(existing, [...items, items[1]]);
  assert.equal(plan.additions.length, 1); assert.equal(plan.skipped.length, 2);
  assert.equal(JSON.stringify({existing, invoices}), before);
  assert.equal(api.plan([...existing,...plan.additions], items).additions.length, 0);
});
test('editing retains import metadata and rolls back catalog on rejected cloud save', async () => {
  const source = fs.readFileSync(new URL('../no-limit-admin-beta/app.js', import.meta.url), 'utf8');
  const existing = {id:'SVC-1',title:'Old',product:'Door',type:'Service',unitPrice:22,importSource:{file:'source.xlsx'},extraLegacy:'retain'};
  const originalDocument = {items:[{title:'Old',unitPrice:22}]};
  const form = new Map([['title','New'],['category','Door'],['product','Base'],['type','Customized'],['unit',''],['unitPrice',''],['description','']]);
  const ctx = vm.createContext({NoLimitServiceCatalog:api,state:{customServices:[existing]},editingCatalogServiceId:'SVC-1',pendingCustomServiceTarget:null,customServiceDialog:{close(){}},location:{hash:'#materials'},renderRoute(){},savePreviewState:async()=>true,nextId:()=> 'SVC-new',documentRecord:originalDocument});
  vm.runInContext(source.slice(source.indexOf('async function saveCustomService('),source.indexOf('\nfunction openCatalogServiceForm(')),ctx);
  await ctx.saveCustomService(form);
  assert.equal(ctx.state.customServices[0].importSource.file,'source.xlsx');
  assert.equal(ctx.state.customServices[0].extraLegacy,'retain');
  assert.equal(ctx.state.customServices[0].unitPrice,null); assert.equal(originalDocument.items[0].unitPrice,22);
  ctx.editingCatalogServiceId='SVC-1'; ctx.savePreviewState=async()=>false;
  const before=JSON.stringify(ctx.state.customServices);
  await assert.rejects(ctx.saveCustomService(new Map([...form,['title','Rejected']])), /could not be saved/);
  assert.equal(JSON.stringify(ctx.state.customServices),before);
});
test('catalog renders real table rows with imported blanks and location types', () => {
  const source=fs.readFileSync(new URL('../no-limit-admin-beta/app.js',import.meta.url),'utf8');
  const section=(a,b)=>source.slice(source.indexOf(a),source.indexOf(b,source.indexOf(a)));
  const ctx=vm.createContext({state:{customServices:api.fromRows(rows)},escapeHtml:value=>String(value??''),formatCurrency:value=>`$${value}`,emptyState:()=>'<p>Empty</p>'});
  vm.runInContext(section('function demoTable(', '\nfunction metric('),ctx);
  vm.runInContext(section('function renderServiceCatalogPanels(', '\nfunction renderMaterials('),ctx);
  const html=ctx.renderServiceCatalogPanels();
  assert.match(html,/Door install/);assert.match(html,/Location/);assert.match(html,/Not entered/);assert.match(html,/\$0/);
  assert.match(html,/PRODUCT/iu);assert.match(html,/data-import-catalog-services/);
});
