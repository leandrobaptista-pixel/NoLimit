import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import vm from 'node:vm';
import { JSDOM } from 'jsdom';

const root = new URL('../', import.meta.url);
const html = readFileSync(new URL('no-limit-admin-beta/index.html', root), 'utf8');
const app = readFileSync(new URL('no-limit-admin-beta/app.js', root), 'utf8');
const catalog = readFileSync(new URL('no-limit-admin-beta/service-catalog.js', root), 'utf8');

function workspace(t) {
  const dom = new JSDOM(html, { url: 'https://workspace.example.invalid/', runScripts: 'outside-only', pretendToBeVisual: true });
  const w = dom.window;
  w.scrollTo = () => {};
  w.confirm = () => true;
  w.alert = (message) => { w.lastAlert = message; };
  w.structuredClone = structuredClone;
  Object.defineProperty(w.crypto, 'randomUUID', { value: randomUUID });
  const context = dom.getInternalVMContext();
  const run = (source) => vm.runInContext(source, context);
  run(app);
  run(catalog);
  run(`currentBetaUser = { id: 'test-admin', role: 'admin', membershipRole: 'owner', name: 'Test Admin' };
    currentAuthSession = null;
    state = emptyOperationalState();
    savePreviewState = async () => true;`);
  t.after(() => w.close());
  return { w, run, json: (source) => JSON.parse(run(`JSON.stringify(${source})`)) };
}
const fields = { title: 'Baseboard installation', description: 'Install baseboard; labor only.', category: 'Trim', unit: 'linear foot', unitPrice: 5 };
const encoded = JSON.stringify(fields);

// Tests load the actual app and catalog scripts into the real entry-page DOM.
// No production requests are made and no real workspace is modified.
test('entry page exposes independent Line Items navigation and loads extension after app', () => {
  assert.match(html, /data-route="line-items"/);
  assert.ok(html.indexOf('app.js?v=__RELEASE__') < html.indexOf('service-catalog.js?v=__RELEASE__'));
});

test('create a reusable service with no invoice, estimate, client or project', async (t) => {
  const { run, json } = workspace(t);
  await run(`NoLimitServiceCatalog.save(${encoded})`);
  assert.equal(json('state.customServices').length, 1);
  for (const key of ['invoices', 'estimates', 'clients', 'projects', 'changeOrders']) assert.equal(json(`state.${key}`).length, 0);
  const service = json('state.customServices[0]');
  assert.equal(service.category, 'Trim');
  assert.equal(service.unit, 'linear foot');
  assert.equal(service.unitPrice, 5);
  assert.match(service.id, /^SVC-/);
});

test('editing and removing a catalog item preserve all document snapshots', async (t) => {
  const { run, json } = workspace(t);
  await run(`NoLimitServiceCatalog.save(${encoded})`);
  run(`const initial = NoLimitServiceCatalog.snapshot(state.customServices[0], { quantity: 100 });
    state.invoices = [{ id: 'INV-1', items: [structuredClone(initial)], total: 500, paid: 50, balance: 450 }];
    state.estimates = [{ id: 'EST-1', items: [structuredClone(initial)], total: 500 }];
    state.changeOrders = [{ id: 'CO-1', ...structuredClone(initial), amount: 500 }];`);
  const history = json('[state.invoices, state.estimates, state.changeOrders]');
  const id = json('state.customServices[0].id');
  await run(`NoLimitServiceCatalog.save({ ...${encoded}, title: 'Updated baseboard', category: 'Stairs', description: 'New scope', unit: 'each', unitPrice: 9 }, state.customServices[0].id)`);
  assert.equal(json('state.customServices[0].id'), id);
  assert.deepEqual(json('[state.invoices, state.estimates, state.changeOrders]'), history);
  await run('NoLimitServiceCatalog.setRemoved(state.customServices[0].id, true)');
  assert.deepEqual(json('[state.invoices, state.estimates, state.changeOrders]'), history);
  assert.equal(json('state.customServices[0].archived'), true);
});

test('removed services disappear from new choices but remain selected in historical documents', async (t) => {
  const { run, json } = workspace(t);
  await run(`NoLimitServiceCatalog.save(${encoded})`);
  const id = json('state.customServices[0].id');
  await run(`NoLimitServiceCatalog.setRemoved('${id}', true)`);
  assert.ok(!run('serviceOptions()').includes(`custom-service:${id}`));
  assert.match(run(`serviceOptions('Trim', '${id}')`), /selected disabled/);
  await run(`NoLimitServiceCatalog.setRemoved('${id}', false)`);
  assert.ok(run('serviceOptions()').includes(`custom-service:${id}`));
});

test('validation rejects blank, nonfinite, negative and duplicate input; zero price is valid', async (t) => {
  const { run, json } = workspace(t);
  for (const mutation of [{ title: '  ' }, { description: '' }, { unit: '' }, { unitPrice: '' }, { unitPrice: -1 }, { unitPrice: 'NaN' }, { unitPrice: 'Infinity' }, { category: 'not a category' }]) {
    await assert.rejects(run(`NoLimitServiceCatalog.save({ ...${encoded}, ...${JSON.stringify(mutation)} })`));
  }
  await run(`NoLimitServiceCatalog.save({ ...${encoded}, unitPrice: 0 })`);
  assert.equal(json('state.customServices[0].unitPrice'), 0);
  await assert.rejects(run(`NoLimitServiceCatalog.save({ ...${encoded}, title: ' BASEBOARD  installation ' })`));
  await run('NoLimitServiceCatalog.setRemoved(state.customServices[0].id, true)');
  await assert.rejects(run(`NoLimitServiceCatalog.save(${encoded})`), /Restore/i);
  assert.equal(json('state.customServices').length, 1);
});

test('server save failures roll back adds, edits and removal', async (t) => {
  const { run, json } = workspace(t);
  await run(`NoLimitServiceCatalog.save(${encoded})`);
  const before = json('state');
  run('savePreviewState = async () => false;');
  await assert.rejects(run(`NoLimitServiceCatalog.save({ ...${encoded}, title: 'A second service' })`), /not saved/);
  assert.deepEqual(json('state'), before);
  await assert.rejects(run(`NoLimitServiceCatalog.save({ ...${encoded}, unitPrice: 90 }, state.customServices[0].id)`));
  assert.deepEqual(json('state'), before);
  await assert.rejects(run('NoLimitServiceCatalog.setRemoved(state.customServices[0].id, true)'));
  assert.deepEqual(json('state'), before);
});

test('unauthorized accounts cannot create, edit or remove services', async (t) => {
  const { run, json } = workspace(t);
  await run(`NoLimitServiceCatalog.save(${encoded})`);
  const before = json('state');
  run(`currentBetaUser = { role: 'viewer', membershipRole: 'viewer' };`);
  await assert.rejects(run(`NoLimitServiceCatalog.save({ ...${encoded}, title: 'Denied' })`));
  await assert.rejects(run(`NoLimitServiceCatalog.save(${encoded}, state.customServices[0].id)`));
  await assert.rejects(run('NoLimitServiceCatalog.setRemoved(state.customServices[0].id, true)'));
  assert.deepEqual(json('state'), before);
  assert.doesNotMatch(run('NoLimitServiceCatalog.render()'), /data-catalog-add/);
});

test('new invoice lines copy category and unit and allow a per-document zero-price override', async (t) => {
  const { run, json } = workspace(t);
  await run(`NoLimitServiceCatalog.save(${encoded})`);
  run(`state.invoices = [{ id: 'INV-1', items: [], total: 0, paid: 0, balance: 0 }];
    pendingRecordType = 'invoiceItem';
    const invoiceForm = new FormData();
    Object.entries({ invoiceId: 'INV-1', category: 'custom-service:' + state.customServices[0].id,
      title: 'Baseboard at project', description: 'Project scope', quantity: '10', unitPrice: '0' }).forEach(([key, value]) => invoiceForm.set(key, value));`);
  await run('saveDataEntry(invoiceForm)');
  const item = json('state.invoices[0].items[0]');
  assert.equal(item.category, 'Trim');
  assert.equal(item.unit, 'linear foot');
  assert.equal(item.quantity, 10);
  assert.equal(item.unitPrice, 0);
  assert.equal(json('state.customServices[0].unitPrice'), 5);
  assert.equal(json('state.invoices[0].total'), 0);
});

test('estimate editing preserves subsequent lines and recalculates their subtotal', async (t) => {
  const { run, json } = workspace(t);
  await run(`NoLimitServiceCatalog.save(${encoded})`);
  run(`state.clients = [{ id: 'CL-1', name: 'Test Client' }];
    state.estimates = [{ id: 'EST-1', revision: 1, items: [NoLimitServiceCatalog.snapshot(state.customServices[0]), { category: 'Stairs', title: 'Other line', description: 'Keep me', quantity: 2, unitPrice: 25 }], total: 55 }];
    pendingRecordType = 'estimate'; pendingRecordId = 'EST-1';
    const estimateForm = new FormData();
    Object.entries({ clientId: 'CL-1', category: 'custom-service:' + state.customServices[0].id, title: 'Project baseboard', description: 'Project scope', quantity: '20', unitPrice: '5', discount: '10', taxPercent: '0', issueDate: '2026-09-30', validUntil: '2026-10-30', status: 'draft' }).forEach(([key, value]) => estimateForm.set(key, value));`);
  await run('saveDataEntry(estimateForm)');
  assert.equal(json('state.estimates[0].items').length, 2);
  assert.equal(json('state.estimates[0].items[1].description'), 'Keep me');
  assert.equal(json('state.estimates[0].total'), 140);
  assert.equal(json('state.estimates[0].revision'), 2);
});

test('legacy normalization follows IDs and does not resurrect renamed or removed services', async (t) => {
  const { run, json } = workspace(t);
  await run(`NoLimitServiceCatalog.save(${encoded})`);
  run(`state.changeOrders = [{ id: 'CO-1', ...NoLimitServiceCatalog.snapshot(state.customServices[0]), amount: 5 }];`);
  await run(`NoLimitServiceCatalog.save({ ...${encoded}, title: 'Renamed service' }, state.customServices[0].id)`);
  await run('NoLimitServiceCatalog.setRemoved(state.customServices[0].id, true)');
  const before = json('state');
  run('const normalized = normalizedSharedState(state);');
  assert.equal(json('normalized.customServices').length, 1);
  assert.equal(json('normalized.customServices[0].archived'), true);
  assert.equal(json('normalized.customServices[0].title'), 'Renamed service');
  assert.deepEqual(json('state'), before);
});

test('historical document editor retains snapshot values after catalog changes and removal', async (t) => {
  const { run, json } = workspace(t);
  await run(`NoLimitServiceCatalog.save(${encoded})`);
  run(`state.clients = [{ id: 'CL-1', name: 'Client', email: 'client@example.invalid' }];
    state.invoices = [{ id: 'INV-1', clientId: 'CL-1', clientName: 'Client', issueDate: 'Sep 30, 2026', dueDate: 'Oct 30, 2026', items: [NoLimitServiceCatalog.snapshot(state.customServices[0], { quantity: 10 })], schedule: [], total: 50, paid: 0, balance: 50 }];`);
  const snapshot = json('state.invoices[0].items[0]');
  await run(`NoLimitServiceCatalog.save({ ...${encoded}, category: 'Stairs', unit: 'each', unitPrice: 40 }, state.customServices[0].id)`);
  await run('NoLimitServiceCatalog.setRemoved(state.customServices[0].id, true)');
  run(`activeDocument = { type: 'invoice', id: 'INV-1', editing: false }; setDocumentEditing(true); harvestDocumentDraft();`);
  assert.deepEqual(json('activeDocument.draftItems[0]'), snapshot);
  assert.deepEqual(json('state.invoices[0].items[0]'), snapshot);
});

test('catalog form is independent and rendered user text is HTML-escaped', async (t) => {
  const { w, run } = workspace(t);
  await run(`NoLimitServiceCatalog.save({ ...${encoded}, title: '<img src=x onerror=alert(1)>' })`);
  assert.match(run('NoLimitServiceCatalog.render()'), /&lt;img/);
  assert.doesNotMatch(run('NoLimitServiceCatalog.render()'), /<img src=x/);
  w.document.getElementById('customServiceDialog').showModal = function () { this.setAttribute('open', ''); };
  run('NoLimitServiceCatalog.openEditor();');
  const form = w.document.getElementById('customServiceForm');
  for (const name of ['title', 'description', 'category', 'unit', 'unitPrice']) assert.ok(form.elements.namedItem(name));
  for (const name of ['invoiceId', 'clientId', 'projectId', 'requestId']) assert.equal(form.elements.namedItem(name), null);
});
