import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { notifyVisitRequest, visitRequestLink } from '../functions/_lib/visit-whatsapp.js';
import { onRequestPost } from '../functions/api/visit-request.js';
const enabled = {
  VISIT_WHATSAPP_ENABLED: 'true', VISIT_REQUEST_LINK_VERIFIED: 'true',
  META_WHATSAPP_MODE: 'test', META_TEST_SENDER_CONFIRMED: 'true',
  META_WHATSAPP_TOKEN: 'fake-token', META_PHONE_NUMBER_ID: '1234',
  META_GRAPH_VERSION: 'v23.0', META_WHATSAPP_RECIPIENTS: '15555550101,15555550102'
};
test('disabled, unverified, and unapproved configurations cannot send', async () => {
  const noSend = () => { throw Error('Must not send'); };
  for (const env of [{}, {...enabled, VISIT_REQUEST_LINK_VERIFIED:'false'},
    {...enabled, META_TEST_SENDER_CONFIRMED:'false'}, {...enabled, META_WHATSAPP_MODE:'production'},
    {...enabled, META_WHATSAPP_RECIPIENTS:'invalid'}]) {
    assert.deepEqual(await notifyVisitRequest(env, 'abc', noSend), []);
  }
});
test('test message uses configured recipients and exact private request link', async () => {
  const calls = [];
  await notifyVisitRequest(enabled, 'request-123', async (url, init) => {
    calls.push(JSON.parse(init.body)); return new Response('{}', {status:200});
  });
  assert.equal(calls.length, 2);
  assert.equal(calls[0].to, '15555550101');
  assert.match(calls[0].text.body, /\?visit=request-123#requests/);
  assert.equal(calls[0].type, 'text');
  assert.equal(new URL(visitRequestLink('x&admin=true')).searchParams.get('visit'), 'x&admin=true');
});
test('production template requires explicit cost approval and template configuration', async () => {
  let body;
  const env = {...enabled, META_WHATSAPP_MODE:'production', META_WHATSAPP_COST_APPROVED:'true', META_VISIT_TEMPLATE:'novo_pedido', META_TEMPLATE_LANGUAGE:'pt_BR', META_WHATSAPP_RECIPIENTS:'15555550101'};
  await notifyVisitRequest(env, 'abc', async (_, init) => {body = JSON.parse(init.body); return new Response('{}');});
  assert.equal(body.template.name, 'novo_pedido');
  assert.equal(body.template.components[0].parameters[0].text, visitRequestLink('abc'));
});
test('deferred integration never notifies, even after successful storage with enabled settings', async () => {
  const original = globalThis.fetch;
  const req = () => new Request('https://nolimitcontractor.net/api/visit-request', {method:'POST',body:JSON.stringify({name:'TEST',email:'test@example.invalid'})});
  try {
    let calls = 0;
    globalThis.fetch = async () => { calls++; return new Response('{}',{status:500}); };
    assert.equal((await onRequestPost({request:req(),env:enabled})).status, 502);
    assert.equal(calls,1);
    let notificationCalls = 0;
    globalThis.fetch = async (url) => {
      if (url.includes('supabase.co')) return new Response(null,{status:201});
      notificationCalls++;
      throw Error('Notifications must stay disconnected');
    };
    assert.equal((await onRequestPost({request:req(),env:enabled})).status,201);
    assert.equal(notificationCalls, 0);
  } finally { globalThis.fetch = original; }
});
test('deep link finds exact authorized record, including completed/archived requests', () => {
  const source = readFileSync(new URL('../no-limit-admin-beta/app.js', import.meta.url),'utf8');
  const start = source.indexOf('function linkedVisitRequest()');
  const end = source.indexOf('function renderVisitRequestCard',start);
  const context = {URLSearchParams,location:{search:'?visit=source-id'},visitIntake:{items:[{id:'admin-id',source_record_id:'source-id',status:'completed',discarded_at:'date'}]}};
  vm.createContext(context); vm.runInContext(source.slice(start,end),context);
  assert.equal(context.linkedVisitRequest().id,'admin-id');
  assert.equal(context.filteredVisitRequests().length,1);
  context.location.search='?visit=unauthorized';
  assert.equal(context.linkedVisitRequest(),null);
});
