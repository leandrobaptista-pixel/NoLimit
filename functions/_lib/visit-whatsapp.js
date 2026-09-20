// Server-only adapter. Deployment alone never enables WhatsApp or charges.
export function visitRequestLink(id) {
  const url = new URL('https://nolimitcontractor.net/no-limit-admin-beta/');
  url.searchParams.set('visit', id);
  url.hash = 'requests';
  return url.href;
}

export async function notifyVisitRequest(env = {}, id, send = fetch) {
  if (env.VISIT_WHATSAPP_ENABLED !== 'true' || env.VISIT_REQUEST_LINK_VERIFIED !== 'true') return [];
  const test = env.META_WHATSAPP_MODE === 'test';
  if (test ? env.META_TEST_SENDER_CONFIRMED !== 'true' :
    env.META_WHATSAPP_MODE !== 'production' || env.META_WHATSAPP_COST_APPROVED !== 'true') return [];
  const recipients = [...new Set(String(env.META_WHATSAPP_RECIPIENTS || '').split(',').map(n => n.trim()).filter(Boolean))];
  // Recipients are server configuration, never supplied by the public form.
  if (!env.META_WHATSAPP_TOKEN || !/^\d+$/.test(env.META_PHONE_NUMBER_ID || '') ||
      !/^v\d+\.\d+$/.test(env.META_GRAPH_VERSION || '') || !recipients.length ||
      recipients.length > 2 || recipients.some(n => !/^[1-9]\d{7,14}$/.test(n)) ||
      (!test && (!/^[a-z0-9_]+$/.test(env.META_VISIT_TEMPLATE || '') || !env.META_TEMPLATE_LANGUAGE))) {
    console.error('Visit WhatsApp configuration incomplete.', id);
    return [];
  }
  const link = visitRequestLink(id);
  return Promise.all(recipients.map(async to => {
    const message = test ? {
      type: 'text', text: { preview_url: false, body: `TESTE — No Limit: novo pedido recebido no site. Abra para atender: ${link}` }
    } : {
      type: 'template', template: {
        name: env.META_VISIT_TEMPLATE,
        language: { code: env.META_TEMPLATE_LANGUAGE },
        components: [{ type: 'body', parameters: [{ type: 'text', text: link }] }]
      }
    };
    try {
      const response = await send(`https://graph.facebook.com/${env.META_GRAPH_VERSION}/${env.META_PHONE_NUMBER_ID}/messages`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${env.META_WHATSAPP_TOKEN}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ messaging_product: 'whatsapp', recipient_type: 'individual', to, ...message }),
        signal: AbortSignal.timeout(10000)
      });
      // API acceptance does not prove delivery. Never log tokens or full numbers.
      console[response.ok ? 'info' : 'error']('Visit WhatsApp provider result.', id, to.slice(-4), response.status);
      return { recipientSuffix: to.slice(-4), accepted: response.ok };
    } catch {
      console.error('Visit WhatsApp provider unavailable.', id, to.slice(-4));
      return { recipientSuffix: to.slice(-4), accepted: false };
    }
  }));
}
