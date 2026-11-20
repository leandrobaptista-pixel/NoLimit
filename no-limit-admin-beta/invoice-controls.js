/* Invoices are client receivables, never project expense entries. */
(() => {
  const C = window.NoLimitProjectControls;
  function create(state, input, id, actor = '') {
    const get = key => input.get(key);
    const project = C.list(state, 'projects').find(p => p.id === get('projectId'));
    const client = C.list(state, 'clients').find(c => c.id === get('clientId'));
    if (!project || !client || project.clientId !== client.id) throw new Error('Select the client that owns this project.');
    const issueDate = C.day(get('issueDate')), dueDate = C.day(get('dueDate'));
    if (!issueDate || !dueDate || dueDate < issueDate) throw new Error('Choose valid invoice and due dates.');
    const reference = String(get('reference') || '').trim().slice(0, 200);
    if (reference && C.list(state, 'invoices').some(i => i.projectId === project.id && i.reference === reference && !['void','cancelled'].includes(i.status))) throw new Error('An invoice with this reference already exists for this project.');
    const titles = input.getAll('itemTitle'), descriptions = input.getAll('itemDescription'), quantities = input.getAll('itemQuantity'), prices = input.getAll('itemPrice');
    if (!titles.length || titles.length > 200 || [descriptions,quantities,prices].some(a => a.length !== titles.length)) throw new Error('Add at least one complete invoice line.');
    const items = titles.map((title, i) => {
      const quantity = Number(quantities[i]), unitPrice = Number(prices[i]);
      if (!String(title).trim() || !String(quantities[i]).trim() || !String(prices[i]).trim() || !Number.isFinite(quantity) || quantity <= 0 || !Number.isFinite(unitPrice) || unitPrice < 0 || quantity > 1e7 || unitPrice > 1e9) throw new Error('Each line needs a title, positive quantity and a valid price.');
      return { category:'Labor', title:String(title).trim().slice(0,300), description:String(descriptions[i] || '').slice(0,15000), quantity, unitPrice };
    });
    const totalCents = items.reduce((sum, item) => sum + Math.round(item.quantity * C.cents(item.unitPrice)), 0);
    if (totalCents <= 0 || !Number.isSafeInteger(totalCents)) throw new Error('Invoice total must be greater than zero.');
    const total = C.dollars(totalCents);
    const record = { id, projectId:project.id, clientId:client.id, clientName:client.name, contractId:'', reference, issueDate, dueDate, terms:String(get('terms') || '').slice(0,300), status:'draft', items, schedule:[], total, paid:0, balance:total, notes:String(get('notes') || '').slice(0,20000), createdAt:new Date().toISOString(), createdBy:actor, source:'direct-project-invoice' };
    return { state:{...state, invoices:[record,...C.list(state,'invoices')]}, record };
  }
  window.NoLimitInvoiceControls = { create };
})();
