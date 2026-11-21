/* Sensitive data lives only in this controller's memory and protected RPCs. */
window.NoLimitCredit = (()=>{
  const M=window.NoLimitCreditMath;
  const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const usd=v=>v==null?'Not entered':new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(v);
  let data=null,loading=false,error='',notice='',generation=0,screen='summary',editing=null,month=M.today().slice(0,7),includeArchived=false,busy=false;
  let client=null,rerender=()=>{},allowed=()=>false,preview=null,entryDraft={};
  function reset(){data=null;preview=null;screen='summary';editing=null;entryDraft={};error='';notice='';loading=false;busy=false;generation++;}
  function configure(options){client=options.client;rerender=options.rerender;allowed=options.allowed;}
  async function load(){
    if(loading||!allowed())return;
    const request=generation;loading=true;error='';rerender();
    try{const r=await client.rpc('nl_credit_read');if(r.error)throw r.error;if(request===generation&&allowed())data=r.data;}
    catch(e){if(request===generation){data=null;error=e.message||'The private credit module could not be loaded.';}}
    finally{if(request===generation){loading=false;rerender();}}
  }
  async function write(operation,item){
    if(busy||!allowed())return;
    const request=generation;busy=true;error='';rerender();
    try{const {error:e}=await client.rpc('nl_credit_write',{operation,item});if(e)throw e;if(request!==generation||!allowed())return;screen='summary';editing=null;notice='Saved to the private credit ledger.';await load();}
    catch(e){if(request===generation)error=e.message||'Could not save. Refresh before retrying.';}
    finally{if(request===generation){busy=false;rerender();}}
  }
  function input(label,name,value='',type='text',extra=''){
    return `<label>${label}<input name="${name}" type="${type}" value="${esc(value??'')}" ${extra}></label>`;
  }
  const table=(headers,rows)=>`<div class="credit-scroll"><table><thead><tr>${headers.map(h=>`<th scope="col">${h}</th>`).join('')}</tr></thead><tbody>${rows.join('')||`<tr><td colspan="${headers.length}">No records yet.</td></tr>`}</tbody></table></div>`;
  function debtForm(){const d=editing||{};return `<article class="panel"><h2>${d.id?'Edit debt':'Add card or financing'}</h2><form id="creditDebtForm"><fieldset ${busy?'disabled':''}><div class="form-grid">
    ${input('Name','name',d.name,'text','required maxlength="120"')}
    <label>Type<select name="kind">${[['card','Credit card'],['mortgage','Mortgage'],['auto','Auto loan'],['loan','Loan'],['other','Other']].map(([k,l])=>`<option value="${k}" ${d.kind===k?'selected':''}>${l}</option>`).join('')}</select></label>
    <label><input name="in_plan" type="checkbox" ${d.in_plan?'checked':''}>Include this card in the monthly payoff budget</label>
    ${input('Opening debt balance (USD)','opening_balance',d.opening_balance,'number','required min="0" step="0.01"')}
    ${input('Annual APR (%)','apr',d.apr==null?'':Number(d.apr)*100,'number','required min="0" max="1000" step="0.0001"')}
    ${input('Current statement minimum / installment (USD)','minimum',d.minimum,'number','min="0" step="0.01"')}
    ${input('Current statement due date','due_date',d.due_date,'date')}
    ${input('Promotional balance to pay off (USD)','promo_balance',d.promo_balance,'number','min="0" step="0.01"')}
    ${input('Promotion deadline','promo_deadline',d.promo_deadline,'date')}
    ${input('Promotional installments remaining','promo_installments',d.promo_installments,'number','min="1" step="1"')}
    <label>Notes<textarea name="notes" maxlength="2000">${esc(d.notes)}</textarea></label></div>
    <p>Use the balance at the start of tracking. Later payments, interest and purchases belong in the ledger. For loans, exclude taxes and insurance paid outside the debt.</p>
    <div class="button-row"><button class="button" type="submit">Save debt</button><button class="button secondary" type="button" data-credit-action="cancel">Cancel</button></div></fieldset></form></article>`;}
  function entryForm(){return `<article class="panel"><h2>Post payment or charges</h2><form id="creditEntryForm"><fieldset ${busy?'disabled':''}><div class="form-grid">
    <label>Debt<select name="debt_id" required><option value="">Choose a debt</option>${data.debts.filter(d=>!d.archived).map(d=>`<option value="${d.id}" ${entryDraft.debt_id===d.id?'selected':''}>${esc(d.name)}</option>`).join('')}</select></label>
    ${input('Posting date','posted_on',entryDraft.posted_on||M.today(),'date',`required max="${M.today()}"`)}
    ${input('Payment applied to debt (USD)','payment',entryDraft.payment,'number','min="0" step="0.01"')}
    ${input('Actual interest charged (USD)','interest',entryDraft.interest,'number','min="0" step="0.01"')}
    ${input('Purchases and fees (USD)','charges',entryDraft.charges,'number','min="0" step="0.01"')}
    ${input('Statement due date this payment covers','invoice_due',entryDraft.invoice_due,'date')}
    <label>Notes<textarea name="notes" maxlength="2000">${esc(entryDraft.notes)}</textarea></label></div>
    <p>The statement due date links a payment to its minimum. Monthly totals use the posting date. Record actual interest separately; APR projections do not change the ledger.</p>
    <div class="button-row"><button class="button" type="submit">Post entry</button><button class="button secondary" type="button" data-credit-action="cancel">Cancel</button></div></fieldset></form></article>`;}
  function render(){
    if(!allowed())return '<section class="page"><article class="panel"><h1>Private credit access required</h1></article></section>';
    const message=`<p role="status" aria-live="polite">${esc(error||notice||'')}</p>`;
    const head=`<section class="page credit-page"><div class="page-head"><div><p class="page-kicker">Private credit ledger · USD</p><h1>Cards &amp; Financing</h1><p>Shared exclusively by Leandro Baptista and Geanniny Freitas.</p></div><button class="button secondary" data-credit-action="refresh" ${busy||loading?'disabled':''}>Refresh</button></div>${message}`;
    if(loading)return head+'<article class="panel"><p role="status">Loading protected records…</p></article></section>';
    if(!data)return head+'<article class="panel"><p>The database must be activated for the two authorized accounts before this module can load.</p><button class="button" data-credit-action="refresh">Try again</button></article></section>';
    if(screen==='debt')return head+debtForm()+'</section>';
    if(screen==='entry')return head+entryForm()+'</section>';
    const s=M.summary(data,month);const rows=s.rows.filter(d=>includeArchived||!d.archived);
    const projection=M.project(s.active.filter(d=>d.in_plan),data.settings.budget);
    const dueMissing=s.active.some(d=>d.balance>0&&(d.minimum==null||!d.due_date));
    const totals=[['All debt balances',s.balance],['Paid in selected month',s.paid],['Card-plan budget remaining',s.available],['All current minimums remaining',s.reserve],['Card-plan shortfall',s.shortfall]];
    return head+`<div class="credit-toolbar"><label>Reporting month<input id="creditMonth" type="month" value="${month}" required></label><button class="button" data-credit-action="add" ${busy?'disabled':''}>Add debt</button><button class="button" data-credit-action="entry" ${busy?'disabled':''}>Post payment / charges</button><label><input id="creditArchived" type="checkbox" ${includeArchived?'checked':''}>Show archived</label></div>
      <div class="credit-metrics">${totals.map(([label,value])=>`<article class="panel"><small>${label}</small><strong>${usd(value)}</strong></article>`).join('')}</div>
      <article class="panel"><form id="creditBudgetForm" class="credit-toolbar"><label>Monthly card payoff budget (USD)<input name="budget" type="number" required min="0" step="0.01" value="${data.settings.budget}" ${busy?'disabled':''}></label><button class="button secondary" ${busy?'disabled':''}>Save budget</button></form><p>The budget and projection apply only to cards included in the payoff plan. Other debts are tracked separately, following the workbook. Statement minimums refer to current due dates, including past due amounts. Payment totals refer to ${esc(month)}. Archived debts remain in the balance and payment history.</p></article>
      <article class="panel"><h2>Debts &amp; current statements</h2>${table(['Debt','Balance','APR','Due','Minimum','Paid for statement','Still required','Status','Actions'],rows.map(d=>`<tr><th scope="row">${esc(d.name)}${d.archived?' (archived)':''}</th><td>${usd(d.balance)}</td><td>${(d.apr*100).toFixed(2)}%</td><td>${esc(d.due_date||'Not entered')}</td><td>${usd(d.minimum)}</td><td>${usd(d.invoicePaid)}</td><td>${usd(d.reserve)}</td><td>${esc(d.status)}</td><td><button class="button secondary" data-credit-edit="${d.id}" ${busy?'disabled':''}>Edit</button> <button class="button secondary" data-credit-archive="${d.id}" ${busy?'disabled':''}>${d.archived?'Restore':'Archive'}</button></td></tr>`))}</article>
      <article class="panel"><h2>Monthly activity</h2>${table(['Debt','Payments','Actual interest','Purchases & fees','Net reduction'],rows.map(d=>`<tr><th scope="row">${esc(d.name)}</th><td>${usd(d.paid)}</td><td>${usd(d.interest)}</td><td>${usd(d.charges)}</td><td>${usd(d.reduction)}</td></tr>`))}</article>
      <article class="panel"><h2>Promotional balances</h2><p>Update the promotional balance and installments from the lender statement. These amounts are tracked separately from the standard APR projection.</p>${table(['Debt','Remaining promotional balance','Deadline','Installments remaining','Target per installment'],s.active.filter(d=>d.promo_balance!=null).map(d=>`<tr><th scope="row">${esc(d.name)}</th><td>${usd(d.promo_balance)}</td><td>${esc(d.promo_deadline||'Not entered')}</td><td>${esc(d.promo_installments||'Not entered')}</td><td>${d.promo_installments?usd(Math.ceil(M.cents(d.promo_balance)/d.promo_installments)/100):'Not entered'}</td></tr>`))}</article>
      <article class="panel"><h2>Card payoff projection</h2><p>Estimates use the selected card plan, constant APR, no new purchases or fees, and extra payments toward the highest APR. Other financing and deferred-interest promotions are tracked separately.</p>
      ${dueMissing?'<p>Enter current minimums and due dates for a statement-based projection. A blank card minimum uses the workbook estimate: the largest of $35, 2% of the opening monthly balance, or estimated interest.</p>':''}
      ${projection.blockedMonth?`<p>Projection stops in month ${projection.blockedMonth}: ${usd(projection.shortfall)} short of minimum payments.</p>`:`<p>${projection.paidOff?`Cards paid off in ${projection.rows.length} months.`:'Card balance remains after 120 months.'}</p>`}
      ${table(['Month','Estimated interest','Card payment','Card balance'],projection.rows.map(r=>`<tr><td>${r.month}</td><td>${usd(r.interest)}</td><td>${usd(r.payment)}</td><td>${usd(r.balance)}</td></tr>`))}</article>
      <article class="panel"><h2>Ledger · ${esc(month)}</h2><p>Void an incorrect entry and post a corrected one. Voided entries remain in history.</p>${table(['Posting date','Debt','Payment','Interest','Purchases / fees','Statement due','Notes','Action'],data.entries.filter(e=>e.posted_on.slice(0,7)===month).map(e=>`<tr class="${e.voided_at?'credit-voided':''}"><td>${esc(e.posted_on)}</td><td>${esc(data.debts.find(d=>d.id===e.debt_id)?.name||'Unknown debt')}</td><td>${usd(e.payment)}</td><td>${usd(e.interest)}</td><td>${usd(e.charges)}</td><td>${esc(e.invoice_due||'Not linked')}</td><td>${esc(e.notes)}</td><td>${e.voided_at?'Voided':`<button class="button secondary" data-credit-void="${e.id}" ${busy?'disabled':''}>Void</button>`}</td></tr>`))}</article>
      ${data.debts.length?'':`<article class="panel"><h2>Import the existing control</h2><p>Select the private import JSON prepared from your workbook. Review it before confirming. The import is available only while this module is empty.</p><label>Private import file<input id="creditImport" type="file" accept=".json,application/json"></label><p id="creditImportPreview" role="status">${preview?`${preview.debts.length} debts and ${preview.entries.length} ledger entries · Monthly budget ${usd(preview.budget)}`:''}</p>${preview?'<button class="button" data-credit-action="import">Confirm initial import</button>':''}</article>`}</section>`;
  }
  function bind(root){
    root.querySelectorAll('[data-credit-action]').forEach(b=>b.addEventListener('click',async()=>{
      if(busy)return;
      const action=b.dataset.creditAction;
      if(action==='refresh'){await load();return;}
      if(action==='add'){screen='debt';editing=null;}
      if(action==='entry'){screen='entry';entryDraft={request_id:crypto.randomUUID()};}
      if(action==='cancel'){screen='summary';editing=null;error='';}
      if(action==='import'&&preview){
        if(!confirm(`Import ${preview.debts.length} debts and ${preview.entries.length} entries into the private credit module?`))return;
        const request=generation;busy=true;rerender();
        try{const {error:e}=await client.rpc('nl_credit_import',{document:preview});if(e)throw e;if(request!==generation||!allowed())return;preview=null;notice='Workbook imported.';await load();}
        catch(e){if(request===generation)error=e.message;}
        finally{if(request===generation){busy=false;rerender();}}return;
      }
      rerender();
    }));
    root.querySelector('#creditMonth')?.addEventListener('change',e=>{if(/^\d{4}-\d{2}$/.test(e.target.value)){month=e.target.value;rerender();}});
    root.querySelector('#creditArchived')?.addEventListener('change',e=>{includeArchived=e.target.checked;rerender();});
    root.querySelectorAll('[data-credit-edit]').forEach(b=>b.addEventListener('click',()=>{editing=data.debts.find(d=>d.id===b.dataset.creditEdit);screen='debt';error='';rerender();}));
    root.querySelectorAll('[data-credit-archive]').forEach(b=>b.addEventListener('click',()=>{const d=data.debts.find(d=>d.id===b.dataset.creditArchive);void write('archive',{id:d.id,revision:d.revision});}));
    root.querySelectorAll('[data-credit-void]').forEach(b=>b.addEventListener('click',()=>{if(confirm('Void this entry? Its amounts will be removed from totals and retained in history.'))void write('void',{id:b.dataset.creditVoid});}));
    root.querySelector('#creditDebtForm')?.addEventListener('submit',e=>{e.preventDefault();const item=Object.fromEntries(new FormData(e.target));item.apr=Number(item.apr)/100;item.in_plan=item.in_plan==='on';if(editing?.id){item.id=editing.id;item.revision=editing.revision;}editing=item;if(item.in_plan&&item.kind!=='card'){error='Only credit cards can join the card payoff plan.';rerender();return;}void write('debt',item);});
    root.querySelector('#creditEntryForm')?.addEventListener('submit',e=>{e.preventDefault();entryDraft={...Object.fromEntries(new FormData(e.target)),request_id:entryDraft.request_id||crypto.randomUUID()};void write('entry',entryDraft);});
    root.querySelector('#creditBudgetForm')?.addEventListener('submit',e=>{e.preventDefault();void write('budget',{budget:new FormData(e.target).get('budget'),revision:data.settings.revision});});
    root.querySelector('#creditImport')?.addEventListener('change',async e=>{
      const file=e.target.files[0];if(!file)return;const request=generation;
      try{if(file.size>5e6)throw Error('Import file must be under 5 MB.');const doc=JSON.parse(await file.text());if(doc.format!=='nolimit-credit-v1'||!Array.isArray(doc.debts)||!Array.isArray(doc.entries)||!Number.isFinite(doc.budget)||doc.budget<0)throw Error('Choose the prepared credit import JSON.');if(request===generation&&allowed()){preview=doc;error='';rerender();}}
      catch(e){if(request===generation){preview=null;error=e.message;rerender();}}
    });
    if(!data&&!loading&&!error)void load();
  }
  return {configure,reset,render,bind};
})();
