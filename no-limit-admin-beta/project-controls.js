/* Project providers, daily work and cash ledgers. All monetary calculations use cents. */
(() => {
  const list=(s,k)=>Array.isArray(s[k])?s[k]:[];
  const cents=v=>Math.round((Number(v)||0)*100);
  const dollars=v=>Math.round(v)/100;
  const today=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'America/New_York',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
  function day(v){
    const text=String(v||'').trim();
    if(/^\d{4}-\d{2}-\d{2}$/.test(text))return Number.isFinite(Date.parse(text))&&new Date(text).toISOString().slice(0,10)===text?text:'';
    if(!text)return '';const d=new Date(text);if(!Number.isFinite(d.getTime()))return '';
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }
  function bounds(value=today()){
    const key=day(value);if(!key)throw new Error('Choose a valid week date.');const d=new Date(key+'T12:00:00Z');
    d.setUTCDate(d.getUTCDate()-((d.getUTCDay()+6)%7));const start=d.toISOString().slice(0,10);d.setUTCDate(d.getUTCDate()+6);return {start,end:d.toISOString().slice(0,10)};
  }
  function number(v,label,{min=0,max=1e9,integer=false}={}){
    if(v===null||v===undefined||String(v).trim()===''||!Number.isFinite(Number(v))||Number(v)<min||Number(v)>max||(integer&&!Number.isInteger(Number(v))))throw new Error(`Enter a valid ${label}.`);
    return Number(v);
  }
  function date(v,label){const key=day(v);if(!key)throw new Error(`Enter a valid ${label}.`);return key;}
  const budget=e=>e.basis==='daily'?Math.round(cents(e.dailyRate)*Number(e.plannedDays||0)):cents(e.contractAmount);
  const activePay=p=>p.status==='paid';
  const workValue=(row,e)=>e.basis==='daily'?Math.round(cents(row.dailyRate)*Number(row.units||0)):cents(row.workValue);
  function charges(s,e){
    if(e.basis==='daily')return list(s,'schedule').filter(r=>r.engagementId===e.id&&r.status==='completed').map(r=>({id:`day:${r.id}`,engagementId:e.id,label:`Daily work ${day(r.workDate)}`,date:day(r.dueDate)||day(r.workDate),amount:workValue(r,e)})).filter(r=>r.date);
    return list(s,'providerMilestones').filter(r=>r.engagementId===e.id&&r.status!=='void').map(r=>({id:`milestone:${r.id}`,engagementId:e.id,label:r.label,date:day(r.dueDate),amount:cents(r.amount)})).filter(r=>r.date);
  }
  function engagementSummary(s,e,asOf=today()){
    const project=list(s,'projects').find(p=>p.id===e.projectId)||{};
    const progress=Math.min(100,Math.max(0,Number(e.progressMode==='scope'?e.scopeProgress:project.progress)||0));
    const payments=list(s,'providerPayments').filter(p=>p.engagementId===e.id&&activePay(p)&&day(p.date)&&day(p.date)<=asOf);
    const paid=payments.reduce((n,p)=>n+cents(p.amount),0), total=budget(e);
    const assignments=list(s,'schedule').filter(r=>r.engagementId===e.id&&r.status!=='cancelled');
    const completed=assignments.filter(r=>r.status==='completed'&&day(r.workDate)<=asOf);
    const recorded=completed.reduce((n,r)=>n+workValue(r,e),0);
    const earned=e.basis==='daily'?recorded:Math.round(total*progress/100);
    const lines=charges(s,e).sort((a,b)=>a.date.localeCompare(b.date)||a.id.localeCompare(b.id)).map(c=>({...c,allocated:0}));
    let pool=0;
    for(const p of payments){const line=lines.find(c=>c.id===p.chargeId);const amount=cents(p.amount);if(line){const used=Math.min(Math.max(0,line.amount-line.allocated),amount);line.allocated+=used;pool+=amount-used;}else pool+=amount;}
    for(const line of lines){const used=Math.min(Math.max(0,line.amount-line.allocated),pool);line.allocated+=used;pool-=used;line.unpaid=Math.max(0,line.amount-line.allocated);}
    const due=lines.filter(c=>c.date<=asOf).reduce((n,c)=>n+c.unpaid,0);
    const overdue=lines.filter(c=>c.date<asOf).reduce((n,c)=>n+c.unpaid,0);
    const benchmark=Math.round(total*progress/100),ahead=Math.max(0,paid-benchmark),unearned=Math.max(0,paid-earned),allowance=cents(e.advanceAllowance);
    const unapproved=Math.max(0,Math.max(ahead,unearned)-allowance);
    return {id:e.id,budget:total,progress,paid,earned,recorded,benchmark,ahead,unearned,allowance,unapproved,due,overdue,lines,advanceBalance:pool,committedUnpaid:Math.max(0,total-paid),plannedWork:assignments.reduce((n,r)=>n+workValue(r,e),0),paymentTermsMissing:e.basis==='fixed'&&!lines.length};
  }
  function clientCash(s,projectId,start='',end='9999-12-31'){
    const invoices=list(s,'invoices').filter(i=>i.projectId===projectId&&!['void','cancelled'].includes(i.status));
    const ids=new Set(invoices.map(i=>i.id));
    const payments=list(s,'payments').filter(p=>(p.projectId===projectId||ids.has(p.invoiceId))&&['received','cleared'].includes(p.status));
    const valid=payments.filter(p=>day(p.date)&&day(p.date)<=end);
    const ledger=valid.filter(p=>!start||day(p.date)>=start).map(p=>({id:p.id,date:day(p.date),amount:cents(p.amount),reference:p.checkNumber||p.milestone||'',method:p.method||'',projectId}));
    if(invoices.length){
      const linked=payments.filter(p=>ids.has(p.invoiceId)),standalone=payments.filter(p=>!ids.has(p.invoiceId));
      const legacy=list(s,'transactions').filter(t=>t.projectId===projectId&&t.type==='receivable'&&t.status==='paid'&&!payments.some(p=>p.legacyTransactionId===t.id));
      const recorded=invoices.reduce((n,i)=>n+cents(i.paid),0);
      // Older projects can have received cash before an invoice existed. Preserve
      // that historical balance, but never add the duplicated ledger twice.
      const historical=Math.max(recorded,legacy.filter(t=>!day(t.date)||day(t.date)<=end).reduce((n,t)=>n+cents(t.amount),0),linked.filter(p=>day(p.date)&&day(p.date)<=end).reduce((n,p)=>n+cents(p.amount),0));
      const historyRows=linked.length?linked:legacy;
      const historyDated=historyRows.filter(p=>day(p.date)).map(p=>({id:p.id,date:day(p.date),amount:cents(p.amount),reference:p.checkNumber||p.milestone||p.category||'',method:p.method||'Legacy receipt',projectId}));
      const standaloneDated=standalone.filter(p=>day(p.date)).map(p=>({id:p.id,date:day(p.date),amount:cents(p.amount),reference:p.checkNumber||p.milestone||'',method:p.method||'',projectId}));
      const rows=[...historyDated,...standaloneDated].filter(p=>p.date<=end&&(!start||p.date>=start));
      return {total:start?rows.reduce((n,p)=>n+p.amount,0):historical+standaloneDated.filter(p=>p.date<=end).reduce((n,p)=>n+p.amount,0),rows,undated:Math.max(0,historical-historyDated.reduce((n,p)=>n+p.amount,0))+standalone.filter(p=>!day(p.date)).reduce((n,p)=>n+cents(p.amount),0),basis:'reconciled invoice / legacy balance and separate project receipts'};
    }
    const tx=list(s,'transactions').filter(t=>t.projectId===projectId&&t.type==='receivable'&&t.status==='paid'&&!payments.some(p=>p.legacyTransactionId===t.id));
    // Older undocumented payment ledgers remain authoritative. New standalone
    // receipts are additive to the existing legacy balance, never a replacement.
    if(payments.some(p=>p.source!=='client-receipt'))return {total:ledger.reduce((n,p)=>n+p.amount,0),rows:ledger,undated:payments.filter(p=>!day(p.date)).reduce((n,p)=>n+cents(p.amount),0),basis:'client payment ledger'};
    const rows=tx.filter(t=>day(t.date)&&day(t.date)<=end&&(!start||day(t.date)>=start)).map(t=>({id:t.id,date:day(t.date),amount:cents(t.amount),reference:t.category||'',method:'Legacy receipt',projectId}));
    return {total:ledger.reduce((n,p)=>n+p.amount,0)+(start?rows.reduce((n,p)=>n+p.amount,0):tx.filter(t=>!day(t.date)||day(t.date)<=end).reduce((n,p)=>n+cents(p.amount),0)),rows:[...rows,...ledger],undated:tx.filter(t=>!day(t.date)).reduce((n,t)=>n+cents(t.amount),0)+payments.filter(p=>!day(p.date)).reduce((n,p)=>n+cents(p.amount),0),basis:'dated client receipts and legacy received transactions'};
  }
  function projectSummary(s,p,asOf=today()){
    const agreements=list(s,'projectEngagements').filter(e=>e.projectId===p.id);
    const providers=agreements.map(e=>({...engagementSummary(s,e,asOf),agreement:e}));
    const paid=providers.reduce((n,e)=>n+e.paid,0),due=providers.reduce((n,e)=>n+e.due,0),earned=providers.reduce((n,e)=>n+e.earned,0);
    const received=clientCash(s,p.id,'',asOf),clientEarned=Math.round(cents(p.contractValue)*Number(p.progress||0)/100);
    const alerts=[];
    for(const summary of providers){const e=summary.agreement,name=list(s,'people').find(x=>x.id===e.personId)?.name||e.personName||'Provider';
      if(summary.paid>summary.budget||summary.recorded>summary.budget)alerts.push({level:'critical',code:'provider-budget',text:`${name}: recorded payments or daily work exceed the registered provider budget.`,amount:Math.max(summary.paid,summary.recorded)-summary.budget});
      if(summary.unapproved>1)alerts.push({level:'warning',code:'provider-advance',text:`${name}: payment is ahead of recorded work / physical progress (${summary.progress}%).`,amount:summary.unapproved});
      else if(Math.max(summary.ahead,summary.unearned)>1)alerts.push({level:'info',code:'approved-advance',text:`${name}: advance covered by the recorded allowance — ${e.advanceReason||'review the agreement'}.`,amount:Math.max(summary.ahead,summary.unearned)});
      if(summary.overdue>0)alerts.push({level:'warning',code:'provider-overdue',text:`${name}: unpaid agreed amount is past its due date.`,amount:summary.overdue});
      if(summary.paymentTermsMissing)alerts.push({level:'info',code:'missing-terms',text:`${name}: contract payment milestones have not been recorded.`,amount:null});
    }
    if(cents(p.contractValue)<=0)alerts.push({level:'info',code:'missing-client-contract',text:'Register the client contract value before comparing receipts with physical progress.',amount:null});
    else if(clientEarned-received.total>1)alerts.push({level:'warning',code:'client-progress',text:`Client receipts are below the value of executed work (${Number(p.progress||0)}%). This is an exposure indicator, not an overdue invoice by itself.`,amount:clientEarned-received.total});
    if(paid>received.total)alerts.push({level:'warning',code:'cash-gap',text:'Provider payments exceed recorded client receipts (provider cash flow only).',amount:paid-received.total});
    if(received.undated>0)alerts.push({level:'info',code:'undated-receipts',text:'Historical client receipts have no matching dated ledger entries and cannot be allocated to a weekly report.',amount:received.undated});
    const legacy=list(s,'schedule').filter(r=>r.projectId===p.id&&!r.engagementId&&r.status!=='cancelled');
    if(legacy.length)alerts.push({level:'info',code:'unlinked-work',text:`${legacy.length} older assignment(s) have no provider agreement / rate. Review them; costs were not invented.`,amount:null});
    return {providers,paid,due,earned,received:received.total,clientEarned,alerts,committed:providers.reduce((n,e)=>n+e.budget,0),committedUnpaid:providers.reduce((n,e)=>n+e.committedUnpaid,0),undated:received.undated};
  }
  function weekly(s,week=today(),projectId='all'){
    const range=bounds(week),within=v=>day(v)>=range.start&&day(v)<=range.end;
    return { ...range,projects:list(s,'projects').filter(p=>projectId==='all'||p.id===projectId).map(p=>{
      const agreements=list(s,'projectEngagements').filter(e=>e.projectId===p.id),ids=new Set(agreements.map(e=>e.id));
      const payments=list(s,'providerPayments').filter(r=>ids.has(r.engagementId)&&activePay(r)&&within(r.date));
      const receipts=clientCash(s,p.id,range.start,range.end);
      const work=list(s,'schedule').filter(r=>r.projectId===p.id&&r.status==='completed'&&within(r.workDate)).map(r=>({row:r,agreement:agreements.find(e=>e.id===r.engagementId)})).filter(x=>x.agreement);
      const paid=payments.reduce((n,r)=>n+cents(r.amount),0),summary=projectSummary(s,p,range.end);
      return {project:p,received:receipts.total,paid,net:receipts.total-paid,recordedWork:work.reduce((n,x)=>n+workValue(x.row,x.agreement),0),due:summary.due,receipts:receipts.rows,payments,work,providers:agreements.map(e=>({agreement:e,paid:payments.filter(r=>r.engagementId===e.id).reduce((n,r)=>n+cents(r.amount),0),work:work.filter(x=>x.agreement.id===e.id).reduce((n,x)=>n+workValue(x.row,e),0),due:engagementSummary(s,e,range.end).due})),undated:summary.undated};
    })};
  }
  function apply(state,type,input,id='',actor='',asOf=today()){
    const s=structuredClone(state);for(const k of ['projectEngagements','providerPayments','providerMilestones','schedule'])if(!Array.isArray(s[k]))s[k]=[];
    const get=k=>input instanceof FormData?input.get(k):input[k];
    const text=(k,max=2000)=>String(get(k)||'').trim().slice(0,max);
    const uid=prefix=>`${prefix}-${crypto.randomUUID().slice(0,8).toUpperCase()}`;
    const save=(collection,record)=>{const i=s[collection].findIndex(r=>r.id===record.id);if(i<0)s[collection].unshift(record);else s[collection][i]=record;};
    let record;
    if(type==='engagement'){
      const old=s.projectEngagements.find(e=>e.id===id),project=list(s,'projects').find(p=>p.id===get('projectId')),person=list(s,'people').find(p=>p.id===get('personId'));
      if(!project||!person)throw new Error('Select a registered project and person / company.');
      if(old&&(old.projectId!==project.id||old.personId!==person.id||old.basis!==get('basis')))throw new Error('The provider, project and agreement basis cannot change after creation. Create a separate agreement.');
      if(!['daily','fixed'].includes(get('basis')))throw new Error('Choose daily or fixed-contract terms.');
      if(!text('scope'))throw new Error('Describe the provider scope.');
      const basis=get('basis'),dailyRate=basis==='daily'?number(get('dailyRate'),'daily rate'):0,plannedDays=basis==='daily'?number(get('plannedDays'),'planned days',{min:.01,max:100000}):0,contractAmount=basis==='fixed'?number(get('contractAmount'),'contract value'):0;
      const advanceAllowance=text('advanceAllowance')?number(get('advanceAllowance'),'advance allowance'):0;
      if(advanceAllowance>0&&!text('advanceReason'))throw new Error('Record a reason for the approved advance allowance.');
      const progressMode=get('progressMode');if(!['project','scope'].includes(progressMode))throw new Error('Choose the progress reference.');
      const scopeProgress=progressMode==='scope'?number(get('scopeProgress'),'scope progress',{max:100}):0;
      if(!['active','paused','completed'].includes(get('status')))throw new Error('Choose a valid agreement status.');
      record={...old,id:id||uid('ENG'),projectId:project.id,personId:person.id,personName:person.name,scope:text('scope'),basis,dailyRate,plannedDays,contractAmount,progressMode,scopeProgress,advanceAllowance,advanceReason:text('advanceReason'),status:get('status'),updatedAt:new Date().toISOString(),updatedBy:actor};
      const milestoneTotal=s.providerMilestones.filter(m=>m.engagementId===record.id&&m.status!=='void').reduce((n,m)=>n+cents(m.amount),0);
      if(milestoneTotal>budget(record))throw new Error('Contract value cannot be lower than the active payment milestones.');
      save('projectEngagements',record);person.projectIds=[...new Set([...(person.projectIds||[]),project.id])];
    }else{
      const originalPayment=type==='providerPayment'?s.providerPayments.find(r=>r.id===id):null;
      if(originalPayment&&originalPayment.engagementId!==get('engagementId'))throw new Error('The payment agreement cannot change.');
      const e=s.projectEngagements.find(r=>r.id===get('engagementId'));if(!e)throw new Error('Link the provider and terms to a project first.');
      const person=list(s,'people').find(p=>p.id===e.personId),project=list(s,'projects').find(p=>p.id===e.projectId);if(!person||!project)throw new Error('The linked provider or project is unavailable.');
      if(type==='schedule'){
        const old=s.schedule.find(r=>r.id===id);if(old?.engagementId&&old.engagementId!==e.id)throw new Error('Create a new assignment to change the provider agreement.');
        if(e.status!=='active'&&!old)throw new Error('Only active agreements can receive new work assignments.');
        const workDate=date(get('workDate'),'work date'),status=get('status');
        if(!['scheduled','confirmed','completed','cancelled'].includes(status))throw new Error('Choose a valid work status.');
        if(status==='completed'&&workDate>asOf)throw new Error('Future work cannot be marked completed.');
        const units=e.basis==='daily'?number(get('units'),'billable days / units',{min:.01,max:1000}):1;
        const dailyRate=e.basis==='daily'?(old?.dailyRate??e.dailyRate):0;
        const workAmount=e.basis==='fixed'?(text('workValue')?number(get('workValue'),'daily contract work value'):status==='completed'?number('', 'daily contract work value'):0):0;
        if(!text('shift')||!text('instructions'))throw new Error('Enter the shift and work instructions.');
        record={...old,id:id||uid('SCH'),engagementId:e.id,projectId:e.projectId,projectName:project.name,personId:e.personId,personName:person.name,personType:person.type,workDate,shift:text('shift',200),instructions:text('instructions'),status,units,dailyRate,workValue:workAmount,dueDate:e.basis==='daily'?date(get('dueDate')||workDate,'payment due date'):'',updatedAt:new Date().toISOString(),updatedBy:actor};
        const paidToDay=s.providerPayments.filter(p=>p.chargeId===`day:${record.id}`&&activePay(p)).reduce((n,p)=>n+cents(p.amount),0);
        if(paidToDay>0&&(status!=='completed'||workValue(record,e)<paidToDay))throw new Error('This daily charge has payments. Keep its completed value or reverse the payment first.');
        save('schedule',record);person.projectIds=[...new Set([...(person.projectIds||[]),project.id])];
      }else if(type==='providerMilestone'){
        if(e.basis!=='fixed')throw new Error('Contract milestones are only for fixed-contract agreements.');
        const old=s.providerMilestones.find(r=>r.id===id);if(old&&old.engagementId!==e.id)throw new Error('The agreement cannot change.');
        const amount=number(get('amount'),'milestone amount',{min:.01}),dueDate=date(get('dueDate'),'milestone due date'),status=get('status')||'active';
        if(!['active','void'].includes(status)||!text('label'))throw new Error('Enter the milestone label and status.');
        const allocated=s.providerPayments.filter(p=>p.chargeId===`milestone:${id}`&&activePay(p)).reduce((n,p)=>n+cents(p.amount),0);
        if(allocated>0&&(status==='void'||cents(amount)<allocated))throw new Error('This milestone has payments. Reverse them before voiding or reducing it.');
        const total=s.providerMilestones.filter(m=>m.engagementId===e.id&&m.id!==id&&m.status!=='void').reduce((n,m)=>n+cents(m.amount),0)+(status==='void'?0:cents(amount));
        if(total>budget(e))throw new Error('Active milestones cannot exceed the provider contract value.');
        record={...old,id:id||uid('MIL'),engagementId:e.id,projectId:e.projectId,personId:e.personId,label:text('label',300),amount,dueDate,status,updatedAt:new Date().toISOString(),updatedBy:actor};save('providerMilestones',record);
      }else if(type==='providerPayment'){
        const old=s.providerPayments.find(r=>r.id===id);
        if(old){if(old.status==='void')throw new Error('This payment is already void.');if(get('status')!=='void'||!text('notes'))throw new Error('Recorded payments are immutable. Void with a reason and create a corrected payment.');record={...old,status:'void',voidReason:text('notes'),voidedAt:new Date().toISOString(),voidedBy:actor};}
        else{
          const amount=number(get('amount'),'payment amount',{min:.01}),paymentDate=date(get('date'),'payment date');if(paymentDate>asOf)throw new Error('Record only payments already made; use work dates or contract milestones for future payables.');
          const chargeId=text('chargeId',150);if(chargeId&&!charges(s,e).some(c=>c.id===chargeId))throw new Error('Select a charge from the chosen provider agreement.');
          if(!text('method')||!text('reference'))throw new Error('Enter the payment method and a check / transfer / cash receipt reference.');
          record={id:uid('PVP'),engagementId:e.id,projectId:e.projectId,personId:e.personId,date:paymentDate,amount,method:text('method',100),reference:text('reference',300),chargeId,status:'paid',notes:text('notes'),createdAt:new Date().toISOString(),createdBy:actor};
        }
        save('providerPayments',record);
      }else throw new Error('Unknown project control entry.');
    }
    return {state:s,record};
  }
  function csv(report){
    const rows=[['Period',report.start,report.end],['Project','Client received','Providers paid','Provider-only net cash','Daily work value recorded','Provider due (week end)']];
    for(const p of report.projects)rows.push([p.project.name,dollars(p.received).toFixed(2),dollars(p.paid).toFixed(2),dollars(p.net).toFixed(2),dollars(p.recordedWork).toFixed(2),dollars(p.due).toFixed(2)]);
    rows.push([],['Project','Provider / agreement','Daily work recorded','Cash paid','Due (week end)']);
    for(const p of report.projects)for(const e of p.providers)rows.push([p.project.name,e.agreement.personName+' / '+e.agreement.scope,dollars(e.work).toFixed(2),dollars(e.paid).toFixed(2),dollars(e.due).toFixed(2)]);
    rows.push([],['Direction','Project','Date','Party / agreement','Amount','Method','Reference']);
    for(const p of report.projects){for(const receipt of p.receipts)rows.push(['Client receipt',p.project.name,receipt.date,'Client',dollars(receipt.amount).toFixed(2),receipt.method,receipt.reference]);for(const payment of p.payments)rows.push(['Provider payment',p.project.name,day(payment.date),p.providers.find(e=>e.agreement.id===payment.engagementId)?.agreement.personName||payment.personId,Number(payment.amount).toFixed(2),payment.method,payment.reference]);}
    return '\uFEFF'+rows.map(row=>row.map(v=>{let t=String(v??'');if(/^[\s]*[=+@-]/.test(t)&&!/^-[0-9]+(?:\.[0-9]+)?$/.test(t))t="'"+t;return '"'+t.replace(/"/g,'""')+'"';}).join(',')).join('\r\n');
  }
  window.NoLimitProjectControls={list,cents,dollars,today,day,bounds,budget,charges,workValue,engagementSummary,clientCash,projectSummary,weekly,apply,csv};
})();
