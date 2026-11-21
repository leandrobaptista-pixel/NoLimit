(function(root){
  'use strict';
  const cents=v=>Math.round(Number(v)*100);
  const money=n=>n/100;
  const today=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'America/New_York',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
  function summary(data,month,asOf=today()){
    const entries=data.entries.filter(e=>!e.voided_at&&e.posted_on<=asOf);
    const rows=data.debts.map(d=>{
      const ledger=entries.filter(e=>e.debt_id===d.id);
      const monthly=ledger.filter(e=>e.posted_on.slice(0,7)===month);
      const sum=(list,key)=>list.reduce((n,e)=>n+cents(e[key]||0),0);
      const balance=cents(d.opening_balance)+sum(ledger,'interest')+sum(ledger,'charges')-sum(ledger,'payment');
      const invoicePaid=d.due_date?sum(ledger.filter(e=>e.invoice_due===d.due_date),'payment'):0;
      const reserve=balance<=0?0:d.minimum==null||!d.due_date?null:Math.max(0,cents(d.minimum)-invoicePaid);
      return {...d,balance:money(balance),paid:money(sum(monthly,'payment')),interest:money(sum(monthly,'interest')),charges:money(sum(monthly,'charges')),reduction:money(sum(monthly,'payment')-sum(monthly,'interest')-sum(monthly,'charges')),invoicePaid:money(invoicePaid),reserve:reserve==null?null:money(reserve),status:balance<=0?'Paid off':reserve==null?'Enter statement':reserve===0?'Minimum covered':d.due_date<asOf?'Past due':'Minimum pending'};
    });
    const active=rows.filter(d=>!d.archived);
    const reserve=active.some(d=>d.balance>0&&d.reserve==null)?null:money(active.reduce((n,d)=>n+(d.balance>0?cents(d.reserve||0):0),0));
    const paid=money(rows.reduce((n,d)=>n+cents(d.paid),0));
    const plan=active.filter(d=>d.in_plan);
    const planPaid=money(rows.filter(d=>d.in_plan).reduce((n,d)=>n+cents(d.paid),0));
    const planReserve=plan.some(d=>d.balance>0&&d.reserve==null)?null:money(plan.reduce((n,d)=>n+cents(d.reserve||0),0));
    const available=money(cents(data.settings.budget)-cents(planPaid));
    return {rows,active,reserve,paid,planReserve,available,shortfall:planReserve==null?null:money(Math.max(0,cents(planReserve)-cents(available))),balance:money(rows.reduce((n,d)=>n+cents(d.balance),0))};
  }
  function project(debts,budget,maxMonths=120){
    let balances=debts.map(d=>Math.max(0,cents(d.balance)));
    const order=debts.map((d,i)=>i).sort((a,b)=>Number(debts[b].apr)-Number(debts[a].apr)||a-b);
    const rows=[];const monthlyBudget=cents(budget);
    for(let month=1;month<=maxMonths&&balances.some(b=>b>0);month++){
      const interest=balances.map((b,i)=>Math.round(b*Number(debts[i].apr)/12));
      const owed=balances.map((b,i)=>b+interest[i]);
      const minimum=balances.map((b,i)=>b===0?0:Math.min(owed[i],debts[i].minimum==null?Math.max(3500,Math.round(b*0.02),interest[i]):cents(debts[i].minimum)));
      const totalMinimum=minimum.reduce((a,b)=>a+b,0);
      if(totalMinimum>monthlyBudget)return {rows,blockedMonth:month,shortfall:money(totalMinimum-monthlyBudget),paidOff:false};
      let remaining=monthlyBudget-totalMinimum;
      const payments=minimum.slice();
      for(const i of order){const extra=Math.min(remaining,owed[i]-payments[i]);payments[i]+=extra;remaining-=extra;}
      balances=owed.map((b,i)=>b-payments[i]);
      rows.push({month,interest:money(interest.reduce((a,b)=>a+b,0)),payment:money(payments.reduce((a,b)=>a+b,0)),balance:money(balances.reduce((a,b)=>a+b,0))});
    }
    return {rows,paidOff:balances.every(b=>b===0),blockedMonth:null,shortfall:0};
  }
  root.NoLimitCreditMath={cents,today,summary,project};
})(typeof window==='undefined'?globalThis:window);
