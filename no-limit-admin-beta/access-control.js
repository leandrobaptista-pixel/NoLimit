/* All decisions are enforced again by PostgreSQL / protected Edge Functions. */
const NoLimitAccess = (() => {
  const areas = [
    ['overview','Overview','Summary of permitted information.'],
    ['requests','Visit requests','Website leads, attachments and internal review. All-project access is required for website intake.'],
    ['documents','Estimates, contracts & invoices','Includes document prices, terms, change orders and takeoff attachments.'],
    ['clients','Clients','Contacts, billing addresses and client records.'],
    ['projects','Projects','Project details and progress. Financial fields require Financial permission.'],
    ['financial','Financial','Costs, rates, receipts, payments and provider agreements.'],
    ['services','Service catalog','Reusable scope and selling prices.'],
    ['materials','Materials & purchasing','Materials, vendors, purchase quantities and costs.'],
    ['team','Employees & subcontractors','People, contact details and project assignments.'],
    ['vendors','Vendors','Vendor directory and project assignments.'],
    ['schedule','Schedule & attendance','Work dates, instructions and attendance. Rates require Financial permission.'],
    ['media','Project media','Private project photos and files. Publishing the website is Owner only.'],
    ['map','Project locations','Addresses for permitted projects.'],
    ['compliance','Compliance & consents','Requirements, W-9, insurance documents and renewal records.'],
    ['reports','Reports','Build and export reports from information this group can view.']
  ];
  const names={admin:'Administration',office:'Finance',manager:'Project managers',team_member:'Employees',vendor:'Vendors',subcontractor:'Subcontractors',viewer:'Read only',developer:'Developer',owner:'Owner'};
  const readonly = new Set(['overview','map','reports']);
  let selected='admin', baseline={}, policyDraft=null, dirty=false, busy=false, status='';
  const esc=v=>escapeHtml(String(v??''));
  const context=()=>currentBetaUser?.access;
  const isOwner=()=>context()?.owner===true;
  const can=(area,action='view')=>isOwner() || context()?.permissions?.[area]?.[action]===true;
  function areaFor(type,id='') {
    if(type==='person')return (state.people.find(p=>p.id===id)?.type || (location.hash==='#vendors'?'Vendor':''))==='Vendor'?'vendors':'team';
    return ({accessUser:'security',request:'requests',siteVisit:'requests',estimate:'documents',contract:'documents',invoice:'documents',invoiceItem:'documents',changeOrder:'documents',client:'clients',project:'projects',projectVendor:'vendors',receipt:'financial',payment:'financial',transaction:'financial',engagement:'financial',providerPayment:'financial',providerMilestone:'financial',catalogService:'services',material:'materials',schedule:'schedule',attendance:'schedule',media:'media',compliance:'compliance',consent:'compliance'})[type] || type;
  }
  const allows=(type,action='create',id='')=>type==='providerPayment'&&id?can('financial','delete')&&can('financial','edit'):can(areaFor(type,id),['projectVendor','invoiceItem'].includes(type)&&action==='create'?'edit':action);
  const routes=()=>Object.keys(window.NoLimitRoutes || {}).filter(r=>['security','website-insights'].includes(r)?isOwner():can(r));
  const snapshot=(candidate=state)=>{baseline=structuredClone(candidate);};
  function sanitize(collection,record){
    const item=structuredClone(record);
    if(isOwner())return item;
    if(!can('financial')){
      const fields=collection==='projects'?['contractValue','cost','outstanding']:collection==='people'?['dailyRate','hourlyRate','rate','balance','totalPaid','payRate']:[];
      fields.forEach(k=>delete item[k]);
      if(collection==='schedule')for(const k of Object.keys(item))if(!['id','shift','status','personId','workDate','projectId','personName','personType','projectName','instructions','engagementId','createdAt','updatedAt','updatedBy'].includes(k))delete item[k];
    }
    if(collection==='projects'&&!can('documents'))delete item.materialTakeoffs;
    if(collection==='projects'&&!can('clients')){delete item.clientId;delete item.clientName;}
    if(collection==='people'&&!can('compliance'))for(const k of ['w9FilePath','w9FileName','insuranceFilePath','insuranceFileName','policyNumber','coverageAmount'])delete item[k];
    return item;
  }
  function changes(){
    const result=[];
    for(const [collection,items] of Object.entries(state)){
      if(!Array.isArray(items))continue;
      const old=new Map((baseline[collection]||[]).map(r=>[r.id,r]));
      for(const r of items){const before=old.get(r.id);old.delete(r.id);if(JSON.stringify(before)!==JSON.stringify(r))result.push({collection,id:r.id,after:sanitize(collection,r)});}
      for(const [id] of old)result.push({collection,id,after:null});
    }
    return result;
  }
  function render(){
    if(!isOwner())return '';
    const groups=context().groups||[];
    if(!policyDraft || policyDraft.group_key!==selected){policyDraft=structuredClone(groups.find(g=>g.group_key===selected));dirty=false;}
    const g=policyDraft, locked=selected==='developer';
    if(!g)return '<p role="alert">Access groups could not be loaded. Reload before making changes.</p>';
    return `<article class="panel nl-access"><div class="panel-head"><div><p class="page-kicker">Owner control</p><h2>Group permissions</h2><p>Select a group, choose its project scope, then check the activities it may use.</p></div><span class="status-pill green">Your account · unrestricted Owner</span></div>
      <div class="nl-access-layout"><nav class="nl-access-groups" aria-label="Access groups">${groups.map(group=>`<button type="button" data-access-group="${esc(group.group_key)}" aria-pressed="${selected===group.group_key}"${busy?' disabled':''}>${esc(names[group.group_key]||group.label)}${group.group_key==='developer'?'<small>Owner only</small>':''}</button>`).join('')}</nav>
      <div class="nl-access-detail"><form id="accessGroupForm"><div class="nl-access-heading"><div><h3>${esc(names[selected]||g.label)}</h3><p>${locked?'Developer and unrestricted system access are reserved to your Owner account. This group cannot be assigned to another user.':'Permissions apply to every active account in this group. Unchecked activities are blocked by the server.'}</p></div><label>Project scope<select name="scope"${locked||busy?' disabled':''}><option value="all"${g.scope==='all'?' selected':''}>All company projects</option><option value="assigned"${g.scope==='assigned'?' selected':''}>Linked person’s assigned projects only</option></select></label></div>
      <div class="nl-permission-scroll"><table class="nl-permission-table"><thead><tr><th scope="col">Department / activity</th>${['View','Create','Edit','Delete'].map(a=>`<th scope="col">${a}</th>`).join('')}</tr></thead><tbody>${areas.map(([key,label,description])=>`<tr><th scope="row">${esc(label)}<small>${esc(description)}</small></th>${['view','create','edit','delete'].map(action=>`<td><input type="checkbox" data-permission="${key}:${action}" aria-label="${esc(action)} ${esc(label)}"${g.permissions[key]?.[action]?' checked':''}${locked||busy||key==='requests'&&g.scope==='assigned'||readonly.has(key)&&action!=='view'?' disabled':''}></td>`).join('')}</tr>`).join('')}</tbody></table></div>
      <div class="nl-access-footer"><p>View is required for changes. Archiving and voiding require Edit and Delete. Delete also controls file removal. Security, invitations, recovery, developer access and website publishing remain Owner only.</p>${locked?'':`<button class="button" type="submit"${busy?' disabled':''}>${busy?'Saving…':'Save group permissions'}</button>`}</div><p id="accessSaveStatus" role="status" aria-live="polite">${esc(status)}</p></form></div></div></article>
      <article class="panel nl-access-members"><div class="panel-head"><div><h2>Accounts & assigned groups</h2><p>Choose a group and link the person or company. Inactive accounts cannot read or save company records.</p></div></div>
      <div class="nl-member-list">${(context().members||[]).map(m=>m.role==='owner'?`<div class="nl-member-row"><div><strong>${esc(m.name)}</strong><small>${esc(m.email)}</small></div><span class="status-pill green">Owner · protected</span><p>Full access · cannot be reassigned or deactivated.</p></div>`:`<form class="nl-member-row" data-member-access="${esc(m.id)}"><div><strong>${esc(m.name)}</strong><small>${esc(m.email)}</small></div><label>Group<select name="group">${groups.filter(g=>g.group_key!=='developer').map(g=>`<option value="${g.group_key}"${m.role===g.group_key?' selected':''}>${esc(names[g.group_key]||g.label)}</option>`).join('')}</select></label><label>Linked person / company<select name="person"><option value="">Not linked</option>${state.people.map(p=>`<option value="${esc(p.id)}"${p.id===m.linkedPersonId?' selected':''}>${esc(p.name)} · ${esc(p.type)}</option>`).join('')}</select></label><label>Status<select name="status"><option value="active"${m.status==='active'?' selected':''}>Active</option><option value="inactive"${m.status!=='active'?' selected':''}>Inactive</option></select></label><button class="button secondary" type="submit"${busy?' disabled':''}>Save account</button><span class="nl-member-status" role="status"></span></form>`).join('')}</div></article>`;
  }
  function bind(){
    if(!isOwner())return;
    content.querySelectorAll('[data-access-group]').forEach(b=>b.addEventListener('click',()=>{if(dirty&&!window.confirm('Discard unsaved permission changes?'))return;selected=b.dataset.accessGroup;policyDraft=null;status='';renderRoute();}));
    content.querySelectorAll('[data-permission]').forEach(input=>input.addEventListener('change',()=>{
      const [area,action]=input.dataset.permission.split(':');policyDraft.permissions[area][action]=input.checked;
      if(action==='view'&&!input.checked)for(const a of ['create','edit','delete']){policyDraft.permissions[area][a]=false;const box=content.querySelector(`[data-permission="${area}:${a}"]`);if(box)box.checked=false;}
      if(action!=='view'&&input.checked){policyDraft.permissions[area].view=true;content.querySelector(`[data-permission="${area}:view"]`).checked=true;}
      dirty=true;
    }));
    content.querySelector('#accessGroupForm [name="scope"]')?.addEventListener('change',e=>{policyDraft.scope=e.target.value;if(policyDraft.scope==='assigned')for(const a of ['view','create','edit','delete'])policyDraft.permissions.requests[a]=false;dirty=true;renderRoute();});
    content.querySelector('#accessGroupForm')?.addEventListener('submit',async e=>{
      e.preventDefault();if(busy||selected==='developer')return;
      busy=true;status='Saving permissions…';renderRoute();
      try{const {data,error}=await window.noLimitSupabaseClient.rpc('nl_save_access_group',{org:betaConfig().organizationId,target_group:selected,target_scope:policyDraft.scope,target_permissions:policyDraft.permissions,expected_version:policyDraft.version});if(error)throw error;currentBetaUser.access=data;policyDraft=null;dirty=false;status='Permissions saved. Server protection is active immediately.';}
      catch(error){status=error.message||'Permissions could not be saved.';}
      finally{busy=false;renderRoute();}
    });
    content.querySelectorAll('[data-member-access]').forEach(form=>form.addEventListener('submit',async e=>{
      e.preventDefault();if(busy)return;const button=form.querySelector('button'),message=form.querySelector('[role="status"]'),f=new FormData(form);button.disabled=true;message.textContent='Saving…';
      try{const {data,error}=await window.noLimitSupabaseClient.rpc('nl_set_member_access',{org:betaConfig().organizationId,target_user:form.dataset.memberAccess,target_group:f.get('group'),target_status:f.get('status'),linked_person:f.get('person')||null});if(error)throw error;currentBetaUser.access=data;message.textContent='Account saved.';}
      catch(error){message.textContent=error.message||'Account could not be saved.';}
      finally{button.disabled=false;}
    }));
  }
  async function refresh(){
    const {data,error}=await window.noLimitSupabaseClient.rpc('nl_access_context',{org:betaConfig().organizationId});
    if(error)throw error;
    const previous=context();const changed=JSON.stringify([previous?.role,previous?.scope,previous?.permissions,previous?.linkedPersonId])!==JSON.stringify([data.role,data.scope,data.permissions,data.linkedPersonId]);
    currentBetaUser.access=data;currentBetaUser.membershipRole=data.role;currentBetaUser.role=mappedBetaRole(data.role);
    return changed;
  }
  function apply(){
    const rules=[['[data-edit-client]','client','edit'],['[data-edit-project]','project','edit'],['[data-edit-estimate]','estimate','edit'],['[data-edit-person]','person','edit'],['[data-convert-estimate]','contract','create'],['[data-send-contract]','contract','edit'],['[data-renewal-person]','compliance','create'],['[data-edit-catalog-service]','catalogService','edit'],['[data-remove-catalog-service]','catalogService','delete'],['[data-restore-catalog-service]','catalogService','edit'],['[data-save-visit]','requests','edit'],['[data-discard-visit]','requests','delete'],['[data-new-takeoff]','project','edit'],['[data-edit-takeoff]','project','edit'],['[data-attach-takeoff]','documents','edit'],['[data-detach-takeoff]','documents','edit']];
    for(const [selector,type,action] of rules)content.querySelectorAll(selector).forEach(b=>{const id=b.dataset.editPerson||'';b.hidden=!allows(type,action,id)||(selector==='[data-remove-catalog-service]'&&!can('services','edit'));});
    content.querySelectorAll('[data-pc]').forEach(b=>b.hidden=!allows(b.dataset.pc,b.dataset.pcId?'edit':'create',b.dataset.pcId));
    content.querySelectorAll('a[href^="#"]').forEach(a=>{const route=a.getAttribute('href').slice(1);if(['security','website-insights'].includes(route)?!isOwner():areas.some(([k])=>k===route)&&!can(route))a.hidden=true;});
  }
  function protectForm(type,id=''){
    if(type==='project'&&id&&!can('clients')){const e=dataForm.elements.namedItem('clientId');if(e){e.disabled=true;e.required=false;e.closest('label').hidden=true;}}
    if(type==='project')for(const k of ['contractValue','cost','outstanding']){const el=dataForm.elements.namedItem(k);if(el&&!can('financial',id?'edit':'create')){el.disabled=true;el.required=false;el.closest('label').hidden=true;}}
    if(type==='person'){
      const el=dataForm.elements.namedItem('type');if(el){for(const o of el.options)o.disabled=!can(o.value==='Vendor'?'vendors':'team',id?'edit':'create');if(el.selectedOptions[0]?.disabled)el.value=[...el.options].find(o=>!o.disabled)?.value||'';}
      if(context()?.scope==='assigned'){const p=dataForm.elements.namedItem('projectId');if(p){p.disabled=true;p.required=false;}}
      if(!can('compliance'))for(const k of ['w9File','insuranceFile','policyNumber','coverageAmount']){const e=dataForm.elements.namedItem(k);if(e){e.disabled=true;e.required=false;e.closest('label').hidden=true;}}
    }
  }
  return {isOwner,can,allows,areaFor,routes,snapshot,changes,render,bind,refresh,apply,protectForm};
})();
