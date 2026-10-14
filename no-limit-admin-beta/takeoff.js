/* Project takeoff records and proposal snapshots use the existing private workspace. */
(function (root) {
  const clone = value => JSON.parse(JSON.stringify(value));
  const fields = ['doorCasing','windowCasing','extension','base','doubleDoors','pocketDoors','singleDoors','windows'];
  const materials = [['doorCasing','Door casing'],['windowCasing','Window casing'],['extension','Extension trim - doors / windows'],['base','Baseboard']];
  const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const number = (value, integer = false) => {
    if (value === '' || value == null) return 0;
    const n = Number(value);
    if (!Number.isFinite(n) || n < 0 || (integer && !Number.isInteger(n))) throw new Error(integer ? 'Door and window counts must be whole numbers, zero or greater.' : 'Measurements must be valid numbers, zero or greater.');
    return n;
  };
  function blankFloor(name) { return {name, ...Object.fromEntries(fields.map(key=>[key,0]))}; }
  function template() { return {title:'Material Takeoff',wastePercent:10,notes:'Combine door and window casing only if the profile and width are identical.\nConfirm trim profiles, widths and stock lengths before ordering.',floors:['2nd Floor','1st Floor','Basement'].map(blankFloor)}; }
  function normalize(value) {
    const wastePercent=number(value.wastePercent);
    if (wastePercent > 100) throw new Error('Waste allowance must be between 0% and 100%.');
    if (!Array.isArray(value.floors) || !value.floors.length) throw new Error('Add at least one floor.');
    return {...value,title:String(value.title||'Material Takeoff').trim(),notes:String(value.notes||''),wastePercent,floors:value.floors.map(floor=>{
      const name=String(floor.name||'').trim();if(!name)throw new Error('Enter a name for each floor.');
      return {name,...Object.fromEntries(fields.map(key=>[key,number(floor[key],!materials.some(([field])=>field===key))]))};
    })};
  }
  function totals(value) {
    const takeoff=normalize(value), sum=key=>Math.round(takeoff.floors.reduce((total,floor)=>total+floor[key],0)*1e6)/1e6;
    const net=Object.fromEntries(fields.map(key=>[key,sum(key)]));net.casing=net.doorCasing+net.windowCasing;
    const order=Object.fromEntries(['casing','extension','base'].map(key=>[key,Math.ceil(Math.round(net[key]*(1+takeoff.wastePercent/100)*1e6)/1e6)]));
    return {net,order,doors:net.doubleDoors+net.pocketDoors+net.singleDoors};
  }
  const fmt=value=>Number(value).toLocaleString('en-US',{maximumFractionDigits:2});
  function documentHTML(value,project,client={}) {
    const takeoff=normalize(value),t=totals(takeoff);
    const table=(headers,rows)=>`<table><thead><tr>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(row=>`<tr>${row.map(cell=>`<td>${esc(cell)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
    return `<section class="takeoff-paper"><img class="takeoff-watermark" src="../assets/brand-kit/no-limit-carpentry-nl-monogram.png" alt="" aria-hidden="true" /><div class="takeoff-paper-content"><header class="takeoff-letterhead"><img src="../assets/brand-kit/no-limit-carpentry-approved-logo.png" alt="No Limit"/><div><h2>No Limit Contractor</h2><p>PREMIUM FINISH CARPENTRY</p><small>nolimitcontractor.net | (848) 466-3339</small></div></header><h1>${esc(takeoff.title)}</h1><p class="takeoff-project"><strong>PROJECT: ${esc(project.name)}</strong>${client.name?`<span>Client: ${esc(client.name)}</span>`:''}</p><p>All trim quantities in linear feet (LF)</p><h3>Net material quantities</h3>${table(['Material',...takeoff.floors.map(f=>f.name),'Total'],materials.map(([key,label])=>[label,...takeoff.floors.map(f=>fmt(f[key])),fmt(t.net[key])]))}<h3>Order quantities - including ${fmt(takeoff.wastePercent)}% waste</h3>${table(['Material','Net quantity','Order quantity'],[['Door and window casing',`${fmt(t.net.casing)} LF`,`${fmt(t.order.casing)} LF`],['Extension trim - doors / windows',`${fmt(t.net.extension)} LF`,`${fmt(t.order.extension)} LF`],['Baseboard',`${fmt(t.net.base)} LF`,`${fmt(t.order.base)} LF`]])}<h3>Door schedule - counts by opening / set</h3>${table(['Location','Double','Pocket','Single','Total'],[...takeoff.floors.map(f=>[f.name,f.doubleDoors,f.pocketDoors,f.singleDoors,f.doubleDoors+f.pocketDoors+f.singleDoors]),['Total',t.net.doubleDoors,t.net.pocketDoors,t.net.singleDoors,t.doors]])}${t.net.windows?`<p class="takeoff-window-count">Windows: ${takeoff.floors.map(f=>`${esc(f.name)}: ${fmt(f.windows)}`).join(' | ')}</p>`:''}<div class="takeoff-notes">${takeoff.notes.split('\n').filter(Boolean).map(note=>`<p>${esc(note)}</p>`).join('')}</div><footer>2137 Aldrin Rd, Suite 6B, Ocean, NJ 07712<br/>nolimitcontractor.net | Office: (848) 466-3339</footer></div></section>`;
  }
  async function persist(list,record,save) {
    const before=clone(list),index=list.findIndex(item=>item.id===record.id);
    if(index<0)list.unshift(clone(record));else list[index]=clone(record);
    try {if(await save()===false)throw new Error('The takeoff could not be saved. Refresh if the workspace changed on another device, then try again.');}
    catch(error){list.splice(0,list.length,...before);throw error;}
  }
  async function attach(estimate,takeoff,save) {
    const before=clone(estimate);
    if(takeoff)estimate.materialTakeoff=clone(normalize(takeoff));else delete estimate.materialTakeoff;
    estimate.revision=Number(estimate.revision||1)+1;
    try {if(await save()===false)throw new Error('The proposal attachment could not be saved.');}
    catch(error){for(const key of Object.keys(estimate))delete estimate[key];Object.assign(estimate,before);throw error;}
  }
  root.NoLimitTakeoff={fields,materials,template,blankFloor,normalize,totals,documentHTML,persist,attach,clone};
})(globalThis);
