/* Shared deterministic quantity rules. Blank measurements remain pending. */
(function(root){
'use strict';
const formulas={length:'LF',perimeter4:'LF',perimeter3:'LF',area:'SF',count:'EA'};
const numeric=(v,nullable=false)=>{if(v===''||v==null){if(nullable)return null;throw Error('Enter a quantity.');}const n=Number(v);if(!Number.isFinite(n)||n<0||n>1000000)throw Error('Use a number between 0 and 1,000,000.');return n;};
function line(value){
 const i={...value};i.id=String(i.id||'');for(const k of ['floor','room','material','profile','source','notes'])i[k]=String(i[k]||'').slice(0,k==='source'?2000:500);
 if(!i.floor.trim()||!i.material.trim())throw Error('Each item needs a floor and material.');
 if(!formulas[i.formula])throw Error('Choose a supported measurement formula.');
 i.count=numeric(i.count,true);if(i.count!==null&&!Number.isInteger(i.count))throw Error('Opening / piece counts must be whole numbers.');
 for(const k of ['width','height','length'])i[k]=numeric(i[k],true);
 i.waste=numeric(i.waste??0);if(i.waste>100)throw Error('Waste must be between 0 and 100%.');
 i.exclude=i.exclude===true; i.reviewed=i.reviewed===true;
 const needed=i.formula==='length'?['length']:['perimeter4','perimeter3','area'].includes(i.formula)?['width','height']:[];
 i.pending=i.count===null||needed.some(k=>i[k]===null);
 i.net=i.pending?null:i.count*(i.formula==='length'?i.length:i.formula==='perimeter4'?2*(i.width+i.height):i.formula==='perimeter3'?i.width+2*i.height:i.formula==='area'?i.width*i.height:1);
 if(i.net!==null&&(!Number.isFinite(i.net)||i.net>100000000))throw Error('Calculated quantity is too large.');
 i.net=i.net===null?null:Math.round(i.net*10000)/10000;i.order=i.net===null?null:(i.exclude?0:Math.ceil(i.net*(1+i.waste/100)*10000)/10000);i.unit=formulas[i.formula];return i;
}
function normalize(v){if(!Array.isArray(v.items)||v.items.length>300)throw Error('Use up to 300 material items.');return {...v,schemaVersion:2,title:String(v.title||'Material Takeoff').slice(0,160),wastePercent:0,notes:String(v.notes||'').slice(0,30000),items:v.items.map(line),sources:Array.isArray(v.sources)?v.sources:[],warnings:Array.isArray(v.warnings)?v.warnings.map(String):[]};}
function parseNotes(text){
 const items=[],warnings=[];let floor='Unassigned floor';
 for(const raw of String(text).split(/\n/)){const s=raw.trim();if(!s)continue;
 const f=s.match(/^(basement|por[aã]o|(?:1st|2nd|3rd|4th|first|second|third)\s+floor|[1234][ºo°]?\s*(?:piso|andar)|attic|half.story)\s*:?\s*$/i);
 if(f){const x=f[1].toLowerCase();floor=/basement|por/.test(x)?'Basement':/1|first/.test(x)?'1st Floor':/2|second/.test(x)?'2nd Floor':/3|third/.test(x)?'3rd Floor':f[1];continue;}
 const m=s.match(/^(.+?)\s*[:=]\s*([\d,]+(?:\.\d+)?)\s*(LF|linear\s*(?:feet|ft)|p[eé]s\s*lineares|SF|sq\s*ft|EA|each|units?)\s*(?:[-—;]\s*(.*))?$/i);
 if(m){const q=Number(m[2].replace(/,/g,'')),unit=m[3].toUpperCase();items.push({id:'',floor,room:'',material:m[1],profile:m[4]||'',formula:/EA|EACH|UNIT/.test(unit)?'count':'length',count:/EA|EACH|UNIT/.test(unit)?q:1,length:/EA|EACH|UNIT/.test(unit)?null:q,width:null,height:null,waste:0,source:s,notes:/SF|SQ/.test(unit)?'Area input: review formula and dimensions.':'',exclude:false,reviewed:false});if(/SF|SQ/.test(unit)){items.pop();warnings.push('Review area measurement: '+s);}continue;}
 warnings.push('Not calculated: '+s);
 }
 return {items,warnings};
}
function totals(v){const t=normalize(v),net={doorCasing:0,windowCasing:0,extension:0,base:0,doubleDoors:0,pocketDoors:0,singleDoors:0,windows:0,casing:0};return {net,order:{casing:0,extension:0,base:0},doors:0,smart:true,pending:t.items.filter(i=>i.pending||!i.reviewed).length};}
root.NoLimitSmartMath={line,normalize,parseNotes,totals,formulas};
})(globalThis);
