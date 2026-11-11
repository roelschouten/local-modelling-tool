/* Concept & process views for the model viewer, with an editor.
 *
 * Added to this fork; not part of the upstream Databricks viewer. Changes to
 * docs/index.html are limited to loading this file and switching to it.
 *
 * Reads domain models written as YAML (data-architecture/model/v1):
 *   kind: BusinessModel  -> concepts (entity | event), relationships with a role
 *                           and cardinality per end, metrics, KPIs
 *   kind: ProcessModel   -> processes, steps (creates / produces / changes / next)
 * and draws them as two views:
 *   Concepts   concepts grouped by domain, relationships labelled with the role
 *              of the first end and the cardinality at each end
 *   Processes  one lane per process: trigger -> steps -> events and entities
 *
 * Edit mode changes concepts, relationships, metrics, KPIs, processes, steps and
 * domain metadata in the browser and redraws at once. "Save" writes a new YAML
 * file with a bumped version (and a changelog entry for business models); the
 * file on disk is never changed here. "Excel" exports the loaded models with one
 * tab per part. Comments inside a YAML file are not kept on save; the comment
 * block at the top of the file is.
 *
 * Requires the globals React, cytoscape and jsyaml. SheetJS is loaded on the
 * first Excel export.
 */
(function(){
'use strict';

const h=React.createElement;
const {useState,useEffect,useMemo,useRef,useCallback}=React;

const LS_KEY='dbx-concept-models-v2';
const XLSX_URL='https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
// Fictional example (Lokkr, an automated parcel-locker company): three domains with
// cross-domain references, metrics/KPIs and processes.
const EXAMPLES=['customer/business.yaml','locker-network/business.yaml','locker-network/processes.yaml',
  'parcel-handling/business.yaml','parcel-handling/processes.yaml'].map(f=>'static/examples/lokkr/'+f);
const DOMAIN_HUES=['teal','blue','violet','green','plum','slate','rose','coral'];
const API='data-architecture/model/v1';
const CARDS=['0..1','1..1','0..*','1..*'];
const KINDS=['entity','event'];
const STATUSES=['draft','agreed','deprecated'];
const AGGREGATIONS=['count','count-distinct','sum','min','max','average'];
const DIRECTIONS=['higher-is-better','lower-is-better','within-range'];
const GRAINS=['hour','day','week','month','quarter','year'];
const SLUG=/^[a-z][a-z0-9]*(-[a-z0-9]+)*$/;
const REF=/^([a-z][a-z0-9]*(-[a-z0-9]+)*:)?[a-z][a-z0-9]*(-[a-z0-9]+)*$/;
const TOKEN=/(?<![\w.:-])([a-z][a-z0-9]*(?:-[a-z0-9]+)*(?::[a-z][a-z0-9]*(?:-[a-z0-9]+)*)?)(?![\w-])/g;

const clone=o=>o==null?o:JSON.parse(JSON.stringify(o));
const same=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
const yload=t=>jsyaml.load(t,{schema:jsyaml.CORE_SCHEMA});
const ydump=o=>jsyaml.dump(o,{schema:jsyaml.CORE_SCHEMA,lineWidth:100,noRefs:true,sortKeys:false});

/* ── storage (per-viewer convenience: survives a refresh, nothing more) ── */
function readStored(){
  try{
    let v=JSON.parse(localStorage.getItem(LS_KEY)||'[]');
    if(!Array.isArray(v)||!v.length) v=JSON.parse(localStorage.getItem('dbx-concept-models')||'[]');
    return Array.isArray(v)?v:[];
  }catch(e){ return []; }
}
function writeStored(files){
  try{ localStorage.setItem(LS_KEY,JSON.stringify(files)); }catch(e){ /* quota or blocked: keep in memory */ }
}

/* ── workspace: the loaded models, as edited, plus what they were on load ─
 * ws = {domains:{[dom]:{dom, business, processes,
 *                       files:{business|processes:{name, header, base}}}}, errors:[]}
 */
function headerOf(text){
  const out=[];
  for(const line of String(text).split('\n')){
    if(/^\s*#/.test(line)) out.push(line);
    else if(!line.trim()&&out.length) out.push(line);
    else if(!line.trim()) continue;
    else break;
  }
  while(out.length&&!out[out.length-1].trim()) out.pop();
  return out.join('\n');
}
function slotOf(doc){ return doc.kind==='BusinessModel'?'business':doc.kind==='ProcessModel'?'processes':null; }

function wsFromFiles(files,prev){
  const ws={domains:prev?{...prev.domains}:{},errors:[]};
  files.forEach(f=>{
    let doc;
    try{ doc=yload(f.text); }
    catch(e){ ws.errors.push(f.name+': not valid YAML ('+(e.reason||e.message)+')'); return; }
    if(!doc||typeof doc!=='object'){ ws.errors.push(f.name+': empty file'); return; }
    const dom=doc.metadata&&doc.metadata.domain, slot=slotOf(doc);
    if(!dom){ ws.errors.push(f.name+': metadata.domain is missing'); return; }
    if(!slot){ ws.errors.push(f.name+': kind must be BusinessModel or ProcessModel'); return; }
    let base=null;
    if(f.baseText){ try{ base=yload(f.baseText); }catch(e){ base=null; } }
    const d=ws.domains[dom]={...(ws.domains[dom]||{dom,business:null,processes:null}),files:{...((ws.domains[dom]||{}).files||{})}};
    d[slot]=doc;
    d.files[slot]={name:f.name,header:headerOf(f.text),base:f.baseText===null?null:(base||clone(doc))};
  });
  Object.values(ws.domains).forEach(d=>{
    if(d.processes&&!d.business) ws.errors.push(d.dom+': processes loaded without a business model');
  });
  return ws;
}
function serialize(doc,header){ return (header?header+'\n\n':'')+ydump(doc); }
function wsToFiles(ws){
  const out=[];
  Object.values(ws.domains).forEach(d=>['business','processes'].forEach(slot=>{
    if(!d[slot]) return;
    const f=d.files[slot]||{name:d.dom+'/'+slot+'.yaml',header:'',base:null};
    out.push({name:f.name,text:serialize(d[slot],f.header),baseText:f.base?serialize(f.base,f.header):null});
  }));
  return out;
}
function isDirty(d,slot){ return !!d[slot]&&(!d.files[slot]||!d.files[slot].base||!same(d[slot],d.files[slot].base)); }
function anyDirty(ws){ return Object.values(ws.domains).some(d=>isDirty(d,'business')||isDirty(d,'processes')); }

/* ── references ──────────────────────────────────────────────────────── */
function titleCase(id){ return String(id).split('-').map(w=>w.charAt(0).toUpperCase()+w.slice(1)).join(' '); }
function keyOf(dom,ref){ return ref.indexOf(':')>=0?ref:dom+':'+ref; }
function domOfKey(k){ return k.slice(0,k.indexOf(':')); }
function relRef(dom,key){ return key.indexOf(dom+':')===0?key.slice(dom.length+1):key; }

/* Every reference a domain makes, as [owner-relative ref, setter] pairs, so one
   walk serves dependsOn sync and renames. kind: 'concept' | 'metric'. */
function eachRef(b,p,fn){
  if(b){
    (b.concepts||[]).forEach(c=>{ if(c.specializes) fn('concept',c.specializes,v=>{c.specializes=v;}); });
    (b.relationships||[]).forEach(r=>(r.ends||[]).forEach(e=>{ if(e.concept) fn('concept',e.concept,v=>{e.concept=v;}); }));
    (b.metrics||[]).forEach(m=>(m.requires||[]).forEach((x,i)=>fn('concept',x,v=>{m.requires[i]=v;})));
    (b.kpis||[]).forEach(k=>{
      (k.uses||[]).forEach((x,i)=>fn('metric',x,v=>{k.uses[i]=v;}));
      (k.dimensions||[]).forEach((x,i)=>{ const [base,attr]=String(x).split('.');
        fn('concept',base,v=>{k.dimensions[i]=v+(attr?'.'+attr:'');}); });
    });
  }
  if(p) (p.processes||[]).forEach(pr=>{
    (pr.triggeredBy||[]).forEach((x,i)=>fn('concept',x,v=>{pr.triggeredBy[i]=v;}));
    (pr.steps||[]).forEach(s=>['creates','produces','changes'].forEach(f=>(s[f]||[]).forEach((x,i)=>fn('concept',x,v=>{s[f][i]=v;}))));
  });
}

/* dependsOn is derived from the references; existing version ranges are kept. */
function syncDepends(ws,dom){
  const d=ws.domains[dom]; if(!d||!d.business) return ws;
  const uses={};
  eachRef(d.business,d.processes,(kind,ref)=>{
    const i=String(ref).indexOf(':'); if(i<0) return;
    const td=ref.slice(0,i), id=ref.slice(i+1); (uses[td]=uses[td]||new Set()).add(id);
  });
  const old=d.business.dependsOn||{}, next={};
  Object.keys(old).concat(Object.keys(uses).filter(k=>!(k in old))).forEach(td=>{
    if(!uses[td]) return;
    const t=ws.domains[td];
    const version=old[td]&&old[td].version||(t&&t.business?'^'+t.business.metadata.version:'unmodelled');
    next[td]={version,uses:[...uses[td]].sort()};
  });
  if(same(next,old)) return ws;
  const b=clone(d.business); b.dependsOn=next;
  return {...ws,domains:{...ws.domains,[dom]:{...d,business:b}}};
}

function updateDoc(ws,dom,slot,fn){
  const d=ws.domains[dom]; if(!d) return ws;
  let doc=clone(d[slot]);
  if(!doc&&slot==='processes') doc={apiVersion:API,kind:'ProcessModel',
    metadata:{domain:dom,name:(d.business&&d.business.metadata.name||dom)+' processes',version:'0.1.0',status:'draft'},processes:[]};
  fn(doc);
  const files={...d.files};
  if(!files[slot]) files[slot]={name:dom+'/'+slot+'.yaml',header:'',base:null};
  return syncDepends({...ws,domains:{...ws.domains,[dom]:{...d,[slot]:doc,files}}},dom);
}

/* Rename a concept or metric id and rewrite every reference to it, in all domains. */
function renameId(ws,dom,kind,oldId,newId){
  let out={...ws,domains:{...ws.domains}};
  const oldQ=dom+':'+oldId, newQ=dom+':'+newId;
  Object.keys(out.domains).forEach(d=>{
    const cur=out.domains[d]; const b=clone(cur.business), p=clone(cur.processes);
    if(d===dom&&b){
      const list=kind==='concept'?b.concepts:b.metrics;
      (list||[]).forEach(x=>{ if(x.id===oldId) x.id=newId; });
    }
    eachRef(b,p,(k,ref,set)=>{
      if(k!==kind) return;
      if(d===dom&&ref===oldId) set(newId);
      else if(ref===oldQ) set(d===dom?newId:newQ);
    });
    if(kind==='metric'&&b) (b.kpis||[]).forEach(k=>{
      if(!k.formula) return;
      const esc=s=>s.replace(/[.*+?^${}()|[\]\\-]/g,'\\$&');
      const swap=(f,from,to)=>f.replace(new RegExp('(?<![\\w.:-])'+esc(from)+'(?![\\w-])','g'),to);
      if(d===dom) k.formula=swap(swap(k.formula,oldQ,newId),oldId,newId);
      else k.formula=swap(k.formula,oldQ,newQ);
    });
    out.domains[d]={...cur,business:b,processes:p};
  });
  Object.keys(out.domains).forEach(d=>{ out=syncDepends(out,d); });
  return out;
}

/* Move a concept to another loaded domain.
   - every reference to it, in every loaded domain, is rewritten
   - relationships of the old domain that no longer have an end there (they only
     existed because of the moved concept) move along, since a relationship is
     owned by a domain it involves
   - dependsOn of all domains follows; an id clash in the target gets a suffix */
function moveConcept(ws,src,id,dst){
  const sd=ws.domains[src], td=ws.domains[dst];
  if(!sd||!sd.business||!td||!td.business||src===dst) return {ws,newId:id,moved:[]};
  const original=(sd.business.concepts||[]).find(c=>c.id===id); if(!original) return {ws,newId:id,moved:[]};
  const newId=takenIds(td.business).has(id)?uniqueId(id,takenIds(td.business)):id;
  const oldQ=src+':'+id, newQ=dst+':'+newId;
  let out={...ws,domains:{...ws.domains}};
  Object.keys(out.domains).forEach(d=>{
    const cur=out.domains[d]; if(!cur.business&&!cur.processes) return;
    const b=clone(cur.business), p=clone(cur.processes);
    eachRef(b,p,(k,ref,set)=>{
      if(k!=='concept') return;
      if((d===src&&ref===id)||ref===oldQ) set(d===dst?newId:newQ);
    });
    out.domains[d]={...cur,business:b,processes:p};
  });
  const sb=out.domains[src].business, tb=out.domains[dst].business;
  const c=clone(original); c.id=newId;
  if(c.specializes){ const k=keyOf(src,c.specializes); c.specializes=relRef(dst,k===oldQ?newQ:k); }
  sb.concepts=sb.concepts.filter(x=>x.id!==id);
  tb.concepts=(tb.concepts||[]).concat([c]);
  const moved=[];
  sb.relationships=(sb.relationships||[]).filter(r=>{
    const keys=(r.ends||[]).map(e=>keyOf(src,e.concept));
    if(!keys.includes(newQ)||keys.some(k=>domOfKey(k)===src)) return true;
    const nr=clone(r); nr.ends.forEach((e,i)=>{ e.concept=relRef(dst,keys[i]); });
    if(takenIds(tb).has(nr.id)) nr.id=uniqueId(nr.id,takenIds(tb));
    tb.relationships=(tb.relationships||[]).concat([nr]); moved.push(nr.id);
    return false;
  });
  Object.keys(out.domains).forEach(d=>{ out=syncDepends(out,d); });
  return {ws:out,newId,moved};
}
function moveRelationship(ws,src,id,dst){
  const sd=ws.domains[src], td=ws.domains[dst];
  if(!sd||!sd.business||!td||!td.business||src===dst) return {ws,newId:id};
  const r=(sd.business.relationships||[]).find(x=>x.id===id); if(!r) return {ws,newId:id};
  const nr=clone(r); nr.ends.forEach(e=>{ e.concept=relRef(dst,keyOf(src,e.concept)); });
  if(takenIds(td.business).has(nr.id)) nr.id=uniqueId(nr.id,takenIds(td.business));
  const sb=clone(sd.business), tb=clone(td.business);
  sb.relationships=sb.relationships.filter(x=>x.id!==id);
  tb.relationships=(tb.relationships||[]).concat([nr]);
  let out={...ws,domains:{...ws.domains,[src]:{...sd,business:sb},[dst]:{...td,business:tb}}};
  out=syncDepends(syncDepends(out,src),dst);
  return {ws:out,newId:nr.id};
}
/* Dependency cycles between loaded domains (the CI checker rejects them). */
function dependencyCycles(ws){
  const found=new Set(), doms=Object.keys(ws.domains);
  const deps=d=>Object.keys((ws.domains[d]&&ws.domains[d].business&&ws.domains[d].business.dependsOn)||{}).filter(x=>ws.domains[x]&&ws.domains[x].business);
  const visit=(d,stack)=>{
    const i=stack.indexOf(d);
    if(i>=0){ const cyc=stack.slice(i), m=cyc.indexOf([...cyc].sort()[0]); found.add(cyc.slice(m).concat(cyc.slice(0,m)).join(' \u2192 ')); return; }
    if(stack.length>doms.length) return;
    deps(d).forEach(x=>visit(x,stack.concat([d])));
  };
  doms.forEach(d=>visit(d,[]));
  return [...found].map(c=>c+' \u2192 '+c.split(' \u2192 ')[0]);
}

function autoRelId(a,z){ const t=ref=>String(ref||'').split(':').pop(); return t(a)+'-relates-to-'+t(z); }
function uniqueId(base,taken){
  let id=base, n=2; while(taken.has(id)) id=base+'-'+(n++); return id;
}
function takenIds(b){
  return new Set(['concepts','relationships','metrics','kpis'].flatMap(s=>(b&&b[s]||[]).map(x=>x.id)));
}
function setOrDelete(obj,key,val){
  const empty=val==null||val===''||(Array.isArray(val)&&!val.length)||(typeof val==='object'&&!Array.isArray(val)&&!Object.keys(val).length);
  if(empty) delete obj[key]; else obj[key]=val;
}

/* ── index: everything the views need, resolved ──────────────────────── */
function buildIndex(ws){
  const ix={domains:{},concepts:{},relationships:[],metrics:{},kpis:[],processes:[],issues:[...ws.errors]};
  const doms=Object.keys(ws.domains).sort();
  doms.forEach((dom,i)=>{
    const b=ws.domains[dom].business;
    ix.domains[dom]={dom,modelled:!!b,hue:DOMAIN_HUES[i%DOMAIN_HUES.length],
      name:b&&b.metadata&&b.metadata.name||dom,version:b&&b.metadata&&b.metadata.version||'',status:b&&b.metadata&&b.metadata.status||''};
    if(!b) return;
    (b.concepts||[]).forEach(c=>{
      ix.concepts[dom+':'+c.id]={key:dom+':'+c.id,dom,id:c.id,name:c.name||c.id,kind:c.kind||'entity',raw:c,external:false};
    });
    (b.metrics||[]).forEach(m=>{ ix.metrics[dom+':'+m.id]={key:dom+':'+m.id,dom,raw:m}; });
  });
  const extHues=DOMAIN_HUES.slice(doms.length%DOMAIN_HUES.length).concat(DOMAIN_HUES);
  let extN=0;
  const concept=(dom,ref,where)=>{
    const k=keyOf(dom,String(ref));
    if(ix.concepts[k]) return ix.concepts[k];
    const td=domOfKey(k), id=k.slice(td.length+1);
    if(ix.domains[td]&&ix.domains[td].modelled) ix.issues.push(dom+': '+where+' refers to '+ref+', which '+td+' does not define');
    if(!ix.domains[td]) ix.domains[td]={dom:td,modelled:false,hue:extHues[extN++%extHues.length],name:td,version:'',status:''};
    return ix.concepts[k]=ix.concepts[k]||{key:k,dom:td,id,name:titleCase(id),kind:'external',raw:{},external:true};
  };
  doms.forEach(dom=>{
    const b=ws.domains[dom].business; if(!b) return;
    (b.relationships||[]).forEach(r=>{
      if(!r.ends||r.ends.length!==2||!r.ends[0].concept||!r.ends[1].concept){ ix.issues.push(dom+': relationship '+r.id+' needs two ends'); return; }
      const [a,z]=r.ends;
      const ca=concept(dom,a.concept,'relationship '+r.id), cz=concept(dom,z.concept,'relationship '+r.id);
      ix.relationships.push({key:dom+':'+r.id,dom,raw:r,
        a:{key:ca.key,role:a.role||'',card:a.cardinality||''},z:{key:cz.key,role:z.role||'',card:z.cardinality||''}});
    });
    (b.metrics||[]).forEach(m=>{ (m.requires||[]).forEach(ref=>concept(dom,ref,'metric '+m.id)); });
    (b.kpis||[]).forEach(k=>{
      const metrics=(k.uses||[]).map(u=>keyOf(dom,u));
      metrics.forEach(mk=>{ if(!ix.metrics[mk]&&ix.domains[domOfKey(mk)]&&ix.domains[domOfKey(mk)].modelled)
        ix.issues.push(dom+': KPI '+k.id+' uses '+mk+', which is not a metric'); });
      const dims=(k.dimensions||[]).map(d=>{ const [base,attr]=String(d).split('.'); return {key:concept(dom,base,'KPI '+k.id).key,attr:attr||''}; });
      const needs=new Set();
      metrics.forEach(mk=>{ const m=ix.metrics[mk]; if(m) (m.raw.requires||[]).forEach(ref=>needs.add(keyOf(m.dom,ref))); });
      dims.forEach(d=>needs.add(d.key));
      ix.kpis.push({key:dom+':'+k.id,dom,raw:k,metrics,dims,needs:[...needs]});
    });
  });
  doms.forEach(dom=>{
    const p=ws.domains[dom].processes; if(!p||!ws.domains[dom].business) return;
    (p.processes||[]).forEach(pr=>{
      const steps=(pr.steps||[]).map(s=>({raw:s,id:s.id,
        creates:(s.creates||[]).map(r=>concept(dom,r,'step '+s.id).key),
        produces:(s.produces||[]).map(r=>concept(dom,r,'step '+s.id).key),
        changes:(s.changes||[]).map(r=>concept(dom,r,'step '+s.id).key),
        next:s.next||[]}));
      steps.forEach(s=>s.next.forEach(n=>{ if(!steps.some(t=>t.id===n))
        ix.issues.push(dom+': process '+pr.id+' step '+s.id+' -> unknown step '+n); }));
      ix.processes.push({key:dom+':'+pr.id,dom,raw:pr,steps,
        triggeredBy:(pr.triggeredBy||[]).map(r=>concept(dom,r,'process '+pr.id).key)});
    });
  });
  validate(ws,ix);
  return ix;
}

/* The checks a reviewer would otherwise do by hand; the CI checker stays the gate. */
function validate(ws,ix){
  const add=s=>ix.issues.push(s);
  dependencyCycles(ws).forEach(c=>add(c.split(' \u2192 ')[0]+': dependency cycle '+c+' (a KPI or relationship that needs both belongs in a domain that depends on both)'));
  const caretOk=(rng,v)=>{ const a=String(rng).replace('^','').split('.').map(Number), b=String(v).split('.').map(Number);
    return a[0]===b[0]&&(b[1]>a[1]||(b[1]===a[1]&&b[2]>=a[2])); };
  Object.values(ws.domains).forEach(d=>{
    const b=d.business, dom=d.dom; if(!b) return;
    if(!/^\d+\.\d+\.\d+$/.test(b.metadata&&b.metadata.version||'')) add(dom+': version must look like 1.2.3');
    const seen=new Set();
    ['concepts','relationships','metrics','kpis'].forEach(s=>(b[s]||[]).forEach(x=>{
      if(!SLUG.test(x.id||'')) add(dom+': id "'+x.id+'" must be lower-case words joined by hyphens');
      if(seen.has(x.id)) add(dom+': duplicate id '+x.id); seen.add(x.id);
    }));
    if(!(b.concepts||[]).length) add(dom+': a business model needs at least one concept');
    (b.concepts||[]).forEach(c=>{
      if(!c.name) add(dom+': concept '+c.id+' has no name');
      if(!c.definition||c.definition.length<20) add(dom+': concept '+c.id+' needs a definition (20+ characters)');
      if(!c.identifiedBy) add(dom+': concept '+c.id+' has no identifiedBy');
      if(c.kind==='event'&&!c.occursAt) add(dom+': event '+c.id+' has no occursAt');
    });
    (b.relationships||[]).forEach(r=>(r.ends||[]).forEach((e,i)=>{
      if(!e.role) add(dom+': relationship '+r.id+' end '+(i+1)+' has no role');
      if(!CARDS.includes(e.cardinality)) add(dom+': relationship '+r.id+' end '+(i+1)+' has no valid cardinality');
    }));
    (b.kpis||[]).forEach(k=>{
      const tokens=new Set([...(k.formula||'').matchAll(TOKEN)].map(m=>m[1])), uses=new Set(k.uses||[]);
      if(!k.formula) add(dom+': KPI '+k.id+' has no formula');
      else if(!same([...tokens].sort(),[...uses].sort())) add(dom+': KPI '+k.id+' formula uses '+[...tokens].join(', ')+' but "uses" lists '+[...uses].join(', '));
    });
    Object.entries(b.dependsOn||{}).forEach(([td,spec])=>{
      const t=ws.domains[td];
      if(t&&t.business&&spec.version!=='unmodelled'&&!caretOk(spec.version,t.business.metadata.version))
        add(dom+': depends on '+td+' '+spec.version+', but '+td+' is now '+t.business.metadata.version);
      if(t&&t.business&&spec.version==='unmodelled') add(dom+': '+td+' is loaded now; set a version range instead of "unmodelled"');
    });
    const p=d.processes;
    if(p){
      const pids=new Set();
      (p.processes||[]).forEach(pr=>{
        if(pids.has(pr.id)) add(dom+': duplicate process id '+pr.id); pids.add(pr.id);
        if(!pr.trigger&&!(pr.triggeredBy||[]).length) add(dom+': process '+pr.id+' has no trigger');
        const sids=new Set();
        (pr.steps||[]).forEach(s=>{
          if(sids.has(s.id)) add(dom+': process '+pr.id+' has duplicate step '+s.id); sids.add(s.id);
          if(!s.actor) add(dom+': step '+pr.id+'/'+s.id+' has no actor');
        });
        (pr.steps||[]).forEach(s=>['produces','creates'].forEach(f=>(s[f]||[]).forEach(ref=>{
          const c=ix.concepts[keyOf(dom,ref)];
          if(c&&!c.external&&f==='produces'&&c.kind!=='event') add(dom+': step '+pr.id+'/'+s.id+' produces '+ref+', which is not an event');
          if(c&&!c.external&&f==='creates'&&c.kind!=='entity') add(dom+': step '+pr.id+'/'+s.id+' creates '+ref+', which is not an entity');
        })));
        (pr.triggeredBy||[]).forEach(ref=>{ const c=ix.concepts[keyOf(dom,ref)];
          if(c&&!c.external&&c.kind!=='event') add(dom+': process '+pr.id+' is triggered by '+ref+', which is not an event'); });
      });
    }
  });
}

function sentences(ix,rel){
  const na=ix.concepts[rel.a.key].name, nz=ix.concepts[rel.z.key].name;
  return ['Each '+na+' '+rel.a.role+' '+rel.z.card+' '+nz,'Each '+nz+' '+rel.z.role+' '+rel.a.card+' '+na];
}
function openQuestionCount(raw){ return (raw&&raw.openQuestions||[]).length; }

function usagesOf(ix,key){
  const out=[];
  ix.relationships.forEach(r=>{ if(r.a.key===key||r.z.key===key) out.push('relationship '+r.raw.id); });
  Object.values(ix.metrics).forEach(m=>{ if((m.raw.requires||[]).some(x=>keyOf(m.dom,x)===key)) out.push('metric '+m.raw.id); });
  ix.kpis.forEach(k=>{ if(k.dims.some(d=>d.key===key)) out.push('KPI '+k.raw.id); });
  ix.processes.forEach(p=>{
    if(p.triggeredBy.includes(key)) out.push('process '+p.raw.id);
    p.steps.forEach(s=>{ if(s.creates.concat(s.produces,s.changes).includes(key)) out.push('step '+p.raw.id+'/'+s.id); });
  });
  Object.values(ix.concepts).forEach(c=>{ if(c.raw.specializes&&keyOf(c.dom,c.raw.specializes)===key) out.push('concept '+c.id); });
  return out;
}

/* ── versions ────────────────────────────────────────────────────────── */
const CMIN={'0..1':0,'1..1':1,'0..*':0,'1..*':1}, CMAX={'0..1':1,'1..1':1,'0..*':Infinity,'1..*':Infinity};
const LEVELS=['none','patch','minor','major'];
function changes(base,cur,slot){
  const out=[]; let level='none';
  const bump=(l,why)=>{ if(LEVELS.indexOf(l)>LEVELS.indexOf(level)) level=l; out.push(why); };
  if(!base){ return {level:'none',reasons:['new file, not saved yet'],isNew:true}; }
  if(same(base,cur)) return {level:'none',reasons:[]};
  const strip=m=>{ const x=clone(m||{}); delete x.version; delete x.changelog; return x; };
  if(!same(strip(base.metadata),strip(cur.metadata))) bump('patch','metadata changed');
  const cmp=(sec,label,fnCommon)=>{
    const a=new Map((base[sec]||[]).map(x=>[x.id,x])), b=new Map((cur[sec]||[]).map(x=>[x.id,x]));
    a.forEach((x,id)=>{ if(!b.has(id)) bump('major','removed '+label+' '+id); });
    b.forEach((x,id)=>{ if(!a.has(id)) bump('minor','added '+label+' '+id);
      else if(!same(a.get(id),x)){ const l=fnCommon?fnCommon(a.get(id),x):null; bump(l?l[0]:'patch',l?l[1]:'changed '+label+' '+id); } });
  };
  if(slot==='business'){
    cmp('concepts','concept',(o,n)=>o.kind!==n.kind?['major','concept '+n.id+' changed kind']:null);
    cmp('relationships','relationship',(o,n)=>{
      for(let i=0;i<2;i++){ const eo=(o.ends||[])[i]||{}, en=(n.ends||[])[i]||{};
        if(eo.concept!==en.concept) return ['major','relationship '+n.id+' now connects '+en.concept];
        if(CMIN[en.cardinality]>CMIN[eo.cardinality]||CMAX[en.cardinality]<CMAX[eo.cardinality])
          return ['major','relationship '+n.id+' cardinality tightened ('+eo.cardinality+' → '+en.cardinality+')']; }
      return null; });
    cmp('metrics','metric',(o,n)=>o.aggregation!==n.aggregation?['major','metric '+n.id+' aggregation changed']:null);
    cmp('kpis','KPI',(o,n)=>o.formula!==n.formula?['major','KPI '+n.id+' formula changed']:null);
    if(!same(base.dependsOn,cur.dependsOn)) bump('patch','dependencies changed');
  } else {
    cmp('processes','process',(o,n)=>{
      const so=new Set((o.steps||[]).map(s=>s.id)), sn=new Set((n.steps||[]).map(s=>s.id));
      for(const s of so) if(!sn.has(s)) return ['major','process '+n.id+' lost step '+s];
      for(const s of sn) if(!so.has(s)) return ['minor','process '+n.id+' gained step '+s];
      return null; });
  }
  return {level,reasons:out};
}
function bumpVersion(v,level){
  const p=String(v||'0.0.0').split('.').map(x=>parseInt(x,10)||0);
  if(level==='major') return (p[0]+1)+'.0.0';
  if(level==='minor') return p[0]+'.'+(p[1]+1)+'.0';
  if(level==='patch') return p[0]+'.'+p[1]+'.'+(p[2]+1);
  return p.join('.');
}
function download(name,text,type){
  const blob=text instanceof Blob?text:new Blob([text],{type:type||'application/x-yaml'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href),2000);
}

/* ── Excel export (SheetJS, loaded on first use) ─────────────────────── */
let _xlsx=null;
function loadXlsx(){
  if(window.XLSX) return Promise.resolve(window.XLSX);
  if(_xlsx) return _xlsx;
  _xlsx=new Promise((res,rej)=>{
    const s=document.createElement('script'); s.src=XLSX_URL; s.async=true;
    s.onload=()=>window.XLSX?res(window.XLSX):rej(new Error('Excel library did not initialise'));
    s.onerror=()=>{ _xlsx=null; rej(new Error('could not load the Excel library from '+XLSX_URL)); };
    document.head.appendChild(s);
  });
  return _xlsx;
}
function workbookRows(ws,ix){
  const j=a=>(a||[]).join(', ');
  const lines=a=>(a||[]).join('\n');
  const mapStr=o=>Object.entries(o||{}).map(([k,v])=>k+': '+(Array.isArray(v)?v.join(', '):v)).join('; ');
  const nm=k=>ix.concepts[k]?ix.concepts[k].name:k;
  const doms=Object.values(ws.domains).filter(d=>d.business).sort((a,b)=>a.dom.localeCompare(b.dom));
  const S={Domains:[['Domain','Name','Version','Status','Owner team','Data owner','Description','Depends on','Unsaved changes']],
    Concepts:[['Domain','Id','Name','Kind','Definition','Identified by','Occurs at','Aliases','Lifecycle','Classifications','Specializes','Abstract','Status','Open questions']],
    Relationships:[['Domain','Id','Concept A','Concept A id','Role A','Cardinality A','Concept B','Concept B id','Role B','Cardinality B','Reads as','Definition','Status','Open questions']],
    Metrics:[['Domain','Id','Name','Definition','Aggregation','Unit','Parameters','Requires','Status','Open questions']],
    KPIs:[['Domain','Id','Name','Definition','Formula','Uses metrics','Unit','Direction','Target','Dimensions','Time grain','Needs concepts','Status','Open questions']],
    Processes:[['Domain','Process id','Process','Definition','Trigger','Triggered by','Step #','Step id','Step','Actor','Creates','Produces','Changes','Next']]};
  doms.forEach(d=>{
    const b=d.business, m=b.metadata||{}, dom=d.dom;
    S.Domains.push([dom,m.name||'',m.version||'',m.status||'',(m.owner||{}).team||'',(m.owner||{}).dataOwner||'',m.description||'',
      Object.entries(b.dependsOn||{}).map(([k,v])=>k+' '+v.version).join(', '),
      (isDirty(d,'business')||isDirty(d,'processes'))?'yes':'']);
    (b.concepts||[]).forEach(c=>S.Concepts.push([dom,c.id,c.name||'',c.kind||'',c.definition||'',c.identifiedBy||'',c.occursAt||'',
      mapStr(c.aliases),(c.lifecycle||[]).join(' → '),mapStr(c.classifiedBy),c.specializes||'',c.abstract?'yes':'',c.status||'',lines(c.openQuestions)]));
    (b.relationships||[]).forEach(r=>{
      const rel=ix.relationships.find(x=>x.key===dom+':'+r.id);
      const [a,z]=r.ends||[{},{}];
      S.Relationships.push([dom,r.id,rel?nm(rel.a.key):a.concept,a.concept||'',a.role||'',a.cardinality||'',
        rel?nm(rel.z.key):z.concept,z.concept||'',z.role||'',z.cardinality||'',rel?sentences(ix,rel).join('\n'):'',r.definition||'',r.status||'',lines(r.openQuestions)]);
    });
    (b.metrics||[]).forEach(x=>S.Metrics.push([dom,x.id,x.name||'',x.definition||'',x.aggregation||'',x.unit||'',
      x.parameters?JSON.stringify(x.parameters):'',j(x.requires),x.status||'',lines(x.openQuestions)]));
    (b.kpis||[]).forEach(k=>{
      const kk=ix.kpis.find(x=>x.key===dom+':'+k.id);
      S.KPIs.push([dom,k.id,k.name||'',k.definition||'',k.formula||'',j(k.uses),k.unit||'',k.direction||'',k.target||'',
        j(k.dimensions),j(k.timeGrain),kk?kk.needs.map(nm).join(', '):'',k.status||'',lines(k.openQuestions)]);
    });
    ((d.processes&&d.processes.processes)||[]).forEach(p=>{
      const steps=p.steps&&p.steps.length?p.steps:[{}];
      steps.forEach((s,i)=>S.Processes.push([dom,p.id,p.name||'',i?'':(p.definition||''),i?'':(p.trigger||''),i?'':j(p.triggeredBy),
        s.id?i+1:'',s.id||'',s.name||'',s.actor||'',j(s.creates),j(s.produces),j(s.changes),j(s.next)]));
    });
  });
  return S;
}
function exportExcel(ws,ix,name){
  return loadXlsx().then(X=>{
    const wb=X.utils.book_new(), S=workbookRows(ws,ix);
    Object.values(S).forEach(rows=>rows.forEach(r=>r.forEach((v,i)=>{ if(typeof v==='string') r[i]=v.trim(); })));
    Object.entries(S).forEach(([tab,rows])=>{
      const sh=X.utils.aoa_to_sheet(rows);
      sh['!cols']=rows[0].map((_,c)=>({wch:Math.min(60,Math.max(10,...rows.map(r=>String(r[c]==null?'':r[c]).split('\n')[0].length+2)))}));
      if(rows.length>1) sh['!autofilter']={ref:X.utils.encode_range({s:{r:0,c:0},e:{r:rows.length-1,c:rows[0].length-1}})};
      X.utils.book_append_sheet(wb,sh,tab);
    });
    X.writeFile(wb,name);
  });
}

/* ── logical model (model.json) ──────────────────────────────────────────
   A first-cut logical model in the viewer's own model.json shape, so the
   Graph, ERD and Design views can open it. Mechanical by design:
   - every entity and event becomes a table with a surrogate key <table>_id;
   - classifiedBy becomes one STRING column per classification, lifecycle a
     <table>_status column, occursAt an occurred_at TIMESTAMP on events;
   - a relationship becomes a foreign key on the side that points to at most
     one of the other (each A ... 0..1/1..1 Z -> A.z_id); many-to-many becomes
     an associative table in the relationship's domain;
   - a concept that specializes another gets a 1..1 foreign key to it;
   - concepts referenced from domains that are not loaded become stub tables.
   Attributes beyond keys and classifications are not in the conceptual model;
   add them in Design mode. Regenerating overwrites those additions. */
function snake(s){ return String(s||'').replace(/([a-z0-9])([A-Z])/g,'$1_$2').replace(/[^A-Za-z0-9]+/g,'_').replace(/^_+|_+$/g,'').toLowerCase()||'x'; }
function maxOne(card){ return /^[01]\.\.1$/.test(String(card||'').trim()); }
function cardText(card){ return ({'0..1':'at most one','1..1':'exactly one','0..*':'zero or more','1..*':'one or more'})[card]||card; }
function conceptToLogical(ws,ix,opts){
  opts=opts||{};
  const only=opts.domain||null;
  const domsOut={}, tables={};
  const domOut=dom=>{
    const sd=snake(dom);
    if(domsOut[sd]) return domsOut[sd];
    const d=ix.domains[dom]||{name:dom}, b=ws.domains[dom]&&ws.domains[dom].business, md=(b&&b.metadata)||{};
    return domsOut[sd]={name:sd,division:'business',description:(md.description||'').trim()||(d.modelled?'':'Not loaded: stub tables only.'),
      database_name:sd,references:d.modelled?'conceptual model '+dom+' v'+(md.version||''):'',tags:'',products:[]};
  };
  const attr=(name,type,term,desc,extra)=>({name,column_name:name,type,business_glossary_term:term||'',description:desc||'',
    value_regex:'',tags:'',foreign_key_to:'',references:'',...(extra||{})});
  const inScope=c=>!only||c.dom===only;
  const tableFor=c=>{
    if(tables[c.key]) return tables[c.key];
    const D=domOut(c.dom), t=snake(c.id), raw=c.raw||{};
    const kind=c.external?'external':c.kind;
    const p={name:t,table_name:t,primary_key:t+'_id',subdomain:kind==='event'?'events':kind==='external'?'external':'entities',
      description:(raw.definition||'').trim()||(c.external?'Defined in '+c.dom+', which is not loaded here.':''),
      type:kind==='event'?'Transactional':'Master',division:'',function:'',
      data_type:kind==='event'?'transactional_data':'master_data',source_domains:'',association_edges:'',
      reference:'concept '+c.dom+':'+c.id,tags:'',
      attributes:[attr(t+'_id','BIGINT',c.name+' ID',
        'Surrogate key. Primary key.'+(raw.identifiedBy?' Business key: '+String(raw.identifiedBy).trim()+'.':''),{tags:'primary_key'})]};
    if(!c.external){
      Object.entries(raw.classifiedBy||{}).forEach(([k,vals])=>p.attributes.push(attr(snake(k),'STRING',titleCase(snake(k).replace(/_/g,'-')),
        'Classification'+(Array.isArray(vals)&&vals.length?': one of '+vals.join(', '):'')+'.')));
      if(Array.isArray(raw.lifecycle)&&raw.lifecycle.length)
        p.attributes.push(attr(t+'_status','STRING',c.name+' Status','Lifecycle state: one of '+raw.lifecycle.join(', ')+'.'));
      if(c.kind==='event') p.attributes.push(attr('occurred_at','TIMESTAMP','Occurred At',String(raw.occursAt||'When the event happened.').trim()));
    }
    D.products.push(p);
    return tables[c.key]={p,D};
  };
  const fkTo=t=>t.D.name+'.'+t.p.name+'.'+t.p.primary_key;
  const fkName=(owner,target,role)=>{
    const used=new Set(owner.p.attributes.map(a=>a.name));
    let n=target.p.primary_key;
    if(!used.has(n)) return n;
    const r=snake(String(role||'').replace(/^(is|are|has|have)\s+/i,''));
    n=r+'_'+target.p.primary_key;
    let i=2; while(used.has(n)) n=r+'_'+target.p.primary_key+'_'+(i++);
    return n;
  };
  // tables for everything in scope, in outline order
  Object.values(ix.concepts).filter(c=>!c.external&&inScope(c))
    .sort((a,b)=>a.dom.localeCompare(b.dom)||(a.kind>b.kind?1:a.kind<b.kind?-1:0)||a.name.localeCompare(b.name))
    .forEach(tableFor);
  let fks=0, assoc=0;
  // specializes: table per type, the subtype points to its supertype (1..1)
  Object.values(ix.concepts).filter(c=>!c.external&&inScope(c)&&c.raw.specializes).forEach(c=>{
    const sk=keyOf(c.dom,String(c.raw.specializes)), sup=ix.concepts[sk]||{key:sk,dom:domOfKey(sk),id:sk.slice(domOfKey(sk).length+1),
      name:titleCase(sk.slice(domOfKey(sk).length+1)),kind:'external',raw:{},external:true};
    const o=tableFor(c), t=tableFor(sup);
    o.p.attributes.push(attr(fkName(o,t,'is a'),'BIGINT',sup.name+' ID (Foreign Key)',
      'Foreign key to '+t.D.name+'.'+t.p.name+': each '+c.name+' is a '+sup.name+' (specializes '+sk+').',
      {foreign_key_to:fkTo(t),nullable:false}));
    fks++;
  });
  ix.relationships.forEach(r=>{
    const A=ix.concepts[r.a.key], Z=ix.concepts[r.z.key];
    if(!A||!Z) return;
    if(only&&r.dom!==only&&A.dom!==only&&Z.dom!==only) return;
    const [s1,s2]=sentences(ix,r);
    const why=' ('+r.raw.id+'). '+s1+'. '+s2+'.';
    const addFk=(from,fromRole,to,toCard)=>{
      const o=tableFor(from), t=tableFor(to);
      const n=fkName(o,t,fromRole);
      o.p.attributes.push(attr(n,'BIGINT',to.name+' ID (Foreign Key)',
        'Foreign key to '+t.D.name+'.'+t.p.name+': each '+from.name+' '+fromRole+' '+cardText(toCard)+' '+to.name+why,
        {foreign_key_to:fkTo(t),nullable:String(toCard).trim()==='1..1'?false:true}));
      fks++;
    };
    if(maxOne(r.z.card)&&(!maxOne(r.a.card)||r.z.card==='1..1'||r.a.card!=='1..1')) addFk(A,r.a.role,Z,r.z.card);
    else if(maxOne(r.a.card)) addFk(Z,r.z.role,A,r.a.card);
    else {
      const D=domOut(r.dom), t=snake(r.raw.id);
      const p={name:t,table_name:t,primary_key:t+'_id',subdomain:'associations',
        description:'Associative table for the many-to-many relationship'+why,type:'associative',division:'',function:'',
        data_type:'master_data',source_domains:'',association_edges:'',reference:'relationship '+r.dom+':'+r.raw.id,tags:'',
        attributes:[attr(t+'_id','BIGINT','','Surrogate key. Primary key.',{tags:'primary_key'})]};
      D.products.push(p);
      const self={p,D}, ta=tableFor(A), tz=tableFor(Z);
      [[A,ta,r.a.role],[Z,tz,r.z.role]].forEach(([c,tc,role])=>{
        p.attributes.push(attr(fkName(self,tc,role),'BIGINT',c.name+' ID (Foreign Key)','Foreign key to '+tc.D.name+'.'+tc.p.name+'.',
          {foreign_key_to:fkTo(tc),nullable:false}));
        fks++;
      });
      assoc++;
    }
  });
  const loaded=Object.values(ix.domains).filter(d=>d.modelled&&(!only||d.dom===only));
  const model={type:'logical',name:only?(ix.domains[only]&&ix.domains[only].name||only):'Logical model from '+loaded.map(d=>d.dom).join(', '),
    version:only?(ix.domains[only]&&ix.domains[only].version||'0.1.0'):'0.1.0',
    description:'Generated from the conceptual models ('+loaded.map(d=>d.dom+' v'+d.version).join(', ')+') on '+new Date().toISOString().slice(0,10)
      +'. Keys and foreign keys follow the relationships; other attributes still need to be added.',
    model_conventions:{primary_key_suffix:'_id',table_id_type:'BIGINT',generated_from:'data-architecture/model/v1'},
    domains:Object.values(domsOut).filter(d=>d.products.length).sort((a,b)=>a.name.localeCompare(b.name))};
  return {model,stats:{tables:Object.keys(tables).length+assoc,fks,assoc,domains:model.domains.length}};
}

/* ── colours ─────────────────────────────────────────────────────────── */
function palette(){
  const cs=getComputedStyle(document.body);
  const v=(n,fb)=>(cs.getPropertyValue(n)||'').trim()||fb;
  const hue=(name)=>({line:v('--g-'+name+'-line','#888'),tile:v('--g-'+name+'-tile','#eee'),
    ink:v('--g-'+name+'-ink','#333'),bg:v('--g-'+name+'-bg','#fafafa'),bd:v('--g-'+name+'-bd','#ccc')});
  const c={ink:v('--ink','#1b1f24'),muted:v('--muted','#5b6570'),faint:v('--faint','#8a939e'),
    line:v('--line','#e3e7ec'),lineStrong:v('--line-strong','#cfd5dc'),bg:v('--bg','#f4f6f8'),
    card:v('--card','#fff'),well:v('--well','#f3f5f8'),brand:v('--brand','#ff3621'),
    flow:v('--flow-neutral','#6d7784'),font:cs.fontFamily||'sans-serif',hues:{}};
  DOMAIN_HUES.concat(['amber']).forEach(n=>{c.hues[n]=hue(n);});
  return c;
}

function stylesheet(c){
  const amber=c.hues.amber;
  return [
    {selector:'node',style:{'label':'data(label)','font-family':c.font,'font-size':11,'color':c.ink,
      'text-wrap':'wrap','text-max-width':136,'text-valign':'center','text-halign':'center',
      'width':156,'height':50,'shape':'round-rectangle','background-color':c.card,
      'border-width':2,'border-color':'data(border)'}},
    {selector:'node[kind="event"]',style:{'background-color':amber.tile,'border-color':amber.line,
      'color':amber.ink,'corner-radius':25}},
    {selector:'node[?external]',style:{'border-style':'dashed','background-color':c.well,'color':c.muted}},
    {selector:'node[oq > 0]',style:{'underlay-color':amber.line,'underlay-opacity':0.28,'underlay-padding':5,
      'underlay-shape':'round-rectangle'}},
    {selector:':parent',style:{'label':'data(label)','text-valign':'top','text-halign':'center',
      'text-margin-y':-6,'font-size':12,'font-weight':'bold','color':'data(ink)',
      'background-color':'data(tile)','background-opacity':0.35,'border-width':1.5,
      'border-color':'data(border)','padding':26,'shape':'round-rectangle'}},
    {selector:':parent[?externalDomain]',style:{'border-style':'dashed','background-opacity':0.15}},
    {selector:'node.step',style:{'border-color':c.ink,'border-width':1.5,'width':168,'height':54}},
    {selector:'node.trigger',style:{'shape':'round-tag','background-color':c.well,'border-color':c.faint,
      'border-width':1,'color':c.muted,'font-size':10,'width':176,'height':58,'text-max-width':150}},
    {selector:'edge',style:{'width':1.4,'line-color':c.flow,'curve-style':'bezier',
      'label':'data(label)','font-family':c.font,'font-size':9.5,'color':c.muted,
      'text-background-color':c.bg,'text-background-opacity':0.9,'text-background-padding':2}},
    {selector:'edge[src]',style:{'source-label':'data(src)','source-text-offset':30,'source-text-margin-y':-8}},
    {selector:'edge[tgt]',style:{'target-label':'data(tgt)','target-text-offset':30,'target-text-margin-y':-8}},
    {selector:'edge[?crossDomain]',style:{'line-style':'dashed'}},
    {selector:'edge.next',style:{'line-color':c.ink,'target-arrow-shape':'triangle','target-arrow-color':c.ink,'width':1.6}},
    {selector:'edge.trigger, edge.produces',style:{'line-color':amber.line,'target-arrow-shape':'triangle',
      'target-arrow-color':amber.line,'color':amber.ink}},
    {selector:'edge.creates',style:{'target-arrow-shape':'triangle','target-arrow-color':c.flow}},
    {selector:'edge.changes',style:{'line-style':'dashed','target-arrow-shape':'vee','target-arrow-color':c.faint,
      'line-color':c.faint}},
    {selector:'edge.freetrigger',style:{'line-style':'dotted','target-arrow-shape':'triangle','target-arrow-color':c.faint,
      'line-color':c.faint}},
    {selector:'.faded',style:{'opacity':0.14}},
    {selector:'edge.faded',style:{'text-opacity':0}},
    {selector:'node.sel',style:{'border-color':c.brand,'border-width':3.5}},
    {selector:'edge.sel',style:{'line-color':c.brand,'width':2.6,'color':c.ink}},
  ];
}

/* ── Concepts view elements ──────────────────────────────────────────── */
function domainNode(ix,dom,c){
  const d=ix.domains[dom], hu=c.hues[d.hue];
  return {group:'nodes',data:{id:'d:'+dom,label:d.modelled?(d.name+(d.version?'  v'+d.version:'')):(dom+'  (not modelled)'),
    tile:hu.tile,border:hu.line,ink:hu.ink,externalDomain:!d.modelled}};
}
function conceptNode(ix,k,c,parent){
  const cc=ix.concepts[k], hu=c.hues[ix.domains[cc.dom].hue];
  return {group:'nodes',data:{id:'c:'+k,parent,label:cc.name,kind:cc.kind,external:cc.external,
    border:hu.line,oq:openQuestionCount(cc.raw)}};
}

function conceptElements(ix,c,opt){
  const inScope=k=>{
    const cc=ix.concepts[k];
    if(!opt.showEvents&&cc.kind==='event') return false;
    if(!opt.showExternal&&cc.external) return false;
    return true;
  };
  let keys=Object.keys(ix.concepts).filter(inScope);
  let rels=ix.relationships.filter(r=>inScope(r.a.key)&&inScope(r.z.key));
  if(opt.domain){
    const own=new Set(keys.filter(k=>ix.concepts[k].dom===opt.domain));
    rels=rels.filter(r=>own.has(r.a.key)||own.has(r.z.key));
    const keep=new Set(own); rels.forEach(r=>{keep.add(r.a.key);keep.add(r.z.key);});
    keys=keys.filter(k=>keep.has(k));
  }
  const used=new Set(rels.flatMap(r=>[r.a.key,r.z.key]));
  keys=keys.filter(k=>!ix.concepts[k].external||used.has(k)||opt.domain);
  const doms=[...new Set(keys.map(k=>ix.concepts[k].dom))].sort((a,b)=>
    (ix.domains[b].modelled-ix.domains[a].modelled)||a.localeCompare(b));
  const els=[];
  doms.forEach(d=>els.push(domainNode(ix,d,c)));
  keys.forEach(k=>els.push(conceptNode(ix,k,c,'d:'+ix.concepts[k].dom)));
  rels.forEach(r=>els.push({group:'edges',data:{id:'r:'+r.key,source:'c:'+r.a.key,target:'c:'+r.z.key,
    label:r.a.role,src:r.a.card,tgt:r.z.card,crossDomain:ix.concepts[r.a.key].dom!==ix.concepts[r.z.key].dom}}));
  // Deterministic seed positions: the layout refines from the same start every time.
  const pos={};
  doms.forEach((d,di)=>{
    const ks=keys.filter(k=>ix.concepts[k].dom===d);
    const cx=(di%3)*900, cy=Math.floor(di/3)*700;
    ks.forEach((k,i)=>{ const a=2*Math.PI*i/Math.max(1,ks.length), r=60+ks.length*18;
      pos['c:'+k]={x:cx+r*Math.cos(a),y:cy+r*Math.sin(a)}; });
  });
  els.forEach(e=>{ if(pos[e.data.id]) e.position=pos[e.data.id]; });
  return {els,layout:conceptLayout()};
}

/* fcose (loaded next to cytoscape) handles compound domains much better than the
   built-in cose; fall back to cose when the extension did not load. */
function hasLayout(name){
  try{ const cy=cytoscape({headless:true}); cy.layout({name}); cy.destroy(); return true; }catch(e){ return false; }
}
let _fcose=null;
function conceptLayout(){
  if(_fcose===null) _fcose=hasLayout('fcose');
  if(_fcose) return {name:'fcose',quality:'proof',randomize:false,animate:false,fit:true,padding:50,
    nodeDimensionsIncludeLabels:true,nodeRepulsion:()=>9000,idealEdgeLength:e=>e.data('crossDomain')?110:170,edgeElasticity:()=>0.3,
    nestingFactor:0.15,gravity:0.2,gravityCompound:1.2,gravityRangeCompound:1.6,numIter:4000,
    tilingPaddingVertical:40,tilingPaddingHorizontal:40,nodeSeparation:110,packComponents:true};
  return {name:'cose',animate:false,randomize:false,fit:true,padding:50,
    nodeRepulsion:()=>600000,idealEdgeLength:()=>170,edgeElasticity:()=>60,nestingFactor:1.2,
    gravity:0.6,numIter:3000,componentSpacing:160,nodeOverlap:40};
}

/* ── Processes view elements: one lane per process, left to right ──── */
function processElements(ix,c,opt){
  const COL=240, ROW=92, CGAP=74, LANE_GAP=120;
  const procs=ix.processes.filter(p=>(!opt.domain||p.dom===opt.domain)&&(!opt.process||p.key===opt.process));
  const els=[], pos={}, placed=new Set();
  let y0=0;
  const addConcept=(k,x,y)=>{
    if(placed.has(k)) return false;
    placed.add(k); els.push(conceptNode(ix,k,c,undefined)); pos['c:'+k]={x,y}; return true;
  };
  procs.forEach(p=>{
    const hu=c.hues[ix.domains[p.dom].hue];
    els.push({group:'nodes',data:{id:'p:'+p.key,label:p.raw.name||p.raw.id,tile:c.well,border:hu.line,ink:c.ink,proc:p.key},
      classes:'proc'});
    const steps=p.steps, depth={};
    if(steps.length){
      depth[steps[0].id]=0;
      for(let i=0;i<steps.length;i++) steps.forEach(s=>{ if(depth[s.id]==null) return;
        s.next.forEach(n=>{ if(steps.some(t=>t.id===n)) depth[n]=Math.max(depth[n]||0,Math.min(depth[s.id]+1,steps.length)); }); });
      let maxD=Math.max(0,...Object.values(depth));
      steps.forEach(s=>{ if(depth[s.id]==null) depth[s.id]=++maxD; });
    }
    const byDepth={};
    steps.forEach(s=>{ (byDepth[depth[s.id]]=byDepth[depth[s.id]]||[]).push(s); });
    const maxDepth=Math.max(0,...steps.map(s=>depth[s.id]));
    const stack=Math.max(1,...Object.values(byDepth).map(a=>a.length));
    const outKeys=[];
    steps.forEach(s=>{ s.produces.concat(s.creates,opt.showChanges?s.changes:[]).forEach(k=>{ if(!outKeys.includes(k)) outKeys.push(k); }); });
    const fresh=outKeys.filter(k=>!placed.has(k));
    const laneH=Math.max(stack*ROW,fresh.length*CGAP);
    const mid=y0+laneH/2-ROW/2;
    Object.entries(byDepth).forEach(([d,list])=>{
      list.forEach((s,i)=>{
        const id='s:'+p.key+'/'+s.id;
        els.push({group:'nodes',data:{id,parent:'p:'+p.key,label:(s.raw.name||s.id)+'\n'+(s.raw.actor||''),
          border:c.ink,proc:p.key,step:s.id},classes:'step'});
        pos[id]={x:220+Number(d)*COL,y:mid+(i-(list.length-1)/2)*ROW};
      });
    });
    const first=steps.length?'s:'+p.key+'/'+steps[0].id:null;
    const tx=-60, ty=mid;
    p.triggeredBy.forEach((k,i)=>{
      addConcept(k,tx,ty+i*CGAP);
      if(first) els.push({group:'edges',data:{id:'e:trg:'+p.key+':'+k,source:'c:'+k,target:first,label:'triggers',proc:p.key},classes:'trigger'});
    });
    if(p.raw.trigger&&first){
      const tid='t:'+p.key;
      els.push({group:'nodes',data:{id:tid,label:p.raw.trigger,border:c.faint,proc:p.key},classes:'trigger'});
      pos[tid]={x:tx-(p.triggeredBy.length?200:0),y:ty};
      els.push({group:'edges',data:{id:'e:ft:'+p.key,source:tid,target:first,label:'',proc:p.key},classes:'freetrigger'});
    }
    const cx=220+(maxDepth+1)*COL+40;
    fresh.forEach((k,i)=>addConcept(k,cx,y0+i*CGAP+(laneH-fresh.length*CGAP)/2+CGAP/2-ROW/2+10));
    steps.forEach(s=>{
      const sid='s:'+p.key+'/'+s.id;
      s.next.forEach(n=>{ if(steps.some(t=>t.id===n))
        els.push({group:'edges',data:{id:'e:n:'+sid+'>'+n,source:sid,target:'s:'+p.key+'/'+n,label:'',proc:p.key},classes:'next'}); });
      const link=(list,cls)=>list.forEach(k=>{
        if(!placed.has(k)) addConcept(k,cx,y0+laneH);
        els.push({group:'edges',data:{id:'e:'+cls+':'+sid+'>'+k,source:sid,target:'c:'+k,label:cls,proc:p.key},classes:cls});
      });
      link(s.produces,'produces'); link(s.creates,'creates');
      if(opt.showChanges) link(s.changes,'changes');
    });
    y0+=laneH+LANE_GAP;
  });
  els.forEach(e=>{ if(pos[e.data.id]) e.position=pos[e.data.id]; });
  return {els,layout:{name:'preset',fit:true,padding:50}};
}

/* ── highlight sets ──────────────────────────────────────────────────── */
function highlightFor(ix,sel,mode){
  if(!sel) return null;
  const ids=new Set();
  const addConcept=k=>ids.add('c:'+k);
  if(sel.type==='concept'){
    addConcept(sel.key);
    if(mode==='concepts') ix.relationships.forEach(r=>{
      if(r.a.key===sel.key||r.z.key===sel.key){ ids.add('r:'+r.key); addConcept(r.a.key); addConcept(r.z.key); } });
    else ix.processes.forEach(p=>{
      p.triggeredBy.forEach(k=>{ if(k===sel.key&&p.steps[0]){ ids.add('e:trg:'+p.key+':'+k); ids.add('s:'+p.key+'/'+p.steps[0].id); } });
      p.steps.forEach(s=>['produces','creates','changes'].forEach(cls=>{
        if(s[cls].includes(sel.key)){ const sid='s:'+p.key+'/'+s.id; ids.add(sid); ids.add('e:'+cls+':'+sid+'>'+sel.key); } })); });
  } else if(sel.type==='rel'){
    const r=ix.relationships.find(x=>x.key===sel.key);
    if(r){ ids.add('r:'+r.key); addConcept(r.a.key); addConcept(r.z.key); }
  } else if(sel.type==='kpi'||sel.type==='metric'){
    const needs=sel.type==='kpi'?(ix.kpis.find(x=>x.key===sel.key)||{}).needs
      :((ix.metrics[sel.key]&&ix.metrics[sel.key].raw.requires)||[]).map(r=>keyOf(domOfKey(sel.key),r));
    if(needs){ needs.forEach(addConcept);
      ix.relationships.forEach(r=>{ if(needs.includes(r.a.key)&&needs.includes(r.z.key)) ids.add('r:'+r.key); }); }
  } else if(sel.type==='process'||sel.type==='step'){
    const p=ix.processes.find(x=>x.key===sel.key);
    if(p){
      if(mode==='concepts'){
        p.triggeredBy.forEach(addConcept);
        p.steps.forEach(s=>s.produces.concat(s.creates,s.changes).forEach(addConcept));
      } else ids.add('__proc__'+p.key);
    }
  } else return null;
  return ids;
}

/* ── cytoscape host ──────────────────────────────────────────────────────
 * resetKey changes (view, filters, a new load): lay out from scratch and fit.
 * Same resetKey (an edit): keep the viewport; in Concepts keep every node where
 * it is and put new nodes next to a neighbour, so editing does not reshuffle.
 */
function CyCanvas({graph,styles,highlight,selectedId,onTap,apiRef,resetKey}){
  const ref=useRef(null), cyRef=useRef(null), tapRef=useRef(onTap), sigRef=useRef(''), keyRef=useRef(null);
  tapRef.current=onTap;
  useEffect(()=>{
    const cy=cytoscape({container:ref.current,elements:[],style:styles,minZoom:0.08,maxZoom:3});
    cy.on('tap',e=>{
      if(e.target===cy) tapRef.current(null);
      else tapRef.current({group:e.target.group(),id:e.target.id(),data:e.target.data()});
    });
    cyRef.current=cy; if(apiRef) apiRef.current=cy;
    let ro=null;
    if(window.ResizeObserver){ ro=new ResizeObserver(()=>cy.resize()); ro.observe(ref.current); }
    return ()=>{ if(ro) ro.disconnect(); cy.destroy(); if(apiRef) apiRef.current=null; };
  },[]);
  useEffect(()=>{
    const cy=cyRef.current; if(!cy) return;
    const sig=graph.els.map(e=>e.group+'|'+e.data.id+'|'+(e.data.parent||'')+'|'+(e.data.source||'')+'|'+(e.data.target||'')).sort().join('\n');
    const reset=keyRef.current!==resetKey;
    const unchanged=!reset&&sig===sigRef.current;
    keyRef.current=resetKey; sigRef.current=sig;
    if(unchanged){
      cy.batch(()=>graph.els.forEach(e=>{ const el=cy.getElementById(e.data.id);
        if(el.nonempty()){ el.data(e.data); el.classes(e.classes||''); } }));
      return;
    }
    const keep={};
    if(!reset) cy.nodes().forEach(n=>{ if(!n.isParent()) keep[n.id()]={x:n.position('x'),y:n.position('y')}; });
    cy.batch(()=>{ cy.elements().remove(); cy.add(graph.els); });
    cy.resize();
    if(reset){
      cy.layout(graph.layout).run();
      cy.fit(undefined,50); if(cy.zoom()>1.4) cy.zoom(1.4); cy.center();
      return;
    }
    if(graph.layout.name==='preset'){ cy.layout({...graph.layout,fit:false}).run(); return; }
    const ext=cy.extent(), centre={x:(ext.x1+ext.x2)/2,y:(ext.y1+ext.y2)/2};
    let n=0;
    cy.nodes().forEach(node=>{
      if(node.isParent()) return;
      if(keep[node.id()]) { node.position(keep[node.id()]); return; }
      const nb=node.neighborhood('node').filter(m=>keep[m.id()]);
      const sib=node.parent().nonempty()?node.parent().children().filter(m=>keep[m.id()]):cy.collection();
      const anchor=nb.nonempty()?keep[nb[0].id()]:sib.nonempty()?keep[sib[0].id()]:centre;
      node.position({x:anchor.x+180,y:anchor.y+70+(n++)*60});
    });
  },[graph,resetKey]);
  useEffect(()=>{ if(cyRef.current) cyRef.current.style(styles); },[styles]);
  useEffect(()=>{
    const cy=cyRef.current; if(!cy) return;
    cy.batch(()=>{
      cy.elements().removeClass('faded sel');
      if(highlight){
        const procOnly=[...highlight].find(x=>x.indexOf('__proc__')===0);
        const pk=procOnly&&procOnly.slice(8);
        cy.elements().forEach(el=>{
          if(el.isNode()&&el.isParent()) return;
          const on=pk?el.data('proc')===pk:highlight.has(el.id());
          if(!on) el.addClass('faded');
        });
      }
      if(selectedId){ const el=cy.getElementById(selectedId); if(el&&el.nonempty()) el.addClass('sel'); }
    });
  },[highlight,selectedId,graph]);
  return h('div',{ref,style:{position:'absolute',top:0,left:0,right:0,bottom:0}});
}

/* ── small UI pieces (reuse the viewer's CSS variables and classes) ───── */
const S={
  side:{width:268,background:'var(--card)',borderRight:'1px solid var(--line)',display:'flex',flexDirection:'column',flexShrink:0,minHeight:0},
  right:{width:300,background:'var(--card)',borderLeft:'1px solid var(--line)',flexShrink:0,overflowY:'auto',padding:'14px 16px',fontSize:12.5,lineHeight:1.5},
  h3:{fontSize:11,fontWeight:700,letterSpacing:'.04em',textTransform:'uppercase',color:'var(--muted)',margin:'14px 0 6px'},
  row:{display:'flex',alignItems:'center',gap:6,padding:'4px 8px',borderRadius:6,cursor:'pointer',fontSize:12.5,color:'var(--ink)'},
  chip:{display:'inline-block',padding:'1px 7px',borderRadius:10,fontSize:10.5,fontWeight:650,marginRight:4},
  q:{background:'var(--g-amber-tile)',border:'1px solid var(--g-amber-bd)',color:'var(--g-amber-ink)',borderRadius:7,padding:'6px 8px',margin:'4px 0',fontSize:12},
  bad:{background:'var(--danger-bg)',border:'1px solid var(--danger-bd)',color:'var(--danger)',borderRadius:7,padding:'6px 8px',margin:'4px 0',fontSize:12},
  field:{display:'flex',flexDirection:'column',gap:3,margin:'0 0 9px',fontSize:12},
  label:{fontSize:10.5,fontWeight:700,letterSpacing:'.03em',textTransform:'uppercase',color:'var(--muted)'},
  input:{width:'100%',boxSizing:'border-box',font:'inherit',fontSize:12.5,padding:'5px 7px',borderRadius:6,
    border:'1px solid var(--line-strong)',background:'var(--card)',color:'var(--ink)'},
  hint:{fontSize:11,color:'var(--muted)'},
  box:{border:'1px solid var(--line)',borderRadius:8,padding:'8px 10px',margin:'6px 0 10px',background:'var(--well)'},
};
function Chip({hue,children}){
  const st=hue?{background:'var(--g-'+hue+'-tile)',color:'var(--g-'+hue+'-ink)',border:'1px solid var(--g-'+hue+'-bd)'}
              :{background:'var(--well)',color:'var(--muted)',border:'1px solid var(--line)'};
  return h('span',{style:{...S.chip,...st}},children);
}
function Section({title,children}){ return h(React.Fragment,null,h('div',{style:S.h3},title),children); }
function Questions({raw}){
  const qs=raw&&raw.openQuestions||[];
  return qs.length?h(Section,{title:'Open questions ('+qs.length+')'},qs.map((q,i)=>h('div',{key:i,style:S.q},q))):null;
}
function Link({onClick,children}){
  return h('a',{href:'#',onClick:e=>{e.preventDefault();onClick();},style:{color:'var(--brand-ink)',textDecoration:'none',fontWeight:600}},children);
}
function Btn({onClick,children,title,accent,danger,disabled,small}){
  return h('button',{className:'btn'+(accent?' btn-accent':''),title,disabled,onClick,
    style:{...(danger?{color:'var(--danger)',borderColor:'var(--danger-bd)'}:{}),...(small?{padding:'3px 8px',fontSize:11.5}:{})}},children);
}

/* Form fields. Text commits on blur or Enter (Cmd/Ctrl+Enter in multi-line), so the
   diagram redraws once per change, not per keystroke. */
function useDraft(value){
  const [v,setV]=useState(value);
  useEffect(()=>{ setV(value); },[value]);
  return [v,setV];
}
function Field({label,hint,children,error}){
  return h('label',{style:S.field},h('span',{style:S.label},label),children,
    error?h('span',{style:{...S.hint,color:'var(--danger)'}},error):hint?h('span',{style:S.hint},hint):null);
}
function Text({label,value,onCommit,multiline,placeholder,mono,validate,hint,rows}){
  const str=value==null?'':String(value);
  const [v,setV]=useDraft(str);
  const [err,setErr]=useState('');
  const commit=()=>{
    if(v===str){ setErr(''); return; }
    const e=validate&&validate(v);
    if(e){ setErr(e); return; }
    setErr(''); onCommit(v);
  };
  const props={value:v,placeholder,onChange:e=>setV(e.target.value),onBlur:commit,
    onKeyDown:e=>{
      if(e.key==='Enter'&&(!multiline||e.metaKey||e.ctrlKey)){ e.preventDefault(); commit(); }
      else if(e.key==='Escape'){ setV(str); setErr(''); }
    },
    style:{...S.input,...(mono?{fontFamily:'ui-monospace,SFMono-Regular,Menlo,monospace',fontSize:12}:{})}};
  return h(Field,{label,hint,error:err},multiline?h('textarea',{...props,rows:rows||3,style:{...props.style,resize:'vertical'}}):h('input',{...props,type:'text'}));
}
function Select({label,value,options,onChange,hint,allowEmpty}){
  const opts=options.map(o=>typeof o==='string'?{value:o,label:o}:o);
  return h(Field,{label,hint},h('select',{value:value||'',onChange:e=>onChange(e.target.value),style:S.input},
    (allowEmpty||!value||!opts.some(o=>o.value===value))&&h('option',{value:''},allowEmpty?'— none —':'— choose —'),
    opts.map(o=>h('option',{key:o.value,value:o.value},o.label))));
}
const splitLines=s=>s.split('\n').map(x=>x.trim()).filter(Boolean);
const splitComma=s=>s.split(',').map(x=>x.trim()).filter(Boolean);
function Lines({label,value,onCommit,hint}){
  return h(Text,{label,hint:hint||'One per line.',multiline:true,value:(value||[]).join('\n'),onCommit:v=>onCommit(splitLines(v))});
}
function CommaList({label,value,onCommit,hint,validate}){
  return h(Text,{label,hint:hint||'Comma-separated.',value:(value||[]).join(', '),onCommit:v=>onCommit(splitComma(v)),validate});
}
function MapLines({label,value,onCommit,hint}){
  const str=Object.entries(value||{}).map(([k,v])=>k+': '+(v||[]).join(', ')).join('\n');
  return h(Text,{label,multiline:true,hint:hint||'One per line: name: value, value',value:str,
    validate:v=>splitLines(v).every(l=>/^[^:]+:/.test(l))?'':'Each line needs "name: values".',
    onCommit:v=>{ const o={}; splitLines(v).forEach(l=>{ const i=l.indexOf(':'); o[l.slice(0,i).trim()]=splitComma(l.slice(i+1)); }); onCommit(o); }});
}
/* A list of references, picked from the model, or typed (other domains, not loaded yet). */
function RefList({label,value,options,onChange,hint,single}){
  const list=single?(value?[value]:[]):(value||[]);
  const [typing,setTyping]=useState(false);
  const [txt,setTxt]=useState('');
  const optLabel=v=>{ const o=options.find(x=>x.value===v); return o?o.label:v; };
  const add=v=>{ if(!v) return; if(single) onChange(v); else if(!list.includes(v)) onChange(list.concat([v])); };
  const commitTyped=()=>{ const v=txt.trim(); if(v&&REF.test(v.split('.')[0])){ add(v); setTxt(''); setTyping(false); } };
  return h(Field,{label,hint},
    h('div',{style:{display:'flex',flexWrap:'wrap',gap:4}},
      list.map(v=>h('span',{key:v,style:{...S.chip,display:'inline-flex',alignItems:'center',gap:4,background:'var(--well)',
          border:'1px solid var(--line)',color:'var(--ink)',fontWeight:500,fontSize:11.5,padding:'2px 4px 2px 8px'}},optLabel(v),
        h('button',{type:'button',title:'Remove',onClick:e=>{e.preventDefault(); single?onChange(''):onChange(list.filter(x=>x!==v));},
          style:{border:0,background:'none',cursor:'pointer',color:'var(--muted)',padding:'0 2px',fontSize:13,lineHeight:1}},'×')))),
    typing
      ?h('input',{autoFocus:true,value:txt,placeholder:'domain:concept-id',onChange:e=>setTxt(e.target.value),
          onBlur:()=>{ commitTyped(); setTyping(false); },
          onKeyDown:e=>{ if(e.key==='Enter'){ e.preventDefault(); commitTyped(); } if(e.key==='Escape') setTyping(false); },style:S.input})
      :h('select',{value:'',style:S.input,onChange:e=>{ if(e.target.value==='__type__') setTyping(true); else add(e.target.value); }},
          h('option',{value:''},single?(list.length?'Change…':'Choose…'):'Add…'),
          options.filter(o=>!list.includes(o.value)).map(o=>h('option',{key:o.value,value:o.value},o.label)),
          h('option',{value:'__type__'},'Type a reference…')));
}
function Toggles({label,value,options,onChange}){
  const list=value||[];
  return h(Field,{label},h('div',{style:{display:'flex',flexWrap:'wrap',gap:4}},options.map(o=>h('button',{key:o,type:'button',
    className:'tbtn',onClick:e=>{e.preventDefault(); onChange(list.includes(o)?list.filter(x=>x!==o):options.filter(x=>x===o||list.includes(x)));},
    style:{fontWeight:650,...(list.includes(o)?{background:'var(--brand)',color:'#fff',borderColor:'var(--brand)'}:{})}},o))));
}

/* ── read-only details ───────────────────────────────────────────────── */
function Details({ix,sel,onSelect,mode,onViewMode}){
  if(!sel||sel.type==='domain') return h('div',{style:{color:'var(--muted)'}},
    h('div',{style:{fontWeight:700,color:'var(--ink)',marginBottom:6}},mode==='concepts'?'Concepts':'Processes'),
    mode==='concepts'
      ?'Click a concept, relationship or KPI. A KPI lights up the concepts it needs.'
      :'Click a step, event or entity to see where it is produced, created or changed.',
    ix.issues.length?h(Section,{title:'Model issues ('+ix.issues.length+')'},ix.issues.map((s,i)=>h('div',{key:i,style:S.bad},s))):null);
  const goConcept=k=>onSelect({type:'concept',key:k});
  const name=k=>ix.concepts[k]?ix.concepts[k].name:k;
  const conceptLink=k=>h(Link,{key:k,onClick:()=>goConcept(k)},name(k));
  const join=list=>list.flatMap((x,i)=>i?[', ',x]:[x]);

  if(sel.type==='concept'){
    const cc=ix.concepts[sel.key]; if(!cc) return null;
    const r=cc.raw, d=ix.domains[cc.dom];
    const rels=ix.relationships.filter(x=>x.a.key===cc.key||x.z.key===cc.key);
    const kpis=ix.kpis.filter(k=>k.needs.includes(cc.key));
    const uses=[];
    ix.processes.forEach(p=>{
      if(p.triggeredBy.includes(cc.key)) uses.push({p,s:null,how:'triggers'});
      p.steps.forEach(s=>['creates','produces','changes'].forEach(how=>{ if(s[how].includes(cc.key)) uses.push({p,s,how}); }));
    });
    return h('div',null,
      h('div',{style:{fontSize:16,fontWeight:750,marginBottom:4}},cc.name),
      h('div',null,h(Chip,{hue:cc.kind==='event'?'amber':null},cc.kind),h(Chip,{hue:d.hue},cc.dom),
        r.status&&h(Chip,null,r.status)),
      cc.external&&h('p',{style:{color:'var(--muted)'}},'Referenced from another domain'+(d.modelled?'':' that has no model loaded yet')+'.'),
      r.definition&&h('p',{style:{margin:'10px 0'}},r.definition),
      r.identifiedBy&&h(Section,{title:'Identified by'},h('div',null,r.identifiedBy)),
      r.occursAt&&h(Section,{title:'Occurs at'},h('div',null,r.occursAt)),
      r.aliases&&h(Section,{title:'Aliases'},Object.entries(r.aliases).map(([l,a])=>h('div',{key:l},h(Chip,null,l),(a||[]).join(', ')))),
      r.lifecycle&&h(Section,{title:'Lifecycle'},h('div',null,r.lifecycle.join(' → '))),
      r.classifiedBy&&h(Section,{title:'Classifications'},Object.entries(r.classifiedBy).map(([k,v])=>
        h('div',{key:k},h('b',null,k),': ',(v||[]).join(', ')))),
      rels.length?h(Section,{title:'Relationships ('+rels.length+')'},rels.map(x=>{
        const [s1,s2]=sentences(ix,x);
        return h('div',{key:x.key,style:{...S.row,display:'block',padding:'5px 6px'},onClick:()=>onSelect({type:'rel',key:x.key})},
          h('div',null,s1),h('div',{style:{color:'var(--muted)'}},s2));
      })):null,
      uses.length?h(Section,{title:'In processes'},uses.map((u,i)=>h('div',{key:i,style:{margin:'3px 0'}},
        h(Link,{onClick:()=>{ onSelect({type:u.s?'step':'process',key:u.p.key,step:u.s&&u.s.id}); if(mode!=='processes') onViewMode('processes'); }},
          (u.p.raw.name||u.p.raw.id)+(u.s?' › '+(u.s.raw.name||u.s.id):'')),
        ' ',h('span',{style:{color:'var(--muted)'}},u.how)))):null,
      kpis.length?h(Section,{title:'Needed by KPIs'},kpis.map(k=>h('div',{key:k.key},
        h(Link,{onClick:()=>onSelect({type:'kpi',key:k.key})},k.raw.name||k.raw.id)))):null,
      h(Questions,{raw:r}),
      mode==='processes'&&h('p',{style:{marginTop:14}},h(Link,{onClick:()=>onViewMode('concepts')},'Show in Concepts →')),
      mode==='concepts'&&uses.length>0&&h('p',{style:{marginTop:14}},h(Link,{onClick:()=>onViewMode('processes')},'Show in Processes →'))
    );
  }
  if(sel.type==='rel'){
    const x=ix.relationships.find(r=>r.key===sel.key); if(!x) return null;
    const [s1,s2]=sentences(ix,x);
    return h('div',null,
      h('div',{style:{fontSize:15,fontWeight:750,marginBottom:4}},x.raw.id),
      h('div',null,h(Chip,{hue:ix.domains[x.dom].hue},'owned by '+x.dom),x.raw.status&&h(Chip,null,x.raw.status)),
      h(Section,{title:'Reads as'},h('div',null,s1),h('div',null,s2)),
      x.raw.definition&&h('p',null,x.raw.definition),
      h(Section,{title:'Ends'},[x.a,x.z].map((e,i)=>h('div',{key:i},conceptLink(e.key),' — ',e.role,' ',h(Chip,null,e.card)))),
      h(Questions,{raw:x.raw}));
  }
  if(sel.type==='kpi'){
    const k=ix.kpis.find(x=>x.key===sel.key); if(!k) return null;
    const r=k.raw;
    return h('div',null,
      h('div',{style:{fontSize:16,fontWeight:750,marginBottom:4}},r.name||r.id),
      h('div',null,h(Chip,{hue:'amber'},'KPI'),h(Chip,{hue:ix.domains[k.dom].hue},k.dom),r.status&&h(Chip,null,r.status)),
      r.definition&&h('p',{style:{margin:'10px 0'}},r.definition),
      h(Section,{title:'Formula'},h('code',{style:{fontSize:11.5,background:'var(--well)',padding:'3px 6px',borderRadius:5,display:'block'}},r.formula),
        h('div',{style:{color:'var(--muted)',marginTop:4}},[r.unit,r.direction,r.target&&('target '+r.target)].filter(Boolean).join(' · '))),
      h(Section,{title:'Metrics'},k.metrics.map(mk=>{ const m=ix.metrics[mk];
        return h('div',{key:mk,style:{margin:'4px 0'}},h(Link,{onClick:()=>m&&onSelect({type:'metric',key:mk})},m?(m.raw.name||m.raw.id):mk),
          m&&h('div',{style:{color:'var(--muted)'}},m.raw.aggregation+' · '+m.raw.unit),
          m&&h('div',null,'needs ',join((m.raw.requires||[]).map(ref=>conceptLink(keyOf(m.dom,ref)))))); })),
      k.dims.length?h(Section,{title:'Dimensions'},h('div',null,join(k.dims.map((d,i)=>h('span',{key:i},conceptLink(d.key),d.attr?'.'+d.attr:''))))):null,
      r.timeGrain&&h(Section,{title:'Time grain'},h('div',null,r.timeGrain.join(', '))),
      h(Questions,{raw:r}));
  }
  if(sel.type==='metric'){
    const m=ix.metrics[sel.key]; if(!m) return null;
    const r=m.raw;
    return h('div',null,
      h('div',{style:{fontSize:16,fontWeight:750,marginBottom:4}},r.name||r.id),
      h('div',null,h(Chip,null,'metric'),h(Chip,{hue:ix.domains[m.dom].hue},m.dom)),
      r.definition&&h('p',{style:{margin:'10px 0'}},r.definition),
      h('div',{style:{color:'var(--muted)'}},(r.aggregation||'')+' · '+(r.unit||'')),
      h(Section,{title:'Needs'},h('div',null,join((r.requires||[]).map(ref=>conceptLink(keyOf(m.dom,ref)))))),
      h(Section,{title:'Used by'},ix.kpis.filter(k=>k.metrics.includes(m.key)).map(k=>h('div',{key:k.key},
        h(Link,{onClick:()=>onSelect({type:'kpi',key:k.key})},k.raw.name||k.raw.id)))));
  }
  if(sel.type==='process'||sel.type==='step'){
    const p=ix.processes.find(x=>x.key===sel.key); if(!p) return null;
    const s=sel.step&&p.steps.find(x=>x.id===sel.step);
    if(s) return h('div',null,
      h('div',{style:{fontSize:16,fontWeight:750,marginBottom:4}},s.raw.name||s.id),
      h('div',null,h(Chip,null,'step'),h(Chip,{hue:ix.domains[p.dom].hue},p.dom)),
      h('p',null,'In ',h(Link,{onClick:()=>onSelect({type:'process',key:p.key})},p.raw.name||p.raw.id),', by ',h('b',null,s.raw.actor||'—')),
      s.produces.length?h(Section,{title:'Produces'},h('div',null,join(s.produces.map(conceptLink)))):null,
      s.creates.length?h(Section,{title:'Creates'},h('div',null,join(s.creates.map(conceptLink)))):null,
      s.changes.length?h(Section,{title:'Changes'},h('div',null,join(s.changes.map(conceptLink)))):null,
      s.next.length?h(Section,{title:'Next'},h('div',null,join(s.next.map(n=>{ const t=p.steps.find(x=>x.id===n);
        return h(Link,{key:n,onClick:()=>onSelect({type:'step',key:p.key,step:n})},t?(t.raw.name||n):n); })))):null);
    return h('div',null,
      h('div',{style:{fontSize:16,fontWeight:750,marginBottom:4}},p.raw.name||p.raw.id),
      h('div',null,h(Chip,null,'process'),h(Chip,{hue:ix.domains[p.dom].hue},p.dom)),
      p.raw.definition&&h('p',{style:{margin:'10px 0'}},p.raw.definition),
      (p.triggeredBy.length||p.raw.trigger)&&h(Section,{title:'Triggered by'},
        p.triggeredBy.length?h('div',null,join(p.triggeredBy.map(conceptLink))):null,
        p.raw.trigger&&h('div',{style:{color:'var(--muted)'}},p.raw.trigger)),
      h(Section,{title:'Steps'},p.steps.map(st=>h('div',{key:st.id,style:S.row,onClick:()=>onSelect({type:'step',key:p.key,step:st.id})},
        st.raw.name||st.id,h('span',{style:{color:'var(--muted)',marginLeft:'auto'}},st.raw.actor||'')))),
      mode==='concepts'&&h('p',{style:{marginTop:14}},h(Link,{onClick:()=>onViewMode('processes')},'Show in Processes →')));
  }
  return null;
}

/* ── editors ─────────────────────────────────────────────────────────── */
function conceptOptions(ix,dom,kind){
  return Object.values(ix.concepts).filter(c=>!kind||c.kind===kind||c.external)
    .sort((a,b)=>(a.dom===dom?0:1)-(b.dom===dom?0:1)||a.dom.localeCompare(b.dom)||a.name.localeCompare(b.name))
    .map(c=>({value:relRef(dom,c.key),label:c.name+(c.dom===dom?'':'  ('+c.dom+')')+(c.external?' · not loaded':'')}));
}
function metricOptions(ix,dom){
  return Object.values(ix.metrics).sort((a,b)=>(a.dom===dom?0:1)-(b.dom===dom?0:1)||a.raw.id.localeCompare(b.raw.id))
    .map(m=>({value:relRef(dom,m.key),label:(m.raw.name||m.raw.id)+(m.dom===dom?'':'  ('+m.dom+')')}));
}
function dimensionOptions(ix,dom){
  const out=[];
  conceptOptions(ix,dom).forEach(o=>{
    out.push(o);
    const c=ix.concepts[keyOf(dom,o.value)];
    Object.keys(c&&c.raw.classifiedBy||{}).forEach(attr=>out.push({value:o.value+'.'+attr,label:o.label+' › '+attr}));
  });
  return out;
}
function EditorHead({title,kind,dom,ix,children}){
  return h('div',{style:{marginBottom:10}},
    h('div',{style:{fontSize:15,fontWeight:750}},title||'—'),
    h('div',{style:{margin:'3px 0 0'}},h(Chip,{hue:kind==='event'?'amber':null},kind),h(Chip,{hue:ix.domains[dom]&&ix.domains[dom].hue},dom),children));
}
function slugCheck(taken,current){
  return v=>!SLUG.test(v)?'Use lower-case words joined by hyphens, e.g. locker-site.':(v!==current&&taken.has(v))?'This id is already used in this domain.':'';
}

function ConceptEditor({ws,ix,sel,act}){
  const cc=ix.concepts[sel.key]; if(!cc||cc.external) return h('p',{style:S.hint},'This concept belongs to another domain. Load that domain to edit it.');
  const dom=cc.dom, c=cc.raw, b=ws.domains[dom].business;
  const upd=fn=>act.update(dom,'business',doc=>{ const x=doc.concepts.find(y=>y.id===c.id); if(x) fn(x); });
  const set=(k,v)=>upd(x=>setOrDelete(x,k,v));
  const usages=usagesOf(ix,cc.key);
  return h('div',null,
    h(EditorHead,{title:c.name,kind:c.kind||'entity',dom,ix},c.status&&h(Chip,null,c.status)),
    h(Text,{label:'Id',value:c.id,mono:true,validate:slugCheck(takenIds(b),c.id),hint:'Renaming rewrites every reference to it.',
      onCommit:v=>{ act.rename(dom,'concept',c.id,v); act.select({type:'concept',key:dom+':'+v}); }}),
    h(Text,{label:'Name',value:c.name,onCommit:v=>set('name',v)}),
    h(Select,{label:'Kind',value:c.kind||'entity',options:KINDS,onChange:v=>upd(x=>{ x.kind=v; if(v!=='event') delete x.occursAt; })}),
    h(Select,{label:'Status',value:c.status,options:STATUSES,allowEmpty:true,onChange:v=>set('status',v)}),
    h(Text,{label:'Definition',value:c.definition,multiline:true,rows:4,onCommit:v=>set('definition',v),hint:'Cmd/Ctrl+Enter or click outside to apply.'}),
    h(Text,{label:'Identified by',value:c.identifiedBy,onCommit:v=>set('identifiedBy',v)}),
    (c.kind==='event')&&h(Text,{label:'Occurs at',value:c.occursAt,onCommit:v=>set('occursAt',v),hint:'The moment (or period) the event stands for.'}),
    h(MapLines,{label:'Aliases',value:c.aliases,onCommit:v=>set('aliases',v),hint:'One language per line, e.g. nl: Pakketautomaat, APL'}),
    h(CommaList,{label:'Lifecycle',value:c.lifecycle,onCommit:v=>set('lifecycle',v),hint:'States in order, comma-separated.'}),
    h(MapLines,{label:'Classifications',value:c.classifiedBy,onCommit:v=>set('classifiedBy',v),hint:'One per line, e.g. sizeClass: S, M, L'}),
    h(RefList,{label:'Specializes',single:true,value:c.specializes,options:conceptOptions(ix,dom,'entity').filter(o=>o.value!==c.id),
      onChange:v=>set('specializes',v)}),
    h('label',{style:{...S.field,flexDirection:'row',alignItems:'center',gap:6}},
      h('input',{type:'checkbox',checked:!!c.abstract,onChange:e=>set('abstract',e.target.checked||null)}),h('span',null,'Abstract (only its subtypes exist)')),
    h(Lines,{label:'Open questions',value:c.openQuestions,onCommit:v=>set('openQuestions',v)}),
    h(MoveConcept,{ws,ix,dom,id:c.id,usages,act}),
    h('div',{style:{display:'flex',gap:6,flexWrap:'wrap',marginTop:6}},
      h(Btn,{onClick:()=>act.addRelationship(dom,c.id)},'+ Relationship from here'),
      h(Btn,{danger:true,disabled:usages.length>0,title:usages.length?'Used by: '+usages.join(', '):'Delete this concept',
        onClick:()=>{ act.update(dom,'business',doc=>{ doc.concepts=doc.concepts.filter(x=>x.id!==c.id); }); act.select(null); }},'Delete')),
    usages.length>0&&h('p',{style:S.hint},'Delete is disabled while it is used by '+usages.length+' item(s): '+usages.slice(0,6).join(', ')+(usages.length>6?'…':'')));
}

function MoveConcept({ws,ix,dom,id,usages,act}){
  const [to,setTo]=useState('');
  const targets=Object.values(ws.domains).filter(d=>d.business&&d.dom!==dom).map(d=>d.dom).sort();
  const preview=useMemo(()=>{
    if(!to) return null;
    const r=moveConcept(ws,dom,id,to);
    const before=new Set(dependencyCycles(ws)), after=dependencyCycles(r.ws).filter(x=>!before.has(x));
    return {...r,cycles:after};
  },[ws,dom,id,to]);
  if(!targets.length) return h('p',{style:{...S.hint,margin:'4px 0 10px'}},'Load or add another domain to move this concept there.');
  return h('div',{style:S.box},
    h('div',{style:{...S.label,marginBottom:6}},'Move to another domain'),
    h(Select,{label:'Target domain',value:to,options:targets,allowEmpty:true,onChange:setTo}),
    preview&&h('div',{style:{...S.hint,marginBottom:8}},
      h('div',null,(usages.length?usages.length+' reference(s) will point to '+to+':'+preview.newId:'Nothing refers to it yet')+'.'),
      preview.newId!==id&&h('div',null,'The id is taken in '+to+'; it becomes '+preview.newId+'.'),
      preview.moved.length?h('div',null,'Moves along (no longer involves '+dom+'): '+preview.moved.join(', ')+'.'):null,
      preview.cycles.map((x,i)=>h('div',{key:i,style:{color:'var(--danger)',marginTop:4}},'Creates a dependency cycle: '+x+'.'))),
    h(Btn,{accent:!!to&&!(preview&&preview.cycles.length),disabled:!to,onClick:()=>{ act.moveConcept(dom,id,to); setTo(''); }},
      to?'Move to '+to:'Move'));
}

function RelEditor({ws,ix,sel,act}){
  const rel=ix.relationships.find(x=>x.key===sel.key);
  const dom=domOfKey(sel.key), id=sel.key.slice(dom.length+1);
  const b=ws.domains[dom]&&ws.domains[dom].business;
  const r=b&&(b.relationships||[]).find(x=>x.id===id); if(!r) return null;
  const upd=fn=>act.update(dom,'business',doc=>{ const x=doc.relationships.find(y=>y.id===id); if(x) fn(x); });
  const opts=conceptOptions(ix,dom);
  const end=i=>{ const e=(r.ends||[])[i]||{};
    return h('div',{style:S.box},
      h('div',{style:{...S.label,marginBottom:6}},'End '+(i===0?'A':'B')),
      h(RefList,{label:'Concept',single:true,value:e.concept,options:opts,onChange:v=>{
        const tail=ref=>String(ref||'').split(':').pop();
        const auto=r.id===autoRelId(r.ends[0].concept,r.ends[1].concept)||new RegExp('^'+autoRelId(r.ends[0].concept,r.ends[1].concept)+'-\\d+$').test(r.id);
        const ends=r.ends.map((x,j)=>j===i?v:x.concept);
        const nid=auto?uniqueId(autoRelId(ends[0],ends[1]),new Set([...takenIds(b)].filter(t=>t!==r.id))):r.id;
        upd(x=>{ x.ends[i].concept=v; x.id=nid; });
        if(nid!==r.id) act.select({type:'rel',key:dom+':'+nid});
      }}),
      h(Text,{label:'Role',value:e.role,onCommit:v=>upd(x=>{ x.ends[i].role=v; }),hint:'As a state: "is located at", "is part of".'}),
      h(Select,{label:'Cardinality',value:e.cardinality,options:CARDS,onChange:v=>upd(x=>{ x.ends[i].cardinality=v; }),
        hint:'How many of this end’s concept per one of the other.'})); };
  return h('div',null,
    h(EditorHead,{title:r.id,kind:'relationship',dom,ix}),
    h(Select,{label:'Owned by domain',value:dom,options:Object.values(ws.domains).filter(d=>d.business).map(d=>d.dom).sort(),
      hint:'A relationship belongs to one of the domains it involves.',
      onChange:v=>{ if(v&&v!==dom) act.moveRelationship(dom,id,v); }}),
    rel&&h('div',{style:{...S.box,background:'var(--card)'}},sentences(ix,rel).map((s,i)=>h('div',{key:i,style:i?{color:'var(--muted)'}:null},s))),
    h(Text,{label:'Id',value:r.id,mono:true,validate:slugCheck(takenIds(b),r.id),
      onCommit:v=>{ upd(x=>{ x.id=v; }); act.select({type:'rel',key:dom+':'+v}); }}),
    end(0),end(1),
    h(Btn,{small:true,onClick:()=>upd(x=>{ x.ends=[x.ends[1],x.ends[0]]; })},'⇅ Swap ends'),
    h('div',{style:{height:10}}),
    h(Select,{label:'Status',value:r.status,options:STATUSES,allowEmpty:true,onChange:v=>upd(x=>setOrDelete(x,'status',v))}),
    h(Text,{label:'Definition',value:r.definition,multiline:true,onCommit:v=>upd(x=>setOrDelete(x,'definition',v))}),
    h(Lines,{label:'Open questions',value:r.openQuestions,onCommit:v=>upd(x=>setOrDelete(x,'openQuestions',v))}),
    h(Btn,{danger:true,onClick:()=>{ act.update(dom,'business',doc=>{ doc.relationships=doc.relationships.filter(x=>x.id!==id); }); act.select(null); }},'Delete relationship'));
}

function KpiEditor({ws,ix,sel,act}){
  const dom=domOfKey(sel.key), id=sel.key.slice(dom.length+1);
  const b=ws.domains[dom]&&ws.domains[dom].business;
  const k=b&&(b.kpis||[]).find(x=>x.id===id); if(!k) return null;
  const upd=fn=>act.update(dom,'business',doc=>{ const x=doc.kpis.find(y=>y.id===id); if(x) fn(x); });
  const set=(f,v)=>upd(x=>setOrDelete(x,f,v));
  return h('div',null,
    h(EditorHead,{title:k.name,kind:'KPI',dom,ix},k.status&&h(Chip,null,k.status)),
    h(Text,{label:'Id',value:k.id,mono:true,validate:slugCheck(takenIds(b),k.id),onCommit:v=>{ upd(x=>{ x.id=v; }); act.select({type:'kpi',key:dom+':'+v}); }}),
    h(Text,{label:'Name',value:k.name,onCommit:v=>set('name',v)}),
    h(Select,{label:'Status',value:k.status,options:STATUSES,allowEmpty:true,onChange:v=>set('status',v)}),
    h(Text,{label:'Definition',value:k.definition,multiline:true,onCommit:v=>set('definition',v)}),
    h(RefList,{label:'Uses metrics',value:k.uses,options:metricOptions(ix,dom),onChange:v=>set('uses',v)}),
    h(Text,{label:'Formula',value:k.formula,mono:true,onCommit:v=>set('formula',v),hint:'Arithmetic over the metric ids listed above, e.g. a / b.'}),
    h('div',{style:{marginTop:-4,marginBottom:10}},h(Btn,{small:true,onClick:()=>act.addMetric(dom,k.id)},'+ New metric for this KPI')),
    h(Text,{label:'Unit',value:k.unit,onCommit:v=>set('unit',v)}),
    h(Select,{label:'Direction',value:k.direction,options:DIRECTIONS,onChange:v=>set('direction',v)}),
    h(Text,{label:'Target',value:k.target,onCommit:v=>set('target',v)}),
    h(RefList,{label:'Dimensions',value:k.dimensions,options:dimensionOptions(ix,dom),onChange:v=>set('dimensions',v)}),
    h(Toggles,{label:'Time grain',value:k.timeGrain,options:GRAINS,onChange:v=>set('timeGrain',v)}),
    h(Lines,{label:'Open questions',value:k.openQuestions,onCommit:v=>set('openQuestions',v)}),
    h(Btn,{danger:true,onClick:()=>{ act.update(dom,'business',doc=>{ doc.kpis=doc.kpis.filter(x=>x.id!==id); }); act.select(null); }},'Delete KPI'));
}

function MetricEditor({ws,ix,sel,act}){
  const dom=domOfKey(sel.key), id=sel.key.slice(dom.length+1);
  const b=ws.domains[dom]&&ws.domains[dom].business;
  const m=b&&(b.metrics||[]).find(x=>x.id===id); if(!m) return null;
  const upd=fn=>act.update(dom,'business',doc=>{ const x=doc.metrics.find(y=>y.id===id); if(x) fn(x); });
  const set=(f,v)=>upd(x=>setOrDelete(x,f,v));
  const usedBy=ix.kpis.filter(k=>k.metrics.includes(sel.key)).map(k=>k.raw.id);
  return h('div',null,
    h(EditorHead,{title:m.name,kind:'metric',dom,ix}),
    h(Text,{label:'Id',value:m.id,mono:true,validate:slugCheck(takenIds(b),m.id),hint:'Renaming also updates KPI formulas.',
      onCommit:v=>{ act.rename(dom,'metric',m.id,v); act.select({type:'metric',key:dom+':'+v}); }}),
    h(Text,{label:'Name',value:m.name,onCommit:v=>set('name',v)}),
    h(Text,{label:'Definition',value:m.definition,multiline:true,onCommit:v=>set('definition',v)}),
    h(Select,{label:'Aggregation',value:m.aggregation,options:AGGREGATIONS,onChange:v=>set('aggregation',v)}),
    h(Text,{label:'Unit',value:m.unit,onCommit:v=>set('unit',v)}),
    h(RefList,{label:'Requires concepts',value:m.requires,options:conceptOptions(ix,dom),onChange:v=>set('requires',v)}),
    h(Select,{label:'Status',value:m.status,options:STATUSES,allowEmpty:true,onChange:v=>set('status',v)}),
    usedBy.length?h('p',{style:S.hint},'Used by KPI '+usedBy.join(', ')+'.'):null,
    h(Btn,{danger:true,disabled:usedBy.length>0,title:usedBy.length?'Used by '+usedBy.join(', '):'Delete this metric',
      onClick:()=>{ act.update(dom,'business',doc=>{ doc.metrics=doc.metrics.filter(x=>x.id!==id); }); act.select(null); }},'Delete metric'));
}

function ProcessEditor({ws,ix,sel,act}){
  const [open,setOpen]=useState(sel.step||null);
  useEffect(()=>{ if(sel.step) setOpen(sel.step); },[sel.step]);
  const dom=domOfKey(sel.key), id=sel.key.slice(dom.length+1);
  const p=ws.domains[dom]&&ws.domains[dom].processes;
  const pr=p&&(p.processes||[]).find(x=>x.id===id); if(!pr) return null;
  const upd=fn=>act.update(dom,'processes',doc=>{ const x=doc.processes.find(y=>y.id===id); if(x) fn(x); });
  const updStep=(sid,fn)=>upd(x=>{ const s=x.steps.find(y=>y.id===sid); if(s) fn(s); });
  const taken=new Set((p.processes||[]).map(x=>x.id));
  const stepIds=(pr.steps||[]).map(s=>s.id);
  const move=(i,d)=>upd(x=>{ const j=i+d; if(j<0||j>=x.steps.length) return; const t=x.steps[i]; x.steps[i]=x.steps[j]; x.steps[j]=t; });
  return h('div',null,
    h(EditorHead,{title:pr.name,kind:'process',dom,ix}),
    h(Text,{label:'Id',value:pr.id,mono:true,validate:slugCheck(taken,pr.id),onCommit:v=>{ upd(x=>{ x.id=v; }); act.select({type:'process',key:dom+':'+v}); }}),
    h(Text,{label:'Name',value:pr.name,onCommit:v=>upd(x=>{ x.name=v; })}),
    h(Text,{label:'Definition',value:pr.definition,multiline:true,onCommit:v=>upd(x=>{ x.definition=v; })}),
    h(RefList,{label:'Triggered by events',value:pr.triggeredBy,options:conceptOptions(ix,dom,'event'),onChange:v=>upd(x=>setOrDelete(x,'triggeredBy',v))}),
    h(Text,{label:'Trigger (free text)',value:pr.trigger,onCommit:v=>upd(x=>setOrDelete(x,'trigger',v)),hint:'For a trigger not modelled as an event.'}),
    h('div',{style:S.h3},'Steps ('+(pr.steps||[]).length+')'),
    (pr.steps||[]).map((s,i)=>{
      const isOpen=open===s.id;
      return h('div',{key:s.id,style:{...S.box,background:isOpen?'var(--card)':'var(--well)'}},
        h('div',{style:{display:'flex',alignItems:'center',gap:4}},
          h('a',{href:'#',onClick:e=>{e.preventDefault(); setOpen(isOpen?null:s.id); act.select({type:'step',key:sel.key,step:s.id});},
            style:{flex:1,fontWeight:650,color:'var(--ink)',textDecoration:'none'}},(i+1)+'. '+(s.name||s.id)),
          h(Btn,{small:true,title:'Move up',disabled:i===0,onClick:()=>move(i,-1)},'↑'),
          h(Btn,{small:true,title:'Move down',disabled:i===pr.steps.length-1,onClick:()=>move(i,1)},'↓'),
          h(Btn,{small:true,danger:true,title:'Delete step',onClick:()=>upd(x=>{ x.steps=x.steps.filter(y=>y.id!==s.id);
            x.steps.forEach(y=>{ if(y.next) y.next=y.next.filter(n=>n!==s.id); if(y.next&&!y.next.length) delete y.next; }); })},'×')),
        isOpen&&h('div',{style:{marginTop:8}},
          h(Text,{label:'Step id',value:s.id,mono:true,validate:slugCheck(new Set(stepIds),s.id),
            onCommit:v=>{ upd(x=>{ x.steps.forEach(y=>{ if(y.id===s.id) y.id=v; if(y.next) y.next=y.next.map(n=>n===s.id?v:n); }); });
              setOpen(v); act.select({type:'step',key:sel.key,step:v}); }}),
          h(Text,{label:'Step name',value:s.name,onCommit:v=>updStep(s.id,y=>{ y.name=v; })}),
          h(Text,{label:'Actor',value:s.actor,onCommit:v=>updStep(s.id,y=>{ y.actor=v; })}),
          h(RefList,{label:'Creates (entities)',value:s.creates,options:conceptOptions(ix,dom,'entity'),onChange:v=>updStep(s.id,y=>setOrDelete(y,'creates',v))}),
          h(RefList,{label:'Produces (events)',value:s.produces,options:conceptOptions(ix,dom,'event'),onChange:v=>updStep(s.id,y=>setOrDelete(y,'produces',v))}),
          h(RefList,{label:'Changes (state of)',value:s.changes,options:conceptOptions(ix,dom),onChange:v=>updStep(s.id,y=>setOrDelete(y,'changes',v))}),
          h(RefList,{label:'Next steps',value:s.next,options:stepIds.filter(x=>x!==s.id).map(x=>({value:x,label:((pr.steps.find(y=>y.id===x)||{}).name||x)})),
            onChange:v=>updStep(s.id,y=>setOrDelete(y,'next',v))})));
    }),
    h('div',{style:{display:'flex',gap:6,flexWrap:'wrap',margin:'4px 0 12px'}},
      h(Btn,{onClick:()=>{ const sid=uniqueId('step-'+(stepIds.length+1),new Set(stepIds));
        upd(x=>{ const last=x.steps[x.steps.length-1]; x.steps.push({id:sid,name:'New step',actor:''});
          if(last&&!(last.next&&last.next.length)) last.next=[sid]; });
        setOpen(sid); act.select({type:'step',key:sel.key,step:sid}); }},'+ Step')),
    h(Btn,{danger:true,onClick:()=>{ act.update(dom,'processes',doc=>{ doc.processes=doc.processes.filter(x=>x.id!==id); }); act.select(null); }},'Delete process'));
}

/* Domain metadata, dependencies and saving. */
function DomainEditor({ws,ix,dom,act}){
  const d=ws.domains[dom]; if(!d||!d.business) return null;
  const b=d.business, m=b.metadata||{};
  const upd=fn=>act.update(dom,'business',doc=>{ doc.metadata=doc.metadata||{}; fn(doc.metadata); });
  return h('div',null,
    h(EditorHead,{title:m.name,kind:'domain',dom,ix},h(Chip,null,'v'+m.version)),
    h(Text,{label:'Name',value:m.name,onCommit:v=>upd(x=>{ x.name=v; })}),
    h(Select,{label:'Status',value:m.status,options:STATUSES,onChange:v=>upd(x=>{ x.status=v; })}),
    h(Text,{label:'Description',value:m.description,multiline:true,onCommit:v=>upd(x=>setOrDelete(x,'description',v))}),
    h(Text,{label:'Owner team',value:(m.owner||{}).team,onCommit:v=>upd(x=>{ x.owner={...(x.owner||{}),team:v}; })}),
    h(Text,{label:'Data owner',value:(m.owner||{}).dataOwner,onCommit:v=>upd(x=>{ x.owner={...(x.owner||{}),dataOwner:v}; })}),
    h('div',{style:S.h3},'Depends on'),
    Object.keys(b.dependsOn||{}).length?Object.entries(b.dependsOn).map(([td,spec])=>h('div',{key:td,style:S.box},
      h('div',{style:{fontWeight:650}},td),
      h(Text,{label:'Version range',value:spec.version,mono:true,validate:v=>/^(\^\d+\.\d+\.\d+|unmodelled)$/.test(v)?'':'Use ^1.2.0 or unmodelled.',
        onCommit:v=>act.update(dom,'business',doc=>{ doc.dependsOn[td].version=v; })}),
      h('div',{style:S.hint},'Uses: '+spec.uses.join(', ')))):h('p',{style:S.hint},'No references to other domains.'),
    h('p',{style:S.hint},'The list of domains and used ids follows the references automatically; only the version range is yours.'),
    b.metadata&&b.metadata.changelog&&b.metadata.changelog.length?h(Section,{title:'Changelog'},b.metadata.changelog.slice(0,8).map((c,i)=>
      h('div',{key:i,style:{margin:'3px 0'}},h('b',null,c.version),' — ',c.change))):null);
}

function SavePanel({ws,ix,dom,act}){
  const [level,setLevel]=useState({});
  const [note,setNoteState]=useState('');
  const noteRef=useRef('');
  const setNote=v=>{ noteRef.current=v; setNoteState(v); };
  const d=ws.domains[dom]; if(!d) return null;
  const slots=['business','processes'].filter(s=>d[s]);
  return h('div',{style:{...S.box,background:'var(--card)'}},
    h('div',{style:{fontWeight:700,marginBottom:6}},'Save '+dom),
    slots.map(slot=>{
      const doc=d[slot], f=d.files[slot]||{}, ch=changes(f.base,doc,slot), cur=doc.metadata&&doc.metadata.version;
      const lv=level[slot]||ch.level;
      const unsaved=ch.level!=='none'||ch.isNew;
      const next=bumpVersion(cur,lv);
      return h('div',{key:slot,style:{margin:'6px 0 10px'}},
        h('div',null,h('b',null,slot==='business'?'Business model':'Processes'),' ',h(Chip,null,'v'+cur),
          unsaved?h(Chip,{hue:'amber'},'unsaved'):h(Chip,null,'saved')),
        ch.reasons.length?h('div',{style:{...S.hint,margin:'4px 0'}},ch.reasons.slice(0,6).join('; ')+(ch.reasons.length>6?'…':'')):null,
        h(Select,{label:'Version change',value:lv,options:[{value:'none',label:'keep '+cur},{value:'patch',label:'patch → '+bumpVersion(cur,'patch')},
            {value:'minor',label:'minor → '+bumpVersion(cur,'minor')},{value:'major',label:'major → '+bumpVersion(cur,'major')}],
          hint:ch.level!=='none'?'Suggested: '+ch.level+' (from what changed).':'',onChange:v=>setLevel({...level,[slot]:v})}),
        h(Btn,{accent:unsaved,onClick:()=>{ act.save(dom,slot,lv,slot==='business'?noteRef.current:''); setLevel({...level,[slot]:undefined}); if(slot==='business') setNote(''); }},
          'Save '+(slot==='business'?'business':'processes')+' as v'+next));
    }),
    d.business&&h(Text,{label:'Changelog entry (business model)',value:note,multiline:true,onCommit:setNote,
      hint:'Left empty, the list of changes is used.'}),
    h('p',{style:S.hint},'Saving downloads a new file, e.g. '+dom+'-business-v'+bumpVersion(d.business?d.business.metadata.version:'0.1.0','minor')+
      '.yaml. Put it in your models repository; the file you loaded is not changed.'));
}

/* ── left panel: loading, filters, outline, add ─────────────────────── */
function Outline({ws,ix,mode,opt,setOpt,sel,onSelect,search,setSearch,files,onLoad,onExample,onClear,serverFiles,onServer,
                  editing,setEditing,act,undo,redo,canUndo,canRedo,onExcel,busy,onJson,onLogical,onOpenJson}){
  const fileRef=useRef(null);
  const [newDom,setNewDom]=useState(null);
  const q=search.trim().toLowerCase();
  const match=s=>!q||String(s||'').toLowerCase().includes(q);
  const doms=Object.values(ix.domains).filter(d=>d.modelled);
  const target=opt.domain||(doms.length===1?doms[0].dom:null);
  const isSel=(type,key,step)=>sel&&sel.type===type&&sel.key===key&&(step===undefined||sel.step===step);
  const item=(type,key,label,extra,step)=>h('div',{key:type+key+(step||''),
    style:{...S.row,background:isSel(type,key,step)?'var(--well)':'transparent',fontWeight:isSel(type,key,step)?700:400},
    onClick:()=>onSelect({type,key,step})},label,extra);
  const concepts=Object.values(ix.concepts).filter(c=>!c.external&&(!opt.domain||c.dom===opt.domain)&&match(c.name))
    .sort((a,b)=>a.dom.localeCompare(b.dom)||(a.kind>b.kind?1:a.kind<b.kind?-1:0)||a.name.localeCompare(b.name));
  const kpis=ix.kpis.filter(k=>(!opt.domain||k.dom===opt.domain)&&match(k.raw.name));
  const metrics=Object.values(ix.metrics).filter(m=>(!opt.domain||m.dom===opt.domain)&&match(m.raw.name));
  const procs=ix.processes.filter(p=>(!opt.domain||p.dom===opt.domain)&&match(p.raw.name));
  const rels=ix.relationships.filter(r=>(!opt.domain||r.dom===opt.domain)&&match(r.raw.id+' '+r.a.role));
  const toggle=(label,key,title)=>h('label',{className:'tbtn',title,style:{gap:5,cursor:'pointer',fontWeight:650,color:'var(--muted)'}},
    h('input',{type:'checkbox',checked:!!opt[key],onChange:e=>setOpt({...opt,[key]:e.target.checked}),style:{margin:0,cursor:'pointer'}}),label);
  const dirty=anyDirty(ws);
  return h('div',{style:S.side},
    h('div',{style:{padding:'12px 15px 8px',borderBottom:'1px solid var(--line)'}},
      h('div',{className:'seg',role:'group',style:{marginBottom:8,display:'inline-flex'}},
        h('button',{className:editing?'':'on',onClick:()=>setEditing(false)},'Explore'),
        h('button',{className:editing?'on':'',onClick:()=>setEditing(true),title:'Change the models; save them as new versioned YAML files'},'✎ Edit')),
      h('div',{style:{display:'flex',gap:6,flexWrap:'wrap'}},
        h('button',{className:'btn',title:'Open business.yaml and processes.yaml files (several domains at once). A model.json opens in the Graph view.',
          onClick:()=>fileRef.current&&fileRef.current.click()},'Open…'),
        serverFiles&&serverFiles.length>0&&h('button',{className:'btn',title:'Models bundled with this app',onClick:onServer},'From app ('+serverFiles.length+')'),
        h('button',{className:'btn',title:'Load the fictional Lokkr example (three domains)',onClick:onExample},'Example'),
        doms.length>0&&h('button',{className:'btn',title:'Export the loaded models to Excel (one tab per part)',disabled:busy,onClick:onExcel},busy?'Exporting…':'Excel'),
        doms.length>0&&h('button',{className:'btn',title:'Download a first-cut logical model (model.json) for the Graph, ERD and Design views'+(opt.domain?' (this domain only)':''),onClick:onJson},'JSON'),
        doms.length>0&&onLogical&&h('button',{className:'btn',title:'Open the generated logical model in the ERD view of this app'+(opt.domain?' (this domain only)':''),onClick:onLogical},'Open as logical →'),
        files.length>0&&h('button',{className:'btn',title:'Remove the loaded models',onClick:onClear},'Clear')),
      h('input',{ref:fileRef,type:'file',multiple:true,accept:'.yaml,.yml,.json,application/json',style:{display:'none'},
        onChange:e=>{ const all=[...e.target.files]; e.target.value='';
          const json=all.filter(f=>/\.json$/i.test(f.name)), yaml=all.filter(f=>!/\.json$/i.test(f.name));
          if(yaml.length) onLoad(yaml);
          if(json.length&&onOpenJson) onOpenJson(json[0]); }}),
      h('div',{style:{color:'var(--muted)',fontSize:11.5,marginTop:6}},
        files.length?files.length+' file(s) loaded':'No conceptual model loaded',
        dirty&&h('span',{style:{color:'var(--g-amber-ink)',fontWeight:700}},' · unsaved changes')),
      doms.length>0&&h('select',{value:opt.domain||'',style:{width:'100%',marginTop:8},
        onChange:e=>setOpt({...opt,domain:e.target.value||null,process:null})},
        h('option',{value:''},'All domains'),doms.map(d=>h('option',{key:d.dom,value:d.dom},d.name+(d.version?' v'+d.version:'')))),
      mode==='processes'&&procs.length>0&&h('select',{value:opt.process||'',style:{width:'100%',marginTop:6},
        onChange:e=>setOpt({...opt,process:e.target.value||null})},
        h('option',{value:''},'All processes'),procs.map(p=>h('option',{key:p.key,value:p.key},p.raw.name||p.raw.id))),
      h('div',{style:{display:'flex',gap:4,flexWrap:'wrap',marginTop:8}},
        mode==='concepts'&&toggle('Events','showEvents','Show events'),
        mode==='concepts'&&toggle('Other domains','showExternal','Show concepts referenced from other domains'),
        mode==='processes'&&toggle('State changes','showChanges','Show which concepts each step changes')),
      editing&&h('div',{style:{...S.box,marginTop:10,marginBottom:0}},
        h('div',{style:{display:'flex',alignItems:'center',gap:6,marginBottom:6}},
          h('span',{style:{...S.label,flex:1}},'Add to '+(target||'…')),
          h(Btn,{small:true,title:'Undo',disabled:!canUndo,onClick:undo},'↶'),
          h(Btn,{small:true,title:'Redo',disabled:!canRedo,onClick:redo},'↷')),
        target?h('div',{style:{display:'flex',gap:4,flexWrap:'wrap'}},
          h(Btn,{small:true,onClick:()=>act.addConcept(target,'entity')},'+ Concept'),
          h(Btn,{small:true,onClick:()=>act.addConcept(target,'event')},'+ Event'),
          h(Btn,{small:true,onClick:()=>act.addRelationship(target,sel&&sel.type==='concept'&&domOfKey(sel.key)===target?sel.key.slice(target.length+1):null)},'+ Relationship'),
          h(Btn,{small:true,onClick:()=>act.addKpi(target)},'+ KPI'),
          h(Btn,{small:true,onClick:()=>act.addProcess(target)},'+ Process')):
          h('div',{style:S.hint},'Pick a domain above to add items to it.'),
        h('div',{style:{marginTop:6}},newDom===null
          ?h(Btn,{small:true,onClick:()=>setNewDom('')},'+ Domain')
          :h(Text,{label:'New domain id',value:newDom,mono:true,placeholder:'e.g. apl-network',
              validate:v=>!SLUG.test(v)?'Lower-case words joined by hyphens.':ws.domains[v]?'That domain is already loaded.':'',
              onCommit:v=>{ act.addDomain(v); setNewDom(null); setOpt({...opt,domain:v}); }}))),
      h('input',{type:'search',placeholder:'Search',value:search,onChange:e=>setSearch(e.target.value),
        style:{width:'100%',marginTop:8,boxSizing:'border-box'}})),
    h('div',{style:{overflowY:'auto',flex:1,padding:'4px 8px 16px'}},
      editing&&doms.length>0&&h(Section,{title:'Domains'},doms.map(d=>{ const wd=ws.domains[d.dom];
        const dd=isDirty(wd,'business')||isDirty(wd,'processes');
        return item('domain',d.dom,d.name,h('span',{style:{marginLeft:'auto'}},dd&&h(Chip,{hue:'amber'},'unsaved'),h(Chip,null,'v'+d.version))); })),
      concepts.length>0&&h(Section,{title:'Concepts ('+concepts.length+')'},concepts.map(c=>item('concept',c.key,c.name,
        h('span',{style:{marginLeft:'auto'}},h(Chip,{hue:c.kind==='event'?'amber':ix.domains[c.dom].hue},c.kind==='event'?'event':c.dom))))),
      editing&&rels.length>0&&h(Section,{title:'Relationships ('+rels.length+')'},rels.map(r=>item('rel',r.key,
        h('span',null,ix.concepts[r.a.key].name,h('span',{style:{color:'var(--muted)'}},' '+r.a.role+' '),ix.concepts[r.z.key].name)))),
      kpis.length>0&&h(Section,{title:'KPIs ('+kpis.length+')'},kpis.map(k=>item('kpi',k.key,k.raw.name||k.raw.id))),
      editing&&metrics.length>0&&h(Section,{title:'Metrics ('+metrics.length+')'},metrics.map(m=>item('metric',m.key,m.raw.name||m.raw.id))),
      procs.length>0&&h(Section,{title:'Processes ('+procs.length+')'},procs.map(p=>item('process',p.key,p.raw.name||p.raw.id)))));
}

function Legend({mode}){
  const box=(st,label)=>h('span',{style:{display:'inline-flex',alignItems:'center',gap:5,marginRight:10}},
    h('span',{style:{display:'inline-block',width:22,height:12,borderRadius:3,border:'2px solid var(--flow-neutral)',background:'var(--card)',...st}}),label);
  return h('div',{style:{position:'absolute',left:14,bottom:12,zIndex:10,background:'var(--card)',border:'1px solid var(--line)',
      borderRadius:9,padding:'6px 10px',fontSize:11,color:'var(--muted)',boxShadow:'var(--shadow)'}},
    box({},'entity'),
    box({borderRadius:8,background:'var(--g-amber-tile)',borderColor:'var(--g-amber-line)'},'event'),
    box({borderStyle:'dashed',background:'var(--well)'},'other domain'),
    box({boxShadow:'0 0 0 4px var(--g-amber-bd)'},'open questions'),
    mode==='concepts'
      ?h('span',null,'label = role of first end · numbers = cardinality at that end')
      :h('span',null,'→ next · orange = triggers / produces · dashed = changes'));
}

function RightPanel({ws,ix,sel,setSel,mode,onViewMode,editing,act,opt}){
  if(!editing) return h(Details,{ix,sel,onSelect:setSel,mode,onViewMode});
  const k=sel?(sel.type+':'+sel.key+':'+(sel.step||'')):'none';
  const editor=!sel?null
    :sel.type==='concept'?h(ConceptEditor,{key:k,ws,ix,sel,act})
    :sel.type==='rel'?h(RelEditor,{key:k,ws,ix,sel,act})
    :sel.type==='kpi'?h(KpiEditor,{key:k,ws,ix,sel,act})
    :sel.type==='metric'?h(MetricEditor,{key:k,ws,ix,sel,act})
    :(sel.type==='process'||sel.type==='step')?h(ProcessEditor,{key:sel.type+':'+sel.key,ws,ix,sel,act})
    :null;
  const dom=sel&&sel.type==='domain'?sel.key:(sel?domOfKey(sel.key):opt.domain||(Object.keys(ws.domains).length===1?Object.keys(ws.domains)[0]:null));
  const domIssues=ix.issues.filter(s=>!dom||s.indexOf(dom+':')===0);
  return h('div',null,
    editor||(sel&&sel.type==='domain'?h(DomainEditor,{key:'d:'+dom,ws,ix,dom,act})
      :h('div',{style:{color:'var(--muted)',marginBottom:10}},h('div',{style:{fontWeight:700,color:'var(--ink)',marginBottom:6}},'Edit'),
        'Select something in the diagram or the list to change it, or add new items on the left. Changes redraw at once; nothing is written until you save.')),
    dom&&ws.domains[dom]&&ws.domains[dom].business&&h(SavePanel,{key:'s:'+dom,ws,ix,dom,act}),
    domIssues.length?h(Section,{title:'Issues'+(dom?' in '+dom:'')+' ('+domIssues.length+')'},domIssues.map((s,i)=>h('div',{key:i,style:S.bad},s))):null);
}

/* ── the view ────────────────────────────────────────────────────────── */
function View({mode,theme,palette:pal,onViewMode,onOpenLogical,onOpenJson}){
  const [hist,setHist]=useState(()=>({past:[],now:wsFromFiles(readStored()),future:[]}));
  const ws=hist.now;
  const [serverFiles,setServerFiles]=useState(null);
  const [sel,setSel]=useState(null);
  const [search,setSearch]=useState('');
  const [opt,setOpt]=useState({domain:null,process:null,showEvents:true,showExternal:true,showChanges:false});
  const [notice,setNotice]=useState('');
  const [editing,setEditing]=useState(false);
  const [busy,setBusy]=useState(false);
  const [loadGen,setLoadGen]=useState(0);
  const cyApi=useRef(null);

  const commit=useCallback((fn,record=true)=>setHist(hs=>{
    const next=fn(hs.now); if(next===hs.now) return hs;
    return record?{past:hs.past.concat([hs.now]).slice(-60),now:next,future:[]}:{...hs,now:next};
  }),[]);
  const undo=useCallback(()=>setHist(hs=>hs.past.length?{past:hs.past.slice(0,-1),now:hs.past[hs.past.length-1],future:[hs.now].concat(hs.future)}:hs),[]);
  const redo=useCallback(()=>setHist(hs=>hs.future.length?{past:hs.past.concat([hs.now]),now:hs.future[0],future:hs.future.slice(1)}:hs),[]);
  const files=useMemo(()=>wsToFiles(ws),[ws]);
  useEffect(()=>{ writeStored(files); },[files]);
  useEffect(()=>{
    if(!editing) return;
    const onKey=e=>{
      if(!(e.metaKey||e.ctrlKey)||e.key.toLowerCase()!=='z') return;
      const t=document.activeElement; if(t&&/^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName)) return;
      e.preventDefault(); e.shiftKey?redo():undo();
    };
    window.addEventListener('keydown',onKey);
    return ()=>window.removeEventListener('keydown',onKey);
  },[editing,undo,redo]);

  const load=useCallback((files,replace)=>{
    setHist(hs=>({past:[],now:wsFromFiles(files,replace?null:hs.now),future:[]}));
    setSel(null); setLoadGen(g=>g+1);
  },[]);
  // Files opened from the app's top bar while this view was not mounted.
  const userLoaded=useRef(false);
  useEffect(()=>{
    inbox.listener=fs=>{ userLoaded.current=true; load(fs,false); };
    if(inbox.pending){ const fs=inbox.pending; inbox.pending=null; inbox.listener(fs); }
    return ()=>{ inbox.listener=null; };
  },[load]);
  useEffect(()=>{
    let alive=true;
    fetch('./api/concept-models',{headers:{'Accept':'application/json'}})
      .then(r=>r.ok?r.json():null).then(j=>{ if(!alive||!j||!Array.isArray(j.files)) return;
        setServerFiles(j.files);
        if(!userLoaded.current&&!readStored().length&&j.files.length) load(j.files,true);
      }).catch(()=>{});
    return ()=>{alive=false;};
  },[load]);

  const onLoad=useCallback(list=>{
    Promise.all(list.map(f=>f.text().then(text=>({name:f.name,text})))).then(fs=>load(fs,false));
  },[load]);
  const onExample=useCallback(()=>{
    Promise.all(EXAMPLES.map(u=>fetch(u).then(r=>{ if(!r.ok) throw new Error(u+' '+r.status); return r.text(); })
      .then(text=>({name:'example/'+u.split('/').slice(-3).join('/'),text}))))
      .then(fs=>load(fs,false)).catch(e=>setNotice('Could not load the example: '+e.message));
  },[load]);

  const ix=useMemo(()=>buildIndex(ws),[ws]);
  // The theme class lands on <body> after this render, so read the colours a frame later.
  const [colours,setColours]=useState(palette);
  useEffect(()=>{ const id=requestAnimationFrame(()=>setColours(palette())); return ()=>cancelAnimationFrame(id); },[theme,pal]);
  const styles=useMemo(()=>stylesheet(colours),[colours]);
  const graph=useMemo(()=>mode==='processes'?processElements(ix,colours,opt):conceptElements(ix,colours,opt),
    [ix,colours,mode,opt]);
  const resetKey=mode+'|'+loadGen+'|'+JSON.stringify(opt)+'|'+Object.keys(ws.domains).sort().join(',');
  const highlight=useMemo(()=>highlightFor(ix,sel,mode),[ix,sel,mode]);
  const selectedId=sel?(sel.type==='concept'?'c:'+sel.key:sel.type==='rel'?'r:'+sel.key
    :sel.type==='step'?'s:'+sel.key+'/'+sel.step:null):null;

  const act=useMemo(()=>({
    select:setSel,
    update:(dom,slot,fn)=>commit(w=>updateDoc(w,dom,slot,fn)),
    rename:(dom,kind,oldId,newId)=>commit(w=>renameId(w,dom,kind,oldId,newId)),
    moveConcept:(src,id,dst)=>{
      const r=moveConcept(ws,src,id,dst);
      commit(()=>r.ws);
      setSel({type:'concept',key:dst+':'+r.newId});
      setNotice('Moved '+id+' to '+dst+(r.moved.length?' with '+r.moved.length+' relationship(s)':'')+'.');
    },
    moveRelationship:(src,id,dst)=>{
      const r=moveRelationship(ws,src,id,dst);
      commit(()=>r.ws);
      setSel({type:'rel',key:dst+':'+r.newId});
    },
    addConcept:(dom,kind)=>{
      const b=ws.domains[dom].business, id=uniqueId(kind==='event'?'new-event':'new-concept',takenIds(b));
      commit(w=>updateDoc(w,dom,'business',doc=>{ doc.concepts=doc.concepts||[];
        doc.concepts.push(kind==='event'?{id,kind,name:'New event',definition:'',identifiedBy:'',occursAt:'',status:'draft'}
          :{id,kind,name:'New concept',definition:'',identifiedBy:'',status:'draft'}); }));
      setSel({type:'concept',key:dom+':'+id});
    },
    addRelationship:(dom,fromId)=>{
      const b=ws.domains[dom].business, cs=b.concepts||[];
      if(!cs.length){ setNotice('Add a concept to '+dom+' first.'); return; }
      const a=fromId||cs[0].id, z=(cs.find(c=>c.id!==a)||cs[0]).id;
      const id=uniqueId(autoRelId(a,z),takenIds(b));
      commit(w=>updateDoc(w,dom,'business',doc=>{ doc.relationships=doc.relationships||[];
        doc.relationships.push({id,status:'draft',ends:[{concept:a,role:'relates to',cardinality:'0..*'},{concept:z,role:'is related to',cardinality:'1..1'}]}); }));
      setSel({type:'rel',key:dom+':'+id});
    },
    addKpi:dom=>{
      const b=ws.domains[dom].business, id=uniqueId('new-kpi',takenIds(b));
      commit(w=>updateDoc(w,dom,'business',doc=>{ doc.kpis=doc.kpis||[];
        doc.kpis.push({id,name:'New KPI',definition:'',formula:'',uses:[],unit:'ratio',direction:'higher-is-better',status:'draft'}); }));
      setSel({type:'kpi',key:dom+':'+id});
    },
    addMetric:(dom,kpiId)=>{
      const b=ws.domains[dom].business, id=uniqueId('new-metric',takenIds(b));
      commit(w=>updateDoc(w,dom,'business',doc=>{ doc.metrics=doc.metrics||[];
        doc.metrics.push({id,name:'New metric',definition:'',aggregation:'count',unit:'',requires:[]});
        if(kpiId){ const k=doc.kpis.find(x=>x.id===kpiId); if(k){ k.uses=(k.uses||[]).concat([id]); k.formula=k.formula?k.formula:id; } } }));
      setSel({type:'metric',key:dom+':'+id});
    },
    addProcess:dom=>{
      const p=ws.domains[dom].processes, id=uniqueId('new-process',new Set(((p&&p.processes)||[]).map(x=>x.id)));
      commit(w=>updateDoc(w,dom,'processes',doc=>{ doc.processes=doc.processes||[];
        doc.processes.push({id,name:'New process',definition:'',trigger:'',steps:[{id:'step-1',name:'First step',actor:''}]}); }));
      setSel({type:'process',key:dom+':'+id});
    },
    addDomain:dom=>{
      commit(w=>({...w,domains:{...w.domains,[dom]:{dom,business:{apiVersion:API,kind:'BusinessModel',
        metadata:{domain:dom,name:titleCase(dom),version:'0.1.0',status:'draft',owner:{team:'TBD',dataOwner:'TBD'},description:'',
          changelog:[{version:'0.1.0',change:'Initial draft.'}]},dependsOn:{},concepts:[],relationships:[],metrics:[],kpis:[]},
        processes:null,files:{business:{name:dom+'/business.yaml',header:'',base:null}}}}}));
      setSel({type:'domain',key:dom});
    },
    save:(dom,slot,level,note)=>{
      const d=ws.domains[dom], f=d.files[slot]||{name:dom+'/'+slot+'.yaml',header:''};
      const doc=clone(d[slot]), ch=changes(f.base,doc,slot);
      const v=bumpVersion(doc.metadata.version,level);
      doc.metadata.version=v;
      if(slot==='business'&&level!=='none'){
        doc.metadata.changelog=[{version:v,change:(note||'').trim()||ch.reasons.join('; ')||'Updated.'}].concat(doc.metadata.changelog||[]);
      }
      const name=dom+'-'+slot+'-v'+v+'.yaml';
      download(name,serialize(doc,f.header));
      commit(w=>{
        const nd={...w.domains[dom],[slot]:doc,files:{...w.domains[dom].files,[slot]:{...f,name,base:clone(doc)}}};
        let out={...w,domains:{...w.domains,[dom]:nd}};
        return out;
      });
      setNotice('Saved '+name+'.');
    },
  }),[ws,commit]);

  const onTap=useCallback(t=>{
    if(!t) return setSel(null);
    const id=t.id;
    if(id.startsWith('c:')) setSel({type:'concept',key:id.slice(2)});
    else if(id.startsWith('r:')) setSel({type:'rel',key:id.slice(2)});
    else if(id.startsWith('s:')) setSel({type:'step',key:t.data.proc,step:t.data.step});
    else if(id.startsWith('p:')) setSel({type:'process',key:id.slice(2)});
    else if(id.startsWith('d:')) setSel({type:'domain',key:id.slice(2)});
    else if(t.data&&t.data.proc) setSel({type:'process',key:t.data.proc});
  },[]);

  const onExcel=()=>{
    setBusy(true);
    const stamp=new Date().toISOString().slice(0,10);
    exportExcel(ws,ix,(opt.domain||'domain-models')+'-'+stamp+'.xlsx')
      .then(()=>setNotice('Excel exported.')).catch(e=>setNotice(e.message)).finally(()=>setBusy(false));
  };

  const logical=()=>{
    const r=conceptToLogical(ws,ix,{domain:opt.domain});
    const stamp=new Date().toISOString().slice(0,10);
    return {...r,file:(opt.domain||'domain-models')+'-logical-'+stamp+'.json',label:(opt.domain||'domain-models')+'-logical'};
  };
  const onJson=()=>{ const r=logical();
    download(r.file,JSON.stringify(r.model,null,2),'application/json');
    setNotice('Downloaded '+r.file+': '+r.stats.tables+' tables, '+r.stats.fks+' foreign keys.'); };
  const onLogical=onOpenLogical?()=>{ const r=logical(); onOpenLogical(r.model,r.label); }:null;

  const empty=!Object.values(ix.domains).some(d=>d.modelled);
  const cy=()=>cyApi.current;
  const png=()=>{ const c=cy(); if(!c) return;
    download((opt.domain||'model')+'-'+mode+'.png',c.png({output:'blob',full:true,scale:2,
      bg:getComputedStyle(document.body).getPropertyValue('--bg').trim()||'#fff'})); };
  useEffect(()=>{ if(!notice) return; const t=setTimeout(()=>setNotice(''),5000); return ()=>clearTimeout(t); },[notice]);

  return h('div',{style:{display:'flex',flex:1,overflow:'hidden'}},
    h(Outline,{ws,ix,mode,opt,setOpt,sel,onSelect:setSel,search,setSearch,files,onLoad,onExample,
      onClear:()=>load([],true),serverFiles,onServer:()=>serverFiles&&load(serverFiles,false),
      editing,setEditing,act,undo,redo,canUndo:hist.past.length>0,canRedo:hist.future.length>0,onExcel,busy,onJson,onLogical,onOpenJson}),
    h('div',{style:{flex:1,position:'relative',background:'var(--bg)',
        backgroundImage:'radial-gradient(circle at 1px 1px,var(--line) 1px,transparent 0)',backgroundSize:'40px 40px'}},
      h(CyCanvas,{graph,styles,highlight,selectedId,onTap,apiRef:cyApi,resetKey}),
      !empty&&h('div',{style:{position:'absolute',top:12,right:12,zIndex:10,display:'flex',gap:6}},
        mode==='concepts'&&h('button',{className:'tbtn',title:'Lay the diagram out again',onClick:()=>{ const c=cy(); if(c){ c.layout(graph.layout).run(); c.fit(undefined,50); } }},'Re-layout'),
        h('button',{className:'tbtn',title:'Fit to view',onClick:()=>cy()&&cy().fit(undefined,50)},'Fit'),
        h('button',{className:'tbtn',title:'Zoom out',onClick:()=>cy()&&cy().zoom({level:cy().zoom()*0.8,renderedPosition:{x:cy().width()/2,y:cy().height()/2}})},'−'),
        h('button',{className:'tbtn',title:'Zoom in',onClick:()=>cy()&&cy().zoom({level:cy().zoom()*1.25,renderedPosition:{x:cy().width()/2,y:cy().height()/2}})},'+'),
        h('button',{className:'tbtn',title:'Download the view as PNG',onClick:png},'PNG')),
      notice&&h('div',{style:{position:'absolute',top:12,left:14,zIndex:11,background:'var(--card)',border:'1px solid var(--line)',
        borderRadius:8,padding:'6px 10px',fontSize:12,boxShadow:'var(--shadow)'}},notice),
      !empty&&h(Legend,{mode}),
      empty&&h('div',{style:{position:'absolute',top:0,left:0,right:0,bottom:0,display:'flex',alignItems:'center',justifyContent:'center',zIndex:20}},
        h('div',{style:{textAlign:'center',maxWidth:460,color:'var(--muted)',fontSize:13.5,lineHeight:1.55}},
          h('div',{style:{fontSize:40,marginBottom:10}},'◇'),
          h('p',null,'Open one or more domain models (',h('code',null,'business.yaml'),', ',h('code',null,'processes.yaml'),
            ') to see and edit their concepts and processes, or start a new domain in Edit mode. Nothing is uploaded; the files stay in this browser.'),
          h('div',{style:{display:'flex',gap:8,justifyContent:'center',marginTop:10}},
            h('button',{className:'btn btn-accent',onClick:onExample},'Open the example model')),
          ws.errors.length>0&&h('div',{style:{...S.q,marginTop:14,textAlign:'left'}},ws.errors.join('\n'))))),
    h('div',{style:{...S.right,width:editing?360:300}},h(RightPanel,{ws,ix,sel,setSel,mode,onViewMode,editing,act,opt})));
}

/* Files handed over by the app (its top-bar Open…), delivered to the view when it mounts. */
const inbox={pending:null,listener:null};
function openFiles(list){
  return Promise.all([...list].map(f=>f.text().then(text=>({name:f.name,text})))).then(fs=>{
    if(inbox.listener) inbox.listener(fs); else inbox.pending=(inbox.pending||[]).concat(fs);
    return fs.length;
  });
}

window.ConceptViews={View,wsFromFiles,buildIndex,changes,bumpVersion,workbookRows,moveConcept,moveRelationship,dependencyCycles,
  conceptToLogical,openFiles,
  MODES:[
    {id:'concepts',label:'Concepts',title:'Conceptual model: business concepts, events and their relationships (YAML)',concept:true},
    {id:'processes',label:'Processes',title:'Process model: triggers, steps and the events and entities they produce (YAML)',concept:true},
  ]};
})();
