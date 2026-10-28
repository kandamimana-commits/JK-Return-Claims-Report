/* Cloud database shared by all pages. localStorage is used only for explicit
   legacy recovery. IndexedDB is a retry outbox, never a report database. */
(function () {
'use strict';
const keys=['pi_outbound_batches','pi_returns_data','pi_tracking_results','pi_claims_history','pi_returns_archive'];
const page=(location.pathname.split('/').pop()||'index.html').toLowerCase();
const needed=({
 'index.html':['pi_report_summary'],
 'pi-report.html':['pi_report_summary'],
 'pi-returns.html':['pi_returns_report','pi_report_summary'],'pi-claims.html':[keys[3]],
 'pi-enter-data.html':[keys[0],keys[1],keys[3],keys[4]],'pi-tracking.html':[keys[0],keys[2],keys[3]],
 'pi-archive.html':['pi_archive_summary'],'pi-archive-outbound.html':[keys[0],keys[4]],
 'pi-archive-returns.html':[keys[4]],'pi-archive-tracking.html':[keys[2]]
})[page]||[];
let cache={},versions={},ready=false,admin=false,busy=false,draftDB,initialized=false;
let refreshing=false,lastCheckAt=0,dataEpoch=0,interactionEpoch=0,reportMonth=null,loading=false;
const deferredCollections=new Set();
let outboundIndex=[];
function calendarMonth(){const p=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Bangkok',year:'numeric',month:'2-digit'}).formatToParts(new Date()).map(x=>[x.type,x.value]));return p.year+'-'+p.month;}
function nextMonth(m){let y=Number(m.slice(0,4)),n=Number(m.slice(5,7))+1;if(n===13){n=1;y++;}return y+'-'+String(n).padStart(2,'0');}
function activeMonth(){return (cache.pi_returns_archive||[]).reduce((m,x)=>/^\d{4}-(0[1-9]|1[0-2])$/.test(x.month||'')&&nextMonth(x.month)>m?nextMonth(x.month):m,calendarMonth());}
let calendarSeen=calendarMonth();
const clone=v=>JSON.parse(JSON.stringify(v));
const norm=v=>String(v||'').replace(/[^A-Za-z0-9]/g,'').toUpperCase();
const clean=v=>String(v||'').replace(/เทเล\s*-/gi,'').replace(/อัพ/g,'').replace(/\s+/g,' ').trim();
function identity(key,x){
 if(key===keys[0])return JSON.stringify([norm(x.tracking),x.date]);
 if(key===keys[1])return JSON.stringify([norm(x.tracking),x.date,clean(x.sku).toUpperCase()||clean(x.product)]);
 if(key===keys[2])return norm(x.tracking);
 if(key===keys[3])return norm(x.tracking)?'claim-'+norm(x.tracking):(x._cloudId||JSON.stringify([x.order,x.dateFiled,x.product]));
 return x.month;
}
function normalize(key,rows){
 if(!Array.isArray(rows))throw Error('รูปแบบข้อมูลไม่ถูกต้อง');
 const result=new Map();
 rows.forEach(raw=>{
  const row=clone(raw);
  if(keys.slice(0,3).includes(key)){row.tracking=norm(row.tracking);if(!row.tracking)throw Error('พบแถวไม่มีเลขพัสดุ');}
  if(keys.slice(0,2).includes(key)&&!/^\d{4}-\d{2}-\d{2}$/.test(row.date||''))throw Error('วันที่ไม่ถูกต้อง');
  if(key===keys[3])row._cloudId=identity(key,row);
  result.set(identity(key,row),row);
 });
 return [...result.values()];
}
function stable(value){
 if(Array.isArray(value))return value.map(stable);
 if(value&&typeof value==='object')return Object.fromEntries(Object.keys(value).sort().map(k=>[k,stable(value[k])]));
 return value;
}
function databaseRows(key,rows){
 const grouped=new Map();
 for(const row of rows){
  const id=key===keys[1]?JSON.stringify([norm(row.tracking),row.date]):key===keys[3]?row._cloudId:identity(key,row);
  if(key===keys[1]){
   if(!grouped.has(id))grouped.set(id,{tracking:norm(row.tracking),date:row.date,lines:[]});
   grouped.get(id).lines.push(row);
  }else grouped.set(id,row);
 }
 if(key===keys[1])for(const row of grouped.values())row.lines.sort((a,b)=>identity(key,a).localeCompare(identity(key,b)));
 return grouped;
}
function buildPatches(changes){
 const patches={};
 for(const [key,change] of Object.entries(changes)){
  const before=databaseRows(key,cache[key]||[]),after=databaseRows(key,change.rows),upsert=[],remove=[];
  for(const [id,row] of after)if(!before.has(id)||JSON.stringify(stable(before.get(id)))!==JSON.stringify(stable(row)))upsert.push(row);
  for(const [id,row] of before)if(!after.has(id))remove.push(key===keys[0]||key===keys[1]?{tracking:row.tracking,date:row.date}:key===keys[2]?{tracking:row.tracking}:key===keys[3]?{_cloudId:row._cloudId}:{month:row.month});
  patches[key]={expected:change.expected,upsert,remove};
 }
 return patches;
}
function mergeIncoming(key,current,incoming){
 const parsed=new Map();
 incoming.forEach(raw=>{
  const row=clone(raw), id=identity(key,row);
  if(key===keys[1]&&parsed.has(id))parsed.get(id).qty=Number(parsed.get(id).qty||0)+Number(row.qty||0);
  else parsed.set(id,row);
 });
 const result=new Map(current.map(x=>[identity(key,x),x]));
 parsed.forEach((x,id)=>result.set(id,x));
 return normalize(key,[...result.values()]);
}
function readWorkbook(buffer,name){
 if(!/\.(csv|txt|tsv)$/i.test(name||''))return window.XLSX.read(buffer,{type:'array',cellDates:true});
 const bytes=new Uint8Array(buffer);let text;
 if(bytes[0]===255&&bytes[1]===254)text=new TextDecoder('utf-16le').decode(bytes);
 else if(bytes[0]===254&&bytes[1]===255)text=new TextDecoder('utf-16be').decode(bytes);
 else {try{text=new TextDecoder('utf-8',{fatal:true}).decode(bytes);}catch(e){text=new TextDecoder('windows-874').decode(bytes);}}
 return window.XLSX.read(text.replace(/^\uFEFF/,''),{type:'string',raw:true,cellDates:true});
}
function message(text,bad=false,action){
 let box=document.getElementById('claimsDatabaseStatus');
 if(!box){box=document.createElement('div');box.id='claimsDatabaseStatus';box.setAttribute('role','status');
  box.style.cssText='padding:12px 72px 12px 20px;background:var(--surface-kraft,#f3e4c8);color:var(--ink,#4a3826);font:14px Sarabun,sans-serif;overflow-wrap:anywhere';document.body.prepend(box);}
 box.replaceChildren(document.createTextNode(text));
 box.style.borderBottom=bad?'2px solid #c94f4f':'1px solid var(--border,#ddd)';
 if(action){const btn=document.createElement('button');btn.textContent=action.text;btn.style.marginLeft='12px';btn.onclick=action.run;box.append(btn);}
}
function lock(on){document.querySelectorAll('main,.wrap').forEach(el=>el.inert=on);}
function describe(error){
 const text=error?.message||'เชื่อมต่อไม่สำเร็จ';
 if(error?.code==='57014'||text.includes('statement timeout'))return 'ฐานข้อมูลยกเลิกคำสั่งเพราะใช้เวลานานเกินกำหนด กรุณาเก็บไฟล์ต้นฉบับและอย่าล้างข้อมูลเบราว์เซอร์';
 if(text.includes('claims_web_page'))return 'กรุณารัน supabase-read-performance-upgrade.sql ในโปรเจ็กต์เว็บเคลมก่อนใช้เว็บรุ่นนี้ (ไม่ต้องรัน migration เก่าซ้ำ)';
 if(text.includes('claims_web_entry_index'))return 'กรุณารัน supabase-read-performance-upgrade.sql ในโปรเจ็กต์เว็บเคลมก่อนใช้เว็บรุ่นนี้';
 if(text.includes('claims_web_monthly_report'))return 'กรุณารัน supabase-monthly-upgrade.sql ในโปรเจ็กต์เว็บเคลมก่อนใช้รายงานรายเดือน';
 if(text.includes('claims_web_apply'))return 'กรุณารัน supabase-upload-delta-upgrade.sql ก่อนอัปโหลดด้วยเว็บรุ่นนี้';
 if(text.includes('DATA_CHANGED'))return 'มีข้อมูลเปลี่ยนจากอีกเครื่อง กรุณาโหลดข้อมูลล่าสุดก่อนทำรายการใหม่ ฉบับร่างยังอยู่ในรายการรอส่ง';
 if(text.includes('claims_web_snapshot'))return 'กรุณารัน supabase-initial-load-upgrade.sql เพื่อเปิดระบบโหลดหน้าแบบใหม่';
 if(text.includes('claims_web_versions'))return 'กรุณารัน supabase-initial-load-upgrade.sql เพื่อเปิดการตรวจข้อมูลใหม่แบบรวดเร็ว';
 if(error?.code==='PGRST202'||text.includes('claims_web_'))return 'ฐานข้อมูลยังไม่พร้อม ให้ผู้ดูแลรัน supabase-migration.sql ก่อน';
 return text;
}
function openOutbox(){
 if(!draftDB)draftDB=new Promise((resolve,reject)=>{
  const req=indexedDB.open('jk888-cloud-recovery',1);
  req.onupgradeneeded=()=>req.result.createObjectStore('drafts',{keyPath:'id'});
  req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);
 });
 return draftDB;
}
async function outbox(method,value){
 const db=await openOutbox();
 return new Promise((resolve,reject)=>{
  const tx=db.transaction('drafts',method==='getAll'?'readonly':'readwrite');
  const req=tx.objectStore('drafts')[method](value);
  tx.oncomplete=()=>resolve(req.result);tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error);
 });
}
async function readCollection(key){
 return (await fetchCollections([key],true))[0][1];
}
async function fetchCollections(selected,full=false){
 if(!selected.length)return [];
 // Reports already return a single server-side aggregate, never raw full-history rows.
 if(selected.length===1&&selected[0]==='pi_report_summary'){
  const data=await requestRead('claims_web_monthly_report',{p_month:reportMonth});
  validateSummary(data,'pi_report_summary');
  return [['pi_report_summary',{rows:data.rows,version:Number(data.revision)}]];
 }
 // Stage every collection. A failed/changed middle page must NEVER become a writable cache.
 for(let attempt=0;attempt<2;attempt++){
  try{
   const expected=await readVersions(selected), all=[];
   const lazy=page==='pi-enter-data.html'&&!full?selected.filter(k=>[keys[0],keys[4]].includes(k)&&(!cache[k]||deferredCollections.has(k))):[];
   const index=lazy.length?await requestRead('claims_web_entry_index',{}):null;
   for(const key of selected){
    let data;
    if(lazy.includes(key)){
     const rows=key===keys[0]?[]:index?.archives;
     if(!Array.isArray(rows)||!Array.isArray(index?.outbound))throw Error('ข้อมูลสรุปหน้าอัปโหลดไม่ครบ');
     data={rows,revision:index.versions?.[key],deferred:true,index:key===keys[0]?index.outbound:undefined};
    }else if(key==='pi_report_summary'){
     data=await requestRead('claims_web_monthly_report',{p_month:reportMonth});validateSummary(data,key);
    }else if(key==='pi_archive_summary'){
     const result=await requestRead('claims_web_snapshot',{p_keys:[key]});data=result?.[key];validateSummary(data,key);
    }else data=await readPages(key,Number(expected[key]));
    if(Number(data.revision)!==Number(expected[key]))throw Error('DATA_CHANGED');
    all.push([key,{rows:data.rows,version:Number(data.revision),deferred:!!data.deferred,index:data.index}]);
   }
   const latest=await readVersions(selected);
   if(selected.some(key=>Number(latest[key])!==Number(expected[key])))throw Error('DATA_CHANGED');
   return all;
  }catch(e){if(attempt===0&&String(e.message).includes('DATA_CHANGED'))continue;throw e;}
 }
}
function validateSummary(data,key){
 if(!data||!Array.isArray(data.rows)||data.rows.length!==1||Number(data.total)!==1||data.revision==null||!Number.isSafeInteger(Number(data.revision)))throw Error('โหลดข้อมูลสรุปได้ไม่ครบ: '+key);
}
async function requestRead(name,params){
 const controller=new AbortController();
 const timer=setTimeout(()=>controller.abort(),20000);
 try{
  let query=window.claimsSupabase.rpc(name,params);
  if(typeof query.abortSignal==='function')query=query.abortSignal(controller.signal);
  const result=await query;
  if(result.error)throw result.error;
  return result.data;
 }finally{clearTimeout(timer);}
}
async function readPages(key,revision){
 const rows=[],archives=new Map(),seenCursors=new Set();
 // Claim records can contain evidence images; keep their individual requests smaller.
 let cursor=null,limit=key==='pi_claims_history'?25:250,readCount=0;
 for(;;){
  let data;
  for(let attempt=0;;attempt++){
   try{data=await requestRead('claims_web_page',{p_key:key,p_after:cursor,p_limit:limit,p_revision:revision});break;}
   catch(e){
    const text=String(e.message||'');
    if(attempt>=2||!(e.code==='57014'||/timeout|fetch|abort|network/i.test(text)))throw e;
    limit=Math.max(5,Math.floor(limit/2));
   }
  }
  if(!data||!Array.isArray(data.rows)||data.revision==null||Number(data.revision)!==revision||!Object.hasOwn(data,'next'))throw Error('โหลดข้อมูลได้ไม่ครบ: '+key);
  if(key==='pi_returns_archive'&&data.fragment){
   const f=data.fragment;
   if(typeof f.key!=='string'||!Number.isSafeInteger(f.offset)||!Number.isSafeInteger(f.total)||f.total<0||!f.meta||f.meta.month!==f.key)throw Error('รูปแบบคลังย้อนหลังไม่ครบ');
   if(!archives.has(f.key)){
    if(f.offset!==0)throw Error('คลังย้อนหลังขาดช่วง');
    const record={...f.meta,rows:[]};archives.set(f.key,{record,total:f.total});rows.push(record);
   }
   const saved=archives.get(f.key);
   if(saved.total!==f.total||saved.record.rows.length!==f.offset||f.offset+data.rows.length>f.total)throw Error('คลังย้อนหลังขาดช่วงหรือซ้ำ');
   saved.record.rows.push(...data.rows);
  }else{
   if(key==='pi_returns_archive'&&data.rows.length)throw Error('ขาดข้อมูลกำกับคลังย้อนหลัง');
   for(const row of data.rows)rows.push(row);
  }
  readCount+=data.rows.length;
  if(!ready)message('กำลังโหลดจาก Supabase เป็นชุด · '+key+' · รับแล้ว '+readCount.toLocaleString()+' รายการ');
  if(data.next===null)break;
  if(typeof data.next!=='object'||Array.isArray(data.next))throw Error('ตำแหน่งโหลดข้อมูลไม่ถูกต้อง');
  const token=JSON.stringify(data.next);
  if(seenCursors.has(token))throw Error('ตำแหน่งโหลดข้อมูลซ้ำ หยุดเพื่อป้องกันข้อมูลไม่ครบ');
  seenCursors.add(token);cursor=data.next;
 }
 for(const {record,total} of archives.values())if(record.rows.length!==total)throw Error('คลังย้อนหลังโหลดได้ไม่ครบ');
 return {rows,revision};
}
function applyCollections(all){
 all.forEach(([key,data])=>{
  cache[key]=data.rows;versions[key]=data.version;
  if(data.deferred){deferredCollections.add(key);if(key===keys[0])outboundIndex=data.index;}
  else deferredCollections.delete(key);
 });
}
async function ensureCollections(selected){
 const missing=selected.filter(k=>!cache[k]||deferredCollections.has(k));
 if(!missing.length)return;
 if(busy||loading)throw Error('กำลังทำรายการอื่น กรุณารอให้เสร็จก่อน');
 busy=true;dataEpoch++;lock(true);message('กำลังโหลดรายละเอียดที่จำเป็นก่อนทำรายการ…');
 try{applyCollections(await fetchCollections(missing,true));renderPage();message('โหลดรายละเอียดครบแล้ว พร้อมทำรายการ');}
 catch(e){message('โหลดรายละเอียดไม่สำเร็จ — ยังไม่ได้บันทึกหรือปิดเดือน: '+describe(e),true);throw e;}
 finally{busy=false;lock(false);}
}
async function load(){applyCollections(await fetchCollections(needed));}
async function readVersions(selected=needed){
 const data=await requestRead('claims_web_versions',{p_keys:selected});
 if(!data||selected.some(key=>data[key]==null||!Number.isSafeInteger(Number(data[key]))))throw Error('ข้อมูลเวอร์ชันตอบกลับไม่ครบ');
 return data;
}
function editing(){return /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement?.tagName||'')||!!document.activeElement?.isContentEditable;}
function renderPage(){
 window.renderCloudPage?.();
 if(page==='pi-claims.html'&&!admin){
  document.querySelectorAll('.claimEdit,.evUploadInput,.evRemoveBtn,[data-delclaim],[data-evtoggle],#syncClaimsBtn').forEach(el=>{el.disabled=true;});
 }
}
async function refresh(force=false){
 if(!ready||busy||refreshing)return;
 // Focus and visibility events often arrive together. Coalesce their cheap checks.
 if(force!==true&&(editing()||Date.now()-lastCheckAt<5000))return;
 refreshing=true;lastCheckAt=Date.now();
 const startData=dataEpoch,startInteraction=interactionEpoch;
 const interrupted=()=>busy||dataEpoch!==startData||interactionEpoch!==startInteraction||editing();
 try{
  const latest=await readVersions();
  if(interrupted())return;
  const changed=needed.filter(key=>Number(latest[key])!==versions[key]);
  if(!changed.length){
   if(calendarSeen!==calendarMonth()){calendarSeen=calendarMonth();renderPage();}
   if(force===true)message('ข้อมูลเป็นเวอร์ชันล่าสุดแล้ว · ไม่ต้องโหลดซ้ำ');
   return;
  }
  // Stage the result; never overwrite a concurrent save or a user's active form.
  const fresh=await fetchCollections(changed);
  if(interrupted())return;
  applyCollections(fresh);dataEpoch++;renderPage();
  message('ข้อมูลจาก Supabase · อัปเดตเฉพาะส่วนที่เปลี่ยน '+new Date().toLocaleTimeString('th-TH'));
 }catch(e){
  if(!interrupted())message('ตรวจข้อมูลใหม่ไม่สำเร็จ — ยังแสดงข้อมูลครั้งล่าสุด: '+describe(e),true,{text:'ลองใหม่',run:()=>refresh(true)});
 }finally{refreshing=false;}
}
async function setReportMonth(month){
 if(!ready||busy)return;
 if(month!==null&&!/^\d{4}-(0[1-9]|1[0-2])$/.test(month))return;
 const previous=reportMonth;reportMonth=month;busy=true;dataEpoch++;lock(true);
 message('กำลังโหลดรายงานเดือนที่เลือก…');
 try{await load();renderPage();message('รายงานรายเดือนจาก Supabase · อัปเดตแล้ว');}
 catch(e){reportMonth=previous;renderPage();message('โหลดเดือนที่เลือกไม่สำเร็จ: '+describe(e),true);}
 finally{busy=false;lock(false);}
}
async function commit(changes,draftId,patches){
 if(!patches)throw Error('ฉบับร่างรุ่นเก่า: ใช้นำเข้าเฉพาะที่ขาด หรือสำรอง JSON ก่อนอัปไฟล์ต้นฉบับใหม่');
 const result=await window.claimsSupabase.rpc('claims_web_apply',{p_request_id:draftId,p_changes:patches});
 if(result.error)throw result.error;
 for(const key of Object.keys(changes))if(result.data?.versions?.[key]===undefined)throw Error('ยังไม่ได้รับผลยืนยันครบ');
 for(const key of Object.keys(changes)){cache[key]=clone(changes[key].rows);versions[key]=Number(result.data.versions[key]);deferredCollections.delete(key);}
 await outbox('delete',draftId).catch(()=>{});
 message('บันทึกลง Supabase สำเร็จ · '+new Date().toLocaleTimeString('th-TH'));
 window.dispatchEvent(new CustomEvent('claims-cloud-saved'));
 if(window.BroadcastChannel){const channel=new BroadcastChannel('jk888-cloud');channel.postMessage('saved');channel.close();}
}
async function saveMany(values){
 if(!ready||!admin)throw Error('กรุณาล็อกอินด้วยบัญชีแอดมินก่อนบันทึก');
 if(busy)throw Error('กำลังทำรายการอื่น กรุณารอให้เสร็จก่อน');
 busy=true;dataEpoch++;lock(true);
 const changes={},draft={id:crypto.randomUUID(),createdAt:new Date().toISOString(),changes};
 let draftSaved=false;
 try{
  for(const key of Object.keys(values).sort()){
   if(deferredCollections.has(key))throw Error('ยังไม่ได้โหลดรายละเอียดครบ ห้ามบันทึกทับจากข้อมูลสรุป: '+key);
   if(versions[key]===undefined)throw Error('ยังไม่ได้โหลดข้อมูล '+key);
   changes[key]={expected:versions[key],rows:normalize(key,values[key])};
  }
  draft.patches=buildPatches(changes);
  await outbox('put',draft);
  draftSaved=true;
  const count=Object.values(draft.patches).reduce((n,p)=>n+p.upsert.length+p.remove.length,0);
  message('กำลังบันทึกเฉพาะรายการที่เปลี่ยน '+count.toLocaleString()+' รายการ… กรุณารอผลยืนยัน');
  await commit(changes,draft.id,draft.patches);
  renderPage();
 }catch(e){
  renderPage();message('บันทึกไม่สำเร็จ: '+describe(e)+(draftSaved?' · เก็บฉบับร่างในรายการรอส่งแล้ว':''),true,{text:'รายการรอส่ง / กู้ข้อมูล',run:recovery});
  throw e;
 }finally{busy=false;lock(false);}
}
function download(value,name){
 const url=URL.createObjectURL(new Blob([JSON.stringify(value,null,2)],{type:'application/json'}));
 const link=document.createElement('a');link.href=url;link.download=name;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
async function recovery(){
 if(!admin||busy)return;
 const dialog=document.createElement('dialog');dialog.style.cssText='max-width:720px;width:90%;padding:24px;border-radius:16px';
 const title=document.createElement('h3');title.textContent='กู้ข้อมูลเก่า / รายการรอส่ง';dialog.append(title);
 const info=document.createElement('p');info.textContent='นำเข้าเฉพาะรายการที่ยังไม่มีบน Cloud โดยเก็บข้อมูลเดิมในเครื่องไว้';dialog.append(info);
 const records=[];
 keys.forEach(key=>['localStorage','sessionStorage'].forEach(storage=>{
  try{const rows=JSON.parse(window[storage].getItem(key)||'[]');if(Array.isArray(rows)&&rows.length)records.push({label:key+' ใน '+storage,values:{[key]:rows}});}catch(e){}
 }));
 try{(await outbox('getAll')).forEach(draft=>{
  const values={};Object.keys(draft.changes).forEach(key=>values[key]=draft.changes[key].rows);
  records.push({label:'ฉบับร่าง '+draft.createdAt,values,draft});
 });}catch(e){info.textContent+=' · อ่านฉบับร่างไม่ได้: '+e.message;}
 if(!records.length){const p=document.createElement('p');p.textContent='ไม่พบข้อมูลเก่าหรือฉบับร่างในเบราว์เซอร์นี้';dialog.append(p);}
 records.forEach(record=>{
  const row=document.createElement('p');
  row.textContent=record.label+' · '+Object.values(record.values).reduce((n,list)=>n+list.length,0).toLocaleString()+' รายการ ';
  const backup=document.createElement('button');backup.textContent='สำรอง JSON';backup.onclick=()=>download(record.values,'jk888-recovery.json');row.append(backup);
  const bring=document.createElement('button');bring.textContent='นำเข้าเฉพาะที่ขาด';
  bring.onclick=async()=>{
   if(busy)return;bring.disabled=true;busy=true;dataEpoch++;lock(true);
   try{
    const values={};let added=0;
    for(const key of Object.keys(record.values)){
     const fresh=await readCollection(key);applyCollections([[key,fresh]]);
     const current=clone(fresh.rows), seen=new Set(current.map(x=>identity(key,x)));
     normalize(key,record.values[key]).forEach(x=>{const id=identity(key,x);if(!seen.has(id)){seen.add(id);current.push(x);added++;}});
     values[key]=current;
    }
    if(!added){info.textContent='รายการเหล่านี้มีอยู่บน Cloud แล้ว หากเป็นการแก้ไข ให้ตรวจ JSON หรือใช้ส่งฉบับร่างเดิม';return;}
    if(!confirm('นำเข้า '+added+' รายการที่ขาดไปยัง Supabase?'))return;
    busy=false;await saveMany(values);renderPage();info.textContent='บันทึกข้อมูลที่ขาด '+added+' รายการแล้ว';
   }catch(e){info.textContent=describe(e);}finally{busy=false;lock(false);bring.disabled=false;}
  };row.append(bring);
  if(record.draft?.patches){
   const retry=document.createElement('button');retry.textContent='ส่งฉบับร่างเดิมอีกครั้ง';
   retry.onclick=async()=>{
    if(busy)return;
    if(!confirm('ส่งรายการแก้ไขในฉบับร่างนี้อีกครั้ง? ระบบจะตรวจว่าข้อมูล Cloud ยังไม่ถูกแก้ไขจากรายการอื่น'))return;
    busy=true;dataEpoch++;lock(true);retry.disabled=true;
    try{await commit(record.draft.changes,record.draft.id,record.draft.patches);await load();renderPage();info.textContent='ส่งฉบับร่างสำเร็จ';}
    catch(e){info.textContent=describe(e);}finally{busy=false;lock(false);retry.disabled=false;}
   };row.append(retry);
  }
  dialog.append(row);
 });
 const close=document.createElement('button');close.textContent='ปิด';close.onclick=()=>{dialog.close();dialog.remove();};dialog.append(close);
 document.body.append(dialog);dialog.showModal();
}
async function bootstrap(event){
 if(initialized||loading)return;
 loading=true;
 const started=performance.now();
 admin=!!event.detail?.admin;window.ClaimsDB.admin=admin;
 if(!admin&&!['index.html','pi-report.html','pi-returns.html','pi-claims.html'].includes(page)){loading=false;return;}
 lock(true);message('กำลังโหลดข้อมูลจาก Supabase…');
 try{
  await load();ready=true;
  document.querySelectorAll('script[data-cloud-page]').forEach(source=>{
   const script=document.createElement('script');script.textContent=source.textContent;document.body.append(script);
  });
  initialized=true;renderPage();lock(false);message('ข้อมูลจาก Supabase · อัปเดต '+new Date().toLocaleTimeString('th-TH')+' · โหลด '+((performance.now()-started)/1000).toFixed(1)+' วินาที');
  const controls=document.createElement('div');controls.style.cssText='padding:6px 20px;display:flex;gap:12px';
  const refreshButton=document.createElement('button');refreshButton.textContent='อัปเดตข้อมูล';refreshButton.onclick=()=>refresh(true);controls.append(refreshButton);
  if(admin){const recover=document.createElement('button');recover.textContent='กู้ข้อมูลเก่า / รายการรอส่ง';recover.onclick=recovery;controls.append(recover);}
  document.getElementById('claimsDatabaseStatus').after(controls);
 }catch(e){ready=false;message('เปิดข้อมูลไม่สำเร็จ: '+describe(e),true,{text:'ลองเชื่อมต่ออีกครั้ง',run:()=>bootstrap(event)});}
 finally{loading=false;}
}
window.ClaimsDB={
 getItem(key){
  if(deferredCollections.has(key))throw Error('ยังไม่ได้โหลดรายละเอียด: '+key);
  if(key===keys[1]&&cache.pi_returns_report){const map=new Map();cache.pi_returns_report.forEach(x=>map.set(identity(keys[1],x),x));return JSON.stringify([...map.values()]);}
  if(!ready||!cache[key])throw Error('ข้อมูลยังไม่พร้อม');return JSON.stringify(cache[key]);
 },
 setItem(key,value){return saveMany({[key]:JSON.parse(value)});},
 read(key){return clone(cache[key]||[]);},isDeferred:key=>deferredCollections.has(key),outboundGroups:()=>clone(outboundIndex),
 ensureCollections,calendarMonth,activeMonth,nextMonth,setReportMonth,saveMany,identity,mergeIncoming,readWorkbook,cleanSku:clean,refresh,recovery,admin:false
};
window.addEventListener('claims-auth-ready',bootstrap);
window.addEventListener('unhandledrejection',event=>{
 message('ทำรายการไม่สำเร็จ: '+describe(event.reason),true,{text:'รายการรอส่ง',run:recovery});event.preventDefault();
});
window.addEventListener('beforeunload',event=>{if(busy){event.preventDefault();event.returnValue='';}});
// Activity after a background request starts makes its result unsafe to apply.
['input','change','focusin','pointerdown','drop'].forEach(type=>document.addEventListener(type,()=>{interactionEpoch++;},true));
window.addEventListener('focus',()=>refresh());
document.addEventListener('visibilitychange',()=>{if(!document.hidden)refresh();});
if(window.BroadcastChannel){const updates=new BroadcastChannel('jk888-cloud');updates.onmessage=()=>{lastCheckAt=0;refresh();};}
setInterval(()=>{if(!document.hidden)refresh();},60000);
})();
