/* Cloud database shared by all pages. localStorage is used only for explicit
   legacy recovery. IndexedDB is a retry outbox, never a report database. */
(function () {
'use strict';
const keys=['pi_outbound_batches','pi_returns_data','pi_tracking_results','pi_claims_history','pi_returns_archive'];
const page=(location.pathname.split('/').pop()||'index.html').toLowerCase();
const needed=({
 'index.html':[keys[0],'pi_returns_report',keys[3]],
 'pi-report.html':[keys[0],'pi_returns_report',keys[3]],
 'pi-returns.html':['pi_returns_report'],'pi-claims.html':[keys[3]],
 'pi-enter-data.html':keys,'pi-tracking.html':[keys[0],keys[2],keys[3]],
 'pi-archive.html':keys,'pi-archive-outbound.html':[keys[0]],
 'pi-archive-returns.html':[keys[4]],'pi-archive-tracking.html':[keys[2]]
})[page]||[];
let cache={},versions={},ready=false,admin=false,busy=false,draftDB,initialized=false;
let refreshing=false,lastCheckAt=0,dataEpoch=0,interactionEpoch=0;
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
 if(text.includes('DATA_CHANGED'))return 'มีข้อมูลเปลี่ยนจากอีกเครื่อง กรุณาโหลดข้อมูลล่าสุดก่อนทำรายการใหม่ ฉบับร่างยังอยู่ในรายการรอส่ง';
 if(text.includes('claims_web_versions'))return 'กรุณารัน supabase-refresh-upgrade.sql เพื่อเปิดการตรวจข้อมูลใหม่แบบรวดเร็ว';
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
 for(let attempt=0;attempt<3;attempt++){
  let rows=[],version=null,total=0;
  try{
   do{
    const result=await window.claimsSupabase.rpc('claims_web_read',{p_key:key,p_offset:rows.length,p_limit:500,p_revision:version});
    if(result.error)throw result.error;
    const data=result.data;
    if(!data||!Array.isArray(data.rows)||data.revision==null)throw Error('ข้อมูลตอบกลับไม่ครบ');
    version=Number(data.revision);total=Number(data.total);
    if(!data.rows.length&&rows.length<total)throw Error('โหลดข้อมูลได้ไม่ครบ');
    rows.push(...data.rows);
   }while(rows.length<total);
   return {rows,version};
  }catch(e){if(!String(e.message).includes('DATA_CHANGED')||attempt===2)throw e;}
 }
}
async function fetchCollections(selected){
 return Promise.all(selected.map(async key=>[key,await readCollection(key)]));
}
function applyCollections(all){
 all.forEach(([key,data])=>{cache[key]=data.rows;versions[key]=data.version;});
}
async function load(){applyCollections(await fetchCollections(needed));}
async function readVersions(){
 const result=await window.claimsSupabase.rpc('claims_web_versions',{p_keys:needed});
 if(result.error)throw result.error;
 if(!result.data||needed.some(key=>result.data[key]==null||!Number.isSafeInteger(Number(result.data[key]))))throw Error('ข้อมูลเวอร์ชันตอบกลับไม่ครบ');
 return result.data;
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
async function commit(changes,draftId){
 const result=await window.claimsSupabase.rpc('claims_web_save',{p_changes:changes});
 if(result.error)throw result.error;
 for(const key of Object.keys(changes))if(result.data?.[key]===undefined)throw Error('ยังไม่ได้รับผลยืนยันครบ');
 for(const key of Object.keys(changes)){cache[key]=clone(changes[key].rows);versions[key]=Number(result.data[key]);}
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
 try{
  for(const key of Object.keys(values).sort()){
   if(versions[key]===undefined)throw Error('ยังไม่ได้โหลดข้อมูล '+key);
   changes[key]={expected:versions[key],rows:normalize(key,values[key])};
  }
  await outbox('put',draft);
  message('กำลังบันทึกลง Supabase… กรุณารอผลยืนยัน');
  await commit(changes,draft.id);
  renderPage();
 }catch(e){
  renderPage();message('บันทึกไม่สำเร็จ: '+describe(e),true,{text:'รายการรอส่ง / กู้ข้อมูล',run:recovery});
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
     const fresh=await readCollection(key);cache[key]=fresh.rows;versions[key]=fresh.version;
     const current=clone(fresh.rows), seen=new Set(current.map(x=>identity(key,x)));
     normalize(key,record.values[key]).forEach(x=>{const id=identity(key,x);if(!seen.has(id)){seen.add(id);current.push(x);added++;}});
     values[key]=current;
    }
    if(!added){info.textContent='รายการเหล่านี้มีอยู่บน Cloud แล้ว หากเป็นการแก้ไข ให้ตรวจ JSON หรือใช้ส่งฉบับร่างเดิม';return;}
    if(!confirm('นำเข้า '+added+' รายการที่ขาดไปยัง Supabase?'))return;
    busy=false;await saveMany(values);renderPage();info.textContent='บันทึกข้อมูลที่ขาด '+added+' รายการแล้ว';
   }catch(e){info.textContent=describe(e);}finally{busy=false;lock(false);bring.disabled=false;}
  };row.append(bring);
  if(record.draft){
   const retry=document.createElement('button');retry.textContent='ส่งฉบับร่างเดิมอีกครั้ง';
   retry.onclick=async()=>{
    if(busy)return;
    if(!confirm('ส่งรายการแก้ไขในฉบับร่างนี้อีกครั้ง? ระบบจะตรวจว่าข้อมูล Cloud ยังไม่ถูกแก้ไขจากรายการอื่น'))return;
    busy=true;dataEpoch++;lock(true);retry.disabled=true;
    try{await commit(record.draft.changes,record.draft.id);await load();renderPage();info.textContent='ส่งฉบับร่างสำเร็จ';}
    catch(e){info.textContent=describe(e);}finally{busy=false;lock(false);retry.disabled=false;}
   };row.append(retry);
  }
  dialog.append(row);
 });
 const close=document.createElement('button');close.textContent='ปิด';close.onclick=()=>{dialog.close();dialog.remove();};dialog.append(close);
 document.body.append(dialog);dialog.showModal();
}
async function bootstrap(event){
 if(initialized)return;
 admin=!!event.detail?.admin;window.ClaimsDB.admin=admin;
 if(!admin&&!['index.html','pi-report.html','pi-returns.html','pi-claims.html'].includes(page))return;
 lock(true);message('กำลังโหลดข้อมูลจาก Supabase…');
 try{
  await load();ready=true;
  document.querySelectorAll('script[data-cloud-page]').forEach(source=>{
   const script=document.createElement('script');script.textContent=source.textContent;document.body.append(script);
  });
  initialized=true;renderPage();lock(false);message('ข้อมูลจาก Supabase · อัปเดต '+new Date().toLocaleTimeString('th-TH'));
  const controls=document.createElement('div');controls.style.cssText='padding:6px 20px;display:flex;gap:12px';
  const refreshButton=document.createElement('button');refreshButton.textContent='อัปเดตข้อมูล';refreshButton.onclick=()=>refresh(true);controls.append(refreshButton);
  if(admin){const recover=document.createElement('button');recover.textContent='กู้ข้อมูลเก่า / รายการรอส่ง';recover.onclick=recovery;controls.append(recover);}
  document.getElementById('claimsDatabaseStatus').after(controls);
 }catch(e){ready=false;message('เปิดข้อมูลไม่สำเร็จ: '+describe(e),true,{text:'ลองเชื่อมต่ออีกครั้ง',run:()=>bootstrap(event)});}
}
window.ClaimsDB={
 getItem(key){
  if(key===keys[1]&&cache.pi_returns_report){const map=new Map();cache.pi_returns_report.forEach(x=>map.set(identity(keys[1],x),x));return JSON.stringify([...map.values()]);}
  if(!ready||!cache[key])throw Error('ข้อมูลยังไม่พร้อม');return JSON.stringify(cache[key]);
 },
 setItem(key,value){return saveMany({[key]:JSON.parse(value)});},
 read(key){return clone(cache[key]||[]);},saveMany,identity,mergeIncoming,readWorkbook,cleanSku:clean,refresh,recovery,admin:false
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
