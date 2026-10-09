"use strict";
const KEY='tcs_queue', BATCH=200;  // หน้าเดียวเลื่อนยาว: วาดทีละ 200 แถว แล้วเติมเองเมื่อเลื่อนใกล้ล่างสุด
let rows=[], isSample=false, fileName='', filter='all', shown=BATCH, cur=[];
const $=id=>document.getElementById(id);
const clean=s=>String(s??'').replace(/[ \s]+/g,' ').trim();
const complete=r=>!!(r.n&&r.a&&r.t);
let seq=0; const newId=()=>Date.now().toString(36)+'-'+(seq++).toString(36);  // เลขประจำแถว (รหัสซ้ำได้ เลยใช้ id แยกแต่ละแถว)

function load(){
  try{const d=JSON.parse(localStorage.getItem(KEY)||'null');
    if(d&&Array.isArray(d.rows)){rows=d.rows;rows.forEach(r=>{if(!r.id)r.id=newId()});isSample=false;fileName=d.file||'';return}}catch(e){}
  rows=[];
}
function save(){
  if(isSample) return;
  try{localStorage.setItem(KEY,JSON.stringify({file:fileName,rows}))}catch(e){toast('บันทึกในเบราว์เซอร์ไม่ได้ พื้นที่อาจเต็ม')}
}
function setSrc(){ if(!isSample) $('src').textContent=rows.length?`ข้อมูลจากไฟล์ ${fileName} · เก็บไว้ในเบราว์เซอร์เครื่องนี้ ปิดเว็บแล้วเปิดใหม่ยังอยู่`:'ยังไม่มีข้อมูล ลากไฟล์ export มาวางได้เลย'; }

// อ่านไฟล์ export: แถวที่คอลัมน์ D เป็นรหัส = เริ่มรายใหม่ / ป้ายในคอลัมน์ F บอกว่าค่าในคอลัมน์ G คืออะไร
// ที่อยู่ยาว ๆ ในไฟล์ export มีคำว่า "เงื่อนไขการชำระเงิน" แทรก → ไม่ตัดทิ้ง แค่ขึ้นสีเหลืองให้คนตรวจลบเอง
const STRAY=(w=>new RegExp([...w].map((_,i)=>w.slice(0,w.length-i)).filter(x=>x.length>=5).join('|'),'g'))('เงื่อนไขการชำระเงิน');
// ธงเหลืองที่อยู่ (ตรวจสดทุกครั้งที่พิมพ์ แก้แล้วหายเอง): 'stray' = มีคำ/เศษคำ "เงื่อนไขการชำระเงิน" ปน (แม้แค่ เ)
// 'label' = ท้ายที่อยู่มีแค่คำนำหน้า เช่น จ. อ. ต. แล้วไม่มีชื่อต่อ / 'zip' = ไม่มีรหัสไปรษณีย์ 5 หลัก
const noZip=a=>!!a&&!/(^|\D)\d{5}(\D|$)/.test(a);
const endLabel=a=>/(^|\s|\d)(จ\.|จังหวัด|อ\.|อำเภอ|ต\.|ตำบล|เขต|แขวง|ถ\.|ถนน|ซ\.|ซอย|ม\.|หมู่ที่|หมู่)\s*$/.test(a);
const addrFlag=r=>!r.a?'':hasStray(r.a)?'stray':endLabel(r.a)?'label':noZip(r.a)?'zip':'';
const hasStray=a=>{a=String(a??'');const h=STRAY.test(a);STRAY.lastIndex=0;return h||/(^|[\d\sก-๙])(เงื่|เงื|เง|เ)$/.test(a.trim())};
function parse(aoa){
  const out=[]; let cur=null;
  for(const r of aoa){
    const d=clean(r[3]), f=clean(r[5]);
    if(/^[A-Z]{2,6}(-[A-Z0-9]+)+$/.test(d)){cur={c:d,n:clean(r[4]),a:'',t:'',s:'draft'};out.push(cur)}
    if(!cur||!f) continue;
    const g=r[6];
    if(f==='ที่อยู่'&&!cur.a)cur.a=clean(g)
    else if(f.startsWith('เลขที่ประจำตัวผู้เสียภาษี')&&!cur.t){
      let t=typeof g==='number'?Math.round(g).toString():clean(g).replace(/\D/g,'');
      if(t.length===12){t='0'+t;cur.pad=1}
      cur.t=t;
    }
  }
  return out;
}
// อ่านไฟล์ที่เป็นตารางอยู่แล้ว (มีแถวหัวตาราง เช่น รหัส | ชื่อ-นามสกุล | ที่อยู่ | เลขประจำตัว หรือชีต Vendor ของ TCS)
function headerMap(r){
  const m={};
  r.forEach((v,j)=>{const h=clean(v).toLowerCase();if(!h)return;
    if(m.t==null&&(h.includes('ประจำตัว')||h.includes('ภาษี')||h.startsWith('vat')||h.includes('tax')))m.t=j;
    else if(m.c==null&&(h.startsWith('รหัส')||h==='no.'||h==='code'))m.c=j;
    else if(m.n==null&&(h.startsWith('ชื่อ')||h==='name'))m.n=j;
    else if(m.a==null&&(h.startsWith('ที่อยู่')||h==='address'))m.a=j;});
  return m;
}
function parseTable(aoa){
  let hi=-1,map=null;
  for(let i=0;i<Math.min(aoa.length,15);i++){const m=headerMap(aoa[i]);if(m.c!=null&&[m.n,m.a,m.t].filter(x=>x!=null).length>=2){hi=i;map=m}}
  if(hi<0) return [];
  const out=[];
  for(const r of aoa.slice(hi+1)){
    const c=clean(r[map.c]);if(!c)continue;
    const rec={c,n:map.n!=null?clean(r[map.n]):'',a:'',t:'',s:'draft'};
    if(map.a!=null)rec.a=clean(r[map.a])
    if(map.t!=null){const g=r[map.t];let t=typeof g==='number'?Math.round(g).toString():clean(g).replace(/\D/g,'');if(t.length===12){t='0'+t;rec.pad=1}rec.t=t}
    out.push(rec);
  }
  return out;
}
function handle(file){
  const fr=new FileReader();
  fr.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      // ลองทุกชีต ใช้ชีตแรกที่อ่านข้อมูลได้ (แบบรายงาน เช่น 2F ก่อน แล้วค่อยแบบตาราง)
      let got=[];
      for(const name of wb.SheetNames){
        const ws=wb.Sheets[name]; if(!ws['!ref']) continue;
        const rg=XLSX.utils.decode_range(ws['!ref']); rg.s.c=0; rg.s.r=0;
        const aoa=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:'',range:rg});
        got=parse(aoa); if(!got.length) got=parseTable(aoa);
        if(got.length) break;
      }
      got.forEach(r=>r.id=newId());
      if(!got.length){toast('ไม่พบข้อมูลในไฟล์นี้ ต้องเป็นไฟล์ export แบบ 2F หรือตารางที่มีหัวคอลัมน์ รหัส / ชื่อ / ที่อยู่ / เลขประจำตัว');return}
      const apply=()=>{
        rows=got; isSample=false; fileName=file.name; filter='all'; shown=BATCH;
        save(); setSrc(); fillPrefix(); render(); toast(`เรียงข้อมูลแล้ว ${got.length.toLocaleString()} ราย ✨`);
        const b=$('drop').getBoundingClientRect(); sparkBurst(b.left+b.width/2,b.top+b.height/2);
      };
      const pending=isSample?0:rows.filter(r=>r.s!=='done').length;
      if(pending) ask({ico:'📂',title:`แทนที่ด้วยไฟล์ ${file.name}?`,
        text:`ตารางตอนนี้ยังมี ${pending.toLocaleString()} แถวที่ยังไม่ลง Excel ถ้าแทนที่ แถวเหล่านั้นจะหายไป แนะนำให้ส่งให้น้องบอทก่อน`,ok:'แทนที่เลย'},apply);
      else apply();
    }catch(err){toast('อ่านไฟล์ไม่ได้ ต้องเป็นไฟล์ .xlsx หรือ .xls')}
  };
  fr.readAsArrayBuffer(file);
}

function dupSet(){const m={},s=new Set();rows.forEach(r=>{m[r.c]=(m[r.c]||0)+1});for(const k in m)if(m[k]>1)s.add(k);return s}
function view(){
  const q=clean($('q').value).toLowerCase(), p=$('prefix').value, dups=dupSet();
  return rows.map((r,i)=>[r,i]).filter(([r])=>{
    if(p&&r.c.split('-')[0]!==p) return false;
    if(q&&!(r.c.toLowerCase().includes(q)||r.n.toLowerCase().includes(q))) return false;
    if(filter==='empty') return !complete(r);
    if(filter==='pad') return r.pad;
    if(filter==='dup') return dups.has(r.c);
    if(filter==='cut') return addrFlag(r);
    if(filter==='fail') return r.s==='fail';
    if(filter==='done') return r.s==='done';
    return true;
  });
}
function chips(){
  const dups=dupSet(), c={all:rows.length,empty:rows.filter(r=>!complete(r)).length,cut:rows.filter(r=>addrFlag(r)).length,pad:rows.filter(r=>r.pad).length,
    dup:rows.filter(r=>dups.has(r.c)).length,fail:rows.filter(r=>r.s==='fail').length,done:rows.filter(r=>r.s==='done').length};
  const def=[['all','ทั้งหมด',''],['empty','มีช่องว่าง','empty'],['cut','ที่อยู่ต้องตรวจ','warn'],['pad','เติม 0 แล้ว','warn'],['dup','รหัสซ้ำ','dup'],['fail','ลงไม่สำเร็จ','empty']];
  $('chips').innerHTML=def.filter(([k])=>k==='all'||c[k]||filter===k).map(([k,l,cl])=>`<button type="button" class="chip ${cl}" data-f="${k}" aria-pressed="${filter===k}"><span class="n">${c[k].toLocaleString()}</span>${l}</button>`).join('');
  const waiting=rows.filter(r=>r.s==='draft'||r.s==='fail').length, ready=rows.filter(r=>r.s==='ready').length;
  $('footmsg').textContent=
    [waiting&&`ยังไม่ได้ส่ง ${waiting.toLocaleString()} ราย`,ready&&`รอน้องบอท ${ready.toLocaleString()} ราย`].filter(Boolean).join(' · ')||(rows.length?'ลง Excel ครบทุกรายแล้ว ✓':'');
}
const esc=s=>String(s).replace(/[&<>"]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[ch]));
let dupsNow=new Set();
function tagsHtml(r){const f=addrFlag(r);return (r.s==='done'?'<span class="tag done">✓ ลง Excel แล้ว</span>':r.s==='ready'?'<span class="tag ready">รอน้องบอท</span>':r.s==='fail'?`<span class="tag fail">⚠️ ลงไม่สำเร็จ${r.msg?': '+esc(r.msg):''}</span>`:'')
  +(f==='stray'?'<span class="tag cut">มีคำ "เงื่อนไข…" ปน · ลบออกเอง</span>':f==='label'?'<span class="tag cut">ท้ายที่อยู่ขาด (เช่น จ. แล้วไม่มีชื่อ)</span>':f==='zip'?'<span class="tag cut">ขาดรหัสไปรษณีย์</span>':'')
  +(r.pad?'<span class="tag pad">เติม 0</span>':'')+(dupsNow.has(r.c)?'<span class="tag dup">รหัสซ้ำ · ลง Excel เป็นแถวใหม่</span>':'')}

const rowHtml=([r,i])=>`<tr>
    <td class="code${dupsNow.has(r.c)?' dup':''}">${esc(r.c)}</td>
    <td><textarea id="n${i}" data-i="${i}" data-k="n" rows="1" placeholder="ว่าง">${esc(r.n)}</textarea></td>
    <td><textarea id="a${i}" data-i="${i}" data-k="a" rows="1" class="${addrFlag(r)?'cut':''}" placeholder="ว่าง">${esc(r.a)}</textarea></td>
    <td><input id="t${i}" data-i="${i}" data-k="t" class="tax${r.t&&r.t.length!==13?' badlen':''}" inputmode="numeric" maxlength="13" value="${esc(r.t)}" placeholder="ว่าง"></td>
    <td><div class="tags">${tagsHtml(r)}</div></td>
    
    <td class="act"><button type="button" class="del" data-del="${i}" aria-label="ลบแถว ${esc(r.c)}">✕</button></td>
  </tr>`;
function render(){
  chips();
  cur=view(); dupsNow=dupSet(); shown=Math.min(Math.max(shown,BATCH),Math.max(cur.length,BATCH));
  $('rows').innerHTML=cur.length?cur.slice(0,shown).map(rowHtml).join(''):`<tr><td colspan="6" class="none">${rows.length?'ไม่มีแถวที่ตรงกับตัวกรองนี้':'ยังไม่มีข้อมูล ลากไฟล์ export มาวางด้านบนได้เลย'}</td></tr>`;
  fitAll(); count();
}
function count(){$('pager').textContent=cur.length?`แสดง ${Math.min(shown,cur.length).toLocaleString()} จาก ${cur.length.toLocaleString()} แถว`:''}
function reset(){shown=BATCH;render();document.querySelector('.tablebox').scrollTop=0}
// เลื่อนใกล้ล่างสุด → เติมอีก 200 แถว (ไม่ต้องกดหน้าถัดไป)
function more(){
  if(shown>=cur.length) return;
  const add=cur.slice(shown,shown+BATCH); shown+=add.length;
  $('rows').insertAdjacentHTML('beforeend',add.map(rowHtml).join('')); fitAll(); count();
}
document.querySelector('.tablebox').addEventListener('scroll',e=>{const t=e.currentTarget;if(t.scrollTop+t.clientHeight>t.scrollHeight-600)more()},{passive:true});
// เบราว์เซอร์ใหม่ (Chrome/Edge) ขยายช่องเองด้วย CSS field-sizing ไม่ต้องคำนวณ → เปลี่ยนหน้าเร็ว
const AUTO_SIZE=!!(window.CSS&&CSS.supports&&CSS.supports('field-sizing','content'));
function fit(el){if(AUTO_SIZE)return;el.style.height='auto';el.style.height=el.scrollHeight+2+'px'}
function fitAll(){if(AUTO_SIZE)return;const els=[...document.querySelectorAll('#rows textarea')];els.forEach(e=>e.style.height='auto');const h=els.map(e=>e.scrollHeight);els.forEach((e,i)=>e.style.height=h[i]+2+'px')}
function fillPrefix(){
  const ps=[...new Set(rows.map(r=>r.c.split('-')[0]))].sort();
  $('prefix').innerHTML='<option value="">ทุกกลุ่มรหัส</option>'+ps.map(p=>`<option>${esc(p)}</option>`).join('');
}
let tt; function toast(m){const t=$('toast');t.textContent=m;t.classList.add('show');clearTimeout(tt);tt=setTimeout(()=>t.classList.remove('show'),2800)}

// ป๊อปอัปยืนยันในหน้าเว็บ
let onOk=null;
function ask({ico='⚠️',title,text,ok='ตกลง',cancel='ยกเลิก'},fn){
  $('mIco').textContent=ico;$('mTitle').textContent=title;$('mText').textContent=text;$('mOk').textContent=ok;$('mCancel').textContent=cancel;
  onOk=fn;$('modal').hidden=false;$('mCancel').focus();
}
$('mOk').onclick=()=>{$('modal').hidden=true;if(onOk)onOk()};
$('mCancel').onclick=()=>{$('modal').hidden=true};
addEventListener('keydown',e=>{if(e.key==='Escape')$('modal').hidden=true});

$('rows').addEventListener('input',e=>{
  const el=e.target;if(!el.dataset.k)return;
  if(el.tagName==='TEXTAREA'){el.value=el.value.replace(/\n/g,' ');fit(el)}
  if(el.dataset.k==='t'){el.value=el.value.replace(/\D/g,'');el.classList.toggle('badlen',el.value.length>0&&el.value.length!==13)}
  const r=rows[+el.dataset.i];r[el.dataset.k]=el.value;r.s='draft';delete r.msg;
  if(el.dataset.k==='a'){el.classList.toggle('cut',!!addrFlag(r))}  // ตรวจสีเหลืองใหม่ทุกครั้งที่พิมพ์
  const tg=el.closest('tr')?.querySelector('.tags');if(tg)tg.innerHTML=tagsHtml(r);
});
$('rows').addEventListener('click',e=>{
  const b=e.target.closest('[data-del]');if(!b)return;
  const i=+b.dataset.del, r=rows[i];
  ask({ico:'✕',title:`ลบแถว ${r.c}?`,text:`${r.n||'(ไม่มีชื่อ)'} จะถูกลบออกจากตารางในเว็บ${r.s==='done'?' ข้อมูลที่ลง Excel ไปแล้วจะไม่หาย':''}`,ok:'ลบแถวนี้'},()=>{
    rows.splice(i,1);save();fillPrefix();render();toast(`ลบ ${r.c} แล้ว`)});
});
$('rows').addEventListener('change',()=>{save();chips()});
$('chips').addEventListener('click',e=>{const b=e.target.closest('[data-f]');if(b){filter=b.dataset.f;reset()}});
$('q').addEventListener('input',reset);
$('prefix').addEventListener('change',reset);
$('file').addEventListener('change',e=>{if(e.target.files[0])handle(e.target.files[0]);e.target.value=''});
const drop=$('drop');
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('over')}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('over')}));
drop.addEventListener('drop',e=>{const f=e.dataTransfer.files[0];if(f)handle(f)});

// ปุ่ม 3: ส่งให้น้องบอท (ส่งทุกแถว รวมแถวที่ยังมีช่องว่างและรหัสซ้ำ — ไฟล์ TCS รับรหัสซ้ำเป็นแถวใหม่ได้)
$('send').addEventListener('click',e=>{
  let n=0;
  rows.forEach(r=>{if(r.s==='draft'||r.s==='fail'){r.s='ready';delete r.msg;n++}});
  save();render();
  toast(n?`ส่งให้น้องบอทแล้ว ${n.toLocaleString()} ราย 🌻`:'ไม่มีรายการใหม่ที่ต้องส่ง');
  if(n){const b=e.currentTarget.getBoundingClientRect();sparkBurst(b.left+b.width/2,b.top+b.height/2)}
});
// ปุ่ม 2: ล้างแถวที่ข้อมูลครบ เหลือไว้แค่แถวที่ยังมีช่องว่าง
$('clearDone').addEventListener('click',()=>{
  const full=rows.filter(complete), notSent=full.filter(r=>r.s!=='done').length;
  if(!full.length){toast('ยังไม่มีแถวที่ข้อมูลครบ');return}
  ask({ico:'🧹',title:`ล้าง ${full.length.toLocaleString()} แถวที่ข้อมูลครบ`,
    text:notSent?`ในนี้มี ${notSent.toLocaleString()} แถวที่ยังไม่ลง Excel ถ้าล้างตอนนี้ บอทจะไม่ได้ข้อมูลแถวนั้น แนะนำให้ส่งให้น้องบอทก่อน`:'ทุกแถวที่จะล้างลง Excel แล้ว จะเหลือไว้แค่แถวที่ยังมีช่องว่างให้เติมต่อ',
    ok:notSent?'ล้างเลย':'ล้าง'},()=>{
      rows=rows.filter(r=>!complete(r));save();setSrc();fillPrefix();render();toast(`ล้างแล้ว เหลือ ${rows.length.toLocaleString()} แถวที่ยังมีช่องว่าง`)});
});
// ปุ่ม 1: ล้างทุกข้อมูล
$('clearAll').addEventListener('click',()=>{
  if(!rows.length){toast('ไม่มีข้อมูลให้ล้าง');return}
  const notSent=rows.filter(r=>r.s!=='done').length;
  ask({ico:'🗑️',title:'ล้างทุกข้อมูลในหน้านี้?',
    text:`จะลบทั้งหมด ${rows.length.toLocaleString()} แถว${notSent?` (ยังไม่ลง Excel ${notSent.toLocaleString()} แถว)`:''} ข้อมูลที่ลง Excel ไปแล้วจะไม่หาย`,ok:'ล้างทั้งหมด'},()=>{
      rows=[];isSample=false;fileName='';try{localStorage.removeItem(KEY)}catch(e){}
      setSrc();fillPrefix();render();toast('ล้างทุกข้อมูลแล้ว')});
});

// ช่องทางให้บอท UiPath (Inject Js Script) — บอทไม่ต้องคลิกหรือพิมพ์อะไรบนหน้าจอ
// sunflow.ready()  → ข้อความ JSON: {file:"ชื่อไฟล์ export ที่ลากเข้ามา", rows:[{id,c,n,a,t}]}  (id = เลขประจำแถว เพราะรหัสซ้ำได้)
//                     file ใช้ตั้งชื่อไฟล์สำรอง เช่น Backup_2F_2026-10-09_0930.xlsx
// sunflow.mark(json) ← บอทส่งผลกลับ: [{id:"…",ok:true},{id:"…",ok:false,msg:"ไฟล์ Excel ถูกเปิดอยู่"}]
//   แถวที่ ok → หายจากเว็บ / แถวที่ไม่ ok → ค้างไว้ ขึ้น ⚠️ พร้อมสาเหตุ · ทำเฉพาะแถวที่ยังรอบอทอยู่ ถ้าพนักงานแก้แถวนั้นระหว่างบอททำงาน แถวจะกลับเป็น "ยังไม่ส่ง" และไม่ถูกติด ✓
window.sunflow={
  ready(){return JSON.stringify({file:fileName.replace(/\.[^.]+$/,''),rows:rows.filter(r=>r.s==='ready').map(({id,c,n,a,t})=>({id,c,n,a,t}))})},
  mark(json){
    let list=[];try{list=typeof json==='string'?JSON.parse(json):json}catch(e){return 'bad json'}
    let ok=0,bad=0;const byId=new Map(rows.filter(r=>r.s==='ready').map(r=>[r.id,r]));
    const gone=new Set();   // แถวที่บอทลง Excel และอ่านกลับตรวจแล้ว → ลบออกจากเว็บ เหลือแต่แถวที่ยังต้องทำ
    for(const x of list){const r=byId.get(x.id);if(!r)continue;if(x.ok){gone.add(r);ok++}else{r.s='fail';r.msg=x.msg||'ไม่ทราบสาเหตุ';bad++}}
    if(gone.size)rows=rows.filter(r=>!gone.has(r));
    save();fillPrefix();render();
    if(ok&&!bad)toast(`น้องบอทลง Excel แล้ว ${ok} ราย ✓ (เอาออกจากตารางแล้ว)`);
    if(bad){   // ลงไม่สำเร็จ → ป๊อปอัปเตือนพร้อมสาเหตุ ค้างไว้จนกว่าจะกด / กด "ส่งอีกครั้ง" = ส่งแถวที่ไม่สำเร็จกลับไปให้บอท บอทลงต่อเองไม่ต้องเปิดบอทใหม่
      const why=[...new Set(list.filter(x=>!x.ok).map(x=>x.msg||'ไม่ทราบสาเหตุ'))].join(' / ');
      const open=/เปิดอยู่/.test(why);
      ask({ico:'⚠️',title:`ลง Excel ไม่สำเร็จ ${bad.toLocaleString()} ราย`+(ok?` (สำเร็จ ${ok.toLocaleString()} ราย)`:''),
        text:`สาเหตุ: ${why}\n\n`+(open?'ปิดไฟล์ Excel ก่อน แล้วกด "ส่งอีกครั้ง" น้องบอทจะลงให้ต่อเอง':'แก้สาเหตุแล้วกด "ส่งอีกครั้ง" น้องบอทจะลงให้ต่อเอง')+'\nถ้ากด "ไว้ทีหลัง" แถวจะค้างไว้พร้อม ⚠️ กด 🌻 ส่งให้น้องบอท เมื่อพร้อม',
        ok:'ส่งอีกครั้ง',cancel:'ไว้ทีหลัง'},
        ()=>{let n=0;rows.forEach(r=>{if(r.s==='fail'){r.s='ready';delete r.msg;n++}});save();render();toast(`ส่งให้น้องบอทอีกครั้ง ${n.toLocaleString()} ราย 🌻`)});
    }
    return `ok:${ok} fail:${bad}`;
  }
};

// ธีมกลางวัน / กลางคืน
function isDark(){const t=document.documentElement.dataset.theme;return t?t==='dark':matchMedia('(prefers-color-scheme: dark)').matches}
function themeLabel(){$('themeBtn').textContent=isDark()?'☀️ ธีมกลางวัน':'🌙 เปลี่ยนธีม'}
$('themeBtn').onclick=()=>{document.documentElement.dataset.theme=isDark()?'light':'dark';themeLabel()};

// พื้นหลังวิบวับ: ดาวกระพริบ + ผงตามเมาส์ (จากเว็บเดิม)
const reduced=matchMedia("(prefers-reduced-motion: reduce)").matches;
const TINTS=["124,108,240","240,165,200","140,190,240","250,205,140"],SPARKS=["✦","✧","⋆","✨"],COLORS=["#a99cff","#f0a5c8","#8cbef0","#fad08c"];
function sparkBurst(x,y,n=24){if(reduced)return;for(let i=0;i<n;i++){const p=document.createElement("div");p.className="dust";p.textContent=SPARKS[i%4];
  p.style.left=x+"px";p.style.top=y+"px";p.style.fontSize=(10+Math.random()*12)+"px";p.style.color=COLORS[i%4];const a=6.283*i/n,dist=70+Math.random()*90;
  p.style.setProperty("--dx",Math.cos(a)*dist+"px");p.style.setProperty("--dy",Math.sin(a)*dist+"px");p.style.animationDuration="1.1s";document.body.appendChild(p);setTimeout(()=>p.remove(),1150)}}
if(!reduced){
  const cv=$('sky'),ctx=cv.getContext("2d");let stars=[],W=0,H=0;
  const resize=()=>{W=cv.width=innerWidth;H=cv.height=innerHeight;stars=Array.from({length:Math.floor(W*H/6000)},()=>({x:Math.random()*W,y:Math.random()*H,r:Math.random()*1.6+.5,a:Math.random()*6.28,sp:Math.random()*.018+.006,dy:Math.random()*.14+.03,t:TINTS[Math.floor(Math.random()*4)]}))};
  const draw=()=>{ctx.clearRect(0,0,W,H);for(const s of stars){s.a+=s.sp;const al=((Math.sin(s.a)+1)/2)*.55+.08;s.y-=s.dy;if(s.y<-5){s.y=H+5;s.x=Math.random()*W}
    ctx.beginPath();ctx.arc(s.x,s.y,s.r,0,6.283);ctx.fillStyle=`rgba(${s.t},${al})`;ctx.shadowBlur=s.r*5;ctx.shadowColor=`rgba(${s.t},.5)`;ctx.fill()}ctx.shadowBlur=0;requestAnimationFrame(draw)};
  addEventListener("resize",()=>{resize();fitAll()});resize();draw();
  let last=0;addEventListener("pointermove",e=>{const now=Date.now();if(now-last<30)return;last=now;const d=document.createElement("div");d.className="dust";
    d.textContent=SPARKS[Math.floor(Math.random()*4)];d.style.left=e.clientX+"px";d.style.top=e.clientY+"px";d.style.color=COLORS[Math.floor(Math.random()*4)];
    d.style.setProperty("--dx",(Math.random()*46-23)+"px");d.style.setProperty("--dy",(Math.random()*38+14)+"px");document.body.appendChild(d);setTimeout(()=>d.remove(),1000)});
}

load();setSrc();fillPrefix();render();themeLabel();
