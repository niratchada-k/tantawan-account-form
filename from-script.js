"use strict";
/* =========================================================
   TCS SunFlow — อ่านไฟล์ export → จัดเป็นตาราง → ส่งให้น้องบอท (UiPath) ลง Excel
   1. ค่าตั้งต้นและตัวช่วย
   2. เก็บข้อมูลในเบราว์เซอร์
   3. อ่านไฟล์ Excel
   4. ตรวจที่อยู่ (ขึ้นสีเหลือง)
   5. วาดตาราง
   6. ป๊อปอัปและแจ้งเตือน
   7. แก้ข้อมูล / ลากไฟล์ / ปุ่ม
   8. คุยกับน้องบอท + ไฟสถานะ
   9. ธีมและพื้นหลัง
   10. เริ่มทำงาน
   ========================================================= */


/* ---------- 1. ค่าตั้งต้นและตัวช่วย ---------- */

const STORAGE_KEY = 'tcs_queue';  // ชื่อที่ใช้เก็บข้อมูลในเบราว์เซอร์
const BATCH = 200;                // วาดตารางทีละ 200 แถว เลื่อนใกล้ล่างสุดค่อยเติม
const BOT_TIMEOUT = 10000;        // น้องบอทไม่แวะมาเกิน 10 วินาที = ไฟสถานะดับ
const BUSY_MAX = 30 * 60000;      // กำลังลง Excel นานสุด 30 นาที เกินนี้ถือว่าน้องบอทหยุดไปแล้ว
const CONTACT = 'ผู้ดูแลระบบ';    // คนที่ให้พนักงานติดต่อเมื่อบันทึกไม่สำเร็จและแก้เองไม่ได้ (เปลี่ยนเป็นชื่อคน/แผนกได้)

// ข้อมูลแต่ละแถว: { id, c: รหัส, n: ชื่อ, a: ที่อยู่, t: เลขภาษี, s: สถานะ, pad: เติม 0 แล้ว, msg: สาเหตุที่ลงไม่สำเร็จ }
// สถานะ s: 'draft' = ยังไม่ส่ง / 'ready' = รอน้องบอท / 'fail' = ลงไม่สำเร็จ  (ลงสำเร็จแล้วแถวจะถูกเอาออกจากตาราง)
let rows = [];
let fileName = '';         // ชื่อไฟล์ export ที่ลากเข้ามา
let filter = 'all';        // ปุ่มกรองที่เลือกอยู่
let shown = BATCH;         // จำนวนแถวที่วาดแล้ว
let current = [];          // แถวที่ผ่านตัวกรองตอนนี้ [แถว, ลำดับใน rows]
let dupCodes = new Set();  // รหัสที่ซ้ำกันในตาราง

const $ = id => document.getElementById(id);
const clean = s => String(s ?? '').replace(/\s+/g, ' ').trim();   // ยุบช่องว่างซ้อน / ขึ้นบรรทัดใหม่ / ช่องว่างพิเศษจาก Excel ให้เหลือเว้นวรรคเดียว
const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
const num = n => n.toLocaleString();
const isComplete = r => !!(r.n && r.a && r.t);

// เลขประจำแถว (รหัสซ้ำได้ เลยต้องมี id แยกแต่ละแถว ให้น้องบอทตอบกลับถูกแถว)
let seq = 0;
const newId = () => Date.now().toString(36) + '-' + (seq++).toString(36);


/* ---------- 2. เก็บข้อมูลในเบราว์เซอร์ (ปิดเว็บแล้วเปิดใหม่ยังอยู่) ---------- */

function load() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
    if (saved && Array.isArray(saved.rows)) {
      rows = saved.rows.filter(r => r.s !== 'done');   // แถวที่ลง Excel แล้วจากเวอร์ชันเก่า ไม่ต้องแสดง
      rows.forEach(r => { if (!r.id) r.id = newId(); });
      fileName = saved.file || '';
    }
  } catch (e) {
    rows = [];
  }
}

function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ file: fileName, rows }));
  } catch (e) {
    toast('บันทึกข้อมูลในเบราว์เซอร์ไม่ได้ (พื้นที่เต็ม)');
  }
  syncBridge();   // อัปเดตงานในตู้ไปรษณีย์ให้น้องบอททุกครั้งที่ข้อมูลเปลี่ยน
}

function showSource() {
  $('src').textContent = rows.length
    ? `ไฟล์ ${fileName}`
    : 'ยังไม่มีข้อมูล';
}


/* ---------- 3. อ่านไฟล์ Excel ---------- */

// เลขภาษี: เก็บเป็นตัวเลขล้วนแบบข้อความ / ถ้าได้ 12 หลัก แปลว่า 0 ข้างหน้าหาย → เติม 0 ให้ครบ 13 หลัก
function toTaxId(value) {
  let t = typeof value === 'number' ? Math.round(value).toString() : clean(value).replace(/\D/g, '');
  const padded = t.length === 12;
  if (padded) t = '0' + t;
  return { t, padded };
}

// แบบรายงาน (เช่น 2F): คอลัมน์ D เป็นรหัส = เริ่มรายใหม่ / คอลัมน์ F เป็นป้าย บอกว่าค่าในคอลัมน์ G คืออะไร
const CODE_PATTERN = /^[A-Z]{2,6}(-[A-Z0-9]+)+$/;
function parseReport(sheet) {
  const out = [];
  let rec = null;
  for (const r of sheet) {
    const code = clean(r[3]);
    const label = clean(r[5]);
    const value = r[6];
    if (CODE_PATTERN.test(code)) {
      rec = { c: code, n: clean(r[4]), a: '', t: '', s: 'draft' };
      out.push(rec);
    }
    if (!rec || !label) continue;
    if (label === 'ที่อยู่' && !rec.a) {
      rec.a = clean(value);
    } else if (label.startsWith('เลขที่ประจำตัวผู้เสียภาษี') && !rec.t) {
      const tax = toTaxId(value);
      rec.t = tax.t;
      if (tax.padded) rec.pad = 1;
    }
  }
  return out;
}

// แบบตาราง (มีแถวหัวคอลัมน์ เช่น รหัส | ชื่อ | ที่อยู่ | เลขประจำตัว): หาว่าคอลัมน์ไหนคืออะไร
function headerMap(r) {
  const m = {};
  r.forEach((v, j) => {
    const h = clean(v).toLowerCase();
    if (!h) return;
    if (m.t == null && (h.includes('ประจำตัว') || h.includes('ภาษี') || h.startsWith('vat') || h.includes('tax'))) m.t = j;
    else if (m.c == null && (h.startsWith('รหัส') || h === 'no.' || h === 'code')) m.c = j;
    else if (m.n == null && (h.startsWith('ชื่อ') || h === 'name')) m.n = j;
    else if (m.a == null && (h.startsWith('ที่อยู่') || h === 'address')) m.a = j;
  });
  return m;
}

function parseTable(sheet) {
  // ใช้แถวหัวตารางแถวสุดท้ายใน 15 แถวแรก ที่มีคอลัมน์รหัส + อย่างน้อย 2 ใน 3 (ชื่อ / ที่อยู่ / เลขภาษี)
  let headerRow = -1, map = null;
  for (let i = 0; i < Math.min(sheet.length, 15); i++) {
    const m = headerMap(sheet[i]);
    if (m.c != null && [m.n, m.a, m.t].filter(x => x != null).length >= 2) { headerRow = i; map = m; }
  }
  if (headerRow < 0) return [];

  const out = [];
  for (const r of sheet.slice(headerRow + 1)) {
    const code = clean(r[map.c]);
    if (!code) continue;
    const rec = { c: code, n: map.n != null ? clean(r[map.n]) : '', a: '', t: '', s: 'draft' };
    if (map.a != null) rec.a = clean(r[map.a]);
    if (map.t != null) {
      const tax = toTaxId(r[map.t]);
      rec.t = tax.t;
      if (tax.padded) rec.pad = 1;
    }
    out.push(rec);
  }
  return out;
}

// อ่านไฟล์: ลองทุกชีต ใช้ชีตแรกที่อ่านข้อมูลได้ (ลองแบบรายงานก่อน แล้วค่อยแบบตาราง)
function readFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    let got = [];
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      for (const name of wb.SheetNames) {
        const ws = wb.Sheets[name];
        if (!ws['!ref']) continue;
        const range = XLSX.utils.decode_range(ws['!ref']);
        range.s.c = 0; range.s.r = 0;   // อ่านตั้งแต่ A1 เสมอ ตำแหน่งคอลัมน์จะได้ไม่เลื่อน
        const sheet = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', range });
        got = parseReport(sheet);
        if (!got.length) got = parseTable(sheet);
        if (got.length) break;
      }
    } catch (err) {
      toast('เปิดไฟล์ไม่ได้ รองรับเฉพาะ .xlsx และ .xls');
      return;
    }
    if (!got.length) {
      toast('ไม่พบข้อมูล Vendor ในไฟล์นี้');
      return;
    }
    got.forEach(r => { r.id = newId(); });

    const apply = () => {
      rows = got; fileName = file.name; filter = 'all'; shown = BATCH;
      save(); showSource(); fillPrefix(); render();
      toast(`นำเข้าข้อมูล ${num(got.length)} รายการ`);
      const b = $('drop').getBoundingClientRect();
      sparkBurst(b.left + b.width / 2, b.top + b.height / 2);
    };
    if (rows.length) {
      ask({ ico: '📂', title: 'แทนที่ข้อมูลเดิม?',
            text: `มี ${num(rows.length)} รายการที่ยังไม่ได้บันทึก หากแทนที่ด้วย ${file.name} รายการเหล่านี้จะถูกลบ`,
            ok: 'แทนที่' }, apply);
    } else {
      apply();
    }
  };
  reader.readAsArrayBuffer(file);
}


/* ---------- 4. ตรวจที่อยู่ (ขึ้นสีเหลือง ตรวจสดทุกครั้งที่พิมพ์) ---------- */
/* กฎทั้งหมดตั้งจากปัญหาที่เจอจริงในไฟล์ 2F (5,416 ราย) */

// คำว่า "เงื่อนไขการชำระเงิน" ที่หลุดมาปนในที่อยู่ (ทั้งคำ หรือเศษคำ 5 ตัวอักษรขึ้นไป) — ที่อยู่ถูกตัดขาด
const STRAY = (w => new RegExp([...w].map((_, i) => w.slice(0, w.length - i)).filter(x => x.length >= 5).join('|'), 'g'))('เงื่อนไขการชำระเงิน');

// มีคำปน หรือท้ายที่อยู่เป็นเศษสั้น ๆ (เ / เง / เงื / เงื่)
function hasStray(a) {
  const found = STRAY.test(a);
  STRAY.lastIndex = 0;
  return found || /(^|[\dA-Za-z\sก-๙])(เงื่|เงื|เง|เ)$/.test(a.trim());
}
// ท้ายที่อยู่มีแค่คำนำหน้า แล้วไม่มีชื่อต่อ เช่น "…อ.เมือง จ." หรือ "…จ. 20000"
const endsWithLabel = a =>
  /(^|\s|\d)(จ\.|จังหวัด|อ\.|อำเภอ|ต\.|ตำบล|เขต|แขวง|ถ\.|ถนน|ซ\.|ซอย|ม\.|หมู่ที่|หมู่)\s*$/.test(a) ||
  /(^|\s)(จ\.|จังหวัด)\s*\d*\s*$/.test(a);
// ไม่มีรหัสไปรษณีย์ 5 หลัก
const noZip = a => !/(^|\D)\d{5}(\D|$)/.test(a);

// ชื่อจังหวัด 77 จังหวัด (มีชื่อจังหวัดแม้ไม่เขียน "จ." นำหน้า ถือว่ามีจังหวัด)
const PROVINCES = ('กรุงเทพ กทม กระบี่ กาญจนบุรี กาฬสินธุ์ กำแพงเพชร ขอนแก่น จันทบุรี ฉะเชิงเทรา ชลบุรี ชัยนาท ชัยภูมิ ชุมพร ' +
  'เชียงราย เชียงใหม่ ตรัง ตราด ตาก นครนายก นครปฐม นครพนม นครราชสีมา นครศรีธรรมราช นครสวรรค์ นนทบุรี นราธิวาส น่าน ' +
  'บึงกาฬ บุรีรัมย์ ปทุมธานี ประจวบคีรีขันธ์ ปราจีนบุรี ปัตตานี อยุธยา พะเยา พังงา พัทลุง พิจิตร พิษณุโลก เพชรบุรี ' +
  'เพชรบูรณ์ แพร่ ภูเก็ต มหาสารคาม มุกดาหาร แม่ฮ่องสอน ยโสธร ยะลา ร้อยเอ็ด ระนอง ระยอง ราชบุรี ลพบุรี ลำปาง ลำพูน เลย ' +
  'ศรีสะเกษ สกลนคร สงขลา สตูล สมุทรปราการ สมุทรสงคราม สมุทรสาคร สระแก้ว สระบุรี สิงห์บุรี สุโขทัย สุพรรณบุรี ' +
  'สุราษฎร์ธานี สุรินทร์ หนองคาย หนองบัวลำภู อ่างทอง อำนาจเจริญ อุดรธานี อุตรดิตถ์ อุทัยธานี อุบลราชธานี').split(' ');

// ที่อยู่ไทยต้องมี ตำบล/แขวง + อำเภอ/เขต + จังหวัด (ที่อยู่ภาษาอังกฤษ/ต่างประเทศไม่ตรวจข้อนี้)
// จับได้ทั้งที่ขาดจริง และที่พิมพ์ผิด เช่น ".กาญจนบุรี" "อแก้งคร้อ" "เขคปทุมวัน" "๖.เขารูปช้าง"
function missingParts(a) {
  if (!/[ก-๙]/.test(a)) return [];
  const miss = [];
  if (!/(ต\.|ตำบล|แขวง)/.test(a)) miss.push('ต.');
  if (!/(อ\.|อำเภอ|เขต)/.test(a)) miss.push('อ.');
  if (!/(จ\.|จังหวัด)/.test(a) && !PROVINCES.some(p => a.includes(p))) miss.push('จ.');
  return miss;
}

// ผลตรวจที่อยู่: ป้ายที่จะแสดง (ว่าง = ที่อยู่สมบูรณ์ ไม่ขึ้นเหลือง)
// ป้ายมี 2 แบบ: "ที่อยู่ต้องตรวจสอบ" (ถูกตัดขาด / จบด้วยคำนำหน้า / ขาด ต. อ. จ.) และ "ไม่มีรหัสไปรษณีย์"
function addrIssues(r) {
  const a = String(r.a ?? '').trim();
  if (!a) return [];                                   // ที่อยู่ว่าง นับใน "มีช่องว่าง" แทน
  const out = [];
  if (hasStray(a) || endsWithLabel(a) || missingParts(a).length) out.push('ที่อยู่ต้องตรวจสอบ');
  if (!hasStray(a) && noZip(a)) out.push('ไม่มีรหัสไปรษณีย์');
  return out;
}
const addrFlag = r => addrIssues(r).length > 0;


/* ---------- 5. วาดตาราง ---------- */

function findDuplicates() {
  const count = {}, dups = new Set();
  rows.forEach(r => { count[r.c] = (count[r.c] || 0) + 1; });
  for (const code in count) if (count[code] > 1) dups.add(code);
  return dups;
}

// แถวที่ผ่านตัวกรอง / ค้นหา / กลุ่มรหัส
function filteredRows() {
  const q = clean($('q').value).toLowerCase();
  const prefix = $('prefix').value;
  return rows.map((r, i) => [r, i]).filter(([r]) => {
    if (prefix && r.c.split('-')[0] !== prefix) return false;
    if (q && !(r.c.toLowerCase().includes(q) || r.n.toLowerCase().includes(q))) return false;
    switch (filter) {
      case 'empty': return !isComplete(r);
      case 'cut':   return !!addrFlag(r);
      case 'pad':   return !!r.pad;
      case 'dup':   return dupCodes.has(r.c);
      case 'fail':  return r.s === 'fail';
      default:      return true;
    }
  });
}

// ปุ่มกรอง (แสดงเฉพาะปุ่มที่มีจำนวน) + ข้อความสรุปด้านล่าง
const CHIPS = [
  ['all',   'ทั้งหมด',          ''],
  ['empty', 'มีช่องว่าง',        'empty'],
  ['cut',   'ที่อยู่ต้องตรวจ',    'warn'],
  ['pad',   'เติม 0 แล้ว',       'warn'],
  ['dup',   'รหัสซ้ำ',           'dup'],
  ['fail',  'บันทึกไม่สำเร็จ',    'empty'],
];
function renderChips() {
  const count = {
    all:   rows.length,
    empty: rows.filter(r => !isComplete(r)).length,
    cut:   rows.filter(r => addrFlag(r)).length,
    pad:   rows.filter(r => r.pad).length,
    dup:   rows.filter(r => dupCodes.has(r.c)).length,
    fail:  rows.filter(r => r.s === 'fail').length,
  };
  $('chips').innerHTML = CHIPS
    .filter(([key]) => key === 'all' || count[key] || filter === key)
    .map(([key, label, cls]) =>
      `<button type="button" class="chip ${cls}" data-f="${key}" aria-pressed="${filter === key}"><span class="n">${num(count[key])}</span>${label}</button>`)
    .join('');

  const waiting = rows.filter(r => r.s === 'draft' || r.s === 'fail').length;
  const ready = rows.filter(r => r.s === 'ready').length;
  $('footmsg').textContent = [
    waiting && `ยังไม่ได้ส่ง ${num(waiting)} รายการ`,
    ready && `รอบันทึก ${num(ready)} รายการ`,
  ].filter(Boolean).join(' · ');
}

// ป้ายในคอลัมน์ "สถานะ"
function tagsHtml(r) {
  let html = '';
  if (r.s === 'ready') html += '<span class="tag ready">รอบันทึก</span>';
  if (r.s === 'fail')  html += `<span class="tag fail" title="${esc(reason(r.msg))}">บันทึกไม่สำเร็จ</span>`;
  for (const t of addrIssues(r)) html += `<span class="tag cut">${t}</span>`;
  if (r.pad)           html += '<span class="tag pad">เติม 0</span>';
  if (dupCodes.has(r.c)) html += '<span class="tag dup">รหัสซ้ำ</span>';
  return html;
}

const rowHtml = ([r, i]) => `
  <tr>
    <td class="code${dupCodes.has(r.c) ? ' dup' : ''}">${esc(r.c)}</td>
    <td><textarea data-i="${i}" data-k="n" rows="1" placeholder="ว่าง">${esc(r.n)}</textarea></td>
    <td><textarea data-i="${i}" data-k="a" rows="1" class="${addrFlag(r) ? 'cut' : ''}" placeholder="ว่าง">${esc(r.a)}</textarea></td>
    <td><input data-i="${i}" data-k="t" class="tax${r.t && r.t.length !== 13 ? ' badlen' : ''}" inputmode="numeric" maxlength="13" value="${esc(r.t)}" placeholder="ว่าง"></td>
    <td><div class="tags">${tagsHtml(r)}</div></td>
    <td class="act"><button type="button" class="del" data-del="${i}" aria-label="ลบแถว ${esc(r.c)}">✕</button></td>
  </tr>`;

function render() {
  dupCodes = findDuplicates();
  renderChips();
  current = filteredRows();
  shown = Math.min(Math.max(shown, BATCH), Math.max(current.length, BATCH));
  $('rows').innerHTML = current.length
    ? current.slice(0, shown).map(rowHtml).join('')
    : `<tr><td colspan="6" class="none">${rows.length ? 'ไม่มีแถวที่ตรงกับตัวกรองนี้' : 'ยังไม่มีข้อมูล ลากไฟล์ export มาวางด้านบน'}</td></tr>`;
  fitAll();
  showCount();
}

function showCount() {
  $('pager').textContent = current.length ? `แสดง ${num(Math.min(shown, current.length))} จาก ${num(current.length)} แถว` : '';
}

// เปลี่ยนตัวกรอง / ค้นหา → กลับไปบนสุด
function resetView() {
  shown = BATCH;
  render();
  document.querySelector('.tablebox').scrollTop = 0;
}

// เลื่อนใกล้ล่างสุด → เติมอีก 200 แถว (หน้าเดียวเลื่อนยาว ไม่ต้องกดหน้าถัดไป)
function loadMore() {
  if (shown >= current.length) return;
  const add = current.slice(shown, shown + BATCH);
  shown += add.length;
  $('rows').insertAdjacentHTML('beforeend', add.map(rowHtml).join(''));
  fitAll();
  showCount();
}

// ช่องข้อความสูงตามเนื้อหา: Chrome / Edge รุ่นใหม่ทำเองด้วย CSS (field-sizing) ไม่ต้องคำนวณ
const AUTO_SIZE = !!(window.CSS && CSS.supports && CSS.supports('field-sizing', 'content'));
function fit(el) {
  if (AUTO_SIZE) return;
  el.style.height = 'auto';
  el.style.height = el.scrollHeight + 2 + 'px';
}
function fitAll() {
  if (AUTO_SIZE) return;
  const els = [...document.querySelectorAll('#rows textarea')];
  els.forEach(e => { e.style.height = 'auto'; });
  const heights = els.map(e => e.scrollHeight);
  els.forEach((e, i) => { e.style.height = heights[i] + 2 + 'px'; });
}

// ตัวเลือก "กลุ่มรหัส" (ส่วนหน้าเครื่องหมาย - ของรหัส เช่น EMP)
function fillPrefix() {
  const prefixes = [...new Set(rows.map(r => r.c.split('-')[0]))].sort();
  $('prefix').innerHTML = '<option value="">ทุกกลุ่มรหัส</option>' + prefixes.map(p => `<option>${esc(p)}</option>`).join('');
}


/* ---------- 6. ป๊อปอัปและแจ้งเตือน ---------- */

let toastTimer;
function toast(message) {
  const t = $('toast');
  t.textContent = message;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 2800);
}

// ป๊อปอัปยืนยัน: กดปุ่มบน = ทำ onOk / ปุ่มล่างหรือ Esc = ปิด
let onOk = null;
function ask({ ico = '⚠️', title, text, ok = 'ตกลง', cancel = 'ยกเลิก' }, fn) {
  $('mIco').textContent = ico;
  $('mTitle').textContent = title;
  $('mText').textContent = text;
  $('mOk').textContent = ok;
  $('mCancel').textContent = cancel || '';
  $('mCancel').hidden = !cancel;             // cancel: null = ป๊อปอัปปุ่มเดียว (แค่ "ตกลง")
  onOk = fn;
  $('modal').hidden = false;
  (cancel ? $('mCancel') : $('mOk')).focus();
}
$('mOk').addEventListener('click', () => { $('modal').hidden = true; if (onOk) onOk(); });
$('mCancel').addEventListener('click', () => { $('modal').hidden = true; });
addEventListener('keydown', e => { if (e.key === 'Escape') $('modal').hidden = true; });


/* ---------- 7. แก้ข้อมูล / ลากไฟล์ / ปุ่ม ---------- */

// พิมพ์แก้ในตาราง: บันทึกทันที แถวกลับเป็น "ยังไม่ส่ง" (กันน้องบอทลงข้อมูลเก่าก่อนแก้)
$('rows').addEventListener('input', e => {
  const el = e.target;
  if (!el.dataset.k) return;
  if (el.tagName === 'TEXTAREA') {
    el.value = el.value.replace(/\n/g, ' ');
    fit(el);
  }
  if (el.dataset.k === 't') {
    el.value = el.value.replace(/\D/g, '');
    el.classList.toggle('badlen', el.value.length > 0 && el.value.length !== 13);
  }
  const r = rows[+el.dataset.i];
  r[el.dataset.k] = el.value;
  r.s = 'draft';
  delete r.msg;
  if (el.dataset.k === 'a') el.classList.toggle('cut', !!addrFlag(r));
  const tags = el.closest('tr')?.querySelector('.tags');
  if (tags) tags.innerHTML = tagsHtml(r);
});
$('rows').addEventListener('change', () => { save(); renderChips(); });

// ปุ่ม ✕ ลบแถว
$('rows').addEventListener('click', e => {
  const btn = e.target.closest('[data-del]');
  if (!btn) return;
  const i = +btn.dataset.del, r = rows[i];
  ask({ ico: '✕', title: `ลบรายการ ${r.c}?`, text: `${r.n || '(ไม่มีชื่อ)'} จะถูกลบออกจากตารางนี้`, ok: 'ลบ' }, () => {
    rows.splice(i, 1);
    save(); fillPrefix(); render();
    toast(`ลบ ${r.c} แล้ว`);
  });
});

// ตัวกรอง / ค้นหา / กลุ่มรหัส / เลื่อนตาราง
$('chips').addEventListener('click', e => {
  const btn = e.target.closest('[data-f]');
  if (btn) { filter = btn.dataset.f; resetView(); }
});
$('q').addEventListener('input', resetView);
$('prefix').addEventListener('change', resetView);
document.querySelector('.tablebox').addEventListener('scroll', e => {
  const box = e.currentTarget;
  if (box.scrollTop + box.clientHeight > box.scrollHeight - 600) loadMore();
}, { passive: true });

// เลือกไฟล์ / ลากไฟล์มาวาง
$('file').addEventListener('change', e => {
  if (e.target.files[0]) readFile(e.target.files[0]);
  e.target.value = '';
});
const drop = $('drop');
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => {
  const f = e.dataTransfer.files[0];
  if (f) readFile(f);
});

// 🌻 ส่งให้น้องบอท: ส่งทุกแถวที่ยังไม่ส่ง / ลงไม่สำเร็จ (รวมแถวที่มีช่องว่างและรหัสซ้ำ)
$('send').addEventListener('click', e => {
  let n = 0;
  rows.forEach(r => {
    if (r.s === 'draft' || r.s === 'fail') { r.s = 'ready'; delete r.msg; n++; }
  });
  save(); render();
  toast(n ? `ส่งบันทึก ${num(n)} รายการ` : 'ไม่มีรายการที่รอส่ง');
  if (n) {
    const b = e.currentTarget.getBoundingClientRect();
    sparkBurst(b.left + b.width / 2, b.top + b.height / 2);
  }
});

// 🗑️ ล้างทุกข้อมูล (ข้อมูลที่ลง Excel ไปแล้วไม่หาย)
$('clearAll').addEventListener('click', () => {
  if (!rows.length) { toast('ไม่มีข้อมูล'); return; }
  ask({ ico: '🗑️', title: 'ล้างข้อมูลทั้งหมด?',
        text: `${num(rows.length)} รายการในหน้านี้จะถูกลบ (ข้อมูลที่บันทึกใน Excel แล้วไม่ได้รับผลกระทบ)`,
        ok: 'ล้างข้อมูล' }, () => {
    rows = []; fileName = '';
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) {}
    syncBridge(); showSource(); fillPrefix(); render();
    toast('ล้างข้อมูลแล้ว');
  });
});


/* ---------- 8. คุยกับน้องบอท + ไฟสถานะ ---------- */
/* น้องบอท (UiPath) ใช้ Inject Js Script อ่าน/เขียน attribute ของกล่องซ่อน #sunflowBridge ("ตู้ไปรษณีย์")
     data-ready = งานที่รอน้องบอท {file, rows:[{id,c,n,a,t}]}   ← น้องบอทอ่าน
     data-ping  = เวลาที่น้องบอทแวะมาล่าสุด                    ← น้องบอทเขียนทุกรอบ (ใช้ทำไฟสถานะ)
     data-mark  = ผลการลง Excel [{id, ok, msg, row}]           ← น้องบอทเขียน เว็บอ่านแล้วล้างทิ้ง (row = แถวใน Excel ที่ลงสำเร็จ)
   ใช้กล่องบนหน้าเว็บแทนตัวแปร เพราะส่วนขยาย Chrome ของ UiPath อาจมองไม่เห็นตัวแปรของเว็บ แต่เห็นหน้าเว็บเสมอ */

const bridge = $('sunflowBridge') || document.body.appendChild(Object.assign(document.createElement('div'), { id: 'sunflowBridge', hidden: true }));

// งานที่รอน้องบอท (file = ชื่อไฟล์ export ไม่มีนามสกุล ใช้ตั้งชื่อไฟล์สำรอง เช่น Backup_2F_2026-10-12_0930.xlsx)
function readyJson() {
  return JSON.stringify({
    file: fileName.replace(/\.[^.]+$/, ''),
    rows: rows.filter(r => r.s === 'ready').map(({ id, c, n, a, t }) => ({ id, c, n, a, t })),
  });
}
function syncBridge() {
  bridge.setAttribute('data-ready', readyJson());
}

// รับผลจากน้องบอท: ok = ลงและตรวจแล้ว → เอาออกจากตาราง / ไม่ ok = ค้างไว้พร้อม ⚠️ และสาเหตุ
// ทำเฉพาะแถวที่ยัง "รอน้องบอท" (ถ้าถูกแก้ระหว่างนั้น แถวจะกลับเป็นยังไม่ส่ง และไม่ถูกนับ)
// สาเหตุแบบสั้นสำหรับแสดงผล (ตัดคำซ้ำที่น้องบอทใส่มา เช่น "ลง Excel ไม่สำเร็จ:" / "(ไม่มีข้อมูลลง Excel)")
function reason(msg) {
  return clean(msg || 'ไม่ทราบสาเหตุ')
    .replace(/^ลง Excel ไม่สำเร็จ:\s*/, '')
    .replace(/\s*จึงไม่บันทึกไฟล์/, '')
    .replace(/\s*\(ไม่มีข้อมูลลง Excel\)/, '')
    .replace(/^ไฟล์ Excel ถูกเปิดอยู่$/, 'ไฟล์ Excel เปิดอยู่') || 'ไม่ทราบสาเหตุ';
}

function applyMark(json) {
  let list;
  try { list = typeof json === 'string' ? JSON.parse(json) : json; }
  catch (e) { return 'bad json'; }

  const waiting = new Map(rows.filter(r => r.s === 'ready').map(r => [r.id, r]));
  const done = new Set();
  let ok = 0, bad = 0;
  for (const x of list) {
    const r = waiting.get(x.id);
    if (!r) continue;
    if (x.ok) { done.add(r); ok++; }
    else { r.s = 'fail'; r.msg = x.msg || 'ไม่ทราบสาเหตุ'; bad++; }
  }
  if (done.size) rows = rows.filter(r => !done.has(r));
  busySince = 0;   // ได้ผลแล้ว = น้องบอทลง Excel เสร็จรอบนี้
  save(); fillPrefix(); render(); updateLight();

  if (ok && !bad) {
    // ป๊อปอัปค้างจนกว่าจะกด "ตกลง" (กันพนักงานพลาดไม่เห็น) + บอกแถวใน Excel ถ้าน้องบอทส่งมา
    const nums = list.filter(x => x.ok && x.row).map(x => +x.row).filter(Boolean);
    const where = nums.length
      ? `แถว ${Math.min(...nums)}` + (nums.length > 1 ? `–${Math.max(...nums)}` : '') + ' · ชีต Vendor\n'
      : '';
    ask({
      ico: '✅',
      title: `บันทึกสำเร็จ ${num(ok)} รายการ`,
      text: where + 'ตรวจสอบความถูกต้องแล้ว',
      ok: 'ตกลง',
      cancel: null,
    });
  }
  if (bad) {
    // ปัญหาที่พนักงานแก้เองได้ (ไฟล์เปิดค้าง) → มีปุ่ม "ลองอีกครั้ง"
    // ปัญหาอื่น (ตรวจซ้ำไม่ตรง / หาไฟล์ไม่เจอ / ไม่มีสิทธิ์ / Excel error) → ปุ่ม "ปิด" อย่างเดียว ให้แจ้ง CONTACT
    // ทุกกรณีน้องบอทไม่บันทึกอะไรเลย (ลงครบหรือไม่ลงเลย) แถวค้างไว้พร้อมป้าย "บันทึกไม่สำเร็จ" กดส่งใหม่ได้เมื่อแก้แล้ว
    const msgs = [...new Set(list.filter(x => !x.ok).map(x => reason(x.msg)))];
    const why = msgs.join(' / ');
    const retry = msgs.every(m => /เปิดอยู่|ถูกใช้งาน|ล็อก/.test(m));
    const head = ok ? `บันทึกสำเร็จ ${num(ok)} รายการ · ไม่สำเร็จ ${num(bad)} รายการ` : null;
    const resend = () => {
      let n = 0;
      rows.forEach(r => { if (r.s === 'fail') { r.s = 'ready'; delete r.msg; n++; } });
      save(); render();
      toast(`ส่งบันทึกอีกครั้ง ${num(n)} รายการ`);
    };
    if (retry) {
      ask({
        ico: '⚠️',
        title: head || 'ไฟล์ Excel เปิดอยู่',
        text: 'กรุณาปิดไฟล์ TCS แล้วกด "ลองอีกครั้ง"\nยังไม่มีข้อมูลถูกบันทึก',
        ok: 'ลองอีกครั้ง',
        cancel: 'ภายหลัง',
      }, resend);
    } else {
      ask({
        ico: '⚠️',
        title: head || `บันทึกไม่สำเร็จ ${num(bad)} รายการ`,
        text: `สาเหตุ: ${why}\nยังไม่มีข้อมูลถูกบันทึก กรุณาแจ้ง${CONTACT}`,
        ok: 'ปิด',
        cancel: null,
      });
    }
  }
  return `ok:${ok} fail:${bad}`;
}

// น้องบอทแวะมา (data-ping) ตอนมีงานรอ = น้องบอทหยิบงานไปแล้ว → กำลังลง Excel
// น้องบอทวางผลใน data-mark → อ่านทันทีแล้วล้างทิ้ง
let busySince = 0;
new MutationObserver(list => {
  for (const m of list) {
    if (m.attributeName === 'data-ping' && rows.some(r => r.s === 'ready')) busySince = Date.now();
  }
  const mark = bridge.getAttribute('data-mark');
  if (mark) {
    bridge.removeAttribute('data-mark');
    applyMark(mark);
  }
  updateLight();
}).observe(bridge, { attributes: true, attributeFilter: ['data-mark', 'data-ping'] });

// ไฟสถานะ 3 แบบ
//   🟡 กำลังลง Excel = น้องบอทหยิบงานไปแล้ว ยังไม่ส่งผลกลับ (ตอนนี้น้องบอทอยู่ใน Excel จึงไม่ได้แวะมาที่เว็บ)
//   🟢 พร้อม        = น้องบอทแวะมาภายใน 10 วินาที
//   ⚪ ยังไม่ทำงาน   = นอกเหนือจากนี้
// ระหว่าง 🟡 ล็อกตาราง / ช่องลากไฟล์ / ปุ่มล่าง ไว้ก่อน กันแก้ข้อมูลที่น้องบอทกำลังลงอยู่
const light = $('botLight');
const LOCK_WHILE_BUSY = ['drop', 'rows', 'clearAll', 'send'];
function updateLight() {
  if (!light) return;
  const now = Date.now();
  const busy = busySince > 0 && now - busySince < BUSY_MAX && rows.some(r => r.s === 'ready');
  const on = now - (+bridge.getAttribute('data-ping') || 0) < BOT_TIMEOUT;
  light.classList.toggle('busy', busy);
  light.classList.toggle('on', on && !busy);
  light.textContent = busy ? '🟡 กำลังบันทึก… กรุณาอย่าปิดหน้านี้'
                    : on   ? '🟢 พร้อมใช้งาน'
                    :        '⚪ ระบบยังไม่เปิด';
  light.title = busy ? 'ระหว่างบันทึกจะแก้ไขข้อมูลในตารางไม่ได้'
              : on   ? 'ส่งบันทึกได้เลย'
              :        'เปิดระบบจากไอคอน TCS SunFlow บนเดสก์ท็อป';
  LOCK_WHILE_BUSY.forEach(id => { const el = $(id); if (el) el.inert = busy; });
  document.body.classList.toggle('botbusy', busy);
}

// เปิดไว้ให้ลองเรียกจาก Console ได้ (เช่น sunflow.ready())
window.sunflow = { ready: readyJson, mark: applyMark };


/* ---------- 9. ธีมและพื้นหลัง ---------- */

function isDark() {
  const t = document.documentElement.dataset.theme;
  return t ? t === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
}
function themeLabel() {
  $('themeBtn').textContent = isDark() ? '☀️ โหมดสว่าง' : '🌙 โหมดมืด';
}
$('themeBtn').addEventListener('click', () => {
  document.documentElement.dataset.theme = isDark() ? 'light' : 'dark';
  themeLabel();
});

// ดาวกระพริบ + ผงประกายตามเมาส์ (จากเว็บเดิม) — ปิดเองถ้าเครื่องตั้งค่าลดการเคลื่อนไหว
const reducedMotion = matchMedia('(prefers-reduced-motion: reduce)').matches;
const TINTS = ['124,108,240', '240,165,200', '140,190,240', '250,205,140'];
const SPARKS = ['✦', '✧', '⋆', '✨'];
const COLORS = ['#a99cff', '#f0a5c8', '#8cbef0', '#fad08c'];

function makeDust(x, y, symbol, color, dx, dy, size, life) {
  const d = document.createElement('div');
  d.className = 'dust';
  d.textContent = symbol;
  d.style.left = x + 'px';
  d.style.top = y + 'px';
  d.style.color = color;
  if (size) d.style.fontSize = size + 'px';
  if (life) d.style.animationDuration = life / 1000 + 's';
  d.style.setProperty('--dx', dx + 'px');
  d.style.setProperty('--dy', dy + 'px');
  document.body.appendChild(d);
  setTimeout(() => d.remove(), (life || 1000) + 50);
}

// ประกายแตกออกเป็นวง (ตอนลากไฟล์สำเร็จ / กดส่ง)
function sparkBurst(x, y, n = 24) {
  if (reducedMotion) return;
  for (let i = 0; i < n; i++) {
    const angle = 2 * Math.PI * i / n, dist = 70 + Math.random() * 90;
    makeDust(x, y, SPARKS[i % 4], COLORS[i % 4], Math.cos(angle) * dist, Math.sin(angle) * dist, 10 + Math.random() * 12, 1100);
  }
}

function startSky() {
  const cv = $('sky'), ctx = cv.getContext('2d');
  let stars = [], W = 0, H = 0;
  const resize = () => {
    W = cv.width = innerWidth;
    H = cv.height = innerHeight;
    stars = Array.from({ length: Math.floor(W * H / 6000) }, () => ({
      x: Math.random() * W, y: Math.random() * H, r: Math.random() * 1.6 + .5,
      a: Math.random() * 6.28, sp: Math.random() * .018 + .006, dy: Math.random() * .14 + .03,
      tint: TINTS[Math.floor(Math.random() * 4)],
    }));
  };
  const draw = () => {
    ctx.clearRect(0, 0, W, H);
    for (const s of stars) {
      s.a += s.sp;
      s.y -= s.dy;
      if (s.y < -5) { s.y = H + 5; s.x = Math.random() * W; }
      const alpha = ((Math.sin(s.a) + 1) / 2) * .55 + .08;
      ctx.beginPath();
      ctx.arc(s.x, s.y, s.r, 0, 2 * Math.PI);
      ctx.fillStyle = `rgba(${s.tint},${alpha})`;
      ctx.shadowBlur = s.r * 5;
      ctx.shadowColor = `rgba(${s.tint},.5)`;
      ctx.fill();
    }
    ctx.shadowBlur = 0;
    requestAnimationFrame(draw);
  };
  addEventListener('resize', () => { resize(); fitAll(); });
  resize();
  draw();

  let last = 0;
  addEventListener('pointermove', e => {
    const now = Date.now();
    if (now - last < 70) return;
    last = now;
    const pick = Math.floor(Math.random() * 4);
    makeDust(e.clientX, e.clientY, SPARKS[pick], COLORS[Math.floor(Math.random() * 4)], Math.random() * 46 - 23, Math.random() * 38 + 14);
  });
}


/* ---------- 10. เริ่มทำงาน ---------- */

load();
showSource();
fillPrefix();
render();
themeLabel();
syncBridge();
updateLight();
setInterval(updateLight, 2000);
if (!reducedMotion) startSky();
