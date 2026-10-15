/* ================= Balance Waste — ระบบบริหารจัดการของเสีย ================= */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const view = $('#view');
const S = { types: [], receivers: [], balance: {}, containers: {}, company: {}, limits: {}, charts: [] };
const store = (fn, k, v) => { try { return fn === 'get' ? localStorage.getItem(k) : fn === 'del' ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch { return null; } };
/* ผู้ใช้งาน: ชื่อที่ตั้งไว้มุมขวาบน (เก็บในเบราว์เซอร์นี้) — ส่งไปกับทุกการบันทึก เป็น “ผู้บันทึก” ในประวัติ / Log */
S.user = store('get', 'wthep_user') || '';

const KIND = {
  store:   { name: 'บันทึกจัดเก็บ', short: 'จัดเก็บ', icon: 'fa-box-archive', c: '#10B981', cls: 'in', sign: '+' },
  ship:    { name: 'ส่งกำจัด / นำออก', short: 'ส่งกำจัด', icon: 'fa-truck-arrow-right', c: '#F97316', cls: 'out', sign: '−' },
  opening: { name: 'ยอดยกมา', short: 'ยอดยกมา', icon: 'fa-flag-checkered', c: '#6366F1', cls: 'open', sign: '+' },
  treat:   { name: 'ข้อมูลภายใน', short: 'ข้อมูลภายใน', icon: 'fa-flask-vial', c: '#7C3AED', cls: 'treat', sign: '−' },
};
// Movement kinds the current view may show (in-house rows: on the unlocked ข้อมูลภายใน page only).
const kinds = () => Object.entries(KIND).filter(([k]) => k !== 'treat' || treatOn());
const CONT = {
  drum200: { name: 'ถัง 200 ลิตร', sub: 'Drum 200 L' },
  ibc:     { name: 'IBC', sub: 'IBC Tank 1,000 L' },
  bigbag:  { name: 'Bigbag', sub: 'Jumbo bag' },
};
const CONT_KEYS = Object.keys(CONT);
const PAGES = {
  dashboard: ['แดชบอร์ด', 'ภาพรวมของเสียในโรงงาน · โควตาตามหนังสืออนุญาต สก.2'],
  store: ['บันทึกจัดเก็บ', 'ลงข้อมูลของเสียที่จัดเก็บในโรงงาน — เลือกประเภท กรอกน้ำหนัก เลือกภาชนะ'],
  ship: ['ส่งกำจัด / นำออก', 'ส่งของเสียให้ผู้รับดำเนินการ — ไม่เกินยอดคงเหลือ'],
  treat: ['ข้อมูลภายใน', 'รายการภายในโรงงาน (บำบัดเอง) — ตัดยอดคงเหลือ ไม่นับโควตา สก.2 · ต้องใส่รหัสผ่าน'],
  balance: ['คงเหลือ & โควตา', 'ยอดคงเหลือในโรงงาน ภาชนะ และการใช้โควตาปีนี้'],
  history: ['ประวัติรายการ', 'ทุกรายการจัดเก็บ / ส่งกำจัด · ค้นหา · ส่งออก Excel'],
  logs: ['บันทึกประวัติ (Log)', 'ทุกการกระทำในระบบ — ใคร ทำอะไร เมื่อไร'],
  settings: ['ตั้งค่าประเภทของเสีย', 'ประเภทของเสีย ปริมาณที่ได้รับอนุญาต (ตัน) และผู้รับดำเนินการ'],
};

/* ---------- API (Google Apps Script web app) ---------- */
// ทุกคำสั่ง = POST text/plain (JSON) พร้อม token ของเซสชันใน body — ไม่เกิด CORS preflight และ token ไม่อยู่ใน URL
const API_URL = String((window.WASTE_CONFIG || {}).apiUrl || '').trim();
const NO_API_MSG = 'ยังไม่ได้ตั้งค่า apiUrl — ใส่ URL ของ Google Apps Script (/exec) ในไฟล์ docs/config.js';
async function gasFetch(url, opt) {
  let j;
  try {
    const res = await fetch(url, { mode: 'cors', credentials: 'omit', cache: 'no-store', ...opt });
    j = await res.json();
  } catch { throw new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ (ตรวจสอบอินเทอร์เน็ต / Google Apps Script)'); }
  if (!j || !j.ok) {
    const e = new Error((j && j.error) || 'เกิดข้อผิดพลาด');
    e.passwordError = !!(j && j.passwordError);
    e.code = (j && j.code) || '';
    throw e;
  }
  return { ok: true, ...(j.data || {}) };
}
const postJson = body => gasFetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify(body) });

/* ---------- รหัสผ่าน 3 ระดับ: ใช้เฉพาะเมนู “ข้อมูลภายใน” และเครื่องมือผู้ดูแล (เว็บเปิดใช้ได้ทุกคนโดยไม่ต้องใส่รหัส) ----------
   viewer = ดูได้อย่างเดียว (เปิดข้อมูลภายใน ดูรายการ / ส่งออก) · staff = พนักงานกรอกข้อมูล (+ บันทึกข้อมูลภายใน)
   admin = ผู้ดูแลระบบ (ทุกอย่าง: แก้ไข / ลบรายการ, ตั้งค่า, Log, รีเซ็ต)
   เซิร์ฟเวอร์ตรวจสิทธิ์ทุกคำสั่ง · token เก็บใน sessionStorage ของแท็บนี้เท่านั้น (ปิดแท็บ / กดล็อก / ครบ 12 ชม. = ล็อก) · รหัสผ่านไม่ถูกเก็บที่ใดเลย */
const ROLE_LV = { viewer: 1, staff: 2, admin: 3 };
const ROLE_TH = { viewer: 'ดูได้อย่างเดียว', staff: 'พนักงานกรอกข้อมูล', admin: 'ผู้ดูแลระบบ' };
const ROLE_IC = { viewer: 'fa-eye', staff: 'fa-user-pen', admin: 'fa-user-shield' };
const ROLE_CAN = { viewer: 'ดูรายการข้อมูลภายใน · ส่งออก Excel', staff: 'ดู + บันทึกข้อมูลภายใน', admin: 'ทุกอย่าง · แก้ไข / ลบ · ตั้งค่า · Log · รีเซ็ต' };
const SESS_KEY = 'wthep_sess';
const ss = (fn, k, v) => { try { return fn === 'get' ? sessionStorage.getItem(k) : fn === 'del' ? sessionStorage.removeItem(k) : sessionStorage.setItem(k, v); } catch { return null; } };
function sessLoad() {
  try {
    const v = JSON.parse(ss('get', SESS_KEY) || 'null');
    if (v && /^[0-9a-f]{64}$/.test(v.t) && ROLE_LV[v.role] && v.exp > Date.now()) return v;
  } catch {}
  ss('del', SESS_KEY);
  return null;
}
S.sess = sessLoad();
const can = role => !!S.sess && ROLE_LV[S.sess.role] >= ROLE_LV[role];
// random id of this browser: only for the per-browser lockout of wrong passwords
const CLIENT_ID = (() => {
  let id = store('get', 'wthep_client') || '';
  if (!/^[0-9a-f]{32}$/.test(id)) {
    try { id = Array.from(crypto.getRandomValues(new Uint8Array(16)), b => b.toString(16).padStart(2, '0')).join(''); } catch { id = ''; }
    if (id) store('set', 'wthep_client', id);
  }
  return id;
})();
/* รายการข้อมูลภายใน (TR-) แสดงเฉพาะหน้า “ข้อมูลภายใน” ที่ปลดล็อกแล้ว — หน้าอื่นขอให้เซิร์ฟเวอร์ซ่อน
   ยอดรวม “บำบัดเอง” (แดชบอร์ด / คงเหลือ) เซิร์ฟเวอร์ส่งให้เฉพาะคำขอที่มี token — ยังไม่ปลดล็อก = ไม่มีการ์ดนี้ */
const treatOn = () => can('viewer') && currentPage() === 'treat';

// Actions that change data: the user's name (top-right) must be set first.
const NAME_REQUIRED = ['store', 'ship', 'opening', 'treat', 'movement_update', 'movement_delete', 'type_save', 'receiver_save', 'receiver_delete', 'type_delete', 'reset_data'];
async function gasPost(payload) {
  if (!API_URL) throw new Error(NO_API_MSG);
  if (NAME_REQUIRED.includes(payload.action) && !S.user) {
    await askUserName(true);
    if (!S.user) throw new Error('กรุณาระบุชื่อผู้ใช้งาน (มุมขวาบน) ก่อนบันทึก');
  }
  const body = { actor: S.user || '', ...payload };
  if (S.sess) {
    body.token = S.sess.t; // in the POST body only, never in a URL
    if (!treatOn()) body.hideTreat = true;
  }
  try {
    return await postJson(body);
  } catch (e) {
    if (e.code === 'AUTH_REQUIRED' && S.sess) sessionEnded('หมดเวลาการปลดล็อก (หรือรหัสผ่านถูกเปลี่ยน) — กรุณาใส่รหัสผ่านใหม่');
    throw e;
  }
}
// Reads: POST (params in the body) — public reads need no token.
const qsObj = params => Object.fromEntries(new URLSearchParams(String(params || '').replace(/^&/, '')));
function readApi(action, params = '') {
  return gasPost({ ...qsObj(params), action });
}
async function api(action, body, params = '') {
  if (!body) return readApi(action, params);
  return gasPost({ ...body, action });
}

/* ---------- ใส่รหัสผ่าน (ปลดล็อก) ----------
   need = ระดับต่ำสุดที่ต้องใช้ · resolve(true) เมื่อได้สิทธิ์พอแล้ว, false เมื่อกดยกเลิก */
function askUnlock(need = 'viewer', why = '') {
  if (can(need)) return Promise.resolve(true);
  if (S.unlockPending) return S.unlockPending;
  S.unlockPending = new Promise(resolve => {
    const box = $('#pwModal'), card = $('#pwCard');
    const treat = need !== 'admin';
    card.className = 'modal-card unlock-card';
    card.innerHTML = `${treat ? '<div class="treat-ic unlock-ic" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-lock"></i></span></div>' : ecoLogo('unlock-ic')}
      <h3 id="ulTitle">${treat ? 'ข้อมูลภายใน' : 'รหัสผ่านผู้ดูแลระบบ'}</h3>
      <p class="hint ul-why">${esc(why || (need === 'staff' ? 'การบันทึกข้อมูลภายในต้องใช้รหัสผ่าน “พนักงานกรอกข้อมูล” หรือ “ผู้ดูแลระบบ”' : treat ? 'เมนูนี้ล็อกไว้ — ใส่รหัสผ่านเพื่อเปิด' : 'การแก้ไข / ลบรายการ ตั้งค่า Log และรีเซ็ตข้อมูล ต้องใช้รหัสผ่านผู้ดูแลระบบ'))}</p>
      <form id="ulForm" autocomplete="off">
        <div class="lg-pw"><i class="fa-solid fa-key"></i><input type="password" id="ulPw" autocomplete="current-password" placeholder="ใส่รหัสผ่าน" aria-labelledby="ulTitle"><button type="button" id="ulEye" aria-label="แสดงรหัสผ่าน"><i class="fa-regular fa-eye"></i></button></div>
        <div class="lg-err" id="ulErr" role="alert"></div>
        <div class="modal-acts"><button class="btn btn-ghost" type="button" id="ulCancel">ยกเลิก</button><button class="btn ${treat ? 'btn-treat' : 'btn-primary'}" id="ulGo" type="submit"><i class="fa-solid fa-lock-open"></i>ปลดล็อก</button></div>
      </form>
      <div class="lg-roles">${Object.keys(ROLE_LV).map(r => `<span class="${ROLE_LV[r] >= ROLE_LV[need] ? '' : 'dim'}"><i class="fa-solid ${ROLE_IC[r]}"></i><b>${ROLE_TH[r]}</b><small>${ROLE_CAN[r]}</small></span>`).join('')}</div>
      <p class="lg-hint">ปลดล็อกเฉพาะแท็บนี้ · ล็อกเองเมื่อกดปุ่มล็อก ปิดแท็บ หรือครบ 12 ชั่วโมง</p>`;
    box.hidden = false;
    const inp = $('#ulPw', card), go = $('#ulGo', card), errEl = $('#ulErr', card);
    if (!API_URL) errEl.textContent = NO_API_MSG;
    const done = v => { box.hidden = true; box.onclick = null; card.innerHTML = ''; card.className = 'modal-card'; resolve(v); };
    $('#ulEye', card).onclick = () => { inp.type = inp.type === 'password' ? 'text' : 'password'; inp.focus(); };
    $('#ulCancel', card).onclick = () => done(false);
    box.onclick = e => { if (e.target === box) done(false); };
    inp.onkeydown = e => { e.stopPropagation(); if (e.key === 'Escape') done(false); };
    $('#ulForm', card).onsubmit = async e => {
      e.preventDefault();
      if (go.disabled) return; // Enter again while checking: one request (each wrong one counts toward the lockout)
      const pw = inp.value;
      if (!pw) { errEl.textContent = 'กรุณาใส่รหัสผ่าน'; inp.focus(); return; }
      if (!API_URL) { errEl.textContent = NO_API_MSG; return; }
      go.disabled = true; go.innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังตรวจสอบ...'; errEl.textContent = '';
      const idle = () => { go.disabled = false; go.innerHTML = '<i class="fa-solid fa-lock-open"></i>ปลดล็อก'; };
      try {
        const r = await postJson({ action: 'login', password: pw, client: CLIENT_ID });
        inp.value = '';
        if (!S.sess || ROLE_LV[r.role] >= ROLE_LV[S.sess.role]) {
          const old = S.sess && S.sess.t;
          S.sess = { t: r.token, role: r.role, exp: r.expires_at };
          ss('set', SESS_KEY, JSON.stringify(S.sess));
          if (old) postJson({ action: 'logout', token: old }).catch(() => {});
        } else postJson({ action: 'logout', token: r.token }).catch(() => {}); // a lower level never replaces a higher one
        rcClear(); paintRole();
        if (!can(need)) {
          errEl.textContent = `รหัสนี้ได้สิทธิ์ “${ROLE_TH[r.role]}” — ต้องเป็น “${ROLE_TH[need]}” ขึ้นไป`;
          idle(); inp.focus();
          return;
        }
        toast(`ปลดล็อกแล้ว — สิทธิ์: ${ROLE_TH[S.sess.role]}`, 'info');
        done(true);
        if (TOTALS_PAGES.includes(currentPage())) route(); // the totals appear at once
      } catch (err) {
        errEl.textContent = err.message; inp.value = ''; inp.focus(); idle();
      }
    };
    setTimeout(() => inp.focus(), 50);
  }).finally(() => { S.unlockPending = null; });
  return S.unlockPending;
}
// Session gone (locked, expired, password changed): back to the open (public) view.
const TOTALS_PAGES = ['dashboard', 'balance']; // pages whose treatment totals depend on the session
function sessionEnded(msg, noRoute) {
  if (!S.sess) return;
  S.sess = null;
  ss('del', SESS_KEY);
  rcClear();
  paintRole();
  if (msg) toast(msg, 'info');
  if (pageNeed()) { closeModal(); route(); }
  else if (!noRoute && TOTALS_PAGES.includes(currentPage())) route(); // the totals disappear at once
}
function lockNow() {
  const tok = S.sess && S.sess.t;
  sessionEnded('ล็อกแล้ว — ต้องใส่รหัสผ่านใหม่เพื่อเปิดข้อมูลภายใน');
  if (tok) postJson({ action: 'logout', token: tok }).catch(() => {});
}
// the level an admin tool needs: ask for the password when this tab does not have it yet
async function needRole(role, why) {
  if (can(role)) return true;
  const ok = await askUnlock(role, why);
  if (!ok) toast(`ต้องใช้รหัสผ่าน “${ROLE_TH[role]}”`, 'err');
  return ok;
}
/* ส่งออก Excel (CSV, UTF-8 BOM ให้ Excel ภาษาไทยอ่านถูก) */
async function downloadCsv(action, params) {
  try {
    const r = await readApi(action, params);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob(['\uFEFF' + r.csv], { type: 'text/csv;charset=utf-8' }));
    a.download = r.filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  } catch (e) { toast(e.message, 'err'); }
}
if (!API_URL) $('#cfgBanner').hidden = false;
// false: the session of this tab has ended on the server (the page is drawn again, locked)
function applyBoot(j) {
  Object.assign(S, { types: j.types, receivers: j.receivers || [], balance: j.balance, containers: j.containers || {}, company: j.company || {}, limits: j.limits || {}, now: j.now });
  // the server is the authority on the level of this session (null: no / expired session)
  if (S.sess && j.role === null) {
    const p = currentPage();
    sessionEnded('หมดเวลาการปลดล็อก (หรือรหัสผ่านถูกเปลี่ยน) — กรุณาใส่รหัสผ่านใหม่', true);
    return !pageNeed(p);
  }
  if (S.sess && j.role && j.role !== S.sess.role && ROLE_LV[j.role]) {
    S.sess.role = j.role; ss('set', SESS_KEY, JSON.stringify(S.sess)); paintRole();
  }
  return true;
}

/* ---------- อ่านหลายรายการในคำขอเดียว (batch) ---------- */
const readsSpec = reads => reads.map(([a, p = '']) => ({ action: a, ...qsObj(p) }));
const readsField = reads => reads.length ? { reads: readsSpec(reads) } : {};
async function readMany(reads) {
  const r = await gasPost({ action: 'batch', reads: readsSpec(reads) });
  return r.results.map(d => ({ ok: true, ...d }));
}

/* ---------- แคชผลการอ่านล่าสุดของแท็บนี้ (แสดงทันที แล้วโหลดใหม่เบื้องหลัง) ----------
   sessionStorage เท่านั้น (หายเมื่อปิดแท็บ) · ใช้เพื่อแสดงผลเท่านั้น — การบันทึกตรวจยอดที่เซิร์ฟเวอร์เสมอ
   Log ไม่เก็บ · ล้างทั้งหมดหลังบันทึก / ปลดล็อก / ล็อก · หน้าข้อมูลภายในไม่อ่าน/เขียนแคชนี้เลย */
const RC_PREFIX = 'wthep_rc:';
const RC_FRESH_MS = 5000;
const rcKey = ([a, p = '']) => RC_PREFIX + a + p;
function rcGet(r) {
  if (treatOn()) return null;
  try { const v = JSON.parse(sessionStorage.getItem(rcKey(r)) || 'null'); return v && v.d ? v : null; } catch { return null; }
}
function rcPut(r, d) {
  if (!d || treatOn()) return;
  try { const s = JSON.stringify({ t: Date.now(), d }); if (s.length < 500000) sessionStorage.setItem(rcKey(r), s); } catch {}
}
function rcClear() {
  for (const st of [sessionStorage, localStorage]) { // localStorage: older cached reads (same wthep_ prefix only)
    try { Object.keys(st).filter(k => k.startsWith(RC_PREFIX)).forEach(k => st.removeItem(k)); } catch {}
  }
}
function afterWrite(r, reads) {
  rcClear();
  if (Array.isArray(r?.reads) && r.reads.length === reads.length) r.reads.forEach((d, i) => rcPut(reads[i], { ok: true, ...d }));
}
async function writeApi(action, body, reads = routeReads()) {
  const r = await api(action, { ...body, ...readsField(reads) });
  afterWrite(r, reads);
  return r;
}

/* ---------- helpers ---------- */
const fmt = (n, d = 0) => Number(n || 0).toLocaleString('th-TH', { maximumFractionDigits: d });
const fmtFix = (n, d = 0) => Number(n || 0).toLocaleString('th-TH', { minimumFractionDigits: d, maximumFractionDigits: d });
/* หน่วยแสดงผล = ตัน (ชีตเก็บเป็น กก.) — แสดงทศนิยม 2 ตำแหน่ง (ยอดจริงในชีตยังเก็บละเอียดถึง กก.) */
const fmtT = kg => (Number(kg || 0) / 1000).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtKgHint = kg => Number(kg || 0).toLocaleString('th-TH', { maximumFractionDigits: 2 });
// kg -> exact ton text for an input (no separators, 2+ decimals), e.g. 1234.56 -> "1.23456", 100 -> "0.10", 1000 -> "1.00"
const tonsStr = kg => {
  const c = Math.round(Math.abs(Number(kg || 0)) * 100);
  const f = String(c % 100000).padStart(5, '0').replace(/0+$/, '');
  return Math.floor(c / 100000) + '.' + f.padEnd(2, '0'); // at least 2 decimals (more only for a fraction of a kg)
};
// Weight typed in tons (the page shows 2 decimals; up to 5 are still read for older kg data), returns kg or NaN.
const parseT = v => { const s = String(v ?? '').replace(/,/g, '').trim(); return /^\d+(\.\d{1,5})?$/.test(s) ? Math.round(Number(s) * 100000) / 100 : NaN; };
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const typeById = id => S.types.find(t => +t.id === +id);
const balOf = id => S.balance[id] || { balance_kg: 0, containers: { drum200: 0, ibc: 0, bigbag: 0 }, shipped_year_kg: 0, permit_kg: 0, quota_pct: null, quota_level: 'ok' };
const dtTH = s => { const d = new Date(s.slice(0, 10) + 'T00:00:00'); return d.toLocaleDateString('th-TH', { day: 'numeric', month: 'short', year: '2-digit' }) + ' ' + s.slice(11, 16) + ' น.'; };
const dateTH = s => new Date(s.slice(0, 10) + 'T00:00:00').toLocaleDateString('th-TH', { day: 'numeric', month: 'short', year: 'numeric' });
const nowParts = () => { const d = new Date(); const p = n => String(n).padStart(2, '0'); return { date: `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`, time: `${p(d.getHours())}:${p(d.getMinutes())}` }; };
// receivers that have waste types, in the order of the settings list (forms, filters, balance page)
const receivers = () => {
  const names = [...new Set(S.types.map(t => t.receiver))], pos = n => { const i = (S.receivers || []).findIndex(r => r.name === n); return i < 0 ? 9999 : i; };
  return names.sort((a, b) => pos(a) - pos(b));
};
const shortRecv = r => String(r || '').replace(/^(บจก\.|บมจ\.|บริษัท)\s*/, '');
// Empty install: no waste type yet. Friendly notice + link to the (admin) settings page.
const noTypesBox = (why = '') => `<div class="card empty-setup" style="text-align:center;padding:34px 20px"><div style="font-size:42px;color:var(--brand)"><i class="fa-solid fa-boxes-stacked"></i></div>
  <h3 style="margin:10px 0 6px;justify-content:center">ยังไม่มีประเภทของเสีย</h3>
  <p class="hint" style="max-width:560px;margin:0 auto 16px">ยังไม่มีประเภทของเสีย — ไปที่ <b>ตั้งค่าประเภทของเสีย (ผู้ดูแลระบบ)</b> เพื่อเพิ่มประเภทและผู้รับดำเนินการ${why ? '<br>' + why : ''}</p>
  <a class="btn btn-primary" href="#settings"><i class="fa-solid fa-gear"></i>ไปที่ ตั้งค่าประเภทของเสีย</a></div>`;
const typeLabel = t => t ? `${t.code} · ${t.name}` : '—';
const lvName = { ok: 'ปกติ', warn: 'ใกล้เต็มโควตา', over: 'เกินโควตา' };

/* ---------- container illustrations (inline SVG) ---------- */
function containerSVG(k, mono = false) {
  const W = mono ? '#fff' : null;
  if (k === 'drum200') {
    const body = W || '#0D9488', top = mono ? 'rgba(255,255,255,.75)' : '#5EEAD4', band = mono ? 'rgba(4,120,87,.55)' : '#115E59';
    return `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M15 12v40c0 3 7.6 5.5 17 5.5S49 55 49 52V12" fill="${body}"/>
      <path d="M15 25c0 3 7.6 5.5 17 5.5S49 28 49 25M15 40c0 3 7.6 5.5 17 5.5S49 43 49 40" stroke="${band}" stroke-width="2.4" fill="none"/>
      <ellipse cx="32" cy="12" rx="17" ry="5.5" fill="${top}"/><ellipse cx="38" cy="11.5" rx="3" ry="1.2" fill="${band}"/>
      <rect x="20" y="16" width="3.5" height="34" rx="1.7" fill="#fff" opacity=".28"/></svg>`;
  }
  if (k === 'ibc') {
    const cage = mono ? '#fff' : '#64748B', tank = mono ? 'rgba(255,255,255,.22)' : '#F1F5F9', liq = mono ? 'rgba(255,255,255,.7)' : '#F59E0B', base = mono ? 'rgba(255,255,255,.85)' : '#A16207';
    return `<svg viewBox="0 0 64 64" aria-hidden="true"><rect x="11" y="12" width="42" height="38" rx="4" fill="${tank}"/>
      <rect x="13" y="29" width="38" height="19" rx="3" fill="${liq}"/>
      <path d="M11 12h42v38H11zM25 12v38M39 12v38M11 25h42M11 38h42" stroke="${cage}" stroke-width="2" fill="none"/>
      <rect x="27" y="7" width="10" height="5" rx="1.5" fill="${cage}"/><rect x="8" y="50" width="48" height="5" rx="1.5" fill="${base}"/>
      <rect x="10" y="55" width="6" height="4" fill="${base}"/><rect x="29" y="55" width="6" height="4" fill="${base}"/><rect x="48" y="55" width="6" height="4" fill="${base}"/></svg>`;
  }
  const bag = mono ? 'rgba(255,255,255,.9)' : '#ECFDF5', edge = mono ? '#fff' : '#059669', loop = mono ? '#FDE68A' : '#F59E0B';
  return `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M17 17c0-7 6-9 6-9M47 17c0-7-6-9-6-9" stroke="${loop}" stroke-width="3.5" fill="none" stroke-linecap="round"/>
    <path d="M14 17h36l4 36c.3 3-1.5 5-4 5H14c-2.5 0-4.3-2-4-5z" fill="${bag}" stroke="${edge}" stroke-width="2.2"/>
    <path d="M14 17l-1 9M50 17l1 9" stroke="${edge}" stroke-width="1.6"/><path d="M21 30h22M20 38h24M19 46h26" stroke="${edge}" stroke-width="1.4" stroke-dasharray="3 3" opacity=".7"/>
    <circle cx="32" cy="38" r="6" fill="none" stroke="${loop}" stroke-width="2"/><path d="M29.5 38l2 2 3.5-4" stroke="${loop}" stroke-width="2" fill="none"/></svg>`;
}
const ecoLogo = (cls = '') => `<div class="eco-logo ${cls}" aria-hidden="true"><span class="ring"></span><span class="core"><i class="fa-solid fa-recycle"></i></span></div>`;
const loader = (msg = 'กำลังโหลดข้อมูล...') => `<div class="loader">${ecoLogo()}<small>${msg}</small></div>`;

/* ---------- UI effects ---------- */
document.addEventListener('pointerdown', e => {
  const b = e.target.closest('.btn, .tpick, .quick button, .seg button');
  if (!b) return;
  const r = b.getBoundingClientRect(), d = Math.max(r.width, r.height);
  const sp = document.createElement('span');
  sp.className = 'ripple';
  Object.assign(sp.style, { width: d + 'px', height: d + 'px', left: e.clientX - r.left - d / 2 + 'px', top: e.clientY - r.top - d / 2 + 'px' });
  if (getComputedStyle(b).position === 'static') b.style.position = 'relative';
  b.style.overflow = 'hidden';
  b.appendChild(sp);
  setTimeout(() => sp.remove(), 700);
});
function toast(msg, type = 'ok') {
  const m = { ok: ['fa-circle-check', '#10B981'], err: ['fa-circle-xmark', '#EF4444'], info: ['fa-recycle', '#0D9488'] }[type];
  const t = document.createElement('div');
  t.className = 'toast';
  t.style.setProperty('--c', m[1]);
  t.innerHTML = `<i class="fa-solid ${m[0]}"></i><div>${esc(msg)}</div>`;
  $('#toasts').appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 400); }, type === 'err' ? 5500 : 3500);
}
const reduceMotion = () => matchMedia('(prefers-reduced-motion: reduce)').matches;
function confetti() {
  if (reduceMotion()) return;
  const cv = $('#confetti'), cx = cv.getContext('2d');
  cv.width = innerWidth; cv.height = innerHeight;
  const cols = ['#10B981', '#14B8A6', '#F59E0B', '#34D399', '#047857', '#FCD34D', '#6EE7B7'];
  const ps = Array.from({ length: 150 }, () => ({
    x: innerWidth / 2 + (Math.random() - .5) * 200, y: innerHeight / 2.2,
    vx: (Math.random() - .5) * 18, vy: -Math.random() * 18 - 6, s: Math.random() * 8 + 5,
    r: Math.random() * 6, vr: (Math.random() - .5) * .4, c: cols[Math.random() * cols.length | 0], leaf: Math.random() < .35,
  }));
  let f = 0;
  const run = S.confettiRun = (S.confettiRun || 0) + 1;
  setTimeout(() => { if (S.confettiRun === run) { S.confettiRun++; cx.clearRect(0, 0, cv.width, cv.height); } }, 2600);
  (function loop() {
    if (S.confettiRun !== run) return;
    cx.clearRect(0, 0, cv.width, cv.height);
    ps.forEach(p => {
      p.vy += .45; p.vx *= .99; p.x += p.vx; p.y += p.vy; p.r += p.vr;
      cx.save(); cx.translate(p.x, p.y); cx.rotate(p.r); cx.fillStyle = p.c; cx.globalAlpha = Math.max(0, 1 - f / 140);
      if (p.leaf) { cx.beginPath(); cx.ellipse(0, 0, p.s / 1.4, p.s / 3, 0, 0, Math.PI * 2); cx.fill(); } else cx.fillRect(-p.s / 2, -p.s / 4, p.s, p.s / 2);
      cx.restore();
    });
    if (++f < 140) requestAnimationFrame(loop); else cx.clearRect(0, 0, cv.width, cv.height);
  })();
}
function modal(html, wide = false) {
  $('#modalCard').className = 'modal-card' + (wide ? ' wide' : '');
  $('#modalCard').innerHTML = html;
  $('#modal').hidden = false;
  return $('#modalCard');
}
function closeModal() {
  if ($('#modal').hidden) return;
  $('#modal').hidden = true;
  const fn = S.onModalClose; S.onModalClose = null;
  if (fn) fn();
}
$('#modal').addEventListener('click', e => { if (e.target.id === 'modal' || e.target.closest('[data-close]')) closeModal(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });
function countUp(root = view) {
  $$('[data-count]', root).forEach(el => {
    const to = +el.dataset.count, dec = +(el.dataset.dec || 0), t0 = performance.now(), dur = reduceMotion() ? 1 : 1200;
    const step = t => {
      const p = Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - p, 4);
      el.textContent = fmtFix(p < 1 ? to * e : to, dec);
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  });
}
function killCharts() { S.charts.forEach(c => c.destroy()); S.charts = []; }
document.addEventListener('animationend', e => {
  if (e.animationName === 'rise' && e.target.parentElement?.classList.contains('stagger')) {
    e.target.style.animation = 'none'; e.target.style.opacity = '1';
  }
});

/* clock */
setInterval(tickClock, 1000); tickClock();
function tickClock() {
  const d = new Date();
  $('#clock').textContent = d.toLocaleTimeString('th-TH', { hour12: false });
  $('#clockDate').textContent = d.toLocaleDateString('th-TH', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
}
$('#menuBtn').onclick = () => $('#sidebar').classList.toggle('open');
$('#navOverlay').onclick = () => $('#sidebar').classList.remove('open');

/* ================= ผู้ใช้งาน (ชื่อมุมขวาบน) ================= */
function paintUser() {
  $('#userName').textContent = S.user || 'ระบุชื่อผู้ใช้';
  $('#userAv').innerHTML = S.user ? esc(S.user.trim().charAt(0)) : '<i class="fa-solid fa-user"></i>';
  $('#userChip').classList.toggle('need', !S.user);
}
/* สถานะการปลดล็อก: ป้ายสิทธิ์ + ปุ่มล็อกมุมขวาบน, กุญแจบนเมนูที่ต้องใช้รหัส (เซิร์ฟเวอร์ตรวจสิทธิ์ซ้ำทุกคำสั่งอยู่แล้ว) */
function paintRole() {
  const r = S.sess ? S.sess.role : '';
  $$('.nav [data-lock]').forEach(el => {
    const locked = !can(el.dataset.lock);
    el.classList.toggle('locked', locked);
    el.title = locked ? `ต้องใส่รหัสผ่าน “${ROLE_TH[el.dataset.lock]}” ขึ้นไป` : '';
  });
  $('#roleTag').innerHTML = r ? `<i class="fa-solid ${ROLE_IC[r]}"></i>${ROLE_TH[r]}` : '';
  $('#roleTag').className = 'role-tag r-' + (r || 'none');
  $('#roleTag').hidden = !r;
  $('#logoutBtn').hidden = !r;
  document.body.dataset.level = r; // not data-role: that marks the elements shown by level
  paintUser();
}
// required = opened because a save needs a name: keep the form, resolve when closed.
function askUserName(required = false) {
  return new Promise(resolve => {
    const m = modal(`<h3><i class="fa-solid fa-user-pen" style="color:var(--brand)"></i> ผู้ใช้งาน</h3>
      <p class="hint">${required ? '<b style="color:var(--bad)">กรุณาระบุชื่อผู้ใช้งานก่อนบันทึกรายการ</b> · ' : ''}ชื่อนี้จะถูกบันทึกเป็น “ผู้บันทึก” ในประวัติและ Log ทุกรายการ</p>
      <div class="field mt"><label>ชื่อ - นามสกุล</label><input id="uName" value="${esc(S.user)}" placeholder="เช่น สมชาย ใจดี" maxlength="60"></div>
      <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn btn-primary" id="uSave"><i class="fa-solid fa-check"></i>บันทึก</button></div>`);
    const inp = $('#uName', m); inp.focus(); inp.select();
    let done = false;
    const finish = () => { if (!done) { done = true; resolve(S.user); } };
    const save = () => {
      const name = inp.value.trim().slice(0, 60);
      if (required && !name) { inp.focus(); return; }
      S.user = name; store('set', 'wthep_user', S.user); paintUser(); closeModal(); finish();
      toast(S.user ? `สวัสดีคุณ${S.user}` : 'ล้างชื่อผู้ใช้แล้ว', 'info');
      if (!required) route();
    };
    $('#uSave', m).onclick = save;
    inp.onkeydown = e => { if (e.key === 'Enter') save(); };
    S.onModalClose = finish;
  });
}
$('#userChip').onclick = () => askUserName(false);
$('#logoutBtn').onclick = lockNow;

/* ================= ROUTER ================= */
const monthFirst = () => nowParts().date.slice(0, 8) + '01';
// ข้อมูลภายใน: since 1 January (in-house rows are few); history: this month
const histDefault = (kind = '') => '&' + new URLSearchParams({ from: kind === 'treat' ? nowParts().date.slice(0, 5) + '01-01' : monthFirst(), to: nowParts().date, type: '', kind, receiver: '', q: '' });
const hashParam = k => new URLSearchParams(location.hash.split('?')[1] || '').get(k) || '';
// ข้อมูลภายใน: #treat = รายการ (ทุกระดับ) · #treat?tab=new = ฟอร์มบันทึก (พนักงานกรอกข้อมูลขึ้นไป)
const treatTab = () => hashParam('tab') === 'new' ? 'new' : 'list';
const currentPage = () => { const page = (location.hash.slice(1) || 'dashboard').split('?')[0]; return PAGES[page] ? page : 'dashboard'; };
function pageReads(p) {
  if (p === 'treat') return treatTab() === 'list' ? [['history', histDefault('treat')]] : [];
  return ({ dashboard: [['dashboard', '']], history: [['history', histDefault()]] })[p] || [];
}
const routeReads = (p = currentPage()) => [['bootstrap', ''], ...pageReads(p)];
const PAGE_FN = {
  dashboard: ([d]) => dashboard(d), history: ([h]) => historyPage(h), balance, logs, settings,
  store: () => txForm('store'), ship: () => txForm('ship'), treat: ([h]) => treatPage(h),
};
['input', 'change'].forEach(ev => view.addEventListener(ev, () => { S.dirty = true; }, true));
view.addEventListener('pointerdown', e => { if (e.target.closest('button, input, select, textarea, label, .ccard, .tpick')) S.dirty = true; }, true);
const sigOf = res => JSON.stringify(res.map((d, i) => (i ? d : { ...d, now: null })));
let routeSeq = 0;
// the level a page needs ('' = open to everyone) — the server checks every action again
const PAGE_ROLE = { treat: 'viewer', logs: 'admin', settings: 'admin' };
const pageNeed = (p = currentPage()) => (p === 'treat' && treatTab() === 'new' ? 'staff' : PAGE_ROLE[p] || '');
const NAME_PAGES = ['store', 'ship'];
const INPUT_PAGES = ['store', 'ship', 'treat', 'history', 'settings'];
function setHead(p) {
  $$('.nav a').forEach(a => a.classList.toggle('active', a.dataset.page === p));
  $('#pageTitle').textContent = PAGES[p][0];
  $('#pageSub').textContent = PAGES[p][1];
  document.title = PAGES[p][0] + ' · Balance Waste';
  $('#sidebar').classList.remove('open');
}
// a page that needs a password this tab does not have: a locked card + the password box
function lockedPage(p, need) {
  const treat = p === 'treat';
  view.innerHTML = `<div class="card locked-page">
    ${treat ? '<div class="treat-ic lock-ic" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-lock"></i></span></div>' : ecoLogo('lock-ic')}
    <h2>${esc(PAGES[p][0])} <span class="tag lk-tag"><i class="fa-solid fa-lock"></i>ล็อกอยู่</span></h2>
    <p>${treat ? (need === 'staff' ? 'การบันทึกข้อมูลภายในต้องใช้รหัสผ่าน “พนักงานกรอกข้อมูล” หรือ “ผู้ดูแลระบบ”' : 'เมนูนี้ต้องใส่รหัสผ่านก่อนเปิด — ยอดรวม “บำบัดเอง” บนแดชบอร์ดจะแสดงหลังปลดล็อกเท่านั้น')
      : 'หน้านี้ต้องใช้รหัสผ่าน “ผู้ดูแลระบบ”'}${S.sess ? `<br><small>สิทธิ์ปัจจุบันของแท็บนี้: <b>${ROLE_TH[S.sess.role]}</b></small>` : ''}</p>
    <div class="lock-levels">${Object.keys(ROLE_LV).map(r => `<div class="ll r-${r} ${ROLE_LV[r] >= ROLE_LV[need] ? '' : 'dim'}"><i class="fa-solid ${ROLE_IC[r]}"></i><b>${ROLE_TH[r]}</b><small>${ROLE_CAN[r]}</small></div>`).join('')}</div>
    <div class="modal-acts" style="justify-content:center"><a class="btn btn-ghost" href="${treat && need === 'staff' && can('viewer') ? '#treat' : '#dashboard'}"><i class="fa-solid fa-arrow-left"></i>${treat && need === 'staff' && can('viewer') ? 'กลับไปรายการ' : 'กลับแดชบอร์ด'}</a>
      <button class="btn ${treat ? 'btn-treat' : 'btn-primary'}" id="lkGo" type="button"><i class="fa-solid fa-key"></i>ใส่รหัสผ่าน</button></div>
  </div>`;
  $('#lkGo').onclick = async () => { if (await askUnlock(need) && currentPage() === p) route(); };
}
async function route() {
  const p = currentPage();
  const need = pageNeed(p);
  if (need && !can(need)) {
    ++routeSeq;
    killCharts(); view.onclick = null; S.onFresh = null;
    setHead(p);
    lockedPage(p, need);
    if (await askUnlock(need) && currentPage() === p && pageNeed() === need) {
      // logging in through the ข้อมูลภายใน menu lands on the dashboard (it now shows the totals); the other locked pages open as they are
      if (p === 'treat') location.hash = 'dashboard'; else route();
    }
    return;
  }
  if ((NAME_PAGES.includes(p) || need === 'staff') && !S.user) {
    await askUserName(true);
    if (!S.user) { toast('กรุณาระบุชื่อผู้ใช้งาน (มุมขวาบน) ก่อนทำรายการ', 'err'); if (currentPage() === p) location.hash = p === 'treat' ? 'treat' : 'dashboard'; return; }
  }
  if (currentPage() !== p) return;
  setHead(p);
  const seq = ++routeSeq;
  const reads = routeReads(p);
  const hits = reads.map(rcGet);
  const cached = hits.every(Boolean) ? hits.map(h => h.d) : null;
  const paint = async res => {
    killCharts(); view.onclick = null; S.onFresh = null; S.dirty = false;
    if (!applyBoot(res[0])) return; // the session ended: the page was drawn again, locked
    updateBadge();
    await PAGE_FN[p](res.slice(1));
    countUp();
  };
  killCharts(); view.onclick = null; S.onFresh = null;
  view.innerHTML = loader();
  view.style.animation = 'none'; view.offsetHeight; view.style.animation = '';
  let shown = null;
  if (cached) {
    try { await paint(cached); shown = sigOf(cached); } catch { view.innerHTML = loader(); }
    if (shown && hits.every(h => Date.now() - h.t < RC_FRESH_MS)) return;
  }
  try {
    const res = await readMany(reads);
    if (seq !== routeSeq) return;
    res.forEach((d, i) => rcPut(reads[i], d));
    if (shown === sigOf(res)) { applyBoot(res[0]); return; }
    if (shown !== null && ((S.dirty && INPUT_PAGES.includes(p)) || p === 'logs')) { if (applyBoot(res[0])) { updateBadge(); S.onFresh?.(); } return; }
    await paint(res);
  } catch (e) {
    if (seq !== routeSeq) return;
    if (shown !== null) { toast(e.message, 'err'); return; }
    view.innerHTML = `<div class="card empty"><i class="fa-solid fa-plug-circle-xmark"></i>${esc(e.message)}</div>`;
  }
}
addEventListener('hashchange', () => route());
function alertCount() {
  const days = S.limits.storageAlertDays || 90;
  let n = 0, bad = false;
  S.types.forEach(t => {
    const b = balOf(t.id);
    if (b.quota_level === 'over') { n++; bad = true; } else if (b.quota_level === 'warn') n++;
    if (b.age_days != null && b.age_days > days) n++;
  });
  return { n, bad };
}
function updateBadge() {
  const { n, bad } = alertCount();
  const el = $('#alertBadge');
  el.textContent = n ? fmt(n) : '';
  el.classList.toggle('bad', bad);
}

/* ================= DASHBOARD ================= */
function needleGauge(pct) {
  const cx = 150, cy = 150, R = 118, MAX = 120;
  const ang = v => (-90 + Math.min(MAX, Math.max(0, v)) / MAX * 180) * Math.PI / 180;
  const pt = (v, r) => [cx + r * Math.sin(ang(v)), cy - r * Math.cos(ang(v))];
  const arc = (a, b, col) => { const [x1, y1] = pt(a, R), [x2, y2] = pt(b, R); return `<path d="M${x1.toFixed(1)} ${y1.toFixed(1)} A${R} ${R} 0 0 1 ${x2.toFixed(1)} ${y2.toFixed(1)}" stroke="${col}" stroke-width="20" fill="none" stroke-linecap="butt"/>`; };
  let ticks = '';
  for (let v = 0; v <= MAX; v += 10) {
    const [x1, y1] = pt(v, R - 16), [x2, y2] = pt(v, R - (v % 50 ? 22 : 28));
    ticks += `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="#94A3B8" stroke-width="${v % 50 ? 1.5 : 2.5}"/>`;
    if (!(v % 20)) { const [tx, ty] = pt(v, R - 40); ticks += `<text x="${tx.toFixed(1)}" y="${(ty + 4).toFixed(1)}" text-anchor="middle" font-size="11" fill="#64748B" font-family="Kanit">${v}</text>`; }
  }
  const deg = -90 + Math.min(MAX, Math.max(0, pct || 0)) / MAX * 180;
  return `<svg class="needle-gauge" viewBox="0 0 300 172" role="img" aria-label="ใช้โควตา ${fmt(pct, 1)}%">
    <defs><linearGradient id="ng" x1="0" x2="1"><stop offset="0" stop-color="#34D399"/><stop offset="1" stop-color="#059669"/></linearGradient></defs>
    ${arc(0, 80, 'url(#ng)')}${arc(80, 100, '#F59E0B')}${arc(100, 120, '#EF4444')}${ticks}
    <g class="tick-spin"><circle cx="150" cy="150" r="26" fill="none" stroke="#A7F3D0" stroke-width="2" stroke-dasharray="4 6"/></g>
    <g class="needle" data-deg="${deg.toFixed(1)}"><path d="M146 150 L150 44 L154 150 Z" fill="#0F2A23"/><circle cx="150" cy="150" r="5" fill="#F59E0B"/></g>
    <circle cx="150" cy="150" r="12" fill="#0F2A23"/><circle cx="150" cy="150" r="5" fill="#FCD34D"/></svg>`;
}
function ringCard(t) {
  const b = balOf(t.id), p = b.quota_pct == null ? 0 : b.quota_pct;
  return `<div class="ring-card ${b.quota_level}" style="--c:${esc(t.color)}" data-p="${Math.min(100, p)}" data-go="balance" title="${esc(typeLabel(t))} · ${esc(t.receiver)}">
    <div class="qring"><span class="spinner-arc"></span><div class="in"><i class="fa-solid ${esc(t.icon)}"></i><b>${fmt(p, 1)}%</b></div></div>
    <div class="rc-name">${esc(t.name)}</div><div class="rc-sub">${esc(t.code)} · ${fmtT(b.shipped_year_kg)} / ${fmtT(b.permit_kg)} ตัน</div></div>`;
}
function tlItem(m) {
  const t = typeById(m.type_id) || {}, k = KIND[m.kind] || KIND.store;
  return `<div class="tl-item"><div class="ti" style="--c:${k.c}"><i class="fa-solid ${k.icon}"></i></div>
    <div class="tx"><b>${esc(t.name || '—')}</b><small>${k.short} · ${esc(m.doc_no)} · ${dtTH(m.moved_at)} · ${esc(m.actor || '')}</small></div>
    <div class="kg ${k.cls}">${k.sign}${fmtT(m.weight_kg)} ตัน</div></div>`;
}
function alertItem(a) {
  const t = typeById(a.type_id) || {};
  if (a.kind === 'quota') {
    return `<div class="alert ${a.level === 'bad' ? 'bad' : ''}"><div class="ai"><i class="fa-solid ${a.level === 'bad' ? 'fa-circle-exclamation' : 'fa-triangle-exclamation'}"></i></div>
      <p><b>${esc(t.code)} ${esc(t.name)}</b> ${a.level === 'bad' ? 'ส่งกำจัดเกินปริมาณที่ได้รับอนุญาต' : 'ใช้โควตาใกล้เต็ม'} — ${fmt(a.pct, 1)}%<br>
      <small>${fmtT(a.used_kg)} / ${fmtT(a.permit_kg)} ตัน · ${esc(shortRecv(t.receiver))}</small></p></div>`;
  }
  return `<div class="alert"><div class="ai"><i class="fa-solid fa-hourglass-half"></i></div>
    <p><b>${esc(t.code)} ${esc(t.name)}</b> จัดเก็บในโรงงานนาน ${fmt(a.days)} วัน<br><small>ล็อตเก่าสุดตั้งแต่ ${dateTH(a.since)} · คงเหลือ ${fmtT(a.balance_kg)} ตัน</small></p></div>`;
}
async function dashboard(d) {
  const k = d.kpi;
  const h = new Date().getHours();
  const greet = h < 12 ? ['สวัสดีตอนเช้า', 'fa-sun'] : h < 17 ? ['สวัสดีตอนบ่าย', 'fa-cloud-sun'] : ['สวัสดีตอนเย็น', 'fa-moon'];
  const withBal = S.types.filter(t => balOf(t.id).balance_kg > 0).length;
  const lv = k.quota_pct == null ? 'ok' : k.quota_pct > 100 ? 'over' : k.quota_pct >= (S.limits.quotaWarnPct || 80) ? 'warn' : 'ok';
  const kpis = [
    ['คงเหลือในโรงงานรวม', k.balance_kg / 1000, 'ตัน', 2, 'fa-warehouse', '#059669', `${withBal} ประเภทที่มีคงเหลือ`],
    ['จัดเก็บเดือนนี้', k.month_in_kg / 1000, 'ตัน', 2, 'fa-box-archive', '#14B8A6', `ทั้งปี ${fmtT(k.year_in_kg)} ตัน`],
    ['ส่งกำจัดปีนี้', k.year_out_kg / 1000, 'ตัน', 2, 'fa-truck-arrow-right', '#F97316', `เดือนนี้ ${fmtT(k.month_out_kg)} ตัน`],
    ['% โควตาที่ใช้ (ปีนี้)', k.quota_pct || 0, '%', 1, 'fa-chart-pie', lv === 'over' ? '#EF4444' : lv === 'warn' ? '#F59E0B' : '#6366F1', `${fmtT(k.year_out_kg)} / ${fmtT(k.permit_kg)} ตัน`],
  ];
  const withPermit = S.types.filter(t => +t.active === 1 || balOf(t.id).balance_kg > 0);
  const cc = d.containers || {};
  const cTotal = CONT_KEYS.reduce((a, c) => a + (cc[c] || 0), 0);
  view.innerHTML = `
  ${S.types.length ? '' : noTypesBox()}
  <div class="reset-topbar"><a class="btn btn-primary btn-sm" href="#store"><i class="fa-solid fa-box-archive"></i>บันทึกจัดเก็บ</a><a class="btn btn-ship btn-sm" href="#ship"><i class="fa-solid fa-truck-arrow-right"></i>ส่งกำจัด</a><a class="btn btn-treat btn-sm" href="#treat"><i class="fa-solid ${can('viewer') ? 'fa-lock-open' : 'fa-lock'}"></i>ข้อมูลภายใน</a>${can('admin') ? '<button class="btn btn-bad btn-sm" id="resetBtn">🗑 รีเซ็ตข้อมูล / Reset data</button>' : ''}</div>
  <div class="hero">
    <div class="hero-orb" aria-hidden="true"><span class="o1"></span><span class="o2"></span><span class="o3"><i class="fa-solid fa-recycle"></i></span>
      <span class="sat"><b><i class="fa-solid fa-leaf"></i></b></span><span class="sat s2"><b><i class="fa-solid fa-droplet"></i></b></span><span class="sat s3"><b><i class="fa-solid fa-seedling"></i></b></span></div>
    <div class="greet"><i class="fa-solid ${greet[1]}"></i>${greet[0]}${S.user ? ' คุณ' + esc(S.user) : ''} · ${new Date().toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</div>
    <h2>Balance <em>Waste</em> · ระบบบริหารจัดการของเสีย</h2>
    <p>${esc(S.company.name || '')}</p>
    <div class="chips">${S.company.factoryReg ? `<span><i class="fa-solid fa-industry"></i>ทะเบียนโรงงาน ${esc(S.company.factoryReg)}</span>` : ''}${S.company.permitNo || S.company.permitYear ? `<span><i class="fa-solid fa-file-shield"></i>${esc([S.company.permitNo, S.company.permitYear ? 'ปี ' + S.company.permitYear : ''].filter(Boolean).join(' · '))}</span>` : ''}${k.permit_kg > 0 ? `<span><i class="fa-solid fa-scale-balanced"></i>อนุญาตรวม ${fmtT(k.permit_kg)} ตัน</span>` : ''}</div>
  </div>

  <div class="grid g-kpi mt stagger">
    ${kpis.map(([l, n, u, dec, ic, c, tr]) => `
      <div class="card kpi" style="--c:${c}">
        <div class="top"><div class="orbit" style="--c:${c}"><span class="dot"></span><span class="dot d2"></span><span class="dot d3"></span><div class="ic"><i class="fa-solid ${ic}"></i></div></div><div class="kpi-label">${l}</div></div>
        <div class="kpi-num"><span data-count="${n}" data-dec="${dec}">0</span><small>${u}</small></div><div class="trend">${tr}</div>
      </div>`).join('')}
  </div>
  ${k.treated_kg !== undefined ? treatSummary(k) : ''}

  <div class="grid g-2 mt">
    <div class="card"><h3><i class="fa-solid fa-chart-column"></i>ปริมาณจัดเก็บ / ส่งกำจัด${k.treated_kg !== undefined ? ' / บำบัดเอง' : ''} 12 เดือน<span class="sub">ตัน</span></h3><div class="chart-box"><canvas id="cMonths"></canvas></div></div>
    <div class="card"><h3><i class="fa-solid fa-gauge-high"></i>โควตารวมปีนี้<span class="sub">เทียบปริมาณที่ได้รับอนุญาต</span></h3>
      <div class="gauge-wrap">${needleGauge(k.quota_pct || 0)}<div class="gauge-val"><span data-count="${k.quota_pct || 0}" data-dec="1">0</span><small>%</small></div>
      <div class="gauge-sub">ส่งกำจัดแล้ว ${fmtT(k.year_out_kg)} ตัน จาก ${fmtT(k.permit_kg)} ตัน<br><span class="tag lv-${lv}">${lvName[lv]}</span></div></div></div>
  </div>

  <div class="sec-h"><i class="fa-solid fa-circle-notch fa-spin" style="color:var(--brand);--fa-animation-duration:3s"></i>การใช้โควตาแยกประเภท <small>ส่งกำจัดปีนี้ / ปริมาณที่ได้รับอนุญาต</small></div>
  <div class="card"><div class="rings stagger">${withPermit.map(ringCard).join('') || '<div class="empty">ยังไม่มีประเภทของเสีย</div>'}</div></div>

  <div class="grid g-2e mt">
    <div class="card"><h3><i class="fa-solid fa-handshake"></i>โควตาตามผู้รับดำเนินการ<span class="sub">ปีนี้</span></h3>
      <div class="recv">${(d.receivers || []).map(r => `<div class="recv-row"><div class="rh"><b>${esc(shortRecv(r.receiver))}</b><em style="color:${r.quota_level === 'over' ? 'var(--bad)' : r.quota_level === 'warn' ? 'var(--amber2)' : 'var(--brand2)'}">${fmt(r.quota_pct, 1)}%</em><span>${fmtT(r.shipped_year_kg)} / ${fmtT(r.permit_kg)} ตัน · ${r.types} ประเภท</span></div>
        <div class="qbar ${r.quota_level}"><span style="width:${Math.min(100, r.quota_pct || 0)}%"></span></div></div>`).join('')}</div></div>
    <div class="card"><h3><i class="fa-solid fa-cubes-stacked"></i>ภาชนะในพื้นที่จัดเก็บ<span class="sub">${fmt(cTotal)} ภาชนะ</span></h3>
      ${cTotal ? '<div class="chart-box sm"><canvas id="cCont"></canvas></div>' : `<div class="empty">${containerSVG('drum200').replace('<svg', '<svg style="width:70px;height:70px"')}<br>ยังไม่มีภาชนะในพื้นที่จัดเก็บ</div>`}</div>
  </div>

  <div class="grid g-2e mt">
    <div class="card"><h3><i class="fa-solid fa-bell"></i>แจ้งเตือน<span class="sub">โควตา ≥ ${S.limits.quotaWarnPct || 80}% · จัดเก็บเกิน ${S.limits.storageAlertDays || 90} วัน</span></h3>
      <div class="alerts">${(d.alerts || []).length ? d.alerts.map(alertItem).join('') : '<div class="all-good"><i class="fa-solid fa-recycle"></i><div><b>ทุกอย่างปกติ</b><br><small>ไม่มีประเภทใดใกล้เต็มโควตาหรือจัดเก็บนานเกินกำหนด</small></div></div>'}</div></div>
    <div class="card"><h3><i class="fa-solid fa-clock-rotate-left"></i>รายการล่าสุด<span class="sub"><a href="#history" style="color:var(--brand2)">ดูทั้งหมด →</a></span></h3>
      <div class="tl">${(d.recent || []).length ? d.recent.map(tlItem).join('') : `<div class="empty"><i class="fa-solid fa-inbox"></i>ยังไม่มีรายการ — เริ่มที่ “บันทึกจัดเก็บ” หรือบันทึก “ยอดยกมา”</div>`}</div></div>
  </div>`;
  view.onclick = e => { const g = e.target.closest('[data-go]'); if (g) location.hash = g.dataset.go; };
  if ($('#resetBtn')) $('#resetBtn').onclick = resetData;
  // animate rings + needle after the first paint
  requestAnimationFrame(() => requestAnimationFrame(() => {
    $$('.ring-card', view).forEach(el => el.style.setProperty('--p', el.dataset.p));
    const n = $('.needle-gauge .needle', view);
    if (n) n.style.transform = `rotate(${n.dataset.deg}deg)`;
  }));
  drawDashCharts(d);
}
/* ของเสียทั้งหมด = บำบัดเอง + ส่งกำจัด + คงเหลือ (สะสม) — เฉพาะแท็บที่ปลดล็อกแล้ว (ยอดรวมเท่านั้น ไม่มีรายการ TR-) */
function treatSummary(k) {
  const tot = k.total_in_kg || 0, pc = v => tot > 0 ? Math.max(0, v / tot * 100) : 0;
  const parts = [
    ['บำบัดเอง', k.treated_kg, 'fa-flask-vial', 'var(--treat)', 'treat', `ปีนี้ ${fmtT(k.year_treat_kg)} ตัน · เดือนนี้ ${fmtT(k.month_treat_kg)} ตัน`],
    ['ส่งกำจัด (ตัดโควตา)', k.shipped_kg, 'fa-truck-arrow-right', 'var(--out)', 'ship', `ปีนี้ ${fmtT(k.year_out_kg)} / ${fmtT(k.permit_kg)} ตัน · ${fmt(k.quota_pct || 0, 1)}%`],
    ['คงเหลือในโรงงาน', k.balance_kg, 'fa-warehouse', 'var(--brand)', 'bal', 'รอส่งกำจัด'],
  ];
  return `<div class="card mt tsum">
    <h3><i class="fa-solid fa-flask-vial brew"></i>ของเสียทั้งหมดในระบบ (สะสม)<span class="sub">สะสมตั้งแต่เริ่มใช้ระบบ</span></h3>
    <div class="tsum-row">
      <div class="tsum-total"><div class="treat-ic" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-recycle"></i></span></div>
        <div><small>ของเสียทั้งหมด (จัดเก็บ + ยอดยกมา)</small><b><span data-count="${tot / 1000}" data-dec="2">0</span> <em>ตัน</em></b><small>= ${fmtKgHint(tot)} กก.</small></div></div>
      <div class="tsum-parts">${parts.map(([l, v, ic, c, cls, sub]) => `<div class="tpart ${cls}" style="--c:${c}">
        <div class="tp-h"><span class="tp-ic"><i class="fa-solid ${ic}"></i></span><span class="tp-l">${l}</span><i class="pc">${fmt(pc(v), 1)}%</i></div>
 <b><span data-count="${v / 1000}" data-dec="2">0</span> <em>ตัน</em></b><small class="tp-s" title="${esc(sub)}">${sub}</small></div>`).join('')}</div>
    </div>
    <div class="tbar" role="img" aria-label="สัดส่วน บำบัดเอง ${fmt(pc(k.treated_kg), 1)}% ส่งกำจัด ${fmt(pc(k.shipped_kg), 1)}% คงเหลือ ${fmt(pc(k.balance_kg), 1)}%">
      <span class="treat" style="width:${pc(k.treated_kg)}%"></span><span class="ship" style="width:${pc(k.shipped_kg)}%"></span><span class="bal" style="width:${pc(k.balance_kg)}%"></span></div>
    <p class="hint" style="margin-top:8px"><i class="fa-solid fa-circle-info"></i> โควตาตามหนังสืออนุญาต สก.2 นับเฉพาะ “ส่งกำจัด” — บำบัดเองไม่ตัดโควตา</p>
  </div>`;
}
function drawDashCharts(d) {
  if (!window.Chart) return;
  Chart.defaults.font.family = 'Prompt, Kanit, sans-serif';
  Chart.defaults.color = '#5B6F68';
  const lbl = d.months.map(m => new Date(m.ym + '-01T00:00:00').toLocaleDateString('th-TH', { month: 'short', year: '2-digit' }));
  const cm = $('#cMonths');
  if (cm) {
    const g = cm.getContext('2d').createLinearGradient(0, 0, 0, 280);
    g.addColorStop(0, '#34D399'); g.addColorStop(1, '#047857');
    const g2 = cm.getContext('2d').createLinearGradient(0, 0, 0, 280);
    g2.addColorStop(0, '#FDBA74'); g2.addColorStop(1, '#EA580C');
    const g3 = c => { const g = c.getContext('2d').createLinearGradient(0, 0, 0, 280); g.addColorStop(0, '#C4B5FD'); g.addColorStop(1, '#6D28D9'); return g; };
    S.charts.push(new Chart(cm, {
      type: 'bar',
      data: { labels: lbl, datasets: [
        { label: 'จัดเก็บ (ตัน)', data: d.months.map(m => m.in_kg / 1000), backgroundColor: g, borderRadius: 8, maxBarThickness: 26 },
        { label: 'ส่งกำจัด (ตัน)', data: d.months.map(m => m.out_kg / 1000), backgroundColor: g2, borderRadius: 8, maxBarThickness: 26 },
        ...(d.months[0] && d.months[0].treat_kg !== undefined ? [{ label: 'บำบัดเอง (ตัน)', data: d.months.map(m => m.treat_kg / 1000), backgroundColor: g3(cm), borderRadius: 8, maxBarThickness: 26 }] : []),
      ] },
      options: { responsive: true, maintainAspectRatio: false, animation: { duration: reduceMotion() ? 0 : 1200, easing: 'easeOutQuart' },
        plugins: { legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle' } }, tooltip: { callbacks: { label: c => `${c.dataset.label}: ${fmtFix(c.parsed.y, 2)}` } } },
        scales: { x: { grid: { display: false } }, y: { beginAtZero: true, grid: { color: 'rgba(6,78,59,.06)' }, ticks: { callback: v => fmtFix(v, 2) } } } },
    }));
  }
  const cc = $('#cCont');
  if (cc) {
    S.charts.push(new Chart(cc, {
      type: 'doughnut',
      data: { labels: CONT_KEYS.map(k => CONT[k].name), datasets: [{ data: CONT_KEYS.map(k => d.containers[k] || 0), backgroundColor: ['#0D9488', '#F59E0B', '#10B981'], borderWidth: 4, borderColor: '#fff', hoverOffset: 12 }] },
      options: { responsive: true, maintainAspectRatio: false, cutout: '64%', animation: { animateRotate: !reduceMotion(), duration: reduceMotion() ? 0 : 1600 },
        plugins: { legend: { position: 'right', labels: { usePointStyle: true, pointStyle: 'circle', padding: 14 } }, tooltip: { callbacks: { label: c => `${c.label}: ${fmt(c.parsed)} ภาชนะ` } } } },
    }));
  }
}

/* ================= STORE / SHIP FORM ================= */
const meLabel = () => S.user || 'ไม่ระบุ (ตั้งชื่อที่มุมขวาบน)';
function typePicker(sel, mode, q) {
  const qq = q.trim().toLowerCase();
  const list = S.types.filter(t => (mode === 'ship' ? true : +t.active === 1))
    .filter(t => !qq || `${t.code} ${t.name} ${t.receiver} ${t.mgmt_code}`.toLowerCase().includes(qq));
  if (!list.length) return '<div class="empty"><i class="fa-solid fa-magnifying-glass"></i>ไม่พบประเภทของเสีย</div>';
  return receivers().map(r => {
    const ts = list.filter(t => t.receiver === r);
    if (!ts.length) return '';
    return `<div class="rgroup"><h4><i class="fa-solid fa-truck-field"></i>${esc(r)} <small>· ${ts.length} ประเภท</small></h4><div class="tpicks">${ts.map(t => {
      const b = balOf(t.id), off = mode === 'ship' && !(b.balance_kg > 0);
      return `<button type="button" class="tpick ${+t.id === +sel ? 'sel' : ''}" data-type="${t.id}" style="--c:${esc(t.color)}" ${off ? 'disabled' : ''}>
        <span class="ti"><i class="fa-solid ${esc(t.icon)}"></i></span>
        <span class="tt"><b>${esc(t.name)}</b><small><span class="code">${esc(t.code)}</span> · จัดการ ${esc(t.mgmt_code)}</small>
        <span class="bal">${off ? 'ไม่มีคงเหลือ' : 'คงเหลือ ' + fmtT(b.balance_kg) + ' ตัน'}</span></span></button>`;
    }).join('')}</div></div>`;
  }).join('');
}
// quick-add buttons (tons)
const QUICK_T = [0.1, 0.5, 1, 5, 10];
const TREAT_METHODS = ['ระบบบำบัดน้ำเสียภายในโรงงาน', 'เตาเผาภายในโรงงาน', 'ระบบปรับเสถียร / ทำให้เป็นของแข็ง', 'ระบบคัดแยก / รีไซเคิลภายในโรงงาน'];
function txForm(page) {
  const n = nowParts();
  const ship = page === 'ship', treat = page === 'treat', out = ship || treat; // out: takes stock out of the plant
  const F = { kind: ship ? 'ship' : treat ? 'treat' : 'store', type_id: null, weight: '', container: '', containers: 1, date: n.date, time: n.time, location: '', manifest_no: '', vehicle: '', method: treat ? TREAT_METHODS[0] : '', note: '', q: '' };
  if (!S.types.length || (!out && !S.types.some(t => +t.active === 1))) {
    view.innerHTML = noTypesBox(out ? '' : 'การบันทึกจัดเก็บจะเปิดใช้งานเมื่อมีประเภทของเสียที่ใช้งานอย่างน้อย 1 ประเภท');
    return;
  }
  const hashType = +new URLSearchParams(location.hash.split('?')[1] || '').get('type');
  if (hashType && typeById(hashType)) F.type_id = hashType;
  view.innerHTML = `<div class="wizard ${ship ? 'ship-mode' : treat ? 'treat-mode' : ''}">
    <div class="steps">
      ${treat ? `<div class="treat-note"><div class="treat-ic sm" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-flask-vial"></i></span></div>
        <div><b>ข้อมูลภายใน</b><small>ตัดยอดคงเหลือในโรงงานเหมือนส่งกำจัด (ไม่เกินยอดคงเหลือ / จำนวนภาชนะ) แต่ <u>ไม่นับ</u> โควตาตามหนังสืออนุญาต สก.2 (แดชบอร์ดแสดงเป็นยอด “บำบัดเอง”) · บันทึกได้เฉพาะพนักงานกรอกข้อมูล / ผู้ดูแลระบบ</small></div></div>` : ''}
      ${out ? '' : `<div class="seg" id="modeSeg"><button type="button" data-mode="store" class="on"><i class="fa-solid fa-box-archive"></i>บันทึกจัดเก็บ</button><button type="button" data-mode="opening"><i class="fa-solid fa-flag-checkered"></i>ยอดยกมา (Opening balance)</button></div>
        <p class="hint" id="modeHint" style="margin-top:-8px"></p>`}
      <div class="card"><div class="step-h"><span class="n">1</span><h3>เลือกประเภทของเสีย</h3><span class="sub">${S.types.length} ประเภท · แยกตามผู้รับดำเนินการ</span></div>
        <div class="picker-top"><div class="search"><i class="fa-solid fa-magnifying-glass"></i><input id="tQ" placeholder="ค้นหา รหัส / ชื่อ / ผู้รับ เช่น 150102 หรือ หลอดไฟ"></div></div>
        <div id="tPick"></div></div>
      <div class="card"><div class="step-h"><span class="n">2</span><h3>น้ำหนัก &amp; ภาชนะบรรจุ</h3></div>
        <div id="liveBal"></div>
        <div class="field"><label>น้ำหนัก (ตัน) · ทศนิยม 2 ตำแหน่ง เช่น 0.25 ตัน = 250 กก.</label>
          <div class="weight-box" id="wBox"><input id="fW" inputmode="decimal" autocomplete="off" placeholder="0.00"><span>ตัน</span></div>
          <div class="weight-help" id="wHelp">&nbsp;</div>
          <div class="quick" id="wQuick">${QUICK_T.map(v => `<button type="button" data-add="${v}">+${v}</button>`).join('')}<button type="button" data-add="0">ล้าง</button>${out ? '<button type="button" data-all="1">ทั้งหมดที่คงเหลือ</button>' : ''}</div></div>
        <div class="field mt"><label>ภาชนะบรรจุ</label><div class="ccards" id="cCards"></div></div>
        <div class="field mt"><label>จำนวนภาชนะ</label><div class="stepper"><button type="button" data-step="-1" aria-label="ลด"><i class="fa-solid fa-minus"></i></button><input id="fN" inputmode="numeric" value="1"><button type="button" data-step="1" aria-label="เพิ่ม"><i class="fa-solid fa-plus"></i></button></div><div class="weight-help" id="nHelp"></div></div>
      </div>
      <div class="card"><div class="step-h"><span class="n">3</span><h3>รายละเอียด</h3></div>
        <div class="fields">
          <div class="field"><label>วันที่</label><input type="date" id="fDate" value="${F.date}" max="${n.date}"></div>
          <div class="field"><label>เวลา</label><input type="time" id="fTime" value="${F.time}"></div>
          ${ship ? `<div class="field"><label>เลขที่ใบกำกับการขนส่ง (Manifest) — ไม่บังคับ</label><input id="fMan" maxlength="60" placeholder="เช่น MN-2569-000123"></div>
            <div class="field"><label>ทะเบียนรถ — ไม่บังคับ</label><input id="fVeh" maxlength="40" placeholder="เช่น 70-1234 ชลบุรี"></div>`
            : treat ? `<div class="field full"><label>ระบบ / วิธีการ — ไม่บังคับ</label><input id="fMeth" maxlength="120" list="methList" value="${esc(F.method)}" placeholder="เช่น ระบบบำบัดน้ำเสียภายในโรงงาน"><datalist id="methList">${TREAT_METHODS.map(m => `<option value="${esc(m)}">`).join('')}</datalist></div>`
            : `<div class="field full"><label>สถานที่จัดเก็บ — ไม่บังคับ</label><input id="fLoc" maxlength="100" placeholder="เช่น โรงเก็บของเสียอันตราย A"></div>`}
          <div class="field full"><label>หมายเหตุ</label><textarea id="fNote" rows="2" maxlength="255"></textarea></div>
        </div></div>
    </div>
    <aside class="card summary" id="sumBox"></aside>
  </div>`;

  const cont = () => {
    const b = F.type_id ? balOf(F.type_id) : null;
    $('#cCards').innerHTML = CONT_KEYS.map(k => {
      const av = b ? b.containers[k] || 0 : 0;
      return `<div class="ccard ${F.container === k ? 'sel' : ''}" data-c="${k}" tabindex="0" role="radio" aria-checked="${F.container === k}" aria-label="${CONT[k].name}">
        <div class="face front">${containerSVG(k)}<b>${CONT[k].name}</b><small>${CONT[k].sub}</small>${out || (b && av) ? `<span class="avail">ในคลัง ${fmt(av)}</span>` : ''}</div>
        <div class="face back"><span class="chk"><i class="fa-solid fa-check"></i></span>${containerSVG(k, true)}<b>${CONT[k].name}</b>${out || (b && av) ? `<span class="avail">ในคลัง ${fmt(av)}</span>` : '<small>เลือกแล้ว</small>'}</div></div>`;
    }).join('');
  };
  const pick = () => { $('#tPick').innerHTML = typePicker(F.type_id, out ? 'ship' : 'store', F.q); };
  let busy = false; // a save is in flight: the button stays disabled even if the summary is redrawn
  const errs = () => {
    const e = {};
    const w = parseT(F.weight);
    if (!F.type_id) e.type = 'กรุณาเลือกประเภทของเสีย';
    if (F.weight === '') e.w = 'กรุณากรอกน้ำหนัก (ตัน)';
    else if (isNaN(w)) e.w = 'น้ำหนักต้องเป็นตัวเลข (ตัน) เช่น 0.25';
    else if (!(w > 0)) e.w = 'น้ำหนักต้องมากกว่า 0';
    if (!F.container) e.c = 'กรุณาเลือกภาชนะ';
    if (!/^\d+$/.test(String(F.containers)) || +F.containers < 1) e.n = 'จำนวนภาชนะต้องเป็นจำนวนเต็มตั้งแต่ 1';
    if (out && F.type_id) { // stock-out forms only (ship / treat): not more than what is in the plant
      const b = balOf(F.type_id);
      if (!e.w && w > b.balance_kg + 1e-9) e.w = `เกินยอดคงเหลือ (${fmtT(b.balance_kg)} ตัน)`;
      if (F.container && !e.n && +F.containers > (b.containers[F.container] || 0)) e.n = `${CONT[F.container].name}ในคลังมี ${fmt(b.containers[F.container] || 0)} ภาชนะ`;
    }
    return e;
  };
  const live = () => {
    const t = typeById(F.type_id), box = $('#liveBal');
    if (!t) { box.innerHTML = ''; return; }
    const b = balOf(t.id), w = parseT(F.weight) || 0, after = Math.round((out ? b.balance_kg - w : b.balance_kg + w) * 100) / 100;
    box.innerHTML = `<div class="live-bal ${out && after < -1e-9 ? 'low' : ''}"><span class="lb-ic"><i class="fa-solid fa-scale-balanced"></i></span>
      <div><small>คงเหลือในโรงงานตอนนี้ · ${esc(t.name)}</small><b>${fmtT(b.balance_kg)}</b> ตัน <small>หลังบันทึก ${fmtT(after)} ตัน · ${CONT_KEYS.map(k => `${CONT[k].name} ${fmt(b.containers[k] || 0)}`).join(' · ')}</small></div></div>`;
  };
  const summary = () => {
    const t = typeById(F.type_id), e = errs(), w = parseT(F.weight), k = KIND[F.kind];
    const ok = !Object.keys(e).length;
    $('#sumBox').innerHTML = `${treat ? '<div class="treat-ic sum-ic" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-flask-vial"></i></span></div>' : ecoLogo('sum-ic')}
      <h3 style="justify-content:center"><i class="fa-solid ${k.icon}" style="color:${k.c}"></i>${k.name}</h3>
      <div class="sum-total"><b>${isNaN(w) ? '0.00' : fmtT(w)} <small style="font-size:16px">ตัน</small></b><small>${isNaN(w) ? '' : '= ' + fmtKgHint(w) + ' กก.'}</small></div>
      <div class="sum-list">
        <div><span>ของเสีย</span><b>${t ? esc(t.name) : '—'}</b></div>
        <div><span>รหัสของเสีย</span><b>${t ? esc(t.code) + ' · จัดการ ' + esc(t.mgmt_code) : '—'}</b></div>
        ${treat ? `<div><span>ระบบ / วิธีการ</span><b>${esc(F.method || '—')}</b></div>
        <div><span>โควตา สก.2</span><b style="color:var(--treat)">ไม่นับ (ข้อมูลภายใน)</b></div>`
          : `<div><span>ผู้รับดำเนินการ</span><b>${t ? esc(t.receiver) : '—'}</b></div>`}
        <div><span>ภาชนะ</span><b>${F.container ? CONT[F.container].name + ' × ' + fmt(F.containers) : '—'}</b></div>
        <div><span>วันที่</span><b>${F.date ? dateTH(F.date) + ' ' + esc(F.time) + ' น.' : '—'}</b></div>
        <div><span>ผู้บันทึก</span><b>${esc(meLabel())}</b></div>
      </div>
      <button class="btn btn-lg ${ship ? 'btn-ship' : treat ? 'btn-treat' : 'btn-primary'}" id="fGo" ${ok ? '' : 'disabled'}><i class="fa-solid ${ship ? 'fa-truck-arrow-right' : treat ? 'fa-flask-vial' : 'fa-floppy-disk'}"></i>${ship ? 'ยืนยันส่งกำจัด' : treat ? 'บันทึกข้อมูลภายใน' : F.kind === 'opening' ? 'บันทึกยอดยกมา' : 'บันทึกจัดเก็บ'}</button>
      <p class="hint" style="margin-top:10px;text-align:center">${ok ? '<i class="fa-solid fa-circle-check" style="color:var(--ok)"></i> พร้อมบันทึก' : esc(Object.values(e)[0])}</p>`;
    $('#fGo').onclick = submit;
    if (busy) { $('#fGo').disabled = true; $('#fGo').innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังบันทึก...'; }
    const we = e.w && F.weight !== '';
    $('#wBox').classList.toggle('bad', !!we);
    $('#wHelp').textContent = we ? e.w : (out && t ? `${treat ? 'ตัดยอด' : 'ส่ง'}ได้ไม่เกิน ${fmtT(balOf(t.id).balance_kg)} ตัน` : '\u00a0');
    $('#wHelp').classList.toggle('bad', !!we);
    $('#nHelp').textContent = e.n || '';
    $('#nHelp').className = 'weight-help' + (e.n ? ' bad' : '');
  };
  const refresh = () => { live(); summary(); };
  const setMode = m => {
    F.kind = m;
    $$('#modeSeg button').forEach(b => b.classList.toggle('on', b.dataset.mode === m));
    $('#modeHint').innerHTML = m === 'opening' ? '<i class="fa-solid fa-circle-info"></i> ยอดยกมา: ของเสียที่จัดเก็บอยู่ในโรงงานก่อนเริ่มใช้ระบบ — บันทึกตามที่ตรวจนับจริง (นับรวมในยอดคงเหลือ ไม่นับเป็นการจัดเก็บประจำเดือน)' : '<i class="fa-solid fa-circle-info"></i> บันทึกของเสียที่นำเข้าพื้นที่จัดเก็บของโรงงาน';
    refresh();
  };
  async function submit() {
    const go = $('#fGo');
    if (!go || go.disabled || busy) return;
    const e = errs();
    if (Object.keys(e).length) { toast(Object.values(e)[0], 'err'); return; }
    busy = true; go.disabled = true; go.innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังบันทึก...';
    // tons go to the server as typed (the server converts them to kg exactly)
    const body = { type_id: F.type_id, weight_t: tonsStr(parseT(F.weight)), container: F.container, containers: String(F.containers), date: F.date, time: F.time, note: F.note };
    if (ship) Object.assign(body, { manifest_no: F.manifest_no, vehicle: F.vehicle });
    else if (treat) body.method = F.method;
    else body.location = F.location;
    try {
      const r = await writeApi(F.kind, body);
      const t = typeById(F.type_id);
      confetti();
      const m = modal(`<div class="success-ic ${ship ? 'ship' : treat ? 'treat' : ''}"><span class="sr"></span><span class="sc"><i class="fa-solid ${ship ? 'fa-truck-arrow-right' : treat ? 'fa-arrows-spin' : 'fa-check'}"></i></span></div>
        <h3 style="justify-content:center">${KIND[F.kind].name}สำเร็จ</h3>
        <div class="doc">${esc(r.doc_no)}</div>
        <div class="receipt"><div class="sum-list">
          <div><span>ของเสีย</span><b>${esc(typeLabel(t))}</b></div>
          <div><span>น้ำหนัก</span><b>${fmtT(parseT(body.weight_t))} ตัน</b></div>
          <div><span>ภาชนะ</span><b>${CONT[F.container].name} × ${fmt(F.containers)}</b></div>
          ${treat && F.method ? `<div><span>ระบบ / วิธีการ</span><b>${esc(F.method)}</b></div>` : ''}
          <div><span>คงเหลือในโรงงาน</span><b>${fmtT(r.balance_kg)} ตัน</b></div>
          <div><span>ผู้บันทึก</span><b>${esc(S.user)}</b></div></div></div>
        <div class="modal-acts">${treat ? '<a class="btn btn-ghost" href="#treat" data-close><i class="fa-solid fa-list"></i>ดูรายการ</a>' : '<a class="btn btn-ghost" href="#balance" data-close><i class="fa-solid fa-scale-balanced"></i>ดูคงเหลือ</a>'}<button class="btn btn-primary" data-close><i class="fa-solid fa-plus"></i>ทำรายการต่อ</button></div>`);
      S.onModalClose = () => route();
      m.querySelector('[data-close].btn-primary')?.focus();
    } catch (err) {
      busy = false;
      toast(err.message, 'err');
      if (err.code === 'AUTH_REQUIRED') return;
      summary();
    }
  }

  // events
  view.onclick = e => {
    const tp = e.target.closest('[data-type]');
    if (tp && !tp.disabled) { F.type_id = +tp.dataset.type; $$('.tpick', view).forEach(b => b.classList.toggle('sel', b === tp)); cont(); refresh(); return; }
    const cc = e.target.closest('[data-c]');
    if (cc) { F.container = cc.dataset.c; $$('.ccard', view).forEach(c => { const on = c === cc; c.classList.toggle('sel', on); c.setAttribute('aria-checked', on); }); refresh(); return; }
    const md = e.target.closest('[data-mode]');
    if (md) { setMode(md.dataset.mode); return; }
    const ad = e.target.closest('[data-add]');
    if (ad) {
      // in whole hundredths of a kg: +0.1 t steps never drift
      const v = +ad.dataset.add, cur = parseT(F.weight) || 0;
      F.weight = v ? tonsStr(Math.round(cur * 100 + v * 100000) / 100) : '';
      $('#fW').value = F.weight; refresh(); return;
    }
    if (e.target.closest('[data-all]') && F.type_id) { F.weight = tonsStr(balOf(F.type_id).balance_kg); $('#fW').value = F.weight; refresh(); return; }
    const st = e.target.closest('[data-step]');
    if (st) { F.containers = Math.max(1, (parseInt(F.containers, 10) || 0) + +st.dataset.step); $('#fN').value = F.containers; refresh(); }
  };
  view.onkeydown = e => { const cc = e.target.closest('.ccard'); if (cc && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); cc.click(); } };
  $('#tQ').oninput = e => { F.q = e.target.value; pick(); };
  $('#fW').oninput = e => { F.weight = e.target.value.trim(); refresh(); };
  $('#fN').oninput = e => { F.containers = e.target.value.trim(); refresh(); };
  $('#fDate').onchange = e => { F.date = e.target.value; summary(); };
  $('#fTime').onchange = e => { F.time = e.target.value; summary(); };
  $('#fNote').oninput = e => { F.note = e.target.value; };
  if (ship) { $('#fMan').oninput = e => { F.manifest_no = e.target.value; }; $('#fVeh').oninput = e => { F.vehicle = e.target.value; }; }
  else if (treat) $('#fMeth').oninput = e => { F.method = e.target.value; summary(); };
  else $('#fLoc').oninput = e => { F.location = e.target.value; };
  // fresh data arrived while the user is typing: update balances only
  S.onFresh = () => { pick(); cont(); refresh(); };
  pick(); cont();
  if (out) refresh(); else setMode('store');
}

/* ================= BALANCE ================= */
function balance() {
  if (!S.types.length) { view.innerHTML = noTypesBox(); return; }
  const days = S.limits.storageAlertDays || 90;
  // per-type in-house totals (บำบัดเอง): only when the server sent them (unlocked tab)
  const showT = S.types.some(t => balOf(t.id).treated_kg !== undefined);
  let filter = 'all';
  const recv = receivers().map(r => {
    const ts = S.types.filter(t => t.receiver === r);
    const sum = k => ts.reduce((a, t) => a + Math.round((balOf(t.id)[k] || 0) * 100), 0) / 100;
    const permit = sum('permit_kg'), used = sum('shipped_year_kg');
    const pct = permit > 0 ? Math.round(used / permit * 1000) / 10 : null;
    return { r, ts, permit, used, bal: sum('balance_kg'), pct, lv: pct == null ? 'ok' : pct > 100 ? 'over' : pct >= (S.limits.quotaWarnPct || 80) ? 'warn' : 'ok' };
  });
  const totalBal = S.types.reduce((a, t) => a + Math.round(balOf(t.id).balance_kg * 100), 0) / 100;
  const cardOf = t => {
    const b = balOf(t.id), pct = b.quota_pct;
    return `<div class="bcard" style="--c:${esc(t.color)}">
      <div class="bh"><span class="bi"><i class="fa-solid ${esc(t.icon)}"></i></span><div class="bt"><b>${esc(t.name)} ${+t.active === 1 ? '' : '<span class="tag off">ปิดใช้งาน</span>'}</b><small>${esc(t.code)} · จัดการ ${esc(t.mgmt_code)} · ${esc(shortRecv(t.receiver))}</small></div></div>
      <div class="big"><b data-count="${b.balance_kg / 1000}" data-dec="2">0</b><span>ตัน คงเหลือ</span></div>
      <div class="flow"><span><i class="fa-solid fa-arrow-down" style="color:var(--in)"></i> รับเข้า <b>${fmtT(b.in_kg)}</b></span><span><i class="fa-solid fa-arrow-up" style="color:var(--out)"></i> ส่งกำจัด <b>${fmtT(b.shipped_kg)}</b></span>${showT && b.treated_kg > 0 ? `<span class="ft"><i class="fa-solid fa-flask-vial"></i> บำบัดเอง <b>${fmtT(b.treated_kg)}</b></span>` : ''}${b.opening_kg ? `<span>(ยกมา ${fmtT(b.opening_kg)})</span>` : ''}</div>
      <div class="cchips">${CONT_KEYS.map(k => `<span class="cchip ${b.containers[k] ? '' : 'zero'}">${containerSVG(k)}${CONT[k].name} <b>${fmt(b.containers[k])}</b></span>`).join('')}</div>
      <div class="qline"><span>ส่งกำจัดปีนี้ <b>${fmtT(b.shipped_year_kg)}</b> / ${fmtT(b.permit_kg)} ตัน</span><span class="tag lv-${b.quota_level}">${pct == null ? '—' : fmt(pct, 1) + '%'}</span></div>
      <div class="qbar ${b.quota_level}"><span style="width:${Math.min(100, pct || 0)}%"></span></div>
      ${b.age_days != null ? `<span class="agepill ${b.age_days > days ? 'old' : ''}"><i class="fa-solid fa-hourglass-half"></i>ล็อตเก่าสุด ${fmt(b.age_days)} วัน (ตั้งแต่ ${dateTH(b.oldest_at)})</span>` : ''}
      <div class="rec-acts" style="margin-top:10px;justify-content:flex-end"><a class="btn btn-sm btn-ghost" href="#store?type=${t.id}"><i class="fa-solid fa-box-archive"></i>จัดเก็บ</a>${b.balance_kg > 0 ? `<a class="btn btn-sm btn-ship" href="#ship?type=${t.id}"><i class="fa-solid fa-truck-arrow-right"></i>ส่งกำจัด</a>` : ''}${can('staff') && b.balance_kg > 0 ? `<a class="btn btn-sm btn-treat" href="#treat?tab=new&type=${t.id}"><i class="fa-solid fa-flask-vial"></i>ข้อมูลภายใน</a>` : ''}</div>
    </div>`;
  };
  const show = () => {
    const list = S.types.filter(t => {
      const b = balOf(t.id);
      if (filter === 'has') return b.balance_kg > 0;
      if (filter === 'alert') return b.quota_level !== 'ok' || (b.age_days != null && b.age_days > days);
      return +t.active === 1 || b.balance_kg > 0 || b.in_kg > 0;
    });
    $('#balGrid').innerHTML = list.length ? list.map(cardOf).join('') : '<div class="empty"><i class="fa-solid fa-leaf"></i>ไม่มีรายการตามตัวกรอง</div>';
    countUp($('#balGrid'));
  };
  view.innerHTML = `
    <div class="recv-cards stagger">${recv.map(x => `<div class="rcard"><h4><i class="fa-solid fa-truck-field"></i>${esc(x.r)}</h4>
      <div style="display:flex;align-items:baseline;gap:8px"><span class="pct" style="color:${x.lv === 'over' ? 'var(--bad)' : x.lv === 'warn' ? 'var(--amber2)' : 'var(--brand2)'}">${x.pct == null ? '—' : fmt(x.pct, 1) + '%'}</span><span class="hint">ใช้โควตาปีนี้</span></div>
      <div class="qbar ${x.lv}"><span style="width:${Math.min(100, x.pct || 0)}%"></span></div>
      <div class="nums"><span>ส่งแล้ว ${fmtT(x.used)} / ${fmtT(x.permit)} ตัน</span><span>คงเหลือในโรงงาน ${fmtT(x.bal)} ตัน</span></div></div>`).join('')}</div>
    <div class="card mt">
      <h3><i class="fa-solid fa-scale-balanced"></i>คงเหลือในโรงงานรวม <span style="font-family:Kanit;color:var(--brand2)">${fmtT(totalBal)} ตัน</span><span class="sub">= ${fmtKgHint(totalBal)} กก.</span></h3>
      <div class="seg" id="balSeg"><button data-f="all" class="on">ทั้งหมด</button><button data-f="has">มีคงเหลือ</button><button data-f="alert">ต้องติดตาม (${alertCount().n})</button></div>
    </div>
    <div class="bal-grid mt" id="balGrid"></div>
    <div class="card mt"><h3><i class="fa-solid fa-table"></i>ตารางสรุปตามหนังสืออนุญาต<span class="sub">${S.company.permitYear ? 'ปี ' + esc(S.company.permitYear) + ' · ' : ''}หน่วย ตัน</span></h3>
      <div class="tbl-wrap"><table><thead><tr><th>#</th><th>รหัส</th><th>ของเสีย</th><th>จัดการ</th><th>ผู้รับดำเนินการ</th><th class="num">อนุญาต (ตัน)</th><th class="num">รับเข้าสะสม</th><th class="num">ส่งกำจัดสะสม</th>${showT ? '<th class="num th-treat">บำบัดเองสะสม</th>' : ''}<th class="num">คงเหลือ</th><th class="num">ส่งปีนี้</th><th class="num">% โควตา</th></tr></thead><tbody>
      ${S.types.map((t, i) => { const b = balOf(t.id); return `<tr><td>${i + 1}</td><td><span class="tcode" style="--c:${esc(t.color)}"><i class="fa-solid ${esc(t.icon)}"></i>${esc(t.code)}</span></td><td>${esc(t.name)}</td><td>${esc(t.mgmt_code)}</td><td>${esc(shortRecv(t.receiver))}</td>
        <td class="num">${fmtFix(t.permit_tons, 2)}</td><td class="num">${fmtT(b.in_kg)}</td><td class="num">${fmtT(b.shipped_kg)}</td>${showT ? `<td class="num" style="color:var(--treat)">${fmtT(b.treated_kg)}</td>` : ''}<td class="num"><b>${fmtT(b.balance_kg)}</b></td><td class="num">${fmtT(b.shipped_year_kg)}</td><td class="num"><span class="tag lv-${b.quota_level}">${b.quota_pct == null ? '—' : fmt(b.quota_pct, 1) + '%'}</span></td></tr>`; }).join('')}
      </tbody></table></div></div>`;
  view.onclick = e => { const f = e.target.closest('[data-f]'); if (f) { filter = f.dataset.f; $$('#balSeg button').forEach(b => b.classList.toggle('on', b === f)); show(); } };
  show();
}

/* ================= HISTORY ================= */
const typeOptions = (sel, filter = () => true) => receivers().map(r => {
  const ts = S.types.filter(t => t.receiver === r && (filter(t) || +t.id === +sel));
  return ts.length ? `<optgroup label="${esc(r)}">${ts.map(t => `<option value="${t.id}" ${+t.id === +sel ? 'selected' : ''}>${esc(t.code)} · ${esc(t.name)}</option>`).join('')}</optgroup>` : '';
}).join('');
// not named history(): a global function declaration would replace window.history
// fixedKind 'treat': the list on the unlocked ข้อมูลภายใน page (rows of that kind only)
function historyPage(initial, fixedKind = '') {
  const n = nowParts();
  const kk = fixedKind ? kinds().filter(([k]) => k === fixedKind) : kinds();
  view.innerHTML = `<div class="card">
    <div class="filters">
      <div class="field"><label>ตั้งแต่วันที่</label><input type="date" id="hFrom" value="${fixedKind ? n.date.slice(0, 5) + '01-01' : monthFirst()}"></div>
      <div class="field"><label>ถึงวันที่</label><input type="date" id="hTo" value="${n.date}"></div>
      <div class="field" ${fixedKind ? 'hidden' : ''}><label>รายการ</label><select id="hKind">${fixedKind ? `<option value="${fixedKind}" selected>${KIND[fixedKind].name}</option>` : `<option value="">ทั้งหมด</option>${kk.map(([k, a]) => `<option value="${k}">${a.name}</option>`).join('')}`}</select></div>
      <div class="field"><label>ประเภทของเสีย</label><select id="hType"><option value="">ทั้งหมด</option>${typeOptions('')}</select></div>
      <div class="field" ${fixedKind ? 'hidden' : ''}><label>ผู้รับดำเนินการ</label><select id="hRecv"><option value="">ทั้งหมด</option>${receivers().map(r => `<option>${esc(r)}</option>`).join('')}</select></div>
      <div class="field"><label>ค้นหา</label><input id="hQ" placeholder="เลขที่ / ผู้บันทึก / ใบกำกับ / หมายเหตุ"></div>
      <button class="btn btn-primary" id="hGo"><i class="fa-solid fa-magnifying-glass"></i>ค้นหา</button>
      <button class="btn btn-amber" id="hCsv"><i class="fa-solid fa-file-excel"></i>Excel</button>
    </div>
    <div id="hRes">${loader()}</div></div>`;
  const params = () => '&' + new URLSearchParams({ from: $('#hFrom').value, to: $('#hTo').value, type: $('#hType').value, kind: $('#hKind').value, receiver: $('#hRecv').value, q: $('#hQ').value });
  let items = [], loadSeq = 0;
  // edit / delete: everyone sees the buttons; the admin password is asked on click (the server checks it)
  const acts = m => `<div class="rec-acts"><button class="btn btn-sm btn-ghost rec-btn" data-medit="${m.id}" title="แก้ไขรายการ">✏️ แก้ไข</button><button class="btn btn-sm btn-ghost rec-btn" data-mdel="${m.id}" title="ลบรายการ">🗑 ลบ</button></div>`;
  const show = () => {
    const tot = k => items.filter(i => i.kind === k).reduce((a, i) => a + Math.round(i.weight_kg * 100), 0) / 100;
    $('#hRes').innerHTML = items.length ? `
      <div class="chips" style="margin-bottom:12px">${kk.map(([k, a]) => `<span class="tag k-${k}"><i class="fa-solid ${a.icon}"></i>${a.short} ${fmtT(tot(k))} ตัน</span>`).join('')}<span class="tag off">${fmt(items.length)} รายการ</span></div>
      <div class="tbl-wrap tbl-desktop" style="max-height:64vh"><table><thead><tr><th>เลขที่เอกสาร</th><th>วันที่ / เวลา</th><th>รายการ</th><th>ของเสีย</th><th>ผู้รับดำเนินการ</th><th class="num">น้ำหนัก (ตัน)</th><th>ภาชนะ</th><th>ที่เก็บ / ใบกำกับ</th><th>หมายเหตุ</th><th>ผู้บันทึก</th><th>จัดการ</th></tr></thead><tbody>
      ${items.map(m => { const k = KIND[m.kind] || KIND.store; return `<tr><td><b>${esc(m.doc_no)}</b></td><td style="white-space:nowrap">${dtTH(m.moved_at)}</td>
        <td><span class="tag k-${m.kind}"><i class="fa-solid ${k.icon}"></i>${k.short}</span></td>
        <td><span class="tcode" style="--c:${esc(m.color)}"><i class="fa-solid ${esc(m.icon)}"></i><span><b>${esc(m.code)}</b> ${esc(m.type_name)}</span></span></td>
        <td style="font-size:12.5px">${m.kind === 'treat' ? '<span style="color:var(--treat)">ภายในโรงงาน</span>' : esc(shortRecv(m.receiver))}</td>
        <td class="num"><b class="kg ${k.cls}">${k.sign}${fmtT(m.weight_kg)}</b></td>
        <td style="white-space:nowrap">${CONT[m.container]?.name || esc(m.container)} × ${fmt(m.containers)}</td>
        <td style="font-size:12.5px">${esc([m.location, m.manifest_no && 'ใบกำกับ ' + m.manifest_no, m.vehicle && 'รถ ' + m.vehicle, m.method && 'ระบบ ' + m.method].filter(Boolean).join(' · ') || '—')}</td>
        <td style="max-width:220px;overflow:hidden;text-overflow:ellipsis">${esc(m.note || '')}</td>
        <td class="by-cell">${m.actor ? `<i class="fa-solid fa-user-check" style="color:var(--ok)"></i> ${esc(m.actor)}` : '—'}</td><td>${acts(m)}</td></tr>`; }).join('')}
      </tbody></table></div>
      <div class="hcards">${items.map(m => { const k = KIND[m.kind] || KIND.store; return `<div class="hcard"><div class="r1"><span class="tag k-${m.kind}"><i class="fa-solid ${k.icon}"></i>${k.short}</span><b>${esc(m.doc_no)}</b><span class="kg ${k.cls}">${k.sign}${fmtT(m.weight_kg)} ตัน</span></div>
        <p><span class="tcode" style="--c:${esc(m.color)}"><i class="fa-solid ${esc(m.icon)}"></i>${esc(m.code)}</span> ${esc(m.type_name)}</p>
        <small>${dtTH(m.moved_at)} · ${CONT[m.container]?.name || ''} × ${fmt(m.containers)} · ${esc(m.actor || '')}${m.manifest_no ? ' · ใบกำกับ ' + esc(m.manifest_no) : ''}${m.method ? ' · ' + esc(m.method) : ''}${m.note ? ' · ' + esc(m.note) : ''}</small>${acts(m)}</div>`; }).join('')}</div>`
      : '<div class="empty"><i class="fa-solid fa-magnifying-glass"></i>ไม่พบรายการตามเงื่อนไข</div>';
  };
  const load = async () => {
    const seq = ++loadSeq, box = $('#hRes');
    box.innerHTML = loader('กำลังค้นหา...');
    try {
      const r = await api('history', null, params());
      if (seq !== loadSeq || !box.isConnected) return;
      items = r.items; show();
    } catch (e) {
      if (seq !== loadSeq || !box.isConnected) return;
      box.innerHTML = `<div class="empty"><i class="fa-solid fa-plug-circle-xmark"></i>${esc(e.message)}</div>`;
    }
  };
  view.onclick = e => {
    const ed = e.target.closest('[data-medit]'), dl = e.target.closest('[data-mdel]');
    const m = ed || dl ? items.find(i => +i.id === +(ed || dl).dataset[ed ? 'medit' : 'mdel']) : null;
    if (!m) return;
    if (ed) editMovement(m); else deleteMovement(m);
  };
  $('#hGo').onclick = load;
  $('#hQ').onkeydown = e => { if (e.key === 'Enter') load(); };
  $('#hCsv').onclick = () => downloadCsv('export', params());
  if (initial) { items = initial.items; show(); } else load();
}

/* ================= ข้อมูลภายใน (ปลดล็อกแล้ว) =================
   รายการ TR- (ทุกระดับ) · ฟอร์มบันทึก (พนักงานกรอกข้อมูล / ผู้ดูแลระบบ) · ปุ่มล็อก */
function treatBar(tab) {
  const r = S.sess.role;
  return `<div class="card treat-bar">
    <div class="treat-ic sm" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-lock-open"></i></span></div>
    <div class="tb-tx"><b>ข้อมูลภายใน · ปลดล็อกแล้ว</b><small><span class="role-tag r-${r}"><i class="fa-solid ${ROLE_IC[r]}"></i>${ROLE_TH[r]}</span> ${ROLE_CAN[r]}</small></div>
    <div class="seg tb-seg" role="tablist">
      <a href="#treat" role="tab" aria-selected="${tab === 'list'}" class="${tab === 'list' ? 'on' : ''}"><i class="fa-solid fa-list"></i>รายการ</a>
      <a href="#treat?tab=new" role="tab" aria-selected="${tab === 'new'}" class="${tab === 'new' ? 'on' : ''}" title="${can('staff') ? '' : 'ต้องใช้รหัสพนักงานกรอกข้อมูล / ผู้ดูแลระบบ'}"><i class="fa-solid ${can('staff') ? 'fa-plus' : 'fa-lock'}"></i>บันทึกข้อมูลภายใน</a>
    </div>
    <button class="btn btn-ghost btn-sm tb-lock" type="button" id="tbLock" title="ล็อกข้อมูลภายใน (ต้องใส่รหัสใหม่)"><i class="fa-solid fa-lock"></i>ล็อก</button>
  </div>`;
}
function treatPage(h) {
  const tab = treatTab();
  if (tab === 'new') txForm('treat'); else historyPage(h, 'treat');
  view.insertAdjacentHTML('afterbegin', treatBar(tab));
  $('#tbLock').onclick = lockNow;
}

/* ================= EDIT / DELETE (ผู้ดูแลระบบ) =================
   เซิร์ฟเวอร์ตรวจสิทธิ์ (admin) และตรวจย้อนทุกรายการตามลำดับเวลา — ถ้ายอดใดจะติดลบจะปฏิเสธและแจ้งเหตุผล */
function recordForm(html, onSave) {
  const m = modal(html + `<div id="eErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px;white-space:pre-line"></div>
    <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn btn-primary" id="eSave"><i class="fa-solid fa-floppy-disk"></i>บันทึก</button></div>`, true);
  const go = $('#eSave', m), err = $('#eErr', m);
  go.onclick = async () => {
    if (go.disabled) return;
    go.disabled = true; go.innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังบันทึก...'; err.textContent = '';
    try {
      const r = await onSave(m);
      closeModal();
      toast(r.message, r.changed === false ? 'info' : 'ok');
      if (r.changed !== false) route();
    } catch (e) {
      err.textContent = e.message; toast(e.message, 'err');
      go.disabled = false; go.innerHTML = '<i class="fa-solid fa-floppy-disk"></i>บันทึก';
    }
  };
  return m;
}
async function editMovement(m) {
  if (!(await needRole('admin', 'การแก้ไขรายการต้องใช้รหัสผ่านผู้ดูแลระบบ'))) return;
  const k = KIND[m.kind] || KIND.store, ship = m.kind === 'ship', treat = m.kind === 'treat';
  const w0 = tonsStr(m.weight_kg); // exact (older kg rows may need up to 5 decimals)
  recordForm(`<h3><i class="fa-solid fa-pen-to-square" style="color:var(--brand)"></i> แก้ไขรายการ ${esc(m.doc_no)}</h3>
    <p><span class="tag k-${m.kind}"><i class="fa-solid ${k.icon}"></i>${k.name}</span></p>
    <p class="hint" style="margin-top:6px"><i class="fa-solid fa-circle-info"></i> เปลี่ยนชนิดรายการไม่ได้ — หากต้องการให้ลบแล้วบันทึกใหม่ · ระบบตรวจยอดคงเหลือย้อนหลังทุกรายการ ไม่ให้ติดลบ</p>
    <p class="hint"><i class="fa-solid fa-user-check" style="color:var(--ok)"></i> บันทึกโดย: <b>${esc(m.actor || 'ไม่ระบุ')}</b> · การแก้ไขนี้จะบันทึกใน Log ในนาม <b>${esc(meLabel())}</b></p>
    <div class="fields mt">
      <div class="field full"><label>ประเภทของเสีย</label><select id="eType">${typeOptions(m.type_id, t => ship || treat || +t.active === 1)}</select></div>
      <div class="field"><label>น้ำหนัก (ตัน)</label><input id="eW" inputmode="decimal" value="${w0}"></div>
      <div class="field"><label>ภาชนะ</label><select id="eC">${CONT_KEYS.map(c => `<option value="${c}" ${c === m.container ? 'selected' : ''}>${CONT[c].name}</option>`).join('')}</select></div>
      <div class="field"><label>จำนวนภาชนะ</label><input id="eN" inputmode="numeric" value="${+m.containers}"></div>
      <div class="field"><label>วันที่</label><input type="date" id="eDate" value="${m.moved_at.slice(0, 10)}"></div>
      <div class="field"><label>เวลา</label><input type="time" id="eTime" value="${m.moved_at.slice(11, 16)}"></div>
      ${ship ? `<div class="field"><label>เลขที่ใบกำกับการขนส่ง</label><input id="eMan" maxlength="60" value="${esc(m.manifest_no)}"></div><div class="field full"><label>ทะเบียนรถ</label><input id="eVeh" maxlength="40" value="${esc(m.vehicle)}"></div>`
        : treat ? `<div class="field full"><label>ระบบ / วิธีการ</label><input id="eMeth" maxlength="120" list="methList" value="${esc(m.method)}"><datalist id="methList">${TREAT_METHODS.map(x => `<option value="${esc(x)}">`).join('')}</datalist></div>`
        : `<div class="field full"><label>สถานที่จัดเก็บ</label><input id="eLoc" maxlength="100" value="${esc(m.location)}"></div>`}
      <div class="field full"><label>หมายเหตุ</label><input id="eNote" maxlength="255" value="${esc(m.note)}"></div>
    </div>`, f => {
    const body = { id: m.id, doc_no: m.doc_no, type_id: $('#eType', f).value, container: $('#eC', f).value, containers: $('#eN', f).value, date: $('#eDate', f).value, time: $('#eTime', f).value, note: $('#eNote', f).value };
    if (ship) Object.assign(body, { manifest_no: $('#eMan', f).value, vehicle: $('#eVeh', f).value });
    else if (treat) body.method = $('#eMeth', f).value;
    else body.location = $('#eLoc', f).value;
    // the weight only when it was changed: an untouched field keeps the stored kg exactly
    const wt = $('#eW', f).value.trim();
    if (wt !== w0) {
      if (isNaN(parseT(wt))) return Promise.reject(new Error('น้ำหนักต้องเป็นตัวเลข (ตัน) เช่น 0.25'));
      body.weight_t = wt;
    }
    return writeApi('movement_update', body);
  });
}
async function deleteMovement(m) {
  const k = KIND[m.kind]?.name || m.kind;
  const msg = [`ลบรายการ ${m.doc_no}?`, '', `${k} ${m.code} ${m.type_name} ${fmtT(m.weight_kg)} ตัน · ${CONT[m.container]?.name || ''} × ${m.containers} · ${dtTH(m.moved_at)}`, '',
    'ยอดคงเหลือจะคำนวณใหม่ (ระบบไม่ยอมให้ยอดย้อนหลังติดลบ) · การลบย้อนกลับไม่ได้ ต้องการลบหรือไม่?'];
  try {
    if (!(await needRole('admin', 'การลบรายการต้องใช้รหัสผ่านผู้ดูแลระบบ'))) return;
    if (!confirm(msg.join('\n'))) return;
    const r = await writeApi('movement_delete', { id: m.id, doc_no: m.doc_no });
    toast(r.message, 'info');
    route();
  } catch (e) { toast(e.message, 'err'); }
}

/* ================= LOGS ================= */
const LOGCAT = { waste: ['ของเสีย', 'fa-recycle', '#059669'], edit: ['แก้ไข / ลบ', 'fa-pen-to-square', '#8B5CF6'], setting: ['ตั้งค่า', 'fa-gear', '#0F2A23'], warn: ['ถูกปฏิเสธ', 'fa-ban', '#F59E0B'] };
const LOGACT = { store: 'fa-box-archive', ship: 'fa-truck-arrow-right', opening: 'fa-flag-checkered', treat: 'fa-flask-vial', treat_unlock: 'fa-lock-open' /* older rows */, movement_update: 'fa-pen-to-square', movement_delete: 'fa-trash-can', type_save: 'fa-gear', receiver_save: 'fa-truck-field', receiver_delete: 'fa-trash-can', type_delete: 'fa-trash-can', reset_data: 'fa-broom' };
async function logs() {
  const n = nowParts();
  let cat = '', items = null, loadedWith = '';
  view.innerHTML = `
    <div class="log-kpi" id="logKpi"></div>
    <div class="card">
      <div class="filters">
        <div class="field"><label>ตั้งแต่วันที่</label><input type="date" id="lFrom" value="${monthFirst()}"></div>
        <div class="field"><label>ถึงวันที่</label><input type="date" id="lTo" value="${n.date}"></div>
        <div class="field" style="flex:2;min-width:200px"><label>ค้นหา</label><input id="lQ" placeholder="เลขที่เอกสาร / ผู้ทำรายการ / รายละเอียด"></div>
        <button class="btn btn-primary" id="lGo"><i class="fa-solid fa-magnifying-glass"></i>ค้นหา</button>
        <button class="btn btn-amber" id="lCsv"><i class="fa-solid fa-file-excel"></i>Excel</button>
      </div>
      <div id="lRes">${loader()}</div>
    </div>`;
  const params = (withCat = true) => '&' + new URLSearchParams({ from: $('#lFrom').value, to: $('#lTo').value, q: $('#lQ').value, cat: withCat ? cat : '' });
  const show = () => {
    const cnt = k => items.filter(i => i.category === k).length;
    $('#logKpi').innerHTML = [['', 'ทั้งหมด', 'fa-list', '#0F2A23', items.length], ...Object.entries(LOGCAT).map(([k, [nm, ic, c]]) => [k, nm, ic, c, cnt(k)])]
      .map(([k, nm, ic, c, v]) => `<div class="lk ${k === cat ? 'on' : ''}" data-cat="${k}" style="--c:${c}"><b>${fmt(v)}</b><span><i class="fa-solid ${ic}"></i>${nm}</span></div>`).join('');
    const list = cat ? items.filter(i => i.category === cat) : items;
    if (!list.length) { $('#lRes').innerHTML = '<div class="empty"><i class="fa-solid fa-clipboard"></i>ไม่พบบันทึกประวัติตามเงื่อนไข</div>'; return; }
    const groups = {};
    list.forEach(i => (groups[i.created_at.slice(0, 10)] ||= []).push(i));
    let k = 0;
    $('#lRes').innerHTML = Object.entries(groups).map(([day, arr]) => `
      <div class="logday"><h4><i class="fa-regular fa-calendar"></i>${new Date(day + 'T00:00:00').toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })} · ${arr.length} รายการ</h4>
        <div class="logline">${arr.map(i => {
          const cc = LOGCAT[i.category] || ['', 'fa-circle'];
          return `<div class="logitem cat-${esc(i.category)}" style="animation-delay:${Math.min(k++, 20) * .03}s">
            <div class="ai"><i class="fa-solid ${LOGACT[i.action] || cc[1]}"></i></div>
            <div class="tx"><p>${esc(i.detail)}</p>
              <div class="mt2"><span><i class="fa-solid fa-tag"></i> ${cc[0]}</span>${i.ref ? `<span><i class="fa-solid fa-hashtag"></i> ${esc(i.ref)}</span>` : ''}
              <span><i class="fa-solid fa-user-check"></i> ผู้ทำรายการ / By: <b>${i.actor ? esc(i.actor) : 'ไม่ระบุ'}</b></span></div></div>
            <div class="tm">${i.created_at.slice(11, 16)} น.</div></div>`;
        }).join('')}</div></div>`).join('');
  };
  const load = async () => {
    $('#lRes').innerHTML = loader();
    try {
      const q = params(false);
      ({ items } = await readApi('logs', q));
      loadedWith = q;
      show();
    } catch (e) {
      $('#lRes').innerHTML = `<div class="empty"><i class="fa-solid fa-lock"></i>${esc(e.message)}</div>`;
    }
  };
  view.onclick = e => { const c = e.target.closest('[data-cat]'); if (c) { cat = c.dataset.cat; if (items && params(false) === loadedWith) show(); else load(); } };
  $('#lGo').onclick = load;
  $('#lQ').onkeydown = e => { if (e.key === 'Enter') load(); };
  $('#lCsv').onclick = () => downloadCsv('logs_export', params());
  load();
}

/* ================= SETTINGS (waste types + receivers) ================= */
const ICONS = ['fa-recycle', 'fa-box-open', 'fa-bottle-water', 'fa-gears', 'fa-spray-can', 'fa-fill-drip', 'fa-biohazard', 'fa-lightbulb', 'fa-microchip', 'fa-jug-detergent', 'fa-note-sticky', 'fa-water', 'fa-faucet-drip', 'fa-oil-can', 'fa-flask', 'fa-battery-half', 'fa-dumpster', 'fa-trash-can', 'fa-leaf', 'fa-skull-crossbones', 'fa-fire', 'fa-droplet', 'fa-box'];
const COLORS = ['#059669', '#14B8A6', '#0EA5E9', '#6366F1', '#A855F7', '#EC4899', '#EF4444', '#F97316', '#F59E0B', '#EAB308', '#B7791F', '#8B5E3C', '#64748B', '#0F2A23'];
function settings() {
  const rs = S.receivers || [];
  const activeTypes = S.types.filter(t => +t.active === 1);
  const permitTotal = activeTypes.reduce((a, t) => a + Math.round(Number(t.permit_tons) * 100), 0) / 100;
  view.innerHTML = `${S.types.length ? '' : '<div class="card empty-setup mt-0"><p class="hint"><i class="fa-solid fa-circle-info"></i> ยังไม่มีประเภทของเสีย — เริ่มด้วย “เพิ่มผู้รับดำเนินการ” (หรือเพิ่มผู้รับใหม่ในฟอร์มประเภทของเสียได้เลย) แล้วกด “เพิ่มประเภทของเสีย”</p></div>'}<div class="reset-topbar"><button class="btn btn-primary" id="tAdd"><i class="fa-solid fa-plus"></i>เพิ่มประเภทของเสีย</button><button class="btn btn-primary" id="rAdd"><i class="fa-solid fa-truck-field"></i>เพิ่มผู้รับดำเนินการ</button><button class="btn btn-bad" id="resetBtn">🗑 รีเซ็ตข้อมูล / Reset data</button></div>
    <div class="card"><h3><i class="fa-solid fa-truck-field"></i>ผู้รับดำเนินการ<span class="sub">${rs.length} ราย · โควตารวมของประเภทที่ใช้งาน ${fmtFix(permitTotal, 2)} ตัน</span></h3>
      <p class="hint"><i class="fa-solid fa-circle-info"></i> ผู้รับและประเภทของเสียที่เพิ่มไว้ จะขึ้นในหน้าภาพรวม (โควตาตามผู้รับ / แยกประเภท) ฟอร์มบันทึก-ส่งกำจัด ยอดคงเหลือ ประวัติ และไฟล์ CSV อัตโนมัติ · ผู้รับที่ยังไม่มีประเภทจะมีโควตา 0 ตัน</p>
      <div class="tlist stagger" id="rList">${rs.map(r => `<div class="titem ${+r.active === 1 ? '' : 'off'}" style="--c:var(--brand)"><span class="ti"><i class="fa-solid fa-truck-field"></i></span>
        <div class="tx"><b>${esc(r.name)} ${+r.active === 1 ? '' : '<span class="tag off">ปิดใช้งาน</span>'}</b><small>${r.reg ? 'ทะเบียน ' + esc(r.reg) + ' · ' : ''}${r.mgmt_code ? 'รหัสจัดการ ' + esc(r.mgmt_code) + ' · ' : ''}${fmt(r.types)} ประเภท · โควตา ${fmtFix(r.permit_tons, 2)} ตัน</small></div>
        <button class="btn btn-sm btn-ghost" data-redit="${r.id}" data-rname="${esc(r.name)}"><i class="fa-solid fa-pen"></i>แก้ไข</button>
        ${r.types ? '' : `<button class="btn btn-sm btn-ghost" data-rdel="${r.id}" data-rname="${esc(r.name)}" title="ลบ (ใช้ได้เมื่อไม่มีประเภทของเสีย)"><i class="fa-solid fa-trash-can"></i></button>`}</div>`).join('') || '<div class="empty">ยังไม่มีผู้รับดำเนินการ</div>'}</div></div>
    <div class="card mt"><h3><i class="fa-solid fa-list-check"></i>ประเภทของเสียตามหนังสืออนุญาต<span class="sub">${S.types.length} รายการ · อนุญาตรวม ${fmtFix(permitTotal, 2)} ตัน</span></h3>
      <div class="tlist stagger">${S.types.map(t => `<div class="titem ${+t.active === 1 ? '' : 'off'}" style="--c:${esc(t.color)}"><span class="ti"><i class="fa-solid ${esc(t.icon)}"></i></span>
        <div class="tx"><b>${esc(t.code)} · ${esc(t.name)} ${+t.active === 1 ? '' : '<span class="tag off">ปิดใช้งาน</span>'}</b><small>อนุญาต ${fmtFix(t.permit_tons, 2)} ตัน · รหัสจัดการ ${esc(t.mgmt_code)} · ${esc(t.receiver)}${t.receiver_reg ? ' (ทะเบียน ' + esc(t.receiver_reg) + ')' : ''}${t.has_moves ? ' · มีรายการแล้ว' : ''}</small></div>
        <button class="btn btn-sm btn-ghost" data-tedit="${t.id}"><i class="fa-solid fa-pen"></i>แก้ไข</button>
        ${t.has_moves ? '' : `<button class="btn btn-sm btn-ghost" data-tdel="${t.id}" title="ลบ (ใช้ได้เมื่อยังไม่มีรายการ)"><i class="fa-solid fa-trash-can"></i></button>`}</div>`).join('')}</div></div>`;
  view.onclick = e => {
    const b = e.target.closest('[data-tedit],[data-redit],[data-rdel],[data-tdel]');
    if (!b) return;
    if (b.dataset.tedit) typeForm(typeById(b.dataset.tedit));
    else if (b.dataset.redit !== undefined) receiverForm(rs.find(r => r.name === b.dataset.rname));
    else if (b.dataset.rdel) confirmDelete('ลบผู้รับดำเนินการ “' + b.dataset.rname + '”?', 'receiver_delete', { id: b.dataset.rdel });
    else if (b.dataset.tdel) { const t = typeById(b.dataset.tdel); confirmDelete('ลบประเภทของเสีย ' + t.code + ' ' + t.name + '?', 'type_delete', { id: t.id }); }
  };
  $('#tAdd').onclick = () => typeForm(null);
  $('#rAdd').onclick = () => receiverForm(null);
  $('#resetBtn').onclick = resetData;
}
async function confirmDelete(msg, action, body) {
  if (!confirm(msg + '\n\n(ลบได้เฉพาะรายการที่ยังไม่ถูกใช้ — ถ้ามีรายการเคลื่อนไหวแล้วให้ปิดใช้งานแทน)')) return;
  try { const r = await writeApi(action, body); toast(r.message, 'ok'); route(); } catch (e) { toast(e.message, 'err'); }
}
// permit tons are typed with 2 decimals (older values with more are shown as they are)
const permitStr = v => { const n = Number(v || 0); return Math.abs(Math.round(n * 100) - n * 100) < 1e-6 ? n.toFixed(2) : String(n); };
function receiverForm(r) {
  const m = recordForm(`<h3><i class="fa-solid ${r ? 'fa-pen-to-square' : 'fa-truck-field'}" style="color:var(--brand)"></i> ${r ? 'แก้ไขผู้รับดำเนินการ' : 'เพิ่มผู้รับดำเนินการ'}</h3>
    <div class="fields mt">
      <div class="field full"><label>ชื่อผู้รับดำเนินการ</label><input id="rName" maxlength="120" value="${esc(r?.name || '')}"></div>
      <div class="field"><label>เลขทะเบียนผู้รับ</label><input id="rReg" inputmode="numeric" maxlength="20" value="${esc(r?.reg || '')}"></div>
      <div class="field"><label>รหัสวิธีการจัดการ (3 หลัก · ไม่บังคับ)</label><input id="rMg" inputmode="numeric" maxlength="3" value="${esc(r?.mgmt_code || '')}"></div>
      <div class="field full"><label class="switch"><input type="checkbox" id="rAct" ${!r || +r.active === 1 ? 'checked' : ''}> ใช้งาน</label></div>
      ${r ? '<p class="hint full"><i class="fa-solid fa-circle-info"></i> การแก้ชื่อ / เลขทะเบียน จะเปลี่ยนในทุกประเภทของเสียของผู้รับนี้ · ปิดใช้งานได้เมื่อปิดประเภทของเสียทั้งหมดของผู้รับแล้ว</p>' : ''}
    </div>`, f => writeApi('receiver_save', { id: r?.id || '', current_name: r && !r.id ? r.name : '', name: $('#rName', f).value, reg: $('#rReg', f).value, mgmt_code: $('#rMg', f).value, active: $('#rAct', f).checked ? 1 : 0 }));
  return m;
}
function typeForm(t) {
  let ic = t?.icon || ICONS[0], col = t?.color || COLORS[0];
  const rs = (S.receivers || []).filter(r => +r.active === 1 || r.name === t?.receiver);
  const NEW = '__new';
  const cur = t ? (rs.find(r => r.name === t.receiver)?.name ?? '') : (rs[0]?.name ?? NEW);
  const m = recordForm(`<h3><i class="fa-solid ${t ? 'fa-pen-to-square' : 'fa-plus'}" style="color:var(--brand)"></i> ${t ? 'แก้ไขประเภทของเสีย' : 'เพิ่มประเภทของเสีย'}</h3>
    <div class="fields mt">
      <div class="field"><label>รหัสของเสีย (6 หลัก)</label><input id="tCode" inputmode="numeric" maxlength="6" value="${esc(t?.code || '')}"></div>
      <div class="field"><label>รหัสวิธีการจัดการ (3 หลัก)</label><input id="tMg" inputmode="numeric" maxlength="3" value="${esc(t?.mgmt_code || '')}"></div>
      <div class="field full"><label>ชื่อของเสีย</label><input id="tName" maxlength="120" value="${esc(t?.name || '')}"></div>
      <div class="field"><label>ปริมาณที่ได้รับอนุญาต (ตัน/ปี) · ทศนิยม 2 ตำแหน่ง</label><input id="tTons" inputmode="decimal" placeholder="0.00" value="${t ? permitStr(t.permit_tons) : ''}"></div>
      <div class="field"><label>ผู้รับดำเนินการ</label><select id="tRecv">${rs.map(r => `<option value="${esc(r.name)}" ${r.name === cur ? 'selected' : ''}>${esc(r.name)}${r.reg ? ' (' + esc(r.reg) + ')' : ''}</option>`).join('')}<option value="${NEW}" ${cur === NEW ? 'selected' : ''}>+ เพิ่มผู้รับรายใหม่…</option></select></div>
      <div class="field full" id="tNewBox" ${cur === NEW ? '' : 'hidden'}><label>ผู้รับรายใหม่ — ชื่อ และเลขทะเบียน</label>
        <div class="fields"><div class="field"><input id="tNewName" maxlength="120" placeholder="ชื่อผู้รับดำเนินการ"></div><div class="field"><input id="tNewReg" inputmode="numeric" maxlength="20" placeholder="เลขทะเบียนผู้รับ"></div></div></div>
      <div class="field full"><label>ไอคอน</label><div class="icon-pick">${ICONS.map(i => `<button type="button" data-ic="${i}" class="${i === ic ? 'sel' : ''}"><i class="fa-solid ${i}"></i></button>`).join('')}</div></div>
      <div class="field full"><label>สี</label><div class="color-pick">${COLORS.map(c => `<button type="button" data-col="${c}" style="--c:${c}" class="${c.toUpperCase() === col.toUpperCase() ? 'sel' : ''}"></button>`).join('')}</div></div>
      <div class="field full"><label class="switch"><input type="checkbox" id="tAct" ${!t || +t.active === 1 ? 'checked' : ''}> ใช้งาน (แสดงในหน้าบันทึกจัดเก็บ และนับโควตาของผู้รับ)</label></div>
    </div>`, f => {
    const tons = $('#tTons', f).value.replace(/,/g, '').trim();
    if (!/^\d+(\.\d{1,2})?$/.test(tons)) return Promise.reject(new Error('ปริมาณที่ได้รับอนุญาต (ตัน) ต้องเป็นตัวเลข ทศนิยมไม่เกิน 2 ตำแหน่ง เช่น 12.50'));
    const sel = $('#tRecv', f).value, isNew = sel === NEW;
    const rc = isNew ? null : (S.receivers || []).find(r => r.name === sel);
    return writeApi('type_save', { id: t?.id || '', code: $('#tCode', f).value, mgmt_code: $('#tMg', f).value, name: $('#tName', f).value, permit_tons: tons,
      ...(isNew ? { receiver: $('#tNewName', f).value, receiver_reg: $('#tNewReg', f).value } : rc?.id ? { receiver_id: rc.id } : { receiver: sel, receiver_reg: rc?.reg || '' }),
      icon: ic, color: col, active: $('#tAct', f).checked ? 1 : 0 });
  });
  const sel = $('#tRecv', m);
  sel.onchange = () => {
    $('#tNewBox', m).hidden = sel.value !== NEW;
    const rc = (S.receivers || []).find(r => r.name === sel.value);
    if (rc?.mgmt_code && !$('#tMg', m).value) $('#tMg', m).value = rc.mgmt_code;
  };
  m.addEventListener('click', e => {
    const i = e.target.closest('[data-ic]'), c = e.target.closest('[data-col]');
    if (i) { ic = i.dataset.ic; $$('[data-ic]', m).forEach(x => x.classList.toggle('sel', x === i)); }
    if (c) { col = c.dataset.col; $$('[data-col]', m).forEach(x => x.classList.toggle('sel', x === c)); }
  });
}

/* ================= RESET DATA ================= */
const RESET_CONFIRM = [
  'ยืนยันการรีเซ็ตข้อมูล?', '',
  'ระบบจะลบข้อมูลต่อไปนี้ทั้งหมด:',
  '• รายการจัดเก็บ / ส่งกำจัด / ยอดยกมา (movements) — ยอดคงเหลือทุกประเภทจะกลับเป็น 0',
  '• บันทึกประวัติ (audit_logs)', '',
  'ข้อมูลที่ยังเก็บไว้: ประเภทของเสีย (waste_types)', '',
  'การลบนี้ย้อนกลับไม่ได้ ต้องการดำเนินการต่อหรือไม่?',
].join('\n');
function resetData() {
  const m = modal(`<h3><i class="fa-solid fa-trash-can" style="color:var(--bad)"></i> รีเซ็ตข้อมูล / Reset data</h3>
    <p class="hint">การรีเซ็ตลบข้อมูลทั้งหมดและย้อนกลับไม่ได้ — ใส่รหัสผ่านผู้ดูแลระบบอีกครั้งเพื่อยืนยัน</p>
    <div class="field mt"><label>รหัสผ่านผู้ดูแล</label><input type="password" id="rPw" autocomplete="off"></div>
    <div id="rErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px"></div>
    <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn btn-bad" id="rGo"><i class="fa-solid fa-trash-can"></i>รีเซ็ตข้อมูล</button></div>`);
  const inp = $('#rPw', m), go = $('#rGo', m), err = $('#rErr', m);
  const idle = () => { go.disabled = false; go.innerHTML = '<i class="fa-solid fa-trash-can"></i>รีเซ็ตข้อมูล'; };
  const submit = async () => {
    if (go.disabled) return;
    let pw = inp.value;
    if (!pw) { err.textContent = 'กรุณาใส่รหัสผ่าน'; inp.focus(); return; }
    if (!confirm(RESET_CONFIRM)) { inp.focus(); return; }
    go.disabled = true; go.innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังรีเซ็ต...'; err.textContent = '';
    try {
      const reads = routeReads();
      const r = await gasPost({ action: 'reset_data', resetPassword: pw, client: CLIENT_ID, ...readsField(reads) });
      afterWrite(r, reads);
      pw = ''; inp.value = '';
      closeModal();
      const c = r.removed || {};
      toast(`รีเซ็ตข้อมูลแล้ว — ลบรายการ ${fmt(c.movements)} · Log ${fmt(c.audit_logs)} แถว`);
      route();
    } catch (e) {
      pw = ''; inp.value = '';
      err.textContent = e.message; idle(); inp.focus();
    }
  };
  go.onclick = submit;
  inp.onkeydown = e => { if (e.key === 'Enter') submit(); };
  inp.focus();
}

paintRole();
route();
