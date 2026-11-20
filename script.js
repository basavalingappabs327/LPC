// Paste your deployed Apps Script Web App URL here:
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzYCU7y-mB5HcFp7Wu2R68w4Vj_MOpQ2_QQqrsNK_JUgrLnFMg0oNUuTEGDEyf00JWTew/exec';

// Field keys match the {placeholders} in LPC_template_ready.docx
const T = (k, l) => ({k, l});
const D = (k, l) => ({k, l, t: 'date'});
const MONTHS = ['ಏಪ್ರಿಲ್','ಮೇ','ಜೂನ್','ಜುಲೈ','ಆಗಸ್ಟ್','ಸೆಪ್ಟಂಬರ್','ಅಕ್ಟೋಬರ್','ನವಂಬರ್','ಡಿಸೆಂಬರ್','ಜನವರಿ','ಫೆಬ್ರವರಿ','ಮಾರ್ಚಿ'];

const SECTIONS = [
 {t:'Basic details', f:[T('place','ಗೆ (to place)'),T('reason','ಪ್ರಯುಕ್ತ (reason)'),T('office','Office name'),T('officeplace','Office place'),
   T('employee','Employee name (ಇಲ್ಲಿನ … ಇವರ)'),T('designation','Designation (ಪದನಾಮ)'),D('paidupto','2. Paid up to (ದಿನಾಂಕ)'),
   T('gpf','3. PF / GPF A/c No.'),D('handoverdate','4. Charge handed over on'),{k:'session',l:'Session',t:'select',o:['ಪೂರ್ವಾಹ್ನ','ಅಪರಾಹ್ನ']},
   T('joiningtime','8. Joining time (days)'),D('today','Certificate date (ದಿನಾಂಕ)')]},
 {t:'Pay details (ವಿವರಗಳು)', f:[T('pay','Pay'),T('stgcount','Stagnation count'),T('stgamt','Stagnation amount'),T('da','DA'),T('hra','HRA'),
   T('ua','UA'),T('cha','CHA / FTA'),T('ma','MA'),T('ppsfn','PP-SFN'),T('totearn','TOTAL (pay)')]},
 {t:'Additional pay items (optional)', x:'xe'},
 {t:'Deductions (ವಜಾಗಳು)', f:[T('it','IT'),T('pt','PT'),T('gis','GIS'),T('licded','LIC'),T('gpfprem','GPF premium'),T('gpfloan','GPF loan'),
   T('kgidprem','KGID premium'),T('kgidloan','KGID loan'),T('nps','NPS'),T('fa','FA'),T('totded','TOTAL (deductions)')]},
 {t:'Additional deductions (optional)', x:'xd'},
 {t:'Recovery table – common', f:[T('financialyear','Financial year (e.g. 2026-27)'),T('kgidno','HRMS KGID No.'),
   T('lic','LIC policy'),T('kgidpolicy','KGID policy'),T('otherpolicy','Other policy')]},
 {t:'Income tax recovered – month-wise (₹)', f:MONTHS.map((m, i) => T(`m${i + 1}_it`, m))},
];

const form = document.getElementById('lpcForm');
const status = document.getElementById('status');
const keys = [];

function field(f) {
  const el = f.t === 'select' ? document.createElement('select') : document.createElement('input');
  if (f.t === 'select') el.innerHTML = '<option value=""></option>' + f.o.map(o => `<option>${o}</option>`).join('');
  else el.type = f.t === 'date' ? 'date' : 'text';
  el.name = f.k; el.dataset.type = f.t || 'text'; keys.push(f.k);
  const lb = document.createElement('label'); lb.append(f.l, el); return lb;
}

// Optional extra rows (name + amount). Up to MAX_EXTRA per side; each side fills rows 1..n of the table.
const MAX_EXTRA = 6;
let sectionN = 0;
function addExtra(list, name = '', amt = '') {
  if (list.children.length >= MAX_EXTRA) return;
  const row = document.createElement('div'); row.className = 'extra-row';
  row.innerHTML = '<input class="xn" placeholder="Name (e.g. CCA)"><input class="xa" placeholder="Amount"><button type="button" class="ghost" title="Remove">×</button>';
  row.children[0].value = name; row.children[1].value = amt;
  row.children[2].onclick = () => { row.remove(); form.dispatchEvent(new Event('input')); };
  list.append(row); form.dispatchEvent(new Event('input'));
}
function extraSection(s) {
  const fs = document.createElement('fieldset');
  fs.dataset.n = String(++sectionN);
  fs.innerHTML = `<legend>${s.t}</legend><div class="extra-list" data-x="${s.x}"></div>`;
  const b = document.createElement('button'); b.type = 'button'; b.textContent = '+ Add item';
  b.onclick = () => addExtra(fs.querySelector('.extra-list'));
  fs.append(b); form.append(fs);
}
const readExtra = x => [...form.querySelectorAll(`[data-x="${x}"] .extra-row`)]
  .map(r => [r.children[0].value.trim(), r.children[1].value.trim()]);

SECTIONS.forEach(s => {
  if (s.x) return extraSection(s);
  const fs = document.createElement('fieldset');
  fs.dataset.n = String(++sectionN);
  fs.innerHTML = `<legend>${s.t}</legend>`;
  const g = document.createElement('div'); g.className = 'grid';
  s.f.forEach(f => g.append(field(f)));
  fs.append(g); form.append(fs);
});

const currentFY = () => { // Apr–Mar financial year, e.g. 2026-27
  const n = new Date(), y = n.getMonth() >= 3 ? n.getFullYear() : n.getFullYear() - 1;
  return y + '-' + String(y + 1).slice(2);
};
const todayISO = () => new Date().toISOString().slice(0, 10);
const KEY = 'lpc_draft_v4'; // draft is saved in this browser only
try {
  const saved = JSON.parse(localStorage.getItem(KEY) || '{}');
  keys.forEach(k => { if (saved[k] !== undefined) form.elements[k].value = saved[k]; });
  ['xe', 'xd'].forEach(x => (saved['_' + x] || []).forEach(([n, a]) => addExtra(form.querySelector(`[data-x="${x}"]`), n, a)));
} catch (e) {}
if (!form.elements.today.value) form.elements.today.value = todayISO();
if (!form.elements.financialyear.value) form.elements.financialyear.value = currentFY();

// TOTAL fields are calculated automatically — keep them read-only
['totearn', 'totded'].forEach(k => {
  const el = form.elements[k];
  if (el) {
    el.readOnly = true;
    el.tabIndex = -1;
    el.title = 'Auto-calculated';
    el.style.fontWeight = '700';
  }
});

/** Parse a number from user input (allows commas, ₹, spaces). Empty → 0. */
function num(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(String(v).replace(/[₹,\s]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

/** Sum amounts from optional extra rows (second input = amount). */
function sumExtra(x) {
  return readExtra(x).reduce((s, [, a]) => s + num(a), 0);
}

const PAY_KEYS = ['pay', 'stgamt', 'da', 'hra', 'ua', 'cha', 'ma', 'ppsfn']; // stgcount is not an amount
const DED_KEYS = ['it', 'pt', 'gis', 'licded', 'gpfprem', 'gpfloan', 'kgidprem', 'kgidloan', 'nps', 'fa'];

function recalcTotals() {
  const earn = PAY_KEYS.reduce((s, k) => s + num(form.elements[k]?.value), 0) + sumExtra('xe');
  const ded = DED_KEYS.reduce((s, k) => s + num(form.elements[k]?.value), 0) + sumExtra('xd');
  // Show blank when zero so the certificate stays clean if nothing entered
  if (form.elements.totearn) form.elements.totearn.value = earn ? String(earn) : '';
  if (form.elements.totded) form.elements.totded.value = ded ? String(ded) : '';
}

form.addEventListener('input', () => {
  recalcTotals();
  try {
    const d = Object.fromEntries(keys.map(k => [k, form.elements[k].value]));
    d._xe = readExtra('xe'); d._xd = readExtra('xd');
    localStorage.setItem(KEY, JSON.stringify(d));
  } catch (e) {}
});

// Initial calc (e.g. after restoring draft)
recalcTotals();

const fmtDate = v => v ? v.split('-').reverse().join('-') : '';

/** Parse "2026-27" / "2026-2027" → { start: 2026, end: 2027 } */
function parseFY(fy) {
  const m = String(fy || '').match(/(\d{4})\s*[-–]\s*(\d{2,4})/);
  if (!m) return null;
  const start = parseInt(m[1], 10);
  let end = parseInt(m[2], 10);
  if (end < 100) end = Math.floor(start / 100) * 100 + end; // 27 → 2027
  return { start, end };
}

function collect() {
  const data = Object.fromEntries(keys.map(k => {
    const el = form.elements[k];
    return [k, el.dataset.type === 'date' ? fmtDate(el.value) : el.value.trim()];
  }));
  ['xe', 'xd'].forEach(x => {
    const items = readExtra(x).filter(([n, a]) => n || a).map(([n, a]) => n && a ? `${n} - ${a}` : n || a);
    for (let i = 1; i <= MAX_EXTRA; i++) data[x + i] = items[i - 1] || '';
  });
  data.officefull = data.office + (data.office && data.officeplace ? ', ' : '') + data.officeplace;

  // Per-month calendar year: Apr–Dec → start year; Jan–Mar → end year
  // e.g. FY 2026-27 → Apr..Dec = 2026, Jan..Mar = 2027
  const fy = parseFY(data.financialyear) || (() => {
    const n = new Date(), y = n.getMonth() >= 3 ? n.getFullYear() : n.getFullYear() - 1;
    return { start: y, end: y + 1 };
  })();
  for (let i = 1; i <= 12; i++) {
    data['fy' + i] = String(i <= 9 ? fy.start : fy.end);
  }
  return data;
}

async function download(format) {
  if (SCRIPT_URL.startsWith('PASTE')) { alert('Set SCRIPT_URL in script.js first.'); return; }
  const btns = document.querySelectorAll('button'); btns.forEach(b => b.disabled = true);
  status.textContent = 'Generating… (10–20 sec)';
  try {
    // text/plain avoids a CORS preflight, which Apps Script cannot answer
    const res = await fetch(SCRIPT_URL, {method: 'POST', body: JSON.stringify({format, data: collect()})});
    const out = await res.json();
    if (!out.ok) throw new Error(out.error);
    const bin = atob(out.data), bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([bytes], {type: out.mime}));
    a.download = out.filename; document.body.append(a); a.click(); a.remove();
    status.textContent = 'Downloaded ✔';
  } catch (e) {
    status.textContent = 'Error: ' + e.message;
  } finally { btns.forEach(b => b.disabled = false); }
}

document.getElementById('btnPdf').onclick = () => download('pdf');
document.getElementById('btnDocx').onclick = () => download('docx');
document.getElementById('btnClear').onclick = () => {
  if (confirm('Clear all fields?')) {
    form.reset();
    form.elements.today.value = todayISO();
    form.elements.financialyear.value = currentFY();
    form.querySelectorAll('.extra-row').forEach(r => r.remove());
    try { localStorage.removeItem(KEY); } catch (e) {}
    recalcTotals();
  }
};
