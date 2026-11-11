import { loadWorkbook } from './xlsx.js';
import { analyze, locate, findSeparators, readPeople, isLightGreen, isBrightGreen, letters, norm, cleanName } from './roster.js';
import { allocate } from './alloc.js';

const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const DEFAULTS = {
  day: new Date().getDate(), sessions: 2, perSector: 3, starts: [],
  sectors: [{ name: 'WEST', letters: 'СЛ' }, { name: 'SOUTH', letters: 'СО' }, { name: 'NORTH', letters: 'СМ' }],
  marks: 'Н, НА', bright: false, extras: '',
};
let cfg = { ...DEFAULTS };
try { cfg = { ...DEFAULTS, ...JSON.parse(localStorage.getItem('roster-cfg') || '{}'), day: DEFAULTS.day }; } catch {}
let sheet = null, info = null, last = null;

// ---------- форма ----------
function drawSectors() {
  $('sectors').innerHTML = cfg.sectors.map((s, i) => `
    <div class="sector">
      <label>Назва<input type="text" data-i="${i}" data-k="name" value="${esc(s.name)}"></label>
      <label>Літери допуску<input type="text" data-i="${i}" data-k="letters" value="${esc(s.letters)}"></label>
      <button type="button" data-del="${i}" aria-label="Видалити сектор" ${cfg.sectors.length < 2 ? 'disabled' : ''}>✕</button>
    </div>`).join('');
}
function drawStarts() {
  const n = Math.max(1, Math.min(12, +$('sessions').value || 1));
  $('starts').innerHTML = Array.from({ length: n }, (_, i) =>
    `<label>Сесія ${i + 1}<input type="time" value="${esc(cfg.starts[i] || '')}"></label>`).join('');
}
$('sessions').addEventListener('input', () => { readForm(); drawStarts(); });
$('starts').addEventListener('change', readForm);
function fillForm() {
  $('day').value = cfg.day; $('sessions').value = cfg.sessions; $('perSector').value = cfg.perSector;
  $('marks').value = cfg.marks; drawStarts();
  $('bright').checked = cfg.bright; $('extras').value = cfg.extras; drawSectors();
}
function readForm() {
  cfg.day = +$('day').value; cfg.sessions = Math.max(1, +$('sessions').value || 1);
  cfg.perSector = Math.max(1, +$('perSector').value || 1);
  cfg.starts = [...document.querySelectorAll('#starts input')].map(i => i.value); cfg.marks = $('marks').value;
  cfg.bright = $('bright').checked; cfg.extras = $('extras').value;
  try { localStorage.setItem('roster-cfg', JSON.stringify({ ...cfg, day: undefined })); } catch {}
}
$('sectors').addEventListener('input', e => { const t = e.target; if (t.dataset.i) cfg.sectors[+t.dataset.i][t.dataset.k] = t.value; readForm(); });
$('sectors').addEventListener('click', e => {
  const i = e.target.dataset && e.target.dataset.del;
  if (i !== undefined && cfg.sectors.length > 1) { cfg.sectors.splice(+i, 1); drawSectors(); readForm(); }
});
$('addSector').onclick = () => { cfg.sectors.push({ name: 'СЕКТОР ' + (cfg.sectors.length + 1), letters: 'С' }); drawSectors(); readForm(); };
document.querySelectorAll('#cParams input:not([data-i]), #cParams textarea, #cParams select').forEach(el => el.addEventListener('change', readForm));

// ---------- файл ----------
async function openFile(file) {
  $('err').textContent = ''; $('go').disabled = true; $('result').hidden = true;
  try {
    const wb = await loadWorkbook(file);
    let found = null;
    for (let i = 0; i < wb.sheets.length && !found; i++) {
      const sh = await wb.sheet(i); const a = analyze(sh);
      if (!a.error) found = { sh, a, name: wb.sheets[i].name };
    }
    if (!found) throw new Error('Не знайшов аркуш із графіком (рядок з числами 1, 2, 3…).');
    sheet = found.sh; info = found.a;
    $('dropText').textContent = file.name;
    $('fileInfo').textContent = `${info.title || 'Графік'} · аркуш «${found.name}» · числа ${info.minDay}–${info.maxDay}`;
    const nameCol = locate(sheet, info.dayRow, info.minDay).nameCol;
    const seps = findSeparators(sheet, info.dayRow, nameCol);
    $('sep').innerHTML = '<option value="0">немає</option>' + seps.list.map(s =>
      `<option value="${s.row}">Рядок ${s.row} — ${esc(s.text || 'без тексту')}</option>`).join('');
    $('sep').value = String(seps.selected);
    $('go').disabled = false;
  } catch (e) {
    sheet = null; $('err').textContent = e.message || String(e);
  }
}
$('file').addEventListener('change', e => e.target.files[0] && openFile(e.target.files[0]));
const drop = $('drop');
['dragenter', 'dragover'].forEach(n => drop.addEventListener(n, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(n => drop.addEventListener(n, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => e.dataTransfer.files[0] && openFile(e.dataTransfer.files[0]));

// ---------- відбір ----------
const allowedSet = (perm, sectors) => new Set(sectors.flatMap((sec, i) => [...letters(sec.letters)].some(ch => perm.includes(ch)) ? [i] : []));
const STATUS = {
  pick: ['Підходить', 'st-ok'], reserve: ['Резерв', 'st-warn'],
  nogreen: ['Не салатова комірка', 'st-no'], mark: ['Позначка Н/НА', 'st-no'],
  noperm: ['Немає допуску на жоден сектор', 'st-no'], bright: ['Яскраво-зелена (вимкнено)', 'st-no'],
};
function buildPeople() {
  const loc = locate(sheet, info.dayRow, cfg.day);
  if (loc.error) throw new Error(loc.error);
  const sepRow = +$('sep').value;
  const marks = new Set(cfg.marks.split(/[,;\s]+/).map(norm).filter(Boolean));
  const allowedOf = perm => allowedSet(perm, cfg.sectors);

  const extras = cfg.extras.split('\n').map(l => cleanName(l)).filter(Boolean).map(l => {
    const t = l.split(' '); let perm = '';
    if (t.length > 1 && /^[A-Za-zА-ЯІЇЄҐа-яіїєґ]{1,8}$/.test(t[t.length - 1])) perm = letters(t.pop());
    return { name: t.join(' '), perm, extra: true, row: null, allowed: allowedOf(perm) };
  });
  const extraKeys = new Set(extras.map(e => norm(e.name)));

  const all = readPeople(sheet, loc, info.dayRow).filter(p => !extraKeys.has(norm(p.name)));
  for (const p of all) {
    p.below = sepRow > 0 && p.row > sepRow;
    p.allowed = allowedOf(p.perm);
    const light = isLightGreen(p.fill), bright = isBrightGreen(p.fill);
    if (!light && !(bright && cfg.bright)) p.status = bright ? 'bright' : 'nogreen';
    else if (marks.has(p.mark)) p.status = 'mark';
    else if (!p.allowed.size) p.status = 'noperm';
    else p.status = 'pick';
    p.note = p.status === 'pick' && p.mark ? `в комірці «${p.mark}» — перевір` : '';
  }
  const okExtras = extras.filter(e => e.allowed.size);
  return { loc, all, extras, okExtras, sepRow };
}

// ---------- розподіл ----------
function run() {
  readForm();
  $('err').textContent = '';
  try {
    if (!sheet) throw new Error('Спочатку завантаж файл.');
    if (!(cfg.day >= 1 && cfg.day <= 31)) throw new Error('Вкажи число місяця від 1 до 31.');
    const data = buildPeople();
    const eligible = data.all.filter(p => p.status === 'pick');
    const tiers = [data.okExtras, eligible.filter(p => !p.below), eligible.filter(p => p.below)];
    const res = allocate(tiers, cfg.sectors.length, cfg.sessions, cfg.perSector);
    last = { data, res, day: cfg.day, saved: Date.now(), snap: JSON.parse(JSON.stringify({ sectors: cfg.sectors, starts: cfg.starts.slice(0, cfg.sessions), perSector: cfg.perSector, sessions: cfg.sessions })) };
    const placed = new Set(res.grid.flat(2).filter(Boolean));
    for (const p of data.all) if (p.status === 'pick' && !placed.has(p)) p.status = 'reserve';
    render(true);
  } catch (e) { $('err').textContent = e.message || String(e); }
}

const startOf = (snap, si) => (snap.starts && snap.starts[si]) || '';
function tag(p) { return p.perm && !p.perm.includes('С') ? `<span class="tag">${esc(p.perm)}</span>` : ''; }

// у Android-застосунку (Capacitor) не працюють завантаження blob і друк — там замість них «Поділитись»
const NATIVE = !!(window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform());
const sharePlugin = () => window.Capacitor && window.Capacitor.Plugins && window.Capacitor.Plugins.Share;
const canShare = () => (NATIVE && !!sharePlugin()) || !!navigator.share;
async function shareResult() {
  const text = asText();
  try {
    if (NATIVE && sharePlugin()) await sharePlugin().share({ title: 'Розподіл по секторах ACC', text, dialogTitle: 'Поділитись розподілом' });
    else await navigator.share({ title: 'Розподіл по секторах ACC', text });
  } catch {}
}

function chipHtml(p, attrs, bad) {
  return `<span class="chip${bad ? ' bad' : ''}" ${attrs}${bad ? ' title="Поза допуском цього сектора"' : ''}>${esc(p.name)}${tag(p)}${bad ? '<span class="tag warn">⚠</span>' : ''}</span>`;
}

function render(fresh = true) {
  const { data, res, snap } = last;
  const { sectors, perSector } = snap;
  const placed = new Set(res.grid.flat(2).filter(Boolean));
  for (const p of data.all) if (p.status === 'pick' || p.status === 'reserve') p.status = placed.has(p) ? 'pick' : 'reserve';
  res.selectedCount = placed.size;
  persist();
  const out = [];
  out.push(`<div class="bar"><h2>Розподіл на ${last.day} число<span class="when">сформовано ${new Date(last.saved).toLocaleString('uk-UA', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' })}</span></h2>
    <button type="button" class="primary" id="again"${sheet ? '' : ' disabled title="Завантаж графік, щоб перемішати знову"'}>Перемішати ще раз</button>
    <button type="button" id="copy">Копіювати</button>
    ${NATIVE ? '' : '<button type="button" id="csv">CSV</button><button type="button" id="print">Друк</button>'}
    ${canShare() ? '<button type="button" id="share">Поділитись</button>' : ''}</div>
    <p class="note hint">Утримуй людину й перетягни на інше місце, щоб поміняти, або у «Резерв». Торкнись місця, щоб поставити людину зі списку чи вписати вручну.</p>`);
  if (res.selectedCount < res.total) out.push(`<p class="warn">Не вистачає людей: заповнено ${res.selectedCount} з ${res.total} місць.</p>`);
  const missing = data.extras.filter(e => !e.allowed.size);
  if (missing.length) out.push(`<p class="warn">Не додано (немає допуску на жоден сектор): ${missing.map(e => esc(e.name)).join(', ')}.</p>`);
  res.grid.forEach((sess, si) => {
    const range = startOf(snap, si) ? 'початок ' + startOf(snap, si) : '';
    out.push(`<div class="session"><h3>Сесія ${si + 1}<span>${range}</span></h3><table><thead><tr><th></th>${
      sectors.map((s, i) => `<th class="t${i % 5}">${esc(s.name)}</th>`).join('')}</tr></thead><tbody>${
      Array.from({ length: perSector }, (_, pos) => {
        return `<tr><td class="time">${pos + 1}</td>${sess.map((col, ci) => {
          const p = col[pos], at = `data-s="${si}" data-c="${ci}" data-p="${pos}"`;
          const bad = p && p.allowed && !p.allowed.has(ci);
          return `<td class="slot" data-drop="slot" ${at}>${p ? chipHtml(p, `data-drag="slot" ${at}`, bad) : '<span class="empty">—</span>'}</td>`;
        }).join('')}</tr>`;
      }).join('')}</tbody></table></div>`);
  });
  const li = tier => {
    const items = res.rest.map((p, i) => ({ p, i })).filter(x => tier === 1 ? x.p.tier < 2 : x.p.tier === 2);
    return items.length ? `<ul>${items.map(({ p, i }) => `<li>${chipHtml(p, `data-drag="rest" data-i="${i}"`)}</li>`).join('')}</ul>` : '<p class="empty">нікого</p>';
  };
  out.push(`<div class="lists"><div class="card" data-drop="rest" data-tier="1"><h2>Резерв</h2>${li(1)}</div>
    <div class="card" data-drop="rest" data-tier="2"><h2>Резерв (індивідуальний графік)</h2>${li(2)}</div></div>
    <form class="addrow" id="addForm"><input type="text" id="addName" placeholder="Додати людину вручну: Прізвище І." required>
      <input type="text" id="addPerm" placeholder="Допуск" maxlength="8"><button type="submit">Додати до списку</button></form>`);
  const counts = {};
  data.all.forEach(p => counts[p.status] = (counts[p.status] || 0) + 1);
  out.push(`<details><summary>Як відібрано: усі ${data.all.length} людей на ${last.day} число</summary>
    <p class="note">Колонка дати: ${colName(data.loc.dayCol)} · прізвища: ${colName(data.loc.nameCol)} · допуски: ${colName(data.loc.permCol)}${data.sepRow ? ' · синя строка: рядок ' + data.sepRow : ''}.</p>
    <table><thead><tr><th>Рядок</th><th>Прізвище</th><th>Допуск</th><th>Статус</th></tr></thead><tbody>${
    data.all.map(p => { const [t, c] = STATUS[p.status]; return `<tr><td>${p.row}</td><td>${esc(p.name)}${p.below ? ' <span class="tag">нижче</span>' : ''}</td><td>${esc(p.permRaw)}</td><td class="${c}">${t}${p.note ? ' — ' + esc(p.note) : ''}</td></tr>`; }).join('')}</tbody></table></details>`);
  const box = $('result'); box.innerHTML = out.join(''); box.hidden = false;
  $('again').onclick = run;
  if ($('print')) $('print').onclick = () => print();
  if ($('share')) $('share').onclick = shareResult;
  $('copy').onclick = async e => {
    const b = e.target;
    try { await navigator.clipboard.writeText(asText()); b.textContent = 'Скопійовано ✓'; } catch { b.textContent = 'Не вдалося скопіювати'; }
    setTimeout(() => b.textContent = 'Копіювати', 1600);
  };
  if ($('csv')) $('csv').onclick = downloadCsv;
  $('addForm').onsubmit = e => { e.preventDefault(); addManual($('addName').value, $('addPerm').value); render(false); };
  if (fresh) box.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ---------- ручні правки ----------
function manualPerson(name, perm) {
  const pm = letters(perm);
  return { name: cleanName(name), perm: pm, manual: true, tier: 1, row: null, allowed: pm ? allowedSet(pm, last.snap.sectors) : null };
}
function addManual(name, perm) {
  name = cleanName(name); if (!name) return null;
  const same = last.res.rest.find(p => norm(p.name) === norm(name));
  if (same) return same;
  const p = manualPerson(name, perm); last.res.rest.push(p); return p;
}
function applyMove(src, dst) {
  const g = last.res.grid, rest = last.res.rest;
  const person = src.kind === 'slot' ? g[src.s][src.c][src.p] : rest[src.i];
  if (!person) return;
  if (dst.drop === 'slot') {
    const T = g[dst.s][dst.c], q = T[dst.p];
    if (src.kind === 'slot') {
      if (src.s === dst.s && src.c === dst.c && src.p === dst.p) return;
      g[src.s][src.c][src.p] = q || null; T[dst.p] = person;
    } else { rest.splice(src.i, 1); T[dst.p] = person; if (q) rest.push(q); }
  } else {
    person.tier = +dst.tier;
    if (src.kind === 'slot') { g[src.s][src.c][src.p] = null; rest.push(person); }
  }
}

// перетягування: миша — після невеликого зсуву, дотик — після утримання
const box0 = $('result');
let drag = null, suppressClick = false, raf = 0;
const num = d => ({ ...d, kind: d.drag, s: +d.s, c: +d.c, p: +d.p, i: +d.i });
function dropTarget(x, y) { const el = document.elementFromPoint(x, y); return el && el.closest('[data-drop]'); }
function startDrag() {
  if (!drag || drag.active) return;
  drag.active = true;
  const g = drag.chip.cloneNode(true); g.className = 'chip ghost'; document.body.appendChild(g); drag.ghost = g;
  drag.chip.classList.add('dragging'); document.body.classList.add('is-dragging');
  try { navigator.vibrate && navigator.vibrate(12); } catch {}
  moveGhost(); tick();
}
function moveGhost() {
  drag.ghost.style.left = drag.x + 'px'; drag.ghost.style.top = drag.y + 'px';
  document.querySelectorAll('.over').forEach(e => e.classList.remove('over'));
  const t = dropTarget(drag.x, drag.y); if (t) t.classList.add('over');
}
function tick() {
  if (!drag || !drag.active) return;
  if (drag.y < 70) scrollBy(0, -14); else if (drag.y > innerHeight - 70) scrollBy(0, 14);
  raf = requestAnimationFrame(tick);
}
function endDrag(drop) {
  if (!drag) return;
  clearTimeout(drag.timer); cancelAnimationFrame(raf);
  const d = drag; drag = null;
  if (!d.active) return;
  const t = dropTarget(d.x, d.y);
  d.ghost.remove(); document.body.classList.remove('is-dragging');
  document.querySelectorAll('.over').forEach(e => e.classList.remove('over'));
  suppressClick = true; setTimeout(() => suppressClick = false, 350);
  if (drop && t) applyMove(num(d.chip.dataset), num(t.dataset));
  render(false);
}
box0.addEventListener('pointerdown', e => {
  const chip = e.target.closest('.chip');
  if (!chip || (e.pointerType === 'mouse' && e.button !== 0)) return;
  drag = { chip, id: e.pointerId, x0: e.clientX, y0: e.clientY, x: e.clientX, y: e.clientY, active: false, touch: e.pointerType !== 'mouse' };
  if (drag.touch) drag.timer = setTimeout(startDrag, 260);
});
addEventListener('pointermove', e => {
  if (!drag || e.pointerId !== drag.id) return;
  drag.x = e.clientX; drag.y = e.clientY;
  if (!drag.active) {
    const dist = Math.hypot(drag.x - drag.x0, drag.y - drag.y0);
    if (drag.touch) { if (dist > 10) { clearTimeout(drag.timer); drag = null; } } else if (dist > 5) startDrag();
  }
  if (drag && drag.active) moveGhost();
});
addEventListener('pointerup', e => { if (drag && e.pointerId === drag.id) endDrag(true); });
addEventListener('pointercancel', e => { if (drag && e.pointerId === drag.id) endDrag(false); });
document.addEventListener('touchmove', e => { if (drag && drag.active) e.preventDefault(); }, { passive: false });
box0.addEventListener('contextmenu', e => { if (e.target.closest('.chip')) e.preventDefault(); });

// торкнулись місця — діалог: людина зі списку, вручну або очистити
const dlg = $('dlg'); let dlgSlot = null;
box0.addEventListener('click', e => {
  if (suppressClick) return;
  const td = e.target.closest('td.slot'); if (td) openDlg(num(td.dataset));
});
function openDlg(a) {
  dlgSlot = a;
  const { res, snap } = last, cur = res.grid[a.s][a.c][a.p];
  $('dlgTitle').textContent = `Сесія ${a.s + 1} · ${snap.sectors[a.c].name} · місце ${a.p + 1}`;
  $('dlgCur').textContent = cur ? `Зараз: ${cur.name}` : 'Місце вільне';
  $('dlgPick').innerHTML = '<option value="">— не обрано —</option>' + res.rest.map((p, i) => `<option value="${i}">${esc(p.name)}${p.perm ? ' (' + esc(p.perm) + ')' : ''}</option>`).join('');
  $('dlgName').value = ''; $('dlgPerm').value = ''; $('dlgClear').hidden = !cur;
  dlg.showModal();
}
function putInSlot(person) {
  const { res } = last, a = dlgSlot, q = res.grid[a.s][a.c][a.p];
  res.grid[a.s][a.c][a.p] = person; if (q) res.rest.push(q);
}
$('dlgSet').onclick = () => {
  const { res } = last, idx = $('dlgPick').value, name = cleanName($('dlgName').value);
  if (idx !== '') { const p = res.rest.splice(+idx, 1)[0]; putInSlot(p); }
  else if (name) {
    const same = res.rest.findIndex(p => norm(p.name) === norm(name));
    putInSlot(same >= 0 ? res.rest.splice(same, 1)[0] : manualPerson(name, $('dlgPerm').value));
  } else { dlg.close(); return; }
  dlg.close(); render(false);
};
$('dlgClear').onclick = () => {
  const { res } = last, a = dlgSlot, p = res.grid[a.s][a.c][a.p];
  if (p) { res.grid[a.s][a.c][a.p] = null; res.rest.push(p); }
  dlg.close(); render(false);
};
$('dlgCancel').onclick = () => dlg.close();
dlg.addEventListener('click', e => { if (e.target === dlg) dlg.close(); });

function colName(n) { let s = ''; while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); } return s; }

function asText() {
  const { res, snap, day } = last; const lines = [`Розподіл на ${day} число`];
  res.grid.forEach((sess, si) => {
    lines.push('', `Сесія ${si + 1}${startOf(snap, si) ? ` (початок ${startOf(snap, si)})` : ''}`);
    snap.sectors.forEach((s, i) => lines.push(`${s.name}: ${sess[i].map(p => p ? p.name : '—').join('; ')}`));
  });
  const rest = res.rest;
  lines.push('', 'Резерв: ' + (rest.filter(p => p.tier < 2).map(p => p.name).join(', ') || 'нікого'),
    'Резерв (індивідуальний графік): ' + (rest.filter(p => p.tier === 2).map(p => p.name).join(', ') || 'нікого'));
  return lines.join('\n');
}
function downloadCsv() {
  const { res, snap, day } = last; const rows = [['Сесія', 'Слот', 'Сектор', 'Прізвище']];
  res.grid.forEach((sess, si) => sess.forEach((col, i) => col.forEach((p, k) => {
    rows.push([si + 1, k + 1, snap.sectors[i].name, p ? p.name : '']);
  })));
  const blob = new Blob(['﻿' + rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(';')).join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: `rozpodil-${day}.csv` });
  a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

// збереження останнього розподілу: лишається на екрані, доки не сформовано новий
const STORE = 'roster-last';
function persist() {
  try {
    const { data, res, snap, day, saved } = last, ids = new Map(), list = [];
    const reg = p => { if (!p) return null; if (!ids.has(p)) { ids.set(p, list.length); list.push(p); } return ids.get(p); };
    const grid = res.grid.map(s => s.map(c => c.map(reg)));
    const rest = res.rest.map(reg), all = data.all.map(reg), extras = data.extras.map(reg);
    localStorage.setItem(STORE, JSON.stringify({
      v: 1, day, snap, saved, loc: data.loc, sepRow: data.sepRow, total: res.total, grid, rest, all, extras,
      people: list.map(p => ({ name: p.name, perm: p.perm, permRaw: p.permRaw, row: p.row, status: p.status, below: p.below, note: p.note,
        tier: p.tier, manual: p.manual, extra: p.extra, allowed: p.allowed ? [...p.allowed] : null })),
    }));
  } catch {}
}
function restore() {
  try {
    const j = JSON.parse(localStorage.getItem(STORE) || 'null');
    if (!j || j.v !== 1) return;
    const ps = j.people.map(p => ({ ...p, allowed: p.allowed ? new Set(p.allowed) : null }));
    const g = i => (i == null ? null : ps[i]);
    last = { day: j.day, snap: j.snap, saved: j.saved,
      data: { loc: j.loc, sepRow: j.sepRow, all: j.all.map(g), extras: j.extras.map(g) },
      res: { grid: j.grid.map(s => s.map(c => c.map(g))), rest: j.rest.map(g), total: j.total, selectedCount: 0 } };
    render(false);
  } catch {}
}

$('go').onclick = run;
fillForm();
restore();
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
