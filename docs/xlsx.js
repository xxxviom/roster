// Мінімальний читач .xlsx без залежностей: текст, злиті комірки та кольори заливки.
const utf8 = new TextDecoder('utf-8');
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const L = (root, name) => Array.from(root.getElementsByTagNameNS('*', name));
const kids = (el, name) => Array.from(el.children).filter(n => n.localName === name);
const parseXml = s => new DOMParser().parseFromString(s, 'application/xml');

async function inflateRaw(u8) {
  const stream = new Blob([u8]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function readZip(buf) {
  const u8 = new Uint8Array(buf);
  const dv = new DataView(buf);
  let e = u8.length - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error('Це не схоже на файл .xlsx');
  const n = dv.getUint16(e + 10, true);
  let p = dv.getUint32(e + 16, true);
  const entries = new Map();
  for (let i = 0; i < n; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true);
    const elen = dv.getUint16(p + 30, true);
    const clen = dv.getUint16(p + 32, true);
    const lho = dv.getUint32(p + 42, true);
    const name = utf8.decode(u8.subarray(p + 46, p + 46 + nlen));
    entries.set(name, { method, csize, lho });
    p += 46 + nlen + elen + clen;
  }
  const bytes = async name => {
    const en = entries.get(name);
    if (!en) return null;
    const start = en.lho + 30 + dv.getUint16(en.lho + 26, true) + dv.getUint16(en.lho + 28, true);
    const data = u8.subarray(start, start + en.csize);
    return en.method === 0 ? data : inflateRaw(data);
  };
  return { bytes, text: async name => { const b = await bytes(name); return b ? utf8.decode(b) : null; } };
}

// ---------- кольори ----------
function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  let h = 0, s = 0; const l = (max + min) / 2;
  if (d) {
    s = d / (1 - Math.abs(2 * l - 1));
    if (max === r) h = ((g - b) / d) % 6; else if (max === g) h = (b - r) / d + 2; else h = (r - g) / d + 4;
    h *= 60; if (h < 0) h += 360;
  }
  return [h, s, l];
}
function hslToRgb(h, s, l) {
  const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m = l - c / 2;
  const [r, g, b] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  return [r + m, g + m, b + m].map(v => Math.round(v * 255));
}
const hex2 = v => v.toString(16).padStart(2, '0');
function applyTint(hex, tint) {
  if (!tint) return hex;
  const [h, s, l0] = rgbToHsl(parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16));
  const l = tint < 0 ? l0 * (1 + tint) : l0 * (1 - tint) + tint;
  return '#' + hslToRgb(h, s, Math.min(1, Math.max(0, l))).map(hex2).join('');
}

function themeColors(xml) {
  if (!xml) return [];
  const scheme = L(parseXml(xml), 'clrScheme')[0];
  if (!scheme) return [];
  return Array.from(scheme.children).map(n => {
    const c = n.firstElementChild;
    const v = c && (c.getAttribute('val') === 'windowText' ? '000000' : c.getAttribute('val') === 'window' ? 'FFFFFF' : c.getAttribute('lastClr') || c.getAttribute('val'));
    return '#' + (v || '000000').slice(-6).toLowerCase();
  }); // порядок: dk1 lt1 dk2 lt2 accent1..6 hlink folHlink
}
const THEME_MAP = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];

function parseFills(stylesDoc, theme) {
  const fillsEl = L(stylesDoc, 'fills')[0];
  const fills = fillsEl ? kids(fillsEl, 'fill').map(f => {
    const pf = kids(f, 'patternFill')[0];
    if (!pf || pf.getAttribute('patternType') !== 'solid') return null;
    const fg = kids(pf, 'fgColor')[0];
    if (!fg) return null;
    if (fg.hasAttribute('rgb')) return '#' + fg.getAttribute('rgb').slice(-6).toLowerCase();
    if (fg.hasAttribute('theme')) {
      const base = theme[THEME_MAP[+fg.getAttribute('theme')]];
      return base ? applyTint(base, parseFloat(fg.getAttribute('tint') || '0')) : null;
    }
    return null;
  }) : [];
  const xfsEl = L(stylesDoc, 'cellXfs')[0];
  const xfFill = xfsEl ? kids(xfsEl, 'xf').map(x => fills[+x.getAttribute('fillId') || 0] || null) : [];
  return xfFill;
}

const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
const key = (r, c) => r * 100000 + c;

function buildSheet(doc, sst, xfFill) {
  const cells = new Map();
  let maxRow = 0, maxCol = 0;
  for (const c of L(doc, 'c')) {
    const m = /^([A-Z]+)(\d+)$/.exec(c.getAttribute('r') || '');
    if (!m) continue;
    const col = colNum(m[1]), row = +m[2];
    const t = c.getAttribute('t');
    const vEl = kids(c, 'v')[0];
    let text = '', num = null;
    if (t === 's' && vEl) text = sst[+vEl.textContent] || '';
    else if (t === 'inlineStr') text = L(c, 't').map(x => x.textContent).join('');
    else if (vEl) { text = vEl.textContent; if (!t || t === 'n') { const f = parseFloat(text); if (!isNaN(f)) num = f; } }
    const fill = xfFill[+c.getAttribute('s') || 0] || null;
    cells.set(key(row, col), { text, num, fill });
    if (text !== '' || fill) { maxRow = Math.max(maxRow, row); maxCol = Math.max(maxCol, col); }
  }
  const mergeMap = new Map();
  for (const mc of L(doc, 'mergeCell')) {
    const [a, b] = mc.getAttribute('ref').split(':');
    const ma = /^([A-Z]+)(\d+)$/.exec(a), mb = /^([A-Z]+)(\d+)$/.exec(b || a);
    if (!ma || !mb) continue;
    const r1 = +ma[2], c1 = colNum(ma[1]), r2 = +mb[2], c2 = colNum(mb[1]);
    if ((r2 - r1 + 1) * (c2 - c1 + 1) > 20000) continue;
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) mergeMap.set(key(r, c), key(r1, c1));
  }
  return {
    maxRow, maxCol,
    // комірка з урахуванням злиття: текст і заливка беруться з головної комірки, якщо власних немає
    get(r, c) {
      const own = cells.get(key(r, c));
      const mk = mergeMap.get(key(r, c));
      const master = mk ? cells.get(mk) : null;
      return {
        text: (own && own.text) || (master && master.text) || '',
        num: own && own.num != null ? own.num : master ? master.num : null,
        fill: (own && own.fill) || (master && master.fill) || null,
      };
    },
  };
}

export async function loadWorkbook(file) {
  const zip = await readZip(await file.arrayBuffer());
  const wbXml = await zip.text('xl/workbook.xml');
  if (!wbXml) throw new Error('Не вдалося прочитати книгу Excel');
  const relsXml = await zip.text('xl/_rels/workbook.xml.rels');
  const rels = new Map(L(parseXml(relsXml), 'Relationship').map(r => [r.getAttribute('Id'), r.getAttribute('Target')]));
  const sheets = L(parseXml(wbXml), 'sheet').map(s => {
    const id = s.getAttributeNS(NS_R, 'id') || s.getAttribute('r:id');
    let t = rels.get(id) || '';
    t = t.startsWith('/') ? t.slice(1) : 'xl/' + t;
    return { name: s.getAttribute('name'), path: t };
  });
  const sstXml = await zip.text('xl/sharedStrings.xml');
  const sst = sstXml ? L(parseXml(sstXml), 'si').map(si => L(si, 't').map(t => t.textContent).join('')) : [];
  const stylesXml = await zip.text('xl/styles.xml');
  const theme = themeColors(await zip.text('xl/theme/theme1.xml'));
  const xfFill = stylesXml ? parseFills(parseXml(stylesXml), theme) : [];
  return {
    sheets,
    async sheet(i) {
      const xml = await zip.text(sheets[i].path);
      return buildSheet(parseXml(xml), sst, xfFill);
    },
  };
}
