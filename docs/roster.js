// Розбір графіка: пошук колонки дати, списку людей, кольорів і рядка-розділювача.
const LAT = { A: 'А', B: 'В', C: 'С', E: 'Е', H: 'Н', I: 'І', K: 'К', M: 'М', O: 'О', P: 'Р', T: 'Т', X: 'Х', Y: 'У' };
export const norm = s => String(s ?? '').replace(/ /g, ' ').trim().toUpperCase().replace(/[A-Z]/g, ch => LAT[ch] || ch);
export const letters = s => norm(s).replace(/[^А-ЯІЇЄҐ]/g, '');
export const cleanName = s => String(s ?? '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();

function hsl(hex) {
  const r = parseInt(hex.slice(1, 3), 16) / 255, g = parseInt(hex.slice(3, 5), 16) / 255, b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  let h = 0, s = 0; const l = (max + min) / 2;
  if (d) {
    s = d / (1 - Math.abs(2 * l - 1));
    if (max === r) h = ((g - b) / d) % 6; else if (max === g) h = (b - r) / d + 2; else h = (r - g) / d + 4;
    h *= 60; if (h < 0) h += 360;
  }
  return [h, s, l];
}
export function isLightGreen(hex) { if (!hex) return false; const [h, s, l] = hsl(hex); return h >= 65 && h <= 150 && s >= 0.25 && l >= 0.5; }
export function isBrightGreen(hex) { if (!hex) return false; const [h, s, l] = hsl(hex); return h >= 90 && h <= 165 && s >= 0.6 && l >= 0.2 && l < 0.5; }
export function isBlue(hex) { if (!hex) return false; const [h, s] = hsl(hex); return h >= 185 && h <= 235 && s >= 0.5; }

const PIB = /^П\.?\s*І\.?\s*[БП]\.?$/;

export function findDayRow(sheet) {
  for (let r = 1; r <= Math.min(sheet.maxRow, 80); r++)
    for (let c = 1; c <= sheet.maxCol; c++)
      if (sheet.get(r, c).num === 1 && sheet.get(r, c + 1).num === 2 && sheet.get(r, c + 2).num === 3) return r;
  return null;
}

export function analyze(sheet) {
  const dayRow = findDayRow(sheet);
  if (!dayRow) return { error: 'Не знайшов рядок із числами місяця (1, 2, 3…)' };
  let title = '';
  for (let r = 1; r < dayRow && !title; r++)
    for (let c = 1; c <= sheet.maxCol; c++) { const t = sheet.get(r, c).text; if (/ГРАФІК/i.test(t)) { title = cleanName(t); break; } }
  const days = [];
  for (let c = 1; c <= sheet.maxCol; c++) { const n = sheet.get(dayRow, c).num; if (Number.isInteger(n) && n >= 1 && n <= 31) days.push(n); }
  return { dayRow, title, minDay: Math.min(...days), maxDay: Math.max(...days) };
}

export function locate(sheet, dayRow, day) {
  let dayCol = null;
  for (let c = 1; c <= sheet.maxCol; c++) if (sheet.get(dayRow, c).num === day) { dayCol = c; break; }
  if (!dayCol) return { error: `У графіку немає колонки для числа ${day}` };
  let nameCol = null;
  for (let r = 1; r <= dayRow; r++)
    for (let c = 1; c <= dayCol; c++) if (PIB.test(norm(sheet.get(r, c).text)) && (nameCol === null || c > nameCol)) nameCol = c;
  if (nameCol === null) nameCol = 2;
  return { dayCol, nameCol, permCol: nameCol + 1 };
}

// Кандидати на «синю строку»: рядки без прізвища, у яких є синя заливка.
export function findSeparators(sheet, dayRow, nameCol) {
  const out = [];
  for (let r = dayRow + 1; r <= sheet.maxRow; r++) {
    if (cleanName(sheet.get(r, nameCol).text)) continue;
    let blue = false, text = '';
    for (let c = 1; c <= sheet.maxCol; c++) {
      const cell = sheet.get(r, c);
      if (isBlue(cell.fill)) blue = true;
      if (!text && cell.text && cell.num === null) text = cleanName(cell.text);
    }
    if (blue) out.push({ row: r, text });
  }
  const pick = out.find(s => s.text) || out[out.length - 1] || null;
  return { list: out, selected: pick ? pick.row : 0 };
}

export function readPeople(sheet, loc, dayRow) {
  const people = [];
  for (let r = dayRow + 1; r <= sheet.maxRow; r++) {
    const name = cleanName(sheet.get(r, loc.nameCol).text);
    if (!name || PIB.test(norm(name))) continue;
    const day = sheet.get(r, loc.dayCol);
    people.push({
      row: r, name,
      permRaw: cleanName(sheet.get(r, loc.permCol).text),
      perm: letters(sheet.get(r, loc.permCol).text),
      mark: norm(day.text),
      fill: day.fill,
    });
  }
  return people;
}
