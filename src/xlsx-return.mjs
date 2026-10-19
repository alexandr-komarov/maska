// Чтение и возврат xlsx. Формулы (f) не трогаются никогда, числа меняются только когда само число
// надо спрятать: ячейка становится строкой inlineStr с меткой. Чтение и замена видят ячейку одинаково.
import { readZip, writeZip, byName } from './zip.mjs';
import { tokenize, join, prefixFor, prepare, collectParas, replaceSegs, editSegs, writeSegs, segText, replaceStr, replaceAllText, tokText,
  authorMap, processProps, propsText, scanLeft, newReport, finishReport, isXmlPart, getAttr, setAttr, decode, encText, attrs, textBuilder } from './ooxml.mjs';

const SNS = ['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main'];
const ANS = ['http://schemas.openxmlformats.org/drawingml/2006/main', 'http://purl.oclc.org/ooxml/drawingml/main'];
const HF = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'];

// абзац в xlsx: строка si, строка ячейки is, текст примечания text; фонетика rPh отдельно
const cfgS = p => ({ para: new Set([p + 'si', p + 'is', p + 'text', p + 'rPh']), text: new Set([p + 't']), space: true });
const cfgD = p => ({ para: new Set([p + 'p']), text: new Set([p + 't']), els: new Map([[p + 'br', '\n'], [p + 'tab', '\t']]), space: false });
const pfx = toks => prefixFor(toks, SNS) ?? '';
const parasNamed = (toks, cfg, name) => collectParas(toks, cfg).filter(pa => toks[pa.start].name === name);

// Связи части: [{id, type, path}], путь считается от папки части
function resolve(dir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const out = dir.split('/').filter(Boolean);
  for (const s of target.split('/')) { if (s === '..') out.pop(); else if (s && s !== '.') out.push(s); }
  return out.join('/');
}
async function relsOf(z, part) {
  const dir = part.replace(/[^/]*$/, ''), r = z[`${dir}_rels/${part.slice(dir.length)}.rels`];
  if (!r) return [];
  return tokenize(await r.text()).filter(t => t.kind === 'tag' && !t.close && /Relationship$/.test(t.name)).map(t => {
    const ext = getAttr(t.raw, 'TargetMode') === 'External', target = getAttr(t.raw, 'Target') || '';
    return { id: getAttr(t.raw, 'Id'), type: (getAttr(t.raw, 'Type') || '').replace(/^.*\//, ''), path: ext ? null : resolve(dir, target) };
  });
}

// Книга: листы в порядке workbook.xml, путь листа по связям, скрытые листы, система дат 1904
async function bookOf(z) {
  const path = 'xl/workbook.xml';
  if (!z[path]) throw new Error('в архиве нет книги Excel (xl/workbook.xml)');
  const rels = await relsOf(z, path), byId = Object.fromEntries(rels.map(r => [r.id, r]));
  const sheets = []; let date1904 = false;
  for (const t of tokenize(await z[path].text())) {
    if (t.kind !== 'tag' || t.close) continue;
    if (/(^|:)workbookPr$/.test(t.name)) date1904 = /^(1|true)$/.test(getAttr(t.raw, 'date1904') || '');
    if (!/(^|:)sheet$/.test(t.name)) continue;
    const id = attrs(t.raw).find(([k]) => /(^|:)id$/.test(k) && k !== 'sheetId')?.[1];
    const r = byId[id];
    if (!r?.path) continue;
    sheets.push({ name: getAttr(t.raw, 'name') ?? '', state: getAttr(t.raw, 'state') || 'visible', path: r.path });
  }
  const styles = rels.find(r => r.type === 'styles')?.path || 'xl/styles.xml';
  const shared = rels.find(r => r.type === 'sharedStrings')?.path || 'xl/sharedStrings.xml';
  return { sheets, date1904, styles, shared };
}

// Стили: для каждого номера стиля ячейки (атрибут s) признак «показывать как дату»
const BUILTIN_DATE = id => (id >= 14 && id <= 17) || id === 22 || (id >= 27 && id <= 36) || (id >= 50 && id <= 58);
const isDateCode = code => /[dy]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''));
async function dateStyles(z, path) {
  if (!z[path]) return [];
  const custom = {}, xfs = []; let inXfs = false;
  for (const t of tokenize(await z[path].text())) {
    if (t.kind !== 'tag') continue;
    const n = t.name.replace(/^[^:]*:/, '');
    if (n === 'numFmt' && !t.close) custom[getAttr(t.raw, 'numFmtId')] = getAttr(t.raw, 'formatCode') || '';
    else if (n === 'cellXfs') inXfs = !t.close && !t.self;
    else if (n === 'xf' && inXfs && !t.close) {
      const id = getAttr(t.raw, 'numFmtId') || '0';
      xfs.push(id in custom ? isDateCode(custom[id]) : BUILTIN_DATE(+id));
    }
  }
  return xfs;
}

// Дата из числа дней Excel: дд.мм.гггг. В системе 1900 есть несуществующее 29.02.1900 (день 60)
export function dateText(x, d1904 = false) {
  if (!isFinite(x)) return null;
  let days = Math.floor(x) + (d1904 ? 1462 : 0);
  if (days < 1 || days > 2958465) return null;
  if (days === 60 && !d1904) return '29.02.1900';
  if (days < 60 && !d1904) days += 1;
  const d = new Date(Date.UTC(1899, 11, 30) + days * 864e5);
  const p2 = n => String(n).padStart(2, '0');
  return `${p2(d.getUTCDate())}.${p2(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`;
}

// Число так, как его видит человек в Excel: 7779.0 → 7779, 9.841726358E9 → 9841726358, шум 0.30000000000000004 → 0.3.
// Целое длиннее 15 цифр Excel хранит с потерей точности: approx
const expand = s => {
  const m = s.match(/^(-?)(\d)(?:\.(\d+))?e\+?(\d+)$/i);
  if (!m) return s;
  const dg = m[2] + (m[3] || ''), exp = +m[4];
  return m[1] + (dg.length <= exp + 1 ? dg + '0'.repeat(exp + 1 - dg.length) : dg.slice(0, exp + 1) + '.' + dg.slice(exp + 1));
};
export function numText(v) {
  const s = v.trim();
  if (!/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(s)) return { text: v };
  const neg = s[0] === '-' ? '-' : '';
  if (/^[-+]?\d+$/.test(s)) { const d = s.replace(/^[-+]?0*(?=\d)/, ''); return { text: neg + d, approx: d.length > 15 }; }
  if (/^[-+]?\d+\.0+$/.test(s)) { const d = s.replace(/^[-+]?0*(?=\d)/, '').replace(/\.0+$/, ''); return { text: neg + d, approx: d.length > 15 }; }
  const x = Number(s);
  if (!isFinite(x)) return { text: v };
  if (Number.isInteger(x)) {
    const plain = Math.abs(x) < 1e21 ? String(Math.abs(x)) : '';
    if (plain && plain.length <= 15) return { text: neg + plain };
    return { text: neg + expand(Math.abs(x).toPrecision(15)), approx: true };
  }
  const r = String(Number(x.toPrecision(15)));
  return { text: /e/i.test(r) ? v : r };
}

const colIdx = ref => { const m = (ref || '').match(/^[A-Z]+/); return m ? [...m[0]].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1 : -1; };
const colName = n => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };

// Ячейки листа: [{ref, row, col, t, s, open, close, f, v, is}]. Ячейка без адреса r получает адрес по месту в строке
function cellsOf(toks, p) {
  const cells = []; let c = null, f = null, row = 0, col = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.kind !== 'tag') continue;
    if (t.name === p + 'row' && !t.close) { row = +getAttr(t.raw, 'r') || row + 1; col = 0; continue; }
    if (t.name === p + 'c') {
      if (t.close) { if (c) { c.close = i; cells.push(c); } c = null; continue; }
      let ref = getAttr(t.raw, 'r');
      if (ref && /^[A-Z]+\d+$/.test(ref)) { col = colIdx(ref); row = +ref.replace(/^[A-Z]+/, ''); } else ref = colName(col) + row;
      if (!t.self) c = { ref, row, col, t: getAttr(t.raw, 't'), s: +(getAttr(t.raw, 's') || 0), open: i };
      col++;
      continue;
    }
    if (!c) continue;
    if (t.name === p + 'f') {
      if (t.self) c.f = { open: i, close: i, text: '' };
      else if (!t.close) f = { open: i, text: '' };
      else if (f) { f.close = i; f.text = toks.slice(f.open + 1, i).map(x => decode(x.raw)).join(''); c.f = f; f = null; }
    } else if (t.name === p + 'v' && !t.close && !t.self) {
      let j = i + 1; while (!(toks[j].kind === 'tag' && toks[j].close)) j++;
      c.v = { open: i, tok: j > i + 1 ? i + 1 : -1, close: j, text: toks.slice(i + 1, j).map(x => decode(x.raw)).join('') };
    } else if (t.name === p + 'is' && !t.close && !t.self) c.is = i;
  }
  return cells;
}

// Что показать в тексте для ячейки. num: число или дата, замена в таком месте идёт по месту
function cellText(c, ctx, inl) {
  if (c.t === 's') return { text: c.v ? ctx.shared[+c.v.text] ?? '' : '' };
  if (c.t === 'inlineStr') return { text: c.is != null ? inl?.get(c.is) ?? '' : '' };
  if (!c.v) return { text: '' };
  if (c.t === 'str' || c.t === 'e' || c.t === 'b') return { text: segText([{ val: c.v.text }]) };
  if (c.t === 'd') {
    const m = c.v.text.match(/^(\d{4})-(\d\d)-(\d\d)/);
    return m ? { text: `${m[3]}.${m[2]}.${m[1]}`, num: true, date: true } : { text: c.v.text, num: true };
  }
  if (ctx.dates[c.s]) { const d = dateText(Number(c.v.text), ctx.d1904); if (d) return { text: d, num: true, date: true }; }
  const r = numText(c.v.text);
  return { text: r.text, num: true, approx: r.approx };
}

async function contextOf(z) {
  const book = await bookOf(z);
  let shared = [];
  if (z[book.shared]) {
    const toks = tokenize(await z[book.shared].text()), p = pfx(toks);
    shared = parasNamed(toks, cfgS(p), p + 'si').map(pa => segText(pa.segs));
  }
  return { book, shared, dates: await dateStyles(z, book.styles), d1904: book.date1904 };
}

// Текст книги той же склейкой, что и замена. Листы построчно, ячейки через табуляцию, пустые строки пропускаются.
// После листа его колонтитулы, примечания к ячейкам и надписи. progress(номер листа, всего, имя) зовётся перед каждым листом
export async function extractXlsxText(bytes, { progress } = {}) {
  const zip = readZip(bytes), z = byName(zip);
  const ctx = await contextOf(z);
  const tb = textBuilder(), seen = new Set();
  const info = { sheets: 0, hidden: [], rows: 0, formulas: 0, dates: 0, approx: [], comments: 0, shapes: 0 };
  const sheets = ctx.book.sheets;
  for (const [k, sh] of sheets.entries()) {
    await progress?.(k, sheets.length, sh.name);
    const e = z[sh.path];
    if (!e) continue;
    info.sheets++;
    const mark = sh.state === 'veryHidden' ? ' (очень скрытый лист)' : sh.state === 'hidden' ? ' (скрытый лист)' : '';
    if (mark) info.hidden.push({ name: sh.name, state: sh.state });
    tb.line(`=== Лист «${sh.name}»${mark} ===`);
    const toks = tokenize(await e.text()), p = pfx(toks);
    const inl = new Map(parasNamed(toks, cfgS(p), p + 'is').map(pa => [pa.start, segText(pa.segs)]));
    const rows = new Map();
    for (const c of cellsOf(toks, p)) {
      if (c.f) info.formulas++;
      const v = cellText(c, ctx, inl);
      if (v.approx) info.approx.push(`${sh.name}!${c.ref}`);
      if (v.date) info.dates++;
      if (v.text === '') continue;
      if (!rows.has(c.row)) rows.set(c.row, []);
      rows.get(c.row).push({ col: c.col, key: `${sh.path}!${c.ref}`, text: v.text, num: v.num });
    }
    for (const [, cells] of [...rows].sort((a, b) => a[0] - b[0])) {
      cells.sort((a, b) => a.col - b.col);
      const parts = []; let col = 0;
      for (const c of cells) { parts.push('\t'.repeat(Math.max(0, c.col - col)), c); col = c.col; }
      tb.line(...parts); info.rows++;
    }
    // колонтитулы листа
    let inHf = 0, hfHead = false;
    for (const [i, t] of toks.entries()) {
      if (t.kind === 'tag' && HF.includes(t.name.slice(p.length))) { inHf += t.close ? -1 : t.self ? 0 : 1; continue; }
      if (inHf && t.kind === 'text') {
        if (!hfHead) { tb.line('--- колонтитулы листа ---'); hfHead = true; }
        tb.line({ key: `${sh.path}@${i}`, text: tokText(t) });
      }
    }
    // примечания к ячейкам и надписи этого листа, по связям листа
    for (const r of await relsOf(z, sh.path)) {
      if (!r.path || !z[r.path] || seen.has(r.path)) continue;
      if (r.type === 'comments') { seen.add(r.path); info.comments += await commentLines(tb, z[r.path], r.path); }
      if (r.type === 'drawing') { seen.add(r.path); info.shapes += await drawingLines(tb, z[r.path], r.path, '--- надписи на листе ---'); }
    }
  }
  // примечания и рисунки без связи с листом
  for (const e of zip.entries) {
    if (seen.has(e.name)) continue;
    if (/^xl\/comments\d*\.xml$/.test(e.name)) info.comments += await commentLines(tb, e, e.name, `=== Примечания (${e.name}) ===`);
    else if (/^xl\/drawings\/drawing\d+\.xml$/.test(e.name)) info.shapes += await drawingLines(tb, e, e.name, `=== Надписи (${e.name}) ===`);
  }
  await propsLines(tb, z);
  info.pivots = zip.entries.filter(e => /^xl\/pivotTables\/.*\.xml$/.test(e.name)).length;
  info.embeddings = zip.entries.filter(e => /^xl\/embeddings\//.test(e.name)).length;
  info.media = zip.entries.filter(e => /^xl\/media\//.test(e.name)).length;
  return { ...tb.done(), info };
}

// Примечания: заголовок с адресом ячейки и текст примечания отдельным местом
async function commentLines(tb, e, part, head) {
  const toks = tokenize(await e.text()), p = pfx(toks);
  let n = 0;
  for (const pa of parasNamed(toks, cfgS(p), p + 'text')) {
    let ref = '';
    for (let i = pa.start; i >= 0; i--) if (toks[i].kind === 'tag' && !toks[i].close && toks[i].name === p + 'comment') { ref = getAttr(toks[i].raw, 'ref') || ''; break; }
    if (head && !n) tb.line(head);
    tb.line(`--- примечание к ${ref} ---`);
    tb.line({ key: `${part}#${pa.start}`, text: segText(pa.segs) });
    n++;
  }
  return n;
}
async function drawingLines(tb, e, part, head) {
  const toks = tokenize(await e.text()), p = prefixFor(toks, ANS) ?? 'a:';
  const paras = collectParas(toks, cfgD(p)).filter(pa => segText(pa.segs).trim());
  if (paras.length) tb.line(head);
  for (const pa of paras) tb.line({ key: `${part}#${pa.start}`, text: segText(pa.segs) });
  return paras.length;
}
// Свойства файла: название, тема, ключевые слова и свои свойства
export async function propsLines(tb, z) {
  let head = false;
  for (const part of ['docProps/core.xml', 'docProps/custom.xml']) {
    if (!z[part]) continue;
    for (const { i, text } of propsText(await z[part].text())) {
      if (!head) { tb.line('=== Свойства файла ==='); head = true; }
      tb.line({ key: `${part}@${i}`, text });
    }
  }
}

// Правки по месту для одной части: ключ «часть@токен», «часть#абзац», «лист!A1»
export const editsFor = (edits, part, sep) => {
  const m = new Map();
  if (edits) for (const [k, v] of edits) if (k.startsWith(part + sep)) { const s = k.slice(part.length + 1); m.set(sep === '!' ? s : +s, v); }
  return m;
};

// Число в ячейке совпадает со строкой из списка: тот же вид, что в тексте (дата, число без хвоста .0),
// или строка из цифр с разделителями не короче 6 цифр (телефон, ИНН). Длинные числа сравниваются по 15 значащим цифрам
function numMatch(disp, nd) {
  if (nd.n === disp) return 1;
  if (!/^-?\d+$/.test(disp) || !/^[\d\s()+\-.]+$/.test(nd.n)) return 0;
  const dg = nd.n.replace(/\D/g, '');
  if (dg.length < 6) return 0;
  if (dg === disp.replace(/^-/, '')) return 1;
  if (dg.length > 15 && Number(disp).toPrecision(15) === Number(dg).toPrecision(15)) return 2;
  return 0;
}

// Excel не принимает кусок <r> с пустым <t> и предлагает восстановить файл. После замены такие куски убираем целиком
// вместе с rPr; если пусты все куски строки, остаётся один <t/>. Трогаем только изменённые строки si, is, text
const changed = pa => pa.segs.some(g => g.el ? g.dead : g.val !== g.orig);
export function dropEmptyRuns(toks, p, paras) {
  for (const pa of paras) {
    const name = toks[pa.start].name;
    if (![p + 'si', p + 'is', p + 'text'].includes(name) || !changed(pa)) continue;
    const runs = []; let cur = null, inT = false;
    for (let i = pa.start + 1; i < toks.length; i++) {
      const t = toks[i];
      if (t.kind === 'tag' && t.close && t.name === name) break;
      if (t.kind === 'tag' && t.name === p + 'r') {
        if (t.self) runs.push({ open: i, close: i, text: '' });
        else if (!t.close) cur = { open: i, text: '' };
        else if (cur) { cur.close = i; runs.push(cur); cur = null; }
        continue;
      }
      if (t.kind === 'tag' && t.name === p + 't' && !t.self) { inT = !t.close; continue; }
      if (cur && inT && (t.kind === 'text' || t.kind === 'cdata')) cur.text += t.raw;
    }
    const empty = runs.filter(r => r.text === '');
    if (!empty.length) continue;
    const blank = r => { for (let k = r.open; k <= r.close; k++) toks[k].raw = ''; };
    if (empty.length === runs.length) { runs.forEach(blank); toks[runs[0].open].raw = `<${p}t/>`; }
    else empty.forEach(blank);
  }
}

// Возврат. replacements: пары [строка, метка] и т. п. (см. prepare); edits: Map ключ места → [{a, b, to}], правки по месту
// Ссылается ли хоть одна формула книги на ячейку, где число стало строкой.
// Протянутая формула (shared) записана один раз, поэтому для неё совпадение по столбцу или строке.
const colNum = s => [...s].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
function markUsed(report) {
  const fs = report.formulas || []; delete report.formulas;
  if (!report.converted.length) return;
  const REF = /(?:(?:'((?:[^']|'')+)'|([A-Za-zА-Яа-яЁё0-9_.]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?|(?:(?:'((?:[^']|'')+)'|([A-Za-zА-Яа-яЁё0-9_.]+))!)?\$?([A-Z]{1,3}):\$?([A-Z]{1,3})(?![A-Za-z0-9(])/g;
  const areas = [];
  for (const f of fs) for (const m of f.text.replace(/"[^"]*"/g, '""').matchAll(REF)) {
    if (m[3]) {
      const c1 = colNum(m[3]), r1 = +m[4], c2 = m[5] ? colNum(m[5]) : c1, r2 = m[6] ? +m[6] : r1;
      areas.push({ sheet: (m[1] || m[2] || f.sheet).replace(/''/g, "'"), c1: Math.min(c1, c2), c2: Math.max(c1, c2), r1: Math.min(r1, r2), r2: Math.max(r1, r2), wide: f.wide });
    } else {
      const c1 = colNum(m[9]), c2 = colNum(m[10]);
      areas.push({ sheet: (m[7] || m[8] || f.sheet).replace(/''/g, "'"), c1: Math.min(c1, c2), c2: Math.max(c1, c2), r1: 1, r2: Infinity, wide: false });
    }
  }
  for (const cv of report.converted) {
    const i = cv.cell.lastIndexOf('!'), sh = cv.cell.slice(0, i), mm = /^([A-Z]+)(\d+)$/.exec(cv.cell.slice(i + 1));
    if (!mm) continue;
    const c = colNum(mm[1]), r = +mm[2];
    cv.used = areas.some(a => a.sheet === sh && (a.wide ? (c >= a.c1 && c <= a.c2) || (r >= a.r1 && r <= a.r2) : c >= a.c1 && c <= a.c2 && r >= a.r1 && r <= a.r2));
  }
}

export async function replaceInXlsx(bytes, replacements, { edits } = {}) {
  const zip = readZip(bytes), z = byName(zip);
  const needles = prepare(replacements);
  const report = newReport();
  const author = authorMap();
  const add = (part, n) => { if (n) report.replaced[part] = (report.replaced[part] || 0) + n; };
  const ctx = await contextOf(z);
  const sheetOf = Object.fromEntries(ctx.book.sheets.map(s => [s.path, s.name]));
  report.hiddenSheets = ctx.book.sheets.filter(s => s.state !== 'visible').map(s => ({ name: s.name, state: s.state }));
  let pivot = false, media = 0;

  for (const e of zip.entries) {
    const name = e.name;
    if (/^xl\/pivot(Cache|Tables)\//.test(name)) pivot = true;
    if (!isXmlPart(name)) {
      if (/^xl\/embeddings\//.test(name)) report.notHandled.push({ kind: 'embedding', part: name, note: 'встроенный объект, внутрь не заглядываем' });
      if (/^xl\/media\//.test(name)) media++;
      continue;
    }
    const src = await e.text();
    let out = src;
    if (name === ctx.book.shared || /^xl\/comments\d*\.xml$/.test(name)) {
      const toks = tokenize(src), p = pfx(toks);
      const paras = collectParas(toks, cfgS(p)), ed = editsFor(edits, name, '#');
      let n = 0;
      for (const pa of paras) {
        if (ed.has(pa.start)) n += editSegs(pa.segs, ed.get(pa.start));
        n += replaceSegs(pa.segs, needles); writeSegs(toks, pa.segs);
      }
      if (name !== ctx.book.shared) {
        // авторы примечаний; имя автора в начале текста примечания заменяем тем же номером
        const names = [];
        for (let i = 0; i < toks.length; i++) {
          const t = toks[i];
          if (t.kind === 'tag' && !t.close && !t.self && t.name === p + 'author' && toks[i + 1]?.kind === 'text') {
            const a = decode(toks[i + 1].raw); names.push([a, author(a)]);
            toks[i + 1] = { raw: encText(author(a)), kind: 'text' }; report.authors++;
          }
        }
        const an = prepare(names);
        if (an.length) for (const pa of paras) { replaceSegs(pa.segs, an); writeSegs(toks, pa.segs); }
      }
      dropEmptyRuns(toks, p, paras);
      add(name, n);
      out = join(toks);
    } else if (sheetOf[name] != null) {
      out = processSheet(name, sheetOf[name], src, needles, report, add, ctx, edits);
    } else if (/^xl\/drawings\/drawing\d+\.xml$/.test(name) || /^xl\/charts\/chart\d+\.xml$/.test(name)) {
      const toks = tokenize(src), p = prefixFor(toks, ANS) ?? 'a:', ed = editsFor(edits, name, '#');
      let n = 0;
      for (const pa of collectParas(toks, cfgD(p))) {
        if (ed.has(pa.start)) n += editSegs(pa.segs, ed.get(pa.start));
        n += replaceSegs(pa.segs, needles); writeSegs(toks, pa.segs);
      }
      n += replaceAllText(toks, needles, (tag, open) => tag === 'c:v' && open.includes('c:strCache'));
      add(name, n);
      out = join(toks);
    } else if (/^xl\/threadedComments\/.*\.xml$/.test(name)) {
      const toks = tokenize(src);
      add(name, replaceAllText(toks, needles, tag => /(^|:)text$/.test(tag)));
      out = join(toks);
    } else if (/^xl\/persons\/.*\.xml$/.test(name)) {
      const toks = tokenize(src);
      for (const t of toks) if (t.kind === 'tag' && !t.close) for (const k of ['displayName', 'userId']) {
        const a = getAttr(t.raw, k); if (a) { t.raw = setAttr(t.raw, k, author(a)); report.authors++; }
      }
      out = join(toks);
    } else if (name === 'xl/workbook.xml') {
      // имена листов и определённые имена не меняем: на них ссылаются формулы
      const toks = tokenize(src);
      for (const t of toks) if (t.kind === 'tag' && !t.close && /(^|:)sheet$/.test(t.name)) {
        const r = replaceStr(getAttr(t.raw, 'name') || '', needles.map(x => ({ ...x, count: 0 })));
        if (r.n) report.notHandled.push({ kind: 'sheetName', sheet: getAttr(t.raw, 'name'), note: 'имя листа содержит строку из списка: не переименовано, на имя ссылаются формулы' });
      }
      let dn = null;
      for (const t of toks) {
        if (t.kind === 'tag' && /(^|:)definedName$/.test(t.name)) { dn = t.close ? null : getAttr(t.raw, 'name'); continue; }
        if (dn && t.kind === 'text') {
          const r = replaceStr(decode(t.raw), needles.map(x => ({ ...x, count: 0 })));
          if (r.n) report.warnings.push({ kind: 'definedName', name: dn, note: 'определённое имя содержит строку из списка, не изменено' });
        }
      }
    } else if (/^docProps\/(core|app|custom)\.xml$/.test(name)) {
      out = processProps(name, src, needles, report, editsFor(edits, name, '@'));
    }
    if (out !== src) { e.setData(out); report.changedParts.push(name); }
  }

  for (const e of zip.entries) if (isXmlPart(e.name)) report.leftovers.push(...scanLeft(e.name, await e.text(), needles));
  // кэш сводных таблиц, заголовки умных таблиц, внешние ссылки не меняем: перечисляем, где осталось
  for (const part of new Set(report.leftovers.map(l => l.part))) {
    const m = part.match(/^xl\/(pivotCache|tables|externalLinks|queryTables|charts|drawings)\//);
    if (m) report.notHandled.push({ kind: m[1], part, note: 'в части остались строки из списка, часть не обрабатывается' });
  }
  if (pivot) report.notHandled.push({ kind: 'pivot', note: 'в книге есть сводные таблицы, их копия данных не обрабатывается' });
  if (media) report.notHandled.push({ kind: 'media', count: media, note: 'картинки не проверяются' });
  markUsed(report);
  finishReport(report, needles);
  return { bytes: await writeZip(zip), report };
}

function processSheet(name, sheet, src, needles, report, add, ctx, edits) {
  const toks = tokenize(src), p = pfx(toks);
  const cells = cellsOf(toks, p), ed = editsFor(edits, name, '!');
  let n = 0;
  // строки внутри ячеек: сначала правки по месту, потом общий список
  const isEd = new Map(cells.filter(c => c.is != null && ed.has(c.ref)).map(c => [c.is, ed.get(c.ref)]));
  const isParas = collectParas(toks, { ...cfgS(p), para: new Set([p + 'is', p + 'rPh']) });
  for (const pa of isParas) {
    if (isEd.has(pa.start)) n += editSegs(pa.segs, isEd.get(pa.start));
    n += replaceSegs(pa.segs, needles); writeSegs(toks, pa.segs);
  }
  // колонтитулы листа
  n += replaceAllText(toks, needles, tag => !!tag && HF.includes(tag.slice(p.length)), editsFor(edits, name, '@'));
  // ячейка получает свою строку вместо числа или общей строки, стиль остаётся
  const toInline = (c, text) => {
    toks[c.open].raw = setAttr(toks[c.open].raw, 't', 'inlineStr');
    toks[c.v.open].raw = `<${p}is><${p}t${/^\s|\s$/.test(text) ? ' xml:space="preserve"' : ''}>${encText(text)}</${p}t></${p}is>`;
    for (let k = c.v.open + 1; k <= c.v.close; k++) toks[k].raw = '';
  };
  for (const c of cells) {
    const where = `${sheet}!${c.ref}`, ce = ed.get(c.ref);
    if (c.f && c.f.text) (report.formulas ||= []).push({ sheet, text: c.f.text, wide: getAttr(toks[c.f.open].raw, 't') === 'shared' });
    if (c.f) {
      for (const nd of needles) if (c.f.text && replaceStr(c.f.text, [{ ...nd, count: 0 }]).n)
        report.warnings.push({ kind: 'formula', cell: where, from: nd.from, note: 'строка из списка внутри формулы, формула не изменена' });
    }
    if (!c.v || c.v.tok < 0) continue;
    if (c.t === 's') {
      // значение заходит в соседнюю ячейку: правка только в этой ячейке, общая строка остаётся для остальных
      if (!ce) continue;
      const segs = [{ val: ctx.shared[+c.v.text] ?? '' }];
      if (!editSegs(segs, ce)) continue;
      replaceSegs(segs, needles);
      toInline(c, segs[0].val); n++;
      continue;
    }
    if (c.t === 'str' || c.t === 'e' || c.t === 'b') {
      // кэш результата формулы: заменяем, иначе данные видны в файле
      const segs = [{ val: c.v.text }];
      const k = (ce ? editSegs(segs, ce) : 0) + replaceSegs(segs, needles);
      if (k) { toks[c.v.tok] = { raw: encText(segs[0].val), kind: 'text' }; n += k; report.warnings.push({ kind: 'cached', cell: where, note: 'заменён кэш результата формулы; после пересчёта Excel вычислит значение заново' }); }
      continue;
    }
    if (c.t && c.t !== 'n' && c.t !== 'd') continue;
    const disp = cellText(c, ctx).text;
    let to = null, approx = false;
    if (ce) { const segs = [{ val: disp }]; if (editSegs(segs, ce)) { replaceSegs(segs, needles); to = segs[0].val; n++; } }
    if (to == null) for (const nd of needles) {
      const m = numMatch(disp, nd);
      if (!m) continue;
      nd.count++; n++; to = nd.to; approx = m === 2;
      break;
    }
    if (to == null) continue;
    if (c.f) {
      // число из формулы: кэш становится строкой, формула остаётся
      toks[c.open].raw = setAttr(toks[c.open].raw, 't', 'str');
      toks[c.v.tok] = { raw: encText(to), kind: 'text' };
      report.warnings.push({ kind: 'cached', cell: where, note: 'кэш формулы совпал с числом из списка: заменён строкой' });
    } else {
      toInline(c, to);
      report.converted.push({ cell: where, to, approx });
    }
  }
  dropEmptyRuns(toks, p, isParas);
  add(name, n);
  return join(toks);
}
