// Файл целиком для интерфейса Маски: чтение, план замены по находкам маскировщика, возврат в исходном формате,
// проверка результата и предупреждения для человека. Без DOM: всё проверяется тестами на Node.
import { extractDocxText, replaceInDocx } from './docx-return.mjs';
import { extractXlsxText, replaceInXlsx } from './xlsx-return.mjs';

const te = new TextEncoder();
export const extOf = name => (String(name).match(/\.([^.]+)$/)?.[1] || '').toLowerCase();
const isOle = u8 => u8.length >= 4 && u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0;
const isZip = u8 => u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4B;

// Понятные сообщения вместо ошибок разбора
export const OLD_FORMAT = {
  doc: 'Это файл старого формата Word (.doc). Откройте его в Word и сохраните как .docx: «Файл», «Сохранить как», тип «Документ Word». Потом загрузите новый файл.',
  xls: 'Это файл старого формата Excel (.xls). Откройте его в Excel и сохраните как .xlsx: «Файл», «Сохранить как», тип «Книга Excel». Потом загрузите новый файл.',
  ole: 'Файл сохранён в старом формате Word или Excel или защищён паролем. Откройте его в Word или Excel, снимите пароль и сохраните как .docx или .xlsx.',
  zip: 'Файл повреждён или это не документ Word и не книга Excel. Откройте его в Word или Excel и сохраните заново как .docx или .xlsx.',
};
export const TEXT_EXT = ['csv', 'md', 'txt'];
// Суммы прячутся в текстах и остаются в таблицах (xlsx, csv), как в плане MVP: третий параметр mask(text, terms, {sums})
export const hideSums = name => !/\.(xlsx|csv)$/i.test(String(name));

// Чтение файла: {kind, ext, text, units, info, note}. kind: docx, xlsx или text.
// progress(номер листа, всего, имя листа) зовётся перед каждым листом книги
export async function readOffice(name, input, { progress } = {}) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const ext = extOf(name);
  if (ext === 'doc' || ext === 'xls') throw new Error(OLD_FORMAT[ext]);
  if (ext === 'docx' || ext === 'xlsx') {
    if (isOle(bytes)) throw new Error(OLD_FORMAT.ole);
    if (!isZip(bytes)) throw new Error(OLD_FORMAT.zip);
    let r;
    try { r = ext === 'docx' ? await extractDocxText(bytes) : await extractXlsxText(bytes, { progress }); }
    catch (e) { throw new Error(/zip|ZIP|CRC|архив/.test(e.message) ? OLD_FORMAT.zip + ' (' + e.message + ')' : e.message); }
    return { kind: ext, ext, text: r.text, units: r.units, info: r.info, note: ext === 'docx' ? docxNote(r.info) : xlsxNote(r.info) };
  }
  if (isOle(bytes)) throw new Error(OLD_FORMAT.ole);
  // текст: UTF-8, если не читается, то кодировка Windows (cp1251), как сохраняет Excel csv
  let text, enc = 'utf-8';
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
  catch { text = new TextDecoder('windows-1251').decode(bytes); enc = 'windows-1251'; }
  const bom = text.charCodeAt(0) === 0xFEFF;
  if (bom) text = text.slice(1);
  const crlf = /\r\n/.test(text);
  return { kind: 'text', ext, text: text.replace(/\r\n?/g, '\n'), units: [], info: { enc, bom, crlf },
    note: `${ext || 'текст'}: как есть` + (enc !== 'utf-8' ? ', кодировка Windows-1251, результат сохранится в UTF-8' : '') };
}

const plural = (n, a, b, c) => { const m = n % 10, h = n % 100; return `${n} ${m === 1 && h !== 11 ? a : m >= 2 && m <= 4 && (h < 12 || h > 14) ? b : c}`; };
function docxNote(i) {
  const extra = [];
  if (i.parts.length) extra.push('ещё: ' + i.parts.join(', ').toLowerCase());
  if (i.revisions) extra.push(`исправления в ${plural(i.revisions, 'абзаце', 'абзацах', 'абзацах')}, удалённый текст в конце`);
  return 'docx: текст, таблицы, надписи' + (extra.length ? '; ' + extra.join('; ') : '');
}
function xlsxNote(i) {
  const extra = [];
  if (i.hidden.length) extra.push(`скрытые листы: ${i.hidden.map(h => '«' + h.name + '»').join(', ')}`);
  if (i.dates) extra.push(`даты показаны как дд.мм.гггг`);
  if (i.comments) extra.push(`примечаний ${i.comments}`);
  return `xlsx: листов ${i.sheets}, строк ${i.rows}, формул ${i.formulas}, в текст ушли значения` + (extra.length ? '; ' + extra.join('; ') : '');
}

// Места, которые задевает отрезок [s, e): места идут по возрастанию начала
function unitsIn(units, s, e) {
  let lo = 0, hi = units.length;
  while (lo < hi) { const m = (lo + hi) >> 1; if (units[m].e <= s) lo = m + 1; else hi = m; }
  const out = [];
  for (let k = lo; k < units.length && units[k].s < e; k++) if (units[k].e > s) out.push(units[k]);
  return out;
}

// Короткое значение (номер «12», «А-5», число меньше 6 цифр) как подстрока встречается в файле и там,
// где маска его не находила: «12» сидит в «2012» и в «112 кв. м». Такое значение меняется только по месту находки
export const isShort = v => v.replace(/\s/g, '').length < 4 || (/^[\d\s.,\/№#()+-]+$/.test(v) && v.replace(/\D/g, '').length < 6);

// План замены по находкам маскировщика (hits: {s, e, value, tag}) и сейфу (vault: {tag, value, forms}).
// Находка внутри одного абзаца или ячейки идёт в общий список «значение → метка»: так оно заменится везде в файле.
// Правки по месту получают находка, которая заходит в соседнюю ячейку или абзац, находка в ячейке с числом или датой
// и короткое значение: метка в первом куске, остальные куски пустеют.
// local: место → значение → сколько раз маска его там нашла, для проверки результата
export function planReplacements(read, hits, vault = []) {
  const units = read.units || [], global = new Map(), edits = new Map(), split = new Set(), placed = new Set(), local = new Map();
  const addEdit = (key, e) => { if (!edits.has(key)) edits.set(key, []); edits.get(key).push(e); };
  const addGlobal = (v, tag) => { if (!global.has(v) && !isShort(v) && !/[\t\n]/.test(v)) global.set(v, tag); };
  for (const h of hits) {
    const over = unitsIn(units, h.s, h.e);
    const one = over.length === 1 && over[0].s <= h.s && h.e <= over[0].e;
    if (!over.length) { addGlobal(h.value, h.tag); continue; }
    if (one && !over[0].num && !isShort(h.value)) { addGlobal(h.value, h.tag); continue; }
    let first = true;
    for (const u of over) {
      const a = Math.max(h.s, u.s) - u.s, b = Math.min(h.e, u.e) - u.s;
      if (b <= a || !read.text.slice(u.s + a, u.s + b).trim()) continue;
      addEdit(u.key, { a, b, to: first ? h.tag : '' }); first = false;
    }
    placed.add(h.value);
    if (!one) { split.add(h.value); continue; }
    addGlobal(h.value, h.tag);
    if (isShort(h.value)) {
      const m = local.get(over[0].key) || new Map();
      m.set(h.value, (m.get(h.value) || 0) + 1); local.set(over[0].key, m);
    }
  }
  // формы из сейфа: то же значение в других местах файла
  for (const v of vault) for (const f of v.forms ? [...v.forms] : [v.value]) if (f && !split.has(f)) addGlobal(f, v.tag);
  return { replacements: [...global], edits, placed, split, local };
}

const count = (s, v) => { let n = 0; for (let a = s.indexOf(v); a >= 0; a = s.indexOf(v, a + v.length)) n++; return n; };
const hideTags = (s, tags) => { for (const t of tags) s = s.split(t).join('\u0001'); return s; };

// Что осталось в файле после замены. Значения из общего списка ищутся во всём тексте результата
// (строки-заголовки листов не считаются: имена листов не меняются, об этом отдельное предупреждение).
// Короткие значения проверяются по месту: в абзаце или ячейке их стало меньше ровно на число находок
export async function verifyReturn(kind, bytes, values, tags = [], { read, local } = {}) {
  const r = kind === 'docx' ? await extractDocxText(bytes) : await extractXlsxText(bytes);
  const body = hideTags(r.text.split('\n').filter(l => !l.startsWith('=== Лист «')).join('\n'), tags);
  const left = [...new Set(values)].filter(v => v && !isShort(v) && body.includes(v));
  if (local?.size) {
    const before = new Map(read.units.map(u => [u.key, read.text.slice(u.s, u.e)]));
    const after = new Map(r.units.map(u => [u.key, hideTags(r.text.slice(u.s, u.e), tags)]));
    for (const [key, m] of local) for (const [v, n] of m)
      if (count(after.get(key) ?? '', v) > count(before.get(key) ?? '', v) - n && !left.includes(v)) left.push(v);
  }
  return { left, text: r.text };
}

const MIME = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  csv: 'text/csv;charset=utf-8', md: 'text/markdown;charset=utf-8', txt: 'text/plain;charset=utf-8' };
export const outName = name => { const m = String(name).match(/^(.*?)(\.[^.]+)?$/); return `${m[1] || 'file'}-maska${m[2] || '.txt'}`; };

// Файл в исходном формате. masked: текст с метками (для csv, md, txt), hits и vault из mask() того же текста
export async function buildReturn({ name, bytes, read, hits, vault, masked }) {
  if (read.kind === 'text') {
    let out = masked;
    if (read.info.crlf) out = out.replace(/\n/g, '\r\n');
    if (read.info.bom || read.ext === 'csv') out = '﻿' + out; // Excel открывает csv в UTF-8 только с BOM
    return { bytes: te.encode(out), name: outName(name), type: MIME[read.ext] || MIME.txt, report: null, warnings: [] };
  }
  const src = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const plan = planReplacements(read, hits, vault);
  const { bytes: out, report } = await (read.kind === 'docx' ? replaceInDocx : replaceInXlsx)(src, plan.replacements, { edits: plan.edits });
  const values = [...new Set([...hits.map(h => h.value), ...plan.replacements.map(r => r[0])])];
  const check = await verifyReturn(read.kind, out, values, [...new Set(hits.map(h => h.tag))], { read, local: plan.local });
  return { bytes: out, name: outName(name), type: MIME[read.kind], report, plan, left: check.left, warnings: returnWarnings(read, report, plan, check.left) };
}

const list = (a, n = 3) => a.slice(0, n).map(x => `«${x}»`).join(', ') + (a.length > n ? ` и ещё ${a.length - n}` : '');
const cells = (a, n = 3) => a.slice(0, n).join(', ') + (a.length > n ? ` и ещё ${a.length - n}` : '');
const PART_NAMES = [[/externalLink/, 'ссылки на другие книги'], [/pivotCache/, 'сводные таблицы'], [/tables\//, 'заголовки умных таблиц'],
  [/styles/, 'названия стилей'], [/charts?\//, 'диаграммы'], [/drawings?\//, 'рисунки'], [/settings/, 'настройки документа'], [/customXml/, 'данные полей формы']];

// Предупреждения перед скачиванием: короткие фразы для человека без технической подготовки
export function returnWarnings(read, report, plan, left = []) {
  const w = [], info = read.info || {};
  const nh = k => report.notHandled.filter(x => x.kind === k);
  const sheetNames = nh('sheetName').map(x => x.sheet);
  // значение, которое не нашлось при замене, но и в результате его нет (его съела более длинная замена), не считается
  if (left.length) w.push(`После замены в файле остались найденные значения (${left.length}): ${list(left)}. Проверьте файл перед отправкой.`);
  if (sheetNames.length) w.push(`В названиях листов остались данные: ${list(sheetNames)}. Маска листы не переименовывает, на них ссылаются формулы. Переименуйте листы в Excel.`);
  const fx = [...new Set(report.warnings.filter(x => x.kind === 'formula').map(x => x.cell))];
  if (fx.length) w.push(`В формулах записаны спрятанные данные (${fx.length}: ${cells(fx)}). Формулы Маска не меняет, после пересчёта Excel покажет эти данные снова. Замените такие формулы значениями.`);
  const dn = [...new Set(report.warnings.filter(x => x.kind === 'definedName').map(x => x.name))];
  if (dn.length) w.push(`В именах диапазонов остались данные (${dn.length}). Их видно в Excel на вкладке «Формулы», «Диспетчер имён».`);
  if (nh('pivot').length) w.push('В книге есть сводные таблицы. Их копия данных внутри файла не обезличивается. Удалите сводные таблицы или сохраните их как значения.');
  const emb = nh('embedding');
  if (emb.length) w.push(`В файле есть вложенные объекты (${emb.length}): документ или таблица внутри файла. Маска их не открывает. Удалите их или проверьте отдельно.`);
  const media = nh('media').reduce((a, x) => a + (x.count || 0), 0);
  if (media) w.push(`В файле есть картинки (${media}). Текст на картинках и в их свойствах Маска не проверяет.`);
  if (report.converted.length) {
    const used = report.converted.filter(c => c.used), n = report.converted.length;
    const head = `В ${plural(n, 'ячейке', 'ячейках', 'ячейках')} число заменено меткой (${cells(report.converted.map(c => c.cell))}).`;
    w.push(used.length
      ? `${head} На ${plural(used.length, 'такую ячейку', 'такие ячейки', 'таких ячеек')} ссылаются формулы (${cells(used.map(c => c.cell))}): там появится ошибка или значение не попадёт в расчёт. Проверьте итоги в файле.`
      : `${head} Обычно это ИНН и номера счетов. Формулы на эти ячейки не ссылаются, расчёты не изменятся.`);
  }
  const hidden = report.hiddenSheets || info.hidden || [];
  if (hidden.length) w.push(`Скрытые листы: ${list(hidden.map(h => h.name))}. Они обезличены, но в Excel их не видно, пока не показать скрытые листы.`);
  if (info.approx?.length) w.push(`В ${plural(info.approx.length, 'ячейке', 'ячейках', 'ячейках')} длинное число, больше 15 цифр (${cells(info.approx)}). Excel хранит только первые 15 цифр, остальные потеряны ещё в исходном файле.`);
  const parts = [...new Set(report.leftovers.map(l => l.part).filter(p => p !== 'xl/workbook.xml' && !/^xl\/worksheets\//.test(p) && !/^xl\/(pivotCache|pivotTables)\//.test(p)))];
  if (parts.length) w.push(`Данные остались в служебных частях файла: ${[...new Set(parts.map(p => PART_NAMES.find(([re]) => re.test(p))?.[1] || 'служебные части'))].join(', ')}. На экране их не видно, но они есть внутри файла.`);
  return w;
}
