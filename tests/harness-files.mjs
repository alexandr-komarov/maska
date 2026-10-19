// Стенд на настоящих файлах: docx и xlsx читаются тем же readOffice из src/office.mjs, что встроен в index.html,
// дальше mask() из index.html и те же счётчики, что в harness.mjs. Счётчики по типам данных: must_types.
// С --index на старую версию index.html, где ещё есть readDocx и readXlsx, стенд читает файлы её функциями.
// Сеанс: файлы идут по порядку разметки с одним сейфом, как их обезличивает человек. Счётчики по словарю сущностей
// из разметки: доля сущностей с одной меткой во всех файлах, расщеплённые (одна сущность, несколько меток),
// склеенные (одна метка у разных сущностей). Сеанс идёт сам, если в разметке есть entities, или с --session.
// С --fresh сейф на каждом файле новый, как до общего сейфа: для сравнения.
// Запуск: node harness-files.mjs <папка с файлами> <разметка.json> [--json путь] [--index index.html] [--session] [--fresh]
//         node harness-files.mjs --find <файл> [--save] [--index index.html]
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { readOffice } from '../src/office.mjs';
import * as safeApi from '../src/safe.mjs';
import { rulesOf } from './mask-api.mjs';

// аргументы
const argv = process.argv.slice(2), pos = [], opt = {};
for (let i = 0; i < argv.length; i++) {
  const a = argv[i];
  if (a === '--save') opt.save = true;
  else if (a === '--session') opt.session = true;
  else if (a === '--fresh') opt.fresh = true;
  else if (a === '--json' || a === '--find' || a === '--index') opt[a.slice(2)] = argv[++i];
  else pos.push(a);
}
const root = new URL('..', import.meta.url).pathname;
const indexPath = opt.index || path.join(root, 'index.html');
const srcBuf = fs.readFileSync(indexPath), src = srcBuf.toString('utf8');

// Для старых версий index.html. В Node нет DOMParser: минимальная замена ровно под то, что зовут readDocx и readXlsx
// (children, localName, textContent, getAttribute, getElementsByTagName[NS]).
const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const unent = s => s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) =>
  e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : +e.slice(1)) : (ENT[e] ?? m));
class Node {
  constructor(qn, attrs, parent) {
    this.nodeName = qn; this.attrs = attrs; this.parent = parent; this.childNodes = [];
    const i = qn.indexOf(':'); this.localName = i < 0 ? qn : qn.slice(i + 1);
    const pref = i < 0 ? '' : qn.slice(0, i);
    const decl = n => { for (let x = n; x; x = x.parent) { const k = pref ? 'xmlns:' + pref : 'xmlns'; if (x.attrs && k in x.attrs) return x.attrs[k]; } return null; };
    this.namespaceURI = decl(this);
  }
  get children() { return this.childNodes.filter(n => n instanceof Node); }
  get textContent() { return this.childNodes.map(n => typeof n === 'string' ? n : n.textContent).join(''); }
  getAttribute(k) { return this.attrs[k] ?? null; }
  all() { const r = []; (function w(n) { for (const c of n.childNodes) if (c instanceof Node) { r.push(c); w(c); } })(this); return r; }
  getElementsByTagName(q) { return this.all().filter(n => q === '*' || n.nodeName === q); }
  getElementsByTagNameNS(ns, q) { return this.all().filter(n => n.namespaceURI === ns && (q === '*' || n.localName === q)); }
}
globalThis.DOMParser = class {
  parseFromString(s) {
    const doc = new Node('#document', {}, null); let cur = doc;
    const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<!DOCTYPE[^>]*>|<\/([^>\s]+)\s*>|<([^\s/>]+)((?:\s+[^\s=>/]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
    let m;
    while ((m = re.exec(s))) {
      if (m[1] !== undefined) cur.childNodes.push(m[1]);
      else if (m[2]) cur = cur.parent || cur;
      else if (m[3]) {
        const attrs = {}; for (const a of m[4].matchAll(/([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attrs[a[1]] = unent(a[2] ?? a[3]);
        const n = new Node(m[3], attrs, cur); cur.childNodes.push(n); if (!m[5]) cur = n;
      } else if (m[6] !== undefined) cur.childNodes.push(unent(m[6]));
    }
    return doc;
  }
};

// те же куски index.html, что видит пользователь. В новой версии чтения в index.html нет: файлы читает readOffice
// из src/office.mjs (тот же код встроен в страницу), у старой версии свои readDocx и readXlsx
const legacy = src.includes('async function readDocx');
// функции сейфа в странице приходят из блока модулей, здесь из src/safe.mjs; старым версиям они не нужны
const maskJs = rulesOf(src) + (legacy ? '' : '\nconst readDocx = null, readXlsx = null;');
const readJs = legacy ? src.slice(src.indexOf('async function unzip'), src.indexOf('async function loadFile')) : '';
const safeNames = Object.keys(safeApi);
const api = new Function(...safeNames, maskJs + '\n' + readJs + '; return {mask, leaks, readDocx, readXlsx};')(...safeNames.map(n => safeApi[n]));

// суммы прячутся в текстах и остаются в таблицах (xlsx, csv), как в плане MVP
const opts = file => ({ sums: !/\.(xlsx|csv)$/i.test(file) });

async function extract(file) {
  if (!legacy) { const r = await readOffice(path.basename(file), new Uint8Array(fs.readFileSync(file))); return { text: r.text, info: r.note }; }
  const ext = path.extname(file).slice(1).toLowerCase();
  const buf = fs.readFileSync(file);
  if (ext === 'docx') return await api.readDocx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.length));
  if (ext === 'xlsx') return await api.readXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.length));
  return { text: buf.toString('utf8'), info: ext + ': как есть' };
}

const TYPES = ['фио', 'организация', 'инн', 'счёт', 'телефон', 'почта', 'адрес', 'номер_договора', 'сумма', 'дата', 'прочее'];
const outDir = path.join(root, 'out');
const outPath = p => { const r = path.dirname(p) === '.' && !p.startsWith('./') ? path.join(outDir, p) : path.resolve(p); fs.mkdirSync(path.dirname(r), { recursive: true }); return r; };

// режим --find: текст файла и что нашла Маска
if (opt.find) {
  const { text, info } = await extract(opt.find);
  const r = api.mask(text, [], opts(opt.find));
  const hits = [...r.hits].sort((a, b) => a.s - b.s);
  const o = [`# ${path.basename(opt.find)} · ${info}`, ''];
  let off = 0;
  for (const [i, line] of text.split('\n').entries()) {
    o.push(String(i + 1).padStart(4) + ' | ' + line);
    for (const h of hits.filter(h => h.s >= off && h.s <= off + line.length)) o.push('       ↳ ' + JSON.stringify(h.value) + ' [' + h.t + ']');
    off += line.length + 1;
  }
  o.push('', '# найдено всего: ' + hits.length);
  for (const h of hits) o.push(`${h.t.padEnd(8)} ${JSON.stringify(h.value)}`);
  const s = o.join('\n');
  if (opt.save) { const p = outPath(`find-${path.basename(opt.find)}.txt`); fs.writeFileSync(p, s + '\n'); console.log('записано: ' + p); } else console.log(s);
  process.exit(0);
}

if (pos.length < 2) { console.error('Запуск: node harness-files.mjs <папка> <разметка.json> [--json путь]\n        node harness-files.mjs --find <файл> [--save]'); process.exit(2); }
const [dir, markFile] = pos;
// разметка: список файлов или {entities, files}; entities: [{id, type, forms: [написание]}]
const markup = JSON.parse(fs.readFileSync(markFile, 'utf8'));
const marks = Array.isArray(markup) ? markup : markup.files, entities = Array.isArray(markup) ? [] : markup.entities || [];
let mustAll = 0, mustOk = 0, keepAll = 0, keepOk = 0, absentAll = 0;
const byType = {}, perFile = [];
for (const c of marks) {
  let ex;
  try { ex = await extract(path.join(dir, c.file)); } catch (e) { console.log(`\n${c.file}: НЕ ПРОЧИТАН (${e.message})`); perFile.push({ file: c.file, error: e.message }); continue; }
  const r = api.mask(ex.text, [], opts(c.file));
  const mt = c.must_types || {};
  const must = [...new Set([...c.must, ...Object.keys(mt)])]; // типизированная строка тоже обязательна
  const miss = must.filter(s => r.masked.includes(s));
  // строка разметки, которой нет в тексте файла, считалась бы спрятанной: печатаем, чтобы поправить разметку
  const absent = [...must, ...(c.keep || [])].filter(s => !ex.text.includes(s));
  const broken = (c.keep || []).filter(s => !r.masked.includes(s));
  let keepN = (c.keep || []).length;
  for (const [val, type] of Object.entries(c.types || {})) {
    const h = r.hits.find(x => x.value === val); keepN++;
    if (!h || h.t !== type) broken.push(`${val} должно быть ${type}, стало ${h ? h.t : 'не найдено'}`);
  }
  for (const [val, t] of Object.entries(mt)) {
    if (!TYPES.includes(t)) console.log(`  ВНИМАНИЕ: неизвестный тип «${t}» у «${val}»`);
    const b = byType[t] ||= { all: 0, ok: 0 }; b.all++; if (!miss.includes(val)) b.ok++;
  }
  mustAll += must.length; mustOk += must.length - miss.length; keepAll += keepN; keepOk += keepN - broken.length;
  console.log(`\n${c.file} (${ex.info}): спрятано ${must.length - miss.length}/${must.length}, лишнего задето ${broken.length}/${keepN}`);
  for (const s of miss) console.log('  ПРОПУЩЕНО:', s);
  for (const s of broken) console.log('  ЗАДЕТО ЛИШНЕЕ:', s);
  for (const s of absent) console.log('  НЕТ В ТЕКСТЕ ФАЙЛА (поправить разметку):', s);
  absentAll += absent.length;
  perFile.push({ file: c.file, hidden: must.length - miss.length, must: must.length, missed: miss, broken });
}
// файлы в папке без разметки
const marked = new Set(marks.map(c => c.file));
const extra = fs.readdirSync(dir).filter(f => /\.(docx|xlsx|csv|md|txt)$/i.test(f) && !marked.has(f));
if (extra.length) console.log('\nБез разметки (не считались): ' + extra.join(', '));
const pct = (a, b) => (100 * a / Math.max(b, 1)).toFixed(1);
console.log(`\nИТОГО: спрятано ${mustOk}/${mustAll} (${pct(mustOk, mustAll)} %), не задето ${keepOk}/${keepAll}`);
if (absentAll) console.log(`ВНИМАНИЕ: строк разметки нет в тексте файлов: ${absentAll}, они считаются спрятанными`);
if (Object.keys(byType).length) {
  console.log('\nПо типам данных (спрятано из скольких):');
  for (const t of [...TYPES, ...Object.keys(byType).filter(t => !TYPES.includes(t))]) if (byType[t]) console.log(`  ${t.padEnd(15)} ${byType[t].ok}/${byType[t].all}`);
}
// Сеанс с одним сейфом
let sessionOut = null;
if (entities.length || opt.session) {
  if (legacy) { console.log('\nСеанс с сейфом: в этой версии index.html сейфа нет'); process.exit(0); }
  let safe = safeApi.safeNew('стенд');
  let sMust = 0, sMustOk = 0, sKeep = 0, sKeepOk = 0;
  const ents = entities.map(e => ({ e, tags: new Set(), files: new Map(), seen: new Set(), n: 0, miss: [] }));
  // место написания: не внутри другого слова или числа
  const W = /[\p{L}\p{N}]/u;
  const places = (text, f) => { const out = []; for (let i = text.indexOf(f); i >= 0; i = text.indexOf(f, i + 1)) if (!(W.test(text[i - 1] || '') && W.test(f[0])) && !(W.test(text[i + f.length] || '') && W.test(f.at(-1)))) out.push(i); return out; };
  for (const c of marks) {
    let ex; try { ex = await extract(path.join(dir, c.file)); } catch { continue; }
    const r = api.mask(ex.text, [], { ...opts(c.file), safe: opt.fresh ? safeApi.safeNew() : safe }); safe = r.safe;
    const must = [...new Set([...c.must, ...Object.keys(c.must_types || {})])], keep = c.keep || [];
    sMust += must.length; sMustOk += must.filter(x => !r.masked.includes(x)).length;
    sKeep += keep.length; sKeepOk += keep.filter(x => r.masked.includes(x)).length;
    for (const o of ents) for (const f of o.e.forms) for (const i of places(ex.text, f)) {
      o.n++; o.seen.add(c.file);
      const tags = [...new Set(r.hits.filter(h => h.s < i + f.length && h.e > i).map(h => h.tag))];
      if (!tags.length) { o.miss.push(`${c.file}: ${f}`); continue; }
      for (const t of tags) { o.tags.add(t); const fs_ = o.files.get(c.file) || new Set(); fs_.add(t); o.files.set(c.file, fs_); }
    }
  }
  const absentE = ents.filter(o => !o.n);
  const byTag = new Map();
  for (const o of ents) for (const t of o.tags) { const s = byTag.get(t) || new Set(); s.add(o.e.id); byTag.set(t, s); }
  const glued = [...byTag].filter(([, s]) => s.size > 1);
  const gluedIds = new Set(glued.flatMap(([, s]) => [...s]));
  const live = ents.filter(o => o.n), split = live.filter(o => o.tags.size > 1);
  const one = live.filter(o => o.tags.size === 1 && !o.miss.length && !gluedIds.has(o.e.id));
  const multi = live.filter(o => o.seen.size >= 2);
  const multiOne = multi.filter(o => one.includes(o));
  console.log(`\nСЕАНС ${opt.fresh ? 'с новым сейфом на каждом файле (--fresh)' : 'с одним сейфом'}, файлы по порядку разметки: спрятано ${sMustOk}/${sMust} (${pct(sMustOk, sMust)} %), не задето ${sKeepOk}/${sKeep}, записей в сейфе ${safe.records.length}`);
  if (live.length) {
    console.log(`Сущностей ${live.length}: одна метка во всех файлах у ${one.length} (${pct(one.length, live.length)} %); из них встречаются в двух файлах и больше ${multi.length}, одна метка у ${multiOne.length} (${pct(multiOne.length, multi.length)} %)`);
    console.log(`Расщеплённые (одна сущность, несколько меток): ${split.length}`);
    for (const o of split) console.log(`  ${o.e.id}: ` + [...o.files].map(([f, t]) => `${f} ${[...t].join(' ')}`).join('; '));
    console.log(`Склеенные (одна метка у разных сущностей): ${glued.length}`);
    for (const [t, s] of glued) console.log(`  ${t}: ${[...s].join(', ')}`);
    const missed = live.filter(o => o.miss.length);
    console.log(`Не спрятаны в части мест: ${missed.length}`);
    for (const o of missed) console.log(`  ${o.e.id}: ${o.miss.join('; ')}`);
  }
  if (absentE.length) console.log(`ВНИМАНИЕ: написаний этих сущностей нет ни в одном файле (поправить разметку): ${absentE.map(o => o.e.id).join(', ')}`);
  sessionOut = { hidden: sMustOk, must: sMust, kept: sKeepOk, keep: sKeep, entities: live.length, one_label: one.length, multi_file: multi.length, multi_file_one_label: multiOne.length,
    split: split.map(o => ({ id: o.e.id, files: Object.fromEntries([...o.files].map(([f, t]) => [f, [...t]])) })),
    glued: glued.map(([t, s]) => ({ tag: t, ids: [...s] })), missed: live.filter(o => o.miss.length).map(o => ({ id: o.e.id, places: o.miss })) };
}

if (opt.json) {
  const p = outPath(opt.json);
  fs.writeFileSync(p, JSON.stringify({
    date: new Date().toISOString(),
    index_sha256: crypto.createHash('sha256').update(srcBuf).digest('hex'),
    summary: { hidden: mustOk, must: mustAll, kept: keepOk, keep: keepAll },
    by_type: byType, files: perFile, session: sessionOut,
  }, null, 1) + '\n');
  console.log('\nзамер записан: ' + p + ' (внутри пропущенные строки как есть)');
}
