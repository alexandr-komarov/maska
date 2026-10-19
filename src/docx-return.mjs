// Чтение и возврат docx. Вёрстка и стили не трогаются: меняется только текст внутри w:t,
// пустые куски остаются пустыми w:t, свойства прогонов и абзацев не меняются.
import { readZip, writeZip, byName } from './zip.mjs';
import { tokenize, join, prefixFor, prepare, collectParas, replaceSegs, editSegs, writeSegs, segText, replaceStr, replaceAllText,
  replaceAttrs, authorMap, processProps, propsText, scanLeft, newReport, finishReport, isXmlPart, getAttr, setAttr, textBuilder } from './ooxml.mjs';

const WNS = ['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main'];
const ANS = ['http://schemas.openxmlformats.org/drawingml/2006/main', 'http://purl.oclc.org/ooxml/drawingml/main'];
const td = new TextDecoder('utf-8', { ignoreBOM: true });

// Части с текстом документа: основной, колонтитулы, сноски, примечания, автотекст
const TEXT_PART = /^word\/(document\d*|header\d+|footer\d+|footnotes|endnotes|comments|glossary\/document)\.xml$/;
// Рисунки SmartArt и диаграммы: текст в a:p/a:t
const DML_PART = /^word\/(diagrams\/data\d+|diagrams\/drawing\d+|charts\/chart\d+)\.xml$/;
const PART_NAME = n => n.replace(/^word\//, '').replace(/\.xml$/, '');

function cfgW(p, ext = {}) {
  return {
    para: new Set([p + 'p']), text: new Set([p + 't', 'm:t']), del: new Set([p + 'delText']), run: p + 'r',
    // табуляция и переносы входят в склейку как \t и \n, мягкий перенос как пустой символ
    els: new Map([[p + 'tab', '\t'], [p + 'br', '\n'], [p + 'cr', '\n'], [p + 'noBreakHyphen', '-'], [p + 'softHyphen', '']]),
    insWrap: new Set([p + 'ins', p + 'moveTo']), delWrap: new Set([p + 'del', p + 'moveFrom']),
    fallback: 'mc:Fallback', space: true, ...ext,
  };
}
const cfgA = p => ({ para: new Set([p + 'p']), text: new Set([p + 't']), els: new Map([[p + 'br', '\n'], [p + 'tab', '\t']]), run: null, fallback: 'mc:Fallback', space: false });

// Два взгляда на абзац: текущий (без удалённого) и исходный (без вставленного)
const viewNow = segs => segs.filter(g => !g.del);
const viewOld = segs => segs.filter(g => !g.ins);

function parasOf(name, toks) {
  if (TEXT_PART.test(name)) { const p = prefixFor(toks, WNS); return p == null ? null : { paras: collectParas(toks, cfgW(p)), p }; }
  if (DML_PART.test(name)) { const p = prefixFor(toks, ANS) ?? 'a:'; return { paras: collectParas(toks, { ...cfgA(p), fallback: null }), p, dml: true }; }
  return null;
}

// Подписи частей в тексте
function partTitle(name) {
  const m = name.match(/^word\/(header|footer)(\d+)\.xml$/);
  if (m) return `${m[1] === 'header' ? 'Верхний' : 'Нижний'} колонтитул ${m[2]}`;
  return { 'word/footnotes.xml': 'Сноски', 'word/endnotes.xml': 'Концевые сноски', 'word/comments.xml': 'Примечания', 'word/glossary/document.xml': 'Автотекст' }[name]
    || (/diagrams/.test(name) ? 'Схема SmartArt' : /charts/.test(name) ? 'Диаграмма' : name);
}
// Порядок: основной документ, колонтитулы, сноски, примечания, остальное
const RANK = n => n === 'word/document.xml' ? 0 : /header/.test(n) ? 1 : /footer/.test(n) ? 2 : /footnotes/.test(n) ? 3 : /endnotes/.test(n) ? 4 : /comments/.test(n) ? 5 : 6;
const order = es => [...es].sort((a, b) => RANK(a.name) - RANK(b.name) || a.name.localeCompare(b.name, 'en', { numeric: true }));

// Текст документа той же склейкой, что и замена. Абзацы через \n, части с заголовками.
// Удалённый в исправлениях текст отдельным блоком: абзацы в редакции до правки.
// units: место каждого абзаца в тексте и его ключ для правок по месту
export async function extractDocxText(bytes) {
  const zip = readZip(bytes), z = byName(zip);
  const tb = textBuilder(), old = [];
  const info = { parts: [], revisions: 0, comments: 0, notes: 0, textboxes: 0 };
  for (const e of order(zip.entries)) {
    if (!TEXT_PART.test(e.name) && !DML_PART.test(e.name)) continue;
    const toks = tokenize(td.decode(await e.data()));
    const r = parasOf(e.name, toks);
    if (!r) continue;
    const lines = [];
    for (const pa of r.paras) {
      if (pa.fb) continue; // запасная копия надписи в mc:Fallback повторяет основную
      const now = segText(viewNow(pa.segs)), was = segText(viewOld(pa.segs));
      lines.push({ key: `${e.name}#${pa.start}`, text: now });
      if (was !== now) old.push({ key: `${e.name}#${pa.start}#old`, text: was });
    }
    if (e.name !== 'word/document.xml') {
      // пустые колонтитулы и колонтитулы из одного номера страницы не показываем
      const all = lines.map(l => l.text).join('\n').trim();
      if (!all || (/header|footer/.test(e.name) && /^(стр\.?\s*)?\d*$/i.test(all))) continue;
      tb.line(`=== ${partTitle(e.name)} ===`);
      info.parts.push(partTitle(e.name));
      if (/comments/.test(e.name)) info.comments += lines.filter(l => l.text.trim()).length;
      if (/notes/.test(e.name)) info.notes += lines.filter(l => l.text.trim()).length;
    } else info.textboxes = (toks.filter(t => t.kind === 'tag' && !t.close && /txbxContent$/.test(t.name)).length);
    for (const l of lines) tb.line(l);
  }
  if (old.length) {
    info.revisions = old.length;
    tb.line('=== Исправления: удалённый текст, абзацы в редакции до правки ===');
    for (const l of old) tb.line(l);
  }
  let head = false;
  for (const part of ['docProps/core.xml', 'docProps/custom.xml']) {
    if (!z[part]) continue;
    for (const { i, text } of propsText(await z[part].text())) {
      if (!head) { tb.line('=== Свойства файла ==='); head = true; }
      tb.line({ key: `${part}@${i}`, text });
    }
  }
  info.embeddings = zip.entries.filter(e => /^word\/embeddings\//.test(e.name)).length;
  info.media = zip.entries.filter(e => /^word\/media\//.test(e.name)).length;
  return { ...tb.done(), info };
}

// Правки по месту для части: номер токена начала абзаца → {now, old}
function paraEdits(edits, part) {
  const m = new Map();
  if (edits) for (const [k, v] of edits) {
    if (!k.startsWith(part + '#')) continue;
    const [start, view] = k.slice(part.length + 1).split('#');
    if (!m.has(+start)) m.set(+start, { now: [], old: [] });
    m.get(+start)[view === 'old' ? 'old' : 'now'].push(...v);
  }
  return m;
}
const tokEdits = (edits, part) => {
  const m = new Map();
  if (edits) for (const [k, v] of edits) if (k.startsWith(part + '@')) m.set(+k.slice(part.length + 1), v);
  return m;
};

// Возврат. replacements: пары [строка, метка] и т. п. (см. prepare); edits: Map ключ места → [{a, b, to}], правки по месту
export async function replaceInDocx(bytes, replacements, { edits } = {}) {
  const zip = readZip(bytes);
  const needles = prepare(replacements);
  const report = newReport();
  const author = authorMap();
  const add = (part, n) => { if (n) report.replaced[part] = (report.replaced[part] || 0) + n; };
  let media = 0;

  for (const e of zip.entries) {
    const name = e.name;
    if (!isXmlPart(name)) {
      if (/^word\/embeddings\//.test(name)) report.notHandled.push({ kind: 'embedding', part: name, note: 'встроенный объект, внутрь не заглядываем' });
      if (/^word\/media\//.test(name)) media++;
      continue;
    }
    const src = td.decode(await e.data());
    let out = src;
    if (TEXT_PART.test(name) || DML_PART.test(name)) {
      const toks = tokenize(src);
      const r = parasOf(name, toks);
      if (!r) { report.notHandled.push({ kind: 'part', part: name, note: 'не найдено пространство имён WordprocessingML' }); continue; }
      let n = 0;
      // правки по месту: сначала по абзацам из текста, потом по их запасным копиям в mc:Fallback с тем же текстом
      const pe = paraEdits(edits, name), byText = new Map();
      for (const pa of r.paras) {
        const ed = pe.get(pa.start);
        if (ed && !pa.fb) byText.set(segText(viewNow(pa.segs)), ed);
      }
      for (const pa of r.paras) {
        const ed = pa.fb ? byText.get(segText(viewNow(pa.segs))) : pe.get(pa.start);
        if (ed) {
          const now = viewNow(pa.segs), old = viewOld(pa.segs);
          n += editSegs(pa.segs, [...ed.now.map(x => ({ ...x, view: now })), ...ed.old.map(x => ({ ...x, view: old }))]);
        }
        n += replaceSegs(viewNow(pa.segs), needles);
        // удалённый текст тоже заменяем: он лежит в файле и виден в режиме исправлений
        n += replaceSegs(viewOld(pa.segs), needles);
        writeSegs(toks, pa.segs);
      }
      add(name, n);
      if (!r.dml) {
        const p = r.p;
        // коды полей (HYPERLINK "mailto:...", MERGEFIELD) и простые поля
        add(name + ' (поля)', replaceAllText(toks, needles, tag => tag === p + 'instrText' || tag === p + 'delInstrText'));
        add(name + ' (поля)', replaceAttrs(toks, needles, (t, k) => k === p + 'instr'));
        // авторы исправлений и примечаний
        for (const t of toks) {
          if (t.kind !== 'tag' || t.close) continue;
          const a = getAttr(t.raw, p + 'author');
          if (a) { t.raw = setAttr(t.raw, p + 'author', author(a)); report.authors++; }
          const ini = getAttr(t.raw, p + 'initials');
          if (ini) t.raw = setAttr(t.raw, p + 'initials', 'А');
        }
      } else {
        // кэш строк диаграммы
        add(name, replaceAllText(toks, needles, (tag, open) => tag === 'c:v' && open.includes('c:strCache')));
      }
      out = join(toks);
    } else if (name === 'word/people.xml') {
      const toks = tokenize(src);
      for (const t of toks) {
        if (t.kind !== 'tag' || t.close) continue;
        for (const k of ['w15:author', 'w15:userId']) { const a = getAttr(t.raw, k); if (a) { t.raw = setAttr(t.raw, k, author(a)); report.authors++; } }
      }
      out = join(toks);
    } else if (/^word\/_rels\/.*\.rels$/.test(name)) {
      // адреса внешних ссылок: mailto, сайты с именами
      const toks = tokenize(src);
      add(name, replaceAttrs(toks, needles, (t, k) => k === 'Target' && getAttr(t.raw, 'TargetMode') === 'External'));
      out = join(toks);
    } else if (/^customXml\/item\d+\.xml$/.test(name)) {
      // данные, привязанные к полям формы: Word при открытии может вернуть их в документ
      const toks = tokenize(src);
      add(name, replaceAllText(toks, needles));
      out = join(toks);
    } else if (/^docProps\/(core|app|custom)\.xml$/.test(name)) {
      out = processProps(name, src, needles, report, tokEdits(edits, name));
    }
    if (out !== src) { e.setData(out); report.changedParts.push(name); }
  }

  // осмотр результата: что из списка осталось в XML
  for (const e of zip.entries) if (isXmlPart(e.name)) report.leftovers.push(...scanLeft(e.name, td.decode(await e.data()), needles));
  if (media) report.notHandled.push({ kind: 'media', count: media, note: 'картинки не проверяются' });
  finishReport(report, needles);
  return { bytes: await writeZip(zip), report };
}
