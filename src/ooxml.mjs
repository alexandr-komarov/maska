// Общее для docx и xlsx: разбор XML на куски без DOM, склейка текста абзаца из кусков,
// поиск строк по склеенному тексту и замена с записью обратно в те же куски.
// Нетронутые куски XML остаются байт в байт как в исходнике.

// Токены: тег, текст, CDATA, прочее (объявление, комментарий). Значения атрибутов могут содержать >
const TOK = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<![^>]*>|<\/?[A-Za-z_][^>"']*(?:(?:"[^"]*"|'[^']*')[^>"']*)*>|[^<]+/g;
export function tokenize(s) {
  const out = [];
  for (const m of s.matchAll(TOK)) {
    const raw = m[0];
    if (raw[0] !== '<') out.push({ raw, kind: 'text' });
    else if (raw.startsWith('<![CDATA[')) out.push({ raw, kind: 'cdata' });
    else if (raw[1] === '!' || raw[1] === '?') out.push({ raw, kind: 'other' });
    else {
      const close = raw[1] === '/', name = raw.match(/^<\/?([^\s/>]+)/)[1];
      out.push({ raw, kind: 'tag', name, close, self: !close && raw.endsWith('/>') });
    }
  }
  if (out.reduce((a, t) => a + t.raw.length, 0) !== s.length) throw new Error('XML не разобран целиком');
  return out;
}
export const join = toks => toks.map(t => t.raw).join('');

const ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
export const decode = s => s.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi, (m, e) =>
  e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : +e.slice(1)) : ENT[e.toLowerCase()] ?? m);
export const encText = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
export const encAttr = s => encText(s).replace(/"/g, '&quot;');
const tokVal = t => t.kind === 'cdata' ? t.raw.slice(9, -3) : decode(t.raw);

export function getAttr(raw, name) {
  const m = raw.match(new RegExp(`\\s${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`));
  return m ? decode(m[1] ?? m[2]) : null;
}
export function setAttr(raw, name, value) {
  const re = new RegExp(`(\\s${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*)(?:"[^"]*"|'[^']*')`);
  if (re.test(raw)) return raw.replace(re, (m, a) => `${a}"${encAttr(value)}"`);
  return raw.replace(/\s*(\/?)>$/, (m, sl) => ` ${name}="${encAttr(value)}"${sl}>`);
}
// все атрибуты тега: [[имя, значение], ...]
export const attrs = raw => [...raw.matchAll(/\s([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)].map(m => [m[1], decode(m[2] ?? m[3])]);

// Префикс пространства имён по корневому элементу: 'w:', '' или null, если не объявлено
export function prefixFor(toks, uris) {
  const root = toks.find(t => t.kind === 'tag');
  if (!root) return null;
  for (const [k, v] of attrs(root.raw)) if (uris.includes(v)) { if (k === 'xmlns') return ''; if (k.startsWith('xmlns:')) return k.slice(6) + ':'; }
  return null;
}

// Сравнение: неразрывные и узкие пробелы равны обычному, мягкий перенос не считается
export const normChar = c => c === ' ' || c === ' ' || c === ' ' ? ' ' : c;
export const norm = s => [...s].filter(c => c !== '­').map(normChar).join('');

// Список замен: пары [строка, метка] или {from,to} или {value,tag}; длинные идут первыми
export function prepare(replacements) {
  const list = replacements.map((r, i) => {
    const [from, to] = Array.isArray(r) ? r : [r.from ?? r.value, r.to ?? r.tag];
    return { i, from: String(from), to: String(to), n: norm(String(from)), count: 0 };
  }).filter(x => x.n.length > 0);
  return list.sort((a, b) => b.n.length - a.n.length || a.i - b.i);
}

// Склейка кусков в строку с картой: позиция в строке → [номер куска, смещение в куске]
function build(segs) {
  let s = ''; const map = [];
  segs.forEach((g, k) => {
    if (g.dead) return;
    if (g.el) { if (g.ch) { s += g.ch; map.push([k, 0]); } return; }
    for (let j = 0; j < g.val.length; j++) { const c = g.val[j]; if (c === '­') continue; s += normChar(c); map.push([k, j]); }
  });
  return { s, map };
}
export const segText = segs => build(segs).s;

// Вырезка куска [s0,o0]..[s1,o1) в склейке и вставка метки. Метка встаёт в первый текстовый кусок,
// остальные куски пустеют, элементы (табуляция, перенос) внутри удаляются. false, если метку некуда поставить
function splice(segs, s0, o0, s1, o1, to) {
  let target = -1;
  for (let k = s0; k <= s1; k++) if (!segs[k].el && !segs[k].dead) { target = k; break; }
  if (target < 0) return false; // совпадение из одних табуляций и переносов
  for (let k = s0; k <= s1; k++) {
    const g = segs[k];
    if (g.dead) continue;
    if (g.el) { g.dead = true; continue; }
    const pre = k === s0 ? g.val.slice(0, o0) : '', post = k === s1 ? g.val.slice(o1) : '';
    g.val = pre + (k === target ? to : '') + post;
  }
  return true;
}

// Замена по кускам: каждое вхождение строки из списка. Возвращает число замен.
export function replaceSegs(segs, needles, onHit) {
  let total = 0;
  for (const nd of needles) {
    const { s, map } = build(segs);
    const occ = [];
    for (let a = s.indexOf(nd.n); a >= 0; a = s.indexOf(nd.n, a + nd.n.length)) occ.push(a);
    for (const a of occ.reverse()) {
      const [s0, o0] = map[a], [s1, l1] = map[a + nd.n.length - 1];
      if (!splice(segs, s0, o0, s1, segs[s1].el ? 1 : l1 + 1, nd.to)) continue;
      nd.count++; total++; onHit?.(nd);
    }
  }
  return total;
}

// Правки по месту: [{view, a, b, to}], a и b - позиции в склейке view (по умолчанию все куски).
// Так меняется ровно то место, где маскировщик нашёл значение, даже если значение заходит в соседнюю ячейку.
// Позиции считаются по кускам до правок, правки идут с конца, пересекающиеся пропускаются.
export function editSegs(segs, edits) {
  const idx = new Map(segs.map((g, k) => [g, k])), maps = new Map(), list = [];
  for (const ed of edits) {
    const view = ed.view || segs;
    if (!maps.has(view)) maps.set(view, build(view).map);
    const map = maps.get(view);
    if (!(ed.b > ed.a) || ed.a < 0 || ed.b > map.length) continue;
    const [k0, o0] = map[ed.a], [k1, l1] = map[ed.b - 1];
    list.push({ view, k0, o0, k1, o1: view[k1].el ? 1 : l1 + 1, p0: idx.get(view[k0]), p1: idx.get(view[k1]), to: ed.to });
  }
  list.sort((x, y) => y.p0 - x.p0 || y.o0 - x.o0);
  let n = 0, lo = null;
  for (const e of list) {
    if (lo && (e.p1 > lo[0] || (e.p1 === lo[0] && e.o1 > lo[1]))) continue;
    if (splice(e.view, e.k0, e.o0, e.k1, e.o1, e.to)) { n++; lo = [e.p0, e.o0]; }
  }
  return n;
}

// Замена в одной строке (атрибут, текст поля, ячейка без разметки)
export function replaceStr(str, needles, onHit) {
  const segs = [{ val: str }];
  const n = replaceSegs(segs, needles, onHit);
  return { val: segs[0].val, n };
}

// Сбор абзацев. cfg: para (имена абзацев), text (текстовые элементы), del (удалённый текст),
// els (элементы-символы внутри прогона: табуляция и т. п.), run, ins/delWrap (обёртки исправлений), fallback
export function collectParas(toks, cfg) {
  const paras = [], stack = [];
  let ins = 0, del = 0, fb = 0;
  // глубина прогона считается внутри своего абзаца: у надписи внутри прогона свои w:tabs в свойствах
  const run = () => stack.at(-1)?.run ?? 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.kind !== 'tag') continue;
    const n = t.name;
    if (cfg.para.has(n)) {
      if (t.self) continue;
      if (!t.close) stack.push({ start: i, segs: [], fb: fb > 0, run: 0 });
      else if (stack.length) paras.push(stack.pop());
      continue;
    }
    if (t.self) {
      if (cfg.els?.has(n) && run() > 0 && stack.length) stack.at(-1).segs.push({ el: true, ch: cfg.els.get(n), open: i, close: i, ins: ins > 0, del: del > 0 });
      continue;
    }
    const d = t.close ? -1 : 1;
    if (cfg.insWrap?.has(n)) { ins += d; continue; }
    if (cfg.delWrap?.has(n)) { del += d; continue; }
    if (cfg.fallback && n === cfg.fallback) { fb += d; continue; }
    if (cfg.run && n === cfg.run) { if (stack.length) stack.at(-1).run += d; continue; }
    if (t.close) continue;
    const isText = cfg.text.has(n), isDel = cfg.del?.has(n), isEl = cfg.els?.has(n) && run() > 0;
    if (!isText && !isDel && !isEl) continue;
    let j = i + 1; const parts = [];
    while (j < toks.length && !(toks[j].kind === 'tag' && toks[j].close && toks[j].name === n)) { if (toks[j].kind === 'text' || toks[j].kind === 'cdata') parts.push(j); j++; }
    if (stack.length) {
      if (isEl) stack.at(-1).segs.push({ el: true, ch: cfg.els.get(n), open: i, close: j, ins: ins > 0, del: del > 0 });
      else if (parts.length) {
        const val = parts.map(k => tokVal(toks[k])).join('');
        stack.at(-1).segs.push({ open: i, tok: parts[0], extra: parts.slice(1), val, orig: val, ins: ins > 0, del: del > 0 || isDel, space: cfg.space });
      }
    }
    i = j;
  }
  return paras.sort((a, b) => a.start - b.start);
}

// Запись изменённых кусков обратно в токены
export function writeSegs(toks, segs) {
  for (const g of segs) {
    if (g.el) { if (g.dead) for (let k = g.open; k <= g.close; k++) toks[k].raw = ''; continue; }
    if (g.val === g.orig) continue;
    toks[g.tok] = { raw: encText(g.val), kind: 'text' };
    for (const k of g.extra) toks[k].raw = '';
    if (g.space && /^\s|\s$|\s\s/.test(g.val) && getAttr(toks[g.open].raw, 'xml:space') !== 'preserve') toks[g.open].raw = setAttr(toks[g.open].raw, 'xml:space', 'preserve');
  }
}

// Замена во всех текстовых узлах части без учёта абзацев (свойства, произвольный XML).
// filter(tagName) решает, какие элементы трогать; edits: номер токена → правки по месту. Возвращает число замен
export function replaceAllText(toks, needles, filter = () => true, edits) {
  let n = 0; const open = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.kind === 'tag') { if (t.close) open.pop(); else if (!t.self) open.push(t.name); continue; }
    if (t.kind !== 'text' && t.kind !== 'cdata') continue;
    if (!filter(open.at(-1), open)) continue;
    const val = tokVal(t), segs = [{ val }];
    const k = (edits?.has(i) ? editSegs(segs, edits.get(i)) : 0) + replaceSegs(segs, needles);
    if (k) { toks[i] = { raw: encText(segs[0].val), kind: 'text' }; n += k; }
  }
  return n;
}
// Текст токена так, как его видит склейка: для правок по месту
export const tokText = t => segText([{ val: tokVal(t) }]);

// Замена в значениях атрибутов; pick(тег, имя атрибута) решает, какие трогать
export function replaceAttrs(toks, needles, pick) {
  let n = 0;
  for (const t of toks) {
    if (t.kind !== 'tag' || t.close) continue;
    for (const [k, v] of attrs(t.raw)) {
      if (!pick(t, k)) continue;
      const r = replaceStr(v, needles);
      if (r.n) { t.raw = setAttr(t.raw, k, r.val); n += r.n; }
    }
  }
  return n;
}

// Обезличивание авторов: одно имя везде получает один номер
export function authorMap() {
  const m = new Map();
  return v => { if (!v) return v; if (!m.has(v)) m.set(v, m.size + 1); return `Автор ${m.get(v)}`; };
}

// docProps: автор, последний редактор, организация, руководитель очищаются, в прочих полях замены по списку
const CLEAR = new Set(['dc:creator', 'cp:lastModifiedBy', 'Company', 'Manager', 'HyperlinkBase']);
export function processProps(name, xml, needles, report, edits) {
  const toks = tokenize(xml);
  let n = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.kind !== 'tag' || t.close || t.self) continue;
    const local = t.name;
    if (CLEAR.has(local) && toks[i + 1]?.kind === 'text') { toks[i + 1] = { raw: '', kind: 'text' }; report.cleared.push(`${name}: ${local}`); }
  }
  n += replaceAllText(toks, needles, undefined, edits);
  if (n) report.replaced[name] = (report.replaced[name] || 0) + n;
  return join(toks);
}

// Текстовые свойства файла для чтения: название, тема, ключевые слова, описание, категория, свои свойства.
// Автор, редактор, организация и руководитель при возврате очищаются, в текст их не берём
const PROP_TEXT = /^(dc:title|dc:subject|cp:keywords|dc:description|cp:category|cp:contentStatus|vt:lpwstr|vt:lpstr|vt:bstr)$/;
export function propsText(xml) {
  const toks = tokenize(xml), out = [], open = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.kind === 'tag') { if (t.close) open.pop(); else if (!t.self) open.push(t.name); continue; }
    if ((t.kind === 'text' || t.kind === 'cdata') && PROP_TEXT.test(open.at(-1) || '')) { const s = tokText(t); if (s.trim()) out.push({ i, text: s }); }
  }
  return out;
}

// Сборка текста для маскировщика с картой мест. Строка состоит из кусков: простой текст (заголовки, разделители)
// или место {key, text} (абзац, ячейка, свойство). По ключу места замена найдёт его в файле.
// Возврат каретки заменяем переводом строки той же длины: поле ввода в браузере всё равно его так заменит
export function textBuilder() {
  const lines = [], units = []; let len = 0;
  return {
    line(...parts) {
      let s = '';
      for (const p of parts) {
        if (p == null) continue;
        if (typeof p === 'string') { s += p.replace(/\r/g, '\n'); continue; }
        const t = p.text.replace(/\r/g, '\n');
        units.push({ s: len + s.length, e: len + s.length + t.length, key: p.key, num: !!p.num });
        s += t;
      }
      lines.push(s); len += s.length + 1;
    },
    get empty() { return lines.length === 0; },
    done() { return { text: lines.join('\n'), units }; },
  };
}

// Итоговый осмотр результата: остались ли строки из списка в каком-либо XML (текст и атрибуты).
// Текст части склеивается без разделителей, поэтому возможны ложные срабатывания на стыке ячеек
export function scanLeft(name, xml, needles) {
  const toks = tokenize(xml);
  let text = '';
  const vals = [];
  for (const t of toks) {
    if (t.kind === 'text' || t.kind === 'cdata') text += tokVal(t);
    else if (t.kind === 'tag' && !t.close) for (const [k, v] of attrs(t.raw)) if (!k.startsWith('xmlns')) vals.push(v);
  }
  const hay = [norm(text), ...vals.map(norm)];
  const out = [];
  for (const nd of needles) {
    let c = 0;
    for (const h of hay) for (let a = h.indexOf(nd.n); a >= 0; a = h.indexOf(nd.n, a + 1)) c++;
    if (c) out.push({ part: name, from: nd.from, count: c });
  }
  return out;
}

export function newReport() {
  return { total: 0, replaced: {}, byValue: [], notFound: [], cleared: [], authors: 0, converted: [], warnings: [], notHandled: [], leftovers: [], changedParts: [] };
}
export function finishReport(report, needles) {
  report.byValue = [...needles].sort((a, b) => a.i - b.i).map(x => ({ from: x.from, to: x.to, count: x.count }));
  report.notFound = report.byValue.filter(x => x.count === 0).map(x => x.from);
  report.total = Object.values(report.replaced).reduce((a, b) => a + b, 0);
  return report;
}
export const isXmlPart = name => /\.(xml|rels|vml)$/i.test(name);
