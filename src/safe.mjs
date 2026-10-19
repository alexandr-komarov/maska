// Сейф: метки и настоящие значения, общие для документов одного клиента или дела.
// Только чистые функции: сейф живёт в памяти страницы и в файле у пользователя, в хранилища браузера не пишется.
// Сейф в памяти: { name, created, changed, words: [слово], skip: [слово], records: [{ tag, type, value, forms: [написание], keys: [ключ] }] }
// skip: сомнительные слова и числа, с которых человек снял метку в окне проверки. В следующих документах они не прячутся

// Слово метки по типу находки. Единственная таблица соответствия: метка [ОРГ_1], [ФИО_2], [ДОГОВОР_3]
export const TAG_WORD = {
  PERSON: 'ФИО', PHONE: 'ТЕЛ', EMAIL: 'ПОЧТА', INN: 'ИНН', KPP: 'КПП', OGRN: 'ОГРН', BIK: 'БИК', SNILS: 'СНИЛС',
  PASSPORT: 'ПАСПОРТ', DOB: 'РОЖДЕНИЕ', ACCOUNT: 'СЧЁТ', CARD: 'КАРТА', ADDRESS: 'АДРЕС', ORG: 'ОРГ', TERM: 'СЛОВО',
  ID: 'ДОГОВОР', NICK: 'НИК', DATE: 'ДАТА', AMOUNT: 'СУММА',
  // сомнительное из окна проверки: слово с заглавной вне словаря, длинное число, всё со знаком @
  NAME: 'НАЗВАНИЕ', NUM: 'НОМЕР', AT: 'КОНТАКТ',
};
const yo = s => s.toUpperCase().replace(/Ё/g, 'Е');
const WORD_TYPE = new Map(Object.entries(TAG_WORD).map(([t, w]) => [yo(w), t]));
export const tagOf = (type, n) => `[${TAG_WORD[type]}_${n}]`;
// Любая метка в тексте с метками: новая [ОРГ_1] и старая #ORG_1# первой версии Маски
export const TAG_ANY = new RegExp(`\\[(?:${Object.values(TAG_WORD).join('|')})_\\d+\\]|#(?:${Object.keys(TAG_WORD).join('|')})_\\d+#`, 'g');

// Метка → {type, n, old}. Старые метки #ORG_1# понимаются, новые не выдаются
export function parseTag(tag) {
  let m = /^\[([А-ЯЁA-Z]+)_(\d+)\]$/.exec(tag);
  if (m && WORD_TYPE.has(yo(m[1]))) return { type: WORD_TYPE.get(yo(m[1])), n: +m[2], old: false };
  m = /^#([A-Z]+)_(\d+)#$/.exec(tag);
  if (m && TAG_WORD[m[1]]) return { type: m[1], n: +m[2], old: true };
  return null;
}

const normKey = s => String(s).toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ').trim();
const normForm = s => normKey(String(s).replace(/[«»"“”„]/g, ''));

export const FORMAT = 'maska-safe';
export const VERSION = 1;

export function safeNew(name = '', now = new Date()) {
  const d = now.toISOString();
  return { name, created: d, changed: d, words: [], skip: [], records: [] };
}

// Файл сейфа. Записи и слова лежат в data: при шифровании его заменит поле encrypted, имя и даты останутся открытыми
export function safeSerialize(safe) {
  return JSON.stringify({
    format: FORMAT, version: VERSION, name: safe.name || '', created: safe.created, changed: safe.changed,
    data: { words: safe.words, skip: safe.skip || [], records: safe.records.map(r => ({ tag: r.tag, type: r.type, value: r.value, forms: r.forms, keys: r.keys })) },
  }, null, 1) + '\n';
}

export const safeFileName = safe => `${String(safe.name || '').replace(/[\\/:*?"<>|\x00-\x1f]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80) || 'без названия'}.сейф.json`;

// Чтение файла сейфа: {safe} или {error} с фразой для человека. Старый сейф первой версии (список меток) тоже читается
export function safeParse(text) {
  const fail = error => ({ error });
  let j;
  try { j = JSON.parse(String(text).replace(/^﻿/, '')); }
  catch { return fail('файл повреждён или это не сейф Маски: его не удалось прочитать'); }
  if (Array.isArray(j)) return oldSafe(j);
  if (!j || typeof j !== 'object' || j.format !== FORMAT) return fail('это не сейф Маски, а другой файл');
  if (typeof j.version !== 'number' || j.version > VERSION) return fail('сейф сохранён более новой версией Маски, откройте его в ней');
  if (j.encrypted) return fail('сейф защищён паролем, эта версия Маски пароли не открывает');
  const d = j.data;
  if (!d || !Array.isArray(d.records) || (d.words != null && !Array.isArray(d.words)) || (d.skip != null && !Array.isArray(d.skip))) return fail('файл сейфа повреждён: нет списка записей');
  const records = [], tags = new Set();
  for (const [i, r] of d.records.entries()) {
    const p = r && typeof r.tag === 'string' && parseTag(r.tag);
    if (!p || typeof r.value !== 'string' || !r.value) return fail(`файл сейфа повреждён: в записи ${i + 1} нет метки или значения`);
    if (r.type != null && r.type !== p.type) return fail(`файл сейфа повреждён: у метки ${r.tag} другой тип`);
    if (tags.has(r.tag)) return fail(`файл сейфа повреждён: метка ${r.tag} записана дважды`);
    tags.add(r.tag);
    const strs = a => Array.isArray(a) ? a.filter(x => typeof x === 'string' && x) : [];
    const forms = strs(r.forms);
    records.push({ tag: r.tag, type: p.type, value: r.value, forms: forms.length ? forms : [r.value], keys: strs(r.keys) });
  }
  const now = new Date().toISOString();
  return { safe: { name: typeof j.name === 'string' ? j.name : '', created: j.created || now, changed: j.changed || now,
    words: (d.words || []).filter(w => typeof w === 'string' && w.trim()), skip: (d.skip || []).filter(w => typeof w === 'string' && w.trim()), records } };
}

// Сейф первой версии: [{tag: '#ORG_1#', type: 'Организация', value, forms}]
function oldSafe(list) {
  const records = [], tags = new Set();
  for (const r of list) {
    const p = r && typeof r.tag === 'string' && parseTag(r.tag);
    if (!p || typeof r.value !== 'string' || tags.has(r.tag)) return { error: 'это не сейф Маски, а другой файл' };
    tags.add(r.tag);
    const forms = Array.isArray(r.forms) ? r.forms.filter(x => typeof x === 'string' && x) : [];
    records.push({ tag: r.tag, type: p.type, value: r.value, forms: forms.length ? forms : [r.value], keys: [] });
  }
  const now = new Date().toISOString();
  return { safe: { name: '', created: now, changed: now, words: [], skip: [], records }, old: true };
}

// ФИО для обратной подстановки: полное лучше инициалов, именительный («Петров Иван Сергеевич») лучше косвенного
const nominative = v => /(?:ич|на)$/.test(v);
const words3 = v => v.trim().split(/\s+/).length >= 3;
function betterPerson(cur, v) {
  if (!words3(v)) return false;
  if (!words3(cur)) return true;
  return !nominative(cur) && nominative(v);
}

// Ключ человека «фамилия|первая буква имени», у фамилии без имени «фамилия|?». Одинокая фамилия идёт к единственному
// человеку с этой фамилией в сейфе, полное ФИО к записи, где фамилия была только одна
function byStem(records, key) {
  const i = key.indexOf('|');
  if (i <= 0) return null;
  const st = normKey(key.slice(0, i)), lone = key.slice(i + 1) === '?';
  if (st === 'np' || st === 'form') return null;
  const same = records.filter(r => r.type === 'PERSON' && r.keys.some(k => normKey(k).startsWith(st + '|')));
  const fit = lone ? same : same.filter(r => r.keys.every(k => k.endsWith('|?')));
  return fit.length === 1 ? fit[0] : null;
}

// Новый документ с сейфом. entities: сущности документа {t, key, value, forms, safeTag?} в порядке появления.
// Сущность получает метку записи сейфа, если совпал ключ того же типа, найденная по сейфу метка (safeTag)
// или одно из написаний; иначе следующий свободный номер своего типа. Возвращает новый сейф и метки по порядку
export function safeMerge(safe, entities, { now = new Date() } = {}) {
  const records = safe.records.map(r => ({ ...r, forms: [...r.forms], keys: [...r.keys] }));
  const byTag = new Map(), byKey = new Map(), byForm = new Map(), next = {};
  const index = r => {
    byTag.set(r.tag, r);
    for (const k of r.keys) if (!byKey.has(r.type + '\u0001' + normKey(k))) byKey.set(r.type + '\u0001' + normKey(k), r);
    for (const f of r.forms) if (!byForm.has(r.type + '\u0001' + normForm(f))) byForm.set(r.type + '\u0001' + normForm(f), r);
  };
  for (const r of records) { index(r); const p = parseTag(r.tag); next[p.type] = Math.max(next[p.type] || 0, p.n); }
  let added = 0, changed = false;
  const tags = [], fresh = [];
  for (const e of entities) {
    const forms = [...(e.forms || [e.value])];
    // одинокая фамилия «петров|?» по ключу не ищется: её запись выбирает byStem, только если человек с такой фамилией один
    const lone = e.t === 'PERSON' && /\|\?$/.test(e.key || '');
    let r = (e.safeTag && byTag.get(e.safeTag)) || (e.key != null && !lone && byKey.get(e.t + '\u0001' + normKey(e.key)));
    if (!r) for (const f of forms) if ((r = byForm.get(e.t + '\u0001' + normForm(f)))) break;
    if (!r && e.t === 'PERSON' && e.key) r = byStem(records, e.key);
    const isNew = !r;
    if (isNew) {
      const n = next[e.t] = (next[e.t] || 0) + 1;
      r = { tag: tagOf(e.t, n), type: e.t, value: e.value, forms: [], keys: [] };
      records.push(r); added++; changed = true;
    }
    if (e.key != null && !e.safeTag && (!lone || isNew) && !r.keys.some(k => normKey(k) === normKey(e.key))) { r.keys.push(e.key); changed = true; }
    for (const f of forms) if (!r.forms.includes(f)) { r.forms.push(f); changed = true; }
    if (r.type === 'PERSON') for (const f of [e.value, ...forms]) if (betterPerson(r.value, f)) { r.value = f; changed = true; }
    index(r);
    tags.push(r.tag); fresh.push(isNew);
  }
  return { safe: { ...safe, records, changed: changed ? now.toISOString() : safe.changed }, tags, fresh, added };
}

// Свои слова: поле «Что ещё спрятать» и словарь сейфа одно и то же
export function safeSetWords(safe, words, { now = new Date() } = {}) {
  const out = [], seen = new Set();
  for (const w of words) { const s = String(w).trim(), k = normKey(s); if (s.length > 1 && !seen.has(k)) { seen.add(k); out.push(s); } }
  if (out.join('\n') === safe.words.join('\n')) return safe;
  return { ...safe, words: out, changed: now.toISOString() };
}

// Слова, с которых снята метка в окне проверки. Список заменяется целиком: слово можно и вернуть под метку
export function safeSetSkip(safe, list, { now = new Date() } = {}) {
  const out = [], seen = new Set();
  for (const w of list) { const s = String(w).trim(), k = normKey(s); if (s && !seen.has(k)) { seen.add(k); out.push(s); } }
  if (out.join('\n') === (safe.skip || []).join('\n')) return safe;
  return { ...safe, skip: out, changed: now.toISOString() };
}

// Написания из сейфа, которые ищутся в новом документе как есть. Свои слова ищет маска сама, даты не ищутся
// (дата выдачи паспорта совпадёт с датой договора), суммы только при включённом «прятать суммы».
// Короткие значения и числа меньше 6 цифр не ищутся: «12» сидит в «2012»
export function safeKnownForms(safe, { sums = true } = {}) {
  const out = [];
  for (const r of safe.records) {
    if (r.type === 'TERM' || r.type === 'DATE' || (r.type === 'AMOUNT' && !sums)) continue;
    for (const f of r.forms) {
      const v = f.trim();
      if (v.length < 4 || /[\t\n]/.test(f) || (/^[\d\s.,\/№#()+-]+$/.test(v) && v.replace(/\D/g, '').length < 6)) continue;
      out.push({ form: f, t: r.type, tag: r.tag });
    }
  }
  return out.sort((a, b) => b.form.length - a.form.length);
}

export function safeSummary(safe) {
  const byType = {};
  for (const r of safe.records) byType[r.type] = (byType[r.type] || 0) + 1;
  return { records: safe.records.length, words: safe.words.length, byType };
}
// Содержимое без дат: изменилось ли что-то с последнего сохранения
export const safeDigest = safe => JSON.stringify([safe.name || '', safe.words, safe.records, safe.skip || []]);

// Обратная подстановка. Нейросеть искажает метки: [орг_1], [ОРГ 1], ОРГ_1, \[ОРГ\_1\], [СЧЕТ_1] вместо [СЧЁТ_1].
// Без скобок метка узнаётся, если в ней подчёркивание или слово заглавными; через пробел без скобок только известная сейфу.
// Метка с окончанием («[ОРГ_1]а») не трогается: просклонять название автомат не умеет, такие места в списке declined.
// Возвращает текст, куски для подсветки и списки неузнанного
const ALT = Object.values(TAG_WORD).sort((a, b) => b.length - a.length).map(w => w.replace(/[ЁЕ]/g, '[ЁЕ]')).join('|');
const RU = new RegExp(`(?<![\\p{L}\\p{N}_])(\\\\?\\[[ \\t]*)?(${ALT})([ \\t]*(?:\\\\?_|-)[ \\t]*|[ \\t]+)?(\\d{1,4})(?!\\p{N})([ \\t]*\\\\?\\])?`, 'giu');
const OLD = new RegExp(`(?<![A-Za-z0-9_])(#)?(${Object.keys(TAG_WORD).join('|')})_(\\d{1,4})(?![0-9])(#)?`, 'g');

export function safeRestore(text, safe) {
  const byTag = new Map(safe.records.map(r => [r.tag, r]));
  const found = [];
  for (const m of text.matchAll(RU)) {
    const [raw, open, word, sep = '', num, close] = m;
    const tag = tagOf(WORD_TYPE.get(yo(word)), +num), known = byTag.has(tag);
    const bracket = !!(open || close), upper = word === word.toUpperCase(), under = sep.includes('_');
    if (!bracket && !under && !(upper && sep && !/^\s+$/.test(sep)) && !(upper && known)) continue;
    found.push({ s: m.index, e: m.index + raw.length, raw, tag, known, sure: bracket || under });
  }
  for (const m of text.matchAll(OLD)) {
    const tag = `#${m[2]}_${m[3]}#`, known = byTag.has(tag);
    if (!known && !(m[1] && m[4])) continue;
    found.push({ s: m.index, e: m.index + m[0].length, raw: m[0], tag, known, sure: true });
  }
  found.sort((a, b) => a.s - b.s || b.e - a.e);
  const parts = [], unknown = [], declined = [];
  let pos = 0, out = '', n = 0;
  for (const f of found) {
    if (f.s < pos) continue;
    if (f.s > pos) { parts.push({ text: text.slice(pos, f.s) }); out += text.slice(pos, f.s); }
    const r = byTag.get(f.tag);
    if (r && /^\p{L}/u.test(text.slice(f.e))) { parts.push({ text: f.raw, declined: f.tag }); declined.push(f.raw); out += f.raw; }
    else if (r) { parts.push({ text: r.value, tag: r.tag, type: r.type }); out += r.value; n++; }
    else if (f.sure) { parts.push({ text: f.raw, unknown: f.tag }); unknown.push(f.raw); out += f.raw; }
    else { parts.push({ text: f.raw }); out += f.raw; }
    pos = f.e;
  }
  if (pos < text.length) { parts.push({ text: text.slice(pos) }); out += text.slice(pos); }
  return { text: out, parts, replaced: n, unknown: [...new Set(unknown)], declined: [...new Set(declined)] };
}
