// Сейф без страницы: метки, файл, пополнение, обратная подстановка. Чистые функции из src/safe.mjs
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { TAG_WORD, tagOf, parseTag, TAG_ANY, safeNew, safeSerialize, safeParse, safeFileName, safeMerge, safeSetWords,
  safeKnownForms, safeSummary, safeDigest, safeRestore, FORMAT, VERSION } from '../src/safe.mjs';

const T0 = new Date('2026-10-04T10:00:00Z'), T1 = new Date('2026-10-05T10:00:00Z');
const ent = (t, key, value, forms = [value]) => ({ t, key, value, forms });

test('метки: русское слово заглавными для каждого типа, разбор новых и старых', () => {
  const html = fs.readFileSync(new URL('../index.html', import.meta.url), 'utf8');
  const label = Object.keys(JSON.parse('{' + html.match(/const LABEL=\{([^}]*)\}/)[1].replace(/(\w+):'/g, '"$1":"').replace(/'/g, '"') + '}'));
  assert.deepEqual(Object.keys(TAG_WORD).sort(), label.sort(), 'слово метки есть у каждого типа из LABEL');
  for (const w of Object.values(TAG_WORD)) assert.match(w, /^[А-ЯЁ]+$/);
  assert.equal(new Set(Object.values(TAG_WORD)).size, Object.keys(TAG_WORD).length, 'слова не повторяются');
  assert.equal(tagOf('ORG', 1), '[ОРГ_1]');
  assert.equal(tagOf('ID', 12), '[ДОГОВОР_12]');
  assert.equal(tagOf('ACCOUNT', 2), '[СЧЁТ_2]');
  assert.deepEqual(parseTag('[ФИО_3]'), { type: 'PERSON', n: 3, old: false });
  assert.deepEqual(parseTag('[СЧЕТ_2]'), { type: 'ACCOUNT', n: 2, old: false });
  assert.deepEqual(parseTag('#ORG_7#'), { type: 'ORG', n: 7, old: true });
  assert.equal(parseTag('[ПРОЧЕЕ_1]'), null);
  assert.equal(parseTag('#FOO_1#'), null);
  assert.deepEqual('[ОРГ_1] и #PERSON_2#, [ИНН_10]'.match(TAG_ANY), ['[ОРГ_1]', '#PERSON_2#', '[ИНН_10]']);
});

test('файл сейфа: формат с версией, данные отдельно от имени и дат, чтение обратно', () => {
  let s = safeNew('Каменный Мост, подряд', T0);
  s = safeMerge(s, [ent('ORG', 'каменный мост', 'ООО «Каменный Мост»'), ent('INN', '7816135799', '7816135799')], { now: T1 }).safe;
  s = safeSetWords(s, ['Складской комплекс', 'проект Север-2'], { now: T1 });
  const text = safeSerialize(s), j = JSON.parse(text);
  assert.equal(j.format, FORMAT);
  assert.equal(j.version, VERSION);
  assert.equal(j.name, 'Каменный Мост, подряд');
  assert.equal(j.changed, T1.toISOString());
  assert.deepEqual(Object.keys(j), ['format', 'version', 'name', 'created', 'changed', 'data'], 'записи только в data: его можно будет зашифровать целиком');
  assert.deepEqual(j.data.words, ['Складской комплекс', 'проект Север-2']);
  assert.deepEqual(j.data.records[0], { tag: '[ОРГ_1]', type: 'ORG', value: 'ООО «Каменный Мост»', forms: ['ООО «Каменный Мост»'], keys: ['каменный мост'] });
  const back = safeParse(text);
  assert.ok(back.safe, back.error);
  assert.equal(safeDigest(back.safe), safeDigest(s));
  assert.equal(safeFileName(s), 'Каменный Мост, подряд.сейф.json');
  assert.equal(safeFileName(safeNew('')), 'без названия.сейф.json');
  assert.equal(safeFileName(safeNew('ООО "Ромашка" / дело 5')), 'ООО Ромашка дело 5.сейф.json');
  assert.deepEqual(safeSummary(s), { records: 2, words: 2, byType: { ORG: 1, INN: 1 } });
});

test('битый или чужой файл: понятная причина, без исключения', () => {
  const ok = JSON.parse(safeSerialize(safeMerge(safeNew('А'), [ent('ORG', 'а', 'ООО «А»')]).safe));
  const cases = [
    ['{"format":"maska-safe",', /повреждён/],
    ['', /повреждён/],
    ['PK\u0003\u0004 бинарный мусор', /повреждён/],
    ['{"name":"заметки","items":[]}', /не сейф Маски/],
    ['[1,2,3]', /не сейф Маски/],
    ['null', /не сейф Маски/],
    [JSON.stringify({ ...ok, version: 2 }), /более новой версией/],
    [JSON.stringify({ ...ok, data: undefined, encrypted: { alg: 'AES-GCM', data: 'xx' } }), /паролем/],
    [JSON.stringify({ ...ok, data: { records: 'нет' } }), /нет списка записей/],
    [JSON.stringify({ ...ok, data: { records: [{ tag: '[ОРГ_1]' }] } }), /в записи 1 нет метки или значения/],
    [JSON.stringify({ ...ok, data: { records: [{ tag: 'ОРГ1', value: 'x' }] } }), /в записи 1/],
    [JSON.stringify({ ...ok, data: { records: [{ tag: '[ОРГ_1]', type: 'INN', value: 'x' }] } }), /другой тип/],
    [JSON.stringify({ ...ok, data: { records: [{ tag: '[ОРГ_1]', value: 'x' }, { tag: '[ОРГ_1]', value: 'y' }] } }), /дважды/],
  ];
  for (const [text, re] of cases) {
    const r = safeParse(text);
    assert.equal(r.safe, undefined, text);
    assert.match(r.error, re, text);
  }
});

test('сейф первой версии: список меток #ORG_1# читается, новые метки продолжают нумерацию', () => {
  const old = JSON.stringify([
    { tag: '#ORG_1#', type: 'Организация', value: 'ООО «Ромашка»', forms: ['ООО «Ромашка»'] },
    { tag: '#PERSON_2#', type: 'ФИО', value: 'Петров Иван Сергеевич', forms: ['Петров Иван Сергеевич', 'Петров И.С.'] },
  ]);
  const r = safeParse(old);
  assert.ok(r.safe && r.old);
  assert.equal(r.safe.records[1].type, 'PERSON');
  const m = safeMerge(r.safe, [ent('ORG', 'ромашка', 'ООО «Ромашка»'), ent('ORG', 'лютик', 'ООО «Лютик»'), ent('PERSON', 'петров|И', 'Петров И.С.')]);
  assert.deepEqual(m.tags, ['#ORG_1#', '[ОРГ_2]', '#PERSON_2#']);
  assert.equal(safeRestore('Договор #ORG_1# и [ОРГ_2], подписал #PERSON_2#', m.safe).text, 'Договор ООО «Ромашка» и ООО «Лютик», подписал Петров Иван Сергеевич');
});

test('пополнение: известное значение получает прежнюю метку, новое следующий свободный номер своего типа', () => {
  let s = safeNew('А');
  let m = safeMerge(s, [ent('ORG', 'каменный мост', 'ООО «Каменный Мост»'), ent('ORG', 'северный ветер строй', 'ООО «Северный Ветер Строй»'), ent('INN', '7816135799', '7816135799')], { now: T0 });
  assert.deepEqual(m.tags, ['[ОРГ_1]', '[ОРГ_2]', '[ИНН_1]']);
  assert.deepEqual(m.fresh, [true, true, true]);
  assert.equal(m.added, 3);
  s = m.safe;
  m = safeMerge(s, [ent('ORG', 'нева-металл', 'АО «Нева-Металл»'), ent('ORG', 'каменный мост', 'Каменный Мост ООО'), ent('PERSON', 'петров|И', 'Петров И.С.')], { now: T1 });
  assert.deepEqual(m.tags, ['[ОРГ_3]', '[ОРГ_1]', '[ФИО_1]']);
  assert.deepEqual(m.fresh, [true, false, true]);
  assert.deepEqual(m.safe.records[0].forms, ['ООО «Каменный Мост»', 'Каменный Мост ООО'], 'новое написание добавлено к записи');
  assert.equal(m.safe.changed, T1.toISOString());
  // документ без нового: сейф не меняется, дата тоже
  const again = safeMerge(m.safe, [ent('ORG', 'каменный мост', 'ООО «Каменный Мост»')], { now: new Date('2027-01-01') });
  assert.equal(again.added, 0);
  assert.equal(again.safe.changed, T1.toISOString());
  // номер после пропуска: свободный номер это максимум плюс один
  const gap = safeParse(JSON.stringify({ format: FORMAT, version: 1, data: { records: [{ tag: '[ОРГ_5]', value: 'ООО «Пять»' }] } })).safe;
  assert.deepEqual(safeMerge(gap, [ent('ORG', 'шесть', 'ООО «Шесть»')]).tags, ['[ОРГ_6]']);
});

test('пополнение: совпадение по написанию без ключа, ё и е, кавычки', () => {
  const s = safeParse(JSON.stringify({ format: FORMAT, version: 1, data: { records: [{ tag: '[ОРГ_1]', type: 'ORG', value: 'ООО «Ёлка»', forms: ['ООО «Ёлка»'] }] } })).safe;
  assert.deepEqual(safeMerge(s, [ent('ORG', 'x', 'ООО "Елка"')]).tags, ['[ОРГ_1]']);
  assert.deepEqual(safeMerge(s, [ent('INN', 'x', 'ООО "Елка"')]).tags, ['[ИНН_1]'], 'другой тип не совпадает');
});

test('пополнение: фамилия без имени идёт к единственному человеку с этой фамилией, ФИО предпочитает именительный', () => {
  let m = safeMerge(safeNew(), [ent('PERSON', 'петров|И', 'Петрова Ивана Сергеевича'), ent('PERSON', 'сидоров|?', 'Сидоровым')]);
  assert.equal(m.safe.records[0].value, 'Петрова Ивана Сергеевича');
  m = safeMerge(m.safe, [ent('PERSON', 'петров|?', 'Петрову'), ent('PERSON', 'сидоров|П', 'Сидоров Павел Петрович'), ent('PERSON', 'петров|И', 'Петров Иван Сергеевич')]);
  assert.deepEqual(m.tags, ['[ФИО_1]', '[ФИО_2]', '[ФИО_1]']);
  assert.equal(m.safe.records[0].value, 'Петров Иван Сергеевич', 'для подстановки именительный падеж');
  assert.equal(m.safe.records[1].value, 'Сидоров Павел Петрович');
  // два человека с одной фамилией: одинокая фамилия не угадывается
  m = safeMerge(m.safe, [ent('PERSON', 'петров|А', 'Петрова Анна Ивановна'), ent('PERSON', 'петров|?', 'Петровой')]);
  assert.deepEqual(m.tags, ['[ФИО_3]', '[ФИО_4]']);
});

test('свои слова: без повторов и пустых, сейф служит словарём', () => {
  const s = safeSetWords(safeNew('', T0), [' Казэнерго ', 'казэнерго', 'А', '', 'проект Север-2'], { now: T1 });
  assert.deepEqual(s.words, ['Казэнерго', 'проект Север-2']);
  assert.equal(s.changed, T1.toISOString());
  assert.equal(safeSetWords(s, ['Казэнерго', 'проект Север-2']), s, 'те же слова: сейф не меняется');
});

test('написания для поиска в новом документе: без коротких, дат, своих слов; суммы по переключателю', () => {
  const s = safeMerge(safeNew(), [
    ent('ORG', 'а', 'ООО «Каменный Мост»', ['ООО «Каменный Мост»', 'Каменный Мост']), ent('ID', 'id|12', '12'), ent('ID', 'id|41-кв/2025', '41-КВ/2025'),
    ent('DATE', '01022015', '01.02.2015'), ent('TERM', 'север', 'Север'), ent('AMOUNT', 's', '4 896 000,00'), ent('INN', '1', '7816135799'),
    ent('ACCOUNT', 'x', 'Петров\tИван'),
  ]).safe;
  const f = safeKnownForms(s).map(x => x.form);
  assert.deepEqual(f, ['ООО «Каменный Мост»', 'Каменный Мост', '4 896 000,00', '41-КВ/2025', '7816135799']);
  assert.ok(!safeKnownForms(s, { sums: false }).some(x => x.t === 'AMOUNT'));
});

test('обратная подстановка: искажённые метки, старые метки, неизвестные и просклонённые', () => {
  const s = safeMerge(safeNew(), [ent('ORG', 'а', 'ООО «Каменный Мост»'), ent('PERSON', 'петров|И', 'Петров Иван Сергеевич'),
    ent('ACCOUNT', '1', '40702810330000087432'), ent('ID', 'id|1', '41-КВ/2025')]).safe;
  const r = safeRestore([
    'Точно: [ОРГ_1], [ФИО_1], [СЧЁТ_1], [ДОГОВОР_1].',
    'Регистр и пробелы: [орг_1], [Фио 1], [ ОРГ_1 ], [ОРГ1].',
    'Без скобок: ОРГ_1, орг_1, ФИО-1, ОРГ 1.',
    'Ё и е: [СЧЕТ_1], [счет 1].',
    'Markdown: \\[ОРГ\\_1\\], **[ФИО_1]**.',
    'Чужое: [ОРГ_9], ОРГ_9, ИНН 7816135799, счёт 1 и Счёт 60, договор 1.',
    'Склонение: [ОРГ_1]у и [ФИО_1]а.',
  ].join('\n'), s);
  const lines = r.text.split('\n');
  assert.equal(lines[0], 'Точно: ООО «Каменный Мост», Петров Иван Сергеевич, 40702810330000087432, 41-КВ/2025.');
  assert.equal(lines[1], 'Регистр и пробелы: ООО «Каменный Мост», Петров Иван Сергеевич, ООО «Каменный Мост», ООО «Каменный Мост».');
  assert.equal(lines[2], 'Без скобок: ООО «Каменный Мост», ООО «Каменный Мост», Петров Иван Сергеевич, ООО «Каменный Мост».');
  assert.equal(lines[3], 'Ё и е: 40702810330000087432, 40702810330000087432.');
  assert.equal(lines[4], 'Markdown: ООО «Каменный Мост», **Петров Иван Сергеевич**.');
  assert.equal(lines[5], 'Чужое: [ОРГ_9], ОРГ_9, ИНН 7816135799, счёт 1 и Счёт 60, договор 1.', 'обычный текст не трогается');
  assert.equal(lines[6], 'Склонение: [ОРГ_1]у и [ФИО_1]а.', 'просклонённая метка остаётся как есть');
  assert.deepEqual(r.unknown, ['[ОРГ_9]', 'ОРГ_9']);
  assert.deepEqual(r.declined, ['[ОРГ_1]', '[ФИО_1]']);
  assert.equal(r.replaced, 16);
  assert.ok(r.parts.some(p => p.unknown === '[ОРГ_9]') && r.parts.some(p => p.tag === '[ОРГ_1]' && p.type === 'ORG'));
  assert.equal(r.parts.map(p => p.text).join(''), r.text);
});
