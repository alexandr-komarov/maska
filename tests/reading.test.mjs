// Чтение docx, xlsx и текста через src/office.mjs: находки стенда harness-files и сквозной цикл.
// Файлы собираются здесь же из XML, данные выдуманы.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { buildZip, readZip, byName } from '../src/zip.mjs';
import { numText, dateText } from '../src/xlsx-return.mjs';
import { readOffice, buildReturn, planReplacements, OLD_FORMAT, hideSums, isShort } from '../src/office.mjs';
import { loadMask } from './mask-api.mjs';
import { listFiles, roundtrip } from './roundtrip.mjs';

const api = loadMask();
const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
const RELS = list => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.join('')}</Relationships>`;
const CT = ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'];
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');

// Книга: sheets = [{name, state, rows: [[значение, ...]]}], значение - строка, число или {v, s, t, f}
function xlsx(sheets, { styles, comments, date1904, files = [] } = {}) {
  const shared = [];
  const sheetXml = sh => `<worksheet ${S}><sheetData>${sh.rows.map((row, i) => `<row r="${i + 1}">${row.map((v, j) => {
    const ref = String.fromCharCode(65 + j) + (i + 1);
    if (v == null) return '';
    if (typeof v === 'number') return `<c r="${ref}"><v>${v}</v></c>`;
    if (typeof v === 'object') return `<c${v.noRef ? '' : ` r="${ref}"`}${v.s != null ? ` s="${v.s}"` : ''}${v.t ? ` t="${v.t}"` : ''}>${v.f ? `<f>${esc(v.f)}</f>` : ''}<v>${esc(String(v.v))}</v></c>`;
    shared.push(v); return `<c r="${ref}" t="s"><v>${shared.length - 1}</v></c>`;
  }).join('')}</row>`).join('')}</sheetData></worksheet>`;
  // файлы листов нумеруются в обратном порядке: имя листа должно браться по связям, а не по номеру файла
  const n = sheets.length, xmls = sheets.map(sheetXml);
  const out = [CT,
    ['xl/workbook.xml', `<workbook ${S}>${date1904 ? '<workbookPr date1904="1"/>' : ''}<sheets>${sheets.map((sh, i) => `<sheet name="${sh.name}" sheetId="${i + 1}"${sh.state ? ` state="${sh.state}"` : ''} r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', RELS([...sheets.map((sh, i) => REL(`rId${i + 1}`, 'worksheet', `worksheets/sheet${n - i}.xml`)), REL('rIdS', 'styles', 'styles.xml'), REL('rIdSS', 'sharedStrings', 'sharedStrings.xml')])],
    ...sheets.map((sh, i) => [`xl/worksheets/sheet${n - i}.xml`, xmls[i]]),
    ['xl/sharedStrings.xml', `<sst ${S}>${shared.map(s => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`],
    ['xl/styles.xml', styles || `<styleSheet ${S}><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`],
  ];
  if (comments) {
    out.push([`xl/worksheets/_rels/sheet${n}.xml.rels`, RELS([REL('rIdC', 'comments', '../comments7.xml')])]);
    out.push(['xl/comments7.xml', `<comments ${S}><authors><author>А</author></authors><commentList>${comments.map(([ref, t]) => `<comment ref="${ref}" authorId="0"><text><r><t xml:space="preserve">${esc(t)}</t></r></text></comment>`).join('')}</commentList></comments>`]);
  }
  return buildZip([...out, ...files]);
}
const DATE_STYLES = `<styleSheet ${S}><numFmts count="1"><numFmt numFmtId="164" formatCode="DD\\.MM\\.YYYY"/></numFmts><cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/><xf numFmtId="4"/></cellXfs></styleSheet>`;

test('числа: экспонента, хвост .0, шум, длинные числа', () => {
  assert.deepEqual(numText('9.841726358E9'), { text: '9841726358' });
  assert.equal(numText('7779.0').text, '7779');
  assert.equal(numText('1284500.5').text, '1284500.5');
  assert.equal(numText('0.30000000000000004').text, '0.3');
  assert.equal(numText('-1.5E3').text, '-1500');
  assert.equal(numText('1E-7').text, '1E-7');
  assert.deepEqual(numText('4.0702810938000001E+19'), { text: '40702810938000000000', approx: true });
  assert.equal(numText('40702810938000012345').approx, true, 'целое длиннее 15 цифр');
  assert.equal(numText('7729654321').approx, false);
});

test('даты: система 1900 с 29.02.1900 и система 1904', () => {
  assert.equal(dateText(32581), '14.03.1989');
  assert.equal(dateText(1), '01.01.1900');
  assert.equal(dateText(59), '28.02.1900');
  assert.equal(dateText(60), '29.02.1900');
  assert.equal(dateText(61), '01.03.1900');
  assert.equal(dateText(45901.75), '01.09.2025');
  assert.equal(dateText(0, true), '01.01.1904');
  assert.equal(dateText(0), null);
});

test('xlsx: имена листов по связям, скрытые листы, даты по стилю, числа, примечания', async () => {
  const bytes = await xlsx([
    { name: 'Первая', rows: [['Сотрудник', 'Дата рождения', 'Сумма', 'ИНН'], ['Орлов Олег Петрович', { v: 32581, s: 1 }, { v: 7779.5, s: 3 }, { v: '9.841726358E9' }], ['Итого', { v: 45901, s: 2 }, { v: '7779.0' }]] },
    { name: 'Вторая', state: 'hidden', rows: [['на втором листе']] },
    { name: 'Служебный', state: 'veryHidden', rows: [['совсем скрытый']] },
  ], { styles: DATE_STYLES, comments: [['A2', 'позвонить Орлову О.П.']] });
  const r = await readOffice('kniga.xlsx', bytes);
  assert.equal(r.text, [
    '=== Лист «Первая» ===', 'Сотрудник\tДата рождения\tСумма\tИНН', 'Орлов Олег Петрович\t14.03.1989\t7779.5\t9841726358', 'Итого\t01.09.2025\t7779',
    '--- примечание к A2 ---', 'позвонить Орлову О.П.',
    '=== Лист «Вторая» (скрытый лист) ===', 'на втором листе',
    '=== Лист «Служебный» (очень скрытый лист) ===', 'совсем скрытый',
  ].join('\n'));
  assert.deepEqual(r.info.hidden.map(h => h.state), ['hidden', 'veryHidden']);
  assert.match(r.note, /скрытые листы: «Вторая», «Служебный»/);
  assert.equal(r.info.dates, 2);
  // сквозной цикл: дата рождения прячется вместе с меткой, ИНН в экспоненте тоже
  const m = api.mask(r.text, ['14.03.1989', '9841726358'], { sums: hideSums('kniga.xlsx') });
  const out = await buildReturn({ name: 'kniga.xlsx', bytes, read: r, hits: m.hits, vault: m.vault, masked: m.masked });
  assert.deepEqual(out.left, []);
  const after = (await readOffice('x.xlsx', out.bytes)).text;
  assert.ok(!after.includes('14.03.1989') && !after.includes('9841726358') && !after.includes('Орлов'));
  // то, чего маска не нашла, остаётся как было (какие даты и числа прятать, решают правила)
  for (const v of ['01.09.2025', '7779.5', '7779']) if (!m.hits.some(h => h.value.includes(v))) assert.ok(after.includes(v), v);
  assert.ok(out.warnings.some(w => /Скрытые листы: «Вторая», «Служебный»/.test(w)));
  assert.ok(out.warnings.some(w => /число заменено меткой/.test(w)));
});

test('xlsx: ячейки без адреса, 2500 строк без лимита', async () => {
  const rows = [[{ v: 1, noRef: true }, { v: 'Б', t: 'str', noRef: true }]];
  for (let i = 2; i <= 2500; i++) rows.push([`строка ${i}`, i]);
  rows.push(['Ковалёв Иван Сергеевич', 'последняя']);
  const bytes = await xlsx([{ name: 'Длинный', rows }]);
  const r = await readOffice('d.xlsx', bytes);
  const lines = r.text.split('\n');
  assert.equal(lines.length, 2502);
  assert.equal(lines[1], '1\tБ');
  assert.equal(lines[2500], 'строка 2500\t2500');
  const m = api.mask(r.text, [], { sums: false });
  const out = await buildReturn({ name: 'd.xlsx', bytes, read: r, hits: m.hits, vault: m.vault, masked: m.masked });
  assert.deepEqual(out.left, []);
  assert.ok(!(await readOffice('d.xlsx', out.bytes)).text.includes('Ковалёв'));
});

test('xlsx: большая книга, 880 тысяч знаков, ход обработки по листам', async () => {
  const sheets = [];
  for (let k = 0; k < 4; k++) {
    const rows = [];
    for (let i = 0; i < 5500; i++) rows.push([`позиция ${k}-${i}`, `склад номер ${i % 50}`, i * 3, `материал марки М${i % 300}`]);
    sheets.push({ name: `Лист${k + 1}`, rows });
  }
  const bytes = await xlsx(sheets);
  const calls = [];
  const t0 = Date.now();
  const r = await readOffice('big.xlsx', bytes, { progress: (k, n, name) => calls.push([k, n, name]) });
  const tRead = Date.now() - t0;
  assert.ok(r.text.length > 880000, `знаков ${r.text.length}`);
  assert.deepEqual(calls.map(c => c[0]), [0, 1, 2, 3]);
  assert.equal(calls[0][1], 4);
  const m = api.mask(r.text, [], { sums: false });
  const t1 = Date.now();
  const out = await buildReturn({ name: 'big.xlsx', bytes, read: r, hits: m.hits, vault: m.vault, masked: m.masked });
  const tRet = Date.now() - t1;
  assert.deepEqual(out.left, []);
  // запас для медленной машины: на этой чтение около 1 с, возврат с проверкой около 2 с
  assert.ok(tRead < 10000 && tRet < 20000, `чтение ${tRead} мс, возврат ${tRet} мс`);
});

test('xlsx: находка через границу ячеек заменяется по месту, общая строка в других ячейках остаётся', async () => {
  const bytes = await xlsx([{ name: 'Кадры', rows: [['Фамилия', 'Имя', 'Отчество'], ['Сидоров', 'Павел', 'Константинович'], ['Павел', 'другой', 'текст']] }]);
  const r = await readOffice('k.xlsx', bytes);
  const m = api.mask(r.text, [], { sums: false });
  const hit = m.hits.find(h => h.value.includes('\t'));
  assert.ok(hit, 'маскировщик нашёл ФИО через табуляцию');
  const plan = planReplacements(r, m.hits, m.vault);
  assert.ok(plan.edits.size >= 2 && plan.split.has(hit.value));
  const out = await buildReturn({ name: 'k.xlsx', bytes, read: r, hits: m.hits, vault: m.vault, masked: m.masked });
  assert.deepEqual(out.left, []);
  const after = (await readOffice('k.xlsx', out.bytes)).text.split('\n');
  assert.equal(after[2], `${hit.tag}`, 'метка в первой ячейке, остальные куски пусты');
  assert.equal(after[3], 'Павел\tдругой\tтекст', 'та же общая строка в другой ячейке не тронута');
});

// docx: sdt на уровне тела, вложенная таблица, надпись, сноска, примечание, колонтитул, удалённый текст
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"';
const p = t => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
async function docx() {
  const box = t => `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent>${p(t)}</w:txbxContent></wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox><w:txbxContent>${p(t)}</w:txbxContent></v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
  const body = `<w:sdt><w:sdtPr/><w:sdtContent>${p('Титул: Голубев Роман Ильич')}</w:sdtContent></w:sdt>`
    + `<w:tbl><w:tr><w:tc>${p('внешняя')}<w:tbl><w:tr><w:tc>${p('вложенная ячейка')}</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>`
    + box('Надпись: Лаптев И.В.')
    + `<w:p><w:r><w:t xml:space="preserve">Подписант: </w:t></w:r><w:del w:id="1" w:author="Корректор"><w:r><w:delText>Сорокин Павел Андреевич</w:delText></w:r></w:del><w:ins w:id="2" w:author="Корректор"><w:r><w:t>Орлов Олег Петрович</w:t></w:r></w:ins></w:p>`;
  return buildZip([CT,
    ['word/document.xml', `<w:document ${W}><w:body>${body}</w:body></w:document>`],
    ['word/footnotes.xml', `<w:footnotes ${W}><w:footnote w:id="1">${p('Сноска: паспорт у Белова Сергея Ивановича')}</w:footnote></w:footnotes>`],
    ['word/endnotes.xml', `<w:endnotes ${W}><w:endnote w:id="1">${p('Концевая: тел. +7 916 245-18-73')}</w:endnote></w:endnotes>`],
    ['word/comments.xml', `<w:comments ${W}><w:comment w:id="0" w:author="Рецензент">${p('Примечание: уточнить у Зайцевой Анны Олеговны')}</w:comment></w:comments>`],
    ['word/header1.xml', `<w:hdr ${W}>${p('ООО «Северный Ветер»')}</w:hdr>`],
    ['word/footer1.xml', `<w:ftr ${W}>${p('стр. 1')}</w:ftr>`],
    ['docProps/core.xml', '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Договор с Голубевым Р.И.</dc:title><dc:creator>Иванов</dc:creator></cp:coreProperties>'],
  ]);
}

test('docx: sdt, вложенная таблица и надпись без дублей, сноски, примечания, колонтитулы, удалённый текст, свойства', async () => {
  const bytes = await docx();
  const r = await readOffice('d.docx', bytes);
  assert.equal(r.text, [
    // пустая строка: абзац, в котором стоит надпись, сам текста не имеет
    'Титул: Голубев Роман Ильич', 'внешняя', 'вложенная ячейка', '', 'Надпись: Лаптев И.В.', 'Подписант: Орлов Олег Петрович',
    '=== Верхний колонтитул 1 ===', 'ООО «Северный Ветер»',
    '=== Сноски ===', 'Сноска: паспорт у Белова Сергея Ивановича',
    '=== Концевые сноски ===', 'Концевая: тел. +7 916 245-18-73',
    '=== Примечания ===', 'Примечание: уточнить у Зайцевой Анны Олеговны',
    '=== Исправления: удалённый текст, абзацы в редакции до правки ===', 'Подписант: Сорокин Павел Андреевич',
    '=== Свойства файла ===', 'Договор с Голубевым Р.И.',
  ].join('\n'));
  assert.match(r.note, /исправления в 1 абзаце/);
  const m = api.mask(r.text, [], { sums: true });
  const out = await buildReturn({ name: 'd.docx', bytes, read: r, hits: m.hits, vault: m.vault, masked: m.masked });
  assert.deepEqual(out.left, []);
  const doc = await byName(readZip(out.bytes))['word/document.xml'].text();
  assert.ok(!doc.includes('Лаптев'), 'надпись заменена и в mc:Fallback');
  assert.ok(!doc.includes('Сорокин'), 'удалённый текст заменён меткой');
});

test('старые форматы и повреждённые файлы: понятное сообщение', async () => {
  const ole = new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0]);
  await assert.rejects(readOffice('Договор.doc', ole), { message: OLD_FORMAT.doc });
  await assert.rejects(readOffice('Реестр.XLS', ole), { message: OLD_FORMAT.xls });
  await assert.rejects(readOffice('переименован.xlsx', ole), { message: OLD_FORMAT.ole });
  await assert.rejects(readOffice('битый.docx', new TextEncoder().encode('не архив')), { message: OLD_FORMAT.zip });
  assert.match(OLD_FORMAT.doc, /сохраните как \.docx/);
  assert.match(OLD_FORMAT.xls, /сохраните как \.xlsx/);
});

test('текст: UTF-8 и Windows-1251, возврат в том же формате', async () => {
  const csv = 'ФИО;ИНН\r\nОрлов Олег Петрович;500100732259\r\n';
  const r = await readOffice('spisok.csv', new TextEncoder().encode('﻿' + csv));
  assert.equal(r.text, csv.replace(/\r\n/g, '\n'));
  const m = api.mask(r.text, [], { sums: hideSums('spisok.csv') });
  const out = await buildReturn({ name: 'spisok.csv', bytes: null, read: r, hits: m.hits, vault: m.vault, masked: m.masked });
  const txt = new TextDecoder('utf-8', { ignoreBOM: true }).decode(out.bytes);
  assert.equal(out.name, 'spisok-maska.csv');
  assert.ok(txt.startsWith('﻿') && txt.includes('\r\n') && !txt.includes('Орлов') && !txt.includes('500100732259'));
  // cp1251: «Привет» в Windows-1251
  const r2 = await readOffice('a.txt', new Uint8Array([0xCF, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2]));
  assert.equal(r2.text, 'Привет');
  assert.match(r2.note, /1251/);
});

test('суммы: в таблицах не прячутся, в текстах прячутся', () => {
  assert.equal(hideSums('a.xlsx'), false);
  assert.equal(hideSums('a.CSV'), false);
  assert.equal(hideSums('a.docx'), true);
  assert.equal(hideSums('a.txt'), true);
});

test('сквозной цикл на tests/fixtures, tests/files, tests/deal: найденного в результате нет', async () => {
  for (const d of ['fixtures', 'files', 'deal']) for (const f of listFiles(new URL(`./${d}/`, import.meta.url).pathname)) {
    const row = await roundtrip(f, api);
    assert.deepEqual(row.left, [], f);
    assert.ok(row.hits > 0, f);
  }
});

test('короткое значение меняется только там, где его нашла маска', async () => {
  const bytes = await xlsx([{ name: 'Л', rows: [['Договор № 12 от 2012 года', 'площадь 112'], ['Акт № 12', 12]] }]);
  const r = await readOffice('k.xlsx', bytes);
  const s = r.text.indexOf('12'), s2 = r.text.indexOf('Акт № 12') + 6;
  const hits = [{ s, e: s + 2, value: '12', tag: '#ID_1#' }, { s: s2, e: s2 + 2, value: '12', tag: '#ID_1#' }];
  assert.ok(isShort('12') && isShort('А-5') && isShort('12345') && !isShort('123456') && !isShort('Орлов'));
  const plan = planReplacements(r, hits, [{ tag: '#ID_1#', value: '12', forms: new Set(['12']) }]);
  assert.deepEqual(plan.replacements, [], 'короткое значение не идёт в общий список');
  const out = await buildReturn({ name: 'k.xlsx', bytes, read: r, hits, vault: [], masked: '' });
  assert.deepEqual(out.left, []);
  const after = (await readOffice('k.xlsx', out.bytes)).text;
  assert.ok(after.includes('Договор № #ID_1# от 2012 года\tплощадь 112'));
  assert.ok(after.includes('Акт № #ID_1#\t12'), 'число 12 в соседней ячейке маска не находила, оно осталось');
});
