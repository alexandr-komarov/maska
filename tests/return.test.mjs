// Возврат файла в исходном формате: zip, docx, xlsx. Запуск: node --test tests/
// Тестовые файлы в tests/fixtures/ собирает node tests/fixtures/make.mjs, данные в них выдуманы.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync, spawnSync } from 'node:child_process';
import { readZip, writeZip, byName, crc32 } from '../src/zip.mjs';
import { replaceInDocx, extractDocxText } from '../src/docx-return.mjs';
import { replaceInXlsx, extractXlsxText } from '../src/xlsx-return.mjs';

const FX = new URL('./fixtures/', import.meta.url).pathname;
const fx = n => new Uint8Array(fs.readFileSync(FX + n));
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'maska-ret-'));
process.on('exit', () => fs.rmSync(tmp, { recursive: true, force: true }));

// zip.testzip и разбор каждого XML в Python
const PY = `
import sys, zipfile, json, xml.dom.minidom as m
z = zipfile.ZipFile(sys.argv[1]); bad = z.testzip(); errs = []
for n in z.namelist():
    if n.endswith(('.xml', '.rels', '.vml')):
        try: m.parseString(z.read(n))
        except Exception as e: errs.append(n + ': ' + str(e))
print(json.dumps({'bad': bad, 'xml': errs}))`;
export function pyCheck(file) { return JSON.parse(execFileSync('python3', ['-c', PY, file], { encoding: 'utf8' })); }

// записи, которые не менялись, совпадают побайтно и стоят в том же порядке
async function sameUntouched(src, out, changed) {
  const a = readZip(src).entries, b = readZip(out).entries;
  assert.deepEqual(b.map(e => e.name), a.map(e => e.name), 'порядок записей');
  for (let i = 0; i < a.length; i++) if (!changed.includes(a[i].name)) assert.deepEqual(b[i].local, a[i].local, `запись ${a[i].name} изменилась`);
}

// LibreOffice: есть ли soffice и какие модули стоят
const prof = `-env:UserInstallation=file://${tmp}/profile`;
function soffice(args) { return spawnSync('soffice', [prof, '--headless', ...args, '--outdir', tmp], { encoding: 'utf8', timeout: 180000 }); }
const hasSoffice = spawnSync('soffice', ['--version'], { encoding: 'utf8' }).status === 0;
let hasWriter = null;
function writerOk() {
  if (hasWriter !== null) return hasWriter;
  fs.writeFileSync(path.join(tmp, 'probe.txt'), 'проба');
  hasWriter = hasSoffice && soffice(['--convert-to', 'odt', path.join(tmp, 'probe.txt')]).status === 0;
  return hasWriter;
}

const R = [
  ['ООО «Ромашка»', '#ORG_1#'], ['Петров Пётр Ильич', '#PERSON_1#'], ['Калинин', '#PERSON_2#'], ['Бета & Ко', '#ORG_2#'],
  ['7729654321', '#INN_1#'], ['Сидор\tСидоров', '#TERM_1#'], ['Сорокин Павел', '#PERSON_3#'], ['Орлов Олег', '#PERSON_4#'],
  ['m.kuznetsova@sev-terminal.example', '#EMAIL_1#'], ['Паспорт 45 18 № 603217', '#PASSPORT_1#'], ['+7 (916) 245-18-73', '#PHONE_1#'],
  ['Кузнецова Марина Олеговна', '#PERSON_5#'], ['40702810938000012345', '#ACCOUNT_1#'], ['Абрамов Дмитрий', '#PERSON_6#'],
  ['500100732259', '#INN_2#'], ['Иванов Иван', '#PERSON_7#'], ['Такого нет нигде', '#ORG_9#'],
];
const has = (text, s) => text.replace(/ /g, ' ').includes(s);

test('zip: чтение и запись без изменений дают те же байты', async () => {
  for (const f of ['synthetic.docx', 'synthetic.xlsx', 'libreoffice.xlsx']) {
    if (!fs.existsSync(FX + f)) continue;
    const src = fx(f);
    assert.deepEqual(await writeZip(readZip(src)), src, f);
  }
});

test('zip: понятные ошибки на чужом формате и ZIP64', () => {
  assert.throws(() => readZip(new TextEncoder().encode('это не архив, а старый .doc')), /не zip/);
  const z64 = new Uint8Array(42), dv = new DataView(z64.buffer);
  dv.setUint32(0, 0x07064b50, true); dv.setUint32(20, 0x06054b50, true); dv.setUint16(30, 0xFFFF, true);
  assert.throws(() => readZip(z64), /ZIP64/);
});

test('zip: CRC32 совпадает с эталоном', () => {
  assert.equal(crc32(new TextEncoder().encode('123456789')), 0xCBF43926);
});

test('docx: замены по всему документу', async t => {
  const src = fx('synthetic.docx');
  const before = (await extractDocxText(src)).text;
  const { bytes, report } = await replaceInDocx(src, R);
  const file = path.join(tmp, 'synthetic-out.docx');
  fs.writeFileSync(file, bytes);
  const after = (await extractDocxText(bytes)).text;
  const z = byName(readZip(bytes));
  const doc = await z['word/document.xml'].text();

  await t.test('архив и XML целы', () => {
    const r = pyCheck(file);
    assert.equal(r.bad, null); assert.deepEqual(r.xml, []);
  });
  await t.test('нетронутые записи совпадают побайтно', () => sameUntouched(src, bytes, report.changedParts));
  await t.test('найденное в извлечённом тексте находится при замене', () => {
    for (const [from] of R) if (has(before, from)) assert.ok(!report.notFound.includes(from), `в тексте есть, заменой не найдено: ${from}`);
    assert.ok(report.notFound.includes('Такого нет нигде'));
  });
  await t.test('исходных строк нет, метки есть', () => {
    for (const [from, to] of R) {
      assert.ok(!has(after, from), `осталось: ${from}`);
      if (has(before, from)) assert.ok(after.includes(to), `нет метки ${to}`);
    }
    assert.deepEqual(report.leftovers, []);
  });
  await t.test('слово из трёх кусков: метка в первом куске с его форматированием', () => {
    assert.match(doc, /<w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">#ORG_1#<\/w:t>/);
    assert.match(doc, /<w:rPr><w:i\/><\/w:rPr><w:t xml:space="preserve"><\/w:t>/);
    assert.match(doc, /<w:color w:val="C00000"\/><\/w:rPr><w:t xml:space="preserve"> <\/w:t>/);
  });
  await t.test('неразрывный пробел, сущность, мягкий перенос', () => {
    assert.ok(after.includes('Представитель: #PERSON_1#, соисполнитель: #PERSON_2# и #PERSON_2# Олег'));
    assert.ok(after.includes('Партнёр: #ORG_2#'));
    assert.ok(!doc.includes('<w:softHyphen/>'), 'мягкий перенос внутри совпадения удалён');
  });
  await t.test('табуляция: вне совпадения остаётся, внутри совпадения удаляется', () => {
    assert.ok(after.includes('ИНН\t#INN_1#, код #TERM_1#'));
    assert.equal(doc.match(/<w:tab\/>/g).length, 1);
  });
  await t.test('колонтитул, сноска, концевая сноска, примечание, надпись, таблица', () => {
    assert.ok(report.replaced['word/header1.xml'] && report.replaced['word/footer1.xml'] && report.replaced['word/footnotes.xml'] && report.replaced['word/endnotes.xml'] && report.replaced['word/comments.xml']);
    assert.equal(report.byValue.find(x => x.from === 'Абрамов Дмитрий').count, 2, 'надпись в mc:Choice и в mc:Fallback');
    assert.ok(after.includes('#PERSON_5#\nСчёт\n#ACCOUNT_1#'));
  });
  await t.test('исправления: вставка и удалённый текст заменены, авторы обезличены', async () => {
    assert.match(doc, /<w:delText>#PERSON_3#<\/w:delText>/);
    assert.ok(after.includes('Подписант: #PERSON_4#'));
    for (const [n, e] of Object.entries(z)) if (/\.xml$/.test(n)) assert.ok(!(await e.text()).includes('Иванов И.И.'), n);
    assert.ok(doc.includes('w:author="Автор 1"'));
  });
  await t.test('поле HYPERLINK и свойства документа', async () => {
    assert.ok(doc.includes('HYPERLINK "mailto:#EMAIL_1#"'));
    const core = await z['docProps/core.xml'].text(), app = await z['docProps/app.xml'].text();
    assert.match(core, /<dc:creator><\/dc:creator>/); assert.match(core, /<cp:lastModifiedBy><\/cp:lastModifiedBy>/);
    assert.match(core, /<dc:title>Договор с #ORG_1#<\/dc:title>/);
    assert.match(app, /<Company><\/Company>/); assert.match(app, /<Manager><\/Manager>/);
  });
  await t.test('LibreOffice открывает результат', { skip: !hasSoffice ? 'нет soffice' : !writerOk() ? 'в LibreOffice нет модуля Writer, docx не открыть' : false }, () => {
    const r = soffice(['--convert-to', 'pdf', file]);
    assert.equal(r.status, 0, r.stderr);
    assert.ok(fs.statSync(file.replace(/docx$/, 'pdf')).size > 0);
    assert.equal(soffice(['--convert-to', 'txt:Text (encoded):UTF8', file]).status, 0);
    const txt = fs.readFileSync(file.replace(/docx$/, 'txt'), 'utf8');
    for (const tag of ['#ORG_1#', '#PERSON_1#', '#PERSON_5#']) assert.ok(txt.includes(tag), tag);
  });
});

// все формулы листа как есть
const formulas = s => s.match(/<(?:\w+:)?f\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?f>)/g) || [];

test('xlsx: замены в книге', async t => {
  const src = fx('synthetic.xlsx');
  const before = (await extractXlsxText(src)).text;
  const { bytes, report } = await replaceInXlsx(src, R);
  const file = path.join(tmp, 'synthetic-out.xlsx');
  fs.writeFileSync(file, bytes);
  const after = (await extractXlsxText(bytes)).text;
  const z0 = byName(readZip(src)), z = byName(readZip(bytes));
  const s1 = await z['xl/worksheets/sheet1.xml'].text();

  await t.test('архив и XML целы', () => { const r = pyCheck(file); assert.equal(r.bad, null); assert.deepEqual(r.xml, []); });
  await t.test('нетронутые записи совпадают побайтно', () => sameUntouched(src, bytes, report.changedParts));
  await t.test('формулы не изменились побайтно', async () => {
    for (const n of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']) assert.deepEqual(formulas(await z[n].text()), formulas(await z0[n].text()), n);
    assert.ok(formulas(s1).length >= 3);
  });
  await t.test('исходных строк в тексте нет, метки есть', () => {
    for (const [from, to] of R) {
      if (!has(before, from)) continue;
      assert.ok(!has(after, from), `осталось: ${from}`);
      assert.ok(after.includes(to), `нет метки ${to}`);
    }
  });
  await t.test('общая строка с форматированием и inlineStr', async () => {
    const ss = await z['xl/sharedStrings.xml'].text();
    assert.match(ss, /<si><t xml:space="preserve"> #PERSON_7# <\/t><\/si>/);
    // пустой кусок <r> Excel не принимает: он удаляется, единственный оставшийся кусок сохраняет своё форматирование
    assert.match(s1, /<c r="A4" t="inlineStr"><is><r><rPr><b\/><\/rPr><t>#PERSON_6#<\/t><\/r><\/is><\/c>/);
    assert.match(ss, /<si><r><rPr><b\/><sz val="11"\/><rFont val="Calibri"\/><\/rPr><t>#PERSON_5#<\/t><\/r><\/si>/);
    assert.equal(ss.match(/<si>/g).length, (await z0['xl/sharedStrings.xml'].text()).match(/<si>/g).length, 'число строк в sst то же');
  });
  await t.test('число, которое надо спрятать, стало строкой со старым стилем', () => {
    assert.match(s1, /<c r="B2" s="1" t="inlineStr"><is><t>#INN_1#<\/t><\/is><\/c>/);
    assert.match(s1, /<c r="B3" t="inlineStr"><is><t>#INN_2#<\/t><\/is><\/c>/);
    assert.ok(report.converted.find(c => c.cell === 'Клиенты!D2' && c.approx), 'счёт с потерей точности');
    assert.match(s1, /<c r="C5"><v>123456<\/v><\/c>/, 'посторонние числа не тронуты');
    assert.match(s1, /<c r="C4"><f>SUM\(C2:C3\)<\/f><v>3200<\/v><\/c>/);
  });
  await t.test('строка внутри формулы и кэш формулы попадают в отчёт', async () => {
    assert.ok(report.warnings.find(w => w.kind === 'formula' && w.cell === 'Клиенты!E2'));
    assert.ok(report.warnings.find(w => w.kind === 'cached' && w.cell === 'Клиенты!F2'));
    assert.ok(report.warnings.find(w => w.kind === 'cached' && w.cell === 'Итоги!B2'));
    assert.ok(report.warnings.find(w => w.kind === 'definedName'));
    assert.match(await z['xl/worksheets/sheet2.xml'].text(), /<c r="B2" s="1" t="str"><f>Клиенты!B2<\/f><v>#INN_1#<\/v><\/c>/);
    // остаток только там, где мы сознательно не меняем: формула и определённое имя
    assert.deepEqual(report.leftovers.map(l => l.part).sort(), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  });
  await t.test('примечание, колонтитул листа, свойства', async () => {
    const cm = await z['xl/comments1.xml'].text();
    assert.ok(cm.includes('<author>Автор 1</author>') && cm.includes('<t>Автор 1:</t>'));
    assert.ok(cm.includes('проверить ИНН #ORG_1#'));
    assert.ok(s1.includes('<oddHeader>&amp;L#ORG_1#&amp;RКонфиденциально</oddHeader>'));
    assert.match(await z['docProps/app.xml'].text(), /<Company><\/Company>/);
  });
  await t.test('LibreOffice открывает результат, в тексте метки', { skip: !hasSoffice && 'нет soffice' }, () => {
    const r = soffice(['--convert-to', 'pdf', file]);
    assert.equal(r.status, 0, r.stderr);
    assert.ok(fs.statSync(file.replace(/xlsx$/, 'pdf')).size > 0);
    assert.equal(soffice(['--convert-to', 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1', file]).status, 0);
    const csv = fs.readFileSync(path.join(tmp, 'synthetic-out-Клиенты.csv'), 'utf8') + fs.readFileSync(path.join(tmp, 'synthetic-out-Итоги.csv'), 'utf8');
    for (const tag of ['#ORG_1#', '#INN_1#', '#INN_2#', '#PERSON_5#', '#PERSON_6#', '#ACCOUNT_1#']) assert.ok(csv.includes(tag), tag);
    for (const [from] of R) assert.ok(!has(csv, from), `в выгрузке LibreOffice осталось: ${from}`);
  });
});

test('xlsx из LibreOffice (csv → xlsx)', { skip: !fs.existsSync(FX + 'libreoffice.xlsx') && 'нет libreoffice.xlsx' }, async () => {
  const src = fx('libreoffice.xlsx');
  const RR = [['Петров Пётр Ильич', '#PERSON_1#'], ['Кузнецова Марина Олеговна', '#PERSON_2#'], ['7729654321', '#INN_1#'], ['500100732259', '#INN_2#'], ['7701234567', '#INN_3#']];
  const { bytes, report } = await replaceInXlsx(src, RR);
  const file = path.join(tmp, 'lo-out.xlsx');
  fs.writeFileSync(file, bytes);
  const r = pyCheck(file); assert.equal(r.bad, null); assert.deepEqual(r.xml, []);
  await sameUntouched(src, bytes, report.changedParts);
  const after = (await extractXlsxText(bytes)).text;
  for (const [from, to] of RR) { assert.ok(!after.includes(from), from); assert.ok(after.includes(to), to); }
  assert.equal(report.converted.length, 3, 'ИНН из csv LibreOffice хранит числами');
  if (hasSoffice) assert.equal(soffice(['--convert-to', 'pdf', file]).status, 0);
});

test('docx: xml:space ставится, когда у куска появился крайний пробел', async () => {
  const { buildZip } = await import('../src/zip.mjs');
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const doc = `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:r><w:t>Ива</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>нов Пётр</w:t></w:r></w:p><w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="100"/></w:tabs></w:pPr><w:r><w:t>A&lt;B &amp; Иванов</w:t></w:r></w:p></w:body></w:document>`;
  const src = await buildZip([['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'], ['word/document.xml', doc]]);
  const { bytes, report } = await replaceInDocx(src, [['Иванов', '#PERSON_1#']]);
  const out = await byName(readZip(bytes))['word/document.xml'].text();
  assert.equal(report.total, 2);
  assert.ok(out.includes('<w:t>#PERSON_1#</w:t>'));
  assert.ok(out.includes('<w:t xml:space="preserve"> Пётр</w:t>'));
  assert.ok(out.includes('<w:t>A&lt;B &amp; #PERSON_1#</w:t>'));
  assert.ok(out.includes('<w:tab w:val="left" w:pos="100"/>'), 'позиции табуляции в свойствах абзаца не трогаются');
});
