// Пустые куски <r> после замены в xlsx: убираются целиком, крайние случаи и все наборы файлов
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { buildZip, readZip, byName } from '../src/zip.mjs';
import { replaceInXlsx } from '../src/xlsx-return.mjs';
import { emptyRuns } from './xlsx-runs.mjs';
import { listFiles, roundtrip } from './roundtrip.mjs';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const run = (t, b) => `<r>${b ? '<rPr><b/></rPr>' : ''}<t xml:space="preserve">${t}</t></r>`;
async function book() {
  const ss = `<sst ${NS} count="3" uniqueCount="3">`
    + `<si>${run('Ива', true)}${run('нов')}${run(' и партнёры')}</si>`     // совпадение с начала: второй кусок пустеет
    + `<si>${run('Петров', true)}${run(' Пётр')}</si>`                     // правка по месту пустым: пустеет первый кусок
    + `<si>${run('Сидоров', true)}${run(' Олег')}</si></sst>`;            // пусты все куски
  const sheet = `<worksheet ${NS}><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c>`
    + `<c r="D1" t="inlineStr"><is>${run('Ко', true)}${run('валёв')}</is></c></row></sheetData></worksheet>`;
  const comments = `<comments ${NS}><authors><author>Иван</author></authors><commentList><comment ref="A1" authorId="0"><text>${run('Иван:', true)}${run('\nпроверить Ковалёв')}</text></comment></commentList></comments>`;
  return buildZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'],
    ['xl/workbook.xml', `<workbook ${NS}><sheets><sheet name="Л" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'],
    ['xl/sharedStrings.xml', ss], ['xl/worksheets/sheet1.xml', sheet], ['xl/comments1.xml', comments],
  ]);
}

test('xlsx: пустые куски <r> удаляются, оставшийся кусок сохраняет форматирование', async () => {
  const edits = new Map([['xl/worksheets/sheet1.xml!B1', [{ a: 0, b: 6, to: '' }]], ['xl/worksheets/sheet1.xml!C1', [{ a: 0, b: 12, to: '' }]]]);
  const { bytes } = await replaceInXlsx(await book(), [['Иванов', '#P1#'], ['Ковалёв', '#P2#']], { edits });
  const z = byName(readZip(bytes));
  const ss = await z['xl/sharedStrings.xml'].text(), s1 = await z['xl/worksheets/sheet1.xml'].text(), cm = await z['xl/comments1.xml'].text();
  assert.ok(ss.includes(`<si>${run('#P1#', true)}${run(' и партнёры')}</si>`), 'совпадение с начала: пустой второй кусок убран');
  assert.ok(s1.includes(`<is>${run('#P2#', true)}</is>`), 'inlineStr: один кусок со своим rPr');
  assert.match(s1, /<c r="B1" t="inlineStr"><is><t xml:space="preserve"> Пётр<\/t><\/is><\/c>/, 'правка по месту в общей строке: ячейка получила свою строку');
  assert.match(s1, /<c r="C1" t="inlineStr"><is><t><\/t><\/is><\/c>/, "вся ячейка ушла в правку: пустая строка без <r>");
  assert.ok(cm.includes('#P2#'));
  for (const [n, e] of Object.entries(z)) if (/\.xml$/.test(n)) assert.equal(emptyRuns(await e.text()), 0, n);
});

test('xlsx: пустой первый кусок и все пустые куски в общей строке', async () => {
  const { bytes } = await replaceInXlsx(await book(), [['Петров', ''], ['Сидоров Олег', '']]);
  const ss = await byName(readZip(bytes))['xl/sharedStrings.xml'].text();
  assert.ok(ss.includes(`<si><r><t xml:space="preserve"> Пётр</t></r></si>`), 'пустой первый кусок убран, второй остался со своим видом');
  assert.ok(ss.includes('<si><t/></si>'), 'все куски пусты: один <t/> без <r>');
  assert.equal(emptyRuns(ss), 0);
});

test('xlsx: ни одного пустого куска <r> после сквозного цикла на наборах', async () => {
  for (const d of ['fixtures', 'files', 'deal']) for (const f of listFiles(new URL(`./${d}/`, import.meta.url).pathname)) {
    if (!f.endsWith('.xlsx')) continue;
    const row = await roundtrip(f);
    for (const e of readZip(row.out.bytes).entries) if (/\.xml$/.test(e.name)) assert.equal(emptyRuns(await e.text()), 0, `${f}: ${e.name}`);
  }
});
