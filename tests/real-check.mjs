// Структурная проверка возврата файла на настоящих docx и xlsx.
// Из каждого файла автоматически берутся 5-10 строк (частые слова с заглавной буквы, числа из 10-12 цифр),
// заменяются метками, результат проверяется: архив, XML, исходных строк нет, нетронутые записи и формулы целы.
// Печатаются только счётчики: содержимое файлов в вывод не попадает.
// Запуск: node tests/real-check.mjs <папка> [папка для результатов, по умолчанию out/real]
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { readZip, byName } from '../src/zip.mjs';
import { replaceInDocx, extractDocxText } from '../src/docx-return.mjs';
import { replaceInXlsx, extractXlsxText } from '../src/xlsx-return.mjs';
import { emptyRuns } from './xlsx-runs.mjs';

const dir = process.argv[2];
if (!dir) { console.log('укажите папку: node tests/real-check.mjs <папка>'); process.exit(2); }
const outDir = process.argv[3] || new URL('../out/real/', import.meta.url).pathname;
fs.mkdirSync(outDir, { recursive: true });

const PY = `
import sys, zipfile, json, xml.dom.minidom as m
z = zipfile.ZipFile(sys.argv[1]); bad = z.testzip(); errs = 0
for n in z.namelist():
    if n.endswith(('.xml', '.rels', '.vml')):
        try: m.parseString(z.read(n))
        except Exception: errs += 1
print(json.dumps({'bad': bad is not None, 'xml': errs}))`;

// похоже на попытку дать указание тому, кто читает файл
const INJECT = /ignore (all|previous|the above)|disregard|system prompt|you are (an?|the) (ai|assistant|model)|assistant:|игнорируй|не обращай внимания на|ты (—|-)? ?(ии|модель|ассистент)|выполни (команду|инструкцию)|claude|chatgpt|openai/gi;

const files = [];
(function walk(d) { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else if (/\.(docx|xlsx)$/i.test(e.name)) files.push(p); } })(dir);
files.sort();

function pick(text) {
  const words = new Map(), nums = new Map();
  for (const m of text.matchAll(/(?<![\p{L}\d])[А-ЯЁA-Z][а-яёa-z]{3,}(?![\p{L}])/gu)) words.set(m[0], (words.get(m[0]) || 0) + 1);
  for (const m of text.matchAll(/(?<![\d.,])\d{10,12}(?!\d|[.,]\d)/g)) nums.set(m[0], (nums.get(m[0]) || 0) + 1);
  const top = (m, k) => [...m].sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : 1)).slice(0, k).map(x => x[0]);
  const n = top(nums, 4), w = top(words, 10 - n.length);
  return [...w.map((s, i) => [s, `#W_${i + 1}#`]), ...n.map((s, i) => [s, `#D_${i + 1}#`])];
}
const formulas = s => s.match(/<(?:\w+:)?f\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?f>)/g) || [];
const nsp = s => s.replace(/ /g, ' ');

const rows = [];
for (const f of files) {
  const ext = f.split('.').pop().toLowerCase(), base = path.basename(f);
  const row = { file: path.relative(dir, f), ext, ok: false, fail: [] };
  rows.push(row);
  try {
    const src = new Uint8Array(fs.readFileSync(f));
    try { readZip(src); } catch (e) { row.fail.push('не zip: ' + e.message.split(' (')[0]); continue; }
    const extract = ext === 'docx' ? extractDocxText : extractXlsxText;
    const before = (await extract(src)).text;
    row.inject = (before.match(INJECT) || []).length;
    // строки-заголовки частей («=== Сноски ===», «--- примечание к A1 ---») пишет чтение, в файле их нет
    const isHead = l => /^(===|---) /.test(l);
    const R = pick(nsp(before).split('\n').filter(l => !isHead(l)).join('\n'));
    row.strings = R.length;
    if (R.length < 5) row.fail.push(`мало строк для проверки: ${R.length}`);
    const { bytes, report } = await (ext === 'docx' ? replaceInDocx : replaceInXlsx)(src, R);
    const outFile = path.join(outDir, base.replace(/\.(docx|xlsx)$/i, '.masked.$1'));
    fs.writeFileSync(outFile, bytes);
    Object.assign(row, {
      replaced: report.total, notFound: report.notFound.length, converted: report.converted.length,
      warnings: report.warnings.reduce((a, w) => (a[w.kind] = (a[w.kind] || 0) + 1, a), {}),
      notHandled: report.notHandled.reduce((a, w) => (a[w.kind] = (a[w.kind] || 0) + 1, a), {}),
      leftoverParts: [...new Set(report.leftovers.map(l => l.part.replace(/\d+(?=\.xml$)/, 'N')))],
      changedParts: report.changedParts.length,
    });
    // архив и XML
    const py = JSON.parse(execFileSync('python3', ['-c', PY, outFile], { encoding: 'utf8' }));
    if (py.bad) row.fail.push('testzip: битая запись');
    if (py.xml) row.fail.push(`некорректных XML: ${py.xml}`);
    // нетронутые записи побайтно
    const a = readZip(src).entries, b = readZip(bytes).entries;
    if (a.map(e => e.name).join('\n') !== b.map(e => e.name).join('\n')) row.fail.push('порядок записей');
    else row.diffUntouched = a.filter((e, i) => !report.changedParts.includes(e.name) && Buffer.compare(e.local, b[i].local) !== 0).length;
    if (row.diffUntouched) row.fail.push(`нетронутые записи изменились: ${row.diffUntouched}`);
    // формулы
    if (ext === 'xlsx') {
      const z0 = byName(readZip(src)), z1 = byName(readZip(bytes));
      let fx = 0, fdiff = 0;
      for (const n of Object.keys(z0)) if (/^xl\/worksheets\/.*\.xml$/.test(n)) {
        const f0 = formulas(await z0[n].text()), f1 = formulas(await z1[n].text());
        fx += f0.length; if (f0.join('\n') !== f1.join('\n')) fdiff++;
      }
      row.formulas = fx;
      if (fdiff) row.fail.push(`формулы изменились на листах: ${fdiff}`);
      // кусок <r> с пустым <t> Excel не принимает и предлагает восстановить файл
      row.emptyRuns = 0;
      for (const e of readZip(bytes).entries) if (/\.xml$/.test(e.name)) row.emptyRuns += emptyRuns(await e.text());
      if (row.emptyRuns) row.fail.push(`пустых кусков <r>: ${row.emptyRuns}`);
    }
    // исходных строк в тексте результата нет; строки заголовков с именами листов считаем отдельно
    const after = nsp((await extract(bytes)).text).split('\n');
    const sheetLines = after.filter(l => l.startsWith('=== Лист «')), body = after.filter(l => !isHead(l)).join('\n');
    row.leftInText = R.filter(([s]) => body.includes(s)).length;
    row.leftInSheetNames = R.filter(([s]) => sheetLines.some(l => l.includes(s))).length;
    // строка только в имени листа метку не получает: имена листов не меняем, это в отчёте отдельно
    const beforeBody = nsp(before).split('\n').filter(l => !isHead(l)).join('\n');
    row.labelsMissing = R.filter(([s, t]) => beforeBody.includes(s) && !after.filter(l => !isHead(l)).join('\n').includes(t)).length;
    if (row.leftInText) row.fail.push(`исходные строки остались в тексте: ${row.leftInText}`);
    if (row.labelsMissing) row.fail.push(`метки не появились: ${row.labelsMissing}`);
    row.ok = row.fail.length === 0;
  } catch (e) { row.fail.push('исключение: ' + e.message.split('\n')[0].slice(0, 120)); }
}

for (const r of rows) {
  const w = Object.entries(r.warnings || {}).map(([k, v]) => `${k}:${v}`).join(',') || '-';
  const nh = Object.entries(r.notHandled || {}).map(([k, v]) => `${k}:${v}`).join(',') || '-';
  console.log(`${r.ok ? 'OK  ' : 'FAIL'} ${r.file}`);
  if (r.strings != null) console.log(`     строк ${r.strings}, замен ${r.replaced}, не найдено ${r.notFound}, чисел в строку ${r.converted}, формул ${r.formulas ?? '-'}, частей изменено ${r.changedParts}, предупреждения ${w}, не обработано ${nh}, остаток в именах листов ${r.leftInSheetNames}, остаток в XML: ${r.leftoverParts?.join(',') || '-'}, похожих на указания мест ${r.inject}, пустых кусков <r> ${r.emptyRuns ?? '-'}`);
  if (r.fail.length) console.log('     ' + r.fail.join('; '));
}
const ok = rows.filter(r => r.ok).length;
console.log(`пустых кусков <r> во всех xlsx: ${rows.reduce((a, r) => a + (r.emptyRuns || 0), 0)}`);
console.log(`\nИТОГО: ${ok} файлов из ${rows.length} прошли (docx ${rows.filter(r => r.ok && r.ext === 'docx').length}/${rows.filter(r => r.ext === 'docx').length}, xlsx ${rows.filter(r => r.ok && r.ext === 'xlsx').length}/${rows.filter(r => r.ext === 'xlsx').length})`);
fs.writeFileSync(path.join(outDir, 'real-check.json'), JSON.stringify(rows, null, 1));
