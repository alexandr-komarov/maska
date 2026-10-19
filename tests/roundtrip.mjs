// Сквозная проверка как в браузере: файл → чтение → mask() из index.html → замена → чтение результата.
// В результате не должно остаться ни одного найденного значения (имена листов считаются отдельно: они не меняются).
// Запуск: node tests/roundtrip.mjs <папка или файл>... [--counts] [--out папка]
//   --counts  печатать только счётчики, без имён файлов и значений (для настоящих файлов с чужими данными)
//   --out     куда писать результаты (по умолчанию не пишутся; для настоящих файлов только out/)
//   --session все файлы одним сейфом по порядку, как их обезличивает человек; без него сейф на каждом файле новый
import fs from 'node:fs';
import path from 'node:path';
import { loadMask } from './mask-api.mjs';
import { readOffice, buildReturn, hideSums } from '../src/office.mjs';
import { safeNew } from '../src/safe.mjs';

const EXT = /\.(docx|xlsx|csv|md|txt)$/i;
export function listFiles(p) {
  if (fs.statSync(p).isFile()) return [p];
  const out = [];
  (function walk(d) { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const q = path.join(d, e.name); if (e.isDirectory()) walk(q); else if (EXT.test(e.name)) out.push(q); } })(p);
  return out.sort();
}

// Один файл: счётчики и список оставшихся значений
// safe: сейф перед этим файлом; обновлённый сейф в row.safe
export async function roundtrip(file, api = loadMask(), safe = safeNew()) {
  const bytes = new Uint8Array(fs.readFileSync(file));
  const read = await readOffice(file, bytes);
  const r = api.mask(read.text, [], { sums: hideSums(file), safe });
  const out = await buildReturn({ name: path.basename(file), bytes, read, hits: r.hits, vault: r.vault, masked: r.masked });
  const row = { ext: read.kind === 'text' ? read.ext : read.kind, hits: r.hits.length, values: new Set(r.hits.map(h => h.value)).size, left: out.left || [], warnings: out.warnings.length, out, safe: r.safe };
  if (read.kind === 'text') {
    // текст отдаётся как есть: проверяем, что в нём нет найденных значений
    const txt = new TextDecoder().decode(out.bytes);
    row.left = [...new Set(r.hits.map(h => h.value))].filter(v => txt.replace(/\r\n/g, '\n').includes(v));
  } else {
    row.edits = [...out.plan.edits.values()].reduce((a, x) => a + x.length, 0);
    row.converted = out.report.converted.length;
    row.sheetNames = out.report.notHandled.filter(x => x.kind === 'sheetName').length;
    row.replaced = out.report.total;
  }
  return row;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const argv = process.argv.slice(2), counts = argv.includes('--counts'), session = argv.includes('--session');
  let safe = safeNew();
  const oi = argv.indexOf('--out'), outDir = oi >= 0 ? argv[oi + 1] : null;
  const dirs = argv.filter((a, i) => !a.startsWith('--') && argv[i - 1] !== '--out');
  if (!dirs.length) { console.log('Запуск: node tests/roundtrip.mjs <папка или файл>... [--counts] [--out папка]'); process.exit(2); }
  const api = loadMask();
  let ok = 0, all = 0, hits = 0, edits = 0, conv = 0, sheet = 0, left = 0, errors = 0, warn = 0;
  const byExt = {};
  for (const d of dirs) for (const [k, f] of listFiles(d).entries()) {
    all++;
    const label = counts ? `файл ${all}` : path.relative(process.cwd(), f);
    let row;
    try { row = await roundtrip(f, api, session ? safe : safeNew()); if (session) safe = row.safe; }
    catch (e) { errors++; console.log(`ОШИБКА ${label}: ${counts ? e.message.split(':')[0].slice(0, 80) : e.message}`); continue; }
    const b = byExt[row.ext] ||= { ok: 0, all: 0 }; b.all++;
    hits += row.hits; edits += row.edits || 0; conv += row.converted || 0; sheet += row.sheetNames || 0; left += row.left.length; warn += row.warnings;
    if (!row.left.length) { ok++; b.ok++; }
    if (outDir) { fs.mkdirSync(outDir, { recursive: true }); fs.writeFileSync(path.join(outDir, `${counts ? 'file-' + all : path.basename(f, path.extname(f))}-maska${path.extname(f)}`), row.out.bytes); }
    const line = `${row.left.length ? 'ОСТАЛОСЬ' : 'OK'}  ${label} (${row.ext}): находок ${row.hits}, правок по месту ${row.edits ?? '-'}, чисел в строку ${row.converted ?? '-'}, в именах листов ${row.sheetNames ?? '-'}, предупреждений ${row.warnings}, осталось ${row.left.length}`;
    console.log(line);
    if (!counts) for (const v of row.left) console.log('   осталось:', JSON.stringify(v));
  }
  if (session) console.log(`\nСеанс с одним сейфом: записей в сейфе ${safe.records.length}`);
  console.log(`\nИТОГО: без остатка ${ok} из ${all} (${Object.entries(byExt).map(([e, b]) => `${e} ${b.ok}/${b.all}`).join(', ')}), ошибок чтения ${errors}; находок ${hits}, правок по месту ${edits}, чисел в строку ${conv}, значений в именах листов ${sheet}, осталось значений ${left}, предупреждений ${warn}`);
  process.exit(ok === all ? 0 : 1);
}
