// Объём сомнительного: сколько строк окна проверки человек увидит на каждом документе набора.
// Группа это одно слово или число во всех написаниях, одна строка окна и один щелчок.
// Заодно счёт разметки с включённым правилом «сомнительное спрятано»: must спрятано, keep не задето.
// Запуск: node tests/doubt-volume.mjs [--list]   с --list печатает сами слова
import fs from 'node:fs';
import { loadMask } from './mask-api.mjs';
import { readOffice, hideSums } from '../src/office.mjs';
const api = loadMask(), list = process.argv.includes('--list');
const DOUBT = { NAME: 'слова', NUM: 'числа', AT: 'с @' };
const SETS = [['cases', 'docs'], ['holdout', 'holdout'], ['blind', 'blind'], ['control', 'control'], ['files', 'files'], ['deal', 'deal']];
for (const [set, dir] of SETS) {
  const j = JSON.parse(fs.readFileSync(new URL(`./${set}.json`, import.meta.url), 'utf8')), marks = Array.isArray(j) ? j : j.files;
  let groups = 0, places = 0, max = 0, mustAll = 0, mustOk = 0, keepAll = 0, keepOk = 0, keepBase = 0;
  console.log(`\n== ${set}`);
  for (const c of marks) {
    const name = c.file, bytes = new Uint8Array(fs.readFileSync(new URL(`./${dir}/${name}`, import.meta.url)));
    const text = (await readOffice(name, bytes)).text, opts = { sums: hideSums(name) };
    const base = api.mask(text, [], opts), r = api.mask(text, [], { ...opts, doubt: true });
    const dz = r.vault.filter(v => DOUBT[v.t]), n = r.hits.filter(h => DOUBT[h.t]).length;
    const by = {}; for (const v of dz) by[DOUBT[v.t]] = (by[DOUBT[v.t]] || 0) + 1;
    const must = [...(c.must || []), ...Object.keys(c.must_types || {})], keep = c.keep || [];
    const broken = keep.filter(s => !r.masked.includes(s));
    mustAll += must.length; mustOk += must.filter(s => !r.masked.includes(s)).length;
    keepAll += keep.length; keepOk += keep.length - broken.length; keepBase += keep.filter(s => base.masked.includes(s)).length;
    groups += dz.length; places += n; max = Math.max(max, dz.length);
    console.log(`${name}: знаков ${text.length}, уверенных находок ${base.hits.length}, сомнительных строк ${dz.length} (мест ${n})` +
      (dz.length ? ': ' + Object.entries(by).map(([k, v]) => `${k} ${v}`).join(', ') : ''));
    if (list) { for (const v of dz) console.log(`    ${v.t} ${JSON.stringify([...v.forms])}`); for (const s of broken) console.log(`    ЗАДЕТО: ${s}`); }
  }
  console.log(`ИТОГО ${set}: файлов ${marks.length}, сомнительных строк ${groups} (мест ${places}), самое большое окно ${max}; ` +
    `с сомнительным спрятано ${mustOk}/${mustAll}, не задето ${keepOk}/${keepAll} (без сомнительного ${keepBase}/${keepAll})`);
}
