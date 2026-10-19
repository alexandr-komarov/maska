// Прогон Маски на тестовом наборе: достаёт mask() и leaks() из index.html и меряет,
// что спрятано из обязательного (must) и что задето из того, что трогать нельзя (keep).
// Запуск: node harness.mjs [путь к index.html или -] [cases|holdout]
import fs from 'node:fs';
import { loadMask } from './mask-api.mjs';
const api = loadMask(process.argv[2] && process.argv[2] !== '-' ? process.argv[2] : new URL('../index.html', import.meta.url));
const set = process.argv[3] || 'cases';
const cases = JSON.parse(fs.readFileSync(new URL(`./${set}.json`, import.meta.url), 'utf8'));
const dir = set === 'cases' ? './docs/' : `./${set}/`;
let mustAll = 0, mustOk = 0, keepAll = 0, keepOk = 0;
for (const c of cases) {
  const text = fs.readFileSync(new URL(dir + c.file, import.meta.url), 'utf8');
  const r = api.mask(text, []);
  const miss = c.must.filter(s => r.masked.includes(s));
  const broken = c.keep.filter(s => !r.masked.includes(s));
  // тип метки: что за данные модель увидит на месте значения
  for (const [val, type] of Object.entries(c.types || {})) {
    const h = r.hits.find(x => x.value === val);
    if (!h || h.t !== type) { broken.push(`${val} должно быть ${type}, стало ${h ? h.t : 'не найдено'}`); keepAll++; } else { keepAll++; keepOk++; }
  }
  mustAll += c.must.length; mustOk += c.must.length - miss.length;
  keepAll += c.keep.length; keepOk += c.keep.length - broken.length;
  console.log(`\n${c.file}: спрятано ${c.must.length - miss.length}/${c.must.length}, лишнего задето ${broken.length}/${c.keep.length}`);
  for (const s of miss) console.log('  ПРОПУЩЕНО:', s);
  for (const s of broken) console.log('  ЗАДЕТО ЛИШНЕЕ:', s);
}
console.log(`\nИТОГО: спрятано ${mustOk}/${mustAll} (${(100 * mustOk / Math.max(mustAll,1)).toFixed(1)} %), не задето ${keepOk}/${keepAll}`);
