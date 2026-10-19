// mask() и leaks() из index.html, как их видит браузер: кусок от const LABEL до function render().
// Третий параметр mask(text, terms, opts) задаёт {sums}: для таблиц суммы не прячутся, см. hideSums в src/office.mjs.
// opts.safe: сейф. Функции сейфа в странице приходят из блока модулей, здесь из src/safe.mjs (блок сверяет page.test.mjs)
import fs from 'node:fs';
import * as safe from '../src/safe.mjs';
export const rulesOf = src => src.slice(src.indexOf('const LABEL'), src.indexOf('function render('));
export function loadMask(indexPath = new URL('../index.html', import.meta.url)) {
  const names = Object.keys(safe);
  return new Function(...names, rulesOf(fs.readFileSync(indexPath, 'utf8')) + '; return {mask, leaks};')(...names.map(n => safe[n]));
}
