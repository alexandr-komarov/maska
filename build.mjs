// Сборка одного файла: код src/*.mjs вставляется в index.html между маркерами.
// Каждый модуль оборачивается в функцию, import становится разбором объекта предыдущего модуля,
// export снимается. Источник правды для чтения и возврата - src/, руками блок в index.html не правим.
// Запуск: node build.mjs          пересобрать index.html
//         node build.mjs --check  только проверить, что блок в index.html совпадает с src/
import fs from 'node:fs';

export const ORDER = ['safe', 'zip', 'ooxml', 'xlsx-return', 'docx-return', 'office'];
// модули, чьи export зовёт страница: сейф из правил и интерфейса, office из интерфейса
export const API = ['safe', 'office'];
export const START = '// ==== модули из src/: собирает build.mjs, руками не править ====';
export const END = '// ==== конец модулей из src/ ====';
const root = new URL('.', import.meta.url);
const ns = n => '__' + n.replace(/-/g, '_');

// Один модуль: функция-обёртка, возвращает объект своих export
export function wrapModule(name, code) {
  const exported = [];
  let body = code.replace(/^import\s*\{([^}]*)\}\s*from\s*'\.\/([\w-]+)\.mjs';?[ \t]*$/gm, (m, names, from) => {
    if (!ORDER.includes(from) || ORDER.indexOf(from) >= ORDER.indexOf(name)) throw new Error(`${name}: импорт из ${from} не по порядку сборки`);
    const list = names.split(',').map(s => s.trim()).filter(Boolean).map(s => s.replace(/\s+as\s+/, ': '));
    return `const { ${list.join(', ')} } = ${ns(from)};`;
  });
  body = body.replace(/^export\s+((?:async\s+)?function\s*\*?\s*|const\s+|let\s+|class\s+)([\w$]+)/gm, (m, kw, id) => { exported.push(id); return kw + id; });
  if (/^\s*(import|export)\b/m.test(body)) throw new Error(`${name}: import или export, который сборка не понимает`);
  if (/<\/script/i.test(body)) throw new Error(`${name}: в коде есть </script>, он оборвёт страницу`);
  return `const ${ns(name)} = (() => {\n${body.trimEnd()}\nreturn { ${exported.join(', ')} };\n})();`;
}

// Блок целиком: модули по порядку и имена из safe.mjs и office.mjs, которые зовут правила и интерфейс
export function buildBlock(read = n => fs.readFileSync(new URL(`src/${n}.mjs`, root), 'utf8')) {
  const parts = ORDER.map(n => wrapModule(n, read(n)));
  const names = n => [...read(n).matchAll(/^export\s+(?:async\s+)?(?:function|const)\s+([\w$]+)/gm)].map(m => m[1]);
  const api = API.map(n => `const { ${names(n).join(', ')} } = ${ns(n)};`);
  return `${START}\n${parts.join('\n')}\n${api.join('\n')}\n${END}`;
}

export function assemble(html, block = buildBlock()) {
  const a = html.indexOf(START), b = html.indexOf(END);
  if (a < 0 || b < a) throw new Error('в index.html нет маркеров блока модулей');
  return html.slice(0, a) + block + html.slice(b + END.length);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const file = new URL('index.html', root), html = fs.readFileSync(file, 'utf8'), out = assemble(html);
  if (process.argv.includes('--check')) {
    if (out !== html) { console.log('index.html не собран из текущих src/: запустите node build.mjs'); process.exit(1); }
    console.log('index.html совпадает с src/');
  } else if (out !== html) { fs.writeFileSync(file, out); console.log('index.html пересобран'); }
  else console.log('index.html уже собран');
}
