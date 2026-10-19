// Страница целиком: сборка не забыта, встроенный код работает как src/, сетевых адресов нет,
// интерфейс возврата на заглушке DOM (браузер здесь не запускается).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { assemble, buildBlock, START, END } from '../build.mjs';
import { readOffice, buildReturn } from '../src/office.mjs';
import { loadMask } from './mask-api.mjs';

const html = fs.readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const script = html.slice(html.indexOf('<script>\n') + 9, html.lastIndexOf('</script>'));
const fx = n => new Uint8Array(fs.readFileSync(new URL(n, import.meta.url)));

test('сборка: блок модулей в index.html совпадает с src/', () => {
  assert.equal(assemble(html), html, 'запустите node build.mjs');
  assert.equal(html.split(START).length, 2);
  assert.ok(html.includes(buildBlock()));
  assert.ok(!/async function (unzip|readDocx|readXlsx)\b/.test(html), 'старые функции чтения убраны');
});

test('сборка: встроенный код даёт те же байты, что src/', async () => {
  const block = html.slice(html.indexOf(START), html.indexOf(END));
  const page = new Function(block + '; return { readOffice, buildReturn };')();
  const api = loadMask();
  for (const f of ['./fixtures/synthetic.docx', './fixtures/synthetic.xlsx', './deal/osv-60.xlsx']) {
    const bytes = fx(f), name = f.split('/').pop();
    const a = await page.readOffice(name, bytes), b = await readOffice(name, bytes);
    assert.equal(a.text, b.text, f);
    const m = api.mask(a.text, [], { sums: false });
    const ra = await page.buildReturn({ name, bytes, read: a, hits: m.hits, vault: m.vault, masked: m.masked });
    const rb = await buildReturn({ name, bytes, read: b, hits: m.hits, vault: m.vault, masked: m.masked });
    assert.deepEqual(ra.bytes, rb.bytes, f);
    assert.deepEqual(ra.warnings, rb.warnings, f);
  }
});

test('без сети: нет внешних адресов в src, href, fetch, import', () => {
  const ext = /^(?:https?:)?\/\//i;
  const found = [];
  for (const m of html.matchAll(/<(\w+)\b([^>]*)>/g)) {
    const [, tag, attrs] = m;
    for (const a of attrs.matchAll(/\b(src|href|srcset|action|data|poster)\s*=\s*["']([^"']*)["']/gi)) {
      const navigation = tag.toLowerCase() === 'a' || /rel=["']canonical["']/i.test(attrs);
      // ссылки, по которым человек переходит сам, и canonical браузер не загружает
      if (navigation && a[1].toLowerCase() === 'href') continue;
      if (ext.test(a[2]) || a[2].startsWith('/')) found.push(`${tag} ${a[1]}=${a[2]}`);
    }
  }
  assert.deepEqual(found, []);
  assert.ok(!/<link[^>]*rel=["']?(stylesheet|preload|prefetch|modulepreload|icon|dns-prefetch|preconnect)/i.test(html), 'нет загружаемых link');
  assert.ok(!/\bfetch\s*\(|XMLHttpRequest|WebSocket|EventSource|sendBeacon|importScripts|\bimport\s*\(|^\s*import\s/m.test(script), 'нет сетевых вызовов и import');
  assert.ok(!/url\(\s*["']?(?:https?:)?\/\//i.test(html), 'нет внешних url() в стилях');
  assert.ok(!/@import/i.test(html), 'нет @import в стилях');
  // запрет сети в самом браузере: CSP в первых строках head, до любого скрипта
  const csp = html.match(/<meta http-equiv="Content-Security-Policy" content="([^"]*)">/);
  assert.ok(csp && csp.index < html.indexOf('<script'), 'CSP в head до скриптов');
  const dir = Object.fromEntries(csp[1].split(';').map(d => d.trim().split(/\s+/)).map(([k, ...v]) => [k, v.join(' ')]));
  for (const k of ['default-src', 'connect-src']) assert.equal(dir[k], "'none'", k);
  assert.equal(dir['form-action'], "'none'");
  assert.ok(!/https?:|\*/.test(csp[1]), 'в CSP нет разрешённых адресов');
});

// Заглушка DOM ровно под то, что зовёт страница
// clipboard: undefined оставляет страницу без navigator.clipboard, как бывает при открытии с диска
function page(o = {}) {
  const els = {}, saved = [], copied = [], execCopy = o.execCopy;
  const clipboard = 'clipboard' in o ? o.clipboard : { writeText: async t => { copied.push(t); } };
  const el = id => els[id] ||= { id, value: '', textContent: '', innerHTML: '', hidden: false, disabled: false, checked: id === 'sums', dataset: {},
    classList: { add() {}, remove() {}, toggle() {} }, scrollIntoView() {}, select() {}, getBoundingClientRect: () => ({ top: 0 }), addEventListener() {} };
  // начальное состояние разметки: кнопка возврата скрыта и выключена, блок предупреждений скрыт
  Object.assign(el('dlFile'), { hidden: true, disabled: true }); el('retWarn').hidden = true; el('out').hidden = true;
  const steps = ['s-doc', 's-check', 's-after', 's-answer'].map(go => ({ ...el('step-' + go), dataset: { go } }));
  // запасной путь копирования: временное поле и execCommand('copy'), execCopy задаёт его ответ
  let held = null;
  const document = { getElementById: el, querySelectorAll: () => steps, addEventListener() {}, body: { scrollHeight: 0, appendChild(n) { held = n; } },
    execCommand: cmd => { if (cmd !== 'copy' || !execCopy) return false; copied.push(held.value); return true; },
    createElement: () => { const a = { style: {}, setAttribute() {}, select() {}, remove() {}, click() { saved.push({ name: a.download }); } }; return a; } };
  const URL_ = { createObjectURL: b => { document.lastBlob = b; return 'blob:x'; }, revokeObjectURL() {} };
  const fn = new Function('document', 'addEventListener', 'innerHeight', 'scrollY', 'navigator', 'URL',
    script + '\n;return { loadFile, loadSafe, prepareReturn, get state() { return state; }, get ready() { return ready; }, get loaded() { return loaded; }, get safe() { return curSafe; } };');
  const api = fn(document, () => {}, 800, 0, { clipboard }, URL_);
  const file = (name, bytes) => (typeof bytes === 'string' ? { name, size: bytes.length, text: async () => bytes }
    : { name, size: bytes.length, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });
  const click = async id => { await els[id].onclick?.({ target: els[id] }); };
  const download = async (id = 'dlFile') => { await click(id); const b = document.lastBlob; return { name: saved.at(-1)?.name, bytes: new Uint8Array(await b.arrayBuffer()) }; };
  const input = (id, value) => { els[id].value = value; els[id].oninput?.({ target: els[id] }); };
  return { api, el, file, click, download, saved, input, copied };
}

test('интерфейс: docx загружается, маскируется, кнопка отдаёт файл в исходном формате', async () => {
  const p = page();
  await p.api.loadFile(p.file('dogovor.docx', fx('./deal/dogovor-podryada.docx')));
  await p.api.prepareReturn();
  assert.match(p.el('fileinfo').textContent, /^dogovor\.docx · docx:/);
  assert.equal(p.el('src').value, (await readOffice('d.docx', fx('./deal/dogovor-podryada.docx'))).text);
  assert.equal(p.el('sums').checked, true, 'для текста суммы прячутся');
  assert.equal(p.el('dlFile').hidden, false);
  assert.equal(p.el('dlFile').disabled, true, 'до окна проверки файл закрыт');
  await p.click('okReview');
  assert.equal(p.el('dlFile').disabled, false);
  assert.ok(p.api.ready && p.api.ready.left.length === 0);
  const d = await p.download();
  assert.equal(p.saved.at(-1).name, 'dogovor-maska.docx');
  const back = await readOffice('x.docx', d.bytes);
  for (const h of p.api.state.hits) assert.ok(!back.text.includes(h.value), h.value);
});

test('интерфейс: xlsx, предупреждения списком, суммы по умолчанию выключены', async () => {
  const p = page();
  await p.api.loadFile(p.file('kniga.xlsx', fx('./fixtures/synthetic.xlsx')));
  await p.api.prepareReturn();
  assert.equal(p.el('sums').checked, false, 'в таблицах суммы по умолчанию не прячутся');
  assert.equal(p.el('dlFile').hidden, false);
  assert.equal(p.el('retWarn').hidden, false);
  // предупреждение о длинном числе идёт от чтения и не зависит от правил маски
  assert.match(p.el('retWarn').innerHTML, /<li>В 1 ячейке длинное число, больше 15 цифр \(Клиенты!D2\)/);
  await p.click('okReview');
  assert.equal((await p.download()).bytes[0], 0x50, 'zip');
});

test('интерфейс: текст изменён после загрузки, старый формат, пример', async () => {
  const p = page();
  await p.api.loadFile(p.file('a.xlsx', fx('./fixtures/synthetic.xlsx')));
  p.el('src').value += '\nдописано руками';
  await p.click('run'); await p.api.prepareReturn();
  assert.equal(p.el('dlFile').hidden, true);
  assert.match(p.el('retInfo').textContent, /Текст изменён после загрузки файла/);
  await p.api.loadFile(p.file('Старый.doc', new Uint8Array([0xD0, 0xCF, 0x11, 0xE0])));
  assert.match(p.el('fileinfo').textContent, /старого формата Word \(\.doc\).*сохраните как \.docx/);
  assert.equal(p.el('dlFile').hidden, true);
  // в сейфе уже есть записи человека: пример сначала спрашивает
  await p.click('sample');
  assert.equal(p.el('safeAsk').hidden, false);
  await p.click('safeYes');
  assert.equal(p.el('dlFile').hidden, true, 'для примера файла нет');
});

test('интерфейс: csv возвращается в csv', async () => {
  const p = page();
  await p.api.loadFile(p.file('spisok.csv', new TextEncoder().encode('ФИО;Телефон\nОрлов Олег Петрович;+7 916 245-18-73\n')));
  await p.api.prepareReturn();
  await p.click('okReview');
  const d = await p.download();
  assert.equal(p.saved.at(-1).name, 'spisok-maska.csv');
  const txt = new TextDecoder().decode(d.bytes);
  assert.ok(txt.includes('[ФИО_1]') && !txt.includes('Орлов') && !txt.includes('245-18-73'));
});

const safeJson = (name, records, words = []) => JSON.stringify({ format: 'maska-safe', version: 1, name, data: { words, records } });

test('сейф: не загружен, предупреждение; после документа напоминание сохранить; сохранение файлом', async () => {
  const p = page();
  assert.match(p.el('safeInfo').innerHTML, /^не загружен\..*Сейф это файл на вашем компьютере.*первый документ клиента/);
  assert.equal(p.el('s-safe').className, 'safe warn');
  await p.api.loadFile(p.file('dogovor.docx', fx('./deal/dogovor-podryada.docx')));
  assert.equal(p.el('safeRemind').hidden, false);
  assert.match(p.el('safeRemind').innerHTML, /<b>Сохраните сейф<\/b>Новых записей в сейфе: \d+\..*Сейф не был открыт/);
  assert.match(p.el('safeInfo').innerHTML, /^новый, без названия · записей \d+ · своих слов 0 · изменён .* · <b>не сохранён<\/b>$/);
  p.input('safeName', 'Каменный Мост, подряд');
  const d = await p.download('safeSave');
  assert.equal(d.name, 'Каменный Мост, подряд.сейф.json');
  const j = JSON.parse(new TextDecoder().decode(d.bytes));
  assert.equal(j.format, 'maska-safe');
  assert.equal(j.data.records.length, p.api.state.vault.length);
  assert.ok(j.data.records.some(r => r.tag === '[ОРГ_1]' && r.forms.includes('ООО «Каменный Мост»')));
  assert.equal(p.el('safeRemind').hidden, true, 'после сохранения напоминания нет');
  assert.match(p.el('safeInfo').innerHTML, /«Каменный Мост, подряд» · записей \d+ .* · сохранён$/);
});

test('сейф: договор и акт по очереди, одна фирма с одной меткой, ответ нейросети возвращается', async () => {
  const p = page();
  await p.api.loadFile(p.file('dogovor.docx', fx('./deal/dogovor-podryada.docx')));
  const tag = p.api.state.hits.find(h => h.value === 'ООО «Каменный Мост»').tag;
  await p.api.loadFile(p.file('akt.docx', fx('./deal/akt-sdachi-priemki.docx')));
  assert.equal(p.api.state.hits.find(h => h.value === 'ООО «Каменный Мост»').tag, tag);
  assert.equal(p.api.state.hits.find(h => h.value === 'Северный Ветер Строй ООО').tag, p.api.safe.records.find(r => r.forms.includes('ООО «Северный Ветер Строй»')).tag);
  p.el('answer').value = `Подрядчик ${tag.toLowerCase()} выполнил работы.`;
  await p.click('restore');
  assert.equal(p.el('copyRestored').dataset.t, `Подрядчик ${p.api.safe.records.find(r => r.tag === tag).value} выполнил работы.`);
  assert.match(p.api.safe.records.find(r => r.tag === tag).value, /Каменный Мост/);
  assert.match(p.el('restoreInfo').textContent, /^подставлено: 1$/);
});

test('сейф: битый файл не трогает текущий, другой сейф заменяет только после подтверждения на странице', async () => {
  const p = page();
  await p.api.loadSafe(p.file('a.сейф.json', safeJson('Клиент А', [{ tag: '[ОРГ_1]', type: 'ORG', value: 'ООО «Альфа»', forms: ['ООО «Альфа»'], keys: ['альфа'] }], ['Север-2'])));
  assert.match(p.el('safeMsg').textContent, /^Открыт сейф «Клиент А»: записей 1, своих слов 1\.$/);
  assert.equal(p.el('terms').value, 'Север-2', 'свои слова сейфа в поле');
  assert.equal(p.el('safeAsk').hidden, true, 'первый сейф открывается без вопроса');
  const before = p.api.safe;
  for (const [name, text] of [['b.json', '{"format":"maska-safe"'], ['c.json', '{"items":[]}'], ['d.json', safeJson('x', [{ tag: '[ОРГ_1]' }])]]) {
    await p.api.loadSafe(p.file(name, text));
    assert.match(p.el('safeMsg').textContent, new RegExp(`^Не удалось открыть сейф «${name}»: .*Текущий сейф не изменился\\.$`));
    assert.equal(p.el('safeMsg').className, 'check warn');
    assert.equal(p.api.safe, before);
  }
  // документ с сейфом А, потом попытка открыть сейф Б
  p.el('src').value = 'ООО «Омега» и ООО «Альфа», объект Север-2';
  await p.click('run');
  assert.equal(p.api.state.masked, '[ОРГ_2] и [ОРГ_1], объект [СЛОВО_1]');
  await p.api.loadSafe(p.file('b.сейф.json', safeJson('Клиент Б', [{ tag: '[ОРГ_1]', type: 'ORG', value: 'ООО «Омега»', forms: ['ООО «Омега»'], keys: ['омега'] }])));
  assert.equal(p.el('safeAsk').hidden, false);
  assert.match(p.el('safeAskText').textContent, /^Заменить сейф «Клиент А» \(записей 3\) на «Клиент Б» \(записей 1\)\? Текущий сейф не сохранён/);
  assert.equal(p.api.safe.name, 'Клиент А', 'до подтверждения сейф прежний');
  await p.click('safeNo');
  assert.equal(p.el('safeAsk').hidden, true);
  assert.equal(p.api.safe.name, 'Клиент А');
  await p.api.loadSafe(p.file('b.сейф.json', safeJson('Клиент Б', [{ tag: '[ОРГ_1]', type: 'ORG', value: 'ООО «Омега»', forms: ['ООО «Омега»'], keys: ['омега'] }])));
  await p.click('safeYes');
  assert.equal(p.api.safe.name, 'Клиент Б');
  // «Север-2» больше не своё слово сейфа А: теперь это сомнительное слово с заглавной, оно спрятано другой меткой
  assert.equal(p.api.state.masked, '[ОРГ_1] и [ОРГ_2], объект [НАЗВАНИЕ_1]', 'документ на экране перемаскирован с сейфом Б, слова А не перешли');
  assert.ok(!p.api.safe.records.some(r => r.value === 'ООО «Альфа»' && r.tag === '[ОРГ_1]'));
  // новый пустой сейф тоже через подтверждение
  await p.click('safeNew');
  assert.match(p.el('safeAskText').textContent, /на новый пустой \(записей 0\)/);
});

test('сейф: файл .json через кнопку документа и перетаскивание открывается как сейф', async () => {
  const p = page();
  await p.el('file').onchange({ target: { files: [p.file('x.json', safeJson('Дело 7', []))] } });
  await new Promise(r => setTimeout(r, 0));
  assert.match(p.el('safeMsg').textContent, /^Открыт сейф «Дело 7»/);
  assert.equal(p.el('fileinfo').textContent, '', 'как документ не читался');
});

test('сейф: значения не пишутся в хранилища браузера', () => {
  assert.ok(!/localStorage|sessionStorage|indexedDB|document\.cookie|caches\.|navigator\.storage/.test(script));
});

// Пример на странице: договор и акт с одним поставщиком, свой сейф «Пример»
const SAMPLE = { org: ['ООО «Ромашка»', 'Ромашка ООО'], buyer: 'АО «Камнегорский завод»', inn: ['0047020267', '0047020316'], num: '47/ЛГ-2026',
  people: ['Иванова Петра Сергеевича', 'Иванов П.С.', 'Соколовой Елены Викторовны', 'Соколова Е.В.'], sums: ['4 180 000 ₽', '4 180 000 руб.'] };
const innOk = s => { const d = [...s].map(Number); return [2, 4, 10, 3, 5, 9, 4, 6, 8].reduce((a, k, i) => a + k * d[i], 0) % 11 % 10 === d[9]; };
const tagFor = (p, value) => p.api.state.hits.find(h => h.value === value)?.tag;

test('пример: договор и акт подряд с одним сейфом, поставщик под одной меткой, исходных значений не осталось', async () => {
  const p = page();
  assert.ok(!html.includes('заявка абонента') && !html.includes('Северный терминал'), 'старый пример убран');
  assert.ok(SAMPLE.inn.every(innOk), 'ИНН в примере с верной контрольной суммой');
  await p.click('sample');
  const dog = p.api.state, src1 = p.el('src').value;
  assert.match(src1, /^ДОГОВОР ПОСТАВКИ/);
  assert.equal(p.api.safe.name, 'Пример');
  assert.equal(tagFor(p, 'ООО «Ромашка»'), '[ОРГ_1]');
  assert.equal(tagFor(p, 'АО «Камнегорский завод»'), '[ОРГ_2]');
  assert.equal(tagFor(p, 'Иванова Петра Сергеевича'), '[ФИО_1]');
  assert.equal(tagFor(p, SAMPLE.num), '[ДОГОВОР_1]');
  assert.equal(tagFor(p, '4 180 000 ₽'), '[СУММА_1]');
  assert.equal(p.el('s-sample').hidden, false);
  assert.match(p.el('sampleTitle').textContent, /шаг 1 из 2/);
  assert.match(p.el('sampleText').textContent, /получил метку \[ОРГ_1\]/);
  assert.equal(p.el('safeRemind').hidden, true, 'сейф примера сохранять не просим');
  assert.match(p.el('safeInfo').innerHTML, /^сейф примера/);

  await p.click('sampleAkt');
  const akt = p.api.state;
  assert.match(p.el('src').value, /^АКТ ПРИЁМА-ПЕРЕДАЧИ/);
  assert.equal(akt.added, 0, 'акт не добавил в сейф ни одной записи: всё узнано по договору');
  assert.equal(tagFor(p, 'Ромашка ООО'), '[ОРГ_1]', 'другое написание поставщика, та же метка');
  assert.equal(tagFor(p, 'АО «Камнегорский завод»'), '[ОРГ_2]');
  assert.equal(tagFor(p, SAMPLE.inn[0]), dog.hits.find(h => h.value === SAMPLE.inn[0]).tag);
  assert.equal(tagFor(p, 'Иванов П.С.'), '[ФИО_1]');
  assert.equal(tagFor(p, 'Соколова Е.В.'), '[ФИО_2]');
  assert.equal(tagFor(p, SAMPLE.num), '[ДОГОВОР_1]');
  assert.equal(tagFor(p, '4 180 000 руб.'), '[СУММА_1]');
  assert.match(p.el('sampleText').textContent, /«Ромашка ООО», но метка у него та же, \[ОРГ_1\]/);
  for (const r of [dog, akt]) {
    assert.ok(r.masked.includes('[ОРГ_1]') && r.masked.includes('[ИНН_1]') && r.masked.includes('[ФИО_1]') && r.masked.includes('[ДОГОВОР_1]') && r.masked.includes('[СУММА_1]'));
    for (const v of [...SAMPLE.org, SAMPLE.buyer, ...SAMPLE.inn, SAMPLE.num, ...SAMPLE.people, ...SAMPLE.sums, 'Ромашка', 'Иванов', 'Соколов', '4 180 000', '00470'])
      assert.ok(!r.masked.includes(v), `осталось в тексте с метками: ${v}`);
  }
  // оба текста помещаются на экран
  for (const t of [src1, p.el('src').value]) assert.ok(t.split('\n').length <= 20 && t.length < 1200);
});

test('пример: сейф человека не трогается без согласия, свой документ не попадает в сейф примера', async () => {
  const p = page();
  await p.api.loadSafe(p.file('a.сейф.json', safeJson('Клиент А', [{ tag: '[ОРГ_1]', type: 'ORG', value: 'ООО «Альфа»', forms: ['ООО «Альфа»'], keys: ['альфа'] }])));
  const mine = p.api.safe;
  await p.click('sample');
  assert.equal(p.el('safeAsk').hidden, false, 'пример спрашивает, прежде чем занять место сейфа');
  assert.match(p.el('safeAskText').textContent, /^Пример откроется в отдельном сейфе.*Заменить сейф «Клиент А» \(записей 1\) на сейф примера/);
  assert.equal(p.api.safe, mine);
  assert.equal(p.el('src').value, '', 'до согласия примера на экране нет');
  await p.click('safeNo');
  assert.equal(p.api.safe, mine, 'отказ: сейф прежний');
  await p.click('sample'); await p.click('safeYes');
  assert.equal(p.api.safe.name, 'Пример');
  assert.ok(!p.api.safe.records.some(r => /Альфа/.test(r.value)), 'записи человека в сейф примера не перешли');
  assert.deepEqual(mine.records.map(r => r.value), ['ООО «Альфа»'], 'сейф человека не изменился');

  // человек возвращает свой сейф: пример уходит с экрана и в сейф не попадает
  await p.api.loadSafe(p.file('a.сейф.json', safeJson('Клиент А', [{ tag: '[ОРГ_1]', type: 'ORG', value: 'ООО «Альфа»', forms: ['ООО «Альфа»'], keys: ['альфа'] }])));
  assert.equal(p.el('safeAsk').hidden, true, 'сейф примера заменяется без вопроса');
  assert.equal(p.api.safe.name, 'Клиент А');
  assert.equal(p.api.safe.records.length, 1);
  assert.equal(p.el('src').value, '');
  assert.equal(p.el('s-sample').hidden, true);
  assert.equal(p.api.state, null);

  // после примера свой файл: сейф примера убран, метки с единицы
  const q = page();
  await q.click('sample'); await q.click('sampleAkt');
  await q.api.loadFile(q.file('spisok.csv', new TextEncoder().encode('Контрагент\nООО «Вектор»\n')));
  assert.equal(q.api.safe.name, '');
  assert.deepEqual(q.api.safe.records.map(r => [r.tag, r.value]), [['[ОРГ_1]', 'ООО «Вектор»']]);
  // свой текст в поле вместо примера
  const w = page();
  await w.click('sample');
  w.input('src', 'Письмо от ООО «Вектор»');
  assert.equal(w.api.safe.records.length, 0);
  assert.equal(w.el('out').hidden, true);
  await w.click('run');
  assert.deepEqual(w.api.safe.records.map(r => [r.tag, r.value]), [['[ОРГ_1]', 'ООО «Вектор»']]);
});

const BANNED = /—|→|←|≈|\+|&|\bvs\b/;
test('просьба для нейросети: метки текущего документа, правила, строка для задачи', async () => {
  const p = page();
  await p.click('copyPrompt');
  const common = p.copied.at(-1);
  assert.match(common, /Чаще всего встречаются такие метки:\n\[ОРГ_1\] организация\n\[ФИО_1\] человек\n\[ДОГОВОР_1\] номер договора/);
  assert.ok(common.endsWith('\nМоя задача: '));
  await p.click('sample'); await p.click('sampleAkt');
  await p.click('copyPrompt');
  const t = p.copied.at(-1);
  assert.match(t, /В этом документе встречаются такие метки:\n\[ОРГ_1\] организация\n\[ФИО_1\] человек\n\[ДОГОВОР_1\] номер договора или другого документа\n\[СУММА_1\] сумма денег\n\[ИНН_1\] ИНН\nМетки с другими словами/);
  assert.ok(!t.includes('[ТЕЛ_') && !t.includes('[КПП_'), 'типов, которых в акте нет, в перечне нет');
  for (const tag of new Set(p.api.state.hits.map(h => h.tag.replace(/_\d+\]/, '_')))) assert.ok(t.includes(tag), tag);
  assert.match(t, /в квадратных скобках, например \[ОРГ_1\]/);
  assert.match(t, /Пиши метки в ответе в точности как в документе/);
  assert.match(t, /Не склоняй и не переводи метки/);
  assert.match(t, /Не раскрывай метки и не придумывай за них настоящие/);
  assert.match(t, /Одинаковая метка означает одно и то же во всех документах/);
  assert.ok(t.endsWith('\nМоя задача: '), 'в конце место для задачи');
  assert.ok(!BANNED.test(t) && !BANNED.test(common), 'без длинного тире и знаков в прозе');
  assert.ok(!/Ромашка|Иванов|0047020267|4 180 000/.test(t), 'настоящих значений в просьбе нет');
  assert.match(p.el('promptInfo').textContent, /^Просьба скопирована/);
  assert.equal(p.el('promptBox').hidden, true);
  // сейф открыт, документа нет: метки берутся из сейфа
  const s = page();
  await s.api.loadSafe(s.file('a.сейф.json', safeJson('Клиент А', [{ tag: '[ТЕЛ_3]', type: 'PHONE', value: '+7 916 245-18-73', forms: ['+7 916 245-18-73'], keys: ['9162451873'] }])));
  await s.click('copyPrompt');
  assert.match(s.copied.at(-1), /В этом документе встречаются такие метки:\n\[ТЕЛ_3\] телефон\nМетки/);
});

test('просьба для нейросети: без navigator.clipboard копируется запасным путём, иначе текст показан для ручного копирования', async () => {
  const a = page({ clipboard: undefined, execCopy: true });
  await a.click('copyPrompt');
  assert.ok(a.copied.at(-1).endsWith('Моя задача: '));
  assert.match(a.el('promptInfo').textContent, /^Просьба скопирована/);
  const b = page({ clipboard: { writeText: async () => { throw new Error('NotAllowedError'); } }, execCopy: true });
  await b.click('copyPrompt');
  assert.equal(b.copied.length, 1, 'отказ буфера: сработал запасной путь');
  const c = page({ clipboard: undefined, execCopy: false });
  await c.click('copyPrompt');
  assert.equal(c.copied.length, 0);
  assert.equal(c.el('promptBox').hidden, false);
  assert.ok(c.el('promptBox').value.endsWith('Моя задача: '));
  assert.match(c.el('promptInfo').textContent, /^Скопировать не получилось/);
});
