// Окно проверки: уверенное и сомнительное раздельно, сомнительное под меткой, снятие метки одним щелчком,
// сейф помнит решение, сохранение закрыто до подтверждения. Браузер здесь не запускается: страница идёт на заглушке DOM
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { loadMask } from './mask-api.mjs';
import { safeNew, safeSetSkip, safeSerialize, safeParse, safeDigest, safeRestore } from '../src/safe.mjs';

const api = loadMask();
const DOUBT = ['NAME', 'NUM', 'AT'];
const dz = r => r.hits.filter(h => DOUBT.includes(h.t)).map(h => h.value);
const TEXT = 'Проект «Полярная звезда» ведёт Нгуен Ван Лонг из Уфы. Заказчик ООО «Ромашка», ИНН 7729654321, пропуск 77001234, связь через ivan@local.\n' +
  'Образцы для проекта заберёт Лонг. Договор подписан, звезда на логотипе остаётся.';

test('правила: без opts.doubt маска работает как раньше, сомнительное не прячется', () => {
  const r = api.mask(TEXT, []);
  assert.deepEqual(dz(r), []);
  assert.ok(r.masked.includes('Полярная звезда') && r.masked.includes('77001234') && r.masked.includes('ivan@local'));
  assert.deepEqual(r.open, []);
});

test('правила: уверенное и сомнительное раздельно, сомнительное под своей меткой', () => {
  const r = api.mask(TEXT, [], { doubt: true });
  assert.deepEqual(r.hits.filter(h => !DOUBT.includes(h.t)).map(h => h.t + ' ' + h.value), ['ORG ООО «Ромашка»', 'INN 7729654321']);
  assert.deepEqual(r.hits.filter(h => DOUBT.includes(h.t)).map(h => h.t + ' ' + h.value),
    ['NAME Полярная звезда', 'NAME Нгуен Ван Лонг', 'NUM 77001234', 'AT ivan@local', 'NAME Лонг']);
  assert.equal(r.masked, 'Проект «[НАЗВАНИЕ_1]» ведёт [НАЗВАНИЕ_2] из Уфы. Заказчик [ОРГ_1], ИНН [ИНН_1], пропуск [НОМЕР_1], связь через [КОНТАКТ_1].\n' +
    'Образцы для проекта заберёт [НАЗВАНИЕ_3]. Договор подписан, звезда на логотипе остаётся.');
  // метки сомнительного возвращаются в ответ нейросети, как остальные
  assert.equal(safeRestore('Пропуск [НОМЕР_1] у [НАЗВАНИЕ_2]', r.safe).text, 'Пропуск 77001234 у Нгуен Ван Лонг');
  assert.match(r.before, /<mark class="NUM" data-open="\[НОМЕР_1\]"/);
});

test('правила: что сомнительным не считается', () => {
  const keep = t => assert.deepEqual(dz(api.mask(t, [], { doubt: true })), [], t);
  keep('Заказчик передаёт Подрядчику Объект по Акту. Стороны подписали Договор в Москве и Санкт-Петербурге.'); // словарь и география
  keep('Оплата прошла. Возврат не нужен.\nНаименование\tКоличество\n- Доставка включена'); // начало предложения, строки, ячейки
  keep('ДОГОВОР АРЕНДЫ НЕЖИЛОГО ПОМЕЩЕНИЯ\nАрбитражный суд Калужской области, Калужский политехнический колледж'); // заголовок заглавными, прилагательные
  keep('Счёт в ПАО Сбербанк и АО «Альфа-Банк», начальник ОТК, НДС 20 %, пост КПП-2, болт М12х120'); // банки, сокращения, размеры
  keep('Итого 1 284 500 за 2024-2025 годы, отчёт от 2025-03-17, остаток 1284500.50, код 1234567'); // суммы, годы, дата, короткое число
  // в таблице суммы остаются: восемь цифр подряд там число, а не номер
  assert.deepEqual(dz(api.mask('Итого\t12845000', [], { doubt: true, sums: false })), []);
  assert.deepEqual(dz(api.mask('Итог: 12845000', [], { doubt: true })), ['12845000']);
});

test('правила: слово в начале предложения берётся, если оно же есть в середине; падежи идут одной меткой', () => {
  const r = api.mask('Вега привезла образцы. Склад Веги закрыт, а у Веги выходной.', [], { doubt: true });
  assert.deepEqual(dz(r), ['Вега', 'Веги', 'Веги']);
  assert.equal(new Set(r.hits.map(h => h.tag)).size, 1);
  assert.deepEqual(dz(api.mask('Вега привезла образцы.', [], { doubt: true })), []);
  // отчество прячется и в начале строки
  assert.deepEqual(dz(api.mask('Ильинична просила перезвонить.', [], { doubt: true })), ['Ильинична']);
});

test('снятие метки: слово из списка сейфа не прячется во всех местах и формах, остальное спрятано', () => {
  const safe = safeSetSkip(safeNew(), ['Лонг', '77001234']);
  const r = api.mask(TEXT, [], { doubt: true, safe });
  assert.deepEqual(dz(r), ['Полярная звезда', 'Нгуен Ван Лонг', 'ivan@local']);
  assert.deepEqual(r.open.map(o => o.t + ' ' + o.value), ['NUM 77001234', 'NAME Лонг']);
  assert.ok(r.masked.includes('пропуск 77001234') && r.masked.includes('заберёт Лонг.'));
  assert.ok(!r.safe.records.some(x => x.value === 'Лонг' || x.value === '77001234'), 'открытое в сейф не пишется');
  const forms = api.mask('Склад Веги закрыт, у Веге выходной, к Вегой.', [], { doubt: true, safe: safeSetSkip(safeNew(), ['Вега']) });
  assert.deepEqual(dz(forms), []);
});

test('снятие метки: сейф хранит список в файле; слово, спрятанное в прошлом документе, после снятия не прячется', () => {
  const first = api.mask(TEXT, [], { doubt: true });
  assert.ok(first.safe.records.some(x => x.tag === '[НОМЕР_1]'));
  // следующий документ с тем же сейфом: число прячется прежней меткой
  const again = api.mask('Пропуск 77001234 сдан.', [], { doubt: true, safe: first.safe });
  assert.equal(again.masked, 'Пропуск [НОМЕР_1] сдан.');
  const s = safeSetSkip(first.safe, ['77001234']);
  assert.notEqual(safeDigest(s), safeDigest(first.safe), 'сейф считается изменённым');
  const file = safeParse(safeSerialize(s)).safe;
  assert.deepEqual(file.skip, ['77001234']);
  assert.equal(api.mask('Пропуск 77001234 сдан.', [], { doubt: true, safe: file }).masked, 'Пропуск 77001234 сдан.');
  assert.equal(safeRestore('[НОМЕР_1]', file).text, '77001234', 'запись осталась: ответы по прежним документам возвращаются');
  assert.deepEqual(safeSetSkip(s, []).skip, []);
  assert.deepEqual(safeParse(JSON.stringify({ format: 'maska-safe', version: 1, data: { records: [] } })).safe.skip, [], 'сейф без списка читается');
  assert.match(safeParse(JSON.stringify({ format: 'maska-safe', version: 1, data: { records: [], skip: 'x' } })).error, /повреждён/);
});

// Заглушка DOM под то, что зовёт страница (как в page.test.mjs), плюс буфер обмена
const html = fs.readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const script = html.slice(html.indexOf('<script>\n') + 9, html.lastIndexOf('</script>'));
function page() {
  const els = {}, saved = [], copied = [];
  const el = id => els[id] ||= { id, value: '', textContent: '', innerHTML: '', hidden: false, disabled: false, checked: id === 'sums', dataset: {},
    classList: { add() {}, remove() {}, toggle() {} }, scrollIntoView() {}, getBoundingClientRect: () => ({ top: 0 }), addEventListener() {} };
  // начальное состояние разметки
  Object.assign(el('dlFile'), { hidden: true, disabled: true }); el('retWarn').hidden = true; el('out').hidden = true; el('selBar').hidden = true;
  const steps = ['s-doc', 's-check', 's-after', 's-answer'].map(go => ({ ...el('step-' + go), dataset: { go } }));
  const document = { getElementById: el, querySelectorAll: () => steps, addEventListener() {}, body: { scrollHeight: 0 },
    createElement: () => { const a = { click() { saved.push({ name: a.download }); } }; return a; } };
  const URL_ = { createObjectURL: b => { document.lastBlob = b; return 'blob:x'; }, revokeObjectURL() {} };
  const fn = new Function('document', 'addEventListener', 'innerHeight', 'scrollY', 'navigator', 'URL',
    script + '\n;return { loadFile, prepareReturn, offerSel, get state() { return state; }, get reviewed() { return reviewed; }, get safe() { return curSafe; } };');
  const api = fn(document, () => {}, 800, 0, { clipboard: { writeText: async t => { copied.push(t); } } }, URL_);
  const click = async (id, target) => { await els[id].onclick?.({ target: target || els[id] }); };
  const run = async text => { el('src').value = text; await click('run'); await api.prepareReturn(); };
  const chip = (id, dataset) => click(id, { dataset });
  return { api, el, click, run, chip, saved, copied };
}
const closed = p => ['dlMasked', 'copy'].every(id => p.el(id).disabled === true);

test('окно: уверенное по типам, сомнительное отдельным списком с кнопками', async () => {
  const p = page();
  await p.run(TEXT);
  const h = p.el('review').innerHTML;
  assert.match(h, /<b>Спрятано уверенно: 2 места<\/b>/);
  assert.match(h, /<span class="ORG">Организация · 1<\/span><span>ООО «Ромашка»<\/span>/);
  assert.match(h, /<span class="INN">ИНН · 1<\/span><span>7729654321<\/span>/);
  assert.match(h, /<b>Сомнительное, тоже спрятано: 5<\/b>/);
  assert.match(h, /Слова с заглавной буквы, которых нет в словаре Маски · 3/);
  assert.match(h, /<button class="dz" data-open="\[НАЗВАНИЕ_2\]" title="[^"]*">Нгуен Ван Лонг<u>не прятать<\/u><\/button>/);
  assert.match(h, /Длинные числа без подписи · 1.*data-open="\[НОМЕР_1\]".*Со знаком @ · 1.*data-open="\[КОНТАКТ_1\]"/s);
  assert.ok(h.indexOf('Спрятано уверенно') < h.indexOf('Сомнительное'), 'уверенное и сомнительное разными блоками');
  assert.ok(!/—|→/.test(h + p.el('gateText').textContent), 'в текстах окна нет длинного тире и стрелок');
});

test('окно: скачивание и копирование закрыты, пока просмотр не подтверждён', async () => {
  const p = page();
  await p.run(TEXT);
  assert.equal(p.api.reviewed, false);
  assert.ok(closed(p));
  assert.equal(p.el('gateTitle').textContent, 'Сохранение закрыто');
  assert.match(p.el('gateText').textContent, /Просмотрите сомнительные места, их 5/);
  assert.equal(p.el('gate').className, 'check warn');
  // щелчки по закрытым кнопкам ничего не отдают
  await p.click('dlMasked'); await p.click('copy'); await p.click('dlFile');
  assert.deepEqual([p.saved, p.copied], [[], []]);
  await p.click('okReview');
  assert.equal(p.api.reviewed, true);
  assert.ok(!p.el('dlMasked').disabled && !p.el('copy').disabled);
  assert.equal(p.el('gateTitle').textContent, 'Проверка пройдена');
  assert.equal(p.el('okReview').hidden, true);
  await p.click('dlMasked'); await p.click('copy');
  assert.equal(p.saved.at(-1).name, 's-metkami.txt');
  assert.deepEqual(p.copied, [p.api.state.masked]);
});

test('окно: новая маскировка и правка текста сбрасывают подтверждение', async () => {
  const p = page();
  await p.run(TEXT);
  await p.click('okReview');
  await p.click('run');
  assert.equal(p.api.reviewed, false, 'после новой маскировки');
  assert.ok(closed(p));
  await p.click('okReview');
  p.el('src').value += '\nДописан Петров Иван Сергеевич';
  p.el('src').oninput();
  assert.ok(closed(p), 'после правки текста');
  assert.equal(p.el('gateTitle').textContent, 'Текст изменён');
  await p.click('dlMasked'); await p.click('copy');
  assert.deepEqual([p.saved, p.copied], [[], []]);
  await p.click('okReview');
  assert.ok(closed(p), 'подтвердить старую маску для нового текста нельзя');
  await p.click('run'); await p.click('okReview');
  assert.ok(!closed(p));
  assert.ok(p.api.state.masked.includes('Дописан [ФИО_1]'));
});

test('окно: файл в исходном формате тоже ждёт подтверждения', async () => {
  const p = page();
  const bytes = new Uint8Array(fs.readFileSync(new URL('./deal/dogovor-podryada.docx', import.meta.url)));
  await p.api.loadFile({ name: 'd.docx', size: bytes.length, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });
  await p.api.prepareReturn();
  assert.equal(p.el('dlFile').disabled, true);
  await p.click('dlFile');
  assert.deepEqual(p.saved, []);
  await p.click('okReview');
  assert.equal(p.el('dlFile').disabled, false);
  await p.click('dlFile');
  assert.equal(p.saved.at(-1).name, 'd-maska.docx');
});

test('окно: один щелчок снимает метку со всех мест, сейф запоминает, слово можно спрятать снова', async () => {
  const p = page();
  await p.run('Груз примет Вега. Склад Веги закрыт, пропуск 77001234 и 77001234.');
  assert.equal(p.api.state.masked, 'Груз примет [НАЗВАНИЕ_1]. Склад [НАЗВАНИЕ_1] закрыт, пропуск [НОМЕР_1] и [НОМЕР_1].');
  assert.match(p.el('review').innerHTML, /Вега, Веги<i>2 места<\/i><u>не прятать<\/u>/);
  await p.click('okReview');
  await p.chip('review', { open: '[НАЗВАНИЕ_1]' });
  assert.equal(p.api.state.masked, 'Груз примет Вега. Склад Веги закрыт, пропуск [НОМЕР_1] и [НОМЕР_1].');
  assert.equal(p.api.reviewed, false, 'снятие метки меняет результат: подтверждение сброшено');
  assert.deepEqual(p.api.safe.skip, ['Вега', 'Веги']);
  assert.ok(!p.api.safe.records.some(r => r.type === 'NAME'), 'записи для открытого слова в сейфе нет');
  assert.match(p.el('review').innerHTML, /Оставлено на виду · 1.*<button class="dz off" data-hide="NAME\|вег">Вега, Веги<i>2 места<\/i><u>спрятать снова<\/u>/s);
  // щелчок по слову прямо в панели «До»
  await p.chip('before', { open: '[НОМЕР_1]' });
  assert.equal(p.api.state.masked, 'Груз примет Вега. Склад Веги закрыт, пропуск 77001234 и 77001234.');
  assert.match(p.el('review').innerHTML, /<b>Сомнительных мест нет<\/b>/);
  // следующий документ с тем же сейфом: решение действует
  await p.run('Вега прислала счёт, пропуск 77001234, новый пропуск 88002345.');
  assert.equal(p.api.state.masked, 'Вега прислала счёт, пропуск 77001234, новый пропуск [НОМЕР_1].');
  const d = JSON.parse(safeSerialize(p.api.safe));
  assert.deepEqual(d.data.skip, ['Вега', 'Веги', '77001234']);
  await p.chip('review', { hide: 'NUM|77001234' });
  assert.match(p.api.state.masked, /пропуск \[НОМЕР_\d\], новый пропуск \[НОМЕР_\d\]\./);
  assert.deepEqual(p.api.safe.skip, ['Вега', 'Веги']);
});

test('окно: ничего не найдено, страница говорит об этом и всё равно показывает проверку', async () => {
  const p = page();
  await p.run('обычный текст без имён и реквизитов, только слова.');
  assert.equal(p.el('out').hidden, false);
  assert.match(p.el('review').innerHTML, /<b>Уверенных находок нет<\/b>.*Правила не нашли в документе.*<b>Сомнительных мест нет<\/b>/s);
  assert.ok(closed(p), 'и здесь сохранение ждёт подтверждения');
  assert.match(p.el('gateText').textContent, /Сомнительных мест нет, просмотрите текст/);
  // уверенных нет, сомнительное есть
  await p.run('обычный текст, но в нём пропуск 77001234.');
  assert.match(p.el('review').innerHTML, /<b>Уверенных находок нет<\/b>.*<b>Сомнительное, тоже спрятано: 1<\/b>/s);
});

test('окно: пропущенное добавляется выделением в тексте', async () => {
  const p = page();
  await p.run('Поставка для стройки на объекте северный берег, второй этап.');
  assert.equal(p.el('selBar').hidden, true);
  p.api.offerSel('  северный\n берег ');
  assert.equal(p.el('selBar').hidden, false);
  assert.equal(p.el('addSel').textContent, 'Спрятать «северный берег» во всём документе');
  await p.click('okReview');
  await p.click('addSel');
  assert.equal(p.el('terms').value, 'северный берег');
  assert.equal(p.api.state.masked, 'Поставка для стройки на объекте [СЛОВО_1], второй этап.');
  assert.deepEqual(p.api.safe.words, ['северный берег'], 'слово ушло в сейф и будет искаться в следующих документах');
  assert.equal(p.api.reviewed, false);
  assert.equal(p.el('selBar').hidden, true);
  for (const bad of ['', 'я', 'слово, запятая', 'д'.repeat(81)]) { p.api.offerSel(bad); assert.equal(p.el('selBar').hidden, true, bad); }
});
