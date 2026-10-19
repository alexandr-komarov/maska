// Главный сценарий: договор, акт и акт сверки первой сделки обезличиваются по очереди с одним сейфом,
// возвращаются в docx и xlsx, затем ответ нейросети с искажёнными метками получает настоящие значения обратно
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { loadMask } from './mask-api.mjs';
import { readOffice, buildReturn, hideSums, isShort } from '../src/office.mjs';
import { readZip, byName } from '../src/zip.mjs';
import { decode } from '../src/ooxml.mjs';
import { safeNew, safeSerialize, safeParse, safeRestore, TAG_ANY, TAG_WORD } from '../src/safe.mjs';
import { emptyRuns } from './xlsx-runs.mjs';

const api = loadMask();
const hasTag = s => new RegExp(TAG_ANY.source).test(s);
const FILES = ['dogovor-podryada.docx', 'akt-sdachi-priemki.docx', 'akt-sverki-1c.xlsx'];
const deal = JSON.parse(fs.readFileSync(new URL('./deal.json', import.meta.url), 'utf8'));
const ent = id => deal.entities.find(e => e.id === id);

// сеанс один раз на все проверки: сейф сохраняется в файл и открывается заново перед каждым документом, как у человека
const run = (async () => {
  let safe = safeNew('Сделка 1');
  const docs = [];
  for (const name of FILES) {
    safe = safeParse(safeSerialize(safe)).safe;
    const bytes = new Uint8Array(fs.readFileSync(new URL('./deal/' + name, import.meta.url)));
    const read = await readOffice(name, bytes);
    const m = api.mask(read.text, [], { sums: hideSums(name), safe });
    safe = m.safe;
    const out = await buildReturn({ name, bytes, read, hits: m.hits, vault: m.vault, masked: m.masked });
    const back = await readOffice(out.name, out.bytes);
    docs.push({ name, bytes, read, m, out, back });
  }
  return { docs, safe };
})();

// метки, под которыми написания сущности ушли в текст файла
function tagsAt(doc, forms) {
  const tags = new Set();
  for (const f of forms) for (let i = doc.read.text.indexOf(f); i >= 0; i = doc.read.text.indexOf(f, i + 1))
    for (const h of doc.m.hits) if (h.s < i + f.length && h.e > i) tags.add(h.tag);
  return [...tags];
}

test('сценарий: одна фирма, один человек и один договор имеют одну метку в трёх файлах', async () => {
  const { docs } = await run;
  for (const id of ['org-kamennyj-most', 'org-severnyj-veter', 'fio-nesterova', 'fio-belousov', 'dog-41-КВ/2025']) {
    const per = docs.map(d => tagsAt(d, ent(id).forms));
    for (const [i, t] of per.entries()) assert.equal(t.length, 1, `${id} в ${FILES[i]}: ${t}`);
    assert.equal(new Set(per.flat()).size, 1, `${id}: ${per.map(t => t.join()).join(' / ')}`);
  }
  const org = tagsAt(docs[0], ent('org-kamennyj-most').forms)[0];
  assert.match(org, /^\[ОРГ_\d+\]$/);
  for (const d of docs) assert.ok(d.back.text.includes(org), `${d.name}: метка фирмы в возвращённом файле`);
  // ИНН и название фирмы под разными метками
  assert.notEqual(tagsAt(docs[0], ['7816135799'])[0], org);
});

test('сценарий: в возвращённых файлах нет исходных значений, ни в тексте, ни в XML частей', async () => {
  const { docs } = await run;
  for (const d of docs) {
    assert.deepEqual(d.out.left, [], d.name);
    const values = [...new Set(d.m.hits.map(h => h.value))].filter(v => !isShort(v) && !/[\t\n]/.test(v));
    assert.ok(values.length > 3, d.name);
    for (const v of values) assert.ok(!d.back.text.includes(v), `${d.name}: «${v}» в тексте`);
    // распакованный архив: каждая часть XML с раскрытыми сущностями
    const z = readZip(d.out.bytes);
    for (const e of z.entries) if (/\.(xml|rels)$/.test(e.name)) {
      const xml = decode(await e.text()), flat = xml.replace(/<[^>]+>/g, '');
      for (const v of values) assert.ok(!xml.includes(v) && !flat.includes(v), `${d.name} ${e.name}: «${v}»`);
    }
  }
});

test('сценарий: метки в квадратных скобках в xlsx записаны строками, формулы и куски <r> целы', async () => {
  const { docs } = await run;
  const d = docs[2], z0 = byName(readZip(d.bytes)), z1 = byName(readZip(d.out.bytes));
  const formulas = s => s.match(/<f\b[^>]*?(?:\/>|>[\s\S]*?<\/f>)/g) || [];
  for (const n of Object.keys(z1)) {
    if (!/\.xml$/.test(n)) continue;
    const a = await z1[n].text();
    assert.equal(emptyRuns(a), 0, n);
    if (/^xl\/worksheets\//.test(n)) {
      assert.deepEqual(formulas(a), formulas(await z0[n].text()), `${n}: формулы не изменились`);
      for (const f of formulas(a)) assert.ok(!hasTag(f), `${n}: метка внутри формулы`);
      // ячейка с меткой: общая строка, строка в ячейке или строка формулы, но не число и не формула
      for (const c of a.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        if (!c[2] || !/[[\]]/.test(decode(c[2]))) continue;
        assert.match(c[1], /\bt="(?:inlineStr|str)"/, `${n}: ${c[0].slice(0, 80)}`);
        assert.ok(!/<f\b/.test(c[2]) || /t="str"/.test(c[1]));
      }
    }
  }
  const ss = await z1['xl/sharedStrings.xml'].text();
  assert.ok(/\[ОРГ_\d+\]/.test(decode(ss)), 'метки лежат в общих строках как текст');
  assert.ok(!/^[=+\-@]/m.test(Object.values(TAG_WORD).map(w => `[${w}_1]`).join('\n')), 'метка не начинается со знака формулы');
});

test('сценарий: ответ нейросети с искажёнными метками получает все значения обратно', async () => {
  const { docs, safe } = await run;
  const tags = [...new Set(docs.flatMap(d => d.back.text.match(TAG_ANY) || []))];
  assert.ok(tags.length > 20, `меток в файлах: ${tags.length}`);
  // как пишет нейросеть: в нижнем регистре, без скобок, с пробелом, с экранированием Markdown, «е» вместо «ё»
  const twist = [t => t, t => t.toLowerCase(), t => t.slice(1, -1), t => t.replace('_', ' '), t => t.replace('[', '\\[').replace('_', '\\_').replace(']', '\\]'), t => t.replace(/Ё/g, 'Е')];
  const lines = tags.map((t, i) => `${i + 1}. По сверке ${twist[i % twist.length](t)} совпадает.`);
  const answer = `Сверил договор, акт и акт сверки.\n${lines.join('\n')}\nРасхождений нет.`;
  const r = safeRestore(answer, safe);
  assert.equal(r.replaced, tags.length);
  assert.deepEqual(r.unknown, []);
  assert.deepEqual(r.declined, []);
  for (const t of tags) assert.ok(r.text.includes(safe.records.find(x => x.tag === t).value), t);
  assert.ok(!hasTag(r.text), 'меток не осталось');
  // фирма возвращается полным названием, человек ФИО в именительном падеже
  const org = safe.records.find(x => x.forms.includes('ООО «Каменный Мост»'));
  assert.match(org.value, /«Каменный Мост»$/);
  assert.equal(safe.records.find(x => x.forms.includes('Нестерова А.В.')).value, 'Нестерова Алла Викторовна');
});
