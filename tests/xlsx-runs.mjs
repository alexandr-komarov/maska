// Куски <r> с пустым <t> в xlsx: Excel на таких предлагает восстановить файл (проверено владельцем в Excel для Mac)
export function emptyRuns(xml) {
  let n = 0;
  for (const m of xml.matchAll(/<((?:[\w.-]+:)?)r(?:\s[^>]*)?(?:\/>|>([\s\S]*?)<\/\1r>)/g)) {
    const body = m[2];
    if (body === undefined) { n++; continue; }
    const ts = [...body.matchAll(new RegExp(`<${m[1]}t(?:\\s[^>]*)?(?:/>|>([\\s\\S]*?)</${m[1]}t>)`, 'g'))];
    if (!ts.length || ts.every(t => !t[1])) n++;
  }
  return n;
}
