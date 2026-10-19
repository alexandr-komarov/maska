// Zip без библиотек: чтение в список записей и запись обратно.
// Непереписанные записи уходят байт в байт (локальный заголовок, сжатые данные, дескриптор),
// изменённые сжимаются deflate-raw через CompressionStream. Только веб-API: работает в браузере и в Node.

// BOM оставляем в строке, чтобы переписанная часть сохранила его
const td = new TextDecoder('utf-8', { ignoreBOM: true }), te = new TextEncoder();

// CRC32 своей таблицей
const CRC = new Uint32Array(256);
for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; CRC[n] = c >>> 0; }
export function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }

async function pipe(u8, stream) { return new Uint8Array(await new Response(new Blob([u8]).stream().pipeThrough(stream)).arrayBuffer()); }
export const inflate = u8 => pipe(u8, new DecompressionStream('deflate-raw'));
export const deflate = u8 => pipe(u8, new CompressionStream('deflate-raw'));

// Чтение. Возвращает {entries, comment}; у записи name, method, raw (сжатые байты), data() и setData()
export function readZip(input) {
  // Buffer из Node режет slice без копии, поэтому приводим к обычному Uint8Array
  const u8 = ArrayBuffer.isView(input) ? new Uint8Array(input.buffer, input.byteOffset, input.byteLength) : new Uint8Array(input);
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('не zip-архив (для docx и xlsx нужен формат Office 2007 и новее; старые .doc и .xls не поддерживаются)');
  if (eocd >= 20 && dv.getUint32(eocd - 20, true) === 0x07064b50) throw new Error('ZIP64 не поддерживается: архив больше 4 ГБ или в нём больше 65535 файлов');
  const disk = dv.getUint16(eocd + 4, true), n = dv.getUint16(eocd + 10, true), cdSize = dv.getUint32(eocd + 12, true), cdOff = dv.getUint32(eocd + 16, true);
  if (n === 0xFFFF || cdOff === 0xFFFFFFFF || cdSize === 0xFFFFFFFF) throw new Error('ZIP64 не поддерживается');
  if (disk !== 0) throw new Error('многотомный zip не поддерживается');
  const comment = u8.slice(eocd + 22, eocd + 22 + dv.getUint16(eocd + 20, true));
  const entries = [];
  let p = cdOff;
  for (let k = 0; k < n; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('битый центральный каталог zip');
    const flags = dv.getUint16(p + 8, true), method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true), usize = dv.getUint32(p + 24, true), crc = dv.getUint32(p + 16, true);
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
    if (csize === 0xFFFFFFFF || usize === 0xFFFFFFFF || off === 0xFFFFFFFF) throw new Error('ZIP64 не поддерживается');
    const nameBytes = u8.slice(p + 46, p + 46 + nl);
    const central = u8.slice(p, p + 46 + nl + el + cl);
    if (dv.getUint32(off, true) !== 0x04034b50) throw new Error('битый локальный заголовок zip');
    const lnl = dv.getUint16(off + 26, true), lel = dv.getUint16(off + 28, true);
    const dataStart = off + 30 + lnl + lel;
    let end = dataStart + csize;
    // дескриптор данных после сжатых байт (флаг 3): с подписью 16 байт, без неё 12
    if (flags & 8) end += dv.getUint32(end, true) === 0x08074b50 ? 16 : 12;
    entries.push(makeEntry({ name: td.decode(nameBytes), nameBytes, flags, method, crc, csize, usize, central,
      local: u8.slice(off, end), raw: u8.slice(dataStart, dataStart + csize) }));
    p += 46 + nl + el + cl;
  }
  return { entries, comment };
}

function makeEntry(e) {
  e.changed = null;
  e.data = async () => {
    if (e.changed) return e.changed;
    if (e.flags & 1) throw new Error(`зашифрованный файл в архиве: ${e.name}`);
    if (e.method === 0) return e.raw;
    if (e.method !== 8) throw new Error(`неизвестный метод сжатия ${e.method}: ${e.name}`);
    const out = await inflate(e.raw);
    if (crc32(out) !== e.crc) throw new Error(`не сходится CRC: ${e.name}`);
    return out;
  };
  e.text = async () => td.decode(await e.data());
  e.setData = u8 => { e.changed = typeof u8 === 'string' ? te.encode(u8) : u8; };
  return e;
}

// Запись. Порядок записей сохраняется, непереписанные копируются как есть
export async function writeZip({ entries, comment = new Uint8Array(0) }) {
  const parts = [], central = [];
  let off = 0;
  for (const e of entries) {
    let local, cen;
    if (!e.changed) {
      local = e.local;
      cen = e.central.slice();
    } else {
      const data = e.changed, crc = crc32(data);
      // mimetype и пустые файлы кладём без сжатия, как принято в ODF и EPUB
      const store = e.name === 'mimetype' || data.length === 0;
      const comp = store ? data : await deflate(data);
      const method = store ? 0 : 8;
      // снимаем дескриптор данных (бит 3) и шифрование (бит 0), признак UTF-8 в имени (бит 11) остаётся
      const flags = e.flags & ~9;
      const cdv = new DataView(e.central.buffer, e.central.byteOffset, e.central.byteLength);
      const time = cdv.getUint16(12, true), date = cdv.getUint16(14, true);
      local = new Uint8Array(30 + e.nameBytes.length + comp.length);
      const lv = new DataView(local.buffer, local.byteOffset);
      lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, flags, true); lv.setUint16(8, method, true);
      lv.setUint16(10, time, true); lv.setUint16(12, date, true); lv.setUint32(14, crc, true);
      lv.setUint32(18, comp.length, true); lv.setUint32(22, data.length, true); lv.setUint16(26, e.nameBytes.length, true); lv.setUint16(28, 0, true);
      local.set(e.nameBytes, 30); local.set(comp, 30 + e.nameBytes.length);
      // в центральной записи меняем флаги, метод, CRC и размеры; дополнительное поле и комментарий оставляем
      cen = e.central.slice();
      const cv = new DataView(cen.buffer, cen.byteOffset);
      cv.setUint16(6, 20, true); cv.setUint16(8, flags, true); cv.setUint16(10, method, true); cv.setUint32(16, crc, true);
      cv.setUint32(20, comp.length, true); cv.setUint32(24, data.length, true);
    }
    new DataView(cen.buffer, cen.byteOffset).setUint32(42, off, true);
    parts.push(local); central.push(cen); off += local.length;
    if (off > 0xFFFFFFFF) throw new Error('архив больше 4 ГБ, ZIP64 не поддерживается');
  }
  const cdOff = off; let cdSize = 0;
  for (const c of central) { parts.push(c); cdSize += c.length; }
  const eocd = new Uint8Array(22 + comment.length), ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, entries.length, true); ev.setUint16(10, entries.length, true);
  ev.setUint32(12, cdSize, true); ev.setUint32(16, cdOff, true); ev.setUint16(20, comment.length, true); eocd.set(comment, 22);
  parts.push(eocd);
  const out = new Uint8Array(parts.reduce((a, b) => a + b.length, 0));
  let q = 0; for (const x of parts) { out.set(x, q); q += x.length; }
  return out;
}

// Сборка архива с нуля, для тестовых файлов: [[имя, строка или байты], ...]
export async function buildZip(files) {
  const entries = files.map(([name, content]) => {
    const nameBytes = te.encode(name);
    const central = new Uint8Array(46 + nameBytes.length), cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(12, 0, true); cv.setUint16(14, (2026 - 1980) << 9 | 10 << 5 | 4, true);
    central.set(nameBytes, 46);
    const e = makeEntry({ name, nameBytes, flags: 0, method: 8, central });
    e.setData(content);
    return e;
  });
  return writeZip({ entries });
}

// Удобство: словарь имя → запись
export const byName = zip => Object.fromEntries(zip.entries.map(e => [e.name, e]));
