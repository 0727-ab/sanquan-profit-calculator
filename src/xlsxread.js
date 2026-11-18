/* ===== Excel 读取：在浏览器本地解析 .xlsx（不上传、不联网、无第三方库） ===== */
const XLSXRead = (() => {
  const td = new TextDecoder('utf-8');
  async function inflate(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const buf = await new Response(new Blob([bytes]).stream().pipeThrough(ds)).arrayBuffer();
    return new Uint8Array(buf);
  }
  // 读取 zip 目录
  function entries(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('这不是有效的 .xlsx 文件（如果是 .xls 旧格式，请在 Excel 里另存为 .xlsx）');
    const n = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
    const out = {};
    for (let i = 0; i < n; i++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true), nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
      const name = td.decode(u8.subarray(p + 46, p + 46 + nlen));
      out[name] = { method, csize, off };
      p += 46 + nlen + elen + clen;
    }
    return { get: async name => {
      const e = out[name]; if (!e) return null;
      const lnl = dv.getUint16(e.off + 26, true), lel = dv.getUint16(e.off + 28, true), start = e.off + 30 + lnl + lel;
      const data = u8.subarray(start, start + e.csize);
      return td.decode(e.method === 0 ? data : await inflate(data));
    }, names: Object.keys(out) };
  }
  const unesc = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, d) => String.fromCharCode(+d)).replace(/&amp;/g, '&');
  const colIdx = s => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  function parseShared(xml) {
    if (!xml) return [];
    const out = []; const re = /<si>([\s\S]*?)<\/si>/g; let m;
    while ((m = re.exec(xml))) { let t = ''; const r2 = /<t[^>]*>([\s\S]*?)<\/t>/g; let k; while ((k = r2.exec(m[1]))) t += k[1]; out.push(unesc(t)); }
    return out;
  }
  function parseSheet(xml, shared) {
    const rows = []; const rre = /<row[^>]*?r="(\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g; let rm;
    const cre = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    while ((rm = rre.exec(xml))) {
      const r = +rm[1] - 1, arr = []; if (!rm[2]) { rows[r] = arr; continue; }
      let cm; cre.lastIndex = 0;
      while ((cm = cre.exec(rm[2]))) {
        const c = colIdx(cm[1]), attrs = cm[2], body = cm[3] || '';
        const t = (attrs.match(/t="([^"]+)"/) || [])[1];
        let v = null;
        if (t === 'inlineStr') { const mm = body.match(/<t[^>]*>([\s\S]*?)<\/t>/); v = mm ? unesc(mm[1]) : ''; }
        else { const mm = body.match(/<v>([\s\S]*?)<\/v>/); if (mm) { const raw = unesc(mm[1]); v = t === 's' ? shared[+raw] : (t === 'str' || t === 'e' ? raw : t === 'b' ? raw === '1' : +raw); } }
        arr[c] = v;
      }
      rows[r] = arr;
    }
    return rows;
  }
  // 返回 { sheetNames, read(name) }
  async function open(arrayBuffer) {
    const z = entries(new Uint8Array(arrayBuffer));
    const wb = await z.get('xl/workbook.xml'); if (!wb) throw new Error('文件里找不到工作簿内容');
    const rels = await z.get('xl/_rels/workbook.xml.rels') || '';
    const relMap = {}; rels.replace(/<Relationship[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"/g, (m, id, tg) => { relMap[id] = tg; });
    rels.replace(/<Relationship[^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"/g, (m, tg, id) => { relMap[id] = relMap[id] || tg; });
    const sheets = []; wb.replace(/<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"/g, (m, n, id) => { sheets.push({ name: unesc(n), path: 'xl/' + relMap[id].replace(/^\/?xl\//, '').replace(/^\//, '') }); });
    const shared = parseShared(await z.get('xl/sharedStrings.xml'));
    return { sheetNames: sheets.map(s => s.name), read: async name => { const s = sheets.find(x => x.name === name); if (!s) return null; return parseSheet(await z.get(s.path), shared); } };
  }
  return { open };
})();
if (typeof module !== 'undefined') module.exports = XLSXRead;
