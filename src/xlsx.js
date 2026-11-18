/* ===== Excel 导出：纯 JS 生成带公式的 .xlsx（无需联网、无需第三方库） ===== */
const XLSX = (() => {
  let NOCACHE = false;
  /* ---------- ZIP（store 模式） ---------- */
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(files) {
    const enc = new TextEncoder(), parts = [], central = []; let off = 0;
    files.forEach(f => {
      const name = enc.encode(f.name), data = typeof f.data === 'string' ? enc.encode(f.data) : f.data, crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, name.length, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true); c.setUint16(28, name.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), name);
      off += 30 + name.length + data.length;
    });
    const cdSize = central.reduce((s, a) => s + a.length, 0), e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, cdSize, true); e.setUint32(16, off, true);
    const all = parts.concat(central, [new Uint8Array(e.buffer)]), out = new Uint8Array(all.reduce((s, a) => s + a.length, 0));
    let p = 0; all.forEach(a => { out.set(a, p); p += a.length; }); return out;
  }

  /* ---------- 工作表模型 ---------- */
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colL = n => { let s = ''; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  // 样式编号
  const S = { txt: 0, title: 1, head: 2, inMoney: 3, money: 4, inInt: 5, int: 6, inPct: 7, pct: 8, inTxt: 9, note: 10, keyMoney: 11, sec: 12, keyInt: 13, wan: 14 };
  class Sheet {
    constructor(name) { this.name = name; this.rows = {}; this.widths = []; }
    set(c, r, cell) { (this.rows[r] = this.rows[r] || {})[c] = cell; }
    xml() {
      let s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
      if (this.widths.length) s += '<cols>' + this.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>';
      s += '<sheetData>';
      Object.keys(this.rows).map(Number).sort((a, b) => a - b).forEach(r => {
        s += `<row r="${r}">`;
        Object.keys(this.rows[r]).map(Number).sort((a, b) => a - b).forEach(c => {
          const x = this.rows[r][c], ref = colL(c) + r, st = x.s ? ` s="${x.s}"` : '';
          if (x.f !== undefined) {
            const isStr = typeof x.v === 'string';
            const v = NOCACHE || x.v === undefined || x.v === null || (typeof x.v === 'number' && !isFinite(x.v)) ? '' : `<v>${esc(isStr ? x.v : +x.v.toFixed(10))}</v>`;
            s += `<c r="${ref}"${st}${isStr ? ' t="str"' : ''}><f>${esc(x.f)}</f>${v}</c>`;
          } else if (typeof x.v === 'number' && isFinite(x.v)) s += `<c r="${ref}"${st}><v>${+x.v.toFixed(10)}</v></c>`;
          else if (x.v !== undefined && x.v !== null && x.v !== '') s += `<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${esc(x.v)}</t></is></c>`;
          else s += `<c r="${ref}"${st}/>`;
        });
        s += '</row>';
      });
      return s + '</sheetData></worksheet>';
    }
  }
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="4"><numFmt numFmtId="167" formatCode="#,##0.0;[Red]-#,##0.0"/><numFmt numFmtId="164" formatCode="&quot;¥&quot;#,##0.00;[Red]-&quot;¥&quot;#,##0.00"/><numFmt numFmtId="165" formatCode="0.0%"/><numFmt numFmtId="166" formatCode="#,##0"/></numFmts>
<fonts count="5"><font><sz val="11"/><name val="Microsoft YaHei"/></font><font><b/><sz val="11"/><name val="Microsoft YaHei"/></font><font><b/><sz val="14"/><name val="Microsoft YaHei"/></font><font><i/><sz val="10"/><color rgb="FF6B6B6B"/><name val="Microsoft YaHei"/></font><font><b/><sz val="11"/><color rgb="FF1C5CAB"/><name val="Microsoft YaHei"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF4CC"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF6"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF3F3F1"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border/><border><bottom style="thin"><color rgb="FF9A9A9A"/></bottom></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="15">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="166" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="164" fontId="4" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="166" fontId="4" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

  function book(sheets) {
    const files = [];
    files.push({ name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '</Types>' });
    files.push({ name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' });
    files.push({ name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>' });
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') + `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` });
    files.push({ name: 'xl/styles.xml', data: STYLES });
    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: s.xml() }));
    return zip(files);
  }

  /* ---------- 由计算器状态生成工作簿 ---------- */
  const sheetName = (n, used) => { let s = String(n || 'SKU').replace(/[\[\]:*?\/\\']/g, '').slice(0, 28) || 'SKU'; let x = s, i = 2; while (used.has(x)) x = s + i++; used.add(x); return x; };
  const q = n => `'${n}'`;

  function laborSheet(L, E) {
    const sh = new Sheet('人工'); sh.widths = [22, 10, 12, 14, 12, 12, 16];
    sh.set(0, 1, { v: '人工成本计算（黄色格可修改）', s: S.title });
    ['岗位', '人数', '月薪', '社保公积金比例', '绩效基数', '绩效系数', '月人工合计'].forEach((h, c) => sh.set(c, 2, { v: h, s: S.head }));
    const roles = L.roles || []; let r = 3;
    const team = E.laborTeam(L);
    roles.forEach((ro, i) => {
      sh.set(0, r, { v: ro.name || '岗位' + (i + 1), s: S.inTxt }); sh.set(1, r, { v: +ro.count || 0, s: S.inInt }); sh.set(2, r, { v: +ro.salary || 0, s: S.inMoney });
      sh.set(3, r, { v: (+ro.social || 0) / 100, s: S.inPct }); sh.set(4, r, { v: +ro.perfBase || 0, s: S.inMoney }); sh.set(5, r, { v: +ro.perfCoef || 0, s: S.inTxt });
      sh.set(6, r, { f: `B${r}*(C${r}*(1+D${r})+E${r}*F${r})`, v: team.roles[i].total, s: S.money }); r++;
    });
    const first = 3, last = Math.max(3, r - 1); if (!roles.length) r = 4;
    const ref = {};
    const row = (label, cell, key) => { sh.set(0, r, { v: label, s: S.txt }); sh.set(1, r, cell); ref[key] = `${q('人工')}!$B$${r}`; r++; };
    r++;
    row('团队月人工合计', { f: roles.length ? `SUM(G${first}:G${last})` : '0', v: team.team, s: S.money }, 'team');
    row('每袋提成（元/袋）', { v: +L.piece || 0, s: S.inMoney }, 'piece');
    row('收入提成比例', { v: (+L.revPct || 0) / 100, s: S.inPct }, 'revPct');
    r++;
    sh.set(0, r, { v: '阶梯激励（门槛需从小到大）', s: S.sec }); sh.set(1, r, { v: '销量门槛', s: S.sec }); sh.set(2, r, { v: '奖励金额', s: S.sec }); r++;
    const tiers = E.sortedTiers(L); ref.tiers = [];
    tiers.forEach((t, i) => { sh.set(0, r, { v: '第' + (i + 1) + '档', s: S.txt }); sh.set(1, r, { v: t.threshold, s: S.inInt }); sh.set(2, r, { v: t.amount, s: S.inMoney }); ref.tiers.push({ t: `${q('人工')}!$B$${r}`, a: `${q('人工')}!$C$${r}`, tv: t.threshold, av: t.amount }); r++; });
    r++; sh.set(0, r, { v: '说明：团队月人工 × 各 SKU 表里的“人工归属比例” = 计入该 SKU 固定费用的人工；提成按袋计入单位成本；达到门槛时发放对应档位的奖励（取达到的最高一档）。', s: S.note });
    return { sh, ref };
  }

  function skuSheet(name, sku, state, E, lref) {
    const sh = new Sheet(name); sh.widths = [30, 16, 16, 46];
    const L = state.labor, R = { base: E.compute(sku, 'base', L), promo: E.compute(sku, 'promo', L) };
    const keys = ['base', 'promo'], col = { base: 1, promo: 2 }, CL = { base: 'B', promo: 'C' };
    sh.set(0, 1, { v: (sku.name || 'SKU') + ' 利润测算（黄色格可修改，其余为公式）', s: S.title });
    ['项目', '基准方案', '促销方案', '说明'].forEach((h, c) => sh.set(c, 2, { v: h, s: S.head }));
    let r = 3; const row = {};
    const line = (key, label, note, fn) => { row[key] = r; sh.set(0, r, { v: label }); if (note) sh.set(3, r, { v: note, s: S.note }); keys.forEach(k => { const c = fn(k, CL[k], r); if (c) sh.set(col[k], r, c); }); r++; };
    const sec = label => { sh.set(0, r, { v: label, s: S.sec }); [1, 2, 3].forEach(c => sh.set(c, r, { v: '', s: S.sec })); r++; };
    sec('输入参数');
    line('list', '标价（元/袋）', '扣渠道费前的单袋售价', k => ({ v: +sku[k].list, s: S.inMoney }));
    line('disc', '商家承担折扣（元/袋）', '直接录折后价时填 0', k => ({ v: +(sku[k].discount || 0), s: S.inMoney }));
    line('vol', '预计销量（袋）', '', k => ({ v: +sku[k].volume, s: S.inInt }));
    line('pur', '采购价', '按箱录入时为元/箱，按袋录入时为元/袋', k => ({ v: +sku[k].purchase, s: S.inMoney }));
    line('box', '每箱袋数（按袋录入填 1）', '', k => ({ v: sku[k].purchaseUnit === 'box' ? +sku[k].bagsPerBox : 1, s: S.inInt }));
    // 有效售价行号先占位，后面回填公式引用
    const effRow = () => row.eff;
    line('labor', '人工（元/袋）', '来自人工表时为公式；不计人工时为 0', (k, C) => sku[k].laborMode === 'calc' ? { f: `${lref.piece}+${lref.revPct}*${C}{EFF}`, v: R[k].laborBag, s: S.money } : sku[k].laborMode === 'none' ? { v: 0, s: S.money } : { v: +sku[k].labor, s: S.inMoney });
    line('alloc', '人工归属比例', '团队月人工中算到本 SKU 的比例（用人工表时生效）', () => ({ v: (sku.laborAlloc === undefined || sku.laborAlloc === '' ? (+L.alloc || 0) : +sku.laborAlloc) / 100, s: S.inPct }));
    line('cold', '冷链（元/袋）', '', k => ({ v: +sku[k].cold, s: S.inMoney }));
    line('loss', '损耗（元/袋）', '', k => ({ v: +sku[k].loss, s: S.inMoney }));
    line('rate', '渠道扣点率', '', k => ({ v: (+sku[k].rate) / 100, s: S.inPct }));
    line('rbase', '扣点基数', '填“有效售价”或“标价”，按合同', k => ({ v: sku[k].rateBase === 'list' ? '标价' : '有效售价', s: S.inTxt }));
    line('fee', '活动费（元）', '可归属本场景的一次性费用', k => ({ v: +(sku[k].fee || 0), s: S.inMoney }));
    // 自定义项目
    const items = sku.items || [], itemRow = {}, orderRow = {};
    const NAT = { cost: '成本', subsidy: '补贴', ref: '参考' }, MET = { bag: '按袋', order: '按单', pct: '按比例', event: '按活动' };
    if (items.length) sec('自定义项目');
    items.forEach((it, i) => {
      const ref0 = C => ({ 标价: C + row.list, 商家折扣: C + row.disc, 有效售价: C + '{EFF}', 销量: C + row.vol, 采购: C + '{PURBAG}', 人工: C + row.labor, 冷链: C + row.cold, 损耗: C + row.loss, 扣点率: C + row.rate, 活动费: C + row.fee });
      line('it' + i, (it.name || '自定义' + (i + 1)) + '（' + NAT[it.nature] + '·' + MET[it.method] + '）', it.src === 'formula' ? '公式：' + (it.vals.base || '') : (it.method === 'pct' ? '比例' : ''), (k, C) => {
        const res = R[k].items[i] || { value: 0 }, isPct = it.method === 'pct';
        const v = isPct ? res.value / 100 : res.value, st = isPct ? (it.src === 'formula' ? S.pct : S.inPct) : (it.src === 'formula' ? S.money : S.inMoney);
        if (it.src === 'formula') {
          const ref = ref0(C); items.slice(0, i).forEach((p, j) => { if (p.name) ref[p.name] = C + itemRow[j]; });
          let f; try { f = E.toExcel(E.parse(it.vals[k] || '0'), ref); } catch (e) { f = '0'; }
          return { f: isPct ? `(${f})/100` : f, v, s: st };
        }
        return { v, s: st };
      });
      itemRow[i] = row['it' + i];
      if (it.method === 'order') { line('io' + i, '  └ 每单袋数', '', () => ({ v: +it.bagsPerOrder || 0, s: S.inInt })); orderRow[i] = row['io' + i]; }
    });
    const calc = state.labor && keys.some(k => sku[k].laborMode === 'calc');
    line('lfix', '人工固定（元）', '团队月人工 × 人工归属比例', (k, C) => sku[k].laborMode === 'calc' ? { f: `${lref.team}*${C}${row.alloc}`, v: R[k].laborFixed, s: S.money } : { v: 0, s: S.money });
    line('tier', '阶梯激励（元）', '来自人工表：达到门槛的奖励', (k, C) => {
      if (sku[k].laborMode !== 'calc' || !lref.tiers.length) return { v: 0, s: S.money };
      let f = '0'; lref.tiers.forEach(t => { f = `IF(${C}${row.vol}>=${t.t},${t.a},${f})`; });
      return { f, v: R[k].tier, s: S.money };
    });
    sec('计算结果（公式，自动重算）');
    line('eff', '有效售价（元/袋）', '标价 − 商家承担折扣', (k, C) => ({ f: `${C}${row.list}-${C}${row.disc}`, v: R[k].eff, s: S.money }));
    line('purbag', '采购（元/袋）', '', (k, C) => ({ f: `${C}${row.pur}/${C}${row.box}`, v: R[k].purchase, s: S.money }));
    line('rev', '销售收入', '有效售价 × 销量', (k, C) => ({ f: `${C}${row.eff}*${C}${row.vol}`, v: R[k].rev, s: S.keyMoney }));
    line('comm', '渠道扣点（元/袋）', '', (k, C) => ({ f: `${C}${row.rate}*IF(${C}${row.rbase}="标价",${C}${row.list},${C}${row.eff})`, v: R[k].commission, s: S.money }));
    const per = (C, it, i) => it.method === 'bag' ? C + itemRow[i] : it.method === 'order' ? `${C}${itemRow[i]}/${C}${orderRow[i]}` : it.method === 'pct' ? `${C}${itemRow[i]}*${it.pctBase === 'list' ? C + row.list : C + row.eff}` : null;
    const sumOf = (C, nat, evt) => { const xs = items.map((it, i) => it.nature === nat ? (evt ? (it.method === 'event' ? C + itemRow[i] : null) : per(C, it, i)) : null).filter(Boolean); return xs.length ? xs.join('+') : '0'; };
    line('cbag', '其他每袋成本', '自定义成本（按袋/按单/按比例）', (k, C) => ({ f: sumOf(C, 'cost', false), v: R[k].costBag, s: S.money }));
    line('sbag', '每袋补贴', '自定义补贴（按袋/按单/按比例）', (k, C) => ({ f: sumOf(C, 'subsidy', false), v: R[k].subBag, s: S.money }));
    line('unit', '单位贡献（元/袋）', '有效售价 − 各项每袋成本 + 每袋补贴', (k, C) => ({ f: `${C}${row.eff}-${C}${row.purbag}-${C}${row.labor}-${C}${row.cold}-${C}${row.loss}-${C}${row.comm}-${C}${row.cbag}+${C}${row.sbag}`, v: R[k].unit, s: S.keyMoney }));
    line('contrib', '贡献利润', '单位贡献 × 销量', (k, C) => ({ f: `${C}${row.unit}*${C}${row.vol}`, v: R[k].contribution, s: S.money }));
    line('fixed', '固定费用', '活动费 + 人工固定 + 阶梯激励 + 按活动计的成本', (k, C) => ({ f: `${C}${row.fee}+${C}${row.lfix}+${C}${row.tier}+${sumOf(C, 'cost', true)}`, v: R[k].fixed, s: S.money }));
    line('sevt', '活动补贴', '按活动总额计的补贴', (k, C) => ({ f: sumOf(C, 'subsidy', true), v: R[k].subEvt, s: S.money }));
    line('profit', '预计场景利润', '贡献利润 − 固定费用 + 活动补贴', (k, C) => ({ f: `${C}${row.contrib}-${C}${row.fixed}+${C}${row.sevt}`, v: R[k].profit, s: S.keyMoney }));
    line('margin', '场景利润率', '', (k, C) => ({ f: `IF(${C}${row.rev}=0,"—",${C}${row.profit}/${C}${row.rev})`, v: R[k].margin === null ? '—' : R[k].margin, s: S.pct }));
    line('urate', '单位贡献率', '单位贡献 ÷ 有效售价', (k, C) => ({ f: `IF(${C}${row.eff}=0,"—",${C}${row.unit}/${C}${row.eff})`, v: R[k].unitRate === null ? '—' : R[k].unitRate, s: S.pct }));
    const mv = (k, C, T) => {
      const U = C + row.unit, F0 = `(${C}${row.fixed}-${C}${row.tier})`, Sb = C + row.sevt;
      const tiers = sku[k].laborMode === 'calc' ? lref.tiers : [];
      if (!tiers.length) return `IF(${U}<=0,"无可用",MAX(0,ROUNDUP((${T}+${F0}-${Sb})/${U},0)))`;
      const cand = (start, bonus) => `MAX(${start},ROUNDUP((${T}+${F0}+${bonus}-${Sb})/${U},0))`;
      const segs = [{ s: '0', b: '0' }].concat(tiers.map(t => ({ s: t.t, b: t.a })));
      let f = cand(segs[segs.length - 1].s, segs[segs.length - 1].b);
      for (let i = segs.length - 2; i >= 0; i--) f = `IF(${cand(segs[i].s, segs[i].b)}<${segs[i + 1].s},${cand(segs[i].s, segs[i].b)},${f})`;
      return `IF(${U}<=0,"无可用",${f})`;
    };
    line('be', '盈亏平衡销量（袋）', '利润 ≥ 0 的最低销量', (k, C) => ({ f: mv(k, C, '0'), v: R[k].be === null ? '无可用' : R[k].be, s: S.keyInt }));
    const eq = R.promo.minVolume(R.base.profit);
    line('eq', '追平基准利润所需销量（袋）', '促销方案达到基准方案利润的最低销量', (k, C) => k === 'promo' ? { f: mv(k, C, `B${row.profit}`), v: eq === null ? '无可用' : eq, s: S.keyInt } : { v: '—' });
    line('gap', '距离追平还差（袋）', '负数表示已超过', (k, C) => k === 'promo' ? { f: `IF(ISNUMBER(C${row.eq}),C${row.eq}-C${row.vol},"—")`, v: eq === null ? '—' : eq - R.promo.V, s: S.int } : { v: '—' });
    // 衍生指标
    const ms = sku.metrics || [];
    if (ms.length) {
      sec('自定义指标');
      ms.forEach((m, i) => line('m' + i, m.name || '指标' + (i + 1), '公式：' + m.expr, (k, C) => {
        const ref = { 标价: C + row.list, 商家折扣: C + row.disc, 有效售价: C + row.eff, 销量: C + row.vol, 采购: C + row.purbag, 人工: C + row.labor, 冷链: C + row.cold, 损耗: C + row.loss, 扣点率: C + row.rate, 活动费: C + row.fee, 收入: C + row.rev, 单位贡献: C + row.unit, 贡献利润: C + row.contrib, 固定费用: C + row.fixed, 场景利润: C + row.profit, 利润率: `IF(ISNUMBER(${C}${row.margin}),${C}${row.margin},0)`, 单位贡献率: `IF(ISNUMBER(${C}${row.urate}),${C}${row.urate},0)` };
        items.forEach((p, j) => { if (p.name) ref[p.name] = C + itemRow[j]; });
        ms.slice(0, i).forEach((p, j) => { if (p.name) ref[p.name] = C + row['m' + j]; });
        const mres = R[k].metrics[i] || {};
        try { return { f: E.toExcel(E.parse(m.expr), ref), v: mres.value, s: S.txt }; } catch (e) { return { v: '公式有误' }; }
      }));
    }
    // 回填占位
    Object.values(sh.rows).forEach(rw => Object.values(rw).forEach(c => { if (c.f) c.f = c.f.replace(/\{EFF\}/g, row.eff).replace(/\{PURBAG\}/g, row.purbag); }));
    return { sh, row, R };
  }

  function build(state, E, analyses) {
    const used = new Set(['说明', '汇总', '人工', '分析提示']);
    const lab = laborSheet(state.labor, E);
    const skus = state.skus.map(s => ({ s, name: sheetName(s.name, used) }));
    const sheets = skus.map(x => Object.assign(skuSheet(x.name, x.s, state, E, lab.ref), { name: x.name }));
    // 汇总
    const sum = new Sheet('汇总'); sum.widths = [24, 16, 16, 16, 16];
    sum.set(0, 1, { v: '各 SKU 汇总（同一渠道、同一月份）', s: S.title });
    ['SKU', '基准收入', '基准利润', '促销收入', '促销利润'].forEach((h, c) => sum.set(c, 2, { v: h, s: S.head }));
    let r = 3;
    sheets.forEach(x => {
      const n = q(x.name); sum.set(0, r, { v: x.name });
      sum.set(1, r, { f: `${n}!B${x.row.rev}`, v: x.R.base.rev, s: S.money }); sum.set(2, r, { f: `${n}!B${x.row.profit}`, v: x.R.base.profit, s: S.money });
      sum.set(3, r, { f: `${n}!C${x.row.rev}`, v: x.R.promo.rev, s: S.money }); sum.set(4, r, { f: `${n}!C${x.row.profit}`, v: x.R.promo.profit, s: S.money }); r++;
    });
    const f1 = 3, l1 = r - 1, tot = k => sheets.reduce((a, x) => a + x.R[k.split('.')[0]][k.split('.')[1]], 0);
    sum.set(0, r, { v: '各 SKU 合计', s: S.sec }); ['B', 'C', 'D', 'E'].forEach((c, i) => sum.set(i + 1, r, { f: `SUM(${c}${f1}:${c}${l1})`, v: tot(['base.rev', 'base.profit', 'promo.rev', 'promo.profit'][i]), s: S.keyMoney })); const tr = r; r++;
    sum.set(0, r, { v: '共同费用（只登记一次）' }); sum.set(2, r, { v: +state.commonFee || 0, s: S.inMoney }); sum.set(4, r, { v: +state.commonFee || 0, s: S.inMoney }); const cr = r; r++;
    sum.set(0, r, { v: '汇总利润' }); sum.set(2, r, { f: `C${tr}-C${cr}`, v: tot('base.profit') - (+state.commonFee || 0), s: S.keyMoney }); sum.set(4, r, { f: `E${tr}-E${cr}`, v: tot('promo.profit') - (+state.commonFee || 0), s: S.keyMoney }); const pr = r; r++;
    sum.set(0, r, { v: '汇总利润率' });
    const bm = tot('base.rev') ? (tot('base.profit') - (+state.commonFee || 0)) / tot('base.rev') : '—', pm = tot('promo.rev') ? (tot('promo.profit') - (+state.commonFee || 0)) / tot('promo.rev') : '—';
    sum.set(2, r, { f: `IF(B${tr}=0,"—",C${pr}/B${tr})`, v: bm, s: S.pct }); sum.set(4, r, { f: `IF(D${tr}=0,"—",E${pr}/D${tr})`, v: pm, s: S.pct }); r += 2;
    sum.set(0, r, { v: '盈亏平衡销量不跨 SKU 相加；共同费用不分摊到单个 SKU。', s: S.note });
    // 分析提示
    const an = new Sheet('分析提示'); an.widths = [18, 100];
    an.set(0, 1, { v: '导出时的分析提示（文字不会随 Excel 修改自动更新）', s: S.title });
    ['SKU', '提示'].forEach((h, c) => an.set(c, 2, { v: h, s: S.head })); r = 3;
    skus.forEach((x, i) => (analyses[i] || []).forEach(a => { an.set(0, r, { v: x.name }); an.set(1, r, { v: a.text }); r++; }));
    // 说明
    const doc = new Sheet('说明'); doc.widths = [110];
    ['三全食品 利润与人工成本测算（由网页计算器导出）', '', '· 黄色格子是输入，可以直接修改；其余格子是公式，会自动重算。',
      '· 每个 SKU 一张表：上半部分是参数，下半部分是计算结果。“汇总”表合计各 SKU；“人工”表是人工成本计算器。',
      '· 盈亏平衡销量和追平销量已考虑阶梯激励。阶梯门槛请保持从小到大。', '· “分析提示”是导出那一刻的文字结论，修改数字后请回到网页重新生成。',
      '· 演示数据全部为假设，请替换为实际采购单、结算单数据。', '', '导出时间：' + new Date().toLocaleString('zh-CN')].forEach((t, i) => doc.set(0, i + 1, { v: t, s: i === 0 ? S.title : S.txt }));
    return book([doc, sum].concat(sheets.map(x => x.sh), [lab.sh, an]));
  }
  return { build, zip, book, Sheet, S, setNoCache: b => { NOCACHE = b; } };
})();
if (typeof module !== 'undefined') module.exports = XLSX;
