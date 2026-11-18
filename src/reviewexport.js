/* ===== 复盘结果导出：Excel（数据表）与 XMind（括号图，结构同手工思维导图） ===== */
const ReviewExport = (() => {
  const uuid = () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); });

  function xmind(tree) {
    const topic = (n, root) => {
      const t = { id: uuid(), class: 'topic', title: n.title };
      if (root) t.structureClass = 'org.xmind.ui.brace.right';
      const nl = n.title.indexOf('\n');
      if (nl > 0) t.attributedTitle = [{ text: n.title.slice(0, nl), 'fo:font-weight': '700' }, { text: n.title.slice(nl) }];
      if (n.children && n.children.length) t.children = { attached: n.children.map(c => topic(c)) };
      return t;
    };
    const content = [{ id: uuid(), revisionId: uuid(), class: 'sheet', title: '损益复盘', rootTopic: topic(tree, true) }];
    return XLSX.zip([
      { name: 'content.json', data: JSON.stringify(content) },
      { name: 'metadata.json', data: JSON.stringify({ dataStructureVersion: '3', creator: { name: '三全利润与人工测算器', version: '2' }, layoutEngineVersion: '5' }) },
      { name: 'manifest.json', data: JSON.stringify({ 'file-entries': { 'content.json': {}, 'metadata.json': {} } }) }
    ]);
  }

  function excel(results, edits, meta) {
    const { Sheet, S, book } = XLSX, V = Review, sheets = [];
    const doc = new Sheet('说明'); doc.widths = [100];
    ['损益复盘（由导入数据自动生成）', '', '分析月份：' + meta.month + '　累计范围：' + meta.cumRange + '　客户列：' + meta.dim, '金额单位：万元；费率 = 费用 ÷ 收入；差异 = 本期费率 − 同期费率。',
      '这里是计算结果（数值），不是公式。原始数据请回到系统导出的明细核对。', '导出时间：' + new Date().toLocaleString('zh-CN')].forEach((t, i) => doc.set(0, i + 1, { v: t, s: i ? S.txt : S.title }));
    sheets.push(doc);
    const used = new Set(['说明']);
    results.forEach(res => {
      let nm = String(res.cust).replace(/[\[\]:*?\/\\']/g, '').slice(0, 28) || '客户'; let x = nm, k = 2; while (used.has(x)) x = nm + k++; used.add(x);
      const sh = new Sheet(x); sh.widths = [26, 14, 14, 14, 14, 3, 26, 14, 14, 14, 14]; let r = 1;
      sh.set(0, r++, { v: res.cust + ' 损益复盘（' + meta.month + '）', s: S.title }); r++;
      ['当月', '累计'].forEach(key => {
        const s = res[key];
        sh.set(0, r, { v: key + '（' + s.months[0] + (s.months.length > 1 ? ' ~ ' + s.months[s.months.length - 1] : '') + '）', s: S.sec }); for (let c = 1; c < 5; c++) sh.set(c, r, { v: '', s: S.sec }); r++;
        ['项目', '本期', '同期', '差额', '增长率'].forEach((h, c) => sh.set(c, r, { v: h, s: S.head })); r++;
        [['收入（万元）', s.rev], ['利润（万元）', s.profit]].forEach(([l, o]) => { sh.set(0, r, { v: l }); sh.set(1, r, { v: o.a, s: S.wan }); sh.set(2, r, { v: o.b, s: S.wan }); sh.set(3, r, { v: o.d, s: S.wan }); sh.set(4, r, o.b > 0 ? { v: o.a / o.b - 1, s: S.pct } : { v: '—' }); r++; });
        r++;
        if (s.fees.length) {
          ['费用率', '本期', '同期', '差异'].forEach((h, c) => sh.set(c, r, { v: h, s: S.head })); r++;
          s.fees.concat(s.tp5 ? [s.tp5] : []).forEach(f => { sh.set(0, r, { v: f.label }); sh.set(1, r, { v: f.a, s: S.pct }); sh.set(2, r, { v: f.b, s: S.pct }); sh.set(3, r, { v: f.d, s: S.pct }); r++; });
          s.cpBreak.forEach(f => { sh.set(0, r, { v: '  CP·' + f.name }); sh.set(1, r, { v: f.a, s: S.pct }); sh.set(2, r, { v: f.b, s: S.pct }); sh.set(3, r, { v: f.d, s: S.pct }); r++; });
          r++;
          ['费用子项（变化最大）', '本期费率', '同期费率', '差异', '本期金额'].forEach((h, c) => sh.set(c, r, { v: h, s: S.head })); r++;
          Object.entries(s.drill).forEach(([fk, list]) => list.slice(0, 3).forEach(d => { sh.set(0, r, { v: fk.slice(1) + '·' + d.name }); sh.set(1, r, { v: d.a, s: S.pct }); sh.set(2, r, { v: d.b, s: S.pct }); sh.set(3, r, { v: d.d, s: S.pct }); sh.set(4, r, { v: d.amtA, s: S.wan }); r++; }));
          r++;
        }
        if (s.hasProduct) {
          ['品类', '本期收入', '同期收入', '收入差额', '利润差额'].forEach((h, c) => sh.set(c, r, { v: h, s: S.head })); r++;
          s.cats.slice().sort((p, q) => p.rev - q.rev).forEach(c => { sh.set(0, r, { v: c.name }); sh.set(1, r, { v: c.revA, s: S.wan }); sh.set(2, r, { v: c.revB, s: S.wan }); sh.set(3, r, { v: c.rev, s: S.wan }); sh.set(4, r, { v: c.profit, s: S.wan }); r++; });
          r++;
        }
        if (s.hasSku) {
          ['规格（收入变化前 10）', '本期收入', '同期收入', '收入差额', '利润差额'].forEach((h, c) => sh.set(c, r, { v: h, s: S.head })); r++;
          s.skus.slice().sort((p, q) => Math.abs(q.rev) - Math.abs(p.rev)).slice(0, 10).forEach(c => { sh.set(0, r, { v: c.name }); sh.set(1, r, { v: c.revA, s: S.wan }); sh.set(2, r, { v: c.revB, s: S.wan }); sh.set(3, r, { v: c.rev, s: S.wan }); sh.set(4, r, { v: c.profit, s: S.wan }); r++; });
          r++;
        }
        sh.set(0, r, { v: '结论', s: S.head }); r++;
        const e = Object.assign(Review.splitConclusions(s), (edits[res.cust] || {})[key] || {});
        [['产品', e.product], ['费用', e.fee]].forEach(([l, t]) => { sh.set(0, r, { v: l + '：' + (t || ''), s: S.note }); r++; });
        r += 2;
      });
      sheets.push(sh);
    });
    return book(sheets);
  }
  return { xmind, excel };
})();
if (typeof module !== 'undefined') module.exports = ReviewExport;
