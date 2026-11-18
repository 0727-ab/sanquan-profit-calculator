/* ===== 损益复盘引擎：读系统导出的“本期 / 同期”明细，生成当月 / 累计对比分析 ===== */
const Review = (() => {
  // 字段清单：必需字段缺一个就拒绝导入；可选字段缺了，对应分析隐藏，绝不当 0 计算
  const NEED = ['日历年/月', '收入', '=产品考核净利润'];
  const OPTIONAL = [['-产品成本', '产品成本'], ['=产品毛利额', '毛利额（毛利率）'], ['-AP', 'AP 费用'], ['-TP', 'TP 费用'], ['-CP', 'CP 费用'], ['-PP', 'PP 费用'], ['-MP', 'MP 费用'], ['-FP', 'FP 费用'],
    ['一级品类', '一级品类（产品维度）'], ['四级规格', '四级规格（规格明细）'], ['-产品组/研发费用合计', '产品组/研发费用'], ['-AE', '公摊费用 AE'], ['所得税', '所得税']];
  const FEES = ['-AP', '-TP', '-CP', '-PP', '-MP', '-FP'];
  const DIMS = ['四级渠道', '三级渠道', '二级渠道', '一级渠道', '考核组织名称', '战区', '大区'];
  // 数据类型模板：每种报表一条，写明需要哪几张表、表头特征。以后新增类型在这里加一条即可。
  const SOURCES = [
    { id: 'pl-yoy', name: '系统损益明细（本期 / 同期）', sheets: ['本期', '同期'], headerHas: ['日历年/月', '收入', '=产品考核净利润'],
      desc: '从系统导出的损益明细，每行是“月份 × 客户 × 规格”，含收入、成本和各项费用；用于当月 / 累计同比复盘。' }
  ];
  function detect(sheetNames) {
    for (const src of SOURCES) { const got = src.sheets.map(n => sheetNames.find(x => x.trim() === n)); if (got.every(Boolean)) return { src, sheets: got }; }
    return null;
  }

  /* ---------- 数据集 ---------- */
  function dataset(rows, label) {
    const hi = rows.findIndex(r => r && r.includes('日历年/月') && r.includes('收入'));
    if (hi < 0) throw new Error(`“${label}”表里找不到表头（需要有“日历年/月”和“收入”两列）`);
    const hdr = rows[hi].map(x => x == null ? '' : String(x).trim()), col = {};
    hdr.forEach((h, i) => { if (h && !(h in col)) col[h] = i; });
    const miss = NEED.filter(n => !(n in col));
    if (miss.length) throw new Error(`“${label}”表缺少必需的列：${miss.join('、')}。请确认导出的是系统损益明细。`);
    if (!DIMS.some(d => d in col)) throw new Error(`“${label}”表里没有可以区分客户的列（${DIMS.join(' / ')}）`);
    const iMonth = col['日历年/月'];
    const data = rows.slice(hi + 1).filter(r => r && r[iMonth] != null && String(r[iMonth]).trim() !== '').map(r => { r[iMonth] = normMonth(r[iMonth]); return r; });
    // 每项费用的子项：从该费用列到下一个“-”或“=”开头的列之间
    const groups = {};
    FEES.forEach(f => {
      if (!(f in col)) { groups[f] = []; return; }
      const s = col[f]; let e = s + 1; while (e < hdr.length && !/^[-=]/.test(hdr[e]) && !DIMS.includes(hdr[e])) e++;
      const sub = [];
      for (let i = s + 1; i < e; i++) { if (/^\d+-/.test(hdr[i])) sub.push({ i, name: hdr[i], level: 1 }); else if (/^\d+\.\d+-/.test(hdr[i])) sub.push({ i, name: hdr[i], level: 2 }); }
      groups[f] = sub;
    });
    const numCols = []; for (let i = col['收入']; i <= col['=产品考核净利润']; i++) numCols.push(i);
    const months = [...new Set(data.map(r => r[iMonth]))].sort();
    const dims = DIMS.filter(d => d in col);
    const has = n => n in col;
    // 导入检查
    const known = new Set([...NEED, ...OPTIONAL.map(o => o[0]), ...DIMS, ...Object.values(groups).flat().map(x => x.name)]);
    const check = { rows: data.length, found: OPTIONAL.filter(o => has(o[0])).map(o => o[1]), missing: OPTIONAL.filter(o => !has(o[0])).map(o => o[1]),
      ignored: hdr.filter(h => h && !known.has(h) && !/^\d+(\.\d+)?-/.test(h)).length, balance: null };
    const parts = ['-产品成本', '-AP', '-TP', '-CP', '-PP', '-MP', '-FP', '-产品组/研发费用合计', '-AE', '所得税'];
    if (parts.every(has)) {
      let rev = 0, gap = 0;
      for (const r of data) { const v = n => +r[col[n]] || 0; rev += Math.abs(v('收入')); gap += Math.abs(v('收入') - parts.reduce((a, n) => a + v(n), 0) - v('=产品考核净利润')); }
      check.balance = { ok: gap <= Math.max(0.01, rev * 0.001), gap };
    }
    return { label, hdr, col, data, groups, numCols, months, dims, has, check };
  }
  function normMonth(v) {
    if (typeof v === 'number') { // Excel 日期序号
      if (v > 20000) { const d = new Date(Math.round((v - 25569) * 864e5)); return d.getUTCFullYear() + '/' + String(d.getUTCMonth() + 1).padStart(2, '0'); }
      return String(v);
    }
    const m = String(v).trim().match(/^(\d{4})\D?(\d{1,2})/); return m ? m[1] + '/' + m[2].padStart(2, '0') : String(v).trim();
  }
  const prevYear = m => { const [y, mm] = m.split('/'); return (+y - 1) + '/' + mm; };
  function values(ds, dim) { return dim in ds.col ? [...new Set(ds.data.map(r => r[ds.col[dim]]).filter(v => v != null && v !== ''))].sort((a, b) => String(a).localeCompare(String(b), 'zh')) : []; }

  /* ---------- 汇总 ---------- */
  function sums(ds, f) {
    const s = new Float64Array(ds.hdr.length), cats = new Map(), skus = new Map();
    const iM = ds.col['日历年/月'], iD = f.dim && f.dim in ds.col ? ds.col[f.dim] : -1, iC = ds.col['一级品类'] ?? -1, iS = ds.col['四级规格'] ?? -1, iG = ds.col['收入'], iP = ds.col['=产品考核净利润'];
    for (const r of ds.data) {
      if (!f.months.has(r[iM])) continue;
      if (iD >= 0 && f.cust !== '全部' && r[iD] !== f.cust) continue;
      if (f.sku && (iS < 0 || r[iS] !== f.sku)) continue;
      for (const i of ds.numCols) { const v = +r[i]; if (v) s[i] += v; }
      const add = (m, k) => { const o = m.get(k) || { rev: 0, profit: 0 }; o.rev += +r[iG] || 0; o.profit += +r[iP] || 0; m.set(k, o); };
      if (iC >= 0) add(cats, r[iC] || '未分类'); if (iS >= 0) add(skus, r[iS] || '未知规格');
    }
    return { s, cats, skus };
  }

  /* ---------- 一段（当月 / 累计）的分析 ---------- */
  function section(dc, dp, f, opts) {
    const A = sums(dc, { dim: f.dim, cust: f.cust, months: new Set(f.cur) }), B = sums(dp, { dim: f.dim, cust: f.cust, months: new Set(f.pri) });
    const g = (X, ds, n) => n in ds.col ? (X.s[ds.col[n]] || 0) : 0;
    const rev = { a: g(A, dc, '收入'), b: g(B, dp, '收入') }, profit = { a: g(A, dc, '=产品考核净利润'), b: g(B, dp, '=产品考核净利润') };
    rev.d = rev.a - rev.b; profit.d = profit.a - profit.b;
    const rate = (X, ds, n) => { const G = g(X, ds, '收入'); return G && n in ds.col ? g(X, ds, n) / G : null; };
    const feeRow = (label, n) => { const a = rate(A, dc, n), b = rate(B, dp, n); return { label, key: n, a, b, d: a !== null && b !== null ? a - b : null }; };
    const fees = [['毛利率', '=产品毛利额']].concat(FEES.map(n => [n.slice(1), n])).filter(([, n]) => n in dc.col && n in dp.col).map(([l, n]) => feeRow(l, n));
    const tp5n = dc.groups['-TP'].find(x => /供价差异/.test(x.name) && x.level === 1);
    const tp5 = tp5n ? feeRow('TP5', tp5n.name) : null;
    const subRate = (X, ds, name) => { const i = ds.hdr.indexOf(name); const G = g(X, ds, '收入'); return i >= 0 && G ? (X.s[i] || 0) / G : 0; };
    const subAmt = (X, ds, name) => { const i = ds.hdr.indexOf(name); return i >= 0 ? (X.s[i] || 0) : 0; };
    const cpBreak = dc.groups['-CP'].filter(x => x.level === 1).map(x => ({ name: clean(x.name), a: subRate(A, dc, x.name), b: subRate(B, dp, x.name) })).map(x => Object.assign(x, { d: x.a - x.b }));
    const drill = {};
    FEES.forEach(n => {
      drill[n] = dc.groups[n].filter(x => x.level === 2).map(x => ({ name: clean(x.name), a: subRate(A, dc, x.name), b: subRate(B, dp, x.name), amtA: subAmt(A, dc, x.name), amtB: subAmt(B, dp, x.name) }))
        .map(x => Object.assign(x, { d: x.a - x.b })).filter(x => Math.abs(x.d) > 1e-6).sort((p, q) => Math.abs(q.d) - Math.abs(p.d)).slice(0, 5);
    });
    const diffMap = (ma, mb) => { const out = []; new Set([...ma.keys(), ...mb.keys()]).forEach(k => { const a = ma.get(k) || { rev: 0, profit: 0 }, b = mb.get(k) || { rev: 0, profit: 0 }; out.push({ name: k, revA: a.rev, revB: b.rev, rev: a.rev - b.rev, profitA: a.profit, profitB: b.profit, profit: a.profit - b.profit }); }); return out; };
    const cats = diffMap(A.cats, B.cats), skus = diffMap(A.skus, B.skus);
    const sec = { rev, profit, fees, tp5, cpBreak, drill, cats, skus, months: f.cur, monthsPri: f.pri, empty: !rev.a && !rev.b,
      hasProduct: '一级品类' in dc.col && '一级品类' in dp.col, hasSku: '四级规格' in dc.col && '四级规格' in dp.col, missingPri: f.pri.filter(m => !dp.months.includes(m)) };
    sec.thr = Math.max(opts.minWan, Math.abs(rev.b) * 0.01, Math.abs(rev.a) * 0.01);
    sec.conclusions = conclude(sec, opts);
    return sec;
  }
  const clean = n => n.replace(/^\d+(\.\d+)?-/, '').trim();

  function analyze(dc, dp, f, opts) {
    opts = Object.assign({ minWan: 1, drillPP: 0.01, topSku: 5 }, opts || {});
    const cumMonths = dc.months.filter(m => m >= f.cumStart && m <= f.month);
    return {
      cust: f.cust, month: f.month, dim: f.dim,
      当月: section(dc, dp, { dim: f.dim, cust: f.cust, cur: [f.month], pri: [prevYear(f.month)] }, opts),
      累计: section(dc, dp, { dim: f.dim, cust: f.cust, cur: cumMonths, pri: cumMonths.map(prevYear) }, opts),
      opts
    };
  }

  /* ---------- 格式 ---------- */
  const wan = v => { const r = Math.round(v); return (r === 0 && v !== 0 ? (+v.toFixed(1)) : r).toLocaleString('zh-CN') + '万'; };
  const swan = v => (v > 0 ? '+' : '') + wan(v);
  const pc = v => v === null || v === undefined ? '—' : (v * 100).toFixed(1) + '%';
  const spc = v => v === null || v === undefined ? '—' : (v > 0 ? '+' : '') + (v * 100).toFixed(1) + '%';
  const growth = (a, b) => b > 0 ? spc(a / b - 1).replace('%', '%') : null;
  function profitPhrase(p) {
    if (p.b > 0 && p.a >= 0) return '增长率' + growth(p.a, p.b);
    if (p.b < 0 && p.a < 0) return p.d >= 0 ? '减亏' + wan(p.d) : '亏损扩大' + wan(-p.d);
    if (p.b <= 0 && p.a > 0) return '扭亏为盈，' + swan(p.d);
    return '由盈转亏，' + swan(p.d);
  }
  // 门槛：变化至少 1 万，且不小于同期收入的 1%（避免零碎变化刷屏）
  const lst = (arr, key, sign, thr = 1) => arr.filter(x => sign > 0 ? x[key] >= thr - 0.5 && Math.round(x[key]) >= 1 : x[key] <= -(thr - 0.5) && Math.round(x[key]) <= -1).sort((p, q) => sign > 0 ? q[key] - p[key] : p[key] - q[key]);
  function headline(sec) {
    return `收入：本期${wan(sec.rev.a)}，同期${wan(sec.rev.b)}，${sec.rev.b > 0 ? '增长率' + growth(sec.rev.a, sec.rev.b) : '差额' + swan(sec.rev.d)}\n利润：本期${wan(sec.profit.a)}，同期${wan(sec.profit.b)}，${profitPhrase(sec.profit)}`;
  }
  function productText(sec, opts) {
    if (!sec.hasProduct) return `收入（${swan(sec.rev.d)}），利润（${swan(sec.profit.d)}）\n（数据里没有“一级品类”列，品类分析不显示）`;
    const n = (opts && opts.topSku) || 5;
    const f = (xs, key) => xs.map(x => x.name + swan(x[key])).join('、') || '无';
    const up = lst(sec.cats, 'rev', 1, sec.thr), dn = lst(sec.cats, 'rev', -1, sec.thr), skUp = lst(sec.skus, 'rev', 1, sec.thr).slice(0, 3), skDn = lst(sec.skus, 'rev', -1, sec.thr).slice(0, n);
    const pup = lst(sec.cats, 'profit', 1, sec.thr), pdn = lst(sec.cats, 'profit', -1, sec.thr), pskDn = lst(sec.skus, 'profit', -1, sec.thr).slice(0, 3);
    return `收入（${swan(sec.rev.d)}）：\n增长品类：${f(up, 'rev')}` + (skUp.length ? `\n（${f(skUp, 'rev')}）` : '') +
      `\n下滑品类：${f(dn, 'rev')}` + (skDn.length ? `\n（${f(skDn, 'rev')}）` : '') +
      `\n\n利润（${swan(sec.profit.d)}）：\n增长品类：${f(pup, 'profit')}\n下滑品类：${f(pdn, 'profit')}` + (pskDn.length ? `\n（${f(pskDn, 'profit')}）` : '');
  }
  function feeText(sec) {
    if (!sec.fees.length) return '（数据里没有费用列，费用分析不显示）';
    let t = sec.fees.map(x => `${x.label}：本期${pc(x.a)}，同期${pc(x.b)}，差异${spc(x.d)}`).join('\n');
    t += '\n\n重点分析：';
    if (sec.tp5) t += `\nTP5：本期${pc(sec.tp5.a)}，同期${pc(sec.tp5.b)}，差异${spc(sec.tp5.d)}`;
    const cp = sec.fees.find(x => x.key === '-CP'), parts = sec.cpBreak.filter(x => Math.abs(x.d) >= 0.0005).sort((p, q) => Math.abs(q.d) - Math.abs(p.d)).slice(0, 3);
    if (cp) t += `\nCP${spc(cp.d)}` + (parts.length ? '，其中' + parts.map(x => x.name + spc(x.d)).join('、') : '');
    sec.fees.filter(x => x.key !== '=产品毛利额' && x.key !== '-CP' && x.d !== null && Math.abs(x.d) >= 0.01).forEach(x => {
      const top = (sec.drill[x.key] || []).filter(s => Math.sign(s.d) === Math.sign(x.d)).slice(0, 2);
      if (top.length) t += `\n${x.label}${spc(x.d)}，主要来自` + top.map(s => s.name + spc(s.d)).join('、');
    });
    return t;
  }

  /* ---------- 结论规则 ---------- */
  function conclude(sec, opts) {
    if (sec.empty) return ['本期和同期都没有数据。'];
    const out = [], rd = sec.rev.d, pd = sec.profit.d, r = sec.rev;
    const g = r.b > 0 ? r.a / r.b - 1 : null;
    if (rd < 0 && pd > 0) out.push(`收入减少${wan(-rd)}，利润反而改善${wan(pd)}，属于“量缩、利改善”的结构，减亏不是靠销量提升带来的`);
    else if (rd > 0 && pd < 0) out.push(`收入增加${wan(rd)}，利润却减少${wan(-pd)}，增收不增利`);
    else if (rd < 0 && pd < 0) out.push(`收入减少${wan(-rd)}，利润减少${wan(-pd)}，量利双降`);
    else if (rd > 0 && pd > 0) out.push(`收入增加${wan(rd)}，利润增加${wan(pd)}，量利双增`);
    const up = lst(sec.cats, 'rev', 1, sec.thr), skUp = lst(sec.skus, 'rev', 1, sec.thr);
    if (!up.length) out.push('没有收入增长的品类' + (skUp.length ? `，仅${skUp[0].name}增长${wan(skUp[0].rev)}` : ''));
    else if (skUp.length) out.push(`收入增量主要来自${skUp[0].name}（${swan(skUp[0].rev)}）`);
    const cannibal = sec.skus.filter(x => x.rev >= sec.thr && x.profit <= -Math.max(3, sec.thr)).sort((p, q) => p.profit - q.profit)[0];
    if (cannibal) out.push(`${cannibal.name}收入${swan(cannibal.rev)}但利润${swan(cannibal.profit)}，增收不增利，吃掉了其他品类的利润改善`);
    const pdn = lst(sec.cats, 'profit', -1, sec.thr);
    if (pdn.length && pd < 0) out.push(`利润主要被${pdn.slice(0, 2).map(x => x.name + swan(x.profit)).join('、')}拖累`);
    const gm = sec.fees.find(x => x.key === '=产品毛利额'), feeD = sec.fees.filter(x => x.key !== '=产品毛利额').reduce((s, x) => s + (x.d || 0), 0);
    if (gm && gm.d !== null && gm.d <= -0.01) {
      if (feeD < -0.005) out.push(Math.abs(gm.d) > Math.abs(feeD) ? `毛利率下滑${(Math.abs(gm.d) * 100).toFixed(1)}个百分点，超过费用率压降的${(Math.abs(feeD) * 100).toFixed(1)}个百分点，费用节约补不上毛利缺口` : `毛利率下滑${(Math.abs(gm.d) * 100).toFixed(1)}个百分点，但费用率压降${(Math.abs(feeD) * 100).toFixed(1)}个百分点，抵消了毛利缺口`);
      else if (feeD > 0.005) out.push(`毛利率下滑${(Math.abs(gm.d) * 100).toFixed(1)}个百分点，费用率又上升${(feeD * 100).toFixed(1)}个百分点，利润受双重挤压`);
      else out.push(`毛利率下滑${(Math.abs(gm.d) * 100).toFixed(1)}个百分点`);
    }
    const big = sec.fees.filter(x => x.key !== '=产品毛利额').filter(x => x.d !== null).sort((p, q) => Math.abs(q.d) - Math.abs(p.d))[0];
    if (big && Math.abs(big.d) >= 0.02) {
      const top = (sec.drill[big.key] || []).filter(s => Math.sign(s.d) === Math.sign(big.d))[0];
      out.push(`${big.label}费率${big.d > 0 ? '上升' : '下降'}${(Math.abs(big.d) * 100).toFixed(1)}个百分点` + (top ? `，主要来自${top.name}（${spc(top.d)}）` : ''));
    }
    const fixedUp = sec.fees.filter(x => ['-PP', '-MP'].includes(x.key) && x.d > 0.005);
    if (g !== null && g < -0.3 && fixedUp.length) out.push(`收入萎缩推高了${fixedUp.map(x => x.label).join('、')}等相对固定费用的费率`);
    return out;
  }

  /* ---------- 思维导图结构（导出 XMind / 复制文字共用） ---------- */
  function splitConclusions(s) {
    return { product: s.conclusions.filter(c => !/费率|毛利/.test(c)).join('；'), fee: s.conclusions.filter(c => /费率|毛利/.test(c)).join('；') || '费用结构变化不大' };
  }
  function tree(results, title, edits) {
    return { title, children: results.map(res => ({ title: res.cust, children: ['当月', '累计'].map(k => {
      const s = res[k], e = Object.assign(splitConclusions(s), ((edits || {})[res.cust] || {})[k] || {});
      return { title: k + '\n' + headline(s), bold: k, children: [
        { title: '产品维度\n' + productText(s, res.opts), children: [{ title: '结论：' + (e.product || '—') }] },
        { title: '费用维度\n' + feeText(s), children: [{ title: '结论：' + (e.fee || '—') }] }
      ] };
    }) })) };
  }
  function treeText(node, d = 0) { return '  '.repeat(d) + '- ' + node.title.replace(/\n/g, '\n' + '  '.repeat(d) + '  ') + '\n' + (node.children || []).map(c => treeText(c, d + 1)).join(''); }

  /* ---------- 由实际数据生成测算方案 ---------- */
  function toScheme(dc, f, price) {
    const X = sums(dc, { dim: f.dim, cust: f.cust, months: new Set(f.months), sku: f.sku });
    const g = n => X.s[dc.col[n]] || 0, G = g('收入');
    if (!(G > 0)) throw new Error('所选范围的收入为 0，无法生成方案');
    const need = ['-产品成本', '-AP', '-TP', '-CP', '-PP', '-MP', '-FP'].filter(n => !(n in dc.col));
    if (need.length) throw new Error('数据缺少这些费用列，无法生成完整方案：' + need.join('、'));
    const r = n => g(n) / G, other = (G - g('-产品成本') - g('-AP') - g('-TP') - g('-CP') - g('-PP') - g('-MP') - g('-FP') - g('=产品考核净利润')) / G;
    const vol = Math.round(G * 10000 / price);
    const p2 = v => Math.round(v * 10000) / 10000;
    const items = [['AP 广告费用', r('-AP')], ['CP 促销费用', r('-CP')], ['MP 管理费用', r('-MP')], ['FP 财务费用', r('-FP')], ['其他（研发/公摊/所得税）', other]]
      .map(([name, v], i) => ({ id: 'act' + i, name, nature: 'cost', method: 'pct', pctBase: 'eff', src: 'num', vals: { base: p2(v * 100), promo: p2(v * 100) } }));
    const base = { list: price, discount: 0, volume: vol, purchase: p2(price * r('-产品成本')), purchaseUnit: 'bag', bagsPerBox: 12, laborMode: 'none', labor: 0, cold: p2(price * r('-PP')), loss: 0, rate: p2(r('-TP') * 100), rateBase: 'eff', fee: 0 };
    return { expectProfit: g('=产品考核净利润') * 10000, revenue: G * 10000, base, items };
  }

  return { SOURCES, detect, splitConclusions, OPTIONAL, NEED, dataset, values, analyze, tree, treeText, headline, productText, feeText, toScheme, prevYear, wan, swan, pc, spc, DIMS };
})();
if (typeof module !== 'undefined') module.exports = Review;
