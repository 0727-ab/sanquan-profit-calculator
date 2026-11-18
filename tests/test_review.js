// 复盘引擎测试：用虚构数据（仓库公开，不放真实经营数据）
const V = require('../src/review.js');
let fails = 0; const eq = (n, g, w) => { const ok = g === w; fails += !ok; console.log((ok ? 'PASS ' : 'FAIL ') + n + ' → ' + g + (ok ? '' : '  (want ' + w + ')')); };
const H = ['日历年/月', '收入', '-产品成本', '=产品毛利额', '-AP', '-TP', '5-供价差异', '-CP', '1-折价', '1.1-价格现扣', '-PP', '-MP', '-FP', '-产品组/研发费用合计', '-AE', '所得税', '=产品考核净利润', '四级渠道', '一级品类', '四级规格'];
// 收入, 成本, 毛利, AP, TP, TP5, CP, 折价, 现扣, PP, MP, FP, 研发, AE, 税 → 净利自动算
const row = (m, c, cat, sku, rev, cost, tp, cp, pp) => { const np = rev - cost - tp - cp - pp; return [m, rev, cost, rev - cost, 0, tp, tp / 2, cp, cp, cp, pp, 0, 0, 0, 0, 0, np, c, cat, sku]; };
const cur = [['整体'], H, row('2026/10', '客户A', '饺类', '水饺1kg', 100, 60, 10, 5, 5), row('2026/10', '客户A', '汤圆', '汤圆1kg', 50, 30, 5, 5, 5), row('2026/09', '客户A', '饺类', '水饺1kg', 80, 48, 8, 4, 4)];
const pri = [['整体'], H, row('2025/10', '客户A', '饺类', '水饺1kg', 80, 40, 8, 8, 4), row('2025/10', '客户A', '汤圆', '汤圆1kg', 70, 42, 7, 7, 7), row('2025/09', '客户A', '饺类', '水饺1kg', 90, 54, 9, 9, 9)];
const dc = V.dataset(cur, '本期'), dp = V.dataset(pri, '同期');
eq('勾稽核对', dc.check.balance.ok, true);
const r = V.analyze(dc, dp, { dim: '四级渠道', cust: '客户A', month: '2026/10', cumStart: '2026/09' });
eq('当月收入', r.当月.rev.a + '/' + r.当月.rev.b, '150/150');
eq('当月利润', r.当月.profit.a + '/' + r.当月.profit.b, '25/27');
eq('当月毛利率', V.pc(r.当月.fees[0].a) + '/' + V.pc(r.当月.fees[0].b), '40.0%/45.3%');
eq('当月 TP5', V.pc(r.当月.tp5.a), '5.0%');
eq('累计收入', r.累计.rev.a + '/' + r.累计.rev.b, '230/240');
eq('品类：饺类收入+20', r.当月.cats.find(x => x.name === '饺类').rev, 20);
eq('结论：毛利率下滑', r.当月.conclusions.some(c => c.includes('毛利率下滑5.3个百分点')), true);
// 缺可选列：去掉品类列 → 产品分析隐藏，不报错
const cut = rows => rows.map((x, i) => i === 0 ? x : x.slice(0, 18).concat([null, x[19]]));
const H2 = H.slice(); H2[18] = '其他';
const cur2 = cut(cur); cur2[1] = H2; const pri2 = cut(pri); pri2[1] = H2;
const r2 = V.analyze(V.dataset(cur2, '本期'), V.dataset(pri2, '同期'), { dim: '四级渠道', cust: '客户A', month: '2026/10', cumStart: '2026/09' });
eq('缺品类列时隐藏产品分析', r2.当月.hasProduct, false);
// 缺必需列 → 拒绝导入
let err = ''; try { V.dataset([H.filter(h => h !== '收入')], '本期'); } catch (e) { err = e.message; } eq('缺必需列拒绝', err.includes('收入'), true);
// 生成方案：测算器利润 = 实际净利
const E = require('../src/engine.js');
const sc = V.toScheme(dc, { dim: '四级渠道', cust: '客户A', sku: '水饺1kg', months: ['2026/10'] }, 20);
eq('生成方案利润一致', Math.round(E.compute({ base: sc.base, promo: sc.base, items: sc.items, metrics: [] }, 'base').profit), Math.round(sc.expectProfit));
console.log(fails ? fails + ' FAILED' : 'ALL PASS');
