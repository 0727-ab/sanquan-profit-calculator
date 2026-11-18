const E = require('../src/engine.js');
let fails = 0;
const eq = (name, got, want) => { const ok = (typeof want === 'number' && typeof got === 'number') ? Math.abs(got - want) < 0.01 : got === want; if (!ok) fails++; console.log((ok ? 'PASS ' : 'FAIL ') + name + ' → ' + got + (ok ? '' : ' (want ' + want + ')')); };
const S = o => Object.assign({ list: 18, discount: 0, volume: 1000, purchase: 10, purchaseUnit: 'bag', labor: 0.4, laborMode: 'manual', cold: 1, loss: 0.2, rate: 5, rateBase: 'eff', fee: 2000 }, o);
const sku = (b, p, extra) => Object.assign({ base: S(b), promo: S(p), items: [], metrics: [] }, extra || {});
const st = { targetMargin: 15, coldMax: 5, laborMax: 8, liftMax: 30, assumedMax: 50 };

// 1 基准与降价
let k = sku({}, { list: 16, volume: 1200 });
let b = E.compute(k, 'base'), p = E.compute(k, 'promo');
eq('1 基准利润', b.profit, 3500); eq('1 降价利润', p.profit, 2320); eq('1 基准BE', b.be, 364); eq('1 降价BE', p.be, 556);
eq('1 等利润销量', p.minVolume(b.profit), 1528); eq('1 利润差', p.profit - b.profit, -1180);
// 2 按箱转袋 + 缺失
k = sku({ purchase: 120, purchaseUnit: 'box', bagsPerBox: 12, cold: '' }, {});
eq('2 缺失冷链不出结果', E.compute(k, 'base').ok, false);
k.base.cold = 0; b = E.compute(k, 'base'); eq('2 采购/袋', b.purchase, 10); eq('2 单位贡献', b.unit, 6.5);
// 3 零销量
k = sku({ volume: 0 }, {}); b = E.compute(k, 'base');
eq('3 收入', b.rev, 0); eq('3 利润', b.profit, -2000); eq('3 利润率', b.margin, null); eq('3 BE', b.be, 364);
// 4 亏损 SKU + 汇总
const A = E.compute(sku({}, {}), 'base'), B = E.compute(sku({ list: 12, fee: 0 }, {}), 'base');
eq('4 B 单位贡献', B.unit, -0.2); eq('4 B 利润', B.profit, -200); eq('4 B BE', B.be, null);
const totRev = A.rev + B.rev, totP = A.profit + B.profit - 1000;
eq('4 汇总收入', totRev, 30000); eq('4 汇总利润', totP, 2300); eq('4 汇总利润率', +(totP / totRev * 100).toFixed(2), 7.67);
// 5 新增变量
k = sku({}, {}, { items: [{ id: 'a', name: '包装费', nature: 'cost', method: 'bag', src: 'num', vals: { base: 0.3, promo: 0.3 } }] });
b = E.compute(k, 'base'); eq('5 单位贡献', b.unit, 5.2); eq('5 利润', b.profit, 3200); eq('5 BE', b.be, 385);
// 6 自定义公式 + 分析
k = sku({}, { list: 16, volume: 1200 }, { metrics: [{ name: '单位贡献率2', expr: '单位贡献 ÷ 有效售价' }] });
b = E.compute(k, 'base'); eq('6 单位贡献率', +(b.metrics[0].value * 100).toFixed(1), 30.6);
const an = E.analyze(k, null, st).map(x => x.text);
eq('6 触发利润率偏低', an.some(t => t.includes('利润率 12.1%')), true);
eq('6 触发降价门槛', an.some(t => t.includes('52.8%')), true);
eq('6 追平差距 328', an.some(t => t.includes('还差 328 袋')), true);
let err = null; try { E.evalFormula('不存在 + 1', { a: 1 }); } catch (e) { err = e.message; } eq('6 错误公式拒绝', !!err, true);
// 人工公式算费用
k = sku({ labor: '' }, {}, { items: [] }); k.base.labor = 0;
k.items = [{ id: 'l', name: '拣货人工', nature: 'cost', method: 'bag', src: 'formula', vals: { base: '=30/60*0.8', promo: '=30/60*0.8' } }];
b = E.compute(k, 'base'); eq('公式算人工 0.40/袋 → 单位贡献', b.unit, 5.5);
// 7 折前/折后等价
const k7a = sku({}, { list: 18, discount: 2, volume: 1200 }), k7b = sku({}, { list: 16, volume: 1200 });
eq('7 折前录法收入', E.compute(k7a, 'promo').rev, 19200); eq('7 折前录法利润', E.compute(k7a, 'promo').profit, 2320); eq('7 折后录法利润', E.compute(k7b, 'promo').profit, 2320);
// 8 补贴
const k8a = sku({}, { list: 16, volume: 1200 }, { items: [{ id: 's', name: '每袋补贴', nature: 'subsidy', method: 'bag', src: 'num', vals: { base: 0, promo: 0.5 } }] });
const k8b = sku({}, { list: 16, volume: 1200 }, { items: [{ id: 's', name: '活动补贴', nature: 'subsidy', method: 'event', src: 'num', vals: { base: 0, promo: 600 } }] });
const p8a = E.compute(k8a, 'promo'), p8b = E.compute(k8b, 'promo'), b8 = E.compute(k8a, 'base');
eq('8 按袋利润', p8a.profit, 2920); eq('8 活动利润', p8b.profit, 2920);
eq('8 按袋追平', p8a.minVolume(b8.profit), 1342); eq('8 活动追平', p8b.minVolume(b8.profit), 1362);
// 9 人工 + 阶梯
const L = { roles: [{ name: '业务员', count: 2, salary: 4000, social: 30, perfBase: 800, perfCoef: 1 }], alloc: 10, piece: 0.1, revPct: 0, tiers: [{ threshold: 1200, amount: 600 }] };
const k9 = sku({}, { list: 16, volume: 1200, laborMode: 'calc' });
const p9 = E.compute(k9, 'promo', L);
eq('9 单位贡献', p9.unit, 3.9); eq('9 固定费用', p9.fixed, 3800); eq('9 利润', p9.profit, 880);
eq('9 1199袋利润', E.compute(k9, 'promo', L, { volume: 1199 }).profit, 1476.1); eq('9 BE', p9.be, 821);
eq('9 人工占收入', +(p9.lab.total / p9.rev * 100).toFixed(1), 10);
eq('9 临界提示', E.analyze(k9, L, st).some(x => x.text.includes('多赚 ¥596')), true);
// 不计人工 & 每个 SKU 单独的归属比例
const kn = sku({}, { list: 16, volume: 1200, laborMode: 'none' });
eq('不计人工 利润', E.compute(kn, 'promo').profit, 2800);
const k9b = sku({}, { list: 16, volume: 1200, laborMode: 'calc' }, { laborAlloc: 30 });
eq('归属30% 固定人工', E.compute(k9b, 'promo', L).laborFixed, 3600);
eq('含人工总额', E.compute(k9b, 'promo', L).laborTotal, 3600 + 120 + 600);
// 公式转 Excel
eq('Excel 转换', E.toExcel(E.parse('IF(销量>=1200, 600, 0)'), { 销量: 'B5' }), 'IF((B5>=1200),600,0)');
console.log(fails ? fails + ' FAILED' : 'ALL PASS');
