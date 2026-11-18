const E = require('../src/engine.js'), X = require('../src/xlsx.js'), D = require('../src/defaults.js'), fs = require('fs');
const st = D.defaultState();
// 加入：自定义项（公式人工、按单冷链、按比例平台费、活动补贴）、人工计算器、指标
st.skus[0].promo.laborMode = 'calc';
st.skus[0].items = [
  { id: 'a', name: '包装费', nature: 'cost', method: 'bag', src: 'num', vals: { base: 0.3, promo: 0.3 } },
  { id: 'b', name: '配送费', nature: 'cost', method: 'order', src: 'num', bagsPerOrder: 6, vals: { base: 6, promo: 6 } },
  { id: 'c', name: '平台服务费', nature: 'cost', method: 'pct', pctBase: 'eff', src: 'num', vals: { base: 2, promo: 2 } },
  { id: 'd', name: '厂家补贴', nature: 'subsidy', method: 'event', src: 'num', vals: { base: 0, promo: 600 } },
  { id: 'e', name: '加班费', nature: 'cost', method: 'bag', src: 'formula', vals: { base: '=IF(销量>1100, 30/60*0.2, 0)', promo: '=IF(销量>1100, 30/60*0.2, 0)' } },
  { id: 'f', name: '竞品价', nature: 'ref', method: 'bag', src: 'num', vals: { base: 17, promo: 17 } }
];
st.skus[0].metrics = [{ name: '每袋利润', expr: '场景利润 / 销量' }, { name: '比竞品贵', expr: '有效售价 - 竞品价' }];
st.skus[1].base.purchaseUnit = 'box'; st.skus[1].base.purchase = 78; st.skus[1].base.bagsPerBox = 12;
st.commonFee = 1000; st.skus[2].promo.laborMode = 'none'; st.skus[1].promo.laborMode = 'calc'; st.skus[1].laborAlloc = 30;
const an = st.skus.map(s => E.analyze(s, st.labor, st.settings));
X.setNoCache(process.argv[2] === 'nocache');
fs.writeFileSync(process.argv[3] || 'out.xlsx', Buffer.from(X.build(st, E, an)));
// 期望值
const exp = {};
st.skus.forEach(s => { exp[s.name] = { base: E.compute(s, 'base', st.labor), promo: E.compute(s, 'promo', st.labor) }; exp[s.name].eq = exp[s.name].promo.minVolume(exp[s.name].base.profit); });
const j = {}; Object.keys(exp).forEach(n => { const e = exp[n]; j[n] = ['base', 'promo'].map(k => ({ rev: e[k].rev, unit: e[k].unit, profit: e[k].profit, be: e[k].be, metrics: e[k].metrics.map(m => m.value) })); j[n].push({ eq: e.eq }); });
fs.writeFileSync('expected.json', JSON.stringify(j));
