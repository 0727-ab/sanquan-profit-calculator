/* ===== 演示数据（全部为假设） ===== */
const LABOR_PRESETS = {
  A: { label: 'A 固定薪资型', roles: [{ name: '仓管', count: 1, salary: 4500, social: 30, perfBase: 0, perfCoef: 0 }, { name: '行政', count: 1, salary: 4000, social: 30, perfBase: 0, perfCoef: 0 }], piece: 0, revPct: 0, tiers: [] },
  B: { label: 'B 底薪 + 计件型', roles: [{ name: '拣货员', count: 2, salary: 3500, social: 30, perfBase: 0, perfCoef: 0 }], piece: 0.1, revPct: 0, tiers: [] },
  C: { label: 'C 绩效激励型', roles: [{ name: '业务员', count: 2, salary: 4000, social: 30, perfBase: 800, perfCoef: 1 }], piece: 0.1, revPct: 0, tiers: [{ threshold: 1200, amount: 600 }] }
};
function scheme(o) { return Object.assign({ list: 18, discount: 0, volume: 1000, purchase: 10, purchaseUnit: 'bag', bagsPerBox: 12, laborMode: 'manual', labor: 0.4, cold: 1, loss: 0.2, rate: 5, rateBase: 'eff', fee: 2000 }, o); }
function defaultState() {
  return {
    version: 1,
    meta: { company: '三全郑州分公司', channel: '演示渠道', month: '2026-12' },
    skus: [
      { id: 's1', name: '速冻水饺 500g', base: scheme({}), promo: scheme({ discount: 2, volume: 1200 }), items: [], metrics: [], laborAlloc: 10, assumed: { purchase: true, labor: true, cold: true, loss: true } },
      { id: 's2', name: '汤圆 400g', base: scheme({ list: 12, purchase: 6.5, labor: 0.3, cold: 0.8, loss: 0.15, fee: 1000, volume: 1500 }), promo: scheme({ list: 12, discount: 1.5, purchase: 6.5, labor: 0.3, cold: 0.8, loss: 0.15, fee: 1000, volume: 1900 }), items: [], metrics: [], laborAlloc: 10, assumed: { purchase: true, labor: true, cold: true, loss: true } },
      { id: 's3', name: '馄饨 300g', base: scheme({ list: 15, purchase: 8.8, labor: 0.35, cold: 0.9, loss: 0.25, fee: 1500, volume: 800 }), promo: scheme({ list: 15, discount: 2, purchase: 8.8, labor: 0.35, cold: 0.9, loss: 0.25, fee: 1500, volume: 1050 }), items: [], metrics: [], laborAlloc: 10, assumed: { purchase: true, labor: true, cold: true, loss: true } }
    ],
    active: 0,
    view: 'promo',
    labor: Object.assign({ preset: 'C', alloc: 10 }, JSON.parse(JSON.stringify(LABOR_PRESETS.C))),
    settings: { targetMargin: 15, coldMax: 5, laborMax: 8, liftMax: 30, assumedMax: 50 },
    commonFee: 0
  };
}
if (typeof module !== 'undefined') module.exports = { defaultState, LABOR_PRESETS, scheme };
