/* ===== 计算引擎：利润 / 人工 / 公式 ===== */
const Engine = (() => {
  const EPS = 1e-9;
  const ceil = x => Math.ceil(x - EPS);
  const num = v => (v === '' || v === null || v === undefined || isNaN(+v)) ? NaN : +v;

  /* ---------- 公式解析（四则运算、比较、IF/MIN/MAX/ROUND） ---------- */
  const FUNCS = { IF: 3, MIN: -1, MAX: -1, ROUND: 2 };
  function tokenize(src) {
    const t = []; let i = 0; const s = String(src).replace(/^\s*=/, '');
    while (i < s.length) {
      const c = s[i];
      if (/\s/.test(c)) { i++; continue; }
      if (/[0-9.]/.test(c)) { let j = i; while (j < s.length && /[0-9.]/.test(s[j])) j++; t.push({ k: 'n', v: parseFloat(s.slice(i, j)) }); i = j; continue; }
      if (/[A-Za-z_一-鿿]/.test(c)) { let j = i; while (j < s.length && /[A-Za-z0-9_一-鿿]/.test(s[j])) j++; t.push({ k: 'id', v: s.slice(i, j) }); i = j; continue; }
      const two = s.slice(i, i + 2);
      if (['>=', '<=', '<>'].includes(two)) { t.push({ k: 'op', v: two }); i += 2; continue; }
      if ('+-*/(),><=×÷（），'.includes(c)) {
        const map = { '×': '*', '÷': '/', '（': '(', '）': ')', '，': ',' };
        t.push({ k: 'op', v: map[c] || c }); i++; continue;
      }
      throw new Error('公式里有无法识别的字符：' + c);
    }
    return t;
  }
  function parse(src) {
    const t = tokenize(src); let p = 0;
    const peek = () => t[p], eat = v => { if (!t[p] || t[p].v !== v) throw new Error('公式缺少 “' + v + '”'); p++; };
    function cmp() { let a = add(); while (peek() && ['>', '<', '>=', '<=', '=', '<>'].includes(peek().v)) { const o = t[p++].v; a = { k: 'bin', o, a, b: add() }; } return a; }
    function add() { let a = mul(); while (peek() && ['+', '-'].includes(peek().v)) { const o = t[p++].v; a = { k: 'bin', o, a, b: mul() }; } return a; }
    function mul() { let a = un(); while (peek() && ['*', '/'].includes(peek().v)) { const o = t[p++].v; a = { k: 'bin', o, a, b: un() }; } return a; }
    function un() { if (peek() && peek().v === '-') { p++; return { k: 'neg', a: un() }; } if (peek() && peek().v === '+') { p++; return un(); } return atom(); }
    function atom() {
      const x = t[p++]; if (!x) throw new Error('公式不完整');
      if (x.k === 'n') return { k: 'n', v: x.v };
      if (x.v === '(') { const e = cmp(); eat(')'); return e; }
      if (x.k === 'id') {
        const F = x.v.toUpperCase();
        if (peek() && peek().v === '(' && F in FUNCS) {
          p++; const args = [];
          if (peek() && peek().v !== ')') { args.push(cmp()); while (peek() && peek().v === ',') { p++; args.push(cmp()); } }
          eat(')');
          if (FUNCS[F] > 0 && args.length !== FUNCS[F]) throw new Error(F + ' 需要 ' + FUNCS[F] + ' 个参数');
          if (!args.length) throw new Error(F + ' 缺少参数');
          return { k: 'fn', f: F, args };
        }
        return { k: 'var', v: x.v };
      }
      throw new Error('公式在 “' + x.v + '” 附近有误');
    }
    const e = cmp(); if (p < t.length) throw new Error('公式在 “' + t[p].v + '” 附近有误'); return e;
  }
  function vars(ast, out = new Set()) {
    if (!ast) return out;
    if (ast.k === 'var') out.add(ast.v);
    ['a', 'b'].forEach(k => ast[k] && vars(ast[k], out));
    (ast.args || []).forEach(a => vars(a, out));
    return out;
  }
  function evalAst(ast, env) {
    switch (ast.k) {
      case 'n': return ast.v;
      case 'var': if (!(ast.v in env)) throw new Error('找不到变量 “' + ast.v + '”'); return env[ast.v];
      case 'neg': return -evalAst(ast.a, env);
      case 'bin': {
        const a = evalAst(ast.a, env), b = evalAst(ast.b, env);
        switch (ast.o) {
          case '+': return a + b; case '-': return a - b; case '*': return a * b;
          case '/': if (b === 0) throw new Error('公式里出现除以 0'); return a / b;
          case '>': return a > b ? 1 : 0; case '<': return a < b ? 1 : 0; case '>=': return a >= b ? 1 : 0;
          case '<=': return a <= b ? 1 : 0; case '=': return Math.abs(a - b) < EPS ? 1 : 0; case '<>': return Math.abs(a - b) >= EPS ? 1 : 0;
        }
      }
      case 'fn': {
        const v = ast.args.map(a => evalAst(a, env));
        if (ast.f === 'IF') return v[0] ? v[1] : v[2];
        if (ast.f === 'MIN') return Math.min(...v);
        if (ast.f === 'MAX') return Math.max(...v);
        if (ast.f === 'ROUND') { const m = Math.pow(10, v[1]); return Math.round(v[0] * m) / m; }
      }
    }
    throw new Error('公式无法计算');
  }
  function toExcel(ast, ref) {
    const r = a => toExcel(a, ref);
    switch (ast.k) {
      case 'n': return String(ast.v);
      case 'var': if (!(ast.v in ref)) throw new Error('导出时找不到变量 ' + ast.v); return ref[ast.v];
      case 'neg': return '(-' + r(ast.a) + ')';
      case 'bin': return '(' + r(ast.a) + ast.o + r(ast.b) + ')';
      case 'fn': return ast.f + '(' + ast.args.map(r).join(',') + ')';
    }
  }
  function evalFormula(src, env) { return evalAst(parse(src), env); }

  /* ---------- 人工计算器 ---------- */
  function sortedTiers(L) {
    return (L.tiers || []).filter(t => num(t.threshold) > 0 && num(t.amount) > 0)
      .map(t => ({ threshold: +t.threshold, amount: +t.amount })).sort((a, b) => a.threshold - b.threshold);
  }
  function tierBonus(L, V) { let b = 0; sortedTiers(L).forEach(t => { if (V >= t.threshold) b = t.amount; }); return b; }
  function laborTeam(L, allocPct) {
    const roles = (L.roles || []).map(r => {
      const base = (+r.count || 0) * (+r.salary || 0), social = base * (+r.social || 0) / 100, perf = (+r.count || 0) * (+r.perfBase || 0) * (+r.perfCoef || 0);
      return { name: r.name, count: +r.count || 0, base, social, perf, total: base + social + perf };
    });
    const sum = k => roles.reduce((s, r) => s + r[k], 0);
    const team = sum('total'), alloc = (allocPct === undefined || allocPct === '' || allocPct === null ? (+L.alloc || 0) : +allocPct) / 100;
    return { roles, team, alloc, base: sum('base') * alloc, social: sum('social') * alloc, perf: sum('perf') * alloc, fixed: team * alloc, people: sum('count') };
  }
  function laborFor(L, V, eff, allocPct) {
    const t = laborTeam(L, allocPct);
    const piece = +L.piece || 0, revPct = (+L.revPct || 0) / 100;
    const perBag = piece + revPct * eff, bonus = tierBonus(L, V);
    const total = t.fixed + perBag * V + bonus;
    return { ...t, piece, revPct, perBag, bonus, commission: perBag * V, total, perBagAll: V > 0 ? total / V : NaN };
  }

  /* ---------- 单个方案 ---------- */
  const CORE = ['list', 'discount', 'volume', 'purchase', 'labor', 'cold', 'loss', 'rate', 'fee'];
  const CORE_NAME = { list: '标价', discount: '商家折扣', volume: '销量', purchase: '采购', labor: '人工', cold: '冷链', loss: '损耗', rate: '扣点率', fee: '活动费' };

  function compute(sku, key, labor, overrides) {
    const s = Object.assign({}, sku[key], overrides || {});
    const errors = [];
    const missing = [];
    const need = ['list', 'volume', 'purchase', 'cold', 'loss', 'rate'];
    if (!s.laborMode || s.laborMode === 'manual') need.push('labor');
    need.forEach(k => { if (isNaN(num(s[k]))) missing.push(CORE_NAME[k]); });
    if (missing.length) return { ok: false, missing, errors: ['请补齐：' + missing.join('、') + '（零成本请明确填 0）'] };
    const V = +s.volume, list = +s.list, disc = +(s.discount || 0), eff = list - disc;
    const purchase = s.purchaseUnit === 'box' ? (+s.purchase) / (+s.bagsPerBox || NaN) : +s.purchase;
    if (s.purchaseUnit === 'box' && !(+s.bagsPerBox > 0)) return { ok: false, missing: ['每箱袋数'], errors: ['采购按箱录入时，请填写每箱袋数'] };
    const rate = (+s.rate) / 100, base = s.rateBase === 'list' ? list : eff, commission = rate * base;
    let lab = null, laborBag = s.laborMode === 'none' ? 0 : (+s.labor || 0), laborFixed = 0, tier = 0;
    if (s.laborMode === 'calc' && labor) { lab = laborFor(labor, V, eff, sku.laborAlloc); laborBag = lab.perBag; laborFixed = lab.fixed; tier = lab.bonus; }
    const env = { 标价: list, 商家折扣: disc, 有效售价: eff, 销量: V, 采购: purchase, 人工: laborBag, 冷链: +s.cold, 损耗: +s.loss, 扣点率: rate, 活动费: +(s.fee || 0) };
    const items = [];
    let costBag = 0, subBag = 0, costEvt = 0, subEvt = 0;
    (sku.items || []).forEach(it => {
      const raw = (it.vals || {})[key];
      let v;
      try {
        if (it.src === 'formula') { if (!raw) throw new Error('公式为空'); v = evalFormula(raw, env); }
        else { v = num(raw); if (isNaN(v)) throw new Error('请填写数值（没有就填 0）'); }
      } catch (e) { errors.push((it.name || '自定义项') + '：' + e.message); v = 0; }
      let perBag = 0, fixed = 0, envVal = v;
      if (it.method === 'bag') perBag = v;
      else if (it.method === 'order') perBag = (+it.bagsPerOrder > 0) ? v / +it.bagsPerOrder : (errors.push((it.name || '') + '：请填写每单袋数'), 0);
      else if (it.method === 'pct') { envVal = v / 100; perBag = envVal * (it.pctBase === 'list' ? list : eff); }
      else if (it.method === 'event') fixed = v;
      if (it.name) env[it.name] = envVal;
      if (it.nature === 'cost') { costBag += perBag; costEvt += fixed; }
      else if (it.nature === 'subsidy') { subBag += perBag; subEvt += fixed; }
      items.push({ id: it.id, name: it.name, nature: it.nature, method: it.method, value: v, perBag, fixed });
    });
    const fee = +(s.fee || 0);
    const unit = eff - purchase - laborBag - (+s.cold) - (+s.loss) - commission - costBag + subBag;
    const F0 = fee + laborFixed + costEvt;           // 不含阶梯激励的固定费用
    const fixed = F0 + tier;
    const rev = eff * V, contribution = unit * V, profit = contribution - fixed + subEvt;
    const margin = rev ? profit / rev : null, unitRate = eff ? unit / eff : null;
    const res = {
      ok: true, errors, V, list, disc, eff, rev, purchase, laborBag, cold: +s.cold, loss: +s.loss, rate, rateBase: s.rateBase, commission,
      costBag, subBag, costEvt, subEvt, fee, laborFixed, tier, F0, fixed, unit, contribution, profit, margin, unitRate, items, lab, env
    };
    res.laborTotal = laborBag * V + laborFixed + tier;
    res.minVolume = target => minVolume(res, labor && s.laborMode === 'calc' ? labor : null, target);
    res.be = res.minVolume(0);
    // 衍生指标
    const env2 = Object.assign({}, env, { 收入: rev, 单位贡献: unit, 贡献利润: contribution, 固定费用: fixed, 场景利润: profit, 利润率: margin === null ? 0 : margin, 单位贡献率: unitRate === null ? 0 : unitRate });
    res.metrics = (sku.metrics || []).map(m => {
      try { const v = evalFormula(m.expr, env2); if (m.name) env2[m.name] = v; return { name: m.name, value: v, unit: m.unit }; }
      catch (e) { return { name: m.name, error: e.message, unit: m.unit }; }
    });
    return res;
  }

  // 利润 ≥ target 的最小销量（考虑阶梯激励）。单位贡献 ≤ 0 时返回 null
  function minVolume(r, labor, target) {
    if (!(r.unit > EPS)) return null;
    const tiers = labor ? sortedTiers(labor) : [];
    const segs = [{ start: 0, bonus: 0 }].concat(tiers.map(t => ({ start: t.threshold, bonus: t.amount })));
    for (let k = 0; k < segs.length; k++) {
      const end = k + 1 < segs.length ? segs[k + 1].start : Infinity;
      const need = (target + r.F0 + segs[k].bonus - r.subEvt) / r.unit;
      const cand = Math.max(segs[k].start, ceil(need));
      if (cand < end) return cand;
    }
    return null;
  }

  /* ---------- 敏感性 & 售价曲线 ---------- */
  function sensitivity(sku, key, labor) {
    const s = sku[key], baseR = compute(sku, key, labor); if (!baseR.ok) return [];
    const keys = ['list', 'volume', 'purchase', 'cold', 'loss', 'rate', 'fee'];
    if (!s.laborMode || s.laborMode === 'manual') keys.splice(3, 0, 'labor');
    return keys.map(k => {
      const v = +s[k] || 0; if (!v) return null;
      const up = compute(sku, key, labor, { [k]: v * 1.1 }).profit - baseR.profit;
      const dn = compute(sku, key, labor, { [k]: v * 0.9 }).profit - baseR.profit;
      return { key: k, name: CORE_NAME[k], up, dn, span: Math.abs(up - dn) };
    }).filter(Boolean).sort((a, b) => b.span - a.span);
  }
  function priceCurve(sku, key, labor) {
    const s = sku[key], L = +s.list; if (!(L > 0)) return null;
    const pts = [];
    for (let i = 0; i <= 40; i++) { const p = L * (0.6 + 0.7 * i / 40); const r = compute(sku, key, labor, { list: p }); if (r.ok) pts.push({ p, profit: r.profit }); }
    let bePrice = null;
    for (let i = 1; i < pts.length; i++) if (pts[i - 1].profit < 0 && pts[i].profit >= 0) { const a = pts[i - 1], b = pts[i]; bePrice = a.p + (0 - a.profit) * (b.p - a.p) / (b.profit - a.profit); break; }
    return { pts, bePrice, cur: L, curProfit: compute(sku, key, labor).profit };
  }

  /* ---------- 自动分析 ---------- */
  const fmtY = v => '¥' + Math.round(v).toLocaleString('zh-CN');
  const fmtP = v => (v * 100).toFixed(1) + '%';
  function analyze(sku, labor, st) {
    const out = [];
    const b = compute(sku, 'base', labor), p = compute(sku, 'promo', labor);
    if (!b.ok || !p.ok) return [{ level: 'warn', text: '有必填参数未填，补齐后才能给出分析。' }];
    [b, p].forEach((r, i) => r.errors.forEach(e => out.push({ level: 'warn', text: (i ? '促销方案' : '基准方案') + ' · ' + e })));
    const eq = p.minVolume(b.profit);
    if (eq === null) out.push({ level: 'bad', text: '促销方案每袋不赚钱，卖多少都追不平基准利润 ' + fmtY(b.profit) + '。' });
    else if (p.V >= eq) out.push({ level: 'good', text: '促销预计卖 ' + p.V.toLocaleString() + ' 袋，已超过追平基准利润所需的 ' + eq.toLocaleString() + ' 袋，多赚 ' + fmtY(p.profit - b.profit) + '。' });
    else out.push({ level: 'info', text: '促销预计卖 ' + p.V.toLocaleString() + ' 袋，距离追平基准利润还差 ' + (eq - p.V).toLocaleString() + ' 袋（需 ' + eq.toLocaleString() + ' 袋）。' });
    if (eq !== null && b.V > 0) { const lift = eq / b.V - 1; if (st.liftMax !== '' && lift * 100 > +st.liftMax) out.push({ level: 'warn', text: '要比基准多卖 ' + fmtP(lift) + ' 才能持平，高于设定的 ' + st.liftMax + '%，请核实能否做到。' }); }
    [['基准', b], ['促销', p]].forEach(([n, r]) => {
      if (r.unit <= 0) out.push({ level: 'bad', text: n + '方案每卖一袋亏 ' + '¥' + (-r.unit).toFixed(2) + '，加量只会扩大亏损。' });
      if (st.targetMargin !== '' && r.margin !== null && r.margin * 100 < +st.targetMargin) out.push({ level: 'warn', text: n + '方案利润率 ' + fmtP(r.margin) + '，低于目标 ' + st.targetMargin + '%。' });
    });
    if (st.coldMax !== '' && p.rev > 0) { const c = p.cold * p.V / p.rev; if (c * 100 > +st.coldMax) out.push({ level: 'warn', text: '促销方案冷链占收入 ' + fmtP(c) + '，高于设定的 ' + st.coldMax + '%。' }); }
    if (p.lab) {
      const share = p.lab.total / p.rev;
      if (st.laborMax !== '' && p.rev > 0 && share * 100 > +st.laborMax) out.push({ level: 'warn', text: '人工占收入 ' + fmtP(share) + '，高于设定的 ' + st.laborMax + '%。' });
      sortedTiers(labor).forEach(t => {
        if (p.V >= t.threshold && p.unit > 0 && p.V < t.threshold + t.amount / p.unit) {
          const prev = compute(sku, 'promo', labor, { volume: t.threshold - 1 }).profit;
          if (prev > p.profit) out.push({ level: 'warn', text: '阶梯激励临界点：卖 ' + (t.threshold - 1).toLocaleString() + ' 袋比现在的 ' + p.V.toLocaleString() + ' 袋多赚 ' + fmtY(prev - p.profit) + '，请核对激励设置。' });
        }
      });
    }
    const fl = sku.assumed || {}, costs = { purchase: p.purchase, labor: p.laborBag, cold: p.cold, loss: p.loss };
    const tot = Object.values(costs).reduce((a, c) => a + c, 0), ass = Object.keys(costs).reduce((a, k) => a + (fl[k] ? costs[k] : 0), 0);
    if (tot > 0 && st.assumedMax !== '' && ass / tot * 100 > +st.assumedMax) out.push({ level: 'info', text: '标为“估计”的成本占每袋成本 ' + fmtP(ass / tot) + '，结果主要基于估计值，仅供参考。' });
    return out;
  }

  return { parse, vars, evalFormula, toExcel, compute, laborFor, laborTeam, tierBonus, sortedTiers, sensitivity, priceCurve, analyze, CORE_NAME, fmtY, fmtP };
})();
if (typeof module !== 'undefined') module.exports = Engine;
