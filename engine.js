/* Etch Spend Forecast Builder: calculation engine.
   Pure functions, no DOM. The same formulas as the Excel model it exports.
   All money is in $K (thousands of US dollars). Months are 1..12 (fiscal). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.EtchEngine = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const LINES = ['headcount', 'materials', 'eval', 'capex'];
  const LINE_NAMES = { headcount: 'Headcount cost', materials: 'Materials', eval: 'Prototype & eval tools', opex: 'Total Opex', capex: 'Capex', total: 'Total spend' };
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const SLIPS = [0, 1, 2, 3, 4, 6];
  const CUTS = [0, 0.05, 0.1, 0.15, 0.2, 0.3];

  const z12 = () => Array(12).fill(0);
  const sum = (a) => a.reduce((s, v) => s + (+v || 0), 0);
  const add = (a, b) => a.map((v, i) => v + b[i]);

  function monthLabels(startMonth, startYear) {
    return Array.from({ length: 12 }, (_, i) => {
      const m = (startMonth - 1 + i) % 12;
      const y = startYear + Math.floor((startMonth - 1 + i) / 12);
      return `${MON[m]}-${String(y).slice(2)}`;
    });
  }

  /** Spread an annual total over 12 months: even, front (weights 12..1) or back (weights 1..12). */
  function phase(total, kind) {
    const t = +total || 0;
    if (kind === 'front') return Array.from({ length: 12 }, (_, k) => (t * (12 - k)) / 78);
    if (kind === 'back') return Array.from({ length: 12 }, (_, k) => (t * (k + 1)) / 78);
    return Array(12).fill(t / 12);
  }

  /** One program under one set of levers. Mirrors the Model tab, section A and B. */
  function runProgram(p, closed, lev) {
    const o = { hires: [], fte: [], headcount: [], materials: [], eval: [], opex: [], capex: [], total: [], dm: [] };
    let fte = +p.startFTE || 0;
    for (let m = 1; m <= 12; m++) {
      const dm = m <= closed ? m : m - lev.slip > closed ? m - lev.slip : 0;
      const eff = (dm ? +p.hires[dm - 1] || 0 : 0) * (m <= closed ? 1 : lev.hiring);
      fte += eff;
      const hc = (fte * (+p.loaded || 0)) / 12;
      const mat = dm ? (+p.materials[dm - 1] || 0) * (1 - lev.matCut) : 0;
      const ev = dm ? (+p.units[dm - 1] || 0) * (+p.unitCost || 0) * (1 - lev.evalCut) : 0;
      const cx = dm ? (+p.capex[dm - 1] || 0) * (1 - lev.capexCut) : 0;
      o.dm.push(dm); o.hires.push(eff); o.fte.push(fte); o.headcount.push(hc); o.materials.push(mat); o.eval.push(ev);
      o.opex.push(hc + mat + ev); o.capex.push(cx); o.total.push(hc + mat + ev + cx);
    }
    return o;
  }

  function depreciation(p, capexArr) {
    const life = Math.max(+p.life || 1, 0.0833), delay = +p.delay || 0;
    const dep = z12();
    for (let m = 1; m <= 12; m++) {
      let s = 0;
      for (let j = 1; j <= 12; j++) if (j <= m - delay) s += capexArr[j - 1];
      dep[m - 1] = s / (life * 12);
    }
    const acc = []; let a = 0, c = 0; const nbv = [];
    for (let m = 0; m < 12; m++) { a += dep[m]; c += capexArr[m]; acc.push(a); nbv.push(c - a); }
    return { dep, acc, nbv, capex: capexArr, life };
  }

  /** Full model for one scenario (lev) under cfg. Returns everything the UI and the export need. */
  function runModel(cfg, lev) {
    const closed = Math.max(0, Math.min(12, Math.round(+cfg.closed || 0)));
    const progs = cfg.programs.map((p) => {
      const sc = runProgram(p, closed, lev);
      const forecast = { headcount: [], materials: [], eval: [], capex: [] };
      for (let m = 1; m <= 12; m++) {
        for (const l of LINES) forecast[l].push(m <= closed ? +p.act[l][m - 1] || 0 : sc[l][m - 1]);
      }
      forecast.opex = add(add(forecast.headcount, forecast.materials), forecast.eval);
      forecast.total = add(forecast.opex, forecast.capex);
      const aop = { headcount: p.aop.headcount.map(Number), materials: p.aop.materials.map(Number), eval: p.aop.eval.map(Number), capex: p.aop.capex.map(Number) };
      aop.opex = add(add(aop.headcount, aop.materials), aop.eval); aop.total = add(aop.opex, aop.capex);
      const act = {};
      for (const l of LINES) act[l] = p.act[l].map((v, i) => (i < closed ? +v || 0 : 0));
      act.opex = add(add(act.headcount, act.materials), act.eval); act.total = add(act.opex, act.capex);
      return { name: p.name, scenario: sc, forecast, aop, act, dep: depreciation(p, forecast.capex) };
    });
    const port = { forecast: {}, aop: {}, act: {}, scenario: {} };
    const keys = ['headcount', 'materials', 'eval', 'opex', 'capex', 'total'];
    for (const k of keys) {
      port.forecast[k] = progs.reduce((a, p) => add(a, p.forecast[k]), z12());
      port.aop[k] = progs.reduce((a, p) => add(a, p.aop[k]), z12());
      port.act[k] = progs.reduce((a, p) => add(a, p.act[k]), z12());
      port.scenario[k] = progs.reduce((a, p) => add(a, p.scenario[k]), z12());
    }
    port.scenario.fte = progs.reduce((a, p) => add(a, p.scenario.fte), z12());
    const dep = { dep: progs.reduce((a, p) => add(a, p.dep.dep), z12()), acc: progs.reduce((a, p) => add(a, p.dep.acc), z12()), nbv: progs.reduce((a, p) => add(a, p.dep.nbv), z12()), capex: port.forecast.capex };

    // variance lines (portfolio and per program)
    const usd = +cfg.thresholds.usd, pct = +cfg.thresholds.pct;
    const mk = (aop, fc, act) => {
      const out = {};
      for (const k of keys) {
        const a = sum(aop[k]), f = sum(fc[k]), v = a - f;
        const pc = a === 0 ? 0 : v / a;
        const ytdAop = sum(aop[k].slice(0, closed)), ytdAct = sum(act[k].slice(0, closed));
        out[k] = { aop: a, fc: f, var: v, pct: pc, fu: Math.abs(v) < 0.5 ? 'On plan' : v > 0 ? 'F' : 'U', ytdAct, ytdAop, ytdVar: ytdAop - ytdAct, review: Math.abs(v) >= usd && Math.abs(pc) >= pct ? 'Review' : 'Within' };
      }
      return out;
    };
    const lines = mk(port.aop, port.forecast, port.act);
    const progLines = progs.map((p) => mk(p.aop, p.forecast, p.act));
    const bridge = [
      { label: 'AOP total spend', change: lines.total.aop, run: lines.total.aop },
      ...['headcount', 'materials', 'eval', 'capex'].map((k) => ({ label: LINE_NAMES[k], key: k, change: -lines[k].var })),
    ];
    let run = lines.total.aop;
    bridge.forEach((b, i) => { if (i) { run += b.change; b.run = run; b.fu = Math.abs(b.change) < 0.5 ? 'On plan' : b.change < 0 ? 'F' : 'U'; b.review = lines[b.key].review; } });
    bridge.push({ label: 'FY27 forecast total spend', change: lines.total.fc, run: lines.total.fc });
    const quarters = [0, 1, 2, 3].map((q) => ({ aop: sum(port.aop.total.slice(q * 3, q * 3 + 3)), fc: sum(port.forecast.total.slice(q * 3, q * 3 + 3)) }));

    const keysOut = {
      aop: lines.total.aop, fc: lines.total.fc, var: lines.total.var, pct: lines.total.pct, opex: lines.opex.fc, capex: lines.capex.fc,
      ytd: lines.total.ytdVar, fte: port.scenario.fte[11], dep: sum(dep.dep),
      runrate: progs.reduce((a, p, i) => a + sum(p.forecast.capex) / (Math.max(+cfg.programs[i].life || 1, 0.0833)), 0),
      nbv: dep.nbv[11],
    };
    return { closed, progs, port, dep, lines, progLines, bridge, quarters, keys: keysOut };
  }

  /** Plan view grid: hiring 100%, one cut on materials, eval and Capex, actuals ignored, slip shifts spend. */
  function sensitivity(cfg) {
    const c0 = Object.assign({}, cfg, { closed: 0 });
    const grid = CUTS.map((cut) => SLIPS.map((slip) => runModel(c0, { hiring: 1, matCut: cut, evalCut: cut, capexCut: cut, slip }).keys.fc));
    const aop = runModel(cfg, { hiring: 1, matCut: 0, evalCut: 0, capexCut: 0, slip: 0 }).keys.aop;
    return { cuts: CUTS, slips: SLIPS, grid, varGrid: grid.map((r) => r.map((v) => aop - v)) };
  }

  /** Auto-fill helpers. */
  function planLines(p) {
    const r = runProgram(p, 0, { hiring: 1, matCut: 0, evalCut: 0, capexCut: 0, slip: 0 });
    return { headcount: r.headcount, materials: r.materials, eval: r.eval, capex: r.capex };
  }
  function autoAop(p, cushion) {
    const pl = planLines(p), out = {};
    for (const l of LINES) out[l] = pl[l].map((v) => v * (1 + cushion));
    return out;
  }
  function autoActuals(p, closed, drift) {
    const pl = planLines(p), out = {};
    for (const l of LINES) out[l] = pl[l].map((v, i) => (i < closed ? v * (1 + (drift[l] || 0)) : 0));
    return out;
  }

  /** The 11 checks shown in the app and in the exported workbook. Each is a difference that should be zero. */
  function checks(cfg, lev, res, sens, scenarioNames, base) {
    const c = res.closed, L = res.lines, P = res.port;
    const sOpen = (a, b) => sum(a.map((v, i) => (i >= c ? v - b[i] : 0)));
    const out = [
      ['Total spend = Opex + Capex (portfolio forecast)', L.total.fc - (L.opex.fc + L.capex.fc)],
      ['Portfolio total = sum of the programs', L.total.fc - sum(res.progLines.map((x) => x.total.fc))],
      ['Closed months in the forecast of record = actuals', sum(P.forecast.total.slice(0, c)) - sum(P.act.total.slice(0, c))],
      ['Open months in the forecast of record = scenario forecast', sOpen(P.forecast.total, P.scenario.total)],
      ['Bridge steps add up to the forecast', res.bridge[4].run - res.bridge[5].change],
      ['Four quarters add up to the full year', sum(res.quarters.map((q) => q.fc)) - L.total.fc],
      ['Depreciation roll-forward: accumulated + net book value = Capex', res.dep.acc[11] + res.dep.nbv[11] - sum(res.dep.capex)],
      ['Sensitivity grid at 0% cut and 0 slip = scenario forecast (Base only)', lev.hiring === 1 && lev.matCut === 0 && lev.evalCut === 0 && lev.capexCut === 0 && lev.slip === 0 ? sens.grid[0][0] - sum(res.port.scenario.total) : 0],
      ['Scenario name is valid (0 = valid)', scenarioNames.includes(base) ? 0 : 1],
      ['Months closed is a whole number from 0 to 12 (0 = valid)', Number.isInteger(+cfg.closed) && cfg.closed >= 0 && cfg.closed <= 12 ? 0 : 1],
      ['Every closed month has actuals typed in (months missing)', P.act.total.slice(0, c).filter((v) => v === 0).length],
    ];
    const list = out.map(([label, diff]) => ({ label, diff, status: Math.abs(diff) < 0.5 ? 'OK' : 'CHECK' }));
    return { list, overall: list.some((x) => x.status === 'CHECK') ? `${list.filter((x) => x.status === 'CHECK').length} TO REVIEW` : 'ALL PASS' };
  }

  /** Run all scenarios plus sensitivity and checks. */
  function analyze(cfg, activeName) {
    const names = cfg.scenarios.map((s) => s.name);
    const sens = sensitivity(cfg);
    const runs = {};
    cfg.scenarios.forEach((s) => {
      const res = runModel(cfg, s);
      res.checks = checks(cfg, s, res, sens, names, s.name);
      runs[s.name] = res;
    });
    return { runs, sens, active: runs[activeName] || runs[names[0]], names };
  }

  return { LINES, LINE_NAMES, MON, SLIPS, CUTS, monthLabels, phase, runProgram, runModel, depreciation, sensitivity, autoAop, autoActuals, planLines, checks, analyze, sum };
});
