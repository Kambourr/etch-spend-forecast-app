/* Etch Spend Forecast Builder: interface. Inputs on the left, live results on the right. */
(function () {
  const E = window.EtchEngine, X = window.EtchExport;
  const LS_KEY = 'etch-forecast-builder-v1';
  const $ = (id) => document.getElementById(id);
  const clone = (o) => JSON.parse(JSON.stringify(o));
  const sum = E.sum;
  const r2 = (v) => Math.round((+v || 0) * 100) / 100;
  const num = (v) => { const n = parseFloat(v); return Number.isFinite(n) ? n : 0; };
  const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const M1 = (k) => `$${(k / 1000).toFixed(1)}M`, M2 = (k) => `$${(k / 1000).toFixed(2)}M`;
  const SM2 = (k) => `${k >= 0 ? '+' : '−'}$${Math.abs(k / 1000).toFixed(2)}M`;
  const K0 = (k) => (Math.round(k) === 0 ? '-' : (k < 0 ? '(' : '') + Math.abs(Math.round(k)).toLocaleString('en-US') + (k < 0 ? ')' : ''));
  const P1 = (p) => `${(Math.abs(p) * 100).toFixed(1)}%`;
  const LINE_LABEL = { headcount: 'Headcount cost', materials: 'Materials', eval: 'Prototype & eval tools', capex: 'Capex' };

  function h(tag, attrs, ...kids) {
    const e = document.createElement(tag);
    for (const k in attrs || {}) {
      const v = attrs[k];
      if (k === 'class') e.className = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else if (k === 'html') e.innerHTML = v;
      else if (v !== false && v != null) e.setAttribute(k, v === true ? '' : v);
    }
    kids.flat().forEach((c) => { if (c != null && c !== false) e.append(c.nodeType ? c : document.createTextNode(c)); });
    return e;
  }

  // ---------------------------------------------------------------- state
  let cfg, ui = { tab: 'Base' }, last = null;
  function fresh() { const c = clone(window.ETCH_EXAMPLE); c.programs.forEach(initPhasing); return c; }
  function initPhasing(p) { p.phasing = p.phasing || {}; }
  function load() {
    try { const s = JSON.parse(localStorage.getItem(LS_KEY) || 'null'); if (s && s.cfg && s.cfg.programs && s.cfg.programs.length && s.v === 1) return s; } catch (e) { /* storage blocked */ }
    return null;
  }
  function save() { try { localStorage.setItem(LS_KEY, JSON.stringify({ v: 1, cfg, ui })); } catch (e) { /* ignore */ } }
  const saved = load();
  cfg = saved ? saved.cfg : fresh(); if (saved && saved.ui) ui = saved.ui;
  cfg.programs.forEach(initPhasing);

  const labels = () => E.monthLabels(cfg.startMonth, cfg.startYear);

  // ---------------------------------------------------------------- row editor (annual total + phasing + 12 months)
  function rowEditor(label, arr, holder, key, opts) {
    opts = opts || {};
    const kind0 = (holder.phasing && holder.phasing[key]) || 'custom';
    const annual = h('input', { type: 'number', step: 'any', class: 'num annual', 'aria-label': `${label} annual total`, value: r2(sum(arr)) });
    const sel = h('select', { class: 'phase', 'aria-label': `${label} phasing` }, ['even', 'front', 'back', 'custom'].map((v) => h('option', { value: v, selected: v === kind0 }, v[0].toUpperCase() + v.slice(1))));
    const cells = arr.map((v, k) => h('input', { type: 'number', step: 'any', class: 'num m', 'aria-label': `${label} month ${k + 1}`, value: r2(v) }));
    const setAll = (vals) => { vals.forEach((v, k) => { arr[k] = v; cells[k].value = r2(v); }); annual.value = r2(sum(arr)); };
    annual.addEventListener('change', () => {
      const t = num(annual.value);
      if (sel.value === 'custom') { const s = sum(arr); setAll(s ? arr.map((v) => (v * t) / s) : E.phase(t, 'even')); }
      else setAll(E.phase(t, sel.value));
      update();
    });
    sel.addEventListener('change', () => { holder.phasing = holder.phasing || {}; holder.phasing[key] = sel.value; if (sel.value !== 'custom') setAll(E.phase(num(annual.value), sel.value)); update(); });
    cells.forEach((c, k) => c.addEventListener('input', () => { arr[k] = num(c.value); sel.value = 'custom'; holder.phasing = holder.phasing || {}; holder.phasing[key] = 'custom'; annual.value = r2(sum(arr)); update(); }));
    const tr = h('tr', { 'data-actuals': opts.actuals ? '1' : null }, h('th', { scope: 'row' }, label), h('td', null, annual), h('td', null, sel), cells.map((c) => h('td', null, c)));
    tr._cells = cells;
    return tr;
  }
  const monthHead = (first) => h('thead', null, h('tr', null, h('th', null, first || ''), h('th', null, 'Annual'), h('th', null, 'Phasing'), labels().map((m) => h('th', null, m))));

  // ---------------------------------------------------------------- inputs panel
  function field(label, input) { return h('label', { class: 'f' }, label, input); }
  function tInput(val, cb, attrs) { const i = h('input', Object.assign({ class: 't', value: val }, attrs || {})); i.addEventListener('input', () => cb(i.value)); return i; }
  function nInput(val, cb, attrs) { const i = h('input', Object.assign({ class: 't', type: 'number', step: 'any', value: val }, attrs || {})); i.addEventListener('input', () => cb(num(i.value))); return i; }

  function renderInputs() {
    const root = $('inputs'); root.innerHTML = '';
    const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

    // 1. basics
    const monthSel = h('select', { class: 't' }, MON.map((m, i) => h('option', { value: i + 1, selected: i + 1 === cfg.startMonth }, m)));
    monthSel.addEventListener('change', () => { cfg.startMonth = +monthSel.value; renderInputs(); update(); });
    const closedIn = nInput(cfg.closed, (v) => { cfg.closed = Math.max(0, Math.min(12, Math.round(v))); syncActualsDisabled(); update(); }, { min: 0, max: 12, step: 1 });
    root.append(h('details', { class: 'card', open: true }, h('summary', null, h('h2', null, '1. Plan basics'), h('span', { class: 'hint' }, 'year, months closed, review rules')),
      h('div', { class: 'body' }, h('div', { class: 'grid2' },
        field('Fiscal year label', tInput(cfg.fyLabel, (v) => { cfg.fyLabel = v; update(); })),
        field('First fiscal month', monthSel),
        field('First month year', nInput(cfg.startYear, (v) => { cfg.startYear = Math.round(v) || 2026; renderInputs(); update(); }, { step: 1 })),
        field('Months closed (0 to 12)', closedIn),
        field('Review threshold ($K)', nInput(cfg.thresholds.usd, (v) => { cfg.thresholds.usd = v; update(); })),
        field('Review threshold (%)', nInput(r2(cfg.thresholds.pct * 100), (v) => { cfg.thresholds.pct = v / 100; update(); }))),
        h('p', { class: 'hint-line' }, 'A line is flagged Review only when its variance is above both thresholds. Closed months use actuals; open months use the forecast.'))));

    // 2. programs
    const progBox = h('div', { class: 'body' });
    cfg.programs.forEach((p, pi) => {
      const rows = h('tbody');
      [['Planned hires (FTE)', 'hires'], ['Materials ($K)', 'materials'], ['Eval tool builds (units)', 'units'], ['Capex ($K)', 'capex']].forEach(([lab, key]) => rows.append(rowEditor(lab, p[key], p, key)));
      progBox.append(h('div', { class: 'prog' },
        h('div', { class: 'prog-head' }, tInput(p.name, (v) => { p.name = v; update(); }, { 'aria-label': 'Program name' }),
          cfg.programs.length > 1 ? h('button', { class: 'btn link', type: 'button', onclick: () => { cfg.programs.splice(pi, 1); renderInputs(); update(); } }, 'Remove') : null),
        h('div', { class: 'grid2' },
          field('Starting FTE', nInput(p.startFTE, (v) => { p.startFTE = v; update(); })),
          field('Loaded cost per FTE ($K/yr)', nInput(p.loaded, (v) => { p.loaded = v; update(); })),
          field('Eval tool unit cost ($K)', nInput(p.unitCost, (v) => { p.unitCost = v; update(); })),
          field('Equipment life (years)', nInput(p.life, (v) => { p.life = v; update(); })),
          field('Months to depreciation', nInput(p.delay, (v) => { p.delay = v; update(); }))),
        h('div', { class: 'scroll' }, h('table', { class: 'rows' }, monthHead('Drivers (base plan)'), rows))));
    });
    progBox.append(h('button', { class: 'btn light', type: 'button', onclick: addProgram }, '+ Add program'),
      h('p', { class: 'hint-line' }, 'Type an annual total and pick a phasing (even, front-loaded, back-loaded), or edit any month. Loaded cost is the full yearly cost of one person.'));
    root.append(h('details', { class: 'card', open: true }, h('summary', null, h('h2', null, '2. Programs and drivers'), h('span', { class: 'hint' }, `${cfg.programs.length} program${cfg.programs.length > 1 ? 's' : ''}`)), progBox));

    // 3. AOP
    const cushion = h('input', { class: 't', type: 'number', step: 'any', value: 2.2, 'aria-label': 'AOP cushion percent' });
    const aopBox = h('div', { class: 'body' }, h('div', { class: 'grid2' }, field('AOP cushion over plan (%)', cushion),
      h('div', { style: 'display:flex;align-items:flex-end' }, h('button', { class: 'btn light', type: 'button', onclick: () => { cfg.programs.forEach((p) => { p.aop = E.autoAop(p, num(cushion.value) / 100); }); renderInputs(); update(); toast('AOP filled from the driver plan'); } }, 'Fill AOP from plan'))),
      h('p', { class: 'hint-line' }, 'The AOP (annual operating plan) is the locked budget. Fill it from the plan plus a cushion, or type your own numbers. Editing a row switches it to Custom.'));
    cfg.programs.forEach((p) => {
      const rows = h('tbody');
      E.LINES.forEach((l) => rows.append(rowEditor(LINE_LABEL[l], p.aop[l], p, 'aop_' + l)));
      aopBox.append(h('h3', { class: 'col-title' }, p.name), h('div', { class: 'scroll' }, h('table', { class: 'rows' }, monthHead('AOP ($K)'), rows)));
    });
    root.append(h('details', { class: 'card' }, h('summary', null, h('h2', null, '3. AOP (budget)'), h('span', { class: 'hint' }, 'locked plan by month')), aopBox));

    // 4. actuals
    const drift = { headcount: -3, materials: 10, eval: 0, capex: 0 };
    const driftIn = {};
    const driftFields = E.LINES.map((l) => { driftIn[l] = h('input', { class: 't', type: 'number', step: 'any', value: drift[l], 'aria-label': `${LINE_LABEL[l]} drift percent` }); return field(`${LINE_LABEL[l]} vs plan (%)`, driftIn[l]); });
    const actBox = h('div', { class: 'body' }, h('div', { class: 'grid2' }, driftFields),
      h('div', { style: 'margin:10px 0' }, h('button', { class: 'btn light', type: 'button', onclick: () => { cfg.programs.forEach((p) => { p.act = E.autoActuals(p, cfg.closed, Object.fromEntries(E.LINES.map((l) => [l, num(driftIn[l].value) / 100]))); }); renderInputs(); update(); toast('Actuals filled for the closed months'); } }, 'Fill actuals for closed months')),
      h('p', { class: 'hint-line' }, 'Actuals only count for closed months. Fill them as plan plus or minus a percent, then edit any month. In real life these come from the general ledger after close.'));
    cfg.programs.forEach((p) => {
      const rows = h('tbody');
      E.LINES.forEach((l) => rows.append(rowEditor(LINE_LABEL[l], p.act[l], p, 'act_' + l, { actuals: true })));
      actBox.append(h('h3', { class: 'col-title' }, p.name), h('div', { class: 'scroll' }, h('table', { class: 'rows' }, monthHead('Actuals ($K)'), rows)));
    });
    root.append(h('details', { class: 'card' }, h('summary', null, h('h2', null, '4. Actuals (closed months)'), h('span', { class: 'hint' }, 'booked results')), actBox));

    // 5. scenarios
    const lev = h('table', { class: 'lev' }, h('thead', null, h('tr', null, h('th', null, 'Lever'), cfg.scenarios.map((s) => h('th', null, s.name)))),
      h('tbody', null, [['Hiring plan achieved (%)', 'hiring', 100], ['Materials cut (%)', 'matCut', 100], ['Eval tool cut (%)', 'evalCut', 100], ['Capex cut (%)', 'capexCut', 100], ['Schedule slip (months)', 'slip', 1]].map(([lab, key, mult]) =>
        h('tr', null, h('th', { scope: 'row' }, lab), cfg.scenarios.map((s) => { const i = h('input', { class: 'num', type: 'number', step: 'any', value: r2(s[key] * mult), 'aria-label': `${s.name} ${lab}` }); i.addEventListener('input', () => { s[key] = key === 'slip' ? Math.max(0, Math.round(num(i.value))) : num(i.value) / mult; update(); }); return h('td', null, i); })))));
    root.append(h('details', { class: 'card' }, h('summary', null, h('h2', null, '5. Scenarios'), h('span', { class: 'hint' }, 'funding cut, schedule slip')),
      h('div', { class: 'body' }, h('div', { class: 'scroll' }, lev), h('p', { class: 'hint-line' }, 'A slip moves open-month spend later (closed months never move). Cuts reduce open-month materials, eval tools and Capex. Hiring is the share of planned hires that happen.'))));
    syncActualsDisabled();
  }
  function syncActualsDisabled() {
    document.querySelectorAll('tr[data-actuals="1"]').forEach((tr) => tr._cells.forEach((c, k) => { c.disabled = k >= cfg.closed; }));
  }
  function addProgram() {
    const n = cfg.programs.length + 1;
    const p = { name: `Program ${n}`, startFTE: 12, loaded: 255, unitCost: 2000, life: 5, delay: 2, phasing: { hires: 'even', materials: 'even', units: 'custom', capex: 'back' },
      hires: E.phase(4, 'even'), materials: E.phase(1800, 'even'), units: [0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0], capex: E.phase(1500, 'back') };
    p.aop = E.autoAop(p, 0.022); p.act = E.autoActuals(p, cfg.closed, { headcount: 0, materials: 0, eval: 0, capex: 0 });
    cfg.programs.push(p); renderInputs(); update(); toast(`Added ${p.name}. Edit its numbers, then fill the AOP and actuals.`);
  }

  // ---------------------------------------------------------------- charts
  const niceStep = (arr, raw, n = 5) => arr.find((s) => raw / s <= n) || arr[arr.length - 1];
  function lineChart(res, aopArr, seriesList, label) {
    const ml = labels(), W = 640, H = 300, l = 46, r = 12, t = 30, b = 34, cw = (W - l - r) / 12;
    const all = [...aopArr, ...seriesList.flatMap((s) => s.data)], max = Math.max(...all, 1);
    const step = niceStep([500, 1000, 2000, 2500, 5000, 10000, 20000], max * 1.05), top = Math.ceil((max * 1.05) / step) * step;
    const x = (i) => l + (i + 0.5) * cw, y = (v) => t + (1 - v / top) * (H - t - b);
    let g = '';
    for (let v = 0; v <= top; v += step) g += `<line class="g" x1="${l}" x2="${W - r}" y1="${y(v)}" y2="${y(v)}"/><text x="${l - 8}" y="${y(v) + 4}" text-anchor="end">$${v / 1000}M</text>`;
    const closed = res.closed;
    const shade = closed ? `<rect class="shade" x="${l}" y="${t}" width="${cw * closed}" height="${H - t - b}"/><text class="tag" x="${l + 8}" y="${t + 14}">ACTUALS</text>` : '';
    const pts = (a) => a.map((v, i) => `${x(i).toFixed(1)},${y(v).toFixed(1)}`).join(' ');
    let lines = `<polyline class="aop" points="${pts(aopArr)}"/>`, leg = `<line class="aop" x1="0" x2="22" y1="0" y2="0"/><text x="28" y="4">AOP</text>`, off = 70;
    seriesList.forEach((s, i) => { lines += `<polyline class="${s.cls}" points="${pts(s.data)}"/>`; leg += `<line class="${s.cls}" x1="${off}" x2="${off + 22}" y1="0" y2="0"/><text x="${off + 28}" y="4">${esc(s.name)}</text>`; off += 28 + s.name.length * 7 + 14; });
    const dots = seriesList.length === 1 ? seriesList[0].data.map((v, i) => `<circle class="${i < closed ? 'dot act' : 'dot'}" cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="4"/>`).join('') : '';
    const xl = ml.map((m, i) => `<text x="${x(i).toFixed(1)}" y="${H - 12}" text-anchor="middle">${m.split('-')[0]}</text>`).join('');
    return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(label)}">${shade}${g}${lines}${dots}${xl}<g transform="translate(${l + 90},10)">${leg}</g></svg>`;
  }
  function waterfall(res, label) {
    const W = 640, H = 300, t = 34, b = 40, l = 16, r = 16, names = ['AOP', 'Headcount', 'Materials', 'Eval tools', 'Capex', 'Forecast'];
    const runs = res.bridge.map((x) => x.run), lo = Math.floor((Math.min(...runs) - 1500) / 1000) * 1000, hi = Math.max(...runs) + 1000;
    const y = (v) => t + (1 - (v - lo) / (hi - lo)) * (H - t - b), slot = (W - l - r) / 6, bw = 62, cx = (i) => l + (i + 0.5) * slot;
    let s = '';
    res.bridge.forEach((st, i) => {
      if (i === 0 || i === 5) { s += `<rect class="bar total" x="${cx(i) - bw / 2}" y="${y(st.run)}" width="${bw}" height="${Math.max(y(lo) - y(st.run), 1)}"/><text class="val" x="${cx(i)}" y="${y(st.run) - 8}" text-anchor="middle">${M2(st.run)}</text>`; }
      else {
        const prev = res.bridge[i - 1].run, now = st.run, flat = Math.abs(st.change) < 0.5, cls = flat ? 'bar flat' : st.change < 0 ? 'bar fav' : 'bar unf';
        const yt = y(Math.max(prev, now)), yb = y(Math.min(prev, now));
        s += `<rect class="${cls}" x="${cx(i) - bw / 2}" y="${yt}" width="${bw}" height="${Math.max(yb - yt, 1.5)}"/><text class="val" x="${cx(i)}" y="${yt - 8}" text-anchor="middle">${flat ? '$0' : SM2(st.change)}</text>`;
      }
      if (i < 5) s += `<line class="link" x1="${cx(i) + bw / 2}" x2="${cx(i + 1) - bw / 2}" y1="${y(st.run)}" y2="${y(st.run)}"/>`;
      s += `<text x="${cx(i)}" y="${H - 18}" text-anchor="middle">${names[i]}</text>`;
    });
    s += `<text class="note" x="${l}" y="14">Axis starts at $${lo / 1000}M. Change in spend vs AOP, by line.</text>`;
    return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(label)}">${s}</svg>`;
  }
  function capexChart(res, label) {
    const ml = labels(), W = 640, H = 300, l = 50, r = 54, t = 30, b = 34, cap = res.dep.capex, dep = res.dep.dep;
    const cmax = Math.max(...cap, 1), dmax = Math.max(...dep, 1);
    const sl = niceStep([50, 100, 200, 250, 500, 1000, 2000, 5000], cmax * 1.05), L = sl * 5, sd = niceStep([1, 2, 5, 10, 20, 25, 50, 100, 200, 500], dmax * 1.05), D = sd * 5;
    const cw = (W - l - r) / 12, x = (i) => l + (i + 0.5) * cw, yl = (v) => t + (1 - v / L) * (H - t - b), yd = (v) => t + (1 - v / D) * (H - t - b);
    let g = '';
    for (let i = 0; i <= 5; i++) g += `<line class="g" x1="${l}" x2="${W - r}" y1="${yl(sl * i)}" y2="${yl(sl * i)}"/><text x="${l - 8}" y="${yl(sl * i) + 4}" text-anchor="end">$${((sl * i) / 1000).toFixed(sl < 100 ? 2 : 1)}M</text><text x="${W - r + 8}" y="${yl(sl * i) + 4}">$${sd * i}K</text>`;
    const bars = cap.map((v, i) => (v > 0 ? `<rect class="bar capex" x="${x(i) - cw * 0.3}" y="${yl(v)}" width="${cw * 0.6}" height="${yl(0) - yl(v)}"/>` : '')).join('');
    const line = `<polyline class="fc" points="${dep.map((v, i) => `${x(i).toFixed(1)},${yd(v).toFixed(1)}`).join(' ')}"/>`;
    const xl = ml.map((m, i) => `<text x="${x(i).toFixed(1)}" y="${H - 12}" text-anchor="middle">${m.split('-')[0]}</text>`).join('');
    const legend = `<g transform="translate(${l},12)"><rect class="bar capex" x="0" y="-6" width="14" height="10"/><text x="20" y="4">Capex spent (left)</text><line class="fc" x1="170" x2="192" y1="0" y2="0"/><text x="198" y="4">Depreciation (right)</text></g>`;
    return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(label)}">${g}${bars}${line}${xl}${legend}</svg>`;
  }

  // ---------------------------------------------------------------- results
  function results() {
    const A = last, names = A.names, compare = ui.tab === 'Compare';
    const nm = compare ? names[0] : names.includes(ui.tab) ? ui.tab : names[0];
    const R = A.runs[nm], k = R.keys, under = k.var >= 0, fy = esc(cfg.fyLabel);
    const tabs = [...names, 'Compare'].map((n) => `<button type="button" role="tab" aria-selected="${(compare ? 'Compare' : nm) === n}" data-tab="${esc(n)}">${esc(n)}</button>`).join('');
    const big = R.lines;
    const topLine = R.bridge.slice(1, 5).reduce((a, b) => (Math.abs(b.change) > Math.abs(a.change) ? b : a));
    const reviewN = R.bridge.slice(1, 5).filter((x) => x.review === 'Review').length;
    const read = [
      `Total spend is forecast at ${M1(k.fc)} against an AOP of ${M1(k.aop)}, which is ${M1(Math.abs(k.var))} ${under ? 'under' : 'over'} plan (${P1(k.pct)}).`,
      `The biggest mover is ${topLine.label}, at ${M2(Math.abs(topLine.change))} ${topLine.change <= 0 ? 'under' : 'over'} AOP.`,
      `Year to date, actual spend is ${M2(Math.abs(k.ytd))} ${k.ytd >= 0 ? 'under' : 'over'} AOP${k.var === 0 ? '.' : `, which is ${Math.round((k.ytd / k.var) * 100)}% of the full-year variance. Check whether any of it is timing that reverses.`}`,
      `Capex of ${M1(k.capex)} this year adds about ${M2(k.runrate)} a year of depreciation once it is all in service.`,
      `${reviewN} of 4 spend lines are above the review thresholds ($${cfg.thresholds.usd}K and ${Math.round(cfg.thresholds.pct * 100)}%).`,
    ];
    const ml = labels();
    let body = '';
    if (!compare) {
      body += `<div class="kpis">
        <div class="kpi main"><div class="l">${fy} forecast</div><div class="v">${M1(k.fc)}</div><div class="s">against ${M1(k.aop)} AOP</div></div>
        <div class="kpi"><div class="l">Variance to AOP</div><div class="v">${M2(Math.abs(k.var))}</div><div class="s"><span class="pill ${under ? '' : 'bad'}">${under ? 'under plan' : 'over plan'} · ${P1(k.pct)}</span></div></div>
        <div class="kpi"><div class="l">Opex · Capex</div><div class="v" style="font-size:22px">${M1(k.opex)} · ${M1(k.capex)}</div><div class="s">forecast</div></div>
        <div class="kpi"><div class="l">Depreciation</div><div class="v">${M2(k.dep)}</div><div class="s">in ${fy}; ${M2(k.runrate)} a year later</div></div>
        <div class="kpi"><div class="l">Ending FTE</div><div class="v">${k.fte % 1 ? k.fte.toFixed(1) : k.fte}</div><div class="s">full-time equivalents</div></div>
        <div class="kpi"><div class="l">Model checks</div><div class="v" style="font-size:22px">${R.checks.overall}</div><div class="s">${R.checks.list.length} tests</div></div></div>
        <div class="readout">${read.map((t) => `<p>${esc(t)}</p>`).join('')}</div>
        <div class="charts">
          <div class="chart-card"><h3>AOP to forecast bridge</h3>${waterfall(R, 'Bridge from the AOP to the forecast by line')}</div>
          <div class="chart-card"><h3>Total spend by month</h3>${lineChart(R, R.port.aop.total, [{ name: 'Forecast of record', cls: 'fc', data: R.port.forecast.total }], 'Monthly total spend, AOP versus forecast of record')}</div>
          <div class="chart-card"><h3>Capex becomes depreciation</h3>${capexChart(R, 'Monthly Capex and depreciation')}</div></div>`;
      body += `<div class="card"><div class="card-head"><h2>By program</h2></div><div class="body"><div class="scroll"><table class="tbl"><thead><tr><th>Program</th><th>AOP</th><th>Forecast</th><th>Variance $</th><th>Var %</th><th>Review</th></tr></thead><tbody>${R.progLines.map((pl, i) => `<tr><th scope="row">${esc(R.progs[i].name)}</th><td>${K0(pl.total.aop)}</td><td>${K0(pl.total.fc)}</td><td class="${pl.total.var < 0 ? 'neg' : ''}">${K0(pl.total.var)}</td><td>${(pl.total.pct * 100).toFixed(1)}%</td><td class="rev">${pl.total.review === 'Review' ? 'Review' : ''}</td></tr>`).join('')}<tr class="total"><th scope="row">Portfolio</th><td>${K0(big.total.aop)}</td><td>${K0(big.total.fc)}</td><td class="${big.total.var < 0 ? 'neg' : ''}">${K0(big.total.var)}</td><td>${(big.total.pct * 100).toFixed(1)}%</td><td class="rev">${big.total.review === 'Review' ? 'Review' : ''}</td></tr></tbody></table></div></div></div>`;
      body += `<div class="card"><div class="card-head"><h2>Variance by line (portfolio, $K)</h2><span class="hint">AOP minus forecast, + = favorable</span></div><div class="body"><div class="scroll"><table class="tbl"><thead><tr><th>Line</th><th>AOP</th><th>Forecast</th><th>Variance $</th><th>Var %</th><th>F/U</th><th>YTD var</th><th>Review</th></tr></thead><tbody>${['headcount', 'materials', 'eval', 'opex', 'capex', 'total'].map((l) => { const v = big[l]; return `<tr class="${l === 'opex' || l === 'total' ? 'total' : ''}"><th scope="row">${E.LINE_NAMES[l]}</th><td>${K0(v.aop)}</td><td>${K0(v.fc)}</td><td class="${v.var < 0 ? 'neg' : ''}">${K0(v.var)}</td><td>${(v.pct * 100).toFixed(1)}%</td><td>${v.fu}</td><td class="${v.ytdVar < 0 ? 'neg' : ''}">${K0(v.ytdVar)}</td><td class="rev">${v.review === 'Review' ? 'Review' : ''}</td></tr>`; }).join('')}</tbody></table></div></div></div>`;
    } else {
      const s1 = names[1], s2 = names[2];
      body += `<div class="charts"><div class="chart-card" style="grid-column:1/-1"><h3>Total spend by month: AOP and every scenario</h3>${lineChart(A.runs[names[0]], A.runs[names[0]].port.aop.total, names.map((n, i) => ({ name: n, cls: ['fc', 's1', 's2'][i] || 's2', data: A.runs[n].port.forecast.total })), 'Monthly total spend for each scenario against the AOP')}</div></div>`;
      body += `<div class="card"><div class="card-head"><h2>Scenario comparison</h2><span class="hint">${M1(A.runs[names[0]].keys.aop)} AOP</span></div><div class="body"><div class="scroll"><table class="tbl"><thead><tr><th></th>${names.map((n) => `<th>${esc(n)}</th>`).join('')}</tr></thead><tbody>${[
        ['Forecast ($K)', (x) => K0(x.fc)], ['Variance to AOP ($K)', (x) => K0(x.var)], ['Variance %', (x) => (x.pct * 100).toFixed(1) + '%'], ['Opex ($K)', (x) => K0(x.opex)], ['Capex ($K)', (x) => K0(x.capex)], ['Ending FTE', (x) => (x.fte % 1 ? x.fte.toFixed(1) : x.fte)], [`${fy} depreciation ($K)`, (x) => K0(x.dep)], ['Depreciation run-rate ($K/yr)', (x) => K0(x.runrate)],
      ].map(([lab, f]) => `<tr><th scope="row">${lab}</th>${names.map((n) => `<td>${f(A.runs[n].keys)}</td>`).join('')}</tr>`).join('')}<tr><th scope="row">Model checks</th>${names.map((n) => `<td>${A.runs[n].checks.overall}</td>`).join('')}</tr></tbody></table></div><p class="hint-line">A slip looks favorable, but the work still happens, so next year's budget has to absorb it. About ${M1(Math.max(...names.map((n) => A.runs[n].keys.fc)) - Math.min(...names.map((n) => A.runs[n].keys.fc)))} separates the highest and lowest scenario.</p></div></div>`;
    }
    // sensitivity + checks
    const S = A.sens, vmax = Math.max(...S.varGrid.flat().map(Math.abs), 1);
    body += `<div class="card"><div class="card-head"><h2>Sensitivity: plan spend by cut and slip</h2><span class="hint">variance to AOP, $K, + = under AOP</span></div><div class="body"><div class="scroll"><table class="tbl heat"><thead><tr><th>Cut \\ Slip (months)</th>${S.slips.map((s) => `<th>${s}</th>`).join('')}</tr></thead><tbody>${S.cuts.map((c, i) => `<tr><th scope="row">${Math.round(c * 100)}%</th>${S.varGrid[i].map((v) => `<td style="background:rgba(16,45,48,${(Math.min(Math.abs(v) / vmax, 1) * 0.28).toFixed(3)})" class="${v < 0 ? 'neg' : ''}">${K0(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div><p class="hint-line">Plan view: hiring at 100%, the cut applies evenly to materials, eval tools and Capex, actuals ignored. Darker = further from AOP.</p></div></div>`;
    body += `<div class="card"><div class="card-head"><h2>Model checks</h2><span class="hint">${esc(nm)}: ${R.checks.overall}</span></div><div class="body"><ul class="checks">${R.checks.list.map((c) => `<li><span>${esc(c.label)}</span><b class="${c.status === 'OK' ? 'ok' : 'bad'}">${c.status}</b></li>`).join('')}</ul></div></div>`;
    $('results').innerHTML = `<div class="tabs" role="tablist" aria-label="Scenario">${tabs}</div>${body}`;
    $('results').querySelectorAll('[data-tab]').forEach((b) => b.addEventListener('click', () => { ui.tab = b.dataset.tab; save(); results(); }));
  }

  let queued = false;
  function update() {
    if (queued) return; queued = true;
    requestAnimationFrame(() => { queued = false; last = E.analyze(cfg, ui.tab === 'Compare' ? cfg.scenarios[0].name : ui.tab); const col = $('col-out'), keep = col.scrollTop, wk = window.scrollY; results(); save(); col.scrollTop = keep; window.scrollTo(0, wk); });
  }
  function toast(msg) {
    const t = h('div', { class: 'toast', role: 'status' }, msg); document.body.append(t); setTimeout(() => t.remove(), 2600);
  }

  // ---------------------------------------------------------------- export
  async function exportXlsx() {
    if (!window.ExcelJS) { toast('The Excel library did not load. Check your connection and reload.'); return; }
    try {
      const A = E.analyze(cfg, ui.tab === 'Compare' ? cfg.scenarios[0].name : ui.tab);
      const active = ui.tab === 'Compare' ? cfg.scenarios[0].name : ui.tab;
      const wb = X.buildWorkbook(window.ExcelJS, cfg, A, active);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = h('a', { href: URL.createObjectURL(blob), download: `Etch_Spend_Forecast_${cfg.fyLabel.replace(/[^A-Za-z0-9]+/g, '') || 'FY'}_${active.replace(/\s+/g, '_')}.xlsx` });
      document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
      toast('Excel downloaded. Open it in Excel and let it recalculate.');
    } catch (e) { console.error(e); toast('Export failed: ' + e.message); }
  }

  // ---------------------------------------------------------------- boot
  $('btn-example').addEventListener('click', () => { cfg = fresh(); ui = { tab: 'Base' }; renderInputs(); update(); toast('Etch example loaded'); });
  $('btn-export').addEventListener('click', exportXlsx);
  renderInputs();
  last = E.analyze(cfg, ui.tab === 'Compare' ? cfg.scenarios[0].name : ui.tab);
  results();
  window.__etch = { get cfg() { return cfg; }, get last() { return last; }, update, exportXlsx };
})();
