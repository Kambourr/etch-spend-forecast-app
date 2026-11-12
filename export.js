/* Etch Spend Forecast Builder: Excel export.
   Builds a six-tab workbook with live formulas from the app inputs (N programs).
   Layout mirrors the published Etch Portfolio Spend Forecast model. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.EtchExport = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const NUM = '#,##0;[Red](#,##0);"-"', PCT = '0.0%;[Red](0.0%);"-"', FTE = '0.0';
  const FONT = { name: 'Calibri', size: 11 };
  const BOLD = { name: 'Calibri', size: 11, bold: true };
  const TITLE = { name: 'Calibri', size: 14, bold: true };
  const SEC = { name: 'Calibri', size: 12, bold: true };
  const GREY = { name: 'Calibri', size: 11, italic: true, color: { argb: 'FF595959' } };
  const thin = { style: 'thin' };
  const BOX = { top: thin, left: thin, bottom: thin, right: thin };
  const UNDER = { bottom: thin }, OVER = { top: thin };
  const LINES = ['Headcount cost', 'Materials', 'Prototype & eval tools', 'Total Opex', 'Capex', 'Total spend'];
  const LK = ['headcount', 'materials', 'eval', null, 'capex', null];
  const colL = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const MC = (k) => colL(3 + k); // month column letter, k = 0..11

  function put(ws, r, c, v, fmt, font, align, border) {
    const cell = ws.getCell(r, c);
    if (typeof v === 'string' && v[0] === '=') cell.value = { formula: v.slice(1) }; else cell.value = v;
    cell.font = font || FONT;
    if (fmt) cell.numFmt = fmt;
    if (align) cell.alignment = align;
    if (border) cell.border = border;
    return cell;
  }
  const hdr = (ws, r, labels, start = 1) => labels.forEach((t, i) => put(ws, r, start + i, t, null, BOLD, { horizontal: i < 2 && start === 1 ? 'left' : 'right', wrapText: true, vertical: 'bottom' }, UNDER));
  const right = { horizontal: 'right' };

  function buildWorkbook(ExcelJS, cfg, A, activeName) {
    const N = cfg.programs.length;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Etch Spend Forecast Builder';
    wb.calcProperties = { fullCalcOnLoad: true };
    const months = cfgMonths(cfg);
    const names = A.names;
    const active = A.runs[activeName] || A.runs[names[0]];

    const SHEETS = {};
    ['Summary', 'Inputs', 'Plan_vs_Actual', 'Model', 'Sensitivity', 'Data'].forEach((n) => (SHEETS[n] = wb.addWorksheet(n)));

    // ---------- layout map ----------
    const IN = { T: 19 + N, H: 20 + N, MN: 21 + N, R0: 22 + N };          // drivers title, header, month#, first driver row
    IN.TH = IN.R0 + 4 * N + 1;                                           // thresholds title
    const AB = Array.from({ length: N }, (_, p) => 8 + 7 * p), AP = 8 + 7 * N;
    const CB = Array.from({ length: N }, (_, p) => AP + 9 + 7 * p), CP = CB[0] + 7 * N;
    const VB = Array.from({ length: N }, (_, p) => CP + 10 + 7 * p), VP = VB[0] + 7 * N;
    const QT = VP + 7;
    const SS = Array.from({ length: N }, (_, p) => 9 + 9 * p), SP = 9 + 9 * N;
    const FB = Array.from({ length: N }, (_, p) => SP + 10 + 7 * p), FP = FB[0] + 7 * N, AF = FP + 7;
    const DB = Array.from({ length: N }, (_, p) => FP + 11 + 5 * p), DP = DB[0] + 5 * N, RR = DP + 5;
    const AOP_TOT = AP + 5, FOR_TOT = FP + 5, ACT_TOT = CP + 5, SCN_TOT = SP + 6;
    const IR = (p, d) => IN.R0 + 4 * p + d; // Inputs driver row

    // ---------- Inputs ----------
    const wi = SHEETS['Inputs'];
    put(wi, 1, 1, 'Inputs and Assumptions', null, TITLE);
    put(wi, 2, 1, 'Figures in $K unless noted. Boxed cells are inputs you can change. Everything else is a formula.');
    put(wi, 4, 1, 'Active scenario'); put(wi, 4, 2, active ? activeName : names[0], null, null, right, BOX); put(wi, 4, 3, 'Pick one of the scenario names in row 9');
    put(wi, 5, 1, 'Fiscal months closed (0 to 12)'); put(wi, 5, 2, cfg.closed, '0', null, right, BOX); put(wi, 5, 3, 'Closed months use Actuals. Open months use the forecast.');
    put(wi, 6, 1, 'Latest closed month'); put(wi, 6, 2, `=IF(B5=0,"None",INDEX($C$${IN.H}:$N$${IN.H},B5))`, null, null, right);
    wi.getCell('B4').dataValidation = { type: 'list', allowBlank: false, formulae: [`"${names.join(',')}"`] };
    wi.getCell('B5').dataValidation = { type: 'whole', operator: 'between', formulae: [0, 12] };
    put(wi, 8, 1, 'Scenario levers', null, BOLD);
    hdr(wi, 9, ['Lever', ...names, 'Active']);
    const lev = [['Hiring plan achieved (%)', 'hiring', true], ['Materials cut (%)', 'matCut', true], ['Prototype / eval tool cut (%)', 'evalCut', true], ['Capex cut (%)', 'capexCut', true], ['Schedule slip (months)', 'slip', false]];
    lev.forEach(([lab, key, pct], i) => {
      const r = 10 + i; put(wi, r, 1, lab);
      cfg.scenarios.forEach((s, j) => put(wi, r, 2 + j, s[key], pct ? '0%' : '0', null, null, BOX));
      put(wi, r, 5, `=INDEX(B${r}:D${r},MATCH($B$4,$B$9:$D$9,0))`, pct ? '0%' : '0');
    });
    put(wi, 10, 6, 'Share of planned hires that actually happen in open months'); put(wi, 14, 6, 'Open-month spend moves this many months later');
    put(wi, 16, 1, 'Program assumptions', null, BOLD);
    hdr(wi, 17, ['Program', 'Starting FTE', 'Loaded cost per FTE ($K per year)', 'Eval tool unit cost ($K)', 'Equipment useful life (years)', 'Months until depreciation starts']);
    wi.getRow(17).height = 48;
    cfg.programs.forEach((p, i) => {
      const r = 18 + i; put(wi, r, 1, p.name);
      [p.startFTE, p.loaded, p.unitCost].forEach((v, j) => put(wi, r, 2 + j, +v || 0, '#,##0.##', null, null, BOX));
      put(wi, r, 5, +p.life || 0, '0.##', null, null, BOX); put(wi, r, 6, +p.delay || 0, '0', null, null, BOX);
    });
    put(wi, 18, 7, 'FTE = full-time equivalent (one full-time employee)');
    put(wi, 19, 7, 'Depreciation columns are illustrative: straight-line, starting after build and qualification');
    put(wi, IN.T, 1, 'Monthly driver plan (base case, before scenario levers)', null, BOLD);
    hdr(wi, IN.H, ['Program', 'Driver', ...months, 'FY total']);
    put(wi, IN.MN, 2, 'Fiscal month #');
    for (let k = 0; k < 12; k++) put(wi, IN.MN, 3 + k, k + 1, null, null, right);
    const drv = [['Planned hires (FTE)', 'hires', FTE], ['Materials ($K)', 'materials', NUM], ['Eval tool builds (units)', 'units', '#,##0.##'], ['Capex ($K)', 'capex', NUM]];
    cfg.programs.forEach((p, i) => drv.forEach(([lab, key, fmt], d) => {
      const r = IR(i, d);
      if (d === 0) put(wi, r, 1, p.name, null, BOLD);
      put(wi, r, 2, lab);
      for (let k = 0; k < 12; k++) put(wi, r, 3 + k, +p[key][k] || 0, fmt, null, null, BOX);
      put(wi, r, 15, `=SUM(C${r}:N${r})`, fmt);
    }));
    put(wi, IN.TH, 1, 'Variance review thresholds', null, BOLD);
    put(wi, IN.TH + 1, 1, 'Dollar threshold ($K)'); put(wi, IN.TH + 1, 2, +cfg.thresholds.usd, '#,##0', null, null, BOX);
    put(wi, IN.TH + 2, 1, 'Percent threshold'); put(wi, IN.TH + 2, 2, +cfg.thresholds.pct, '0%', null, null, BOX);
    put(wi, IN.TH + 1, 3, 'A line is flagged Review only when its variance is above both thresholds');
    wi.getColumn(1).width = 30; wi.getColumn(2).width = 24; wi.getColumn(15).width = 12;
    for (let k = 3; k <= 14; k++) wi.getColumn(k).width = 11;
    wi.views = [{ state: 'frozen', ySplit: 3 }];

    // ---------- Plan_vs_Actual ----------
    const wp = SHEETS['Plan_vs_Actual'];
    put(wp, 1, 1, 'Plan vs Actual: AOP budget, actuals and variance', null, TITLE);
    put(wp, 2, 1, 'Three stacked sections share the same month columns. $K. Variance = AOP minus forecast, so positive = favorable (underspend).');
    hdr(wp, 4, ['Program', 'Line', ...months, 'Total']);
    put(wp, 5, 2, 'Fiscal month #');
    for (let k = 0; k < 12; k++) put(wp, 5, 3 + k, k + 1, null, null, right);
    const portfolio = (ws, top, bases) => {
      put(ws, top, 1, 'Portfolio total', null, BOLD);
      LINES.forEach((ln, i) => {
        const r = top + i; put(ws, r, 2, ln, null, ln === 'Total Opex' || ln === 'Total spend' ? BOLD : null);
        for (let k = 0; k < 13; k++) {
          const c = colL(3 + k);
          const f = ln === 'Total Opex' ? `=SUM(${c}${top}:${c}${top + 2})` : ln === 'Total spend' ? `=${c}${top + 3}+${c}${top + 4}` : '=' + bases.map((b) => `${c}${b + i}`).join('+');
          put(ws, r, 3 + k, f, NUM, ln === 'Total Opex' || ln === 'Total spend' ? BOLD : null, null, ln === 'Total Opex' ? OVER : null);
        }
      });
    };
    const gridSection = (ws, bases, top, key, onlyClosed) => {
      bases.forEach((b, p) => {
        put(ws, b, 1, cfg.programs[p].name, null, BOLD);
        LINES.forEach((ln, i) => {
          const r = b + i; const tot = ln === 'Total Opex' || ln === 'Total spend';
          put(ws, r, 2, ln, null, tot ? BOLD : null);
          for (let k = 0; k < 12; k++) {
            const c = colL(3 + k);
            if (ln === 'Total Opex') put(ws, r, 3 + k, `=SUM(${c}${b}:${c}${b + 2})`, NUM, BOLD, null, OVER);
            else if (ln === 'Total spend') put(ws, r, 3 + k, `=${c}${b + 3}+${c}${b + 4}`, NUM, BOLD);
            else if (!onlyClosed || k < cfg.closed) put(ws, r, 3 + k, +cfg.programs[p][key][LK[i]][k] || 0, NUM);
          }
          put(ws, r, 15, `=SUM(C${r}:N${r})`, NUM, tot ? BOLD : null, null, ln === 'Total Opex' ? OVER : null);
          if (ln === 'Total Opex') put(ws, r, 2, ln, null, BOLD, null, OVER);
        });
      });
      portfolio(ws, top, bases);
    };
    put(wp, 6, 1, 'A. AOP budget (annual operating plan, locked)', null, SEC);
    gridSection(wp, AB, AP, 'aop', false);
    put(wp, AP + 7, 1, `B. Actuals (closed months only: ${months[0]} to ${months[Math.max(cfg.closed - 1, 0)]})`, null, SEC);
    gridSection(wp, CB, CP, 'act', true);
    put(wp, CP + 7, 1, 'C. Variance by program and line', null, SEC);
    put(wp, CP + 8, 1, 'FY forecast comes from the Model tab. YTD (year to date) covers closed months only.', null, GREY);
    hdr(wp, CP + 9, ['Program', 'Line', 'FY AOP', 'FY Forecast', 'Variance $', 'Variance %', 'F / U', 'YTD Actual', 'YTD AOP', 'YTD Variance $', 'Review?']);
    wp.getRow(CP + 9).height = 30;
    [...VB, VP].forEach((vb, g) => {
      const ab = g < N ? AB[g] : AP, cb = g < N ? CB[g] : CP, fb = g < N ? FB[g] : FP;
      for (let i = 0; i < 6; i++) {
        const r = vb + i; const tot = i === 3 || i === 5;
        if (i === 0) put(wp, r, 1, g === N ? 'Portfolio total' : cfg.programs[g].name, null, BOLD);
        put(wp, r, 2, LINES[i], null, tot ? BOLD : null, null, i === 3 ? OVER : null);
        const f = [
          [3, `=O${ab + i}`, NUM], [4, `=Model!O${fb + i}`, NUM], [5, `=C${r}-D${r}`, NUM], [6, `=IF(C${r}=0,0,E${r}/C${r})`, PCT],
          [7, `=IF(ABS(E${r})<0.5,"On plan",IF(E${r}>0,"F","U"))`, null],
          [8, `=SUMPRODUCT(($C$5:$N$5<=Inputs!$B$5)*C${cb + i}:N${cb + i})`, NUM], [9, `=SUMPRODUCT(($C$5:$N$5<=Inputs!$B$5)*C${ab + i}:N${ab + i})`, NUM], [10, `=I${r}-H${r}`, NUM],
          [11, `=IF(AND(ABS(E${r})>=Inputs!$B$${IN.TH + 1},ABS(F${r})>=Inputs!$B$${IN.TH + 2}),"Review","Within")`, null],
        ];
        f.forEach(([c, v, fmt]) => put(wp, r, c, v, fmt, tot ? BOLD : null, c === 7 || c === 11 ? right : null, i === 3 ? OVER : null));
      }
    });
    put(wp, QT, 1, 'D. Quarterly view (portfolio total spend)', null, SEC);
    hdr(wp, QT + 2, ['Quarter', 'Months', 'AOP', 'Forecast', 'Variance $', 'Variance %']);
    [['Q1', 'C', 'E'], ['Q2', 'F', 'H'], ['Q3', 'I', 'K'], ['Q4', 'L', 'N']].forEach(([q, a, z], i) => {
      const r = QT + 3 + i;
      put(wp, r, 1, `${q} ${cfg.fyLabel}`); put(wp, r, 2, `=${a}4&" to "&${z}4`);
      put(wp, r, 3, `=SUM(${a}${AOP_TOT}:${z}${AOP_TOT})`, NUM); put(wp, r, 4, `=SUM(Model!${a}${FOR_TOT}:${z}${FOR_TOT})`, NUM);
      put(wp, r, 5, `=C${r}-D${r}`, NUM); put(wp, r, 6, `=IF(C${r}=0,0,E${r}/C${r})`, PCT);
    });
    const QF = QT + 7;
    put(wp, QF, 1, 'Full year', null, BOLD);
    [[3, `=SUM(C${QT + 3}:C${QT + 6})`, NUM], [4, `=SUM(D${QT + 3}:D${QT + 6})`, NUM], [5, `=C${QF}-D${QF}`, NUM], [6, `=IF(C${QF}=0,0,E${QF}/C${QF})`, PCT]].forEach(([c, f, fmt]) => put(wp, QF, c, f, fmt, BOLD, null, OVER));
    wp.getColumn(1).width = 30; wp.getColumn(2).width = 24; wp.getColumn(15).width = 14;
    for (let k = 3; k <= 14; k++) wp.getColumn(k).width = 11;
    wp.views = [{ state: 'frozen', xSplit: 2, ySplit: 5 }];

    // ---------- Model ----------
    const wm = SHEETS['Model'];
    put(wm, 1, 1, 'Model: scenario forecast, forecast of record and Capex depreciation', null, TITLE);
    put(wm, 2, 1, 'All formulas ($K). Three stacked sections share the same month columns. Drivers come from Inputs.');
    put(wm, 3, 1, 'Scenario:'); put(wm, 3, 2, '=Inputs!B4', null, BOLD); put(wm, 3, 4, 'Months closed:'); put(wm, 3, 5, '=Inputs!B5');
    hdr(wm, 4, ['Program', 'Line', ...months, 'FY total']);
    put(wm, 5, 2, 'Fiscal month #'); put(wm, 6, 2, 'Driver month used');
    for (let k = 0; k < 12; k++) {
      const c = MC(k);
      put(wm, 5, 3 + k, `=Inputs!${c}${IN.MN}`);
      put(wm, 6, 3 + k, `=IF(${c}5<=Inputs!$B$5,${c}5,IF(${c}5-Inputs!$E$14>Inputs!$B$5,${c}5-Inputs!$E$14,0))`);
    }
    put(wm, 6, 15, '0 = no spend that month (slipped out)');
    put(wm, 7, 1, 'A. Scenario forecast (driver-based, all 12 months)', null, SEC);
    const SFL = ['Effective hires (FTE)', 'Ending FTE', 'Headcount cost', 'Materials', 'Prototype & eval tools', 'Total Opex', 'Capex', 'Total spend'];
    SS.forEach((s, p) => {
      put(wm, s, 1, cfg.programs[p].name, null, BOLD);
      SFL.forEach((ln, i) => put(wm, s + i, 2, ln, null, i === 5 || i === 7 ? BOLD : null));
      for (let k = 0; k < 12; k++) {
        const c = MC(k);
        const dr = (d) => `INDEX(Inputs!$C$${IR(p, d)}:$N$${IR(p, d)},${c}$6)`;
        put(wm, s, 3 + k, `=IF(${c}$6=0,0,${dr(0)})*IF(${c}$5<=Inputs!$B$5,1,Inputs!$E$10)`, FTE);
        put(wm, s + 1, 3 + k, `=Inputs!$B$${18 + p}+SUM($C${s}:${c}${s})`, FTE);
        put(wm, s + 2, 3 + k, `=${c}${s + 1}*Inputs!$C$${18 + p}/12`, NUM);
        put(wm, s + 3, 3 + k, `=IF(${c}$6=0,0,${dr(1)})*(1-Inputs!$E$11)`, NUM);
        put(wm, s + 4, 3 + k, `=IF(${c}$6=0,0,${dr(2)})*Inputs!$D$${18 + p}*(1-Inputs!$E$12)`, NUM);
        put(wm, s + 5, 3 + k, `=SUM(${c}${s + 2}:${c}${s + 4})`, NUM, BOLD, null, OVER);
        put(wm, s + 6, 3 + k, `=IF(${c}$6=0,0,${dr(3)})*(1-Inputs!$E$13)`, NUM);
        put(wm, s + 7, 3 + k, `=${c}${s + 5}+${c}${s + 6}`, NUM, BOLD);
      }
      put(wm, s, 15, `=SUM(C${s}:N${s})`, FTE); put(wm, s + 1, 15, `=N${s + 1}`, FTE);
      [2, 3, 4, 5, 6, 7].forEach((i) => put(wm, s + i, 15, `=SUM(C${s + i}:N${s + i})`, NUM, i === 5 || i === 7 ? BOLD : null, null, i === 5 ? OVER : null));
      put(wm, s + 5, 2, 'Total Opex', null, BOLD, null, OVER);
    });
    put(wm, SP, 1, 'Portfolio total', null, BOLD);
    ['Ending FTE', 'Headcount cost', 'Materials', 'Prototype & eval tools', 'Total Opex', 'Capex', 'Total spend'].forEach((ln, i) => {
      const r = SP + i; put(wm, r, 2, ln, null, i === 4 || i === 6 ? BOLD : null, null, i === 4 ? OVER : null);
      for (let k = 0; k < 13; k++) { const c = colL(3 + k); put(wm, r, 3 + k, '=' + SS.map((s) => `${c}${s + 1 + i}`).join('+'), i === 0 ? FTE : NUM, i === 4 || i === 6 ? BOLD : null, null, i === 4 ? OVER : null); }
    });
    put(wm, SP + 8, 1, 'B. Forecast of record (actuals for closed months, scenario forecast for open months)', null, SEC);
    const smap = { 0: 2, 1: 3, 2: 4, 4: 6 };
    FB.forEach((b, p) => {
      put(wm, b, 1, cfg.programs[p].name, null, BOLD);
      LINES.forEach((ln, i) => {
        const r = b + i; const tot = i === 3 || i === 5; put(wm, r, 2, ln, null, tot ? BOLD : null, null, i === 3 ? OVER : null);
        for (let k = 0; k < 12; k++) {
          const c = MC(k);
          const f = i in smap ? `=IF(${c}$5<=Inputs!$B$5,Plan_vs_Actual!${c}${CB[p] + i},${c}${SS[p] + smap[i]})` : i === 3 ? `=SUM(${c}${b}:${c}${b + 2})` : `=${c}${b + 3}+${c}${b + 4}`;
          put(wm, r, 3 + k, f, NUM, tot ? BOLD : null, null, i === 3 ? OVER : null);
        }
        put(wm, r, 15, `=SUM(C${r}:N${r})`, NUM, tot ? BOLD : null, null, i === 3 ? OVER : null);
      });
    });
    portfolio(wm, FP, FB);
    put(wm, AF, 2, 'Actual (A) or Forecast (F)');
    for (let k = 0; k < 12; k++) put(wm, AF, 3 + k, `=IF(${MC(k)}$5<=Inputs!$B$5,"A","F")`, null, null, right);
    put(wm, FP + 9, 1, 'C. Capex depreciation (straight-line, illustrative)', null, SEC);
    put(wm, FP + 10, 1, 'Capex in a month starts depreciating after the delay on Inputs. Memo only: not included in the Opex totals above.', null, GREY);
    const DL = ['Capex spent', 'Depreciation expense', 'Accumulated depreciation', 'Net book value'];
    DB.forEach((b, p) => {
      put(wm, b, 1, cfg.programs[p].name, null, BOLD); DL.forEach((ln, i) => put(wm, b + i, 2, ln));
      for (let k = 0; k < 12; k++) {
        const c = MC(k);
        put(wm, b, 3 + k, `=${c}${FB[p] + 4}`, NUM);
        put(wm, b + 1, 3 + k, `=SUMPRODUCT(($C$5:$N$5<=${c}$5-Inputs!$F$${18 + p})*$C${b}:$N${b})/(Inputs!$E$${18 + p}*12)`, NUM);
        put(wm, b + 2, 3 + k, `=SUM($C${b + 1}:${c}${b + 1})`, NUM);
        put(wm, b + 3, 3 + k, `=SUM($C${b}:${c}${b})-${c}${b + 2}`, NUM);
      }
      put(wm, b, 15, `=SUM(C${b}:N${b})`, NUM); put(wm, b + 1, 15, `=SUM(C${b + 1}:N${b + 1})`, NUM); put(wm, b + 2, 15, `=N${b + 2}`, NUM); put(wm, b + 3, 15, `=N${b + 3}`, NUM);
    });
    put(wm, DP, 1, 'Portfolio total', null, BOLD);
    DL.forEach((ln, i) => {
      const r = DP + i; put(wm, r, 2, ln, null, i === 1 || i === 3 ? BOLD : null);
      for (let k = 0; k < 13; k++) { const c = colL(3 + k); put(wm, r, 3 + k, '=' + DB.map((b) => `${c}${b + i}`).join('+'), NUM, i === 1 || i === 3 ? BOLD : null); }
    });
    put(wm, RR, 1, 'Annualized depreciation run-rate once all FY Capex is in service ($K per year)');
    put(wm, RR, 15, '=' + DB.map((b, p) => `O${b}/Inputs!$E$${18 + p}`).join('+'), NUM);
    wm.getColumn(1).width = 30; wm.getColumn(2).width = 26; wm.getColumn(15).width = 14;
    for (let k = 3; k <= 14; k++) wm.getColumn(k).width = 11;
    wm.views = [{ state: 'frozen', xSplit: 2, ySplit: 7 }];

    // ---------- Sensitivity (values from the app at export time) ----------
    const wsn = SHEETS['Sensitivity'];
    put(wsn, 1, 1, 'Sensitivity: plan spend by funding cut and schedule slip', null, TITLE);
    put(wsn, 2, 1, 'Values computed by the app when you exported (plan view: hiring at 100%, the cut applies evenly to materials, eval tools and Capex, actuals ignored). Not live: re-export after changing inputs. $K.');
    put(wsn, 4, 1, 'Total spend ($K)', null, BOLD);
    hdr(wsn, 5, ['Cut % \\ Slip (months)', ...A.sens.slips]);
    A.sens.cuts.forEach((cu, i) => {
      put(wsn, 6 + i, 1, cu, '0%', BOLD, { horizontal: 'left' });
      A.sens.grid[i].forEach((v, j) => put(wsn, 6 + i, 2 + j, Math.round(v * 100) / 100, NUM));
    });
    put(wsn, 13, 1, 'Variance vs AOP ($K, positive = under AOP)', null, BOLD);
    hdr(wsn, 14, ['Cut % \\ Slip (months)', ...A.sens.slips]);
    A.sens.cuts.forEach((cu, i) => {
      put(wsn, 15 + i, 1, cu, '0%', BOLD, { horizontal: 'left' });
      A.sens.varGrid[i].forEach((v, j) => put(wsn, 15 + i, 2 + j, Math.round(v * 100) / 100, NUM));
    });
    put(wsn, 22, 1, 'Scenario snapshot (values at export, active inputs)', null, BOLD);
    hdr(wsn, 23, ['Scenario', 'FY forecast', 'Variance to AOP', 'Variance %', 'Opex', 'Capex', 'Ending FTE']);
    names.forEach((n, i) => {
      const k = A.runs[n].keys, r = 24 + i; put(wsn, r, 1, n);
      [[k.fc, NUM], [k.var, NUM], [k.pct, PCT], [k.opex, NUM], [k.capex, NUM], [k.fte, FTE]].forEach(([v, f], j) => put(wsn, r, 2 + j, Math.round(v * 10000) / 10000, f));
    });
    wsn.getColumn(1).width = 46; for (let k = 2; k <= 8; k++) wsn.getColumn(k).width = 14;

    // ---------- Data (tidy) ----------
    const wd = SHEETS['Data'];
    put(wd, 1, 1, 'Data: one row per program, line and month', null, TITLE);
    put(wd, 2, 1, 'Clean long-format table for pivot tables, Tableau or Power BI. All values are formulas pointing at Plan_vs_Actual and Model. $K.');
    hdr(wd, 4, ['Program', 'Line', 'Type', 'Month #', 'Month', 'Fiscal quarter', 'AOP', 'Actual', 'Forecast of record', 'Variance (AOP - forecast)', 'A / F']);
    wd.getRow(4).height = 32;
    let dr = 5;
    cfg.programs.forEach((p, pi) => [[0, 'Headcount cost', 'Opex'], [1, 'Materials', 'Opex'], [2, 'Prototype & eval tools', 'Opex'], [4, 'Capex', 'Capex']].forEach(([i, ln, ty]) => {
      for (let k = 0; k < 12; k++) {
        const c = MC(k);
        [p.name, ln, ty, k + 1, months[k], `Q${Math.floor(k / 3) + 1}`, `=Plan_vs_Actual!${c}${AB[pi] + i}`, `=Plan_vs_Actual!${c}${CB[pi] + i}`, `=Model!${c}${FB[pi] + i}`, `=G${dr}-I${dr}`, `=Model!${c}$${AF}`]
          .forEach((v, j) => put(wd, dr, 1 + j, v, j >= 6 && j <= 9 ? NUM : null, null, j === 3 || j === 10 ? right : null));
        dr++;
      }
    }));
    wd.autoFilter = `A4:K${dr - 1}`;
    [30, 24, 8, 9, 9, 10, 11, 11, 14, 16, 7].forEach((w, i) => (wd.getColumn(i + 1).width = w));
    wd.views = [{ state: 'frozen', ySplit: 4 }];

    // ---------- Summary ----------
    const ws = SHEETS['Summary'];
    const V = 'Plan_vs_Actual';
    put(ws, 1, 1, `Etch Portfolio Spend Forecast: ${cfg.fyLabel} Summary`, null, TITLE);
    put(ws, 2, 1, `Generated by the Etch Spend Forecast Builder. Illustrative data unless you replaced it. $K (thousands of US dollars). Fiscal year ${months[0]} to ${months[11]}.`);
    put(ws, 3, 1, 'Scenario', null, BOLD); put(ws, 3, 2, '=Inputs!B4', null, null, right); put(ws, 3, 3, 'Closed through', null, BOLD); put(ws, 3, 4, '=Inputs!B6', null, null, right); put(ws, 3, 5, 'Model checks', null, BOLD);
    hdr(ws, 5, ['Key numbers', 'Value', 'Note']);
    const keys = [['AOP total spend ($K)', `=${V}!C${VP + 5}`, NUM, 'Opex plus Capex'], ['Forecast total spend ($K)', `=${V}!D${VP + 5}`, NUM, 'Actuals to date plus scenario forecast'],
      ['Variance to AOP ($K)', `=${V}!E${VP + 5}`, NUM, 'Positive = favorable (under AOP)'], ['Variance to AOP (%)', `=${V}!F${VP + 5}`, PCT, ''],
      ['Opex forecast ($K)', `=${V}!D${VP + 3}`, NUM, 'Operating expenses'], ['Capex forecast ($K)', `=${V}!D${VP + 4}`, NUM, 'Capital expenditures'],
      ['YTD actual vs AOP ($K)', `=${V}!J${VP + 5}`, NUM, 'Positive = favorable'], ['Ending FTE (scenario)', `=Model!O${SP}`, FTE, 'Full-time equivalents at fiscal year end'],
      ['Depreciation on new Capex ($K)', `=Model!O${DP + 1}`, NUM, 'Memo: not in the Opex total'], ['Annual depreciation run-rate ($K)', `=Model!O${RR}`, NUM, 'Future yearly cost once all Capex is in service']];
    keys.forEach(([a, f, fmt, n], i) => { put(ws, 6 + i, 1, a); put(ws, 6 + i, 2, f, fmt); put(ws, 6 + i, 3, n, null, GREY); });
    hdr(ws, 17, ['Read-out (updates with the scenario)']);
    ws.mergeCells('A17:G17');
    const reads = [
      '="Total spend is forecast at $"&TEXT(B7/1000,"0.0")&"M against an AOP of $"&TEXT(B6/1000,"0.0")&"M, which is $"&TEXT(ABS(B8)/1000,"0.0")&"M "&IF(B8>=0,"under","over")&" plan ("&TEXT(ABS(B9),"0.0%")&")."',
      '="The biggest mover is "&INDEX(A27:A30,MATCH(MAX(F27:F30),F27:F30,0))&", at $"&TEXT(MAX(F27:F30)/1000,"0.00")&"M "&IF(INDEX(B27:B30,MATCH(MAX(F27:F30),F27:F30,0))<=0,"under","over")&" AOP."',
      '="Year to date, actual spend is $"&TEXT(ABS(B12)/1000,"0.00")&"M "&IF(B12>=0,"under","over")&" AOP"&IF(B8=0,".",", which is "&TEXT(B12/B8,"0%")&" of the full-year variance. Check whether any of it is timing that reverses.")',
      '="Capex of $"&TEXT(B11/1000,"0.0")&"M this year adds about $"&TEXT(B15/1000,"0.00")&"M a year of depreciation once it is all in service."',
      `=COUNTIF(E27:E30,"Review")&" of 4 spend lines are above the review thresholds ($"&TEXT(Inputs!B${IN.TH + 1},"#,##0")&"K and "&TEXT(Inputs!B${IN.TH + 2},"0%")&")."`,
    ];
    reads.forEach((f, i) => { const r = 18 + i; put(ws, r, 1, f, null, null, { wrapText: true, vertical: 'top' }); ws.mergeCells(`A${r}:G${r}`); ws.getRow(r).height = i === 0 || i === 2 ? 30 : 16; });
    put(ws, 24, 1, 'AOP to forecast bridge (portfolio)', null, SEC);
    hdr(ws, 25, ['Step', 'Change in spend ($K)', 'Running total ($K)', 'F / U', 'Review?', 'Size of move ($K)']); ws.getRow(25).height = 30;
    put(ws, 26, 1, 'AOP total spend', null, BOLD); put(ws, 26, 2, `=${V}!C${VP + 5}`, NUM); put(ws, 26, 3, '=B26', NUM);
    [['Headcount cost', VP], ['Materials', VP + 1], ['Prototype & eval tools', VP + 2], ['Capex', VP + 4]].forEach(([nm, vr], i) => {
      const r = 27 + i; put(ws, r, 1, nm); put(ws, r, 2, `=-${V}!E${vr}`, NUM); put(ws, r, 3, `=C${r - 1}+B${r}`, NUM);
      put(ws, r, 4, `=IF(ABS(B${r})<0.5,"On plan",IF(B${r}<0,"F","U"))`, null, null, right);
      put(ws, r, 5, `=IF(AND(ABS(B${r})>=Inputs!$B$${IN.TH + 1},ABS(${V}!F${vr})>=Inputs!$B$${IN.TH + 2}),"Review","Within")`, null, null, right);
      put(ws, r, 6, `=ABS(B${r})`, NUM);
    });
    put(ws, 31, 1, 'Forecast total spend', null, BOLD, null, OVER); put(ws, 31, 2, `=${V}!D${VP + 5}`, NUM, BOLD, null, OVER); put(ws, 31, 3, '=C30', NUM, BOLD, null, OVER);
    for (let c = 4; c <= 6; c++) put(ws, 31, c, null, null, BOLD, null, OVER);
    put(ws, 32, 1, 'A negative change means spending below AOP (favorable). A positive change means spending above AOP (unfavorable).', null, GREY);
    put(ws, 34, 1, 'By program', null, SEC);
    hdr(ws, 35, ['Program', 'AOP ($K)', 'Forecast ($K)', 'Variance $ (+ = fav.)', 'Variance %', 'Review?']); ws.getRow(35).height = 30;
    cfg.programs.forEach((p, i) => {
      const r = 36 + i, vr = VB[i] + 5; put(ws, r, 1, p.name);
      ['C', 'D', 'E', 'F'].forEach((col, j) => put(ws, r, 2 + j, `=${V}!${col}${vr}`, col === 'F' ? PCT : NUM));
      put(ws, r, 6, `=${V}!K${vr}`, null, null, right);
    });
    const qT = 37 + N; // by quarter title
    put(ws, qT, 1, 'By quarter', null, SEC); hdr(ws, qT + 1, ['Quarter', 'AOP ($K)', 'Forecast ($K)']);
    for (let i = 0; i < 4; i++) { put(ws, qT + 2 + i, 1, `Q${i + 1}`); put(ws, qT + 2 + i, 2, `=${V}!C${QT + 3 + i}`, NUM); put(ws, qT + 2 + i, 3, `=${V}!D${QT + 3 + i}`, NUM); }
    const cT = qT + 7;
    put(ws, cT, 1, 'Cumulative spend by month', null, SEC); hdr(ws, cT + 1, ['Month', 'Cumulative AOP ($K)', 'Cumulative forecast ($K)']);
    for (let k = 0; k < 12; k++) {
      const c = MC(k); put(ws, cT + 2 + k, 1, `=${V}!${c}4`);
      put(ws, cT + 2 + k, 2, `=SUM(${V}!$C$${AOP_TOT}:${c}${AOP_TOT})`, NUM); put(ws, cT + 2 + k, 3, `=SUM(Model!$C$${FOR_TOT}:${c}${FOR_TOT})`, NUM);
    }
    const kT = cT + 16; // checks title
    put(ws, kT, 1, 'Model checks', null, SEC); put(ws, kT + 1, 1, 'Each check is a difference that should be zero. Status turns to CHECK if it is not.', null, GREY);
    hdr(ws, kT + 2, ['Check', 'Difference', 'Status']);
    const lastVB = VB.map((vb) => `${V}!D${vb + 5}`).join('+');
    const chk = [
      ['Total spend = Opex + Capex (portfolio forecast)', `=${V}!D${VP + 5}-(${V}!D${VP + 3}+${V}!D${VP + 4})`],
      ['Portfolio total = sum of the programs', `=${V}!D${VP + 5}-(${lastVB})`],
      ['Closed months in the forecast of record = actuals', `=SUMPRODUCT((Model!$C$5:$N$5<=Inputs!$B$5)*Model!C${FOR_TOT}:N${FOR_TOT})-SUMPRODUCT((${V}!$C$5:$N$5<=Inputs!$B$5)*${V}!C${ACT_TOT}:N${ACT_TOT})`],
      ['Open months in the forecast of record = scenario forecast', `=SUMPRODUCT((Model!$C$5:$N$5>Inputs!$B$5)*(Model!C${FOR_TOT}:N${FOR_TOT}-Model!C${SCN_TOT}:N${SCN_TOT}))`],
      ['Bridge steps add up to the forecast', '=C30-B31'],
      ['Four quarters add up to the full year', `=${V}!D${QF}-${V}!D${VP + 5}`],
      ['Depreciation roll-forward: accumulated + net book value = Capex', `=Model!O${DP + 2}+Model!O${DP + 3}-Model!O${DP}`],
      ['Sensitivity grid at 0% cut and 0 slip = scenario forecast (neutral levers only)', `=IF(AND(Inputs!E10=1,Inputs!E11=0,Inputs!E12=0,Inputs!E13=0,Inputs!E14=0),Sensitivity!B6-Model!O${SCN_TOT},0)`],
      ['Scenario name is valid (0 = valid)', '=IF(ISNUMBER(MATCH(Inputs!B4,Inputs!B9:D9,0)),0,1)'],
      ['Months closed is a whole number from 0 to 12 (0 = valid)', '=IF(AND(Inputs!B5>=0,Inputs!B5<=12,Inputs!B5=INT(Inputs!B5)),0,1)'],
      ['Every closed month has actuals typed in (months missing)', `=SUMPRODUCT((${V}!$C$5:$N$5<=Inputs!$B$5)*(${V}!C${ACT_TOT}:N${ACT_TOT}=0))`],
    ];
    chk.forEach(([lab, f], i) => {
      const r = kT + 3 + i; put(ws, r, 1, lab, null, null, { wrapText: true, vertical: 'top' }); if (lab.length > 42) ws.getRow(r).height = 30;
      put(ws, r, 2, f, '#,##0.0;[Red](#,##0.0);"-"'); put(ws, r, 3, `=IF(ABS(B${r})<0.5,"OK","CHECK")`, null, null, right);
    });
    const kl = kT + 3 + chk.length - 1;
    put(ws, kl + 1, 1, 'Overall', null, BOLD, null, OVER);
    put(ws, kl + 1, 2, null, null, null, null, OVER);
    put(ws, kl + 1, 3, `=IF(COUNTIF(C${kT + 3}:C${kl},"CHECK")=0,"ALL PASS",COUNTIF(C${kT + 3}:C${kl},"CHECK")&" TO REVIEW")`, null, BOLD, right, OVER);
    put(ws, 3, 6, `=C${kl + 1}`, null, BOLD, right);
    const green = { font: { color: { argb: 'FF006100' }, bold: true } }, red = { font: { color: { argb: 'FF9C0006' }, bold: true } };
    ws.addConditionalFormatting({ ref: `C${kT + 3}:C${kl + 1}`, rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"OK"'], style: green }, { type: 'cellIs', operator: 'equal', formulae: ['"ALL PASS"'], style: green }, { type: 'cellIs', operator: 'equal', formulae: ['"CHECK"'], style: red }] });
    ws.addConditionalFormatting({ ref: 'F3', rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"ALL PASS"'], style: green }, { type: 'cellIs', operator: 'notEqual', formulae: ['"ALL PASS"'], style: red }] });
    const a0 = kl + 4;
    put(ws, a0, 1, 'About this workbook', null, SEC); hdr(ws, a0 + 1, ['Tab', 'What it does']);
    [['Summary', 'Key numbers, read-out, bridge and model checks.'], ['Inputs', 'Scenario switch, months closed, levers, program and depreciation assumptions, monthly driver plan, review thresholds.'],
      ['Plan_vs_Actual', 'AOP, actuals and variance by program and line, plus a quarterly view.'], ['Model', 'Scenario forecast, forecast of record and Capex depreciation.'],
      ['Sensitivity', 'Funding cut vs schedule slip grid and a scenario snapshot (values from export).'], ['Data', 'Tidy table, one row per program, line and month.']]
      .forEach(([a, b], i) => { put(ws, a0 + 2 + i, 1, a); put(ws, a0 + 2 + i, 2, b); });
    const n0 = a0 + 9;
    ['How to use it: pick a scenario in Inputs B4, set months closed in Inputs B5, then read this page. Boxed cells on Inputs are the only inputs, besides the typed AOP and actuals on Plan_vs_Actual.',
      'Sign convention: variance = AOP minus forecast. Positive means favorable (underspend). Negative, shown in red brackets, means unfavorable (overspend).',
      'Opex vs Capex: R&D spending is expensed as incurred (ASC 730, Accounting Standards Codification topic 730). Equipment that lasts more than a year is capitalized and depreciated.']
      .forEach((t, i) => put(ws, n0 + i, 1, t));
    [40, 20, 20, 14, 12, 17, 3].forEach((w, i) => (ws.getColumn(i + 1).width = w));
    ws.views = [{ state: 'frozen', ySplit: 4 }];

    wb.worksheets.forEach((s) => { s.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }; });
    wb.views = [{ activeTab: 0 }];
    return wb;
  }

  function cfgMonths(cfg) { return require_labels(cfg); }
  function require_labels(cfg) {
    const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return Array.from({ length: 12 }, (_, i) => { const m = (cfg.startMonth - 1 + i) % 12; const y = cfg.startYear + Math.floor((cfg.startMonth - 1 + i) / 12); return `${MON[m]}-${String(y).slice(2)}`; });
  }
  return { buildWorkbook };
});
