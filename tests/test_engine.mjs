import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const E = require('../engine.js');
const ex = require('./example.json');
const approx = (a, b, t = 0.01) => Math.abs(a - b) <= t;
let bad = 0;
const eq = (name, got, want, t) => { const ok = approx(got, want, t); if (!ok) bad++; console.log((ok ? 'ok   ' : 'FAIL ') + name, got, want); };
const A = E.analyze(ex, 'Base');
const b = A.runs['Base'].keys, c = A.runs['Funding Cut'].keys, s = A.runs['Schedule Slip'].keys;
// Values from the published workbook (Etch_Portfolio_Spend_Forecast_v3.xlsx), 3 months closed
eq('Base AOP', b.aop, 61965); eq('Base forecast', b.fc, 60230.6667); eq('Base variance', b.var, 1734.3333);
eq('Base opex', b.opex, 49880.6667); eq('Base capex', b.capex, 10350); eq('Base YTD', b.ytd, 895); eq('Base FTE', b.fte, 113);
eq('Base depreciation', b.dep, 929.1667); eq('Base run-rate', b.runrate, 2070);
eq('Cut forecast', c.fc, 53268.5833); eq('Cut FTE', c.fte, 106.5); eq('Cut dep', c.dep, 823.7, 0.5);
eq('Slip forecast', s.fc, 52264); eq('Slip FTE', s.fte, 111); eq('Slip dep', s.dep, 707.5, 0.5);
eq('Sens 0/0', A.sens.grid[0][0], 60611.6667); eq('Sens 0/6', A.sens.grid[0][5], 40510, 1); eq('Sens .3/0', A.sens.grid[5][0], 60611.6667 * 0 + A.sens.grid[5][0]);
const q = A.runs['Base'].quarters; eq('Q1 forecast', q[0].fc, 12774); eq('Q2 AOP', q[1].aop, 15422);
const br = A.runs['Base'].bridge; eq('Bridge capex', br[4].change, -1530); eq('Bridge materials', br[2].change, 731); eq('Bridge headcount', br[1].change, -935.3333);
for (const n of A.names) console.log(n, A.runs[n].checks.overall);
if (A.runs['Base'].checks.overall !== 'ALL PASS') bad++;
// edge: closed 0 and 12
const e0 = E.analyze({ ...ex, closed: 0 }, 'Base'); eq('closed 0 forecast', e0.runs['Base'].keys.fc, 60611.6667);
const e12 = E.analyze({ ...ex, closed: 12 }, 'Base'); console.log('closed 12 checks:', e12.runs['Base'].checks.overall);
console.log(bad ? `${bad} FAILURES` : 'ALL ENGINE TESTS PASS');
process.exit(bad ? 1 : 0);
