import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const E = require('../engine.js');
const cfg = process.argv[2] ? require(process.argv[2]) : require('./example.json');
const name = process.argv[3] || 'Base';
const A = E.analyze(cfg, name); const r = A.runs[name];
console.log(JSON.stringify({ keys: r.keys, forTotal: r.port.forecast.total, dep: r.dep.dep, bridge: r.bridge.map((b) => b.change), overall: r.checks.overall, nprog: cfg.programs.length, grid00: A.sens.grid[0][0] }));
