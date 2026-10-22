// Copies templates.js (the single source of models + evaluation templates) into
// backend/SPARK_Code.gs between the SPARK_TEMPLATES markers.
// Run from the repo root:  node tools/sync-templates.js
// Then paste the whole Code.gs into Apps Script and redeploy (new version, same deployment).
const fs = require('fs');
const src = fs.readFileSync('templates.js', 'utf8');
const body = src.slice(src.indexOf('const SPARK = ')).trim();
const gsPath = 'backend/SPARK_Code.gs';
const gs = fs.readFileSync(gsPath, 'utf8');
const START = '// ---- SPARK_TEMPLATES (generated from templates.js — do not edit here) ----';
const END = '// ---- end SPARK_TEMPLATES ----';
const a = gs.indexOf(START), b = gs.indexOf(END);
if (a < 0 || b < 0) throw new Error('SPARK_TEMPLATES markers not found in ' + gsPath);
fs.writeFileSync(gsPath, gs.slice(0, a + START.length) + '\n' + body + '\n' + gs.slice(b));
console.log('Synced templates.js into ' + gsPath);
