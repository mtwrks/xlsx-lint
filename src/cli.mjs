#!/usr/bin/env node
/**
 * xlsx-lint — find formulas that break their own pattern, in CI.
 *
 * Reads .xlsx workbooks locally. No network, no upload, no telemetry.
 *
 *   xlsx-lint model.xlsx
 *   xlsx-lint "reports/*.xlsx" --severity high
 *   xlsx-lint model.xlsx --json
 *
 * Exit codes:
 *   0  no findings at or above the chosen severity
 *   1  findings at or above the chosen severity
 *   2  usage or read error
 */

import { readFileSync, existsSync, statSync, readdirSync } from "node:fs";
import { join, resolve, basename, dirname } from "node:path";
import { inspect, CHECKS } from "./engine.mjs";

const SEV_ORDER = { high: 0, med: 1, low: 2 };
const SEV_LABEL = { high: "review first", med: "worth checking", low: "note" };

const args = process.argv.slice(2);
const flag = (n) => args.includes(n);
const val = (n, d) => { const i = args.indexOf(n); return i >= 0 && args[i + 1] ? args[i + 1] : d; };

if (!args.length || flag("--help") || flag("-h")) {
  console.log(`
xlsx-lint — find formulas that break their own pattern, in CI

  xlsx-lint <file-or-glob>...        inspect one or more workbooks
  xlsx-lint . --recursive            every .xlsx under a directory

Options
  --severity <high|med|low>   minimum severity that fails the run (default: high)
  --json                      machine-readable output
  --quiet                     only print failures
  --list-checks               print the twelve checks and exit
  --recursive, -r             recurse into directories
  --version, -v

Runs entirely on your machine. No network, no upload, no telemetry.
`.trim());
  process.exit(0);
}

if (flag("--version") || flag("-v")) { console.log("1.0.0"); process.exit(0); }

if (flag("--list-checks")) {
  for (const c of CHECKS) console.log(`  ${c.sev.padEnd(5)} ${c.id.padEnd(15)} ${c.title}`);
  process.exit(0);
}

const minSev = val("--severity", "high");
if (!(minSev in SEV_ORDER)) { console.error(`xlsx-lint: unknown severity "${minSev}"`); process.exit(2); }
const asJson = flag("--json");
const quiet = flag("--quiet");
const recursive = flag("--recursive") || flag("-r");

/* ── collect targets ─────────────────────────────────────────────────────
   Deliberately no glob dependency: a shell usually expands patterns, and a
   directory walk covers what it does not. Keeping this at zero dependencies
   is the point of the tool. */
const isXlsx = (p) => /\.xlsx$/i.test(p) && !basename(p).startsWith("~$");
const walk = (dir, out = []) => {
  for (const e of readdirSync(dir, { withFileTypes: true })) {
    if (e.name === "node_modules" || e.name.startsWith(".")) continue;
    const p = join(dir, e.name);
    if (e.isDirectory()) { if (recursive) walk(p, out); }
    else if (isXlsx(p)) out.push(p);
  }
  return out;
};

const targets = [];
for (const a of args) {
  if (a.startsWith("-") || a === minSev) continue;
  if (!existsSync(a)) { console.error(`xlsx-lint: no such file or directory: ${a}`); process.exit(2); }
  statSync(a).isDirectory() ? walk(a, targets) : (isXlsx(a) ? targets.push(a) : null);
}
if (!targets.length) { console.error("xlsx-lint: no .xlsx files matched."); process.exit(2); }

/* ── run ─────────────────────────────────────────────────────────────────── */
const results = [];
let failing = 0;

for (const file of targets) {
  let r;
  try { r = inspect(readFileSync(resolve(file))); }
  catch (err) { console.error(`xlsx-lint: ${file}: ${err.message}`); process.exit(2); }

  const gating = r.findings.filter((f) => SEV_ORDER[f.sev] <= SEV_ORDER[minSev]);
  const total = gating.reduce((a, f) => a + f.count, 0);
  if (total) failing++;
  results.push({ file, r, gating, total });
}

if (asJson) {
  console.log(JSON.stringify({
    version: "1.0.0",
    severity: minSev,
    files: results.map(({ file, r, total }) => ({
      file,
      sheets: r.sheetsRead,
      formulaCells: r.totalFormulas,
      checksRun: r.checksRun,
      findings: r.findings.map((f) => ({ id: f.id, severity: f.sev, title: f.title, count: f.count, locations: f.items })),
      coverage: r.coverage.map((c) => ({ id: c.id, count: c.count })),
      gatingFindings: total,
    })),
    filesWithFindings: failing,
  }, null, 2));
  process.exit(failing ? 1 : 0);
}

for (const { file, r, gating, total } of results) {
  if (quiet && !total) continue;
  console.log(`\n${file}`);
  console.log(`  ${r.sheetsRead} sheet${r.sheetsRead === 1 ? "" : "s"} · ${r.totalFormulas.toLocaleString()} formula cells · ${r.checksRun} checks run`);

  if (!r.findings.length) {
    console.log("  no flagged patterns");
    console.log("  note: this does not establish that the workbook is correct — only that these 12 patterns were absent");
    continue;
  }
  for (const f of r.findings) {
    const gate = SEV_ORDER[f.sev] <= SEV_ORDER[minSev];
    console.log(`  ${gate ? "✗" : "·"} [${SEV_LABEL[f.sev]}] ${f.title} — ${f.count}`);
    for (const item of f.items.slice(0, 10)) console.log(`      ${item}`);
    if (f.count > 10) console.log(`      … and ${f.count - 10} more`);
  }
  const clean = r.coverage.filter((c) => !c.count).length;
  console.log(`  ${clean} of ${r.checksRun} checks found nothing`);
  if (total) console.log(`  FAIL — ${total} finding${total === 1 ? "" : "s"} at or above "${minSev}"`);
}

if (!quiet || failing) {
  console.log(`\n${failing ? `${failing} of ${results.length} file(s) have findings at or above "${minSev}"`
                            : `${results.length} file(s) clean at or above "${minSev}"`}`);
  console.log("Flagged items are prompts to look, not findings of error. Some will be deliberate.\n");
}
process.exit(failing ? 1 : 0);
