#!/usr/bin/env node
/**
 * xlsx-lint CLI behaviour tests.
 *
 * Exercises the published binary the way CI would call it, because exit codes
 * are the product here: a lint that reports correctly but exits 0 on failure
 * is worse than no lint at all — it turns a red build green.
 *
 * Run:  node test/cli.test.mjs
 */

import { execFileSync } from "node:child_process";
import { readFileSync, writeFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const CLI = resolve(HERE, "..", "src", "cli.mjs");
const FX = resolve(HERE, "fixtures");

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`  ${cond ? "PASS" : "FAIL"}  ${name}${!cond && detail ? `   [${detail}]` : ""}`);
};

/** Run the CLI, capturing stdout and exit code rather than throwing. */
function run(args) {
  try {
    const out = execFileSync(process.execPath, [CLI, ...args], { encoding: "utf8" });
    return { code: 0, out };
  } catch (e) {
    return { code: e.status ?? -1, out: (e.stdout || "") + (e.stderr || "") };
  }
}

console.log("\nxlsx-lint — CLI behaviour\n");

/* ── exit codes: the load-bearing behaviour ──────────────────────────────── */
const defective = join(FX, "fixture-defective.xlsx");
const clean = join(FX, "fixture-clean.xlsx");
const demo = join(FX, "fixture-demo.xlsx");

const d = run([defective]);
ok("defective workbook exits 1", d.code === 1, `exit ${d.code}`);
ok("defective output names the pattern finding", /differ from the pattern beside them/.test(d.out));
ok("defective output shows the flagged cell", /Model!D2/.test(d.out));

const c = run([clean]);
ok("clean workbook exits 0", c.code === 0, `exit ${c.code}`);
ok("clean output says no flagged patterns", /no flagged patterns/.test(c.out));
ok("clean output refuses to imply correctness",
   /does not establish that the workbook is correct/.test(c.out));

/* Severity gating: demo has one high finding plus med and low ones. */
const dh = run([demo, "--severity", "high"]);
ok("demo at --severity high exits 1", dh.code === 1, `exit ${dh.code}`);
ok("demo reports 8 of 12 checks found nothing", /8 of 12 checks found nothing/.test(dh.out));

/* Low severity is a superset, so it must also fail. */
const dl = run([demo, "--severity", "low"]);
ok("demo at --severity low also exits 1", dl.code === 1, `exit ${dl.code}`);

/* ── json ────────────────────────────────────────────────────────────────── */
const j = run([demo, "--json"]);
ok("--json exits 1 when findings gate", j.code === 1, `exit ${j.code}`);
let parsed = null;
try { parsed = JSON.parse(j.out); } catch {}
ok("--json emits valid JSON", !!parsed);
if (parsed) {
  const f0 = parsed.files[0];
  ok("json reports the file", f0.file.includes("fixture-demo"));
  ok("json reports 12 checks run", f0.checksRun === 12);
  ok("json coverage lists all 12 checks", f0.coverage.length === 12);
  ok("json includes checks that found nothing",
     f0.coverage.filter((x) => x.count === 0).length === 8);
  ok("json findings carry severity and locations",
     f0.findings.every((x) => x.severity && Array.isArray(x.locations)));
  ok("json pattern finding has exactly one location",
     f0.findings.find((x) => x.id === "pattern")?.count === 1);
}

/* ── usage ───────────────────────────────────────────────────────────────── */
ok("--list-checks prints twelve", run(["--list-checks"]).out.trim().split("\n").length === 12);
ok("--help exits 0", run(["--help"]).code === 0);
ok("--version prints a version", /^\d+\.\d+\.\d+$/.test(run(["--version"]).out.trim()));

const missing = run([join(FX, "does-not-exist.xlsx")]);
ok("a missing file exits 2", missing.code === 2, `exit ${missing.code}`);

/* A non-xlsx input must not be silently treated as clean. */
const tmp = mkdtempSync(join(tmpdir(), "xlsxlint-"));
const notZip = join(tmp, "notes.xlsx");
writeFileSync(notZip, "this is not a zip at all");
const bad = run([notZip]);
ok("a non-zip .xlsx exits 2 rather than passing", bad.code === 2, `exit ${bad.code}`);
ok("the read error is explained", /valid \.xlsx|ZIP directory/i.test(bad.out));

/* ── directory walk ──────────────────────────────────────────────────────── */
const dirRun = run([FX, "--severity", "high"]);
ok("a directory inspects every workbook in it", (dirRun.out.match(/\.xlsx/g) || []).length >= 5);
ok("a directory with defective workbooks exits 1", dirRun.code === 1, `exit ${dirRun.code}`);

/* ── unicode ─────────────────────────────────────────────────────────────── */
const uni = run([join(FX, "fixture-unicode.xlsx")]);
ok("unicode sheet names survive to the output", /Résumé — Q1 ✓/.test(uni.out));

/* ── the privacy claim must be structurally true ─────────────────────────── */
const sources = ["src/cli.mjs", "src/engine.mjs"]
  .map((p) => readFileSync(resolve(HERE, "..", p), "utf8")).join("\n");
ok("no fetch in the shipped source", !/\bfetch\s*\(/.test(sources));
ok("no http(s) request or URL", !/https?:\/\//.test(sources.replace(/^\s*\*.*$/gm, "")));
ok("no node:http / node:https import", !/node:https?/.test(sources));
ok("no child process execution", !/child_process/.test(sources));
/* Match a tracking CALL, not the word. The first version of this check
   searched for "telemetry" and failed on the CLI's own privacy promise
   ("No network, no upload, no telemetry") — the opposite of tracking. */
ok("no telemetry call", !/(sendBeacon|gtag\s*\(|dataLayer|mixpanel\.|posthog|amplitude)/i.test(sources));
ok("zero runtime dependencies declared",
   !JSON.parse(readFileSync(resolve(HERE, "..", "package.json"), "utf8")).dependencies);

console.log(`\n${fail ? `${fail} of ${pass + fail} CHECKS FAILED` : `ALL ${pass} CHECKS PASSED`}\n`);
process.exit(fail ? 1 : 0);
