#!/usr/bin/env node
/**
 * Parity: xlsx-lint (Node) vs Workbook Inspector Pro (browser).
 *
 * The two share no code. The browser product parses OOXML with DOMParser and
 * decompresses with DecompressionStream; Node has neither combination, so
 * xlsx-lint reimplements extraction. That is a standing drift risk, and this
 * test is the control on it.
 *
 * For every fixture, both engines run and must agree on:
 *   - the count for each of the twelve checks
 *   - the EXACT location strings for each check, compared by SHA-256 digest
 *     so the assertion is byte-level rather than approximate
 *   - total formula cells and sheets read
 *
 * The digest is computed by the browser product's own sha256Hex on one side
 * and node:crypto on the other, so the comparison does not depend on either
 * implementation being trusted.
 *
 * Run:  node test/parity.test.mjs
 */

import { readFileSync, writeFileSync, mkdtempSync, existsSync } from "node:fs";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join, dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { inspect } from "../src/engine.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const PRODUCT = resolve(HERE, "..", "..", "workbook-inspector-pro");
const SHIPPING = join(PRODUCT, "Workbook Inspector Pro.html");
const FIXTURES_DIR = resolve(HERE, "fixtures");
const FIXTURES = ["fixture-defective", "fixture-clean", "fixture-unicode", "fixture-large", "fixture-demo"];

const CHROME = [
  process.env.CHROME_PATH,
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  "/usr/bin/google-chrome", "/usr/bin/chromium",
].find((p) => p && existsSync(p));
if (!CHROME) { console.error("FAIL: no Chrome found; parity cannot be established."); process.exit(2); }
if (!existsSync(SHIPPING)) {
  console.log("\nparity: SKIPPED — the commercial product is not present in this checkout.");
  console.log("This test compares xlsx-lint against Workbook Inspector Pro, which is not open source.");
  console.log("Everything xlsx-lint ships is covered by test/cli.test.mjs, which does run here.\n");
  process.exit(0);
}

/* Fixture drift guard. The fixtures are duplicated into this repository so it
   is self-contained for outside contributors. When the product IS present,
   assert the copies are byte-identical to the originals — otherwise parity
   could pass against fixtures that have quietly diverged, which would make
   the whole comparison meaningless. */
for (const f of ["fixture-defective", "fixture-clean", "fixture-unicode", "fixture-large", "fixture-demo"]) {
  const here = readFileSync(join(FIXTURES_DIR, `${f}.xlsx`));
  const there = join(PRODUCT, "test", `${f}.xlsx`);
  if (existsSync(there) && !here.equals(readFileSync(there))) {
    console.error(`FAIL: ${f}.xlsx has drifted from the product's copy. Re-copy before trusting parity.`);
    process.exit(2);
  }
}

const bytes = {};
for (const f of FIXTURES) bytes[f] = readFileSync(join(FIXTURES_DIR, `${f}.xlsx`)).toString("base64");

/* ── browser side ────────────────────────────────────────────────────────── */
const PROBE = `
<script>
(function(){
 function go(){
  var FX = ${JSON.stringify(bytes)};
  function b64(s){ var bin=atob(s), u=new Uint8Array(bin.length);
    for(var i=0;i<bin.length;i++) u[i]=bin.charCodeAt(i); return u.buffer; }
  (async function(){
   var out = {};
   try{
    for (var k in FX){
      var r = await inspect(b64(FX[k]), k + ".xlsx");
      var per = {};
      r.coverage.forEach(function(c){
        per[c.id] = { count: c.count,
                      digest: sha256Hex(new TextEncoder().encode(c.items.join("\\n"))) };
      });
      out[k] = { checks: per, totalFormulas: r.totalFormulas, sheetsRead: r.sheetsRead };
    }
   } catch(e){ out.__error = String(e && e.message); }
   document.title = "PARITY:" + JSON.stringify(out);
  })();
 }
 if (document.readyState === "complete") go(); else window.addEventListener("load", go);
})();
<\/script>`;

const src = readFileSync(SHIPPING, "utf8");
const cut = src.lastIndexOf("</body>");            // the product's own source contains "</body>" in a template string
const dir = mkdtempSync(join(tmpdir(), "parity-"));
const page = join(dir, "parity.html");
writeFileSync(page, src.slice(0, cut) + PROBE + "\n" + src.slice(cut), "utf8");

const dom = execFileSync(CHROME, [
  "--headless=new", "--disable-gpu", "--no-sandbox",
  "--virtual-time-budget=60000", "--dump-dom", pathToFileURL(page).href,
], { encoding: "utf8", maxBuffer: 1 << 28 });

const m = dom.match(/PARITY:(\{.*?\})<\/title>/s);
if (!m) { console.error("FAIL: the browser product did not report."); process.exit(2); }
const browser = JSON.parse(m[1].replace(/&quot;/g, '"').replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">"));
if (browser.__error) { console.error("FAIL: browser side threw:", browser.__error); process.exit(2); }

/* ── node side + comparison ──────────────────────────────────────────────── */
const sha = (s) => createHash("sha256").update(Buffer.from(s, "utf8")).digest("hex");

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`  ${cond ? "PASS" : "FAIL"}  ${name}${!cond && detail ? `   [${detail}]` : ""}`);
};

console.log("\nxlsx-lint ↔ Workbook Inspector Pro — engine parity\n");

for (const f of FIXTURES) {
  console.log(`── ${f}`);
  const b = browser[f];
  if (!b) { ok(`${f}: browser produced a result`, false); continue; }
  const n = inspect(readFileSync(join(FIXTURES_DIR, `${f}.xlsx`)));

  ok(`${f}: same formula-cell count`, n.totalFormulas === b.totalFormulas,
     `node ${n.totalFormulas} vs browser ${b.totalFormulas}`);
  ok(`${f}: same sheets read`, n.sheetsRead === b.sheetsRead,
     `node ${n.sheetsRead} vs browser ${b.sheetsRead}`);

  let allMatch = true, mismatched = [];
  for (const c of n.coverage) {
    const bc = b.checks[c.id];
    const nd = sha(c.items.join("\n"));
    const same = bc && bc.count === c.count && bc.digest === nd;
    if (!same) { allMatch = false; mismatched.push(`${c.id}(node ${c.count}/browser ${bc ? bc.count : "-"})`); }
  }
  ok(`${f}: all 12 checks identical, locations byte-for-byte`, allMatch, mismatched.join(" "));
}

console.log(`\n${fail ? `PARITY FAILED — ${fail} of ${pass + fail}` : `PARITY HOLDS — ${pass} checks`}\n`);
process.exit(fail ? 1 : 0);
