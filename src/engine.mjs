/**
 * xlsx-lint engine — Node, zero dependencies.
 *
 * Reads an .xlsx workbook and runs the same twelve structural and
 * formula-pattern checks as Workbook Inspector.
 *
 * ── On the duplicated implementation ──────────────────────────────────────
 * The browser product uses DOMParser and DecompressionStream. Node has the
 * second but not the first, so the OOXML extraction here is a separate
 * implementation, which means it CAN drift from the shipped product.
 *
 * That risk is not left to discipline. `test/parity.test.mjs` runs the real
 * browser product in headless Chrome and this engine over the same five
 * fixtures and asserts identical findings, check by check. If the two ever
 * disagree, the build fails.
 *
 * The shape-normalisation function is copied verbatim rather than
 * reimplemented, because it is the load-bearing logic and a paraphrase of it
 * would be a different product.
 */

import { inflateRawSync } from "node:zlib";

/* ── ZIP ──────────────────────────────────────────────────────────────────
   Minimal central-directory reader. inflateRawSync rather than
   DecompressionStream so the CLI stays synchronous and simple. */
export function readZip(buf) {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0 && i > buf.length - 22 - 65557; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("Not a valid .xlsx file (no ZIP directory found).");

  const count = dv.getUint16(eocd + 10, true);
  let off = dv.getUint32(eocd + 16, true);
  const entries = {};
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(off, true) !== 0x02014b50) break;
    const method   = dv.getUint16(off + 10, true);
    const compSize = dv.getUint32(off + 20, true);
    const nameLen  = dv.getUint16(off + 28, true);
    const extraLen = dv.getUint16(off + 30, true);
    const cmtLen   = dv.getUint16(off + 32, true);
    const localOff = dv.getUint32(off + 42, true);
    const name = buf.subarray(off + 46, off + 46 + nameLen).toString("utf8");
    entries[name] = { method, compSize, localOff };
    off += 46 + nameLen + extraLen + cmtLen;
  }

  const get = (name) => {
    const e = entries[name];
    if (!e) return null;
    const lNameLen  = dv.getUint16(e.localOff + 26, true);
    const lExtraLen = dv.getUint16(e.localOff + 28, true);
    const start = e.localOff + 30 + lNameLen + lExtraLen;
    const raw = buf.subarray(start, start + e.compSize);
    if (e.method === 0) return raw.toString("utf8");
    if (e.method !== 8) throw new Error("Unsupported compression in the file.");
    return inflateRawSync(raw).toString("utf8");
  };
  return { names: Object.keys(entries), get };
}

/* ── XML helpers (no DOMParser) ───────────────────────────────────────────
   OOXML from Excel is machine-generated and regular, so targeted extraction
   is adequate here. It is not a general XML parser and is not offered as one. */
const decode = (s) => s
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
  .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d))
  .replace(/&amp;/g, "&");                       // last: avoid double-decoding

const attr = (tag, name) => {
  const m = new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`).exec(tag);
  return m ? decode(m[1]) : null;
};

/* ── shape normalisation — copied verbatim from the shipped product ─────── */
export const colToNum = (c) => { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
export const splitRef = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? { col: colToNum(m[1]), row: +m[2] } : null; };

const REF_RE = /(\$?)([A-Z]{1,3})(\$?)(\d{1,7})/g;
export function shapeOf(formula, col, row) {
  return formula.replace(/"(?:[^"]|"")*"/g, '""')
                .replace(REF_RE, (m, ac, c, ar, r) => {
    const cn = colToNum(c), rn = +r;
    if (cn > 16384 || rn > 1048576) return m;
    const cpart = ac ? "C$" + cn : "C[" + (cn - col) + "]";
    const rpart = ar ? "R$" + rn : "R[" + (rn - row) + "]";
    return rpart + cpart;
  });
}

const VOLATILE = ["NOW","TODAY","RAND","RANDBETWEEN","RANDARRAY","OFFSET","INDIRECT","CELL","INFO"];
const ERRORS   = ["#REF!","#DIV/0!","#VALUE!","#N/A","#NAME?","#NULL!","#NUM!","#SPILL!","#CALC!"];

/** Same catalogue and ids as the shipped product, so parity is comparable. */
export const CHECKS = [
  { id: "calc-iterative", sev: "high", title: "Iterative calculation is switched on" },
  { id: "calc-manual",    sev: "high", title: "Automatic calculation is switched off" },
  { id: "name-ref",       sev: "high", title: "Named ranges pointing at #REF!" },
  { id: "external",       sev: "high", title: "Links to other workbooks" },
  { id: "error-cells",    sev: "high", title: "Cells currently showing an error value" },
  { id: "pattern",        sev: "high", title: "Formulas that differ from the pattern beside them" },
  { id: "hidden",         sev: "med",  title: "Hidden sheets" },
  { id: "hardcoded",      sev: "med",  title: "Numbers typed inside formulas" },
  { id: "indirect",       sev: "med",  title: "INDIRECT and OFFSET" },
  { id: "volatile",       sev: "low",  title: "Volatile functions" },
  { id: "longf",          sev: "low",  title: "Very long formulas" },
  { id: "merged",         sev: "low",  title: "Merged cells" },
];

/* ── inspection ───────────────────────────────────────────────────────── */
export function inspect(buf) {
  const zip = readZip(buf);
  const wbXml = zip.get("xl/workbook.xml");
  if (!wbXml) throw new Error("This doesn't look like an Excel .xlsx workbook.");

  const relMap = {};
  const relsXml = zip.get("xl/_rels/workbook.xml.rels");
  if (relsXml) {
    for (const tag of relsXml.match(/<Relationship\b[^>]*>/g) || []) {
      let t = (attr(tag, "Target") || "").replace(/^\/xl\//, "").replace(/^\.\//, "");
      relMap[attr(tag, "Id")] = t.startsWith("xl/") ? t : "xl/" + t;
    }
  }

  const sheets = [];
  for (const tag of wbXml.match(/<sheet\b[^>]*\/?>/g) || []) {
    const rid = attr(tag, "r:id") || attr(tag, "id");
    sheets.push({ name: attr(tag, "name"), state: attr(tag, "state") || "visible", path: relMap[rid] || null });
  }

  const hits = {};
  const put = (id, item) => { (hits[id] = hits[id] || []).push(item); };

  const calcTag = (wbXml.match(/<calcPr\b[^>]*\/?>/) || [])[0];
  if (calcTag) {
    if (attr(calcTag, "iterate") === "1") put("calc-iterative", "Workbook setting: calcPr iterate = 1");
    const mode = attr(calcTag, "calcMode");
    if (mode && mode !== "auto") put("calc-manual", "Workbook setting: calcMode = " + mode);
  }
  for (const s of sheets) if (s.state !== "visible") put("hidden", `${s.name}  (${s.state})`);

  const externalNames = [];
  for (const m of wbXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const nm = attr("<x " + m[1] + ">", "name");
    const txt = decode(m[2]);
    if (ERRORS.some((e) => txt.includes(e))) put("name-ref", `${nm} → ${txt}`);
    if (/\[\d+\]/.test(txt)) externalNames.push(`${nm} → ${txt}`);
  }
  const extParts = zip.names.filter((n) => n.startsWith("xl/externalLinks/externalLink"));
  const extTargets = [];
  for (const n of zip.names.filter((n) => n.startsWith("xl/externalLinks/_rels/"))) {
    const x = zip.get(n); if (!x) continue;
    for (const tag of x.match(/<Relationship\b[^>]*>/g) || []) {
      const t = attr(tag, "Target");
      if (t) extTargets.push(decodeURIComponent(t));
    }
  }
  for (const t of (extTargets.length ? extTargets : extParts)) put("external", t);
  for (const n of externalNames) put("external", "named range → " + n);

  let totalCells = 0, totalFormulas = 0, sheetsRead = 0;
  const perSheet = [];

  for (const sh of sheets) {
    if (!sh.path) continue;
    let xml; try { xml = zip.get(sh.path); } catch { continue; }
    if (!xml) continue;
    sheetsRead++;
    let sCells = 0, sFormulas = 0;

    for (const m of xml.matchAll(/<mergeCell\b[^>]*\bref\s*=\s*"([^"]+)"/g)) put("merged", `${sh.name}!${m[1]}`);

    for (const rowM of xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
      const rowXml = rowM[1];
      const rowFormulas = [];

      for (const cm of rowXml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        totalCells++; sCells++;
        const tag = "<c " + cm[1] + ">";
        const inner = cm[2] || "";
        const ref = attr(tag, "r") || "";
        const pos = splitRef(ref);

        const vM = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner);
        if (attr(tag, "t") === "e" && vM) put("error-cells", `${sh.name}!${ref}  ${decode(vM[1])}`);

        const fM = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(inner);
        if (!fM) continue;
        totalFormulas++; sFormulas++;
        const f = decode(fM[1]); if (!f) continue;
        const F = f.toUpperCase();

        for (const v of VOLATILE) {
          if (new RegExp("(^|[^A-Z0-9_.])" + v + "\\s*\\(").test(F)) {
            put(v === "INDIRECT" || v === "OFFSET" ? "indirect" : "volatile",
                `${sh.name}!${ref}  =${f.slice(0, 90)}`);
            break;
          }
        }
        if (f.length > 250) put("longf", `${sh.name}!${ref}  (${f.length} characters)`);

        const stripped = f.replace(/"(?:[^"]|"")*"/g, '""').replace(REF_RE, " ");
        const nums = (stripped.match(/(?<![A-Z0-9_.])\d+(?:\.\d+)?/gi) || [])
                      .filter((n) => !["0","1","2","-1","100"].includes(n));
        if (nums.length) put("hardcoded", `${sh.name}!${ref}  =${f.slice(0, 90)}   → ${[...new Set(nums)].join(", ")}`);

        if (pos) rowFormulas.push({ ref, col: pos.col, row: pos.row, f, shape: shapeOf(f, pos.col, pos.row) });
      }

      rowFormulas.sort((a, b) => a.col - b.col);
      let run = [];
      const flush = () => {
        if (run.length >= 3) {
          const tally = {};
          run.forEach((x) => tally[x.shape] = (tally[x.shape] || 0) + 1);
          const shapes = Object.keys(tally);
          if (shapes.length > 1) {
            const dominant = shapes.sort((a, b) => tally[b] - tally[a])[0];
            run.filter((x) => x.shape !== dominant).forEach((x) => {
              put("pattern", `${sh.name}!${x.ref}  =${x.f.slice(0, 90)}   → differs from the ${tally[dominant]} cells beside it`);
            });
          }
        }
        run = [];
      };
      for (const cell of rowFormulas) {
        if (run.length && cell.col !== run[run.length - 1].col + 1) flush();
        run.push(cell);
      }
      flush();
    }
    perSheet.push({ name: sh.name, state: sh.state, cells: sCells, formulas: sFormulas });
  }

  const coverage = CHECKS.map((c) => ({ ...c, items: hits[c.id] || [], count: (hits[c.id] || []).length }));
  const findings = coverage.filter((c) => c.count > 0);
  const order = { high: 0, med: 1, low: 2 };
  findings.sort((a, b) => order[a.sev] - order[b.sev] || b.count - a.count);

  const counts = { high: 0, med: 0, low: 0 };
  findings.forEach((f) => counts[f.sev] += f.count);

  return { sheets, perSheet, sheetsRead, totalCells, totalFormulas, coverage, findings, counts, checksRun: CHECKS.length };
}
