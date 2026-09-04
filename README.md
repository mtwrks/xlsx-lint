# xlsx-lint

**Find Excel formulas that break their own pattern — in CI, before anyone relies on the number.**

Zero dependencies. Runs entirely on your machine. Nothing is uploaded, there is no service, and
there is no API key.

```bash
npx xlsx-lint model.xlsx
```

---

## The problem

Someone fills a formula across a row and it stops one cell short. Or a cell gets overtyped with a
constant. The spreadsheet still opens, still calculates, still looks right — and the number is
wrong.

```
B2  =A2*$A$1
C2  =B2*$A$1
D2  =C2        ← flagged: differs from the 4 cells beside it
E2  =D2*$A$1
F2  =E2*$A$1
```

`xlsx-lint` normalises every formula so that references become relative positions. Two cells doing
the same job from different places produce an identical structure — so the one that doesn't stands
out. **Compared by structure, not by value**, so it is caught whatever the numbers happen to be.

## Install

```bash
npx xlsx-lint model.xlsx          # no install
npm i -g xlsx-lint                # or globally
```

Node 18+. No other requirements.

## Usage

```bash
xlsx-lint model.xlsx                      # one workbook
xlsx-lint reports/ --recursive            # every .xlsx in a tree
xlsx-lint model.xlsx --severity low       # fail on anything, not just the worst
xlsx-lint model.xlsx --json               # machine-readable
xlsx-lint --list-checks                   # the twelve checks
```

**Exit codes:** `0` clean at the chosen severity · `1` findings · `2` usage or read error.

## GitHub Action

```yaml
- uses: vaananenvesamatti-ship-it/xlsx-lint@v1
  with:
    path: models/
    severity: high
```

Fails the build when a workbook in `models/` has a formula that breaks its own pattern.

## The twelve checks

| Severity | Checks |
| --- | --- |
| **review first** | iterative calculation on · automatic calculation off · named ranges pointing at `#REF!` · links to other workbooks · cells showing an error value · **formulas that differ from the pattern beside them** |
| **worth checking** | hidden sheets · numbers typed inside formulas · `INDIRECT` and `OFFSET` |
| **note** | volatile functions · very long formulas · merged cells |

Every run reports **all twelve**, including the ones that found nothing — because knowing what was
checked and came back clean is most of what a check is worth.

## Your files stay on your machine

No network calls of any kind. The published source contains no `fetch`, no `http`/`https` import,
no child process, and no telemetry — and [the test suite asserts each of those](test/cli.test.mjs)
rather than asking you to take it on trust.

## What this does not do

- **It is not an audit and not a certification.** It runs a fixed set of pattern checks.
- **A clean result does not mean the workbook is correct.** It means those twelve patterns were
  absent. Defects outside them will not appear.
- **Flagged items are prompts to look, not findings of error.** Many will be deliberate — a
  subtotal, a first period, a genuine exception. You decide.
- **It does not check your logic**, only structure and formula patterns.
- **Values are not recalculated.** It reads what is stored in the file.
- `.xlsx` only. Not `.xls`, not `.xlsb`. Password-protected and macro-only workbooks cannot be read.

## Verification

The engine is shared with a browser product, and the two are held to byte-level agreement:
[`test/parity.test.mjs`](test/parity.test.mjs) runs both over five fixtures — defective, clean,
unicode, large and a realistic model — and asserts identical findings and identical location
strings for all twelve checks. If they ever diverge, the build fails.

```bash
npm test    # CLI behaviour (32 checks) + engine parity (15 checks)
```

## If you need a report rather than an exit code

`xlsx-lint` answers *pass or fail*. If you need to **hand someone evidence of what you checked** —
a dated review report with your name, the client reference, a SHA-256 fingerprint of the file, your
sign-off on each finding, and the coverage list — that is **Workbook Inspector Pro**, a one-time
purchase built on this same engine, from [Mintworks](https://mtwrks.gumroad.com).

This tool stays free and complete for what it does. Nothing is withheld to sell the other one.

## Licence

MIT.
