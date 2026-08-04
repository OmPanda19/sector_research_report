# Master Industry Database — Indian Steel Industry

Institutional database for the Indian steel industry. **Data cutoff: 03 August 2026.**
Fiscal-year convention: **FY2026 = 1 April 2025 to 31 March 2026.**

**Deliverable:** `Master Industry Database.xlsx` — 18 worksheets (the 14 mandated, plus Cover,
Contents, Data Dictionary and Conflicts Log), 31 Excel Tables, 15 companies.

---

## Contents of the workbook

| # | Sheet | Content |
|---|---|---|
| — | Cover | Purpose, coverage, basis of preparation, data-integrity standard, known limitations |
| — | Contents | Hyperlinked navigation index |
| 1 | Company List | Coverage-universe control record: identifiers, ownership, route, assets, consolidation basis, latest reported period |
| 2 | Production Capacity | India capacity FY2024–FY2026 vs the 300 Mtpa NSP-2017 target; company capacity and pipeline with approval status; SAIL plant-wise capacity and utilisation |
| 3 | Production Volume | India crude/finished steel, DRI, pig iron; world and China comparatives; company production and sales volumes |
| 4 | Revenue | Consolidated revenue from operations FY2022–FY2026, plus issuer headline measures where they differ; live CAGR formula |
| 5 | EBITDA | Two bases — as reported and standardised — with other-income treatment documented per company; live margin formula |
| 6 | EBITDA per Ton | Annual and quarterly, with the volume denominator stated on every row; FY2027 guidance where given |
| 7 | Steel Prices | Domestic and international HRC/rebar with assessment basis and 52-week ranges; FY2026 price path around the safeguard duty |
| 8 | Iron Ore Prices | 62% Fe CFR China: FY2016–FY2026 fiscal-year averages, calendar averages, monthly, quarterly; NMDC notified ex-mine prices |
| 9 | Coking Coal Prices | Premium HCC FOB Australia from both public series, presented separately; thermal coal reference explicitly segregated |
| 10 | Demand & Consumption | Twelve fiscal years of apparent consumption, the full supply-demand balance, per-capita intensity, latest quarter |
| 11 | Imports & Exports | Twelve fiscal years by volume and value; product mix; country shares; monthly net-trade turning points |
| 12 | Government Policies | 20 dated and scoped measures with assessed sector impact |
| 13 | Industry KPIs | 24 dashboard-ready indicators with stable KPI IDs |
| 14 | Sources | 47 sources with publisher, date, page, URL and explicit verification status |
| 15 | Data Dictionary | Every column defined, with the reason it exists |
| 16 | Conflicts Log | 16 conflicts and definitional traps: competing figures, cause, resolution, residual risk |

---

## Design decisions that matter

**Two EBITDA columns, not one.** The peer set does not define EBITDA consistently. Shyam Metalics
includes other income (₹2,333 cr + ₹204 cr = the reported ₹2,537 cr, reconciling exactly); Jindal
Stainless excludes it (₹5,560 cr reported vs ₹5,561 cr standardised); SAIL headlines EBITDA on a
**standalone** basis while filing consolidated. A single EBITDA column would have silently mixed
three definitions. Every row states which.

**Issuer headline vs statutory measure, both carried.** Jindal Steel headlines "Gross Revenue" of
₹62,412 cr, which includes GST *and* other income. The comparable figure is ₹53,225 cr, verified by
summing the four reported quarters. Both are in the workbook, labelled.

**No estimation.** Where a figure is not public it says `Data Not Publicly Available` with the
reason. Nothing is interpolated. Calculated figures are marked `Computed` or `Derived`; figures
lifted from a document are marked `As reported`.

**No coking-coal fiscal-year average is asserted.** The Ministry of Steel and broker series
disagree by up to US$40/t, and the broker's 52-week high (US$217/t) sits *below* the Ministry's
Feb-2026 level (US$246/t) — so they are demonstrably different assessments. Averaging them would
manufacture false precision. See Conflict C01.

**Scalability.** Fiscal years run left to right so a new year is one appended column; no formula
depends on a column count. Every data range is an Excel Table, so models and Power Query bind to
auto-expanding ranges. Header rows are frozen and repeat on print. No merged cells anywhere.

---

## Reproducing the build

```bash
pip install openpyxl pypdf
python3 build_workbook.py     # writes ./Master Industry Database.xlsx
python3 qc_workbook.py        # structural + value audit
python3 qc_formulas.py        # resolves every formula reference and recomputes independently
```

### Repository layout

```
build_workbook.py        entry point
qc_workbook.py           structural audit: sheets, tables, headers, number formats, spot-checks
qc_formulas.py           formula audit: reference resolution + independent recomputation
builder/style.py         formatting engine (single source of truth for house style)
builder/sources.py       47-source registry, data dictionary, 16-entry conflicts log
builder/data.py          company list, capacity, volume, revenue, EBITDA, EBITDA/t
builder/data_market.py   prices, demand, trade, policy, KPIs
research/NOTES_*.md      research working papers with full derivations and conflict analysis
research/fetch_pdf.py    PDF retrieval + text extraction used during research
research/scrape_screener.py  multi-year backbone extraction and cross-checking
```

`research/raw/` (cached source documents) is gitignored: those files are third-party copyright and
every one is fully cited on the Sources sheet, so the research is reproducible without
redistributing them.

---

## Verification performed

- Structural audit: 18 sheets, 31 uniquely-named tables, 0 merged cells, gridlines off and freeze
  panes set on every sheet, no numeric cell left on `General` format, no duplicate or blank table
  headers.
- Value spot-checks: 22 headline figures re-checked against the primary filings, all tie.
- Formula audit: all 52 formulas resolve to numeric or blank operands; the 9 that reference
  `Data Not Publicly Available` cells are `ISNUMBER`-guarded inside `IFERROR` and correctly return
  `""` rather than erroring. All four formula families recomputed independently in Python.

## Known limitations

1. No fiscal-year coking coal average (see above). A licensed PRA series — e.g. Platts PLV HCC FOB
   Australia — is required before coking coal enters a cost model.
2. Sales volumes were not sourced for six of the eight Supporting Coverage names, so EBITDA per
   tonne is not computed for them. None has been inferred.
3. RINL and ESL Steel (Electrosteel) have no publicly available FY2026 statutory financials.
4. Company capacity is not sourced for several Supporting names.
5. BSE scrip codes, ISINs and Bloomberg/Refinitiv identifiers are deliberately **not asserted** —
   they were not independently verified in this cycle. NSE symbols were verified and are included.
6. `Data Not Publicly Available` in numeric columns makes those columns mixed-type. Safe in Excel;
   in Power Query replace the sentinel with null *before* setting the column type to decimal.

Prepared from public sources only. Contains no non-public or price-sensitive information.


---


# Industry Financial Model — Indian Steel Industry

**Deliverable:** `Industry Financial Model.xlsx` — 8 worksheets, 1,138 live formulas, **built to be copy-pasted**.

**Horizon: FY2027E–FY2033E (7 years). Base year FY2026A.**

```bash
pip install openpyxl
python3 build_model.py     # writes ./Industry Financial Model.xlsx
python3 qc_model.py        # structural + formula-evaluation audit
```

## v1.0 was broken. What went wrong and what changed

The first version of this file would not open cleanly. Two defects, both mine:

1. **Eight Excel Tables had a blank header cell.** Excel treats that as invalid content and triggers
   the "we found a problem with some content" repair prompt.
2. **199 cells were plain text instead of formulas.** A helper that built driver links returned
   `'02 Model Assumptions'!E$122` without a leading `=`, so a large part of the model never
   calculated at all.

Separately — and more fundamentally — the architecture was wrong for the way the file is used.
v1.0 leant on cross-sheet formulas and named ranges. **Neither survives a copy-paste into another
workbook:** cross-sheet references become external links or `#REF!`, and named ranges become
`#NAME?`.

### v2.0 design rules

| Rule | Why |
|---|---|
| **Zero cross-sheet formulas** | A pasted sheet cannot reference sheets that didn't come with it |
| **Zero defined names** | Named ranges don't travel with a pasted sheet |
| **Zero Excel Tables** | They caused the v1.0 repair prompt and complicate pasting |
| **Zero merged cells** | Break sorting, filtering and paste alignment |
| Scenario selector is an **integer 1–4**, not a text label | No text lookup can fail after pasting |
| Only paste-safe functions | `INDEX, CHOOSE, IF, IFERROR, ISNUMBER, ISBLANK, MIN, MAX, SUM, SUMIFS, SUMPRODUCT, ABS, COUNT, COUNTIF, MEDIAN` |

The consequence — and it is a deliberate trade-off — is that **the whole calculation chain lives on
one sheet** (`Model`), as 20 labelled sections in column A. That is the only way a pasted sheet
keeps working. All 26 originally-specified worksheet topics are present as those sections; the
mapping is on `Read Me First`.

**Paste into a blank sheet starting at cell A1.** A few formulas use an absolute reference to the
scenario cell, so pasting at A1 guarantees identical behaviour.

## Sheets

| Sheet | Formulas? | Content |
|---|---|---|
| `Read Me First` | none | How to use it, what changed, limitations, section map |
| `Model` | 1,138, all intra-sheet | S0 Control → S1 Driver matrix → S2 Active drivers → S3 Pipeline → S4 Macro → S5 Demand → S6 Capacity → S7 Supply → S8 Utilisation → S9 Price → S10 Raw materials → S11 Cost curve → S12 Revenue/EBITDA/margin → S13 Working capital & cash flow → S14 Capital allocation → S15 Cycle & trade → S16 ESG → S17 Company build → S18 DCF → S19 Audit checks |
| `Scenario Comparison` | none (values) | All four scenarios side by side, FY2033E |
| `Sensitivity` | intra-sheet | Tornado + **two live two-variable grids** + unsourced-assumption materiality |
| `Comparable Valuation` | intra-sheet | Share price = user input; all multiples compute live |
| `Assumption Register` | none | The 13 mandated fields per assumption |
| `Forecast Horizon` | intra-sheet | Weighted 5/7/10-year evaluation with live scoring |
| `Sources` | none | Forward-looking assumption sources + change log |

## Verification — this time by actually computing the formulas

Structural checks alone did not catch the v1.0 defects, so `qc_model.py` now contains a
**recursive-descent Excel formula evaluator** (lazy `IF`/`IFERROR`, array broadcasting for the
`SUMPRODUCT(ABS(range-range))` audit formulas). It evaluates the sheet and compares against
`fmodel/chain.py`:

```
Cross-sheet references            : 0 OK
Defined names                     : 0 OK
Excel Tables                      : 0 OK
Merged cells                      : 0 OK
Text-that-should-be-formula       : 0 OK

Base Case     15 line items x 7 years = 105 cells   OK
Bull Case     15 line items x 7 years = 105 cells   OK
Bear Case     15 line items x 7 years = 105 cells   OK
Stress Case   15 line items x 7 years = 105 cells   OK

In-workbook audit checks: 26 evaluated, 25 PASS, 1 WARN, 0 FAIL, 0 ERROR
Database reconciliation : 4/4 tie exactly
Zip container intact, round-trips through openpyxl
QC PASSED - 14 checks OK, 0 warnings, 0 failures
```

**420 of 420 evaluated cells match** the independent Python implementation across all four scenarios.

The evaluator also found a **real bug**: `chain.py` had the FY2026 opening working capital flexing
with the scenario's days driver. A historical base-year balance must not do that, or FY2027 embeds
an artificial step change. The sheet was right; the mirror was wrong. Fixed.

## The one warning, and it matters

`K21` warns that **terminal value is 83.4% of enterprise value** in the Base Case. That is a real
finding, not noise: the industry DCF is dominated by the exit multiple — which is **unsourced**
(6.0x). In Bear it is 238% and in Stress −123%, because explicit-period free cash flow is negative,
at which point the DCF construct breaks down entirely.

**Read the industry DCF as a sanity frame, not a valuation.** Re-base the exit multiple on the
observed median on `Comparable Valuation` once you enter share prices.

## Design decisions carried over from v1.0

**The lagged price–utilisation feedback (D23)** — effective realisation = driver × (1 + elasticity ×
(prior-year utilisation − normal utilisation)). Lagged, so no circularity, but capacity genuinely
transmits into earnings. Before adding it, the capacity drivers showed **literally zero** EBITDA
sensitivity. It also gives the economically correct, counter-intuitive result that **faster capacity
delivery is EBITDA-negative**.

**Independent validation** — built bottom-up from a 20-project tracker plus a secondary-sector
residual, the model produces **302.0 Mtpa of FY2031 capacity vs the National Steel Policy target of
300 Mtpa** (within 0.7%), without being calibrated to it.

**Scenarios calibrated to observed margin envelopes**, not set independently:

| Scenario | FY2033E margin | FY2033E EBITDA (₹ cr) | Calibrated against |
|---|---|---|---|
| Bull | 25.2% | 535,067 | FY2022 observed peak (Tata 26%, JSW 27%, JSPL 30%) |
| Base | 16.8% | 315,788 | FY2026 actual 17.9% — mild erosion, not expansion |
| Bear | 8.0% | 122,426 | Below the FY2024 observed trough of ~13% |
| Stress | 9.9% (troughs −9.2% FY2028) | 157,106 | Worst observation: SAIL FY2016, −7% |

Stress carries *higher* nominal rupee realisation than Bear despite worse demand, because a rupee
collapse to 113/US$ lifts import parity ~28%. The FY2022 evidence is decisive: coking coal at
US$354/t coincided with *record* margins.

## What must be fixed before transaction use

Unchanged from v1.0. Four inputs unsourced, two left as user inputs, **none fabricated**. Unsourced
parameters are held **constant across scenarios** — flexing one to manufacture a scenario would be
dishonest.

| Gap | ID | Verdict |
|---|---|---|
| Price–utilisation elasticity | D23 | **Most important unsourced input** — decides whether capacity matters at all |
| Raw material share of cash cost | D13 | Material, second-order; direction depends on relative escalation |
| Net working capital days | D18 | Material for cash flow and credit; immaterial for EBITDA |
| Equity risk premium and beta | WACC | Valuation only; ~11% of EV per 120bps |
| Share prices | Comps sheet | User input — multiples compute live once entered |
| Coking coal series | D11 | Highest-priority data gap; FY2026 anchor is a Mar-2026 *spot* point |

### Repository files

```
build_model.py           entry point (8 sheets, copy-paste safe)
qc_model.py              structural audit + Excel formula evaluator
fmodel/assumptions.py    26 drivers x 4 scenarios, horizon analysis, 20-project tracker, register
fmodel/chain.py          Python mirror of the chain (scenario table, tornado, verification)
fmodel/companies.py      company FY2026 actuals with explicit basis discipline
```
