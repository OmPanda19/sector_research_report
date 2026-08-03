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

**Deliverable:** `Industry Financial Model.xlsx` — 30 worksheets, 1,750 live formulas, 13 Excel
Tables, 0 merged cells. Consumes `Master Industry Database.xlsx`; contradicts nothing in it.

**Horizon: FY2027E–FY2033E (7 years). Base year FY2026A.**

```bash
pip install openpyxl
python3 build_financial_model.py     # writes ./Industry Financial Model.xlsx
python3 qc_financial_model.py        # 4-pass audit incl. reconciliation to the database
```

## Why 7 years

Scored on five weighted criteria (live formulas on sheet `0H`, so a reviewer can change a weight
and see whether the conclusion survives). 7-year scores **4.85** against 5-year 2.55 and 10-year 2.75.

The decisive point: a 5-year horizon ends in FY2031, the same year the announced capacity pipeline
finishes commissioning — so its terminal year embeds artificially depressed utilisation and would
understate terminal value. 7 years adds the two-year ramp tail (empirically supported: NMDC Steel's
Nagarnar reached only ~79% of nameplate in its third year), and in doing so spans ~one full industry
cycle, making the forecast average cycle-neutral. 10 years buys nothing: steady state is reached by
FY2033, no dated project extends past FY2031, no institutional source publishes an Indian GDP path
that far out, and years 8–10 carry <12% of present value at a ~12% WACC.

## The design decision that matters most

**A lagged price–utilisation feedback (driver D23).** Effective realisation = driver realisation ×
(1 + elasticity × (prior-year utilisation − normal utilisation)). Because it is *lagged* there is no
circular reference, but it means capacity, delivery rates and secondary-sector additions genuinely
transmit into price and therefore EBITDA.

Without it, capacity is decorative — and I verified that: before adding it, the pipeline delivery
factor and secondary additions drivers showed **literally zero** EBITDA sensitivity. It also produces
the economically correct, counter-intuitive result that **faster capacity delivery is
EBITDA-negative**, because it depresses utilisation and therefore price.

## Independent validation

Built bottom-up from a dated 20-project tracker plus a secondary-sector residual, the model produces
India crude steel capacity of **302.0 Mtpa in FY2031 against the National Steel Policy 2017 target
of 300 Mtpa** — within 0.7%. It was not calibrated to that target.

## Scenarios are calibrated, not invented

| Scenario | FY2033E margin | Calibrated against |
|---|---|---|
| Bull | 25.2% (peaks 26.0% in FY2030) | FY2022 observed cycle peak: Tata 26%, JSW 27%, JSPL 30% |
| Base | 16.8% | FY2026 actual 17.9% — base assumes mild erosion, not expansion |
| Bear | 8.0% | Below the FY2024 observed trough of ~13% |
| Stress | 9.9% (troughs −9.2% in FY2028) | Worst observation in the database: SAIL FY2016, −7% |

Three defects were found and fixed during sanity-checking, and they are worth noting because they
are the errors this kind of model usually ships with:

1. **Bear and stress originally produced multi-year negative industry EBITDA.** Cause: I had flexed
   prices down *and* raw materials up. That is internally contradictory — in a demand-led downturn
   iron ore and coking coal *fall*. Bear was rebuilt as a genuine demand-led downturn.
2. **Stress now carries *higher* nominal rupee realisation than bear**, despite worse demand,
   because a rupee collapse to 113/US$ lifts the rupee landed cost of imports ~28%. The FY2022
   evidence is decisive here: a coking coal spike to US$354/t coincided with *record* margins.
3. **Bull originally hit a 28.8% industry margin** — above any observed peak. Retuned to 26.0%.

## Calibration anchors (all derived from the database, not assumed)

| Anchor | Value | Derivation |
|---|---|---|
| Blended realisation | ₹59,974/t | ₹4,79,191 cr ÷ 79.90 Mt for Tata India, JSW India, SAIL, JSPL (49.4% of India output) |
| Weighted EBITDA/t | ₹10,733/t | Same four; ties to reported EBITDA sum within 0.6% |
| Implied cash cost/t | ₹49,241/t | Realisation less EBITDA/t; dispersion across the four is <8% |
| Demand elasticity | 1.10x | Consumption CAGR 7.13% (FY15–FY26) ÷ real GDP CAGR ~6.5% |
| FY2026 avg USD/INR | 88.4 | Tata Steel's dual-currency disclosure: ₹10,900/US$124, ₹15,213/US$172, ₹2.32trn/US$26bn |
| D&A | 5.28% of revenue | ₹30,715 cr ÷ ₹5,81,646 cr, four majors; tight 5.15–5.96% range |
| Normal utilisation | 80% | FY2024 actual 80.4% (vs FY2025 76.0%, FY2026 76.4%) |

Cost is calibrated **top-down**, not built bottom-up from consumption coefficients — those could not
be sourced from any primary document, and inventing tonnes-of-ore-per-tonne-of-steel would have put
false precision into the single most important cost line.

## What must be fixed before transaction use

Four inputs are unsourced and flagged Low confidence; two are left as user inputs. **None is
fabricated.** Sheet `22` section E quantifies exactly how much of the answer rests on each:

| Gap | ID | Verdict |
|---|---|---|
| Price–utilisation elasticity | D23 | **Most important unsourced input** — sets whether capacity matters at all |
| Raw material share of cash cost | D13 | Material, second-order; direction depends on relative escalation |
| Net working capital days | D18 | Material for cash flow and credit; immaterial for EBITDA |
| Equity risk premium and beta | R08 | Material for valuation only; ~11% of EV per 120bps of WACC |
| Share prices | sheet 23 | User input — all multiples compute live once entered |
| Coking coal series | D11 | Highest-priority data gap; FY2026 "anchor" is a Mar-2026 *spot* point |

Unsourced parameters are held **constant across scenarios** — flexing an unsourced input to
manufacture a scenario would be dishonest.

## Architecture

Inputs (`00`, `02`, `07`) → calculations (`03`–`20`) → outputs (`21`–`24`) → checks (`25`) →
documentation (`0H`, `26`). No assumption is hardcoded in a formula. Column alignment is identical
on every forecast sheet — **D = FY2026A, E:K = FY2027E:FY2033E** — so any cell can reference the same
column on any other sheet.

**Scenario switching:** one validated dropdown on `01 Control Panel` → named range `ScenID` →
every driver on `02` resolves via `INDEX` across its four scenario rows. No sheet contains a
scenario-specific hardcode.

**On the database boundary:** values are imported as static numbers with full provenance rather than
live external links. Live cross-workbook links break on move or rename, prompt on open, and return
stale values silently. Each row records its exact database location, the refresh formula is supplied
as text, and every value was extracted *programmatically* from the database file at build time — so a
transcription error is structurally impossible.

## Verification performed

- **Structure:** 30 sheets, 13 uniquely-named tables, 0 merged cells, gridlines off, freeze panes set.
- **Formula integrity:** all 1,750 formulas parsed; every cross-sheet reference resolves to a real
  sheet and a populated cell; 0 unguarded blank references; 0 error tokens.
- **Reconciliation to the database:** crude steel production, capacity, consumption and finished
  production all tie exactly — the model does not contradict the database.
- **Independent recomputation:** supply identity closes to 0.000 Mt across all years and scenarios;
  FY2026 base identity closes to 0.002 Mt; EBITDA unit bridge exact; utilisation ≤ 100%.
- **In-workbook:** 30 live audit checks on sheet `25`, including a **regression test** (A27) that
  compares 11 live FY2033 outputs against values computed independently at build time by
  `fmodel/chain.py`.

`qc_financial_model.py` reports **37 checks OK, 0 warnings, 0 failures.**

No Excel or LibreOffice was available in the sandbox, so formulas were verified by reference
resolution and independent recomputation rather than by opening the file — worth a spot-check on open.

## Known limitations (standing disclosure, sheet 25)

1. **No endogenous supply response** — the model does not idle capacity when margins turn negative,
   which in reality would arrest a decline. Read the stress case as a *lower bound* on margin.
2. Price feedback uses prior-year utilisation, which slightly dampens the cycle.
3. Coking coal has no fiscal-year average (database Conflict C01).
4. Working capital is not balance-sheet-derived.
5. No cost curves below the four majors; the secondary sector (~half of output) has no cost
   representation.
6. Share prices not sourced, so no trading multiple is asserted.
7. Invested capital is a replacement-cost proxy, so ROIC is indicative.
8. India-only scope — Tata Steel's Netherlands going-concern uncertainty is outside this model.

### Additional repository files

```
build_financial_model.py     entry point (30 sheets)
qc_financial_model.py        4-pass audit incl. database reconciliation
fmodel/assumptions.py        26 drivers x 4 scenarios, horizon analysis, 20-project tracker, register
fmodel/chain.py              Python mirror of the Excel chain (scenario table, tornado, tie-out)
fmodel/companies.py          company FY2026 actuals with explicit basis discipline
```
