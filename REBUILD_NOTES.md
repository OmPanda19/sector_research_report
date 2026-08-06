# Rebuild notes

What changed in this rebuild of `Industry Financial Model.xlsx`, and why. Read this before
the workbook.

---

## 1. The Support & Audit tables are gone

They have been replaced by a single **reference register** on `25 Audit Checks`.

- Every documented row on a model sheet now carries a short code — `FH01`, `CP07`, `ES14` —
  in a narrow **`Ref`** column appended to the right of its table.
- The code is a **hyperlink**. Click it and you land on that row's full entry in the register:
  method, exact formula, primary and secondary source, assumption, reasoning, cross-check,
  confidence and refresh frequency.
- The register opens with an index of all 31 sheets, also hyperlinked, showing each sheet's
  code range and entry count.

The `Ref` column is **appended on the right**, not inserted on the left. Inserting a column
programmatically does not repair cross-sheet references, so an insert would have broken the
model silently; appending cannot. Nothing was shifted and no formula was re-pointed.

The worked example from the brief holds exactly: on `00H Forecast Horizon`, the WEIGHTED TOTAL
row carries **`FH01`** in a new `Ref` column, with the `Ref` header on the table's header row,
and everything about FH01 lives on `Audit Checks`.

## 2. Colour conventions

| Fill / font | Meaning |
|---|---|
| Blue font | Manual input, including modelled estimates whose basis is in the register |
| Black font | Formula computed on this sheet |
| Green font | Link to another sheet |
| **Orange fill** | Sourced from an external workbook — `Master Industry Database.xlsx` |
| **Red fill** | Missing data — no source and no defensible estimate |

The orange/red pair follows the wording in the brief (*"External Workbook – yellow/orange"*,
*"Missing Data – Red"*). Both are defined **once**, as `EXT_RGB` and `MISS_RGB` at the top of
`gen/style.py`, and `tools/audit.py` imports those same two constants — so the audit can never
drift from the build. Swapping those two strings reverses the convention everywhere.

## 3. Two arithmetic defects fixed, not flagged

Both were previously disclosed and left in place. That was the wrong call.

| Cell | Was | Now |
|---|---|---|
| `Comparable Valuation` L108:L113 | `*10000` — every EV/tonne read 1,000× too high | `*10`. Tata Steel now reads **Rs 1,22,792/t**, sensibly above the Rs 55,000/t replacement cost |
| `Comparable Valuation` D135 | hard-coded `7.0x` against driver D19 base of 6.0x | `='Model Assumptions'!$K$26` — reads the driver, so the DCF and the scenario valuation finally agree |

Both are in the research block and are carried individually in the audit tool's
disclosed-edit allow-list.

## 4. Formula simplification

Formulas longer than 240 characters in the model area: **29 → 0**. Every simplification was
verified to reproduce the value it replaced.

| Where | Was | Now |
|---|---|---|
| `Control Panel` H7 | `COUNT(...)=182` | `COUNT(...)=ROWS(...)*COLUMNS(...)` — no magic number, and it survives a driver being added |
| `Control Panel` H8 | two `ISNUMBER` guards around a subtraction | guards removed — if an anchor becomes text you want to SEE the failure |
| `Control Panel` H10 | `IFERROR(...)` masking a broken link | `IFERROR` dropped, real status text still passed through |
| Everywhere | `MEDIAN(0,100,x)` as a clamp | `MIN(100,MAX(0,x))` — "not below nought, not above a hundred", in reading order |
| `Industry Cycle Model` rows 70–71 | one **550-character** formula per scenario | one cross-sheet reference. The eight indicator scores are now ordinary rows in the scenario engine |
| `Sensitivity Analysis` rows 47–51 | ten `IF(ISNUMBER(axis),axis,0)` guards | axis centres hold a real `0` with a number format that still displays "Base" |

The Cycle simplification is verified numerically: base case row 70 reproduces the live score
at **59.83** exactly.

## 5. Missing data: 1,717 → 430

A forecast model with 1,717 blank cells is a template, not a model. The rule applied
throughout:

> An estimate is always expressed as a **relationship to something sourced** — a beta to real
> GDP, a differential to a benchmark price, an intensity per tonne, a ratio to a net total —
> never as a free-floating number. Every one carries its basis into the register.

All modelled estimates live in one place, `gen/estimates.py`, so they can be reviewed as a set
rather than hunted through five fill modules.

| Sheet | Was | Now |
|---|---|---|
| ESG Model | 223 | 0 |
| Raw Material Forecast | 169 | 45 |
| Comparable Valuation | 164 | 138 |
| Macroeconomic Model | 135 | 2 |
| Steel Demand Model | 130 | 0 |
| Steel Price Forecast | 126 | 2 |
| Capital Allocation | 117 | 0 |
| Working Capital Model | 109 | 0 |
| Industry Cycle Model | 76 | 4 |
| Cash Flow Model | 74 | 6 |
| Cost Curve | 71 | 57 |
| Capacity Expansion Tracker | 60 | 60 |
| EBITDA Model | 52 | 0 |
| Steel Supply Model | 44 | 2 |
| Trade Model | 35 | 35 |
| Capacity Forecast | 34 | 34 |
| Revenue Forecast | 33 | 33 |
| Margin Analysis | 29 | 0 |
| Sensitivity Analysis | 20 | 6 |
| Scenario Manager | 10 | 0 |
| others | 11 | 6 |
| **Total** | **1,717** | **430** |

### Nothing was added to any total

Every decomposition is a **carve-out**, verified to reconcile:

- Working capital: five components as ratios to the net, signed ratios summing to exactly
  1.00, so the split reconciles to driver D18 in all eight years and rescales when the
  scenario moves the driver to 40, 52 or 62 days. Cash conversion cycle equals the driver
  exactly.
- Cost build-up: rows 38–47 of `11 Cost Curve` sum to **Rs 49,241/t**, the calibrated cash
  cost, exactly. Both share columns sum to 1.000. `13 EBITDA Model` and `14 Margin Analysis`
  read the *same* shares, so three sheets state one cost split and cannot disagree.
- Revenue less total opex still equals EBITDA at Rs 172,737 cr; margins are unchanged in all
  four scenarios (16.8 / 25.2 / 8.0 / 9.9%).
- Capital allocation cash retained is still exactly zero in every year — the ESG, technology
  and safety capex lines are memoranda carved out of spend already deducted, not additions.
- Price bridge and margin bridge both close to the rupee / basis point.

### The best result of the exercise

`10 Raw Material Forecast` now carries a **bottom-up** cost build-up from published
consumption coefficients (worldsteel: ~1.6 t ore and ~450 kg coke per tonne of pig iron) and
sourced prices. It is a **cross-check, not an input** — the model still calibrates cash cost
top-down from observed realisation less observed EBITDA.

`tools/audit.py` now runs that comparison:

```
raw material cost cross-check: bottom-up Rs 32,072/t against top-down Rs 30,814/t, gap +4.1%
```

Two estimates resting on completely different evidence, agreeing within 4%. The model could
not perform this check at all before, and the audit now warns if the gap ever exceeds 35%.

### Zeros that were previously blanks

Some cells were flagged as missing data when the true answer was **zero**, and the zero is
itself the disclosure. These now hold explicit blue zeros with the reason in the register:
debt raised and repaid, dividends, buybacks, equity issuance, M&A (this is an unlevered
aggregate that applies every rupee of free cash flow to net debt); FY2026A carbon cost (no
CBAM was payable); scrap in the raw material basket; FX in the margin bridge (the exchange
rate enters only through dollar-denominated ore and coal, so a non-zero FX line would
double-count the largest driver in the model).

## 6. What is still red, and why

430 cells. These are the cases where there is no source **and** no defensible way to model
the value. Each has an entry in the register.

- **`23 Comparable Valuation`, 138 cells** — market capitalisation and enterprise value for
  the supporting and unlisted names, and the FY2021–FY2026 historical multiples. All need a
  share price the workbook does not have. Prices *could* be typed in, but they would corrupt
  the peer median that the entire sheet is measured against, so they are not.
  Book value **is** now derived, by inverting return on equity, which resolved the P/B column,
  the valuation bridge, the implied valuation table and the scenario summary. The resulting
  multiples land where these names actually trade — SAIL 1.01x, JSW 1.96x, Tata Steel 2.42x.
- **`07 Capacity Expansion Tracker`, 60 cells** — per-project existing capacity, total capacity
  on completion, and announcement dates. Project-level disclosure that does not exist for most
  of the pipeline.
- **`11 Cost Curve`, 57 cells** — company cost curve ranks 7–15. Implied cash cost can only be
  derived for producers disclosing *both* realisation and EBITDA per tonne. That is four names.
- **`10 Raw Material Forecast`, 45 cells** — FY2021–FY2026 fiscal-year averages for coking coal
  and its derivatives. **The Master Industry Database explicitly refuses to assert these** and
  records the reason as Conflict C01: the Ministry of Steel put March-2026 at US$225/t while
  IDBI Capital put it at US$186/t. Modelling a forecast from a documented relationship is
  defensible; manufacturing history that the sources decline to provide is not.
- **`19 Trade Model` 35, `06 Capacity Forecast` 34, `12 Revenue Forecast` 33** — regional trade
  and capacity splits, and product-wise volume. Not yet reached in this pass.
- **Remainder, 24 cells** — genuinely not applicable rather than unavailable: the FY2026A
  column of the free cash flow bridge (no FY2025 comparative to bridge from), the Overheating
  cycle phase (no scenario reaches that score), FY2026A price growth (no FY2025 fiscal-year
  average realisation exists).

## 7. Verification

```
python build_ifm.py        # rebuild the workbook
python tools/audit.py      # 23 checks
python tools/blanks.py     # per-sheet inventory of red and orange cells
```

`tools/audit.py`: **23 pass, 0 warn, 0 fail**, including six widened bridge-closure checks and
the new raw-material cross-check. Every formula in the workbook evaluates without error in all
four scenarios, and the 44 independent FY2033E tie-out values are still reproduced exactly by
the scenario engine.
