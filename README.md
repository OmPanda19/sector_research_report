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
