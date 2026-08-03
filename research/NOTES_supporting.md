# Research Notes — Supporting Coverage + Multi-Year Backbone

## SOURCE-BASIS WARNING (critical methodology point for the workbook)

Indian steel companies do NOT define EBITDA identically. Verified examples:

| Company | FY26 company-reported EBITDA | Revenue − operating expenses (excl. other income) | Difference | Explanation |
|---|---|---|---|---|
| Shyam Metalics | Rs 2,537 cr | Rs 2,333 cr | Rs 204 cr | Company EBITDA INCLUDES other income (other income FY26 = Rs 204 cr — reconciles exactly) |
| Jindal Stainless | Rs 5,560 cr | Rs 5,561 cr | ~0 | Company EBITDA EXCLUDES other income |
| Jindal Steel | Rs 9,660 cr (reported, = 9,099 adj + 561 FX) | Rs 9,644 cr | Rs 16 cr | Broadly excludes other income |
| Tata Steel | Rs 34,848 cr | Rs 34,352 cr | Rs 496 cr | Includes certain other operating items |
| JSW Steel | Rs 29,821 cr | Rs 29,464 cr | Rs 357 cr | Includes certain other operating items |
| SAIL | Rs 13,146 cr (STANDALONE, per press release) | Rs 12,000 cr (consolidated) | n/a | Basis mismatch: standalone vs consolidated |

=> WORKBOOK DESIGN DECISION: carry TWO EBITDA columns on the EBITDA sheet:
   (1) "EBITDA — As Reported by Company" (company definition, primary source, footnoted)
   (2) "EBITDA — Standardised (Revenue from operations less operating expenses; excl. other income,
        finance cost, D&A, exceptionals)" — comparable across the peer set
   Plus an explicit "EBITDA Definition / Other income included?" metadata column.
   This is exactly the kind of thing a Tier 1 comp sheet must not fudge.

---

## MULTI-YEAR BACKBONE (consolidated, Rs crore)
Source: screener.in company pages (SECONDARY aggregator of NSE/BSE filings), accessed 03-Aug-2026.
"Sales" = Revenue from operations. "Operating Profit" = revenue less operating expenses, EXCLUDING other
income => this is the STANDARDISED EBITDA basis.
Annual columns map Mar-2015 ... Mar-2026; where the row carried one extra value it is a trailing-twelve-month
(TTM) column and has been discarded.
Every FY26 figure below has been cross-checked against a primary or news-of-filing source where noted.

### CORE

TATA STEEL (TATASTEEL) — consolidated
| FY | Sales | Std EBITDA | OPM% | Other Inc | Interest | Depn | PBT | Net Profit | EPS |
|---|---|---|---|---|---|---|---|---|---|
| FY22 | 243,959 | 63,490 | 26% | 1,300 | 5,462 | 9,101 | 50,227 | 41,749 | 32.88 |
| FY23 | 243,353 | 32,300 | 13% | 1,569 | 6,299 | 9,335 | 18,235 | 8,075 | 7.17 |
| FY24 | 229,171 | 22,248 | 10% | -6,005 | 7,508 | 9,882 | -1,147 | -4,910 | -3.55 |
| FY25 | 218,543 | 25,298 | 12% | 877 | 7,341 | 10,421 | 8,413 | 3,174 | 2.74 |
| FY26 | 232,140 | 34,352 | 15% | 738 | 7,167 | 11,955 | 15,969 | 10,886 | 8.65 |
Balance sheet: Borrowings FY22 75,561 / FY23 84,893 / FY24 87,082 / FY25 94,801 / FY26 92,382
Total assets FY26 296,515; Equity capital 1,247; Reserves 100,920
VERIFIED FY26 vs primary press release: Sales 232,140 ✓ ; Net profit 10,886 ✓

JSW STEEL (JSWSTEEL) — consolidated
| FY | Sales | Std EBITDA | OPM% | Other Inc | Interest | Depn | PBT | Net Profit | EPS |
|---|---|---|---|---|---|---|---|---|---|
| FY22 | 146,371 | 39,114 | 27% | 1,600 | 4,968 | 6,001 | 29,745 | 20,938 | 85.49 |
| FY23 | 165,960 | 18,470 | 11% | 1,561 | 6,902 | 7,474 | 5,655 | 4,139 | 17.14 |
| FY24 | 175,006 | 28,157 | 16% | 1,500 | 8,105 | 8,172 | 13,380 | 8,973 | 36.03 |
| FY25 | 168,824 | 22,725 | 13% | 73 | 8,412 | 9,309 | 5,077 | 3,491 | 14.33 |
| FY26 | 185,470 | 29,464 | 16% | 18,489 | 9,102 | 9,601 | 29,250 | 25,508 | 91.26 |
(FY26 "other income" of 18,489 in the aggregator includes the Rs 18,051 cr BPSL slump-sale exceptional gain)
Balance sheet: Borrowings FY22 72,237 / FY23 80,853 / FY24 87,984 / FY25 98,752 / FY26 99,310
Total assets FY26 269,658; Equity capital 305; Reserves 99,748
VERIFIED FY26 vs primary presentation: Sales 185,470 ✓ ; PAT 25,508 ✓ ; Interest 9,102 ✓ ; Depn 9,601 ✓

SAIL — consolidated
| FY | Sales | Std EBITDA | OPM% | Other Inc | Interest | Depn | PBT | Net Profit | EPS |
|---|---|---|---|---|---|---|---|---|---|
| FY22 | 103,477 | 21,363 | 21% | 902 | 1,698 | 4,275 | 16,292 | 12,243 | 29.64 |
| FY23 | 104,448 | 8,038 | 8% | 1,856 | 2,037 | 4,964 | 2,892 | 2,177 | 5.27 |
| FY24 | 105,378 | 11,149 | 11% | 665 | 2,474 | 5,278 | 4,062 | 3,067 | 7.42 |
| FY25 | 102,479 | 10,690 | 10% | 1,006 | 2,793 | 5,651 | 3,252 | 2,372 | 5.74 |
| FY26 | 110,811 | 12,000 | 11% | 651 | 2,158 | 5,988 | 4,506 | 3,373 | 8.17 |
Balance sheet: Borrowings FY22 17,284 / FY23 30,773 / FY24 36,323 / FY25 36,934 / FY26 31,928
Total assets FY26 135,896; Equity capital 4,131; Reserves 56,225
VERIFIED FY26 vs primary filing: Sales 110,811 ✓ ; Net profit 3,373 ✓ ; Interest 2,158 ✓ ; Depn 5,988 ✓
NOTE: aggregator Std EBITDA 12,000 vs my direct derivation from the filing 12,894 (PBT before exceptionals
& JV share + finance + depn). The 894 gap equals other income (651) plus classification of certain items.
For the workbook use the DIRECT DERIVATION from the audited filing = 12,894 for standardised consolidated
EBITDA, and 13,146 for company-reported STANDALONE EBITDA. Disclose both bases.

JINDAL STEEL (JINDALSTEL) — consolidated
| FY | Sales | Std EBITDA | OPM% | Other Inc | Interest | Depn | PBT | Net Profit | EPS |
|---|---|---|---|---|---|---|---|---|---|
| FY22 | 51,166 | 15,559 | 30% | -1,884 | 1,888 | 2,097 | 9,690 | 6,766 | 56.40 |
| FY23 | 53,212 | 9,942 | 19% | -539 | 1,446 | 2,691 | 5,266 | 3,974 | 31.11 |
| FY24 | 50,354 | 10,202 | 20% | 156 | 1,294 | 2,822 | 6,241 | 5,943 | 58.21 |
| FY25 | 50,129 | 9,488 | 19% | -1,065 | 1,312 | 2,768 | 4,344 | 2,846 | 27.57 |
| FY26 | 53,225 | 9,644 | 18% | -543 | 1,517 | 3,171 | 4,413 | 3,361 | 33.01 |
Balance sheet: Borrowings FY22 13,502 / FY23 13,046 / FY24 16,472 / FY25 18,406 / FY26 22,610
Total assets FY26 97,762; Equity capital 102; Reserves 50,797
CONFLICT D2 RESOLVED: FY26 consolidated revenue from operations = Rs 53,225 cr.
Verification: sum of FY26 quarterly revenue = 12,294 + 11,686 + 13,027 + 16,218 = 53,225 ✓
and the Mar-26 quarter of 16,218 independently matches Business Standard (Rs 16,217.93 cr) and
Economic Times (Rs 16,218 cr) reporting of the exchange filing. Livemint's Rs 53,553 cr and
Economic Times' Rs 54,320 cr are on wider bases (incl. other income / total income). Company IR
headline "Gross Revenue Rs 62,412 cr" includes GST and other income and is NOT peer-comparable.
Carry BOTH in the workbook, clearly labelled.

JINDAL STAINLESS (JSL) — consolidated
| FY | Sales | Std EBITDA | OPM% | Other Inc | Interest | Depn | PBT | Net Profit | EPS |
|---|---|---|---|---|---|---|---|---|---|
| FY22 | 32,733 | 5,090 | 16% | 171 | 344 | 759 | 4,159 | 3,109 | 58.59 |
| FY23 | 35,697 | 3,586 | 10% | 236 | 325 | 724 | 2,774 | 2,084 | 25.68 |
| FY24 | 38,562 | 4,511 | 12% | 515 | 554 | 879 | 3,592 | 2,693 | 32.95 |
| FY25 | 39,312 | 4,469 | 11% | 438 | 612 | 956 | 3,339 | 2,500 | 30.41 |
| FY26 | 42,955 | 5,560 | 13% | 310 | 568 | 1,060 | 4,242 | 3,185 | 38.74 |
Balance sheet: Borrowings FY22 4,007 / FY23 3,958 / FY24 6,052 / FY25 6,402 / FY26 7,460
Total assets FY26 40,704; Equity capital 165; Reserves 19,626
VERIFIED FY26 vs PRIMARY (exchange-filed 4QFY26 earnings call transcript, 05-May-2026, JSL):
Consolidated EBITDA Rs 5,560 cr (+~19% YoY) ✓ exactly matches ; Consolidated PAT Rs 3,185 cr (+~27%) ✓ ;
Revenue Rs 42,955 cr = 10,207 + 10,893 + 10,518 + 11,337 ✓ (quarterly sum ties exactly)
Deliveries FY26 2.57 Mt (+8% YoY); Q4FY26 deliveries 0.64 Mt
Q4FY26: EBITDA Rs 1,455 cr (+37% YoY, +3% QoQ); PAT Rs 834 cr (+41% YoY)
Net debt 31-Mar-26 Rs 3,040 cr; ND/EBITDA 0.55x; ND/Equity 0.15x
Derived FY26 EBITDA/t = 5,560 cr / 2.57 Mt = Rs 21,634/t
Capacity: total melt capacity 4.2 Mtpa = 3.0 Mtpa India + 1.2 Mtpa Indonesia (Indonesian SSMS
commissioned ahead of schedule in FY26). Downstream: 1.1 Mtpa HRAP + 0.17 Mtpa CRAP lines at Jajpur
upcoming; additional Rs 900 cr committed to cold rolling at Hisar & Kharagpur; CRAP capacity to
2.67 Mtpa by FY28. Sales volume target 3.5 Mtpa by FY29.
Dividend FY26: Rs 1 interim + Rs 3 final = Rs 4/share (FV Rs 2); total payout ~Rs 330 cr
FY27 guidance: volume growth 7–9%; EBITDA/t Rs 18,000–20,000
FY26 exports = 8% of total sales. Q4FY26 grade mix: 200-series ~38%, 300-series ~43%, 400-series ~19%
ESG: EcoVadis 71/100 (Bronze) in Q4FY26; 315 MW solar/wind hybrid partially commissioned
(with Oyster Renewable Energy)
MANAGEMENT FLAG (policy-relevant): MD Abhyuday Jindal called the "temporary suspension of QCO"
a concern, noting inferior imported material continuing to enter India at scale.
Subsidiary: Jindal Stainless Steelway Ltd commenced first SS fabrication facility at Patalganga
(~Rs 125 cr initial investment).

### SUPPORTING

SHYAM METALICS & ENERGY (SHYAMMETL) — consolidated
| FY | Sales | Std EBITDA | OPM% | Other Inc | Interest | Depn | PBT | Net Profit | EPS |
|---|---|---|---|---|---|---|---|---|---|
| FY22 | 10,394 | 2,601 | 25% | 59 | 23 | 272 | 2,364 | 1,724 | 67.61 |
| FY23 | 12,658 | 1,499 | 12% | 105 | 93 | 474 | 1,037 | 843 | 33.43 |
| FY24 | 13,195 | 1,570 | 12% | 159 | 133 | 656 | 940 | 1,029 | 37.07 |
| FY25 | 15,138 | 1,866 | 12% | 231 | 144 | 711 | 1,241 | 909 | 32.53 |
| FY26 | 18,552 | 2,333 | 13% | 204 | 192 | 882 | 1,462 | 1,060 | 38.34 |
Balance sheet: Borrowings FY22 543 / FY23 1,172 / FY24 597 / FY25 789 / FY26 1,005
Total assets FY26 20,061; Equity capital 278; Reserves 11,245
VERIFIED FY26 (news-of-filing, ET Manufacturing 11-May-2026 + Quartr summaries of investor update):
Revenue Rs 18,552 cr (+22% YoY vs Rs 15,158 cr) ✓ ; COMPANY-REPORTED EBITDA Rs 2,537 cr (+21% YoY) ;
PAT Rs 1,061.17 cr (+17.85% vs Rs 908.10 cr) ✓
Q4FY26: PAT Rs 319.09 cr (+45.8%); EBITDA Rs 756 cr (+33%); EBITDA margin 14.4%
NOTE: company EBITDA Rs 2,537 cr = std EBITDA 2,333 + other income 204. Reconciles exactly.
Small FY25 revenue base difference: aggregator 15,138 vs company-quoted 15,158 (Rs 20 cr) — immaterial,
likely other-operating-revenue classification. Flag Confidence = High for FY26, Medium for FY25 base.
Auditor: MSKA & Associates LLP; audit report unmodified. Board approved 11-May-2026.

GODAWARI POWER & ISPAT (GPIL) — consolidated
| FY | Sales | Std EBITDA | OPM% | Interest | Depn | PBT | Net Profit |
|---|---|---|---|---|---|---|---|
| FY22 | 5,397 | 1,868 | 35% | 20 | 105 | 1,918 | 1,467 |
| FY23 | 5,745 | 1,134 | 20% | 20 | 124 | 1,083 | 793 |
| FY24 | 5,445 | 1,328 | 24% | 60 | 141 | 1,256 | 936 |
| FY25 | 5,370 | 1,194 | 22% | 55 | 155 | 1,092 | 813 |
| FY26 | 5,381 | 1,253 | 23% | 58 | 178 | 1,098 | 802 |
Borrowings: FY22 428 / FY23 317 / FY24 52 / FY25 309 / FY26 443
FY26 quarterly revenue sum check: 1,323 + 1,308 + 1,139 + 1,610 = 5,380 vs annual 5,381 ✓ (rounding)
Confidence: Medium-High (aggregator of filings; FY26 internally consistent). Needs primary IR
verification for a final institutional sign-off.

APL APOLLO TUBES (APLAPOLLO) — consolidated. NOTE: structural steel TUBES fabricator, NOT a
crude-steel producer. Include for downstream/value-added read-across; exclude from crude-steel
capacity and EBITDA/t-of-crude-steel comparisons.
| FY | Sales | Std EBITDA | OPM% | Interest | Depn | PBT | Net Profit |
|---|---|---|---|---|---|---|---|
| FY22 | 13,063 | 946 | 7% | 44 | 109 | 832 | 619 |
| FY23 | 16,166 | 1,022 | 6% | 67 | 138 | 863 | 642 |
| FY24 | 18,119 | 1,193 | 7% | 113 | 176 | 978 | 732 |
| FY25 | 20,690 | 1,199 | 6% | 133 | 201 | 960 | 757 |
| FY26 | 23,079 | 1,802 | 8% | 125 | 231 | 1,557 | 1,203 |
Borrowings: FY22 581 / FY23 873 / FY24 1,144 / FY25 634 / FY26 498
Confidence: Medium-High (aggregator). Needs primary IR verification.

MUKAND LTD (MUKANDLTD) — consolidated. Specialty/alloy steel long products.
| FY | Sales | Std EBITDA | OPM% | Interest | Depn | PBT | Net Profit |
|---|---|---|---|---|---|---|---|
| FY22 | 4,643 | 239 | 5% | 161 | 45 | 151 | 176 |
| FY23 | 5,568 | -172 | -3.1% | 177 | 52 | 172 | 172 |
| FY24 | 5,175 | 293 | 6% | 131 | 50 | 127 | 104 |
| FY25 | 4,890 | 285 | 6% | 130 | 51 | 118 | 76 |
| FY26 | 4,890 | 172 | 3.5% | 150 | 62 | 512 | 604 |
Borrowings: FY22 2,036 / FY23 1,505 / FY24 1,489 / FY25 1,559 / FY26 1,697
FY26 shows PBT (512) and Net Profit (604) far above Std EBITDA (172) => material non-operating /
exceptional gain and/or JV-share contribution in FY26. FLAG: requires primary filing to explain.
Confidence: Medium. Do not use FY26 Mukand PAT for operating comparisons without reading the filing.
FY26 quarterly revenue sum check: 1,129 + 1,161 + 1,329 + 1,269 = 4,888 vs annual 4,890 ✓

GALLANTT ISPAT / GALLANTT (GALLANTT) — consolidated
| FY | Sales | Std EBITDA | OPM% | Interest | Depn | PBT | Net Profit |
|---|---|---|---|---|---|---|---|
| FY22 | 3,017 | 296 | 10% | 20 | 92 | 237 | 176 |
| FY23 | 4,057 | 364 | 9% | 27 | 100 | 240 | 141 |
| FY24 | 4,227 | 448 | 11% | 28 | 116 | 311 | 225 |
| FY25 | 4,293 | 694 | 16% | 22 | 120 | 568 | 401 |
| FY26 | 4,419 | 716 | 16% | 42 | 130 | 604 | 484 |
Borrowings: FY22 387 / FY23 538 / FY24 462 / FY25 378 / FY26 548
NOTE: Gallantt Ispat Ltd was amalgamated with Gallantt Metal Ltd; the surviving listed entity is
Gallantt Ispat Ltd (renamed / NSE symbol GALLANTT). Pre-merger series are not like-for-like.
Confidence: Medium. Minor inconsistency between annual (4,419) and quarterly-sum (4,438) FY26 revenue
in the aggregator — flag and verify against the filing before use.

KIRLOSKAR FERROUS INDUSTRIES (KIRLFER) — consolidated. Pig iron, castings, and (via Oliver Engineering /
ISMT) seamless tubes & alloy steel.
| FY | Sales | Std EBITDA | OPM% | Interest | Depn | PBT | Net Profit |
|---|---|---|---|---|---|---|---|
| FY22 | 3,748 | 643 | 17% | 30 | 92 | 533 | 300 |
| FY23 | 6,417 | 836 | 13% | 95 | 173 | 617 | 437 |
| FY24 | 6,146 | 857 | 14% | 120 | 239 | 453 | 298 |
| FY25 | 6,564 | 762 | 12% | 144 | 256 | 408 | 294 |
| FY26 | 6,889 | 850 | 12% | 125 | 267 | 497 | 507 |
Borrowings: FY22 1,208 / FY23 971 / FY24 1,224 / FY25 1,278 / FY26 1,036
FY26 quarterly revenue sum check: 1,698 + 1,755 + 1,618 + 1,817 = 6,888 vs annual 6,889 ✓
Confidence: Medium-High. FY26 Net Profit (507) > PBT (497) => tax credit / deferred tax write-back;
verify in filing.

### SUPPORTING — DELISTED / NO PUBLIC FINANCIALS

UTTAM GALVA STEELS LTD — **Data Not Publicly Available (FY23 onwards)**
Status established from primary/quasi-primary record:
- Admitted to corporate insolvency resolution process under IBC.
- NCLT Mumbai approved the resolution plan of AM Mining India Pvt Ltd (ArcelorMittal group)
  by order delivered 14-Oct-2022 (IBBI order record).
- AM Mining India confirmed completion of the acquisition in Nov-2022 (ETInfra, 11-Nov-2022).
- Trading in the equity shares was SUSPENDED with effect from 17-Oct-2022 and the scrip was
  DELISTED from exchange records with effect from 08-Dec-2022, the company having complied with the
  NCLT order of 14-Oct-2022 (exchange circulars).
Consequence: no post-FY22 public consolidated financial statements. Last aggregator-available annual
series ends FY22. Company is now within the ArcelorMittal group perimeter, i.e. economically inside
AM/NS India's ownership chain rather than an independent listed comparable.
Last available annual series (consolidated, Rs cr) — for historical reference only:
FY19 Sales 757, Std EBITDA -66, Net Profit -2,146
FY20 Sales 521, Std EBITDA -46, Net Profit -1,414
FY21 Sales 654, Std EBITDA -17, Net Profit -236
FY22 Sales 840, Std EBITDA -57, Net Profit -256
RECOMMENDATION for the workbook: retain the row for continuity of the mandated coverage universe,
populate FY23–FY26 as "Data Not Publicly Available", and record the delisting/resolution facts in Notes.

ELECTROSTEEL STEELS LTD (now ESL STEEL LTD) — **FY25/FY26 financials Data Not Publicly Available**
Status:
- Greenfield integrated steel plant at Siyaljori, Chandankyari, Bokaro district, Jharkhand.
- Resolved under IBC; Vedanta Ltd acquired control (initially ~90% stake; subsequently increased).
  Renamed ESL Steel Limited. Equity delisted — no quarterly exchange reporting as a listed equity.
- Commissioned capacity 1.5 Mtpa. Product range: pig iron, billets, TMT bars, wire rods,
  ductile iron pipes (source: eslsteel.com; vedantalimited.com steel business page).
- Expansion: Vedanta targeting doubling Bokaro steelmaking capacity to 3.0 Mtpa by end-CY2026,
  ahead of the timeline in the annual report (Livemint, Jan-2026 — SECONDARY, single source, treat as
  indicative not confirmed).
- Vedanta Ltd demerger effective 01-May-2026 places the steel business in "Vedanta Iron & Steel";
  group has articulated a long-run 15 Mtpa aspiration across Bokaro and Goa (SECONDARY).
- Vedanta Ltd's Q4/FY26 press release does NOT disclose an ESL steel revenue/EBITDA/production line
  item (it discloses iron ore pig iron 895 kt, up 10% YoY, and ferrochrome 101 kt for the iron-ore
  business). Therefore standalone ESL FY26 revenue/EBITDA cannot be sourced from Vedanta's release.
- ESL Steel Ltd annual reports exist (e.g. FY2024) but the FY2026 annual report was not publicly
  available as at the 03-Aug-2026 cutoff.
RECOMMENDATION: retain the row; populate financials as "Data Not Publicly Available" with the above
explanation; populate capacity and ownership from primary company/parent sources.

---

## ENTITIES DELIBERATELY EXCLUDED FROM THE 15-NAME UNIVERSE (but relevant context)
- NMDC Steel Ltd (NSL), Nagarnar — 3 Mtpa greenfield CPSE, FY26 crude steel 2.382 Mt, HRC 2.325 Mt
  (JPC CPSE data). Not in the mandated coverage list but material to India supply; captured in the
  Production Volume sheet as an industry participant.
- NMDC Ltd, KIOCL, MOIL — raw-material CPSEs; captured in raw-material sheets, not as steel comps.
