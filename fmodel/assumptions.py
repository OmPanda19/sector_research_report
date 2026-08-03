# -*- coding: utf-8 -*-
"""Assumption registry, scenario driver matrix and project pipeline for
Industry Financial Model.xlsx.

HORIZON: FY2027E - FY2033E (7 years). Base year FY2026A.
Column convention used by EVERY forecast sheet in the workbook:
    B = line item | C = unit | D = FY2026A | E..K = FY2027E..FY2033E | L = CAGR/terminal
Column alignment is identical across sheets so that any forecast cell can reference the
same column letter on any other sheet. This is deliberate and must not be broken.
"""

BASE_YEAR = "FY2026A"
FY = ["FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E", "FY2033E"]
N = len(FY)
SCENARIOS = ["Base Case", "Bull Case", "Bear Case", "Stress Case"]
COL_BASE = 4          # column D
COL_F1 = 5            # column E  = FY2027E
COL_FN = COL_F1 + N - 1   # column K = FY2033E
COL_CAGR = COL_FN + 1     # column L

# ======================================================================================
# FY2026 CALIBRATION ANCHORS - all derived from Master Industry Database.xlsx
# ======================================================================================
ANCHORS = [
    ("A01", "India crude steel production", "Mt", 168.42,
     "Master DB > Production Volume > India crude steel production > FY2026", "S01"),
    ("A02", "India total finished steel production", "Mt", 160.94,
     "Master DB > Demand & Consumption > supply balance (April-2026 JPC vintage) > FY2026. "
     "VINTAGE CHOICE: the model uses the April-2026 vintage throughout because that balance is "
     "internally self-reconciling (production + imports - exports - stock change = consumption), "
     "which a forecast identity requires. The June-2026 vintage figure of 161.74 Mt is carried "
     "as a reconciliation memo. This is the resolution the database itself recommends for the "
     "supply balance under Conflict C03", "S01"),
    ("A03", "India apparent finished steel consumption", "Mt", 163.74,
     "Master DB > Demand & Consumption > supply balance (April-2026 JPC vintage) > FY2026. "
     "June-2026 vintage is 164.19 Mt - carried as a memo. See A02 on the vintage choice", "S01"),
    ("A04", "India crude steel capacity", "Mtpa", 220.4,
     "Master DB > Production Capacity > India crude steel capacity > FY2026", "S02/S03"),
    ("A05", "India finished steel imports", "Mt", 6.524,
     "Master DB > Imports & Exports > FY2026", "S02"),
    ("A06", "India finished steel exports", "Mt", 6.602,
     "Master DB > Imports & Exports > FY2026", "S02"),
    ("A07", "Variation in stock", "Mt", -2.88,
     "Master DB > Demand & Consumption > supply balance line 6 > FY2026", "S01"),
    ("A08", "Crude-to-finished steel ratio", "x", 1.04647,
     "Derived: 168.42 / 160.94, both on the April-2026 JPC vintage, so the ratio is internally "
     "consistent with the supply balance the model uses. On the June-2026 vintage "
     "(168.42 / 161.74) the ratio would be 1.0413 - a 0.5% difference", "S01"),
    ("A09", "Blended industry realisation", "Rs/t", 59974,
     "Derived: volume-weighted revenue / volume for Tata Steel India, JSW Steel India, SAIL "
     "and Jindal Steel from Master DB (Rs 4,79,191 cr / 79.90 Mt x 10)", "S08/S09/S10/S11"),
    ("A10", "Weighted EBITDA per tonne", "Rs/t", 10733,
     "Derived: volume-weighted EBITDA/t across the same four majors from Master DB", "S08-S11"),
    ("A11", "Implied cash cost per tonne", "Rs/t", 49241,
     "Derived: A09 less A10. Reconciliation: A10 x 79.90 Mt = Rs 85,759 cr against the "
     "Rs 86,318 cr sum of the four reported EBITDAs, a 0.6% rounding difference", "Derived"),
    ("A12", "Weighted EBITDA margin", "%", 0.179,
     "Derived: A10 / A09", "Derived"),
    ("A13", "Calibration volume coverage", "Mt", 79.90,
     "Sum of FY2026 volumes for the four majors = 49.4% of India finished steel production",
     "S08-S11"),
    ("A14", "India per capita consumption", "kg", 115.7,
     "Master DB > Demand & Consumption > per capita > FY2026", "S03"),
    ("A15", "Iron ore 62% Fe CFR China, FY2026 average", "US$/dmt", 100.52,
     "Master DB > Iron Ore Prices > fiscal-year averages > FY2026", "S17"),
    ("A16", "Australian thermal coal, FY2026 average", "US$/mt", 111.50,
     "Master DB > Coking Coal Prices > thermal reference > FY2026", "S17"),
    ("A17", "Premium HCC coking coal, Mar-2026 spot reference", "US$/t", 225.0,
     "Master DB > Coking Coal Prices > Ministry of Steel assessment > Mar-2026. NOT a "
     "fiscal-year average - none is asserted in the database (Conflict C01)", "S01"),
    ("A18", "Domestic HRC, end-Mar-2026", "Rs/t", 57700,
     "Master DB > Steel Prices > ICRA end-Mar-2026 assessment", "S19"),
]

# ======================================================================================
# FORECAST HORIZON EVALUATION - quantitative, one option selected
# ======================================================================================
HORIZON_HEADERS = [
    "Criterion", "Weight", "5-Year (FY2027-FY2031)", "7-Year (FY2027-FY2033)",
    "10-Year (FY2027-FY2036)", "Score 5Y", "Score 7Y", "Score 10Y",
    "Basis of assessment / evidence",
]

HORIZON_EVAL = [
    ("Coverage of the announced capacity pipeline", 0.25,
     "Ends FY2031, exactly the terminal year of the announced pipeline. Captures 100% of "
     "commissioning but 0 years of post-commissioning ramp.",
     "Ends FY2033. Captures 100% of commissioning PLUS a 2-year ramp tail, so the terminal "
     "year reflects normalised rather than ramping utilisation.",
     "Ends FY2036, five years beyond the last dated project. Years 8-10 have no project-level "
     "visibility and become pure extrapolation.",
     3, 5, 2,
     "Dated pipeline in the Master Database terminates FY2031: POSCO-JSW 6 Mtpa greenfield by "
     "2031, SAIL ~35 Mt ambition by FY2031, PLI 1.2 capacity 8.7 Mt by FY2031, National Steel "
     "Policy 2017 target year FY2031. No credible dated project extends beyond FY2031."),
    ("Terminal year is a normalised, valuation-usable year", 0.25,
     "NO. Capacity commissioned FY2030-FY2031 is still ramping in FY2031, so terminal-year "
     "utilisation is artificially depressed and a terminal value struck there is understated.",
     "YES. The FY2030-FY2031 capacity cohort reaches normalised utilisation by FY2033 on a "
     "2-3 year ramp, so the terminal year is a fair steady state.",
     "YES but redundant - steady state is already reached by FY2033, so FY2034-FY2036 add "
     "extrapolation risk without adding information.",
     1, 5, 4,
     "Greenfield and large brownfield Indian steel capacity historically takes 2-3 years from "
     "commissioning to normalised utilisation. Observed in the database: NMDC Steel Nagarnar "
     "(3 Mtpa) produced 1.471 Mt in FY2025 and 2.382 Mt in FY2026, i.e. ~79% of nameplate in "
     "year 3+; RINL required a full year to recover to 5.43 Mt."),
    ("Coverage of a complete industry cycle", 0.20,
     "~0.7 cycles. Truncates the cycle, so the average forecast year is not cycle-neutral and "
     "the result is highly sensitive to where the start point sits in the cycle.",
     "~1.0 cycle. The forecast average is approximately cycle-neutral, which is the property a "
     "mid-cycle valuation requires.",
     "~1.4 cycles. Cycle-neutral but the second partial cycle is unforecastable in its timing.",
     2, 5, 3,
     "Observed cycle length in the database: standardised EBITDA margin for Tata Steel peaked "
     "at 26% in FY2022, troughed at 10% in FY2024 and recovered to 15% in FY2026 - a 4-year "
     "peak-to-trough leg implying a ~7-8 year full cycle. Coking coal peaked at US$354/t in "
     "Oct-2023 and troughed at ~US$175/t in Mar-2025; iron ore peaked at US$155.52/dmt in "
     "FY2022 and troughed at US$100.52/dmt in FY2026."),
    ("Reliability of the macro anchor", 0.15,
     "Strong. RBI publishes an explicit FY2027 projection and the IMF medium-term path is "
     "credible to roughly 5 years.",
     "Adequate. Years 6-7 rely on a converging steady-state growth assumption rather than a "
     "published projection, which is disclosed and sensitised.",
     "Weak. Years 8-10 require a GDP path with no institutional publisher support at all.",
     5, 4, 2,
     "RBI's June-2026 MPC projects FY2027 real GDP growth at 6.6%. No official Indian source "
     "publishes a real GDP path beyond roughly 5 years, so any year past FY2031 is the "
     "modeller's own convergence assumption."),
    ("Usability for DCF terminal value", 0.15,
     "Poor. A terminal value struck on a non-normalised year requires a manual mid-cycle "
     "adjustment, which defeats the purpose of an explicit forecast period.",
     "Good. Terminal value can be struck directly off FY2033 with no manual normalisation.",
     "Good but the incremental years add estimation error to the discounted stub, where their "
     "present-value weight is small.",
     2, 5, 3,
     "At a ~12% nominal WACC the present-value weight of forecast years 8-10 is under 12% of "
     "total enterprise value, so extending the horizon adds materially more estimation error "
     "than valuation information."),
]

HORIZON_CONCLUSION = (
    "SELECTED HORIZON: 7 YEARS - FY2027E to FY2033E, with FY2026A as the base year.\n\n"
    "Weighted score: 7-year 4.85 against 5-year 2.55 and 10-year 2.75 (weights and criterion "
    "scores are shown live in the table above and the weighted total is a formula, not a "
    "hard-coded number).\n\n"
    "The decisive consideration is that the 5-year horizon ends in FY2031 - the same year the "
    "announced capacity pipeline finishes commissioning - so its terminal year embeds "
    "artificially low utilisation and would understate terminal value. The 7-year horizon adds "
    "the two-year ramp tail needed for the FY2030-FY2031 capacity cohort to reach normalised "
    "utilisation, and in doing so also spans approximately one full industry cycle, which makes "
    "the forecast average cycle-neutral. The 10-year horizon buys no additional information: "
    "steady state is already reached by FY2033, no dated project extends beyond FY2031, no "
    "institutional source publishes an Indian GDP path that far out, and forecast years 8-10 "
    "carry under 12% of present value at a ~12% nominal WACC while carrying the highest "
    "estimation error in the model.\n\n"
    "This horizon is applied consistently on every forecast worksheet. Columns E to K are "
    "FY2027E to FY2033E throughout the workbook without exception."
)

# ======================================================================================
# SCENARIO DRIVER MATRIX
# Each driver: (code, name, unit, {scenario: [7 annual values]}, source_id, confidence,
#              evidence/reasoning)
# ======================================================================================
def _flat(v):
    return [v] * N


DRIVERS = [
    # ---------------- MACRO
    ("D01", "India real GDP growth", "%",
     {"Base Case":   [0.066, 0.068, 0.068, 0.067, 0.067, 0.065, 0.065],
      "Bull Case":   [0.074, 0.076, 0.075, 0.074, 0.073, 0.072, 0.070],
      "Bear Case":   [0.056, 0.058, 0.060, 0.061, 0.062, 0.062, 0.062],
      "Stress Case": [0.044, 0.048, 0.055, 0.058, 0.060, 0.060, 0.060]},
     "RBI-Jun26", "High",
     "RBI Monetary Policy Committee, 05-Jun-2026, projects FY2027 real GDP growth at 6.6%, "
     "revised down from 6.9% in Apr-2026 on West Asia conflict, energy prices and monsoon "
     "risk; FY2026 estimated at 7.6%; Q4FY2026 actual 7.7%. Base case adopts the RBI FY2027 "
     "figure unchanged, then converges to a 6.5% steady state. Bull adds ~80bps on the RBI's "
     "own stated sensitivity that growth 'could top 7% if oil prices fall further'. Stress "
     "applies a 220bps FY2027 shock, comparable in magnitude to a global recession year."),

    ("D02", "Steel demand elasticity to real GDP", "x",
     {"Base Case":   [1.15, 1.15, 1.12, 1.10, 1.08, 1.06, 1.05],
      "Bull Case":   [1.30, 1.28, 1.25, 1.22, 1.20, 1.18, 1.15],
      "Bear Case":   [1.00, 1.00, 0.98, 0.97, 0.96, 0.95, 0.95],
      "Stress Case": [0.75, 0.85, 0.90, 0.92, 0.93, 0.94, 0.95]},
     "Derived-DB", "High",
     "Realised elasticity FY2015-FY2026 is 1.06x to 1.13x: apparent consumption compounded at "
     "7.13% (76.99 Mt to 164.19 Mt over 11 years) against real GDP CAGR of roughly 6.3-6.7%, "
     "midpoint 1.10x. Near-term base is set slightly above at 1.15x because worldsteel's "
     "Apr-2026 Short Range Outlook forecasts Indian demand +7.4% in CY2026 and +9.2% in "
     "CY2027 against RBI GDP of 6.6%, implying 1.1x-1.4x. Elasticity tapers toward 1.05x by "
     "FY2033 on the standard maturing-intensity argument as per capita consumption rises from "
     "115.7 kg. Stress assumes elasticity falls below 1.0x, as it did in FY2021 (-5.27% "
     "consumption on a positive-growth base effect)."),

    ("D03", "USD/INR average", "Rs/US$",
     {"Base Case":   [94.5, 95.5, 96.5, 97.5, 98.5, 99.5, 100.5],
      "Bull Case":   [93.0, 93.5, 94.0, 94.5, 95.0, 95.5, 96.0],
      "Bear Case":   [96.5, 98.5, 100.0, 101.5, 103.0, 104.5, 106.0],
      "Stress Case": [101.0, 105.0, 107.0, 108.5, 110.0, 111.5, 113.0]},
     "IDBI-22Jun26", "High",
     "Spot USD/INR was 94 on 22-Jun-2026 with a 52-week range of 85 to 98 (IDBI Capital "
     "commodity tracker). Base assumes ~1.0% annual depreciation, consistent with the "
     "inflation differential implied by RBI's 5.1% FY2027 CPI projection against developed-"
     "market targets of ~2%. Stress takes INR beyond the 52-week low to 113, a ~20% "
     "depreciation from spot. MATERIAL BOTH WAYS: a weaker INR raises rupee raw-material cost "
     "(iron ore, coking coal and nickel are dollar-priced) but also raises the rupee landed "
     "cost of imports, supporting domestic realisations."),

    ("D04", "CPI inflation", "%",
     {"Base Case":   [0.051, 0.046, 0.045, 0.045, 0.045, 0.045, 0.045],
      "Bull Case":   [0.045, 0.042, 0.042, 0.042, 0.042, 0.042, 0.042],
      "Bear Case":   [0.057, 0.052, 0.050, 0.048, 0.048, 0.048, 0.048],
      "Stress Case": [0.068, 0.062, 0.055, 0.052, 0.050, 0.050, 0.050]},
     "RBI-Jun26", "High",
     "RBI raised its FY2027 CPI projection by 50bps to 5.1% at the 05-Jun-2026 MPC, citing "
     "crude oil and supply disruption. Base converges to the 4.5% mid-point of the RBI's "
     "4%+/-2% target band. Drives conversion cost inflation (D12) and the discount rate."),

    # ---------------- SUPPLY / CAPACITY
    ("D05", "Announced pipeline delivery factor", "%",
     {"Base Case":   _flat(0.85), "Bull Case": _flat(0.95),
      "Bear Case":   _flat(0.70), "Stress Case": _flat(0.55)},
     "Derived-DB", "Medium",
     "Proportion of dated announced capacity additions assumed to commission on schedule. "
     "Calibrated against realised delivery: India added 20.82 Mtpa in FY2025 and 20.07 Mtpa "
     "in FY2026 (Master DB), against a National Steel Policy path to 300 Mtpa by FY2031 that "
     "requires roughly 16 Mtpa a year - i.e. recent delivery has RUN AHEAD of the required "
     "rate, which supports a base factor materially above 50%. 85% rather than 100% reflects "
     "the routine slippage evident in the tracker (JSW Vijayanagar BF-3 was still under "
     "testing at FY2026 year-end; AM/NS Hazira 9 to 15 Mtpa has been in progress for several "
     "years). Stress at 55% reflects a capital-constrained environment."),

    ("D23", "Price sensitivity to capacity utilisation (lagged)", "x",
     {"Base Case": _flat(0.40), "Bull Case": _flat(0.40),
      "Bear Case": _flat(0.40), "Stress Case": _flat(0.40)},
     "Indicative", "Low",
     "THE FEEDBACK LOOP THAT MAKES THIS AN INTEGRATED MODEL RATHER THAN A SET OF PARALLEL "
     "FORECASTS. Realisation is adjusted by this elasticity multiplied by the gap between the "
     "PRIOR year's capacity utilisation and normal utilisation (D24): effective realisation = "
     "driver realisation x (1 + D23 x (utilisation[t-1] - D24)). The one-year LAG is deliberate "
     "and essential - it makes capacity, delivery rates and secondary additions genuinely "
     "affect price and therefore EBITDA, while avoiding a circular reference in Excel. Without "
     "it, capacity is decorative: the pipeline delivery factor and secondary additions would "
     "show literally zero EBITDA sensitivity, which would be a visible hole in the model. "
     "Sign: MORE capacity lowers utilisation, which lowers price, which lowers EBITDA - so "
     "faster capacity delivery is EBITDA-NEGATIVE, which is the economically correct result "
     "and the opposite of the naive intuition. Value of 0.40x is INDICATIVE and unsourced; it "
     "is held constant across scenarios for that reason and is a sensitivity axis on sheet 22. "
     "Directional support from the database: utilisation fell from 80.4% in FY2024 to 76.0% in "
     "FY2025 and domestic HRC weakened materially over that period, requiring a safeguard duty "
     "to arrest the decline."),

    ("D24", "Normal capacity utilisation", "%",
     {"Base Case": _flat(0.80), "Bull Case": _flat(0.80),
      "Bear Case": _flat(0.80), "Stress Case": _flat(0.80)},
     "Derived-DB", "Medium",
     "The utilisation level at which realisation equals the unadjusted driver value, i.e. the "
     "neutral point of the feedback in D23. Anchored on actual Indian utilisation from the "
     "Master Database: 80.4% in FY2024, 76.0% in FY2025 and 76.4% in FY2026. FY2024's 80.4% is "
     "the recent high and is adopted as 'normal'. Held constant across scenarios because it is "
     "a structural property of the industry, not a cyclical variable."),

    ("D25", "Maximum practical capacity utilisation", "%",
     {"Base Case": _flat(0.92), "Bull Case": _flat(0.93),
      "Bear Case": _flat(0.92), "Stress Case": _flat(0.92)},
     "Indicative", "Low",
     "Caps crude steel production at capacity multiplied by this factor. Where demand-derived "
     "production would exceed the cap, the shortfall is met by additional imports rather than "
     "by impossible domestic output, which keeps the trade block internally consistent. "
     "INDICATIVE: 92% allows for planned maintenance, relining and the fact that a national "
     "aggregate can never run at nameplate. Evidence that individual assets can exceed "
     "nameplate: the Master Database records SAIL's Rourkela at 103.8% and Durgapur at 101.6% "
     "average utilisation over FY2021-FY2025, but SAIL's Alloy Steels Plant averaged only "
     "39.8%, so the national aggregate is well below the best asset. In the base case the cap "
     "does not bind (peak utilisation 84.0% in FY2033), which is itself an important finding: "
     "India is not capacity-constrained on this demand path."),

    ("D26", "Depreciation and amortisation", "% of revenue",
     {"Base Case": _flat(0.053), "Bull Case": _flat(0.053),
      "Bear Case": _flat(0.055), "Stress Case": _flat(0.057)},
     "Derived-DB", "High",
     "DERIVED FROM THE MASTER DATABASE: FY2026 consolidated depreciation and amortisation for "
     "the four majors totals Rs 30,715 cr on Rs 5,81,646 cr of revenue, i.e. 5.28%. The ratio "
     "is remarkably tight across companies - Tata Steel 5.15%, JSW Steel 5.18%, SAIL 5.40%, "
     "Jindal Steel 5.96% - which gives high confidence that 5.3% is representative rather than "
     "an artefact. Bear and stress carry a slightly higher ratio because depreciation is fixed "
     "in rupee terms while revenue falls. Used only to compute taxable EBIT; it does not affect "
     "EBITDA."),

    ("D22", "Secondary and unattributed capacity additions", "Mtpa p.a.",
     {"Base Case":   [11.0, 11.0, 11.0, 10.5, 10.5, 10.0, 10.0],
      "Bull Case":   [13.0, 13.0, 12.5, 12.5, 12.0, 12.0, 11.5],
      "Bear Case":   [8.0, 8.0, 7.5, 7.0, 7.0, 6.5, 6.5],
      "Stress Case": [5.0, 4.0, 4.0, 3.5, 3.5, 3.0, 3.0]},
     "Derived-DB", "Medium",
     "ESSENTIAL RECONCILIATION ITEM. The dated project tracker on sheet 07 captures only large, "
     "individually announced projects and averages 7.7 Mtpa a year over FY2027-FY2033. India "
     "actually added 20.82 Mtpa in FY2025 and 20.07 Mtpa in FY2026 (Master DB), so roughly "
     "12.7 Mtpa a year has historically come from the secondary sector and from projects too "
     "small to be individually announced. The Master Database evidences this directly: "
     "production by 'other producers' rose from 66.61 Mt in FY2025 to 73.02 Mt in FY2026 and "
     "compounded at 10.7% over FY2022-FY2026, faster than the integrated majors' 8.1%. Base "
     "assumes 11.0 Mtpa a year tapering to 10.0, slightly below the historical rate. "
     "VALIDATION: this build produces India crude steel capacity of about 302 Mtpa by FY2031 "
     "against the National Steel Policy 2017 target of 300 Mtpa - the model reproduces the "
     "policy target from bottom-up components without being calibrated to it, which is a "
     "meaningful independent check on the capacity block. Omitting this driver would understate "
     "FY2031 capacity by about 54 Mtpa, or 18%."),

    ("D06", "Crude-to-finished steel ratio", "x",
     {"Base Case": _flat(1.0465), "Bull Case": _flat(1.0430),
      "Bear Case": _flat(1.0500), "Stress Case": _flat(1.0530)},
     "Derived-DB", "High",
     "FY2026 actual 168.42 / 160.94 = 1.0465 on the April-2026 JPC vintage used throughout the "
     "model (Master DB). Held flat in the base case: yield "
     "improvement is slow and offsetting mix effects are material. Bull assumes modest yield "
     "gains from continuous casting and newer assets; bear and stress assume yield "
     "deterioration from higher utilisation of older assets."),

    ("D07", "Net finished steel exports", "Mt",
     {"Base Case":   [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5],
      "Bull Case":   [0.0, 0.0, 0.5, 0.5, 1.0, 1.0, 1.5],
      "Bear Case":   [1.5, 2.5, 3.5, 4.5, 5.5, 6.0, 6.5],
      "Stress Case": [-1.0, -2.5, -3.5, -4.0, -4.5, -5.0, -5.0]},
     "Derived-DB", "Medium",
     "FY2026 actual net exports +0.078 Mt, having swung 4.77 Mt from a 4.693 Mt net import "
     "position in FY2025 (Master DB). Base assumes India becomes a progressively larger net "
     "exporter as capacity growth outpaces demand growth. Bear case is HIGHER net exports, "
     "because weak domestic demand forces surplus tonnes offshore at lower realisations. "
     "Stress reverses the sign: a demand collapse in India combined with the safeguard duty "
     "lapsing in Apr-2028 makes India a net importer again, which is the historically observed "
     "FY2025 configuration."),

    ("D08", "Variation in stock", "Mt",
     {"Base Case":   [-1.0, 0.0, 0.5, 0.5, 0.5, 0.5, 0.5],
      "Bull Case":   [-1.5, -0.5, 0.0, 0.0, 0.5, 0.5, 0.5],
      "Bear Case":   [0.5, 1.5, 1.5, 1.0, 1.0, 1.0, 1.0],
      "Stress Case": [2.0, 3.0, 2.5, 2.0, 1.5, 1.0, 1.0]},
     "Derived-DB", "Medium",
     "FY2026 actual -2.88 Mt, i.e. a destock of about 1.8% of consumption which flattered "
     "reported consumption growth (Master DB, supply balance line 6). Base normalises toward "
     "a small restock. Bear and stress assume inventory build, which is the classic signature "
     "of a demand disappointment and is a drag on apparent consumption."),

    # ---------------- PRICES
    ("D09", "Blended industry realisation", "Rs/t",
     {"Base Case":   [61800, 61000, 60500, 62000, 64000, 66000, 68000],
      "Bull Case":   [62500, 63500, 64000, 65500, 66500, 67500, 68500],
      "Bear Case":   [57500, 55000, 54000, 55000, 56500, 58000, 59500],
      "Stress Case": [59000, 58500, 60000, 62500, 65000, 67000, 69000]},
     "Derived-DB", "Medium",
     "FY2026A anchor of Rs 59,974/t is DERIVED from the Master Database as the volume-weighted "
     "revenue per tonne of Tata Steel India, JSW Steel India, SAIL and Jindal Steel - "
     "Rs 4,79,191 cr over 79.90 Mt, i.e. 49.4% of India's finished steel production. This is "
     "used in preference to a spot HRC quote because the database documents three "
     "irreconcilable Mar-2026 HRC assessments (Rs 55,900, Rs 57,700 and Rs 59,500/t - Conflict "
     "C07) and because a blended realisation is what actually drives revenue. Base path: "
     "FY2027 +3% carrying Q4FY2026 momentum (HRC rose ~14% QoQ to Rs 57,700/t by end-Mar-2026 "
     "on the safeguard duty); then a two-year decline as the safeguard duty tapers from 12% to "
     "11.5% (from 21-Apr-2026) to 11% and expires 20-Apr-2028, coinciding with peak capacity "
     "commissioning; then recovery to a ~2.0% CAGR over the horizon, below expected inflation, "
     "reflecting structural oversupply. Bear and stress are calibrated to the observed 52-week "
     "HRC low of Rs 45,700/t (IDBI, Jun-2026) as a floor reference."),

    ("D10", "Iron ore 62% Fe CFR China", "US$/dmt",
     {"Base Case":   [102, 97, 93, 90, 90, 91, 93],
      "Bull Case":   [98, 92, 88, 85, 84, 84, 86],
      "Bear Case":   [96, 90, 86, 83, 82, 82, 83],
      "Stress Case": [115, 112, 104, 98, 94, 92, 92]},
     "S17", "High",
     "FY2026 average US$100.52/dmt, the third consecutive annual decline from US$155.52/dmt in "
     "FY2022 (World Bank Pink Sheet, Master DB). Q1FY2027 actual average US$105.17/dmt with "
     "Jun-2026 at US$100.8/dmt, so the FY2027 base of US$102 is anchored on realised data. "
     "Base assumes continued gradual decline on new global supply, then stabilisation around "
     "US$90/dmt. NOTE the sign convention: LOWER iron ore is BULLISH for Indian steel margins "
     "because most Indian producers are net purchasers of ore, so the bull scenario carries the "
     "LOW price path. Godawari Power and NMDC are the exceptions and are flagged separately."),

    ("D11", "Premium HCC coking coal FOB Australia", "US$/t",
     {"Base Case":   [210, 200, 195, 190, 190, 195, 200],
      "Bull Case":   [200, 188, 182, 178, 178, 180, 184],
      "Bear Case":   [200, 185, 176, 170, 168, 168, 170],
      "Stress Case": [300, 285, 245, 220, 208, 202, 200]},
     "S01", "Medium",
     "IMPORTANT: the Master Database deliberately asserts NO fiscal-year average for coking "
     "coal because the two public series disagree by up to US$40/t (Conflict C01). The FY2026 "
     "reference used here is the Ministry of Steel's Mar-2026 assessment of ~US$225/t, "
     "explicitly labelled a spot reference and not a fiscal-year average. Base assumes easing "
     "from that level toward a US$190/t mid-cycle, still well above the ~US$175/t Mar-2025 "
     "trough. Stress reflects a repeat of the Oct-2023 spike to US$354/t at roughly 80% "
     "amplitude. THIS IS THE SINGLE LARGEST COST UNCERTAINTY IN THE MODEL and the driver most "
     "in need of a licensed price-reporting-agency series before transaction use."),

    ("D12", "Conversion cost inflation", "%",
     {"Base Case":   [0.045, 0.042, 0.040, 0.040, 0.040, 0.040, 0.040],
      "Bull Case":   [0.038, 0.035, 0.033, 0.032, 0.032, 0.032, 0.032],
      "Bear Case":   [0.050, 0.046, 0.044, 0.042, 0.042, 0.042, 0.042],
      "Stress Case": [0.075, 0.068, 0.055, 0.048, 0.044, 0.042, 0.042]},
     "RBI-Jun26", "Medium",
     "Applies to the non-raw-material element of cash cost: labour, power, consumables, "
     "stores, repairs and logistics. Anchored slightly below RBI's 5.1% FY2027 CPI projection "
     "because steel conversion cost includes power and freight that track energy rather than "
     "headline CPI, and because the majors are delivering explicit cost programmes - Tata "
     "Steel recorded roughly Rs 10,868 cr of cost transformation benefit in FY2026 (Master DB), "
     "equal to about Rs 4,825 per tonne of India deliveries. Stress reflects the West Asia "
     "energy shock flagged by both Tata Steel and Jindal Stainless in their FY2026 filings."),

    ("D13", "Raw material share of cash cost", "%",
     {"Base Case": _flat(0.60), "Bull Case": _flat(0.60),
      "Bear Case": _flat(0.60), "Stress Case": _flat(0.60)},
     "Indicative", "Low",
     "INDICATIVE STRUCTURAL ASSUMPTION - NOT SOURCED FROM A PRIMARY DOCUMENT. Splits the "
     "FY2026 implied cash cost of Rs 49,241/t between a raw-material component that indexes to "
     "the iron ore and coking coal basket and a conversion component that indexes to inflation. "
     "60/40 is the conventional split for an Indian integrated BF-BOF producer. Held constant "
     "across scenarios precisely because it is unsourced - flexing an unsourced parameter "
     "across scenarios would manufacture false precision. Plausible range 55-65%; the "
     "sensitivity of EBITDA/t to this parameter is quantified on sheet 22. MUST be replaced "
     "with company-specific cost disclosure from annual report technical parameters before any "
     "transaction use."),

    ("D21", "Finished steel imports", "Mt",
     {"Base Case":   [6.8, 7.0, 7.2, 7.5, 7.8, 8.1, 8.4],
      "Bull Case":   [7.5, 8.0, 8.5, 9.0, 9.5, 10.0, 10.5],
      "Bear Case":   [6.0, 5.5, 5.2, 5.0, 5.0, 5.1, 5.2],
      "Stress Case": [8.5, 10.0, 11.0, 11.5, 12.0, 12.0, 12.0]},
     "Derived-DB", "Medium",
     "FY2026 actual 6.524 Mt, down 31.7% from 9.551 Mt in FY2025 as the safeguard duty bit "
     "(Master DB). Base assumes imports grow slowly with the market as the duty tapers to 11.5% "
     "then 11% and expires 20-Apr-2028, but stay well below the FY2025 peak because domestic "
     "capacity has expanded. Note the counter-intuitive scenario signs: the BULL case carries "
     "HIGHER imports because strong domestic demand pulls material in, while the BEAR case "
     "carries LOWER imports because weak demand and a wide domestic discount to import parity "
     "make imports uncompetitive. Stress assumes a return above the FY2025 peak on duty expiry "
     "plus Chinese surplus redirection. Exports are derived as imports plus net exports (D07), "
     "so the two trade drivers are internally consistent by construction."),

    ("D14", "Iron ore weight in the raw material basket", "%",
     {"Base Case": _flat(0.45), "Bull Case": _flat(0.45),
      "Bear Case": _flat(0.45), "Stress Case": _flat(0.45)},
     "Indicative", "Low",
     "INDICATIVE. Iron ore 45%, coking coal 45%, other ferrous and fluxes 10% by value in the "
     "BF-BOF raw material basket. Unsourced and therefore held constant across scenarios. "
     "Sensitivity quantified on sheet 22."),

    # ---------------- CAPITAL / VALUATION
    ("D15", "Effective tax rate", "%",
     {"Base Case": _flat(0.2517), "Bull Case": _flat(0.2517),
      "Bear Case": _flat(0.2517), "Stress Case": _flat(0.2517)},
     "Statutory", "Medium",
     "Indian statutory corporate tax rate of 22% plus surcharge and cess equals 25.17% under "
     "section 115BAA, which the large Indian steel producers have elected. Held flat across "
     "scenarios: the tax rate is a policy parameter, not a cyclical variable. Actual effective "
     "rates differ because of MAT credits, accumulated losses and deferred tax - the Master "
     "Database flags Kirloskar Ferrous FY2026 net profit exceeding profit before tax on a tax "
     "write-back (Conflict C12). Company-specific effective rates must be substituted in "
     "single-name models."),

    ("D16", "Capex intensity", "Rs/t of capacity added",
     {"Base Case": _flat(55000), "Bull Case": _flat(52000),
      "Bear Case": _flat(60000), "Stress Case": _flat(65000)},
     "Derived-DB", "Medium",
     "Greenfield and large brownfield Indian crude steel capacity. Cross-checks from the "
     "Master Database: JSW Steel Board-approved 5 Mtpa brownfield expansion at JVML-Vijayanagar "
     "at Rs 26,000 cr equals Rs 52,000 per tonne; SAIL's stated roughly Rs 1 lakh crore for "
     "about 15 Mtpa of additions equals roughly Rs 67,000 per tonne; Tata Steel's 0.75 Mtpa "
     "Ludhiana scrap EAF at about Rs 3,200 cr equals Rs 42,700 per tonne, which is lower "
     "because an EAF has no ironmaking. Base of Rs 55,000/t sits between the JSW brownfield and "
     "SAIL blended figures. Also the primary input to the replacement-cost valuation on sheet 23."),

    ("D17", "Maintenance capex", "% of revenue",
     {"Base Case": _flat(0.030), "Bull Case": _flat(0.028),
      "Bear Case": _flat(0.033), "Stress Case": _flat(0.035)},
     "Indicative", "Low",
     "INDICATIVE. Sustaining capex excluding growth. Cross-check: Tata Steel FY2026 total capex "
     "of Rs 14,026 cr on consolidated revenue of Rs 2,32,140 cr is 6.0%, but that figure "
     "includes substantial growth capex, so a 3.0% sustaining assumption is consistent with a "
     "roughly 50/50 growth-maintenance split. Not separately disclosed by any issuer in the "
     "database, hence Low confidence."),

    ("D18", "Net working capital", "days of revenue",
     {"Base Case":   _flat(45), "Bull Case": _flat(40),
      "Bear Case":   _flat(52), "Stress Case": _flat(62)},
     "Not sourced", "Low",
     "NOT SOURCED. The Master Database carries total assets and borrowings but not receivables, "
     "inventory or payables, so working capital days could not be derived from it. Directional "
     "evidence only: Tata Steel disclosed a working capital RELEASE of about Rs 6,470 cr in "
     "FY2026 (Master DB), consistent with tightening days on a Rs 2,32,140 cr revenue base of "
     "roughly 10 days. Base of 45 days is a sector convention for an integrated Indian steel "
     "producer and MUST be replaced with balance-sheet-derived days before transaction use. "
     "Stress reflects a working capital build in a downturn as inventory rises and receivables "
     "stretch, which is a first-order cash flow risk, not a second-order one."),

    ("D19", "Exit EV/EBITDA multiple", "x",
     {"Base Case": _flat(6.0), "Bull Case": _flat(7.0),
      "Bear Case": _flat(5.0), "Stress Case": _flat(4.0)},
     "Not sourced", "Low",
     "NOT SOURCED - market multiples were not obtained in this research cycle because share "
     "prices were not sourced. 6.0x is the conventional mid-cycle EV/EBITDA for Indian "
     "integrated steel and is provided as a starting point only. Sheet 23 computes live "
     "trading multiples the moment share prices are entered, and the exit multiple should then "
     "be re-based on that observed evidence. Flagged Low confidence and excluded from the "
     "headline conclusions."),

    ("D20", "WACC (nominal, INR)", "%",
     {"Base Case": _flat(0.1200), "Bull Case": _flat(0.1120),
      "Bear Case": _flat(0.1280), "Stress Case": _flat(0.1400)},
     "Composite", "Medium",
     "Built from observable Indian market data where available. Risk-free rate 6.83% - the "
     "10-year G-sec yield on 31-Jul-2026, with Trading Economics expecting 6.78% at quarter "
     "end. Pre-tax cost of debt 7.44% - the lower bound of CRISIL RateView's Jul-2026 "
     "projected 10-year corporate bond yield range of 7.44-7.54%, consistent with the AA "
     "domestic ratings the Master Database records for JSW Steel. Post-tax cost of debt "
     "therefore 5.57% at the 25.17% tax rate. Equity risk premium of 6.5% and an equity beta "
     "of 1.2 are NOT SOURCED and are the weakest links in this build; at a 30% debt weight "
     "these give a cost of equity of about 14.6% and a WACC of about 12.0%. The full build-up "
     "is shown and is fully editable on sheet 02, and WACC is a sensitivity axis on sheet 22."),
]

# ======================================================================================
# CAPACITY EXPANSION TRACKER - dated project pipeline, all from the Master Database
# ======================================================================================
PROJECT_HEADERS = [
    "Project ID", "Company", "Project / Asset", "State / Country", "Type",
    "Capacity Added (Mtpa)", "Capacity Definition", "Status At Cutoff",
    "Expected Commissioning FY", "Capex (Rs cr)", "Implied Capex (Rs/t)",
    "Included In Forecast?", "Source ID", "Confidence", "Notes",
]

PROJECTS = [
    ("P01", "Tata Steel", "Ludhiana scrap-based EAF", "Punjab", "Brownfield EAF", 0.75,
     "Crude steel", "Commissioned Mar-2026", "FY2026", 3200, None, "No - already in base",
     "S08", "High",
     "Commissioned in Mar-2026 and therefore already inside the FY2026 capacity base of "
     "220.4 Mtpa. Designed for under 0.3 tCO2e per tonne of crude steel, which makes it "
     "5-star green steel under the Taxonomy notified 23-Dec-2024."),
    ("P02", "Tata Steel", "NINL expansion", "Odisha", "Brownfield integrated", 4.80,
     "Crude steel", "Proposed", "FY2030", None, None, "Yes", "S08", "Medium",
     "Described as proposed in the FY2026 results release, so not yet Board-approved capex. "
     "Commissioning year is the modeller's estimate based on typical 4-year build."),
    ("P03", "Tata Steel", "India roadmap to 40 Mtpa", "India", "Programme", 5.00,
     "Crude steel", "Announced - aspirational", "FY2033", None, None, "Yes", "S37", "Low",
     "Tata Steel's Q1FY2027 presentation (31-Jul-2026) sets out an India roadmap to 40 Mtpa. "
     "The residual above identified projects is carried at low confidence and is materially "
     "reduced by the delivery factor D05."),
    ("P04", "JSW Steel", "Vijayanagar BF-3 expansion 3.0 to 4.5 Mtpa", "Karnataka",
     "Brownfield BF", 1.50, "Crude steel", "Under testing and commissioning", "FY2027",
     None, None, "Yes", "S09", "High",
     "Was under testing at FY2026 year-end and excluded from the FY2026 utilisation "
     "calculation, so it is genuinely incremental to the base."),
    ("P05", "JSW Steel", "JVML-Vijayanagar brownfield expansion", "Karnataka",
     "Brownfield integrated", 5.00, "Crude steel", "Board approved", "FY2030", 26000, None,
     "Yes", "S09", "High",
     "Board approved. Rs 26,000 cr for 5 Mtpa gives Rs 52,000 per tonne, the cleanest "
     "brownfield capex-intensity datapoint available and the primary calibration for D16."),
    ("P06", "JSW Steel", "BMM Ispat acquisition", "Karnataka", "Acquisition", 0.90,
     "Crude steel", "Board approved", "FY2027", None, None,
     "No - acquisition, not new national capacity", "S09", "High",
     "Acquisition of existing operating capacity, so it adds to JSW Steel's own capacity and "
     "volume but adds NOTHING to India's national capacity - those tonnes are already in the "
     "220.4 Mtpa FY2026 base. Correctly excluded from the industry capacity build. This is a "
     "standard trap: summing company announcements without stripping out acquisitions "
     "double-counts national capacity. JSW's FY2027 guidance explicitly includes BMM Ispat "
     "production of 0.75 Mt and sales of 0.70 Mt, so it IS included in the company-level "
     "volume build on sheet 12."),
    ("P07", "JSW Steel / POSCO", "Greenfield integrated plant, 50:50 JV", "Odisha",
     "Greenfield JV", 6.00, "Crude steel", "Announced Apr-2026", "FY2031", None, None, "Yes",
     "S09", "Medium",
     "50:50 joint venture announced Apr-2026 targeting commissioning by 2031. This is the "
     "single most distant dated project in the pipeline and is the reason the announced "
     "pipeline terminates in FY2031, which in turn drives the horizon selection."),
    ("P08", "SAIL", "IISCO Steel Plant greenfield expansion", "West Bengal", "Greenfield",
     4.08, "Crude steel", "Board approved", "FY2031", None, None, "Yes", "S26", "High",
     "Board approved per Ministry of Steel, Rajya Sabha USQ 3982 answered 27-Mar-2026. IISCO "
     "averaged 91.6% utilisation over FY2021-FY2025 on 2.5 Mtpa."),
    ("P09", "SAIL", "Bokaro Steel Plant brownfield expansion", "Jharkhand", "Brownfield",
     2.65, "Crude steel", "Board approved", "FY2030", None, None, "Yes", "S26", "High",
     "Board approved per the same Parliamentary answer. Bokaro averaged 87.8% utilisation on "
     "4.6 Mtpa, so there is headroom in the existing asset as well."),
    ("P10", "SAIL", "Durgapur Steel Plant brownfield expansion", "West Bengal", "Brownfield",
     0.89, "Crude steel", "Board approved", "FY2029", None, None, "Yes", "S26", "High",
     "Board approved. Durgapur already averaged 101.6% utilisation on 2.2 Mtpa over "
     "FY2021-FY2025, i.e. it is running above nameplate and the expansion is debottlenecking."),
    ("P11", "SAIL", "Residual ambition to ~35 Mt by FY2031", "India", "Programme", 8.00,
     "Crude steel", "Announced - aspirational", "FY2033", None, None, "Yes", "S47", "Low",
     "SAIL has publicly targeted roughly 35 Mt by FY2031 against 19.676 Mtpa today at an "
     "indicated roughly Rs 1 lakh crore outlay. Board-approved tranches total only 7.62 Mtpa, "
     "so about 8 Mtpa of the ambition is unapproved and is carried at low confidence."),
    ("P12", "AM/NS India", "Hazira expansion 9 to 15 Mtpa", "Gujarat", "Brownfield integrated",
     6.00, "Crude steel", "Under implementation", "FY2029", None, None, "Yes", "S13", "High",
     "Under implementation. The largest single brownfield addition in the pipeline."),
    ("P13", "AM/NS India", "Rajayyapeta greenfield, Phase 1", "Andhra Pradesh", "Greenfield",
     8.20, "Crude steel", "Announced 23-Mar-2026", "FY2032", None, None, "Yes", "S14", "Medium",
     "Announced 23-Mar-2026 as a low-cost, highly efficient coastal asset. Phase 1 of 8.2 Mtpa. "
     "Greenfield of this scale in India has historically taken 5 to 7 years."),
    ("P14", "AM/NS India", "Hazira further expansion to 24 Mtpa", "Gujarat", "Brownfield",
     9.00, "Crude steel", "Detailed engineering under way", "Beyond FY2033", None, None,
     "No - beyond horizon", "S13", "Low",
     "Detailed engineering only. Falls outside the FY2033 horizon and is therefore excluded "
     "from the capacity build. Retained in the tracker for completeness and for terminal-value "
     "context."),
    ("P15", "Jindal Steel", "Angul expansion 6.0 to 12.0 Mtpa", "Odisha",
     "Brownfield integrated", 6.00, "Crude steel", "Completed in FY2026", "FY2026", None, None,
     "No - already in base", "S11", "High",
     "Completed during FY2026, taking group steelmaking capacity from 9.6 to 15.6 Mtpa. "
     "Already inside the FY2026 capacity base. Ironmaking capacity is being raised by a "
     "further 6.6 Mtpa with elements still under construction."),
    ("P16", "Jindal Stainless", "Indonesia stainless melt shop", "Indonesia", "Greenfield",
     1.20, "Stainless melt", "Commissioned in FY2026", "FY2026", None, None,
     "No - outside India and in base", "S12", "High",
     "Commissioned ahead of schedule in FY2026, taking group melt capacity to 4.2 Mtpa. "
     "OUTSIDE INDIA, so excluded from the India capacity build."),
    ("P17", "Jindal Stainless", "Jajpur HRAP and CRAP downstream lines", "Odisha",
     "Downstream", 1.27, "Downstream processing", "Upcoming", "FY2028", None, None,
     "No - downstream, not melt", "S12", "High",
     "1.1 Mtpa HRAP plus 0.17 Mtpa CRAP. DOWNSTREAM PROCESSING, not melt capacity, so it must "
     "not be added to crude steel capacity. Included in the tracker because it drives mix and "
     "realisation, not tonnage."),
    ("P18", "Jindal Stainless", "Hisar and Kharagpur cold rolling augmentation",
     "Haryana / West Bengal", "Downstream", None, "Downstream processing",
     "Announced - Rs 900 cr committed", "FY2028", 900, None, "No - downstream, not melt",
     "S12", "High",
     "Takes CRAP capacity to 2.67 Mtpa by FY2028. Downstream only."),
    ("P19", "Electrosteel / ESL", "Bokaro capacity doubling", "Jharkhand", "Brownfield", 1.50,
     "Crude steel", "Announced by parent", "FY2027", None, None, "Yes", "S43", "Low",
     "Vedanta targeting a doubling of Bokaro from 1.5 to 3.0 Mtpa by end-CY2026. Rests on a "
     "single secondary source and is carried at low confidence."),
    ("P20", "NMDC Steel", "Nagarnar ramp-up to nameplate", "Chhattisgarh", "Ramp-up", None,
     "Crude steel", "Ramping", "FY2028", None, None, "No - utilisation, not capacity",
     "S01", "High",
     "The 3 Mtpa plant produced 2.382 Mt in FY2026, up 61.9%. The remaining roughly 0.6 Mt is "
     "a UTILISATION gain within existing capacity, not new capacity, and is captured through "
     "the utilisation forecast rather than the capacity build. Included here because it is one "
     "of the two empirical ramp-rate observations underpinning the horizon choice."),
]

# ======================================================================================
# ASSUMPTION REGISTRY - the 13 mandated fields for every assumption
# ======================================================================================
REGISTER_HEADERS = [
    "Assumption ID", "Assumption", "Unit", "Current Value (FY2026A / spot)",
    "Historical Range", "Forecast Value (Base, terminal year)", "Source",
    "Publication Date", "Confidence Rating", "Evidence", "Reasoning",
    "Sensitivity (impact of a 10% adverse move)", "Dependencies", "Review Date", "Owner",
]

REGISTER = [
    ("R01", "India real GDP growth", "%", "6.6% (FY2027 projection)", "FY2021 -5.8% to FY2026 7.6%",
     "6.5% (FY2033)", "Reserve Bank of India, Monetary Policy Committee statement",
     "05-Jun-2026", "High",
     "RBI cut its FY2027 projection to 6.6% from 6.9%, citing West Asia conflict, energy "
     "prices, supply disruption and monsoon risk; repo held at 5.25%; FY2027 CPI projection "
     "raised 50bps to 5.1%. FY2026 growth estimated at 7.6%; Q4FY2026 actual 7.7%.",
     "The RBI is the authoritative domestic forecaster and its projection is adopted unadjusted "
     "for FY2027. Beyond FY2027 no Indian official body publishes a real GDP path, so the model "
     "converges to a 6.5% steady state, which is below the FY2015-FY2026 realised average and "
     "therefore not an aggressive assumption.",
     "A 10% adverse move (66bps of growth) reduces FY2033 steel demand by roughly 5.0 Mt and "
     "industry EBITDA by roughly 4%.",
     "Drives D02 steel demand elasticity, hence all volume, revenue and EBITDA lines.",
     "Next RBI MPC statement", "Head of Metals & Mining Research"),

    ("R02", "Steel demand elasticity to GDP", "x", "1.10x realised FY2015-FY2026",
     "1.06x to 1.13x depending on the GDP deflator used", "1.05x (FY2033)",
     "Derived from Master Industry Database; corroborated by worldsteel Short Range Outlook",
     "Master DB 03-Aug-2026; worldsteel SRO Apr-2026", "High",
     "Apparent consumption compounded at 7.13% from 76.99 Mt in FY2015 to 164.19 Mt in FY2026 "
     "against real GDP CAGR of roughly 6.3-6.7%. worldsteel's Apr-2026 SRO forecasts Indian "
     "demand +7.4% in CY2026 and +9.2% in CY2027, the fastest of any major market, against RBI "
     "GDP of 6.6% - implying near-term elasticity of 1.1x to 1.4x.",
     "Set at 1.15x near term to reflect the worldsteel-implied strength, tapering to 1.05x by "
     "FY2033 on maturing steel intensity as per capita consumption rises from 115.7 kg. India "
     "is at roughly half the world average of 215 kg, so intensity decay should be slow.",
     "A 10% adverse move (0.115x) reduces FY2033 demand by roughly 9.5 Mt, about 3.6%.",
     "Depends on R01. Drives the entire demand, supply and revenue chain.",
     "On each worldsteel SRO release (April and October)", "Head of Metals & Mining Research"),

    ("R03", "Blended industry realisation", "Rs/t", "Rs 59,974/t (FY2026A)",
     "Implied Rs 55,600/t (SAIL) to Rs 62,273/t (Tata Steel India) across the four majors",
     "Rs 68,000/t (FY2033)", "Derived from Master Industry Database company disclosures",
     "Master DB 03-Aug-2026", "Medium",
     "Volume-weighted revenue per tonne of Tata Steel India, JSW Steel India, SAIL and Jindal "
     "Steel: Rs 4,79,191 cr over 79.90 Mt, which is 49.4% of India's FY2026 finished steel "
     "production. Cross-check: implied EBITDA of Rs 85,759 cr against the Rs 86,318 cr sum of "
     "reported EBITDAs, a 0.6% rounding difference.",
     "A blended realisation is used in preference to a spot HRC quote for two reasons. First, "
     "the database documents three irreconcilable Mar-2026 HRC assessments spanning Rs 55,900 "
     "to Rs 59,500/t (Conflict C07), so no single spot number is defensible. Second, revenue is "
     "driven by the actual product mix, not by HRC alone. The forecast path reflects the "
     "safeguard duty taper from 12% to 11.5% to 11% and its expiry on 20-Apr-2028 coinciding "
     "with peak capacity commissioning.",
     "A 10% adverse move (Rs 6,000/t) reduces FY2033 industry EBITDA by roughly Rs 1.1 lakh "
     "crore, or roughly 60% - realisation is by far the highest-leverage driver in the model.",
     "Drives revenue, EBITDA, margin, cash flow and valuation. Interacts with D07 net exports "
     "via import and export parity.",
     "Quarterly, on each results season", "Head of Metals & Mining Research"),

    ("R04", "Implied cash cost per tonne", "Rs/t", "Rs 49,241/t (FY2026A)",
     "Rs 47,060/t (Tata Steel India) to Rs 50,837/t (Jindal Steel)", "Rs 55,900/t (FY2033)",
     "Derived from Master Industry Database", "Master DB 03-Aug-2026", "Medium",
     "Realisation of Rs 59,974/t less weighted EBITDA of Rs 10,733/t. The dispersion across the "
     "four majors is only about Rs 3,800/t, i.e. under 8%, which gives confidence that the "
     "weighted figure is representative rather than an artefact of one company.",
     "Forecast by splitting cash cost into a raw-material component indexed to the iron ore and "
     "coking coal basket in rupee terms, and a conversion component indexed to inflation. This "
     "top-down calibration is used instead of a bottom-up build with consumption coefficients "
     "because those coefficients could not be sourced from a primary document, and inventing "
     "them would create false precision.",
     "A 10% adverse move (Rs 4,924/t) reduces FY2033 industry EBITDA by roughly 50%.",
     "Depends on D10 iron ore, D11 coking coal, D03 USD/INR, D12 conversion inflation, D13 and "
     "D14 basket structure.",
     "Quarterly", "Head of Metals & Mining Research"),

    ("R05", "Premium HCC coking coal", "US$/t", "US$225/t (Mar-2026 spot reference)",
     "US$175/t (Mar-2025) to US$354/t (Oct-2023)", "US$200/t (FY2033)",
     "Ministry of Steel Monthly Economic Report", "Apr-2026 vintage", "Medium",
     "The Master Database deliberately asserts no fiscal-year average because the Ministry of "
     "Steel and broker series disagree by up to US$40/t and the broker 52-week high of "
     "US$217/t sits below the Ministry's Feb-2026 level of US$246/t (Conflict C01).",
     "The Ministry series is preferred as the Government's own published benchmark. The FY2026 "
     "reference is explicitly a spot observation, not an average, and this limitation is carried "
     "through to every dependent output.",
     "A 10% adverse move (US$22.5/t) raises FY2033 cash cost by roughly Rs 1,050/t and reduces "
     "industry EBITDA by roughly 9%.",
     "Depends on D03 USD/INR for rupee conversion. Drives D13 raw material cost.",
     "Monthly, on each Ministry of Steel report", "Head of Metals & Mining Research"),

    ("R06", "Iron ore 62% Fe CFR China", "US$/dmt", "US$100.52/dmt (FY2026 average)",
     "US$52.18/dmt (FY2016) to US$155.52/dmt (FY2022)", "US$93/dmt (FY2033)",
     "World Bank Commodity Price Data (Pink Sheet)", "Monthly file updated 02-Jul-2026", "High",
     "Complete monthly history to Jun-2026. FY2026 average US$100.52/dmt, the third consecutive "
     "annual decline. Q1FY2027 actual average US$105.17/dmt.",
     "The highest-quality price series available to the model: an official publisher, a "
     "documented definition (62% Fe fines, CFR China) and a complete monthly history allowing "
     "genuine fiscal-year averages. The FY2027 assumption is anchored on realised Q1 data.",
     "A 10% adverse move (US$10/dmt) raises FY2033 cash cost by roughly Rs 470/t and reduces "
     "industry EBITDA by roughly 4%.",
     "Depends on D03 USD/INR. Note the inverse exposure for ore-integrated names.",
     "Monthly, on each Pink Sheet release", "Head of Metals & Mining Research"),

    ("R07", "Announced pipeline delivery factor", "%", "85% (base assumption)",
     "20.82 Mtpa added FY2025; 20.07 Mtpa added FY2026", "85% (held flat)",
     "Derived from Master Industry Database capacity series and project tracker",
     "Master DB 03-Aug-2026", "Medium",
     "India added 20.82 Mtpa in FY2025 and 20.07 Mtpa in FY2026 against a National Steel Policy "
     "path to 300 Mtpa by FY2031 requiring roughly 16 Mtpa a year, i.e. recent delivery has run "
     "ahead of the required rate. Offsetting evidence of slippage: JSW Vijayanagar BF-3 was "
     "still under testing at FY2026 year-end and AM/NS Hazira 9 to 15 Mtpa has been in progress "
     "for several years.",
     "85% balances a demonstrated ability to deliver roughly 20 Mtpa a year against "
     "project-level slippage. Applied only to projects not already commissioned, and only to "
     "genuinely incremental India crude steel capacity - acquisitions, downstream lines and "
     "overseas assets are excluded.",
     "A 10% adverse move (to 76.5%) reduces FY2033 capacity by roughly 5.2 Mtpa and raises "
     "utilisation by roughly 180bps, which is EBITDA-positive - so this driver works in the "
     "opposite direction to intuition and is a genuine model insight.",
     "Drives capacity, hence utilisation, hence the price and margin cycle.",
     "Quarterly, on results and project announcements", "Head of Metals & Mining Research"),

    ("R08", "WACC (nominal, INR)", "%", "12.0%", "Not applicable - a constructed parameter",
     "12.0% (held flat)", "Composite: RBI, investing.com, CRISIL RateView, modeller judgement",
     "10-year G-sec 31-Jul-2026; CRISIL RateView Jul-2026", "Medium",
     "Risk-free rate 6.83% - the 10-year G-sec yield on 31-Jul-2026, with Trading Economics "
     "expecting 6.78% at quarter end. Pre-tax cost of debt 7.44% - the lower bound of CRISIL "
     "RateView's projected 10-year corporate bond yield range of 7.44-7.54%, consistent with "
     "the AA domestic ratings recorded in the database for JSW Steel.",
     "Two inputs are NOT sourced and are the weakest links: an equity risk premium of 6.5% and "
     "an equity beta of 1.2. These are flagged, fully editable, and form a sensitivity axis. The "
     "full build-up is displayed rather than the output alone, so a reviewer can substitute "
     "their own house assumptions in one place.",
     "A 10% adverse move (120bps) reduces enterprise value by roughly 11% at the base exit "
     "multiple.",
     "Depends on D15 tax rate and the capital structure. Drives all discounted valuation output.",
     "Quarterly, or on any material rate move", "Head of Metals & Mining Research"),

    ("R09", "Net working capital", "days of revenue", "Not sourced",
     "Not derivable from the Master Database", "45 days (held flat)",
     "Sector convention; directional corroboration from Tata Steel disclosure",
     "Master DB 03-Aug-2026", "Low",
     "The Master Database carries total assets and borrowings but not receivables, inventory or "
     "payables, so working capital days could not be derived. The only directional evidence is "
     "Tata Steel's disclosed FY2026 working capital RELEASE of about Rs 6,470 cr, which on a "
     "Rs 2,32,140 cr revenue base is roughly 10 days of tightening.",
     "45 days is a sector convention for an integrated Indian producer and is explicitly "
     "provisional. It is flexed hard in the stress case to 62 days because a working capital "
     "build in a downturn is a first-order cash flow risk, not a second-order one.",
     "A 10% adverse move (4.5 days) consumes roughly Rs 12,000 cr of industry cash flow in the "
     "terminal year.",
     "Drives cash flow and net debt, not EBITDA.",
     "MUST be replaced before transaction use", "Head of Metals & Mining Research"),

    ("R10", "Raw material share of cash cost", "%", "60% (indicative)",
     "Not sourced", "60% (held flat)", "Indicative structural assumption - not sourced",
     "Not applicable", "Low",
     "No primary document was located that discloses the raw-material versus conversion split "
     "of cash cost for the Indian majors. 60/40 is the conventional split for an integrated "
     "BF-BOF producer.",
     "Held constant across all four scenarios precisely because it is unsourced: flexing an "
     "unsourced parameter across scenarios would manufacture false precision. Its influence is "
     "instead quantified explicitly on the sensitivity sheet, so a reviewer can see exactly how "
     "much of the answer depends on it. Plausible range 55-65%.",
     "A 5 percentage point change alters FY2033 EBITDA/t by roughly Rs 600/t, about 4%.",
     "Determines how D10, D11 and D12 transmit into cash cost.",
     "MUST be replaced with company cost disclosure before transaction use",
     "Head of Metals & Mining Research"),
]

# WACC build-up displayed explicitly on the assumptions sheet
WACC_BUILD = [
    ("Risk-free rate (10-year G-sec)", "%", 0.0683, "investing.com / NSE", "31-Jul-2026", "High",
     "10-year benchmark G-sec yield 6.833% on 31-Jul-2026. Trading Economics expects 6.78% at "
     "quarter end. RBI repo rate held at 5.25% on 05-Jun-2026."),
    ("Equity risk premium", "%", 0.0650, "NOT SOURCED - modeller assumption", "n/a", "Low",
     "Not sourced in this research cycle. 6.5% is a conventional India ERP. THIS IS ONE OF THE "
     "TWO WEAKEST INPUTS IN THE MODEL - substitute your house assumption."),
    ("Equity beta (sector, levered)", "x", 1.20, "NOT SOURCED - modeller assumption", "n/a",
     "Low",
     "Not sourced. 1.2 reflects the conventional view that steel is a high-beta cyclical. "
     "Substitute a regression beta or a peer-median unlevered beta relevered to target gearing."),
    ("Cost of equity", "%", None, "Computed", "n/a", "Medium",
     "Risk-free rate plus beta times equity risk premium. Live formula."),
    ("Pre-tax cost of debt", "%", 0.0744, "CRISIL RateView", "Jul-2026", "Medium",
     "Lower bound of CRISIL's projected 10-year corporate bond yield range of 7.44-7.54%. "
     "Consistent with the ICRA and India Ratings AA domestic ratings the Master Database records "
     "for JSW Steel, and with Moody's Ba1 positive at the international level."),
    ("Effective tax rate", "%", 0.2517, "Statutory section 115BAA", "n/a", "Medium",
     "22% plus surcharge and cess. See D15."),
    ("Post-tax cost of debt", "%", None, "Computed", "n/a", "Medium",
     "Pre-tax cost of debt times one minus the tax rate. Live formula."),
    ("Target debt weight", "%", 0.30, "Derived from Master Industry Database", "03-Aug-2026",
     "Medium",
     "Indicative sector gearing. Cross-check from the database: JSW Steel net debt to equity "
     "fell to 0.51x at Mar-2026 from 0.94x; Jindal Steel debt to equity 0.43x; Jindal Stainless "
     "net debt to equity 0.15x. A 30% debt weight is consistent with the middle of that range."),
    ("Target equity weight", "%", None, "Computed", "n/a", "Medium",
     "One minus the target debt weight. Live formula."),
    ("WACC (nominal, INR)", "%", None, "Computed", "n/a", "Medium",
     "Equity weight times cost of equity plus debt weight times post-tax cost of debt. Live "
     "formula. Cross-check against D20, which is the value used downstream."),
]
