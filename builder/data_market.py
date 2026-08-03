# -*- coding: utf-8 -*-
"""Market data: steel / iron ore / coking coal prices, demand, trade, policy, KPIs."""
from .style import NA

# ======================================================================================
# 7. STEEL PRICES
# ======================================================================================
STEEL_PRICE_HEADERS = [
    "Series", "Product / Grade", "Geography", "Assessment Basis", "Unit", "Currency",
    "Observation Date / Period", "Value", "52-Week High", "52-Week Low",
    "Period Type", "Source ID", "Source Tier", "Methodology", "Confidence", "Notes",
]

STEEL_PRICES = [
    # --- FY2026 domestic HRC path (ICRA)
    ("Domestic HRC - FY2026 price path", "Hot rolled coil", "India",
     "ICRA assessment; basis not fully specified; excludes GST", "Rs/t", "INR",
     "Early Dec-2025", 46000, None, None, "Point-in-time", "S19", 2, "As reported", "Medium",
     "Trough of the FY2026 cycle. Driven by excess supply over demand following the lapse of the provisional safeguard duty in Nov-2025. An alternative ICRA vintage cites ~Rs 47,000/t by mid-Nov-2025."),
    ("Domestic HRC - FY2026 price path", "Hot rolled coil", "India",
     "ICRA assessment; excludes GST", "Rs/t", "INR",
     "Sep-2025", 49500, None, None, "Point-in-time", "S19", 2, "As reported", "Medium",
     "Correction from the post-safeguard-duty highs of early FY2026."),
    ("Domestic HRC - FY2026 price path", "Hot rolled coil", "India",
     "ICRA assessment; excludes GST", "Rs/t", "INR",
     "Early Feb-2026", 53800, None, None, "Point-in-time", "S19", 2, "As reported", "Medium",
     "Rebound following reinstatement of the safeguard duty via the definitive notification of 30-Dec-2025, plus higher coking coal costs."),
    ("Domestic HRC - FY2026 price path", "Hot rolled coil", "India",
     "ICRA assessment, end-of-period; excludes GST", "Rs/t", "INR",
     "End Mar-2026", 57700, None, None, "Point-in-time", "S19", 2, "As reported", "High",
     "RECOMMENDED REFERENCE for end-FY2026 domestic HRC. Up ~14% quarter on quarter in Q4FY2026 on higher coking coal and iron ore costs. ICRA states the assessment point and the QoQ construction, making this figure reproducible - see Conflict C07."),
    # --- Alternative Mar-2026 assessments (Conflict C07)
    ("Domestic steel price - broker assessment", "Hot rolled coil (broker 'Indian steel price')",
     "India", "IDBI Capital monthly assessment", "Rs/t", "INR",
     "Mar-2026", 59500, None, None, "Month", "S20", 2, "As reported", "Medium",
     "Up 10.2% month on month; described as the highest level in 45 months. Differs from ICRA - see Conflict C07."),
    ("Domestic steel price - trade assessment", "Hot rolled coil", "India",
     "Trade assessment as reported by ETInfra", "Rs/t", "INR",
     "Mar-2026", 55900, None, None, "Point-in-time", "S40", 2, "As reported", "Low",
     "Reported alongside a Dec-2025 quarter comparative of Rs 47,317/t, i.e. an 18% increase attributed to the extended safeguard duty."),
    ("Domestic rebar - trade assessment", "Primary rebar", "India",
     "Trade assessment as reported by ETInfra", "Rs/t", "INR",
     "Mar-2026", 59800, None, None, "Point-in-time", "S40", 2, "As reported", "Low",
     "Against a Dec-2025 quarter comparative of Rs 47,615/t, i.e. a 25% increase. Longs outperformed flats on the safeguard duty because the duty covers FLAT products only, tightening flat supply while construction demand lifted longs."),
    # --- Jun-2026 cross-regional snapshot with ranges
    ("Regional HRC snapshot", "Hot rolled coil", "India", "IDBI Capital weekly spot; excludes GST",
     "Rs/t", "INR", "Week ended 22-Jun-2026", 58200, 59600, 45700, "Point-in-time",
     "S21", 2, "As reported", "Medium",
     "The 52-week range of Rs 45,700-59,600/t is a ~30% swing and is the single best illustration of why FY2027 earnings are highly price-sensitive. Flat week on week; down 0.9% month on month."),
    ("Regional HRC snapshot", "Hot rolled coil", "China", "FOB China, converted to INR by the broker",
     "Rs/t", "INR", "Week ended 22-Jun-2026", 48580, 49052, 41977, "Point-in-time",
     "S21", 2, "As reported", "Medium",
     "The Indian premium over Chinese FOB was ~Rs 9,620/t (~20%) at this date. This spread is the core determinant of import pressure and hence of the political durability of the safeguard duty."),
    ("Regional HRC snapshot", "Hot rolled coil", "South East Asia",
     "CNF South East Asia, converted to INR by the broker", "Rs/t", "INR",
     "Week ended 22-Jun-2026", 50938, 51410, 44807, "Point-in-time", "S21", 2, "As reported",
     "Medium", "Down 0.9% week on week."),
    ("Regional HRC snapshot", "Hot rolled coil", "CIS", "FOB CIS, converted to INR by the broker",
     "Rs/t", "INR", "Week ended 22-Jun-2026", 51410, 51410, 41505, "Point-in-time",
     "S21", 2, "As reported", "Medium",
     "Up 0.9% week on week and up 21.1% year on year - the strongest year-on-year move of the four regional benchmarks, reflecting West Asia supply-chain disruption."),
    ("International steel price", "Hot rolled coil", "China",
     "Broker assessment in USD", "US$/t", "USD", "Mar-2026", 500, None, None, "Month",
     "S20", 2, "As reported", "Medium",
     "Up 7.5% month on month to an 8-month high, attributed to supply disruption from Iran, higher energy and freight costs and improved mill profitability."),
    # --- Rebar benchmark
    ("Domestic rebar benchmark", "Rebar IS 1786 Fe 550D, 12-32 mm, BF route", "India (ex-Mumbai)",
     "Distributor-to-dealer, ex-Mumbai, EXCLUDES 18% GST", "Rs/t", "INR",
     "31-Jul-2026", 50900, None, None, "Point-in-time", "S22", 2, "As reported", "High",
     "BigMint benchmark. Up Rs 2,200/t week on week from Rs 48,700/t on 24-Jul-2026, on tightening supply and inventory restocking. The most precisely specified steel price in this workbook - grade, size, route, location, channel and tax treatment are all stated."),
    # --- Import parity
    ("Import parity", "Hot rolled coil", "India",
     "Domestic price versus landed cost of imports", "US$/t", "USD",
     "Apr-2026", None, -29, -45, "Point-in-time", "S19", 2, "As reported", "Medium",
     "Domestic HRC traded at a DISCOUNT of USD 29-45/t to landed imports at the start of FY2027, keeping imports uncompetitive and supporting domestic realisations. Shown in the high/low columns as a range; a negative sign denotes a discount to import parity."),
    # --- MoS retail series: existence documented, values not tabulated
    ("Ministry of Steel Mumbai retail series", "TMT 10 mm", "Mumbai (retail)",
     "Retail, INCLUDES GST; month-end", "Rs/t", "INR",
     "Monthly, 31-Oct-2023 to 31-Mar-2026", NA, None, None, "Month", "S01", 1,
     "Not tabulated at source", "High",
     "The Ministry of Steel publishes this month-end series as a CHART ONLY in its Monthly Economic Report; the underlying values are not tabulated in the document. Recorded here so the series is registered in the database with its correct definition and frequency. Obtain the numeric series from JPC on request. NOTE: GST-inclusive retail prices are NOT comparable with the ex-works and trade assessments above."),
    ("Ministry of Steel Mumbai retail series", "HR coils 2.50 mm", "Mumbai (retail)",
     "Retail, INCLUDES GST; month-end", "Rs/t", "INR",
     "Monthly, 31-Oct-2023 to 31-Mar-2026", NA, None, None, "Month", "S01", 1,
     "Not tabulated at source", "High", "As above."),
    ("Ministry of Steel Mumbai retail series", "CR coils 0.63 mm", "Mumbai (retail)",
     "Retail, INCLUDES GST; month-end", "Rs/t", "INR",
     "Monthly, 31-Oct-2023 to 31-Mar-2026", NA, None, None, "Month", "S01", 1,
     "Not tabulated at source", "High", "As above."),
]

# ======================================================================================
# 8. IRON ORE PRICES
# ======================================================================================
IRON_ORE_FY_HEADERS = [
    "Fiscal Year", "Iron Ore 62% Fe Fines, CFR China (US$/dmt)",
    "Months In Average", "Period Type", "Source ID", "Methodology", "Confidence", "Notes",
]

IRON_ORE_FY = [
    ("FY2016", 52.18, 12), ("FY2017", 67.81, 12), ("FY2018", 68.98, 12),
    ("FY2019", 71.99, 12), ("FY2020", 95.66, 12), ("FY2021", 128.03, 12),
    ("FY2022", 155.52, 12), ("FY2023", 117.20, 12), ("FY2024", 119.91, 12),
    ("FY2025", 103.97, 12), ("FY2026", 100.52, 12),
]

IRON_ORE_CY_HEADERS = [
    "Calendar Year", "Iron Ore 62% Fe Fines, CFR China (US$/dmt)",
    "Period Type", "Source ID", "Methodology", "Confidence", "Notes",
]

IRON_ORE_CY = [
    ("2015", 55.9), ("2016", 58.4), ("2017", 71.8), ("2018", 69.8), ("2019", 93.8),
    ("2020", 108.9), ("2021", 161.7), ("2022", 121.3), ("2023", 120.6),
    ("2024", 109.4), ("2025", 100.2),
]

IRON_ORE_MONTHLY_HEADERS = [
    "Month", "Iron Ore 62% Fe Fines, CFR China (US$/dmt)",
    "Australian THERMAL Coal, FOB Newcastle 6000 kcal/kg (US$/mt)",
    "Fiscal Year", "Fiscal Quarter", "Period Type", "Source ID", "Confidence",
]

# (label, iron ore, thermal coal, FY, FQ)
IRON_ORE_MONTHLY = [
    ("Apr-2025", 97.2, 98.6, "FY2026", "Q1"), ("May-2025", 97.0, 104.4, "FY2026", "Q1"),
    ("Jun-2025", 92.3, 109.0, "FY2026", "Q1"), ("Jul-2025", 97.3, 112.9, "FY2026", "Q2"),
    ("Aug-2025", 99.7, 112.2, "FY2026", "Q2"), ("Sep-2025", 103.3, 106.3, "FY2026", "Q2"),
    ("Oct-2025", 103.5, 107.5, "FY2026", "Q3"), ("Nov-2025", 102.5, 112.6, "FY2026", "Q3"),
    ("Dec-2025", 104.6, 107.7, "FY2026", "Q3"), ("Jan-2026", 105.5, 109.8, "FY2026", "Q4"),
    ("Feb-2026", 98.8, 118.4, "FY2026", "Q4"), ("Mar-2026", 104.5, 138.6, "FY2026", "Q4"),
    ("Apr-2026", 106.1, 130.9, "FY2027", "Q1"), ("May-2026", 108.6, 136.9, "FY2027", "Q1"),
    ("Jun-2026", 100.8, 138.5, "FY2027", "Q1"),
]

IRON_ORE_QTR_HEADERS = [
    "Quarter", "Iron Ore 62% Fe Fines, CFR China (US$/dmt)",
    "Australian THERMAL Coal (US$/mt)", "Period Type", "Source ID", "Methodology", "Confidence",
]

IRON_ORE_QTR = [
    ("Q1 FY2026", 95.50, 104.00), ("Q2 FY2026", 100.10, 110.47),
    ("Q3 FY2026", 103.53, 109.27), ("Q4 FY2026", 102.93, 122.27),
    ("Q1 FY2027", 105.17, 135.43),
]

IRON_ORE_DOMESTIC_HEADERS = [
    "Series", "Producer / Assessor", "Grade / Specification", "Basis", "Unit", "Currency",
    "Effective / Observation Date", "Value", "Period Type", "Source ID", "Source Tier",
    "Confidence", "Notes",
]

IRON_ORE_DOMESTIC = [
    ("NMDC notified price", "NMDC Limited", "Bailadila Lump, 65.5% Fe, 10-40 mm",
     "Ex-mine; EXCLUDES royalty, DMF, NMET, cess and other levies", "Rs/t", "INR",
     "05-Apr-2026", 5300, "Point-in-time", "S23", 1, "High",
     "Price revision filed with the exchanges; represented an increase of up to 11.1%."),
    ("NMDC notified price", "NMDC Limited", "Bailadila Fines, 64% Fe, minus 10 mm",
     "Ex-mine; EXCLUDES royalty, DMF, NMET, cess and other levies", "Rs/t", "INR",
     "05-Apr-2026", 4500, "Point-in-time", "S23", 1, "High", "As above."),
    ("NMDC notified price", "NMDC Limited", "Bailadila Lump, 65.5% Fe, 10-40 mm",
     "Ex-mine; excludes levies", "Rs/t", "INR", "06-May-2026", 5500, "Point-in-time",
     "S23", 1, "High", "Increase of Rs 200/t on both grades."),
    ("NMDC notified price", "NMDC Limited", "Bailadila Fines, 64% Fe, minus 10 mm",
     "Ex-mine; excludes levies", "Rs/t", "INR", "06-May-2026", 4700, "Point-in-time",
     "S23", 1, "High", "Increase of Rs 200/t."),
    ("NMDC notified price", "NMDC Limited", "Bailadila Lump, 65.5% Fe, 10-40 mm",
     "Ex-mine; excludes levies", "Rs/t", "INR", "10-Jul-2026", 5450, "Point-in-time",
     "S23", 1, "High",
     "Reduction of Rs 50/t on lump; fines held at Rs 4,700/t. Reported as a second successive reduction. Equivalent prices INCLUDING taxes and charges were reported as Rs 6,745/t for lump and Rs 5,815/t for fines - a ~24% gross-up over the ex-mine notified price, which is material for any landed-cost build."),
    ("NMDC notified price", "NMDC Limited", "Bailadila Fines, 64% Fe, minus 10 mm",
     "Ex-mine; excludes levies", "Rs/t", "INR", "10-Jul-2026", 4700, "Point-in-time",
     "S23", 1, "High", "Unchanged from 06-May-2026."),
    ("Ministry of Steel domestic assessment", "Ministry of Steel / JPC", "Iron ore lumps",
     "Market assessment as published by the Ministry", "Rs/t", "INR",
     "Mar-2026", 4800, "Month", "S01", 1, "High",
     "The Ministry's own published domestic iron ore lump price, described as showing a slight increase on the previous month. This differs from NMDC's notified Rs 5,300/t for a specific 65.5% Fe grade - a construct difference, not a conflict. The Ministry has published this series monthly since Apr-2023."),
    ("Regional domestic spot", "IDBI Capital", "Iron ore, Odisha", "Broker weekly assessment",
     "Rs/t", "INR", "Week ended 22-Jun-2026", 6900, "Point-in-time", "S21", 2, "Medium",
     "Down 1.4% week on week and down 14.8% year on year; equal to the 52-week low (range Rs 6,900-8,100/t). Note IDBI's own Mar-2026 monthly note put 'domestic iron ore' at Rs 7,850/t, consistent with a ~12% correction through Q1FY2027, though grade and location definitions differ between the two quotes."),
    ("Regional domestic spot", "IDBI Capital", "Iron ore pellets, ex-Barbil",
     "Broker weekly assessment", "Rs/t", "INR", "Week ended 22-Jun-2026", 8000,
     "Point-in-time", "S21", 2, "Medium",
     "52-week range Rs 7,900-9,250/t. The pellet-over-fines premium is the key margin driver for pellet-integrated producers such as Godawari Power and Shyam Metalics."),
    ("International spot", "IDBI Capital", "Iron ore, 62% Fe import fines",
     "Broker weekly assessment", "US$/t", "USD", "Week ended 22-Jun-2026", 96,
     "Point-in-time", "S21", 2, "Medium",
     "52-week range USD 88-107/t. Broadly consistent with the World Bank series (Jun-2026 monthly average USD 100.8/dmt), which is a reassuring independent cross-check on the international benchmark."),
    ("International spot", "IDBI Capital", "Iron ore pellets", "Broker weekly assessment",
     "US$/t", "USD", "Week ended 22-Jun-2026", 144, "Point-in-time", "S21", 2, "Medium",
     "Up 25.1% year on year; 52-week range USD 111-150/t."),
    ("Domestic assessment", "IDBI Capital", "Iron ore, domestic", "Broker monthly assessment",
     "Rs/t", "INR", "Mar-2026", 7850, "Month", "S20", 2, "Medium",
     "Flat month on month in Mar-2026 while the international benchmark rose 2.4% to USD 102/t."),
]

# ======================================================================================
# 9. COKING COAL PRICES
# ======================================================================================
COKING_COAL_HEADERS = [
    "Benchmark / Assessment", "Assessor", "Grade", "Delivery Term", "Unit", "Currency",
    "Observation Date / Period", "Value", "52-Week High", "52-Week Low",
    "Period Type", "Source ID", "Source Tier", "Confidence", "Notes",
]

COKING_COAL = [
    # --- Ministry of Steel / JPC series (PREFERRED)
    ("Premium HCC - Ministry of Steel / JPC assessment", "Ministry of Steel / JPC",
     "Hard coking coal (HCC)", "FOB Australia", "US$/t", "USD",
     "Oct-2023", 354, None, None, "Month", "S01", 1, "High",
     "Cycle PEAK. The Ministry publishes this series as a chart in its Monthly Economic Report with narrative annotation of key levels. PREFERRED SERIES for modelling Indian integrated-mill cost - see Conflict C01."),
    ("Premium HCC - Ministry of Steel / JPC assessment", "Ministry of Steel / JPC",
     "Hard coking coal (HCC)", "FOB Australia", "US$/t", "USD",
     "Mar-2025", 175, None, None, "Month", "S01", 1, "High",
     "Cycle TROUGH, described by the Ministry as 'a low of around US$175 per tonne'. This marks the start of FY2026 and explains the strong first-half FY2026 margins for BF-BOF producers."),
    ("Premium HCC - Ministry of Steel / JPC assessment", "Ministry of Steel / JPC",
     "Hard coking coal (HCC)", "FOB Australia", "US$/t", "USD",
     "Feb-2026", 246, None, None, "Month", "S01", 1, "High",
     "FY2026 high point, up ~41% from the Mar-2025 trough. The principal cost headwind behind Q4FY2026 margin commentary from Tata Steel, SAIL and Jindal Steel."),
    ("Premium HCC - Ministry of Steel / JPC assessment", "Ministry of Steel / JPC",
     "Hard coking coal (HCC)", "FOB Australia", "US$/t", "USD",
     "Mar-2026", 225, None, None, "Month", "S01", 1, "High",
     "Eased from the Feb-2026 high. LATEST Ministry-published observation available at the data cutoff."),
    # --- Broker series (cross-check, DO NOT BLEND)
    ("Coking coal - IDBI Capital tracker", "IDBI Capital", "Coking coal (grade not specified)",
     "FOB Australia", "US$/t", "USD", "Feb-2026", 214, None, None, "Month", "S20", 2, "Low",
     "IMPLIED, not directly stated: the broker reports Mar-2026 at USD 186/t 'flat month on month after declining 13% month on month in Feb-2026', which implies ~USD 214/t for Feb-2026. Shown to evidence the magnitude of the divergence from the Ministry series - see Conflict C01. DO NOT BLEND with the Ministry rows above."),
    ("Coking coal - IDBI Capital tracker", "IDBI Capital", "Coking coal (grade not specified)",
     "FOB Australia", "US$/t", "USD", "Mar-2026", 186, None, None, "Month", "S20", 2, "Medium",
     "Flat month on month. Sits USD 39/t BELOW the Ministry's ~USD 225/t for the same month, which is why the two series cannot be treated as the same assessment."),
    ("Coking coal - IDBI Capital tracker", "IDBI Capital", "Coking coal (grade not specified)",
     "FOB Australia", "US$/t", "USD", "Week ended 22-Jun-2026", 200, 217, 135,
     "Point-in-time", "S21", 2, "Medium",
     "Flat week on week; up 14.3% month on month; up 48.1% year on year. DECISIVE EVIDENCE for Conflict C01: this series' 52-week HIGH of USD 217/t is BELOW the Ministry's Feb-2026 level of USD 246/t, so they are demonstrably different assessments."),
    # --- Thermal coal reference (explicitly NOT coking coal)
    ("Australian THERMAL coal - energy cost reference ONLY", "World Bank",
     "THERMAL coal, 6000 kcal/kg - NOT coking coal", "FOB Newcastle (futures)",
     "US$/mt", "USD", "FY2026 average", 111.50, None, None, "Annual (FY)", "S17", 1, "High",
     "DO NOT USE AS A COKING COAL PROXY. This is thermal coal, priced off energy content, with no consistent relationship to metallurgical coal - see Conflict C15. Included solely as an input reference for captive power and for gas/coal-based DRI economics. Fiscal-year averages: FY2022 176.58, FY2023 343.40, FY2024 145.10, FY2025 131.89, FY2026 111.50. Calendar 2022 averaged USD 344.90/mt at the energy-crisis peak."),
    ("Australian THERMAL coal - energy cost reference ONLY", "World Bank",
     "THERMAL coal, 6000 kcal/kg - NOT coking coal", "FOB Newcastle (futures)",
     "US$/mt", "USD", "Q1 FY2027 average", 135.43, None, None, "Quarter", "S17", 1, "High",
     "Up 30% from the Q1FY2026 average of USD 104.00/mt. Rising through Q4FY2026 and into Q1FY2027 (Mar-2026 USD 138.6/mt), consistent with the West Asia energy disruption flagged by Tata Steel and Jindal Stainless. A cost headwind for DRI-EAF and captive-power-heavy producers."),
]

COKING_COAL_NOTE = (
    "NO FISCAL-YEAR AVERAGE IS ASSERTED FOR COKING COAL. Neither of the two available series is "
    "published as a complete monthly history in the public domain, and the two disagree by up to "
    "US$40/t (see Conflict C01). Computing a fiscal-year average from four annotated chart points "
    "would create a false impression of precision. Obtain a licensed price-reporting-agency series "
    "- for example Platts Premium Low-Vol Hard Coking Coal FOB Australia - before using coking coal "
    "in a formal cost model or a DCF."
)

# ======================================================================================
# 10. DEMAND & CONSUMPTION
# ======================================================================================
DEMAND_HEADERS = [
    "Fiscal Year", "Total Finished Steel Consumption (Mt)", "Growth (%)",
    "Total Finished Steel Production (Mt)", "Crude Steel Production (Mt)",
    "Crude Steel Capacity (Mtpa)", "Consumption As % Of Production",
    "Period Type", "Source ID", "Confidence", "Notes",
]

# (fy, consumption, growth, finished_prod, crude_prod, capacity, note)
DEMAND = [
    ("FY2015", 76.99, 0.0390, 81.86, None, None, ""),
    ("FY2016", 81.52, 0.0588, 81.90, None, None, ""),
    ("FY2017", 84.04, 0.0309, 91.54, None, None, ""),
    ("FY2018", 90.71, 0.0794, 95.01, None, None, ""),
    ("FY2019", 98.71, 0.0882, 101.29, None, None, ""),
    ("FY2020", 100.17, 0.0147, 102.62, None, None, ""),
    ("FY2021", 94.89, -0.0527, 96.20, 103.545, None,
     "Only contraction in the series, caused by COVID-19 construction and industrial shutdowns."),
    ("FY2022", 105.75, 0.1144, 113.60, None, None, ""),
    ("FY2023", 119.89, 0.1337, 123.20, 127.20, None, ""),
    ("FY2024", 136.29, 0.1370, 139.15, 144.30, 179.51, ""),
    ("FY2025", 152.13, 0.1162, 146.69, 152.18, 200.33,
     "India became a net importer of total finished steel in FY2025 for the first time in several years."),
    ("FY2026", 164.19, 0.0790, 161.74, 168.42, 220.4,
     "Provisional. June-2026 JPC vintage; the April-2026 vintage shows consumption of 163.74 Mt and finished production of 160.94 Mt - see Conflict C03."),
]

DEMAND_SUPPLY_BALANCE_HEADERS = [
    "Line", "Unit", "FY2023", "FY2024", "FY2025", "FY2026 (P)",
    "Period Type", "Source ID", "Confidence", "Notes",
]

DEMAND_SUPPLY_BALANCE = [
    ("1. Crude steel production", "Mt", 127.20, 144.30, 152.18, 168.42,
     "Memo item; not part of the finished steel balance below."),
    ("2. Finished steel production", "Mt", 123.20, 139.15, 146.69, 160.94, ""),
    ("3. Finished steel imports", "Mt", 6.02, 8.32, 9.55, 6.52, ""),
    ("4. Finished steel exports", "Mt", 6.72, 7.49, 4.86, 6.60, ""),
    ("5. Availability (2 + 3 - 4)", "Mt", 122.50, 139.99, 151.38, 160.86, ""),
    ("6. Variation in stock", "Mt", 2.61, 3.69, -0.75, -2.88,
     "NEGATIVE in FY2025 and FY2026, i.e. inventory was DRAWN DOWN, which added to apparent consumption. The FY2026 destock of 2.88 Mt is ~1.8% of consumption - a genuine demand signal, and it flatters the reported consumption growth."),
    ("7. Apparent finished steel consumption (5 - 6)", "Mt", 119.89, 136.29, 152.13, 163.74,
     "April-2026 JPC vintage, internally consistent with lines 2 to 6. The later June-2026 vintage shows 164.19 Mt - see Conflict C03."),
]

PER_CAPITA_HEADERS = [
    "Geography", "Unit", "Period", "Per Capita Finished Steel Consumption",
    "Period Type", "Source ID", "Confidence", "Notes",
]

PER_CAPITA = [
    ("India", "kg", "FY2026", 115.7, "Annual (FY)", "S03", "High",
     "JPC. The Mar-2026 Ministry vintage stated 115 kg - see Conflict C05. India remains at roughly half the world average, which is the quantitative core of the long-run India steel demand thesis."),
    ("World", "kg", "CY2024", 215.0, "Calendar Year", "S03", "High",
     "World Steel Association, Steel Statistical Yearbook 2025. CALENDAR YEAR - not directly comparable to India's fiscal-year figure, though the gap is far too large to be explained by the period difference."),
    ("China", "kg", "CY2024", 604.0, "Calendar Year", "S03", "High",
     "World Steel Association, SSY-2025. China consumes ~5.2x India per capita. Convergence to even half the Chinese level would imply roughly a tripling of Indian demand."),
]

QUARTERLY_DEMAND_HEADERS = [
    "Period", "Finished Steel Consumption (Mt)", "Growth YoY (%)",
    "Finished Steel Production (Mt)", "Production Growth YoY (%)",
    "Period Type", "Source ID", "Confidence", "Notes",
]

QUARTERLY_DEMAND = [
    ("Q1 FY2027 (Apr-Jun 2026)", 41.57, 0.083, 41.00, 0.059, "Quarter", "S05", "High",
     "Consumption growth of 8.3% OUTPACED production growth of 5.9%, indicating continued domestic demand strength and implying either further destocking or higher net imports in Q1FY2027. This is the most recent industry datapoint available at the data cutoff."),
]

# ======================================================================================
# 11. IMPORTS & EXPORTS
# ======================================================================================
TRADE_ANNUAL_HEADERS = [
    "Fiscal Year", "Finished Steel Imports (Mt)", "Finished Steel Exports (Mt)",
    "Net Trade (Mt)", "Semis Exports ('000 t)", "Pig Iron Imports ('000 t)",
    "Pig Iron Exports ('000 t)", "Import Value (Rs crore)",
    "Total Steel Export Value (Rs crore)", "Period Type", "Source ID", "Confidence", "Notes",
]

# (fy, imp_mt, exp_mt, semis_kt, pigimp_kt, pigexp_kt, impval, expval, note)
TRADE_ANNUAL = [
    ("FY2015", 9.320, 5.595, 640, 23, 540, 44893, 31283, ""),
    ("FY2016", 11.711, 4.079, 639, 22, 297, 45044, 24083,
     "Peak import year of the last decade; triggered the 2015-16 safeguard duty and minimum import price measures."),
    ("FY2017", 7.224, 8.242, 1192, 34, 387, 34104, 38182, ""),
    ("FY2018", 7.483, 9.620, 1994, 16, 518, 39484, 52812, ""),
    ("FY2019", 7.835, 6.361, 2183, 67, 319, 49317, 40900, ""),
    ("FY2020", 6.768, 8.355, 2819, 11, 422, 44683, 45102, ""),
    ("FY2021", 4.752, 10.784, 6553, 9, 1099, 32154, 67132,
     "Record export year as India supplied a world short of Chinese material during COVID-19 disruption."),
    ("FY2022", 4.669, 13.494, 4878, 26, 1213, 46298, 122222,
     "All-time high finished steel exports of 13.49 Mt and export value of Rs 1,22,222 cr."),
    ("FY2023", 6.022, 6.716, 1621, 118, 629, 64454, 65117,
     "Exports halved after the 15% export duty was imposed in May-2022; the duty was withdrawn on 19-Nov-2022."),
    ("FY2024", 8.320, 7.487, 1055, 366, 385, 68193, 64634, ""),
    ("FY2025", 9.551, 4.858, 1403, 326, 287, 80737, 46870,
     "India was a NET IMPORTER of total finished steel. Imports at a decade high in value terms (Rs 80,737 cr) with exports at a decade low in volume. This is the evidential basis for the DGTR safeguard investigation."),
    ("FY2026", 6.524, 6.602, 1721, 182, 613, 61407, 56762,
     "Provisional. India returned to NET EXPORTER status by a narrow 0.08 Mt. Imports fell 31.7% and exports rose 35.9% - the clearest measurable effect of the safeguard duty."),
]

TRADE_PRODUCT_HEADERS = [
    "Direction", "Product", "Unit", "FY2025", "FY2026 (P)", "Change (%)",
    "Share Of FY2026 Total (%)", "Period Type", "Source ID", "Confidence", "Notes",
]

TRADE_PRODUCT = [
    ("Imports", "HR coil / strip", "Mt", 3.96, 2.31, -0.416, None,
     "Annual (FY)", "S01", "High",
     "The single largest swing factor: HR coil imports fell 1.65 Mt, accounting for 54% of the total 3.03 Mt fall in imports. HR coil is squarely within the safeguard duty's tariff headings."),
    ("Imports", "CR coil / sheets", "Mt", 1.56, 1.17, -0.249, None,
     "Annual (FY)", "S01", "High", ""),
    ("Imports", "GP / GC sheets and coil", "Mt", 1.36, 1.01, -0.255, None,
     "Annual (FY)", "S01", "High", ""),
    ("Imports", "Plates", "Mt", 1.18, 0.67, -0.433, None, "Annual (FY)", "S01", "High",
     "Largest percentage fall among the top five imported products."),
    ("Imports", "Electrical sheets", "Mt", 0.49, 0.49, 0.003, None,
     "Annual (FY)", "S01", "High",
     "ESSENTIALLY FLAT while every other major import category fell sharply. Electrical steel - particularly CRGO - is a genuine domestic capability gap, which is why it is a PLI Specialty Steel category and why Budget 2026-27 extended customs duty exemptions on CRGO inputs to 31-Mar-2028."),
    ("Imports", "Others", "Mt", 1.00, 0.86, -0.134, None, "Annual (FY)", "S01", "High", ""),
    ("Imports", "TOTAL", "Mt", 9.55, 6.52, -0.317, 1.0, "Annual (FY)", "S01", "High", ""),
    ("Exports", "HR coil / strip", "Mt", 1.09, 2.53, 1.314, None,
     "Annual (FY)", "S01", "High",
     "MORE THAN DOUBLED. HR coil went from 1.09 Mt exported to 2.53 Mt while HR coil imports fell from 3.96 Mt to 2.31 Mt - a 3.09 Mt swing in the net trade position of a single product category in one year."),
    ("Exports", "Pipes", "Mt", 0.71, 1.05, 0.481, None, "Annual (FY)", "S01", "High",
     "Value-added export growth, consistent with APL Apollo's revenue trajectory."),
    ("Exports", "GP / GC sheets and coil", "Mt", 1.14, 0.91, -0.201, None,
     "Annual (FY)", "S01", "High", "The only top-five export category to decline."),
    ("Exports", "Bars and rods", "Mt", 0.62, 0.75, 0.202, None, "Annual (FY)", "S01", "High", ""),
    ("Exports", "CR coil / sheets", "Mt", 0.59, 0.59, -0.010, None,
     "Annual (FY)", "S01", "High", ""),
    ("Exports", "Others", "Mt", 0.70, 0.77, 0.103, None, "Annual (FY)", "S01", "High", ""),
    ("Exports", "TOTAL", "Mt", 4.86, 6.60, 0.359, 1.0, "Annual (FY)", "S01", "High", ""),
]

TRADE_COUNTRY_HEADERS = [
    "Direction", "Country", "Unit", "FY2025 Share", "FY2026 Share (P)",
    "Period Type", "Source ID", "Confidence", "Notes",
]

TRADE_COUNTRY = [
    ("Imports", "Korea", "%", 0.294, 0.354, "Annual (FY)", "S01", "High",
     "Largest source in both years and GAINED share as total imports fell, i.e. Korean material proved most resilient to the safeguard duty. Korea has a free trade agreement with India (CEPA), which is central to the ongoing policy debate."),
    ("Imports", "China", "%", 0.265, 0.235, "Annual (FY)", "S01", "High",
     "Lost share as the safeguard duty bit. Chinese crude steel output fell 4.4% in CY2025, also reducing exportable surplus."),
    ("Imports", "Japan", "%", 0.211, 0.202, "Annual (FY)", "S01", "High",
     "Broadly held share. Japan also has a CEPA with India."),
    ("Imports", "Vietnam", "%", 0.073, NA, "Annual (FY)", "S01", "High",
     "Disclosed separately in FY2025 but folded into 'Others' in the FY2026 full-year disclosure. Vietnam was 6.3% of March-2026 monthly imports, so it remains a material source."),
    ("Imports", "Others", "%", 0.156, 0.210, "Annual (FY)", "S01", "High",
     "Korea, China and Japan together were 79.1% of FY2026 imports, against 77.0% in FY2025 - import sourcing CONCENTRATED rather than diversified under the safeguard duty."),
    ("Exports", "Italy", "%", 0.146, 0.162, "Annual (FY)", "S01", "High",
     "Largest destination in both years. EU exposure makes the EU CBAM, in its definitive regime from 01-Jan-2026, a direct commercial issue for Indian exporters."),
    ("Exports", "Vietnam", "%", NA, 0.117, "Annual (FY)", "S01", "High",
     "New entrant to the top destinations in FY2026 and 32.3% of March-2026 monthly exports - a sharp, recent redirection of Indian material into South East Asia."),
    ("Exports", "Belgium", "%", 0.112, 0.109, "Annual (FY)", "S01", "High", ""),
    ("Exports", "UAE", "%", 0.101, 0.074, "Annual (FY)", "S01", "High", ""),
    ("Exports", "Spain", "%", 0.086, 0.073, "Annual (FY)", "S01", "High", ""),
    ("Exports", "Nepal", "%", 0.110, 0.065, "Annual (FY)", "S01", "High",
     "Share roughly halved."),
    ("Exports", "Taiwan", "%", NA, 0.057, "Annual (FY)", "S01", "High",
     "New entrant to the top destinations in FY2026."),
    ("Exports", "United Kingdom", "%", 0.068, NA, "Annual (FY)", "S01", "High",
     "Disclosed in FY2025, folded into 'Others' in FY2026. UK import quota changes announced Mar-2026 are relevant."),
    ("Exports", "Others", "%", 0.377, 0.343, "Annual (FY)", "S01", "High",
     "EU destinations (Italy, Belgium, Spain) alone were 34.4% of FY2026 exports, which is the scale of India's CBAM exposure."),
]

TRADE_MONTHLY_HEADERS = [
    "Month", "Exports (Mt)", "Imports (Mt)", "Net Trade (Mt)", "Period Type",
    "Source ID", "Confidence", "Notes",
]

TRADE_MONTHLY = [
    ("Mar-2025", 0.45, 0.58, -0.13, "Month", "S01", "High", "Net importer."),
    ("Feb-2026", 0.54, 0.56, -0.02, "Month", "S01", "High", "Marginally net importer."),
    ("Mar-2026", 0.58, 0.52, 0.06, "Month", "S01", "High",
     "Net exporter. Exports up 29.1% year on year; imports down 9.5%."),
]

TRADE_NARRATIVE = (
    "India had been a NET IMPORTER of steel in every month since April 2024 with only five "
    "exceptions - June 2025, October 2025, November 2025, January 2026 and March 2026. For "
    "FY2026 as a whole India was a net EXPORTER, with exports exceeding imports by 0.08 Mt. "
    "The clustering of the net-export months in the second half of FY2026, after the definitive "
    "safeguard duty notification of 30 December 2025, is the clearest month-level evidence of "
    "the policy's trade effect. Source: S01 (Joint Plant Committee, provisional)."
)

# ======================================================================================
# 12. GOVERNMENT POLICIES
# ======================================================================================
POLICY_HEADERS = [
    "Policy ID", "Policy / Measure", "Category", "Issuing Authority", "Instrument",
    "Announced / Notified Date", "Effective From", "Effective To / Review",
    "Status At Cutoff", "Key Quantitative Terms", "Products / Scope Covered",
    "Assessed Sector Impact", "Source ID", "Confidence", "Notes",
]

POLICIES = [
    ("P01", "Provisional safeguard duty on steel flat products", "Trade remedy",
     "Ministry of Finance, Government of India", "Customs notification (safeguard duty)",
     "21-Apr-2025", "21-Apr-2025", "Lapsed on expiry of 200 days, ~07-Nov-2025",
     "Superseded by the definitive measure (P03)",
     "12% ad valorem for 200 days",
     "Five categories of non-alloy and alloy steel FLAT products, including hot rolled coils, sheets and plates",
     "POSITIVE for integrated flat producers; NEGATIVE for flat-steel consumers and downstream converters such as APL Apollo. Imposed amid fears of trade diversion into India following steep US duties on Chinese goods.",
     "S29", "High",
     "Imposed on an emergency basis. Its lapse in Nov-2025 coincides with domestic HRC falling to ~Rs 46,000/t in early Dec-2025."),

    ("P02", "DGTR final findings in the safeguard investigation", "Trade remedy",
     "Directorate General of Trade Remedies, Ministry of Commerce & Industry",
     "Final findings and recommendation", "18-Aug-2025", "Recommendation only",
     "Recommendation adopted via P03", "Implemented",
     "Recommended a three-year duty tapering 12% / 11.5% / 11%",
     "Non-alloy and alloy steel flat products",
     "Set the template for the definitive measure. DGTR recorded 'a recent, sudden, sharp and significant increase in imports of the product under consideration'. Industry was divided, with primary producers in favour and MSME user industries opposed.",
     "S28", "High",
     "The taper is the mechanism by which protection is designed to withdraw as domestic capacity ramps up."),

    ("P03", "Definitive safeguard duty on steel flat products", "Trade remedy",
     "Ministry of Finance, Government of India", "Customs notification (safeguard duty)",
     "30-Dec-2025", "Year 1 dated from 21-Apr-2025", "20-Apr-2028, with a mid-term review provision",
     "IN FORCE - year 2 rate of 11.5% applies from 21-Apr-2026",
     "12% for 21-Apr-2025 to 20-Apr-2026; 11.5% for 21-Apr-2026 to 20-Apr-2027; 11% for 21-Apr-2027 to 20-Apr-2028",
     "Non-alloy and alloy steel flat products under tariff headings 7208, 7209, 7210, 7211, 7212, 7225 and 7226",
     "THE SINGLE MOST IMPORTANT POLICY DRIVER OF FY2026 EARNINGS. Domestic HRC rose from ~Rs 46,000/t in early Dec-2025 to Rs 57,700/t by end-Mar-2026, a ~25% move. FY2026 imports fell 31.7% and exports rose 35.9%, returning India to net exporter status. The 11.5% year-2 rate and the mid-term review are the key FY2027 monitorables.",
     "S27, S25", "High",
     "Scope independently confirmed by the Ministry of Steel in Rajya Sabha USQ 3350 (20-Mar-2026). Note the year-1 window back-fills the provisional period, so the levy appears continuous on the face of the notification despite an effective interruption - see Conflict C08. Covers FLAT products only, which is why longs outperformed flats on price."),

    ("P04", "Steel and Steel Products (Quality Control) Order, 2024", "Standards / non-tariff",
     "Ministry of Steel and Bureau of Indian Standards", "Quality Control Order",
     "2024", "2024", "Ongoing, subject to amendment", "In force, materially amended",
     "Mandatory BIS certification for covered steel and steel products",
     "Steel and steel products specified in the schedules to the Order",
     "A non-tariff barrier that operates alongside the safeguard duty. The Ministry of Steel describes it as 'banning sub-standard/defective steel products in domestic market as well as imports'. Its progressive dilution through FY2026 and FY2027 partially offsets the protection delivered by the safeguard duty.",
     "S25, S32", "High",
     "The governing instrument for the QCO regime. Subsequent measures P05 to P08 successively narrowed or deferred its application."),

    ("P05", "Extension of exemption from mandatory quality-norm compliance",
     "Standards / non-tariff", "Ministry of Steel", "Executive order",
     "~20-Nov-2025", "~20-Nov-2025", "Extended to March 2026", "Superseded by later measures",
     "Exemption extended for certain steel and stainless steel grades; separately, steel grades not covered by any QCO no longer require clarification or an NOC from the Ministry of Steel",
     "Certain specified steel and stainless steel grades",
     "MIXED. Eased input availability for downstream users and removed a significant administrative bottleneck on imports. Negative at the margin for domestic producers of the exempted grades. Motivated by preserving availability of critical steel products while domestic capacity ramps up.",
     "S32", "Medium",
     "For QCO-covered grades importers must still ensure the overseas manufacturer holds a BIS licence for those grades."),

    ("P06", "BIS amendments deferring enforcement for specified steel products",
     "Standards / non-tariff", "Bureau of Indian Standards / Ministry of Steel",
     "Amendment to the QCO", "Feb-2026", "Date of publication of the notification",
     "Deferred by three years from publication", "In force",
     "Three-year deferral of BIS enforcement",
     "Mild steel used in metal arc welding electrodes; steels for die blocks used in drop forging",
     "Narrow in scope. Relief for specific downstream engineering users; immaterial at sector level.",
     "S32", "Medium", "Made under the Steel and Steel Products (Quality Control) Order, 2024."),

    ("P07", "Suspension of the Quality Control Order for stainless steel products",
     "Standards / non-tariff", "Ministry of Steel", "Executive order",
     "Order dated 27-Apr-2026", "27-Apr-2026", "Not specified - described as temporary",
     "IN FORCE - active FY2027 risk",
     "Suspension of the requirement to use only BIS-certified items for various stainless steel products",
     "Various stainless steel products",
     "MATERIALLY NEGATIVE for domestic stainless producers, principally Jindal Stainless. Stated intent was to ease the compliance burden on domestic producers, especially MSMEs. Stainless-sector MSMEs have since flagged an import surge and urged the Government to reintroduce the Order.",
     "S32, S42, S12", "High",
     "CORROBORATED BY A PRIMARY ISSUER SOURCE: on Jindal Stainless's exchange-filed 4QFY26 earnings call the Managing Director called the suspension 'a matter of concern and a discouraging setback for quality-focused domestic industry players', noting inferior imported material continuing to enter India at scale. A rare instance where a policy risk is confirmed in an issuer's own filed disclosure."),

    ("P08", "Transition Facilitation (Quality Control) Order, 2026", "Standards / non-tariff",
     "Ministry of Commerce & Industry", "Quality Control Order",
     "25-Jun-2026", "2026", "Five-year transition window", "IN FORCE",
     "Alternative risk-based compliance mechanism; simplified registration-based BIS certification in place of the conventional inspection-intensive process; easing across several sectors for five years",
     "Multiple sectors, including steel",
     "Structurally reduces the effectiveness of QCOs as a non-tariff barrier for the next five years, while lowering compliance cost for importers and for smaller domestic manufacturers. Directionally NEGATIVE for incumbent large domestic producers that had relied on QCO-based protection.",
     "S33", "Medium",
     "Described by the Ministry as facilitating 'a smooth transition for industry while maintaining quality assurance and consumer protection'."),

    ("P09", "Production Linked Incentive Scheme for Specialty Steel", "Industrial policy",
     "Ministry of Steel", "Central sector scheme", "Original scheme 2021; PLI 1.1 launched 06-Jan-2025",
     "Production period FY2026 to FY2030 for PLI 1.1", "FY2030 / FY2031 depending on round",
     "IN FORCE across three rounds",
     "Budgetary outlay Rs 6,322 crore. Cumulative across all three rounds as at 24-Mar-2026: 171 projects from 104 companies; committed investment Rs 55,993 crore; investment realised Rs 23,827 crore (42.6%); specialty steel capacity created 24 Mt; import substitution through incremental production valued at Rs 6,000 crore",
     "Five categories: coated/plated steel products; high strength and wear-resistant steel; specialty rails; alloy steel products and steel wires; electrical steel",
     "The principal industrial-policy lever aimed at India's value-added and import-dependent segments, especially electrical steel where imports were flat at 0.49 Mt in FY2026 while all other categories fell. Ministry expects specialty steel production to reach 42 Mt by end-FY2027.",
     "S24, S45, S44", "High",
     "Round detail: Round 1 44 projects / 19 companies; Round 2 42 projects / 25 companies; Round 3 85 projects / 60 companies. PLI 1.1 reduced investment thresholds and allowed excess production to be carried forward for incentive claims, to attract applications in sub-categories that drew none in the first round. PLI 1.2 secured Rs 11,887 crore of commitments for 8.7 Mt of capacity by FY2031."),

    ("P10", "National Steel Policy 2017", "Industrial policy", "Ministry of Steel",
     "Policy framework", "2017", "2017", "FY2031 horizon", "IN FORCE",
     "Target of 300 Mtpa crude steel capacity by FY2031, with per capita consumption and raw material security objectives",
     "Whole sector",
     "The anchor long-term planning document. Against the 300 Mtpa target, capacity stood at 220.4 Mt in FY2026, so ~80 Mtpa must be added in five years - roughly 16 Mtpa a year, against ~20 Mtpa actually added in each of FY2025 and FY2026. The target is therefore achievable on current run-rate, which is a materially more constructive read than the market generally assumes.",
     "S03, S02", "High",
     "Reducing the carbon footprint of the industry is also stated as an objective of the Policy."),

    ("P11", "Domestically Manufactured Iron & Steel Products (DMI&SP) Policy",
     "Public procurement", "Ministry of Steel", "Procurement policy",
     "In force during the period", "In force", "Ongoing", "IN FORCE",
     "Preference for domestically manufactured iron and steel in Government procurement",
     "Government and public sector procurement of iron and steel",
     "Provides a protected demand base for domestic producers, particularly for SAIL and RINL given their long-products and rails exposure to railways and infrastructure procurement.",
     "S25", "High", "Listed by the Ministry of Steel as a measure to safeguard India's steel trade position."),

    ("P12", "Countervailing duty on welded stainless steel pipes and tubes",
     "Trade remedy", "Ministry of Finance / DGTR", "Countervailing duty",
     "In force during the period", "In force", "Ongoing", "IN FORCE",
     "Countervailing duty in place", "Welded stainless steel pipes and tubes from China and Vietnam",
     "Narrow but directly supportive of domestic stainless pipe and tube producers, partially offsetting the effect of the stainless QCO suspension (P07).",
     "S25", "High", "Confirmed by the Ministry of Steel in Rajya Sabha USQ 3350."),

    ("P13", "Union Budget 2026-27 customs measures for steel", "Fiscal / customs",
     "Ministry of Finance", "Finance Act / customs notifications", "Union Budget 2026-27",
     "FY2027", "Exemptions extended to 31-Mar-2028", "IN FORCE",
     "Nil Basic Customs Duty continued on ferro-nickel; BCD exemption extended to 31-Mar-2028 on ferrous scrap, on magnesium oxide coated cold rolled steel coils used to manufacture cold rolled grain oriented (CRGO) steel, and on specified goods for the manufacture of CRGO steel",
     "Ferro-nickel; ferrous scrap; CRGO steel inputs",
     "POSITIVE for stainless producers (ferro-nickel is a primary input for Jindal Stainless), for scrap-based EAF operators including Tata Steel's Ludhiana EAF, and for prospective CRGO manufacturers. Directly targets the electrical steel import gap that PLI also addresses.",
     "S25", "High",
     "The CRGO input exemptions are the fiscal complement to the PLI electrical steel category."),

    ("P14", "Steel export duty - imposition and withdrawal", "Trade / fiscal",
     "Ministry of Finance", "Customs notification", "Imposed May-2022; withdrawn 19-Nov-2022",
     "May-2022", "Withdrawn with effect from 19-Nov-2022",
     "WITHDRAWN - tariff on finished steel is NIL (0%)",
     "15% export duty on finished steel, reduced to nil",
     "Finished steel",
     "Historical but essential context for interpreting the FY2023 volume series. Exports halved from 13.49 Mt in FY2022 to 6.72 Mt in FY2023. The duty was scrapped following severe industry pushback and a 15-20% correction in domestic steel prices.",
     "S02", "High", "Effected by Notification 58/2022-Customs."),

    ("P15", "Taxonomy for Green Steel", "Decarbonisation", "Ministry of Steel",
     "Notified taxonomy", "23-Dec-2024", "23-Dec-2024", "Ongoing", "IN FORCE",
     "Green steel defined as emission intensity below 2.2 tCO2e per tonne of finished steel. Star rating: 5-star below 1.6; 4-star 1.6 to 2.0; 3-star 2.0 to 2.2; above 2.2 not eligible",
     "All steel produced in India",
     "Establishes the measurement framework that will underpin green procurement, any future carbon pricing, and export documentation for CBAM purposes. Described as the world's first national green steel taxonomy. A 37% government procurement share for green steel has been PROPOSED but is not an enacted mandate.",
     "S25, S46", "High",
     "Notification date of 23-Dec-2024 confirmed by the Ministry of Steel in Rajya Sabha USQ 3350. A Guideline for Green Steel Certification dated 30-Dec-2025 operationalises the 2.2 tCO2e/tfs threshold."),

    ("P16", "EU Carbon Border Adjustment Mechanism (CBAM) - definitive regime",
     "Decarbonisation / trade", "European Union", "EU regulation",
     "Definitive regime commenced", "01-Jan-2026", "Ongoing", "IN FORCE",
     "Carbon cost applied at the EU border on embedded emissions of covered goods including iron and steel",
     "Iron and steel imports into the EU",
     "DIRECTLY MATERIAL. EU destinations - Italy, Belgium and Spain - were 34.4% of India's FY2026 finished steel exports. Indian BF-BOF emission intensity is well above the EU benchmark, so CBAM is a structural margin headwind on the EU export book and a strategic driver of the green steel investment case. The Ministry of Steel's answer to Rajya Sabha USQ 3350 was specifically framed around CBAM readiness.",
     "S08, S25", "High",
     "CONFIRMED FROM A PRIMARY ISSUER SOURCE: Tata Steel's 4QFY2026 press release states that in Europe 'import safeguards and roll out of the Carbon Border Adjustment Mechanism from 1st January has improved pricing conditions' - i.e. in the near term CBAM has supported EU domestic prices, which benefits Tata Steel's own Netherlands operations even as it raises the barrier for India-origin exports."),

    ("P17", "UK steel import quota changes", "Trade remedy (overseas)",
     "United Kingdom", "Import quota regime", "Announced Mar-2026", "2026", "Not specified",
     "IN FORCE", "Changes to UK steel import quotas", "Steel imports into the UK",
     "Relevant to Tata Steel specifically through Tata Steel UK. Management expects the changes to 'bring greater balance to a market where demand conditions continue to be cause for concern'. Supportive of the halving of Tata Steel UK's EBITDA loss to GBP 217m in FY2026.",
     "S08", "High", "Sourced from Tata Steel's own FY2026 results release."),

    ("P18", "Netherlands environmental enforcement against Tata Steel Netherlands",
     "Environmental regulation (overseas)",
     "Dutch Environment Agency and the Province (Netherlands)",
     "Enforcement action and notices", "Letter issued 23-Apr-2026", "2026", "Ongoing",
     "ACTIVE - unresolved at the data cutoff",
     "More than EUR 20 million of penalties paid in FY2026 in relation to the coke and gas plants; stated intention to revoke operating permits and trigger early closure of those plants",
     "Tata Steel Netherlands, IJmuiden site",
     "SEVERE AND COMPANY-SPECIFIC. Tata Steel Netherlands' financial statements have been prepared taking into account a MATERIAL UNCERTAINTY RELATED TO GOING CONCERN, in discussion with its auditors. Many penalties relate to exceedances for which the company states no technically and operationally feasible best practice is currently available globally given the 40 to 50 year vintage of the coke ovens. Tata Steel is exploring all options including legal recourse. Separately, Dutch requirements on classification and disposal of steel slag now exceed EU standards.",
     "S08", "High",
     "This is the single largest identifiable downside risk in the coverage universe and is disclosed in the issuer's own results release. Any Tata Steel valuation must sensitise for impairment or closure of the Netherlands operation."),

    ("P19", "RINL revival support", "State support / restructuring",
     "Government of India", "Equity infusion and support package",
     "Package under implementation; a further infusion reported under consideration Apr-2026",
     "FY2025-FY2026", "Ongoing", "IN FORCE",
     "A further equity infusion of approximately Rs 8,097 crore was reported to be under consideration",
     "Rashtriya Ispat Nigam Limited",
     "TRANSFORMATIONAL FOR RINL AND MATERIAL FOR INDIA SUPPLY. FY2026 crude steel output rose 51% to 5.43 Mt and turnover 22% to over Rs 22,300 crore, with a return to monthly cash profit of Rs 54 crore in Jan-2026 against a cash loss of Rs 486 crore in Sep-2024. Roughly 1.85 Mt of incremental supply, equal to ~1.1% of India's FY2026 crude steel output, came from RINL's restart alone. The Andhra Pradesh government has stated there is no proposal to privatise RINL.",
     "S16", "Medium",
     "The Rs 8,097 crore figure rests on press reporting and should be treated as indicative until a formal announcement."),

    ("P20", "West Asia conflict - supply chain and energy disruption",
     "Geopolitical (not a policy, tracked here for completeness)",
     "Not applicable", "External event", "Escalated from Q4FY2026", "Q4FY2026", "Ongoing",
     "ACTIVE - flagged as continuing into FY2027",
     "Not quantified by issuers",
     "Energy, oil, freight, currency markets; availability of propane, LPG, natural gas and ammonia; shipping-lane routing",
     "SECTOR-WIDE COST AND LOGISTICS HEADWIND. Tata Steel states that developments in West Asia began to pressure supply chains and input costs in Q4FY2026 and that these pressures continue into FY2027, with calibrated mitigating actions under way. Jindal Stainless cites constrained availability of industrial gases critical to stainless steelmaking, plus route diversions, extended transit times and intermittent cargo delays. Also cited as a driver of the Mar-2026 spike in Chinese and Indian steel prices and of the rise in Australian thermal coal from USD 118/mt in Feb-2026 to USD 139/mt in Mar-2026.",
     "S08, S12, S20", "High",
     "Included because two separate issuers independently flagged it in filed FY2026 disclosures, making it the most consistently cited forward risk in the coverage universe."),
]

# ======================================================================================
# 13. INDUSTRY KPIs
# ======================================================================================
KPI_HEADERS = [
    "KPI ID", "KPI", "Category", "Unit", "FY2024", "FY2025", "FY2026",
    "Latest Available", "Latest Period", "Period Type", "Source ID",
    "Methodology", "Confidence", "Why This KPI Matters",
]

KPIS = [
    ("K01", "Crude steel production", "Supply", "Mt", 144.30, 152.18, 168.42, 168.42,
     "FY2026", "Annual (FY)", "S01", "As reported", "High",
     "The headline supply number for the world's second-largest producer and the fastest-growing at scale."),
    ("K02", "Crude steel capacity", "Supply", "Mtpa", 179.51, 200.33, 220.4, 220.4,
     "FY2026", "Annual (FY), year-end", "S02, S03", "As reported", "High",
     "The denominator for utilisation and the gauge of progress towards the 300 Mtpa National Steel Policy target."),
    ("K03", "Capacity utilisation (crude steel)", "Supply", "%", 0.8039, 0.7596, 0.7642, 0.7642,
     "FY2026", "Annual (FY)", "S01, S02", "Computed: production divided by year-end capacity",
     "High",
     "Fell from 80% to 76% as capacity outran production. On a year-end denominator this understates true utilisation in heavy commissioning years, but the direction is the point: India is building ahead of demand, which caps domestic pricing power."),
    ("K04", "Total finished steel production", "Supply", "Mt", 139.15, 146.69, 161.74, 161.74,
     "FY2026", "Annual (FY)", "S02", "As reported", "High",
     "The relevant supply measure for the finished steel balance and for market-share work."),
    ("K05", "Apparent finished steel consumption", "Demand", "Mt", 136.29, 152.13, 164.19, 41.57,
     "Q1 FY2027", "Annual (FY); latest is a quarter", "S02, S05", "As reported", "High",
     "The demand anchor. Grew at a 10.6% CAGR over FY2022-FY2026. Latest quarter shows 8.3% year-on-year growth."),
    ("K06", "Finished steel consumption growth", "Demand", "%", 0.1370, 0.1162, 0.0790, 0.083,
     "Q1 FY2027", "Annual (FY); latest is a quarter", "S02, S05", "As reported", "High",
     "Decelerating on the annual view from 13.7% to 7.9%, but the latest quarter re-accelerated to 8.3%. The most important single forward indicator in the sector."),
    ("K07", "Per capita finished steel consumption", "Demand", "kg", None, None, 115.7, 115.7,
     "FY2026", "Annual (FY)", "S03", "As reported", "High",
     "Roughly half the world average of 215 kg and about a fifth of China's 604 kg. The quantitative core of the multi-decade India demand thesis."),
    ("K08", "Sponge iron (DRI) production", "Supply mix", "Mt", 51.56, 55.76, 60.30, 60.30,
     "FY2026", "Annual (FY)", "S02", "As reported", "High",
     "India is the world's largest DRI producer. DRI-EAF is exposed to thermal coal, gas and scrap rather than coking coal, which materially changes the sector's cost sensitivity relative to a BF-BOF-dominated industry."),
    ("K09", "Pig iron production", "Supply mix", "Mt", 7.36, 8.33, 8.37, 8.37, "FY2026",
     "Annual (FY)", "S02", "As reported", "High",
     "Structurally declining in share. A read on merchant foundry and casting demand."),
    ("K10", "Finished steel imports", "Trade", "Mt", 8.320, 9.551, 6.524, 6.524, "FY2026",
     "Annual (FY)", "S02", "As reported", "High",
     "Down 31.7% in FY2026. The most direct measure of the safeguard duty's effect."),
    ("K11", "Finished steel exports", "Trade", "Mt", 7.487, 4.858, 6.602, 6.602, "FY2026",
     "Annual (FY)", "S02", "As reported", "High",
     "Up 35.9% in FY2026 from a decade low."),
    ("K12", "Net trade position (exports less imports)", "Trade", "Mt", -0.833, -4.693, 0.078,
     0.06, "Mar-2026", "Annual (FY); latest is a month", "S01, S02",
     "Computed from reported imports and exports", "High",
     "India swung from a 4.69 Mt net import position in FY2025 to a 0.08 Mt net export position in FY2026 - a 4.77 Mt swing in one year. This determines whether domestic prices set at import parity or export parity, which drives the entire domestic price deck."),
    ("K13", "Import penetration (imports as % of consumption)", "Trade", "%", 0.0610, 0.0628,
     0.0397, 0.0397, "FY2026", "Annual (FY)",
     "S02", "Computed: imports divided by consumption", "High",
     "Fell from 6.3% to 4.0%. Below 5% is generally consistent with domestic producers retaining pricing power."),
    ("K14", "Export intensity (exports as % of production)", "Trade", "%", 0.0538, 0.0331,
     0.0408, 0.0408, "FY2026", "Annual (FY)",
     "S02", "Computed: exports divided by finished steel production", "High",
     "India remains a fundamentally domestic market: only ~4% of finished output is exported, so global steel prices matter through the import channel rather than the export channel."),
    ("K15", "Iron ore price, 62% Fe fines CFR China", "Input cost", "US$/dmt", 119.91, 103.97,
     100.52, 105.17, "Q1 FY2027", "Annual (FY) average; latest is a quarter average",
     "S17", "Computed: average of World Bank monthly nominal series", "High",
     "Down for a third consecutive year, a persistent tailwind for non-integrated producers and a headwind for iron-ore-integrated names such as Godawari Power and NMDC."),
    ("K16", "Premium HCC coking coal, FOB Australia", "Input cost", "US$/t", None, None, None,
     225, "Mar-2026", "Month", "S01", "As reported (Ministry of Steel assessment)", "Medium",
     "NO FISCAL-YEAR AVERAGE IS ASSERTED - see Conflict C01. The available observations show a rise from a ~USD 175/t trough in Mar-2025 to ~USD 246/t in Feb-2026 before easing to ~USD 225/t in Mar-2026. This is the dominant swing factor in BF-BOF cost and the reason Q4FY2026 margin commentary was cost-focused across the integrated names."),
    ("K17", "Australian thermal coal, FOB Newcastle", "Input cost", "US$/mt", 145.10, 131.89,
     111.50, 135.43, "Q1 FY2027", "Annual (FY) average; latest is a quarter average",
     "S17", "Computed: average of World Bank monthly nominal series", "High",
     "THERMAL, not coking coal. Fell through FY2026 but rose 30% into Q1FY2027 versus Q1FY2026, a fresh headwind for DRI-EAF and captive-power-heavy producers."),
    ("K18", "Domestic HRC price", "Realisation", "Rs/t", None, None, 57700, 58200,
     "Week ended 22-Jun-2026", "Point-in-time", "S19, S21", "As reported", "Medium",
     "FY2026 column shows the ICRA end-March 2026 assessment. The 52-week range of Rs 45,700-59,600/t is a ~30% swing, which is the single largest source of FY2027 earnings uncertainty in the sector."),
    ("K19", "India-China HRC spread", "Realisation", "Rs/t", None, None, None, 9620,
     "Week ended 22-Jun-2026", "Point-in-time", "S21",
     "Computed: Indian HRC less China HRC FOB, both as converted to INR by the source", "Medium",
     "A ~20% premium. This spread is what the safeguard duty defends; if it widens materially further, pressure to dilute the duty at the mid-term review will grow."),
    ("K20", "Number of companies in coverage universe", "Coverage", "Number", 15, 15, 15, 15,
     "FY2026", "Point-in-time", "S30", "Computed", "High",
     "Seven Core Coverage and eight Supporting Coverage names, of which twelve are listed, one is an unlisted joint venture, one is an unlisted state-owned enterprise and two are delisted."),
    ("K21", "Aggregate coverage revenue (listed Indian names, consolidated)", "Coverage",
     "Rs crore", None, None, 693921, 693921, "FY2026", "Annual (FY)",
     "S30", "Computed: sum of FY2026 revenue from operations for the eleven listed names with data",
     "Medium",
     "Sum of Tata Steel, JSW Steel, SAIL, Jindal Steel, Jindal Stainless, Shyam Metalics, Godawari Power, APL Apollo, Mukand, Gallantt and Kirloskar Ferrous. Excludes AM/NS India (USD, calendar year), RINL, Uttam Galva and Electrosteel. Not an industry total - it double-counts nothing but omits the large secondary sector."),
    ("K22", "PLI Specialty Steel investment realised", "Policy", "Rs crore", None, None, 23827,
     23827, "24-Mar-2026", "Point-in-time", "S24", "As reported", "High",
     "Against Rs 55,993 crore committed, i.e. 42.6% realised. The gap is the pipeline of specialty steel capacity still to be built and is a direct read on import substitution in electrical and high-strength steels."),
    ("K23", "Safeguard duty rate in force", "Policy", "%", None, 0.12, 0.12, 0.115,
     "From 21-Apr-2026", "Point-in-time", "S27", "As reported", "High",
     "12% applied for year 1 (21-Apr-2025 to 20-Apr-2026); 11.5% applies for year 2 from 21-Apr-2026; 11% for year 3. A mid-term review is provided for and is the key policy monitorable for FY2027."),
    ("K24", "Green steel threshold", "Policy / ESG", "tCO2e/tfs", None, 2.2, 2.2, 2.2,
     "In force", "Point-in-time", "S25, S46", "As reported", "High",
     "Steel below 2.2 tCO2e per tonne of finished steel qualifies as green, with 5-star below 1.6, 4-star 1.6-2.0 and 3-star 2.0-2.2. Sets the measurement basis for future green procurement and for CBAM documentation."),
]
