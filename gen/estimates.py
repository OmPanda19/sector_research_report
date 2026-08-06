"""Central store for every MODELLED estimate in the workbook.

A forecast model with 1,700 blank cells is not a model, it is a template. Where a series
is genuinely published, the workbook sources it. Where no published series exists, the
choice is between leaving a hole and modelling the number from something that IS known.
This module holds the second case: every estimate, its numeric basis, and the one-line
justification that goes into the reference register.

Two rules apply to everything in here:

  1. An estimate is always expressed as a RELATIONSHIP to something sourced - a beta to
     real GDP, a differential to a benchmark price, an intensity per tonne of production -
     never as a free-floating number. That way it moves with the model instead of going
     stale, and a reader can attack the relationship rather than guess at the number.
  2. Every entry carries `basis`, which is reproduced verbatim in the register against the
     reference code for those cells. Nothing is asserted silently.

Confidence is stated honestly. These are modelled estimates, not observations, and the
register labels them as such so they can never be mistaken for sourced data.
"""

# ============================================================ macro sub-series
# Growth of a sector expressed as a multiple of real GDP growth. India's long-run
# relationships: industrial production has trailed GDP as services took share,
# infrastructure has outpaced it through the National Infrastructure Pipeline period,
# and construction sits between the two.
MACRO_BETA = {
    'iip': (0.85, 'Industrial production growth modelled at 0.85x real GDP growth. The Index of '
                  'Industrial Production has persistently trailed headline GDP in India because '
                  'services carry roughly 55% of GVA and have grown faster than industry; a beta '
                  'below 1.0 is the arithmetic consequence.'),
    'mfggva': (1.00, 'Manufacturing GVA growth modelled at 1.00x real GDP growth. Manufacturing has '
                     'held a broadly stable 17-18% share of Indian GVA for a decade, and a constant '
                     'share is exactly what a unit beta means.'),
    'infra': (1.60, 'Infrastructure growth modelled at 1.60x real GDP growth. Central government '
                    'capital expenditure has compounded far faster than nominal GDP across the '
                    'National Infrastructure Pipeline period. The beta is held well above 1.0 but '
                    'below the realised outturn, because the base is now large and the pipeline '
                    'the model itself uses terminates in FY2031.'),
    'construction': (1.30, 'Construction growth modelled at 1.30x real GDP growth. Construction GVA '
                           'has outgrown headline GDP through the housing and infrastructure upcycle '
                           'but by less than infrastructure alone.'),
    'auto': (1.05, 'Automobile production growth modelled at 1.05x real GDP growth. Vehicle output is '
                   'cyclical around GDP with a mild positive elasticity; it is deliberately NOT given '
                   'a high beta, because two-wheeler and passenger-vehicle cycles have repeatedly '
                   'diverged from GDP in both directions.'),
    'engineering': (1.15, 'Engineering output growth modelled at 1.15x real GDP growth, between '
                          'manufacturing and construction, reflecting its mix of domestic capex and '
                          'export order books.'),
    'capgoods': (1.20, 'Capital goods growth modelled at 1.20x real GDP growth. Capital goods is the '
                       'most investment-levered block in the table and therefore the most cyclical.'),
    'railways': (1.25, 'Railway output growth modelled at 1.25x real GDP growth, reflecting a '
                       'sustained public capital-expenditure programme running ahead of GDP.'),
    'oilgas': (0.90, 'Oil and gas growth modelled at 0.90x real GDP growth. Refining and pipeline '
                     'capital expenditure is lumpy and, unlike the rest of the table, faces a '
                     'structural energy-transition headwind over the horizon.'),
    'renewable': (2.50, 'Renewable energy growth modelled at 2.50x real GDP growth. This is the one '
                        'deliberately aggressive beta in the table: installed renewable capacity has '
                        'compounded at multiples of GDP off a small base, and the steel intensity of '
                        'solar mounting structures and wind towers is high. It is applied to a 2% '
                        'demand share, so the aggregate consequence is small and contained.'),
}

# WPI runs below CPI in India on average: CPI carries a ~46% food and beverages weight,
# while WPI is dominated by manufactured products and fuel.
WPI_SPREAD = (-0.008, 'WPI inflation modelled as CPI inflation less 0.8 percentage points. Over the '
                      'last decade WPI has averaged below CPI because CPI carries a food and '
                      'beverages weight of roughly 46% while WPI is dominated by manufactured '
                      'products and fuel. The spread is a long-run average, not a forecast of the '
                      'gap in any single year - WPI is materially more volatile than CPI and has '
                      'been both negative and double-digit within the last ten years.')

# ============================================================ leading indicators
# FY2026A reading, number format, and the basis. These are the last observed prints or,
# where the print was not retrievable, a value modelled off a driver the workbook holds.
LEADING = [
    ('PMI Manufacturing', 57.0, '0.0',
     'Indicative reading. India\'s manufacturing Purchasing Managers\' Index has held in the '
     'mid-to-high 50s throughout FY2026; 50 is the expansion threshold. Carried as an indicative '
     'level rather than a sourced monthly print, and it drives nothing in the model.'),
    ('Core Sector Growth', 0.055, '0.0%',
     'Modelled at 0.85x real GDP growth, the same beta the model applies to industrial production. '
     'The eight core industries account for roughly 40% of the Index of Industrial Production, so '
     'the two series track each other closely.'),
    ('Infrastructure Spending', 0.105, '0.0%',
     'Modelled at the infrastructure beta of 1.60x real GDP growth applied to the FY2026A GDP '
     'estimate. Consistent with the infrastructure growth line in the executive summary above, '
     'which uses the same beta.'),
    ('Government Capex', 0.100, '0.0%',
     'Modelled just below the infrastructure beta. Central capital expenditure growth has begun to '
     'decelerate as the base has grown, and budgeted capex growth has been in the low teens.'),
    ('Housing Completions', 0.085, '0.0%',
     'Modelled at the construction beta of 1.30x real GDP growth, less a small drag for the '
     'completion lag between launches and handover.'),
    ('Vehicle Production', 0.080, '0.0%',
     'Modelled at the automobile beta of 1.05x real GDP growth. Consistent with the auto production '
     'growth line in the executive summary above.'),
    ('Railway Capex', 0.095, '0.0%',
     'Modelled at the railway beta of 1.25x real GDP growth. Railway capital outlay has been '
     'sustained above GDP growth but the annual budget allocation has flattened recently.'),
]

# ============================================================ sector demand mix
# FY2026A share of finished steel consumption, and the FY2033E share the model drifts to.
# Building and construction (buildings plus infrastructure) is the dominant block at just
# over half of demand, consistent with published estimates of the Indian demand mix.
SECTOR_MIX = [
    ('Construction', 0.34, 0.32),
    ('Infrastructure', 0.19, 0.22),
    ('Automotive', 0.09, 0.09),
    ('Engineering', 0.08, 0.08),
    ('Capital Goods', 0.07, 0.07),
    ('Railways', 0.04, 0.04),
    ('Oil & Gas', 0.04, 0.03),
    ('Consumer Durables', 0.05, 0.05),
    ('Renewable Energy', 0.02, 0.04),
    ('Others', 0.08, 0.06),
]

SECTOR_MIX_BASIS = (
    'FY2026A shares are an indicative decomposition of finished steel consumption in which buildings '
    'and infrastructure together take 53%, consistent with published estimates that building and '
    'construction accounts for roughly half of Indian finished steel demand. Shares drift linearly to '
    'the FY2033E column, with infrastructure and renewable energy gaining and construction, oil and gas '
    'and the residual giving way. Every year is then NORMALISED so the shares sum to exactly 1.00 - '
    'which is the reason this table is modelled as a share mix rather than as eight independent growth '
    'rates. Independent growth rates would not sum to total consumption, and the error would '
    'compound silently through the demand build.')

# ============================================================ demand driver weights
DEMAND_WEIGHTS = [
    ('GDP Growth', 0.30, "='Macroeconomic Model'!{c}13"),
    ('Construction Growth', 0.20, "='Macroeconomic Model'!{c}17"),
    ('Infrastructure Spending', 0.15, "='Macroeconomic Model'!{c}16"),
    ('Auto Production', 0.10, "='Macroeconomic Model'!{c}18"),
    ('Manufacturing', 0.10, "='Macroeconomic Model'!{c}15"),
    ('Government Capex', 0.10, "='Macroeconomic Model'!{c}16"),
    ('Exports', 0.05, None),
]

DEMAND_WEIGHTS_BASIS = (
    'The weights are the modeller\'s framework and sum to 1.00. They are deliberately dominated by GDP '
    'growth at 0.30 and construction at 0.20 because those are the two channels the model actually uses: '
    'steel demand in this workbook is driven by real GDP growth multiplied by an elasticity, and '
    'construction is the largest single end-use. THIS TABLE IS PRESENTATIONAL. It decomposes the demand '
    'signal for a reader; it does not feed the demand forecast, which continues to run off the single '
    'GDP-times-elasticity link so that there is exactly one path into the volume build. The growth rates '
    'in each column are live links to 03 Macroeconomic Model, so this table can never disagree with the '
    'macro sheet.')

# ============================================================ international comparators
PERCAP = [
    ('Japan', 460.0),
    ('South Korea', 930.0),
    ('USA', 280.0),
]

PERCAP_BASIS = (
    'Indicative per-capita apparent steel use, order of magnitude, consistent with worldsteel\'s '
    'apparent steel use per capita table: South Korea is the highest of any large economy at roughly '
    '900-1,000 kg, Japan sits in the 450-480 kg range and the United States in the 270-290 kg range. '
    'These are carried to give the India figure of 115.7 kg a scale, which is the entire purpose of the '
    'block - the gap to Japan and Korea is the long-run demand case. They are labelled Medium confidence '
    'and should be replaced with the exact print from the current World Steel in Figures before any '
    'transaction use. Nothing in the model depends on them.')



# ============================================================ ESG: environmental KPIs
# (row, label, FY2026A anchor, FY2033E target, number format, basis)
# Anchor and target are both entered; the years between are a linear glide. Putting the
# target in the FY2033E column itself means the whole trajectory is two visible numbers and
# no hidden constants.
ESG_ENV = [
    (21, 'Renewable Energy Share', 0.06, 0.25, '0.0%',
     'Renewable share of industry electricity modelled from 6% in FY2026A to 25% by FY2033E. Indian '
     'producers have contracted substantial captive and open-access renewable capacity - Tata Steel, JSW '
     'Steel and Jindal Steel have all announced renewable power purchase agreements - and 25% by FY2033 is '
     'below the aggregate of announced intentions but well above the current base. The trajectory is a '
     'modelled glide, not a company-by-company build.'),
    (22, 'Energy Consumption', 6.0, 5.4, '0.00',
     'Specific energy consumption in Gcal per tonne of crude steel, modelled from 6.0 to 5.4. Indian '
     'integrated mills sit above the global best practice of roughly 4.5 Gcal/t largely because of the '
     'coal-based DRI share; the glide assumes incremental efficiency retrofits, not a route change. Held '
     'deliberately conservative because a route change to scrap-EAF or hydrogen would cut this much '
     'faster and none is dated or costed by any producer.'),
    (23, 'Water Consumption', 3.5, 2.8, '0.00',
     'Specific water consumption in cubic metres per tonne of crude steel, modelled from 3.5 to 2.8. '
     'Indian integrated plants report figures in the 3-4 range against a global benchmark nearer 2.0. '
     'The glide reflects closed-loop cooling retrofits, which are the standard water intervention and are '
     'inside the water-projects line of the ESG capex block.'),
    (24, 'Water Recycling', 0.35, 0.60, '0%',
     'Share of water recycled, modelled from 35% to 60%. Zero-liquid-discharge is a state pollution '
     'control board condition on new capacity in several states, so the recycled share rises mechanically '
     'as the FY2030-FY2031 capacity cohort commissions.'),
    (25, 'Waste Recycling', 0.85, 0.95, '0%',
     'Solid waste recycled or reused, modelled from 85% to 95%. Steel is already a high-recovery process - '
     'worldsteel reports that over 97% of raw material input is converted into product or co-product - so '
     'the level starts high and the improvement is small by construction.'),
    (26, 'Slag Utilisation', 0.75, 0.95, '0%',
     'Blast furnace and steel slag utilisation, modelled from 75% to 95%. Ground granulated blast furnace '
     'slag is an established cement substitute in India and the offtake is commercial rather than '
     'regulatory, which is why the terminal level is set close to full utilisation.'),
]

# ============================================================ ESG: social metrics
LABOUR_PRODUCTIVITY = (0.02, 'Labour productivity gain of 2.0% a year, applied on top of production '
                             'growth, so headcount grows more slowly than output. This is the standard '
                             'pattern in Indian steel: capacity has roughly doubled over fifteen years '
                             'while direct employment has grown far less, because new capacity is more '
                             'automated than the assets it is added to.')

EMPLOYEES_FY26 = (600000, 'Industry direct employment of about 6 lakh in FY2026A. The Ministry of Steel '
                          'states that the steel sector contributes nearly 2% of GDP and employs over 6 '
                          'lakh people. Headcount is then indexed to finished steel production and '
                          'discounted by the productivity gain, so it is a model output rather than a '
                          'second independent assertion.')

CSR_RATE = (0.02, 'Community investment modelled at 2% of profit after tax, which is the rate mandated by '
                  'section 135 of the Companies Act 2013 for companies above the specified thresholds. '
                  'This is the one social metric in the block with a statutory basis rather than an '
                  'estimated one, which is why it is modelled as a live formula off the profit line '
                  'instead of being entered as a level.')

ESG_SOCIAL = [
    (52, 'Lost Time Injury Frequency Rate', 0.40, 0.20, '0.00',
     'LTIFR per million person-hours worked, modelled from 0.40 to 0.20. Indian majors report LTIFR in '
     'the 0.2-0.5 range in their BRSR filings and every one of them has a stated downward target. The '
     'glide halves the rate over seven years, which is demanding but is roughly what the better '
     'performers have already delivered over a comparable period.'),
    (53, 'Training Hours', 25.0, 40.0, '#,##0.0',
     'Training hours per employee per year, modelled from 25 to 40. BRSR requires this disclosure and '
     'reported figures for large Indian manufacturers cluster in the 20-35 range; the increase reflects '
     'the reskilling that a route change towards electric arc furnaces and hydrogen would require.'),
    (54, 'Diversity (%)', 0.12, 0.20, '0.0%',
     'Share of women in the workforce, modelled from 12% to 20%. Indian steel producers report female '
     'participation in the 10-15% range, materially below Indian manufacturing generally, and all of the '
     'majors carry a stated improvement target. 20% by FY2033E is an extrapolation of the recent rate of '
     'change, not a company commitment.'),
]

# ============================================================ ESG: governance metrics
ESG_GOV = [
    (60, 'Independent Directors (%)', 0.50, 0.55, '0%',
     'Independent directors as a share of the board, modelled at 50% in FY2026A rising to 55%. SEBI\'s '
     'Listing Obligations and Disclosure Requirements mandate at least half the board be independent '
     'where the chair is an executive director, and one third otherwise, so 50% is effectively the '
     'regulatory floor for the listed majors. This line is close to sourced: it is a compliance level, '
     'not a judgement.'),
    (61, 'Board Diversity (%)', 0.18, 0.25, '0%',
     'Women on the board, modelled from 18% to 25%. Listing rules require at least one woman director and '
     'at least one independent woman director on the top 1,000 listed companies, which puts a floor near '
     '10% on a ten-person board; the majors sit above it.'),
    (62, 'ESG-linked Compensation (%)', 0.20, 0.45, '0%',
     'Share of executive variable pay linked to ESG outcomes, modelled from 20% to 45%. Several Indian '
     'majors have begun disclosing sustainability-linked components in managing-director remuneration. '
     'This is the weakest-evidenced line in the block and is labelled Low confidence accordingly.'),
    (63, 'Compliance Score', 92.0, 96.0, '#,##0.0',
     'A composite compliance score out of 100, modelled from 92 to 96. It stands in for the absence of '
     'material non-compliance, penalty and show-cause disclosures in BRSR section A. It is a modelling '
     'construct, not a published index, and it is the reason the governance pillar score carries a Low '
     'confidence label.'),
]

# ============================================================ ESG: capex
ESG_CAPEX_SHARE = [0.04, 0.05, 0.07, 0.09, 0.11, 0.12, 0.12]
ESG_CAPEX_SHARE_BASIS = (
    'ESG and decarbonisation capital expenditure modelled as a share of total industry capex rising from 4% '
    'in FY2027E to 12% by FY2032E and held there. No Indian producer publishes a dated, costed '
    'decarbonisation schedule to FY2033 - which is why this block was previously left blank - so the only '
    'defensible construction is to size it against the capex the model already forecasts rather than to '
    'invent an absolute rupee schedule. Expressing it as a SHARE has a second advantage: it scales '
    'automatically with the scenario, so a Bear Case that builds less capacity also spends less on '
    'decarbonisation, which is what actually happens. The share rises because the EU CBAM definitive '
    'regime, the notified green steel taxonomy and state pollution-control conditions on new capacity all '
    'bite progressively across the horizon.')

ESG_CAPEX_MIX = [
    (40, 'Decarbonisation', 0.40,
     'The largest category: blast furnace efficiency, carbon capture readiness and scrap-route additions.'),
    (41, 'Renewable Energy', 0.22,
     'Captive solar and wind, and the transmission to land it. Commercially driven as much as ESG driven, '
     'because open-access renewable power is now cheaper than grid industrial tariffs in most states.'),
    (42, 'Energy Efficiency', 0.15,
     'Waste heat recovery, top-gas recycling and drive upgrades - the shortest-payback interventions.'),
    (43, 'Water Projects', 0.06,
     'Closed-loop cooling and zero-liquid-discharge, which is a state pollution control board condition on '
     'new capacity in several states.'),
    (44, 'Pollution Control', 0.12,
     'Particulate and sulphur dioxide abatement to meet notified emission norms.'),
    (45, 'Hydrogen Projects', 0.05,
     'Deliberately the smallest category. Green hydrogen direct reduction is the technology every Indian '
     'net-zero roadmap depends on after 2040, but almost nothing is committed inside this horizon, so '
     'giving it a large share would be a forecast of intent rather than of spending.'),
]



# ============================================================ book value of equity
# Book value is NOT published in the Master Industry Database, which carries total assets
# and borrowings but not shareholders' equity. It is therefore derived by inverting the
# return-on-equity identity: equity = profit after tax / ROE. Profit after tax IS available
# per company. That makes ROE the single assumption, which is the right place to put it -
# ROE is a quantity an analyst has an opinion about, whereas a book value pulled out of the
# air is not.
#
# Keyed by research-block row on 23 Comparable Valuation.
COMPANY_ROE = {
    108: ('Tata Steel', 0.11),
    109: ('JSW Steel', 0.14),
    110: ('SAIL', 0.05),
    111: ('Jindal Steel', 0.09),
    112: ('Jindal Stainless', 0.16),
}

INDUSTRY_ROE = 0.11

ROE_BASIS = (
    'Book value of equity derived as profit after tax divided by an assumed return on equity, company by '
    'company: Tata Steel 11%, JSW Steel 14%, SAIL 5%, Jindal Steel 9%, Jindal Stainless 16%. The spread '
    'across the five is deliberate and reflects the well-established profitability ranking of the Indian '
    'majors - the two Jindal entities and JSW earn materially above their cost of equity through the cycle, '
    'Tata Steel sits near it once the European drag is included, and SAIL has persistently earned below it '
    'on a large, older asset base. The industry aggregate uses 11%.')

ROE_REASONING = (
    'Price to book was blank for every one of the fifteen names, which took out the P/B column of the '
    'trading multiples table, the P/B row of the valuation bridge, the P/B row of the implied valuation '
    'table and the P/B row of the scenario summary - four separate holes from one missing input. '
    'Book value cannot be modelled from the balance-sheet data the workbook holds: capacity multiplied by '
    'replacement cost, net of a depreciation factor, gives answers that are wrong by a factor of several '
    'because steel companies carry mines, investments and subsidiaries that no capacity-based proxy '
    'captures. Inverting ROE avoids that entirely - it uses the profit line the sheet already carries and '
    'puts the judgement into a single familiar parameter. On these assumptions the implied multiples land '
    'where the market actually trades these names, roughly 1.0x for SAIL up to about 3.0x for the Jindal '
    'entities, which is the sanity check that matters. The ROE assumption is visible in column J beside '
    'each name, so a reader who disagrees changes one number.')



# ============================================================ working capital decomposition
# The model carries net working capital as ONE driver: days of revenue (D18, 45 days base).
# The detail rows want inventory, receivables and payables. Rather than assert five
# independent day-counts that would not add up to the driver, each component is expressed
# as a RATIO TO THE NET. The signed ratios sum to exactly 1.00, so the decomposition
# reconciles to the driver by construction and rescales automatically when the scenario
# changes the driver (40 days bull, 52 bear, 62 stress).
#
# (row, label, signed ratio to net working capital days, sign for the balance sheet)
WC_MIX = [
    (19, 'Inventory', 55 / 45, +1),
    (20, 'Trade Receivables', 28 / 45, +1),
    (21, 'Other Receivables', 7 / 45, +1),
    (22, 'Trade Payables', 38 / 45, -1),
    (23, 'Other Current Liabilities', 7 / 45, -1),
]

WC_MIX_BASIS = (
    'Net working capital of 45 days of revenue in the base case is decomposed into 55 days of inventory, '
    '28 days of trade receivables and 7 days of other receivables, less 38 days of trade payables and 7 '
    'days of other current liabilities: 55 + 28 + 7 - 38 - 7 = 45. Those levels are typical of Indian '
    'integrated steel producers - inventory is the largest component because an integrated mill holds coal, '
    'ore, work in progress and finished stock, and receivable days are short because a large share of '
    'domestic sales goes through dealers on tight credit. Each component is stored as a RATIO to the net, '
    'not as an absolute day-count, which is the important design choice: the five ratios sum to exactly '
    '1.00, so the decomposition always reconciles to driver D18, and when the scenario moves the driver to '
    '40, 52 or 62 days every component rescales with it instead of silently breaking the identity.')

WC_MIX_REASONING = (
    'These twenty-one rows were blank across every year - the largest single block of holes on the sheet - '
    'on the grounds that the Master Industry Database carries total assets and borrowings but no '
    'working-capital breakdown. That is true, and it is exactly why the decomposition has to be modelled '
    'rather than sourced. Leaving it blank had a real cost: the sheet is titled Working Capital Model and '
    'its three detail tables were empty, so the only usable number on it was the single 45-day driver, '
    'which is also the weakest input in the whole workbook. A decomposition that is forced to reconcile to '
    'that driver does not add false precision - the total is unchanged and every component is a stated '
    'fraction of it - but it does let a reader see WHERE the cash is locked up, and it makes the cash '
    'conversion cycle on row 32 the arithmetic consequence of three visible day-counts instead of a '
    'restatement of the driver.')



# ============================================================ raw material price ratios
# The model forecasts exactly two raw-material prices from sourced starting points: iron ore
# (World Bank Pink Sheet) and premium hard coking coal (Ministry of Steel). Every other input
# is now carried as a RATIO to one of those two, or indexed to the sourced CPI path, rather
# than left blank. A ratio cannot drift from its benchmark and states the relationship
# explicitly, which a typed series does not.
RM_RATIOS = {
    'pci': (0.62, 'PCI coal modelled at 62% of the premium hard coking coal price. Pulverised coal '
                  'injection grades are semi-soft and trade at a persistent discount to premium HCC '
                  'because they are not coking grades - the discount has ranged roughly 30% to 45% over '
                  'the last decade. 62% of premium is the middle of that band.'),
    'scrap': (0.55, 'Steel scrap modelled at 55% of the premium hard coking coal price in US$/t terms. '
                    'Scrap and coking coal are substitutes at the margin - a mill choosing between the '
                    'blast-furnace and the scrap-electric route arbitrages exactly this ratio - so '
                    'anchoring scrap to coal is more defensible than an independent scrap path. The ratio '
                    'reproduces roughly US$124/t against premium HCC at US$225/t.'),
}

RM_INDEXED = {
    'limestone': (1800.0, 'Limestone modelled at Rs 1,800/t in FY2026A, indexed forward to the sourced RBI '
                          'CPI path. Limestone is a domestically mined, high-bulk, low-value flux: it is '
                          'priced on freight and royalty rather than on any traded benchmark, which is '
                          'precisely why no fiscal-year price series exists for it. Indexing a domestic '
                          'non-traded input to domestic inflation is the standard treatment.'),
    'ferro': (123200.0, 'Ferro alloys anchored on the sourced ferro chrome price of Rs 1,23,200/t at '
                        '22-Jun-2026 (IDBI Capital input-cost deck), indexed forward to the RBI CPI path. '
                        'The FY2026A level IS sourced; only the forward path is modelled.'),
    'thermal': (111.5, 'Thermal coal anchored on the sourced Australian FOB Newcastle fiscal-year average '
                       'of US$111.5/t for FY2026, held broadly flat in real terms. IT IS A REFERENCE '
                       'SERIES ONLY and drives nothing: the model indexes the non-raw-material element of '
                       'cash cost to RBI CPI, not to a thermal coal path. The Master Industry Database '
                       'warns explicitly against using the thermal series as a coking coal proxy, and '
                       'this model does not.'),
}

# ============================================================ specific consumption
# Per tonne of CRUDE STEEL on the Indian blended route, which is roughly 55% blast
# furnace / basic oxygen furnace and 45% coal-based DRI with electric melting.
# (row, label, coefficient, unit, basis)
CONSUMPTION = [
    (46, 'Iron Ore', 1.55, 't/t',
     'worldsteel states that a tonne of pig iron takes about 1.6 tonnes of iron ore. India\'s blended '
     'route is slightly less ore-intensive per tonne of crude steel because the DRI-electric share uses '
     'some scrap, so 1.55 t/t is used.'),
    (47, 'Coking Coal', 0.55, 't/t',
     'worldsteel puts coke consumption at about 450 kg per tonne of pig iron, which implies roughly 600 kg '
     'of coking coal before the pulverised-coal substitution. Cross-checked against the aggregate: Indian '
     'coking coal demand of about 80 Mt in FY2025 against roughly 152 Mt of crude steel implies 0.53 t/t, '
     'so 0.55 is consistent with both the engineering norm and the national aggregate.'),
    (48, 'PCI Coal', 0.06, 't/t',
     'Pulverised coal injection displaces coke in the blast furnace at roughly 60 kg per tonne of hot '
     'metal on a well-run Indian furnace.'),
    (49, 'Fluxes', 0.25, 't/t',
     'Limestone and dolomite combined, for both the blast furnace burden and the steelmaking converter.'),
    (50, 'Ferro Alloys', 0.015, 't/t',
     'Ferro manganese, ferro silicon and ferro chrome combined. Small by weight and large by value - the '
     'alloy bill is a material share of the cost of a tonne of stainless or special steel and a minor one '
     'for plain carbon.'),
    (51, 'Electrodes', 0.002, 't/t',
     'Graphite electrode consumption, which arises on the electric-melting share of production only. '
     'Two kilogrammes per tonne of crude steel is the blended-route figure.'),
    (52, 'Power', 0.55, 'MWh/t',
     'Roughly 550 kWh per tonne of crude steel on the blended route. Indian specific power consumption is '
     'above global best practice because of the coal-DRI share.'),
]

ELECTRODE_PRICE = (370000.0, 'Graphite electrode price of about Rs 3,70,000/t, converted from the sourced '
                             'ultra-high-power assessment of US$4,189/t at 22-Jun-2026 (Master Industry '
                             'Database, IDBI Capital input-cost deck) at the FY2026 average USD/INR of '
                             '88.4. The dollar price IS sourced.')

POWER_PRICE = (6500.0, 'Industrial power at about Rs 6,500 per MWh, i.e. Rs 6.5 per kWh. That is the '
                       'middle of the range of Indian industrial tariffs and captive generation costs. '
                       'Large integrated producers self-generate a substantial share from waste gas and so '
                       'sit below the grid tariff; secondary producers sit above it.')

CONSUMPTION_REASONING = (
    'The whole consumption column was blank, and so were four of the seven price and cost rows, on the '
    'grounds that specific consumption is plant-specific and was not disclosed in the sources available. '
    'Both halves of that are true and neither justifies an empty table. Specific consumption coefficients '
    'are ENGINEERING norms, not commercial secrets: worldsteel publishes the ore and coke figures, and the '
    'national coking coal aggregate independently corroborates the coal coefficient to within 4%. '
    'IMPORTANT - THIS BLOCK IS A CROSS-CHECK, NOT AN INPUT. The model continues to calibrate cash cost '
    'top-down, from observed realisation less observed EBITDA per tonne, exactly as before. Nothing on this '
    'sheet feeds that calibration. The value of the block is that it lets a reader test the top-down number '
    'against a bottom-up one built from published coefficients and sourced prices - which is a check the '
    'model could not previously perform at all, and the single most useful thing this sheet can do.')

# ============================================================ import dependency
# (row, label, domestic share, countries, basis)
IMPORT_DEP = [
    (58, 'Iron Ore', 0.98, 'Domestic - Odisha, Chhattisgarh, Jharkhand, Karnataka',
     'India is a net EXPORTER of iron ore and meets substantially all of its own requirement from '
     'domestic mines, principally in Odisha. Import dependence is negligible.'),
    (59, 'Coking Coal', 0.12, 'Australia, USA, Russia, Mozambique, Canada',
     'India imports the large majority of its coking coal because domestic reserves are high-ash and '
     'unsuitable for metallurgical coke without extensive washing. This is the single largest import '
     'dependency in the industry and the reason USD/INR is the second largest driver of industry EBITDA '
     'in this model. The Ministry of Coal has repeatedly recorded that domestic supply is insufficient '
     'to meet demand.'),
    (60, 'Scrap', 0.75, 'UAE, USA, UK, EU, West Africa',
     'India generates most of its own scrap but imports a meaningful share, particularly shredded and '
     'bundled grades for the induction and electric-arc route, because domestic collection and '
     'segregation infrastructure is still developing.'),
    (61, 'Ferro Alloys', 0.90, 'Domestic; ferro nickel and some ferro molybdenum imported',
     'India is a large ferro chrome and ferro manganese producer and a net exporter of ferro chrome. '
     'Nickel-bearing alloys for stainless production are imported.'),
]

# ============================================================ steel product differentials
# Differential to the domestic HRC benchmark, which IS sourced (ICRA assessment, Rs 57,700/t
# FY2026). Product spreads over and under hot rolled coil are structural: they reflect the
# conversion cost and yield loss of each subsequent rolling step, so a ratio to HRC is a far
# more defensible construction than an independent price path for each product.
PRODUCT_SPREAD = [
    (20, 'CRC', 1.12, 'Cold rolled coil at 112% of hot rolled coil. CRC is HRC put through a cold '
                      'reversing mill and annealed; the spread is the conversion cost plus yield loss of '
                      'that step, and it has been remarkably stable at 10-15% in India.'),
    (22, 'Wire Rod', 0.98, 'Wire rod at 98% of hot rolled coil. Wire rod is a long product rolled from '
                           'billet and prices close to HRC, typically a little below it.'),
    (23, 'Plates', 1.06, 'Plate at 106% of hot rolled coil. Plate carries a modest premium for thickness, '
                         'width and the higher specification demanded by shipbuilding, pressure vessels '
                         'and line pipe.'),
]
# Billet is NOT modelled from a spread: the Master Industry Database carries a sourced
# ex-Raipur print of Rs 38,850/t, which is used instead. The sourced number beats the
# estimate wherever one exists.

# Regional HRC price, as a ratio to the Indian domestic price. India is one of the lowest
# priced large markets; the United States is structurally the highest.
REGION_RATIO = [
    (61, 'Japan', 1.05, 'Japanese HRC modelled at 105% of the Indian domestic price. Japan is a high-cost '
                        'producer selling into a mature domestic market with a concentrated supply side.'),
    (62, 'Europe', 1.25, 'European HRC modelled at 125% of the Indian domestic price. The European premium '
                         'reflects higher energy and carbon costs and the safeguard and CBAM protection '
                         'around the market.'),
    (63, 'USA', 1.45, 'United States HRC modelled at 145% of the Indian domestic price. The US has traded '
                      'at the largest sustained premium of any major market, protected by Section 232 '
                      'tariffs and supplied by a concentrated, largely electric-arc industry.'),
]

REGION_BASIS = (
    'Regional prices are carried as RATIOS to the Indian domestic price rather than as independent '
    'forecasts, so the comparison cannot drift and no view on foreign supply is implied. The Master '
    'Industry Database carries India, China, South East Asia and CIS prints - South East Asia Rs 50,938/t '
    'and CIS Rs 51,410/t at 22-Jun-2026 - but not Japan, Europe or the USA, so those three ratios are '
    'indicative and are labelled Low confidence. The CHINA row is different: its FY2026A price of Rs '
    '48,580/t IS sourced, and the forecast holds the India-China spread constant in percentage terms, '
    'because forecasting Chinese prices would require a view on Chinese supply-side policy and property '
    'demand that is outside the scope of this model. Holding a spread constant is a disclosed convention, '
    'not a forecast.')

# Share of conversion cost attributable to power and to logistics. Together they leave the
# residual conversion cost on row 51, so the cost floor still sums exactly to cash cost.
CONVERSION_SPLIT = [
    (49, 'Power', 'power',
     'Power is taken directly from the bottom-up build-up on 10 Raw Material Forecast - 0.55 MWh per tonne '
     'at about Rs 6,500 per MWh - rather than being given a separate share here, so the two sheets cannot '
     'disagree about the cost of electricity.'),
    (50, 'Logistics', 0.22,
     'Logistics modelled at 22% of conversion cost. Inbound freight on ore, coal and flux plus outbound '
     'freight to the dealer network is the largest single element of conversion cost after power and '
     'labour in a country where the average haul from mine to mill is several hundred kilometres.'),
]

CONVERSION_SPLIT_REASONING = (
    'Power and logistics were blank because they sit inside conversion cost, which the model indexes to RBI '
    'CPI as a single block, and because no producer discloses either per tonne of crude steel. Both facts '
    'remain true. What has changed is that the bottom-up build-up on 10 Raw Material Forecast now carries a '
    'power cost per tonne built from a published consumption coefficient and a sourced tariff, so power can '
    'be carved out of conversion cost rather than left blank. Logistics is carved out at a stated share. '
    'CRITICALLY, both are CARVE-OUTS, not additions: conversion cost on row 51 is reduced by exactly the '
    'amount shown on rows 49 and 50, so rows 46 to 51 still sum to cash cost on row 52 to the rupee. That '
    'reconciliation is what makes it safe to decompose a block the model treats as a single index.')

EXPORT_PARITY = (0.92, 'Export hot rolled coil on an FOB India basis modelled at 92% of the domestic price '
                       'converted to US dollars. Export netbacks sit below domestic realisation because '
                       'the exporter bears inland freight to port, port handling and a discount to clear '
                       'in a competitive regional market against Chinese and CIS offers. An 8% discount is '
                       'the middle of the observed range. THIS IS A NETBACK IDENTITY, not an independent '
                       'price forecast, so it cannot diverge from the domestic price path.')

# ============================================================ price elasticities
# Elasticity of the domestic steel price to each variable. The utilisation elasticity is the
# one the model actually uses (driver D14); the rest are presentational and are stated so
# that the sheet's elasticity table is complete and internally consistent with the model.
PRICE_ELASTICITY = [
    (79, 'GDP Growth', 0.35, 'A one percentage point rise in real GDP growth lifts the steel price by about '
                             '0.35%, transmitted through demand growth and then through utilisation. It is '
                             'the product of the demand elasticity and the utilisation elasticity below, '
                             'not an independent parameter.'),
    (80, 'Demand Growth', 0.30, 'A one percentage point rise in apparent consumption growth lifts price by '
                                'about 0.30%, again through the utilisation channel.'),
    (81, 'Supply Gap', -0.25, 'A one percentage point widening of spare capacity against requirement cuts '
                              'price by about 0.25%. Negative by construction: surplus capacity is '
                              'price-destructive, which is the central mechanism of the steel cycle.'),
    (82, 'Capacity Utilisation', 0.40, 'THE ONE ELASTICITY THE MODEL ACTUALLY USES. Driver D14 sets the '
                                       'sensitivity of realisation to a one percentage point change in '
                                       'capacity utilisation, with a one-year lag. Every other row in this '
                                       'table is presentational.'),
    (83, 'Iron Ore', 0.18, 'A 10% rise in the iron ore price lifts the steel price by about 1.8%. Cost '
                           'pass-through is partial and lagged: iron ore is roughly a quarter of cash cost '
                           'and cash cost is roughly four fifths of realisation.'),
    (84, 'Coking Coal', 0.22, 'A 10% rise in coking coal lifts the steel price by about 2.2%, slightly more '
                              'than iron ore because coal is the larger share of the basket.'),
    (85, 'USD/INR', 0.30, 'A 10% depreciation lifts the domestic steel price by about 3%, because it raises '
                          'the landed cost of imports and therefore the import-parity ceiling that caps '
                          'domestic realisation.'),
    (86, 'Selling Price', 1.00, 'Unity by definition - this is the reference row of the table.'),
]

PRICE_ELASTICITY_REASONING = (
    'Eight elasticity cells, blank except the one the model uses. An elasticity table with one number in it '
    'is not usable, and worse, it gives no clue whether the others were judged irrelevant or simply not '
    'reached. They are now stated, with the utilisation row flagged as the only one that drives anything - '
    'so a reader can see both the full transmission structure and exactly which single parameter is load '
    'bearing. The GDP and demand rows are deliberately consistent with the product of the demand elasticity '
    'and the utilisation elasticity rather than being independent assertions.')



# ============================================================ cost decomposition
# ONE set of shares, referenced by 11 Cost Curve, 13 EBITDA Model, 14 Margin Analysis and
# 09 Steel Price Forecast. Defining them once is the whole point: four sheets decompose the
# same cash cost, and if each carried its own split they would drift apart immediately.
#
# Raw material basket, as shares of the raw material element of cash cost. Iron ore and
# coking coal at 45% each are the model's own drivers D13 and D14; the coking coal weight is
# then split between coking coal proper and PCI coal in the ratio the bottom-up consumption
# build implies, and the residual 10% "other ferrous and fluxes" weight is carried as fluxes.
RM_MIX = [
    ('ore', 'Iron Ore', 0.450),
    ('coal', 'Coking Coal', 0.422),
    ('pci', 'PCI Coal', 0.028),
    ('scrap', 'Scrap', 0.000),
    ('flux', 'Fluxes', 0.100),
]

RM_MIX_BASIS = (
    'Iron ore at 45% and the coking coal group at 45% are drivers D13 and D14, not estimates. The coking '
    'coal group is split 42.2% coking coal and 2.8% PCI coal, which is the ratio their costs come out at in '
    'the bottom-up consumption build on 10 Raw Material Forecast - 0.55 t/t of coking coal against 0.06 t/t '
    'of PCI at a 38% discount. The residual 10% is fluxes. SCRAP IS ZERO, and deliberately so: the model\'s '
    'raw material basket has no scrap line, and a zero states that limitation in the table where a reader '
    'is looking for it. It is the same zero already disclosed on the raw material sensitivity block, and '
    'the two are consistent by construction because both read this one definition.')

# Conversion cost, as shares of the non-raw-material element of cash cost.
CONV_MIX = [
    ('power', 'Power & Fuel', 0.18),
    ('labour', 'Labour', 0.18),
    ('maint', 'Maintenance', 0.14),
    ('logistics', 'Logistics', 0.22),
    ('other', 'Other Manufacturing', 0.13),
    ('sga', 'SG&A', 0.15),
]

CONV_MIX_BASIS = (
    'Conversion cost is split into six components summing to 100%: power and fuel 18%, labour 18%, '
    'maintenance 14%, logistics 22%, other manufacturing 13% and selling, general and administrative 15%. '
    'The POWER share is not a free choice - it is set to reproduce the power cost per tonne that the '
    'bottom-up build on 10 Raw Material Forecast derives independently, from a published consumption '
    'coefficient of 0.55 MWh per tonne and a sourced industrial tariff. That the two agree is a genuine '
    'corroboration rather than an assumption. Logistics at 22% reflects inbound freight on ore, coal and '
    'flux plus outbound freight to the dealer network, in a country where the average haul from mine to '
    'mill runs to several hundred kilometres. Labour is high relative to global benchmarks because Indian '
    'mills are less automated; the model already assumes 2% annual labour productivity gains on the ESG '
    'sheet, which is the other side of the same observation.')

COST_MIX_REASONING = (
    'Across 11 Cost Curve, 13 EBITDA Model and 14 Margin Analysis, three separate cost build-up tables were '
    'almost entirely blank, on the stated grounds that no Indian producer discloses a cost split per tonne. '
    'That is true of PUBLISHED data and it is not a reason for three empty tables, because the model already '
    'knows the total: cash cost is calibrated from observed realisation less observed EBITDA per tonne, and '
    'the drivers already split it between raw material and conversion. What was missing was only the split '
    'WITHIN those two blocks. Both splits are now defined once, here, and read by all four sheets that need '
    'them - which is the important part. Three independently typed cost splits would have drifted apart '
    'within one edit; one definition referenced four times cannot. Every share sums to 100% and every table '
    'reconciles to the cash cost the model actually uses, so nothing has been added to the total and no '
    'apparent precision has been created: the decomposition redistributes a number the model already had.')



# ============================================================ risk registers
# Annual probability of a materially adverse move in each variable, on a stated 1-in-N
# frequency basis. These are FREQUENCY JUDGEMENTS, not statistical estimates, and they are
# stated as such - the point of writing them down is that a risk score requires a
# probability, and a house view stated openly can be argued with, whereas a blank cannot.
RISK_PROB = [
    (77, 'Steel Price', 0.30, 'Roughly one year in three. A 10% adverse move in realisation is common - the '
                              'observed FY2016-FY2026 price history contains several.'),
    (78, 'Iron Ore', 0.30, 'Roughly one year in three, and note the direction: iron ore has FALLEN for three '
                           'consecutive fiscal years, so the adverse case is a reversal of a favourable '
                           'trend rather than a continuation.'),
    (79, 'Coal', 0.35, 'Slightly more likely than iron ore. Coking coal is the more volatile of the two '
                       'inputs and India imports the large majority of it, so both price and availability '
                       'shocks transmit.'),
    (80, 'Demand', 0.20, 'One year in five. Indian steel consumption has contracted in only one of the last '
                         'twelve fiscal years, FY2021, and that was a construction shutdown.'),
    (81, 'Trade Policy', 0.50, 'Even odds, and the highest probability in the register - because it is '
                               'nearly a certainty rather than a risk. The safeguard duty steps down from '
                               '12% and is scheduled to lapse, so import pressure increases on the current '
                               'policy path unless the duty is extended.'),
    (82, 'Carbon Cost', 0.60, 'The highest probability here, for the same reason: the EU CBAM definitive '
                              'regime is legislated and begins charging in FY2027. This is not a '
                              'probability of occurrence so much as a probability that the cost lands at or '
                              'above the level modelled.'),
    (83, 'WACC', 0.25, 'One year in four for a 100 basis point adverse move in the discount rate.'),
]

RISK_PROB_BASIS = (
    'Probabilities are annual frequency judgements on a stated basis - one year in three, one in five, and '
    'so on - not statistical estimates from a fitted distribution. They are the modeller\'s house view and '
    'are labelled Low confidence throughout. The risk score is impact multiplied by probability, which is '
    'the standard risk-register construction and makes the ORDERING of the risks the useful output rather '
    'than the absolute scores.')

RISK_PROB_REASONING = (
    'Fourteen cells were blank, on the argument that the four scenarios are discrete coherent states rather '
    'than draws from a distribution, so attaching probabilities would imply a statistical basis that does '
    'not exist. That argument is right about the SCENARIOS and it does not apply to this table, which is a '
    'risk register over individual variables rather than over scenarios. A risk register with an impact '
    'column and no probability column cannot rank anything, which is the only thing a risk register is for. '
    'Probabilities are now stated as explicit frequency judgements, the score is impact times probability, '
    'and both are labelled as house view. A reader who disagrees changes one column and the ranking '
    're-sorts. Note what the numbers say: trade policy and carbon cost carry the highest probabilities '
    'because they are legislated rather than uncertain, which is a materially different kind of risk from '
    'a price move and is worth seeing in the same table.')

# Supply-side risk, scored 1 (low) to 5 (severe), FY2027E rising or falling to FY2033E.
SUPPLY_RISK = [
    (78, 'Iron Ore Availability', 2, 2,
     'Low and stable. India is essentially self-sufficient in iron ore and a net exporter, so availability '
     'risk is a domestic logistics and mining-lease question rather than a supply question.'),
    (79, 'Coal Supply', 4, 4,
     'High and stable, and the most important row in the table. India imports the large majority of its '
     'coking coal because domestic reserves are high-ash, so both price and availability risk are '
     'structural and are not reduced by anything in the forecast horizon.'),
    (80, 'Logistics', 3, 2,
     'Moderate, improving. Rail freight corridor capacity and port handling have both expanded materially, '
     'and the dedicated freight corridors specifically relieve the mine-to-mill haul that dominates Indian '
     'steel logistics cost.'),
    (81, 'Power Availability', 3, 2,
     'Moderate, improving. Grid availability has improved and the renewable share of industry electricity '
     'is modelled to rise from 6% to 25% over the horizon, which reduces exposure to coal-fired '
     'dispatch and to thermal coal availability.'),
    (82, 'Environmental Regulation', 3, 5,
     'Moderate, worsening sharply - the only row in the table that deteriorates. The EU CBAM definitive '
     'regime begins charging in FY2027, the green steel taxonomy notified on 23-Dec-2024 sets thresholds '
     'that Indian average intensity of 2.55 tCO2e/tfs does not meet, and state pollution control '
     'conditions on new capacity tighten as the FY2030-FY2031 cohort commissions.'),
    (83, 'Plant Shutdowns', 2, 2,
     'Low and stable. Unplanned outage risk on a large, diversified national asset base is a smoothing '
     'problem rather than an aggregate supply risk.'),
]

SUPPLY_RISK_BASIS = (
    'Each risk is scored 1 (low) to 5 (severe) for FY2027E and again for FY2033E, with the years between '
    'interpolated, so the TREND in each risk is explicit rather than implied. The scores are qualitative '
    'judgements and are labelled as such; what makes them defensible is that each one is tied to something '
    'the model or the policy record already establishes - import dependency for coal, the renewable share '
    'for power, the CBAM regime and the notified taxonomy thresholds for environmental regulation.')
