"""Structural map of the workbook: legacy block boundaries and column grammar."""

# First row of the pre-existing research block on each sheet. Nothing at or below
# this row may be written to. Sheets absent from this map have no legacy block.
LEGACY = {
    'Control Panel': 35,
    'Macroeconomic Model': 99,
    'Steel Demand Model': 112,
    'Steel Supply Model': 88,
    'Capacity Forecast': 94,
    'Capacity Expansion Tracker': 75,
    'Capacity Utilisation': 64,
    'Steel Price Forecast': 95,
    'Raw Material Forecast': 83,
    'Cost Curve': 76,
    'Revenue Forecast': 93,
    'EBITDA Model': 98,
    'Margin Analysis': 87,
    'Working Capital Model': 69,
    'Cash Flow Model': 95,
    'Capital Allocation': 76,
    'Industry Cycle Model': 86,
    'Trade Model': 73,
    'ESG Model': 94,
    'Scenario Manager': 68,
    'Sensitivity Analysis': 95,
    'Comparable Valuation': 106,
    'Industry Dashboard': 30,
    'Audit Checks': 13,
    'Sources': 13,
}

# Legacy block year columns: index 0 = FY2026A, 1..7 = FY2027E..FY2033E
LEG = 'CDEFGHIJ'
# New-table grammar A: label in A, FY2026A in B, FY2027E..FY2033E in C..I
NEW8 = 'BCDEFGHI'
# New-table grammar B: label in A, FY2027E..FY2033E in B..H (no actual column)
NEW7 = 'BCDEFGH'

YEARS = ['FY2026A', 'FY2027E', 'FY2028E', 'FY2029E', 'FY2030E', 'FY2031E', 'FY2032E', 'FY2033E']
FCST = YEARS[1:]

# Sheet code used to prefix Support & Audit reference IDs
CODE = {
    'Cover': 'CV', 'Index': 'IX', 'Database Import': 'DI', 'Forecast Horizon': 'FH',
    'Control Panel': 'CP', 'Model Assumptions': 'MA', 'Model Calibration': 'MC',
    'Macroeconomic Model': 'MM', 'Steel Demand Model': 'DM', 'Steel Supply Model': 'SM',
    'Capacity Forecast': 'CF', 'Capacity Expansion Tracker': 'CT', 'Capacity Utilisation': 'CU',
    'Steel Price Forecast': 'PF', 'Raw Material Forecast': 'RM', 'Cost Curve': 'CC',
    'Revenue Forecast': 'RF', 'EBITDA Model': 'EM', 'Margin Analysis': 'MG',
    'Working Capital Model': 'WC', 'Cash Flow Model': 'CX', 'Capital Allocation': 'CA',
    'Industry Cycle Model': 'IC', 'Trade Model': 'TM', 'ESG Model': 'ES',
    'Scenario Manager': 'SC', 'Sensitivity Analysis': 'SA', 'Comparable Valuation': 'VL',
    'Industry Dashboard': 'DB', 'Audit Checks': 'AC', 'Sources': 'SR',
}

# Navigation targets used by the Control Panel and Index hyperlinks
NAV = {
    'Database Import': 'Database Import', 'Model Assumptions': 'Model Assumptions',
    'Model Calibration': 'Model Calibration', 'Macro Model': 'Macroeconomic Model',
    'Demand Model': 'Steel Demand Model', 'Supply Model': 'Steel Supply Model',
    'Capacity Model': 'Capacity Forecast', 'Price Model': 'Steel Price Forecast',
    'Revenue Model': 'Revenue Forecast', 'EBITDA Model': 'EBITDA Model',
    'Valuation': 'Comparable Valuation', 'Dashboard': 'Industry Dashboard',
    'Audit': 'Audit Checks', 'Sources': 'Sources',
}

# Reasons used for ORANGE cells, kept central so the wording is consistent.
NA = {
    'macro_sub': ('Series not held in the Industry Financial Model or the Master Industry Database, and no '
                  'institutional forecast to FY2033 exists. MOSPI publishes IIP and GVA, RBI publishes WPI and '
                  'the policy repo rate, and Brent is a market print - none were retrievable in this research '
                  'cycle. Populate from MOSPI / RBI / EIA before transaction use.'),
    'macro_actual': ('FY2026 actual not retrievable in this research cycle. The model deliberately carries only '
                     'the macro series it can source (real GDP, CPI, USD/INR, 10-year G-sec).'),
    'sector_split': ('Sector-wise steel consumption splits are published by the Ministry of Steel and worldsteel '
                     'but were not retrievable in this research cycle. No split is asserted rather than an '
                     'invented one, because sector shares drive the demand mix and would propagate silently.'),
    'intl': ('International per-capita and price comparators beyond India, China and the world average were not '
             'retrievable. The Master Industry Database carries only India 115.7 kg, China 604 kg and world '
             '215 kg (worldsteel SSY-2025).'),
    'product_price': ('No fiscal-year average product-level price series exists in the public domain for India. '
                      'The Master Industry Database records this explicitly as a material limitation: only '
                      'point-in-time HRC and rebar assessments are available, not CRC, wire rod, plate or billet '
                      'fiscal-year averages.'),
    'company_cost': ('Implied cash cost can only be derived for producers that disclose BOTH realisation and '
                     'EBITDA per tonne. That is four names - Tata Steel, JSW Steel, SAIL and Jindal Steel. No '
                     'cost disclosure exists for the remaining eleven names in the coverage universe.'),
    'share_price': ('Share prices, market capitalisation and book value were not sourced in this research cycle. '
                    'This is disclosed as standing limitation 6 on 25 Audit Checks. Enter prices in the shaded '
                    'column on 23 Comparable Valuation section A and every trading multiple resolves.'),
    'esg': ('Company and industry ESG operating metrics (emissions, energy, water, safety, workforce, board '
            'composition) are disclosed in BRSR filings and sustainability reports but were not retrievable in '
            'this research cycle. Nothing is asserted.'),
    'esg_capex': ('No Indian producer discloses a decarbonisation capex schedule by category to FY2033. Tata '
                  'Steel, JSW Steel and SAIL disclose aggregate net-zero ambitions without a dated, costed plan.'),
    'raw_hist': ('No fiscal-year average series exists for this input. PCI coal, scrap, limestone, dolomite and '
                 'ferro alloys are not carried in the Master Industry Database; coking coal deliberately asserts '
                 'no fiscal-year average (Conflict C01).'),
    'consumption_coef': ('Specific consumption coefficients per tonne of crude steel (ore, coal, flux, alloy, '
                         'electrode, power) are plant-specific and were not disclosed in the sources available. '
                         'The model instead calibrates cash cost top-down from observed realisation less EBITDA.'),
    'trade_region': ('Regional export and import volumes are not published; only percentage shares by country '
                     'are available, and only for the top destinations. Shares are carried on 19 Trade Model.'),
    'hist_multiple': ('Historical trading multiples require a share-price history that was not sourced. '
                      'See standing limitation 6 on 25 Audit Checks.'),
    'financing': ('The model is an unlevered INDUSTRY aggregate. It does not forecast debt draw-downs, '
                  'repayments, dividends, buybacks or equity issuance, because those are capital-structure '
                  'decisions taken at company level. Net debt is rolled forward from free cash flow only.'),
    'crude_fy22': ('The Joint Plant Committee series carried in the Master Industry Database records "Data Not '
                   'Publicly Available" for FY2022 crude steel production. Only finished steel production '
                   '(113.60 Mt) is available for that year.'),
    'regional_cap': ('State-wise and region-wise installed capacity is not published by the Joint Plant '
                     'Committee at the aggregate level. Only SAIL discloses plant-wise capacity.'),
    'wc_detail': ('Inventory, receivable and payable balances are not derivable from the Master Industry '
                  'Database, which carries total assets and borrowings but not a working-capital breakdown. '
                  'Driver D18 models net working capital as a single days-of-revenue parameter. This is '
                  'standing limitation 4 on 25 Audit Checks.'),
    'supply_risk': ('Forward risk scoring for input availability, logistics, power and environmental compliance '
                    'is a qualitative overlay for which no published index exists. Nothing is asserted.'),
    'peer_fin': ('Enterprise value, price/earnings and price/book require a share price, which was not sourced. '
                 'Revenue, EBITDA and net debt are available and are shown.'),
}
