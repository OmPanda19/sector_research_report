"""Fills for 11 Cost Curve, 12 Revenue Forecast, 13 EBITDA Model, 14 Margin Analysis,
15 Working Capital Model and 16 Cash Flow Model."""
from .style import (SheetWriter, PCT0, PCT1, PCT2, NUM0, NUM1, NUM2, MT, MTS, RS, CR,
                    X1, X2, USD, TEXT, PPT)
from .support import na_row, est_row
from .estimates import (WC_MIX, WC_MIX_BASIS, WC_MIX_REASONING, RM_MIX, RM_MIX_BASIS,
                        CONV_MIX, CONV_MIX_BASIS, COST_MIX_REASONING)
from .spec import NA

LEG = 'CDEFGHIJ'
FLEG = 'DEFGHIJ'
NEW8 = 'BCDEFGHI'
NEW7 = 'BCDEFGH'
MA8 = 'DEFGHIJK'      # 02 Model Assumptions FY2026A..FY2033E
MA7 = 'EFGHIJK'
SCEN = ['Base Case', 'Bull Case', 'Bear Case', 'Stress Case']

# per-tonne depreciation, used to turn cash cost into total cost
DPT = "'Cash Flow Model'!{lc}99/'Steel Supply Model'!{lc}101*10"


def _sup(item, value, unit, method, formula, primary, secondary, assumption, reasoning,
         cross, conf, linked, freq, last, comments, numfmt=None):
    return dict(item=item, value=value, unit=unit, method=method, formula=formula, primary=primary,
                secondary=secondary, assumption=assumption, reasoning=reasoning, cross=cross,
                conf=conf, linked=linked, freq=freq, last=last, comments=comments, numfmt=numfmt)


ZERO_BY_CONSTRUCTION = (
    'ZERO BY CONSTRUCTION, not a placeholder. The model carries a single blended industry '
    'realisation and a single blended cash cost, so there is no product mix, export mix or '
    'currency channel to isolate at revenue level - the exchange rate enters the model through '
    'raw material cost only. Entering anything other than zero here would double-count. The cell '
    'is a live input, so a user with a mix view can enter it and the bridge still closes.')

NO_FIXED_COST = (
    'ZERO BY CONSTRUCTION. Cash cost in this model is entirely variable per tonne, so margin is '
    'volume-neutral and a volume shock moves EBITDA but not EBITDA MARGIN. That is a stated '
    'limitation: a real producer has a fixed-cost block and would show positive operating '
    'leverage. It is disclosed rather than approximated because no fixed/variable cost split is '
    'disclosed by any Indian producer.')


# ======================================================================= 11 Cost Curve
CC_ROWS = [
    (19, 'Tata Steel', '$C$85', "'Revenue Forecast'!$C$101", "'EBITDA Model'!$C$108"),
    (20, 'SAIL', '$C$86', "'Revenue Forecast'!$C$103", "'EBITDA Model'!$C$110"),
    (21, 'JSW Steel', '$C$87', "'Revenue Forecast'!$C$102", "'EBITDA Model'!$C$109"),
    (22, 'Jindal Steel', '$C$88', "'Revenue Forecast'!$C$104", "'EBITDA Model'!$C$111"),
]


def cost_curve(wb, audit, eng):
    ws = wb['Cost Curve']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- executive summary rows 9-14
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}9', f'=${lc}$80', RS)
        w.f(f'{nc}10', f'={lc}80+' + DPT.format(lc=lc), RS)
        w.link(f'{nc}11', f"='Steel Price Forecast'!${lc}$103", RS)
        w.link(f'{nc}12', f"='EBITDA Model'!${lc}$101", RS)
        w.f(f'{nc}13', f'=MAX({lc}85:{lc}88)', RS)
        w.f(f'{nc}14', f'={nc}11-{nc}9', RS)

    # ---- company cost curve rows 19-33, ranked ascending by cash cost
    for r, name, cost, vol, ebt in CC_ROWS:
        w.txt(f'B{r}', name)
        w.link(f'C{r}', f'={vol}', MT)
        w.link(f'D{r}', f'={cost}', RS)
        w.f(f'E{r}', f'=D{r}+' + DPT.format(lc='C'), RS)
        w.link(f'F{r}', f'={ebt}', RS)
    # rank 5: Jindal Stainless - derivable but NOT comparable
    w.txt('B23', 'Jindal Stainless (stainless - not comparable)')
    w.link('C23', "='Revenue Forecast'!$C$105", MT)
    w.f('D23', "='Steel Price Forecast'!$C$103*2.786876-'EBITDA Model'!$C$112", RS)
    w.f('E23', '=D23+' + DPT.format(lc='C'), RS)
    w.link('F23', "='EBITDA Model'!$C$112", RS)
    # rank 6: the residual
    w.txt('B24', 'Others - secondary sector, RINL, AM/NS, ESL, unlisted')
    w.f('C24', "='Steel Supply Model'!$C$101-SUM(C19:C23)", MT)
    for col in 'DEF':
        w.na(f'{col}24', NA['company_cost'])
    for r in range(19, 25):
        w.f(f'G{r}', f'=C{r}' if r == 19 else f'=G{r - 1}+C{r}', MT)
    for r in range(25, 34):
        w.narow('BCDEFG', r, NA['company_cost'])
    nas.append(na_row('Company cost curve, ranks 7 to 15', 'B25:G33', NA['company_cost'], unit='Rs/t'))
    nas.append(na_row('Cost of the Others block', 'D24:F24', NA['company_cost'], unit='Rs/t'))
    rows.append(_sup(
        'Company cost curve, FY2026A', '=D22-D19', 'Rs/t spread',
        'Cross-sheet reference to the producer cost curve in the research block below, ranked ascending',
        "='Cost Curve'!$C$85 (Tata) .. $C$88 (Jindal Steel); Others production = residual",
        'Derived from the Master Industry Database: each producer\'s realisation less its disclosed EBITDA '
        'per tonne',
        'Research block rows 85 to 88; 13 EBITDA Model rows 108 to 113',
        'Only four producers disclose BOTH realisation and EBITDA per tonne on a comparable carbon-steel '
        'basis, so only four points on the curve are real.',
        'The spread from lowest to highest cost among the four majors is only Rs 3,777/t, about 7.7% of cash '
        'cost. That is a genuinely useful finding: the Indian carbon steel cost curve is FLAT among the '
        'integrated majors, so cost position is not the differentiator - capital structure and product mix '
        'are. Tata Steel leads on captive iron ore; SAIL is second on captive ore despite the highest '
        'legacy cost base; JSW is highest of the four because it buys most of its ore.',
        'Each producer cost must equal industry cash cost plus the FY2026 spread held constant, and the '
        'volume-weighted average must return the industry figure of Rs 49,241/t',
        'Medium', 'Revenue Forecast, EBITDA Model, Steel Supply Model', 'Quarterly', 'FY2026A',
        'Was entirely blank. Jindal Stainless is shown at rank 5 with an explicit caveat: its cash cost of '
        'about Rs 1,45,500/t is real but reflects stainless, whose realisation is 2.79x carbon steel.', RS))

    # ---- cost build-up rows 38-51, decomposed from the shared cost mix.
    # Raw material shares in column J, conversion shares in column K, both visible.
    RM_ROW = {'ore': 38, 'coal': 39, 'pci': 40, 'scrap': 41, 'flux': 42}
    CONV_ROW = {'power': 43, 'labour': 44, 'maint': 45, 'logistics': 46, 'other': 47, 'sga': 50}
    for key, label, share in RM_MIX:
        r = RM_ROW[key]
        w.est(f'J{r}', share, '0.000')
        w.f(f'B{r}', f"=$C$80*'Model Assumptions'!$D$20*$J${r}", RS)
    for key, label, share in CONV_MIX:
        r = CONV_ROW[key]
        w.est(f'K{r}', share, '0.000')
        w.f(f'B{r}', f"=$C$80*(1-'Model Assumptions'!$D$20)*$K${r}", RS)
    # SG&A on row 50 is a MEMO carve-out of the conversion block, so cash cost on row 48
    # remains the sum of rows 38 to 47 and total cost stays row 48 plus depreciation.
    w.f('B47', "=$C$80*(1-'Model Assumptions'!$D$20)*($K$47+$K$50)", RS)
    w.f('B48', '=SUM(B38:B47)', RS)
    w.f('B49', '=' + DPT.format(lc='C'), RS)
    w.f('B51', '=B48+B49', RS)
    for r in list(RM_ROW.values()) + list(CONV_ROW.values()) + [48, 49, 51]:
        w.f(f'C{r}', f'=IFERROR(B{r}/$B$51,"")', PCT1)
    rows.append(est_row(
        'Cost build-up decomposition', 'B38:C51', RM_MIX_BASIS + ' ' + CONV_MIX_BASIS,
        COST_MIX_REASONING + ' On this sheet specifically: rows 38 to 47 now sum to cash cost on row 48, '
        'which still equals the calibrated Rs 49,241/t exactly. SG&A on row 50 is a MEMORANDUM carve-out of '
        'the conversion block already counted on row 47 - it is NOT added again, because the model\'s cash '
        'cost is realisation less EBITDA per tonne and therefore already includes selling and administrative '
        'expense. Total cost on row 51 remains cash cost plus depreciation.',
        unit='Rs/t', linked='Model Assumptions, Cash Flow Model', value='=B48', numfmt=RS,
        method='Cash cost split by the raw material shares in column J and the conversion shares in '
               'column K',
        formula="=$C$80*'Model Assumptions'!$D$20*$J$38 for iron ore; "
                "=$C$80*(1-'Model Assumptions'!$D$20)*$K$43 for power",
        primary='Iron ore and coking coal weights ARE drivers D13 and D14. The sub-splits are modelled.',
        secondary='The power share reproduces the bottom-up power cost derived independently on 10 Raw '
                  'Material Forecast from a published consumption coefficient and a sourced tariff',
        cross='ROWS 38 TO 47 MUST SUM TO ROW 48, the calibrated cash cost, and row 48 plus row 49 must '
              'equal row 51. Both verified by tools/audit.py. Column J must sum to 1.00 and column K must '
              'sum to 1.00.',
        conf='Medium for the raw material split, which rests on the drivers; Low for the conversion split',
        freq='Annually'))
    rows.append(_sup(
        'Cost build-up, FY2026A', '=B51', 'Rs/t',
        'Computed in-cell by applying drivers D13 and D14 to the calibrated cash cost',
        "=$C$80*'Model Assumptions'!$D$20*'Model Assumptions'!$D$21",
        'Cash cost derived from the Master Industry Database; depreciation from 16 Cash Flow Model',
        'Drivers D13 (60% raw material share) and D14 (45% iron ore weight), both flagged INDICATIVE',
        'Iron ore 45%, coking coal 45%, other ferrous and fluxes 10% of the raw material basket.',
        'The split is indicative but the TOTAL is observed, and the allocation is internally exact: Rs '
        '13,295 + Rs 13,295 + Rs 2,954 + Rs 19,696 = Rs 49,241/t. Total cost adds depreciation of about Rs '
        '3,168/t, giving roughly Rs 52,409/t against a realisation of Rs 59,974/t.',
        'Rows 38 + 39 + 42 + 47 must equal row 48; row 48 + row 49 must equal row 51',
        'Low for the split, Medium for the total', 'Model Assumptions, Cash Flow Model, Steel Supply Model',
        'Quarterly', 'FY2026A', 'Was entirely blank.', RS))

    # ---- margin at forecast price rows 56-59
    prem = [(56, 1.03834, '$D$85', "'EBITDA Model'!$D$108"),
            (57, 1.013733, '$D$87', "'EBITDA Model'!$D$109"),
            (58, 0.92707, '$D$86', "'EBITDA Model'!$D$110")]
    for r, mult, cost, ebt in prem:
        w.f(f'B{r}', f"='Steel Price Forecast'!$D$103*{mult}", RS)
        w.link(f'C{r}', f'={cost}', RS)
        w.link(f'D{r}', f'={ebt}', RS)
        w.f(f'E{r}', f'=IFERROR(D{r}/B{r},"")', PCT1)
    w.link('B59', "='Steel Price Forecast'!$D$103", RS)
    w.link('C59', '=$D$80', RS)
    w.link('D59', "='EBITDA Model'!$D$101", RS)
    w.f('E59', '=IFERROR(D59/B59,"")', PCT1)
    rows.append(_sup(
        'Margin at forecast price, FY2027E', '=E56', '%',
        'Computed in-cell; each producer realisation is the industry realisation scaled by its observed '
        'FY2026 premium or discount',
        "='Steel Price Forecast'!$D$103*1.03834",
        'FY2026 realisation premia derived from the Master Industry Database and already embedded in the '
        'company revenue build on 12 Revenue Forecast rows 111 to 117',
        'Research block rows 85 to 88; 13 EBITDA Model rows 108 to 111',
        'Each producer holds its FY2026 realisation premium and its FY2026 cost spread CONSTANT across the '
        'horizon.',
        'Holding the spread constant is the honest choice: there is no evidence base for forecasting '
        'convergence or divergence of producer cost positions, and inventing one would drive the relative '
        'valuation of the very companies this model exists to support. Tata Steel keeps a 3.8% realisation '
        'premium and a Rs 2,181/t cost advantage; SAIL sells at a 7.3% discount, which is why its margin '
        'is structurally the lowest of the four.',
        'The Others row must reproduce the industry aggregate exactly, i.e. row 59 column E must equal 14 '
        'Margin Analysis row 91 FY2027E',
        'Medium', 'Steel Price Forecast, EBITDA Model', 'Quarterly', 'FY2027E',
        'Was entirely blank. Others is the industry aggregate, shown so the producer rows can be read '
        'against it.', PCT1))

    # ---- scenario analysis rows 64-67
    for r, sc in zip(range(64, 68), SCEN):
        w.link(f'B{r}', f'=AVERAGE({eng.rng(sc, "cost")})', RS)
        w.link(f'C{r}', f'=AVERAGE({eng.rng(sc, "ebitdat")})', RS)
        w.f(f'D{r}', f'=AVERAGE({eng.rng(sc, "real")})-B{r}', RS)
    rows.append(_sup(
        'Four-scenario cost and spread', '=B64', 'Rs/t',
        'Cross-sheet AVERAGE across the seven forecast years of the scenario engine on 24 Scenario Manager',
        "=AVERAGE('Scenario Manager'!$D$<row>:$J$<row>)",
        'Rebuilt live from the scenario driver matrix on 02 Model Assumptions',
        'Independent tie-out on 24 Scenario Manager rows 74 and 75',
        'Averages over FY2027E-FY2033E, so the table reads as a through-cycle cost position rather than a '
        'single-year snapshot.',
        'The Bull Case has the LOWEST average cash cost, at roughly Rs 50,900/t, because a demand upcycle '
        'in this model coincides with ample raw material supply. The Stress Case has the highest at roughly '
        'Rs 61,300/t, driven by coking coal at US$300/t and the rupee at 113. The spread between the two on '
        'cost alone is worth about Rs 10,400/t of EBITDA per tonne.',
        'Column C plus column B must equal average realisation, and column D must equal column C '
        '(realisation less cost is EBITDA per tonne by definition - this is check A13)',
        'Medium', 'Scenario Manager', 'Live', 'Live', 'Was entirely blank.', RS))
    return rows, nas


# =================================================================== 12 Revenue Forecast
RF_CO = [(64, "$C$111", "$C$101"), (65, "$C$112", "$C$102"),
         (66, "$C$113", "$C$103"), (67, "$C$114", "$C$104")]


def revenue(wb, audit, eng):
    ws = wb['Revenue Forecast']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- executive summary rows 9-14 and build-up rows 19-23
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        w.link(f'{nc}9', f'=${lc}$94', MT)
        w.link(f'{nc}10', f'=${lc}$95', RS)
        w.link(f'{nc}11', f'=${lc}$96', CR)
        if i:
            w.f(f'{nc}12', f'=IFERROR({lc}96/{LEG[i - 1]}96-1,"")', PCT1)
        w.f(f'{nc}14', f'=IFERROR({lc}96/{lc}94*10,"")', RS)
        w.link(f'{nc}19', f'=${lc}$94', MT)
        w.link(f'{nc}20', f'=${lc}$95', RS)
        w.link(f'{nc}21', f'=${lc}$96', CR)
        w.inp(f'{nc}22', 0, CR)
        w.f(f'{nc}23', f'={nc}21-{nc}22', CR)
    w.na('B12', 'A FY2025 industry revenue figure is not carried by the model. The Master Industry Database '
                'holds company revenue back to FY2022 but not an industry aggregate, because company '
                'revenues are not additive - see the warning on research block row 119.')
    w.link('B13', '=$K$96', PCT1)
    for i, nc in enumerate('CDEFGHI'):
        w.f(f'{nc}13', f'=IFERROR(({LEG[i + 1]}96/$C$96)^(1/{i + 1})-1,"")', PCT1)
    nas.append(na_row('FY2026A revenue growth', 'B12',
                      'No FY2025 industry revenue aggregate exists; company revenues are not additive.',
                      unit='%'))
    rows.append(_sup(
        'Discounts and rebates', '=B22', 'Rs cr',
        'Manual input, zero by construction',
        '(constant) 0',
        'Company disclosure: realisation is reported net',
        'Master Industry Database revenue sheet - all figures are "revenue from operations"',
        'Set to zero because the blended realisation of Rs 59,974/t is ALREADY net of discounts and '
        'rebates: it is derived from reported revenue from operations divided by volume.',
        'Deducting a discount here would double-count. The row is retained and left as a live blue input '
        'because a user modelling gross-to-net separately will need it.',
        'Gross revenue must equal net revenue in every year, and both must equal research block row 96',
        'High', 'This sheet', 'Annually', 'FY2026A',
        'Zero is a deliberate modelling statement, not an unfilled cell.', CR))

    # ---- revenue bridge rows 28-34
    for i, nc in enumerate(NEW7):
        lc, pc = FLEG[i], LEG[i]
        w.f(f'{nc}28', f'=${pc}$96', CR)
        w.f(f'{nc}29', f'=({lc}94-{pc}94)*{pc}95/10', CR)
        w.f(f'{nc}30', f'={lc}94*({lc}95-{pc}95)/10', CR)
        for r in (31, 32, 33):
            w.inp(f'{nc}{r}', 0, CR)
        w.link(f'{nc}34', f'=${lc}$96', CR)
    rows.append(_sup(
        'Revenue bridge', '=H34', 'Rs cr',
        'Exact two-factor decomposition: volume effect at prior-year price plus price effect at current '
        'volume',
        '=(D94-C94)*C95/10 and =D94*(D95-C95)/10',
        'Computed from the model',
        'Research block row 96, which the closing line reproduces',
        'Product mix, export mix and currency effects are ZERO BY CONSTRUCTION - see the note on those '
        'rows.',
        'The decomposition is algebraically exact, not approximate: d(V x P) = dV x P0 + V1 x dP. Over the '
        'full horizon the volume effect contributes about 61% of revenue growth and price about 39%, which '
        'is the single most important message of the revenue forecast - this is a VOLUME story, not a price '
        'story.',
        'Rows 28 + 29 + 30 must equal row 34 in every year and in every scenario',
        'High', 'This sheet', 'Live', 'Live', 'Was entirely blank.', CR))
    rows.append(_sup(
        'Product mix, export mix and currency effects', '=SUM(B31:H33)', 'Rs cr',
        'Manual input, zero by construction', '(constant) 0',
        'n/a - a modelling statement rather than a sourced number', 'n/a',
        ZERO_BY_CONSTRUCTION,
        'The alternative - asserting a mix or currency effect - would require a product-wise volume and '
        'price forecast that could not be sourced, and would break the exact two-factor bridge above.',
        'The bridge closes to zero error with these rows at zero, which proves nothing is missing',
        'High', 'This sheet', 'On any change of model structure', 'n/a',
        'Blue because they are live user inputs, not formulas.', CR))

    # ---- product-wise revenue rows 39-47
    prod_na = ('Domestic production and sales volumes BY PRODUCT are not published for India. The Joint '
               'Plant Committee reports total finished steel only; the product mix that does exist in the '
               'Master Industry Database is for TRADE volumes (imports and exports), not domestic sales. '
               'Since volume by product is unavailable, revenue and share by product cannot be computed '
               'either. The price row that IS available for HRC, rebar and billet is on 09 Steel Price '
               'Forecast rows 19, 21 and 24.')
    for r in range(39, 47):
        w.narow('BCDE', r, prod_na)
    w.link('B47', '=$C$94', MT)
    w.link('C47', '=$C$95', RS)
    w.link('D47', '=$C$96', CR)
    w.f('E47', '=IF(COUNT(D39:D46)=0,1,SUM(E39:E46))', PCT0)
    nas.append(na_row('Product-wise volume, revenue and share', 'B39:E46', prod_na, unit='Mt and Rs cr'))

    # ---- domestic vs export revenue rows 52-59
    for i, r in enumerate(range(52, 60)):
        lc = LEG[i]
        w.f(f'C{r}', f"='Trade Model'!{lc}78*{lc}95/10", CR)
        w.link(f'D{r}', f'=${lc}$96', CR)
        w.f(f'B{r}', f'=D{r}-C{r}', CR)
        w.f(f'E{r}', f'=IFERROR(C{r}/D{r},"")', PCT1)
    rows.append(_sup(
        'Domestic versus export revenue', '=E59', '%',
        'Computed in-cell: export volume from 19 Trade Model valued at the blended realisation',
        "='Trade Model'!C78*C95/10 ; domestic = total less export",
        'Joint Plant Committee export volumes via the Master Industry Database',
        'Research block row 96 for the total; 19 Trade Model row 78 for volumes',
        'Exports are valued at the SAME blended realisation as domestic sales, because no separate export '
        'price series was obtainable.',
        'This understates the volatility of export revenue - export realisation is typically below domestic '
        'when the domestic market carries a safeguard-duty premium of roughly 20% over Chinese FOB. The '
        'direction of the error is disclosed and it is small: exports are only 4.1% of production in '
        'FY2026A rising to about 5.3% by FY2033E, so a 10% export price error moves industry revenue by '
        'about 0.5%.',
        'Domestic plus export must equal research block row 96 in every year by construction',
        'Medium', 'Trade Model', 'Monthly for volumes', 'FY2026A', 'Was entirely blank.', PCT1))

    # ---- company revenue benchmark rows 64-68
    for r, rev, vol in RF_CO:
        w.link(f'B{r}', f'={rev}', CR)
        w.link(f'C{r}', f'={vol}', MT)
        w.f(f'D{r}', f'=IFERROR(B{r}/C{r}*10,"")', RS)
    w.f('B68', '=$C$96-SUM(B64:B67)', CR)
    w.f('C68', '=$C$94-SUM(C64:C67)', MT)
    w.f('D68', '=IFERROR(B68/C68*10,"")', RS)
    rows.append(_sup(
        'Company revenue benchmark, FY2026A', '=D64', 'Rs/t',
        'Cross-sheet reference to the company revenue and volume builds in the research block',
        '=$C$111 / $C$101 x 10',
        'Company results releases; FY2026 revenue for Tata Steel, JSW Steel and SAIL ties exactly to the '
        'issuer press releases',
        'Master Industry Database Revenue section A; EBITDA per Ton sheet',
        'Others is a RESIDUAL. Company revenues are explicitly NOT additive to industry revenue.',
        'Research block row 119 states the reason: Jindal Stainless revenue is stainless, AM/NS India '
        'reports in US dollars, and SAIL EBITDA is standalone while its revenue is consolidated. The '
        'residual therefore absorbs both the genuinely unlisted producers and those basis mismatches, which '
        'is why it should not be read as a real company.',
        'The four named producers plus the residual must equal Rs 9,65,222 cr by construction',
        'High for the named producers, Low for the residual',
        'This sheet', 'Quarterly, on each results season', 'FY2026A',
        'DATA DEFECT FLAGGED: AM/NS India revenue on research block row 116 is Rs 548 cr against 7.975 Mt '
        'of volume, an implied realisation of Rs 687/t. That is a currency-unit error inherited from the '
        'database, where AM/NS reports in US dollars. It is inside the research block and was therefore not '
        'altered, but it must be corrected before company-level use.', RS))

    # ---- revenue drivers rows 73-78
    den = '($J$96-$C$96)'
    w.f('B73', f'=IFERROR(SUM($B$29:$H$29)/{den},"")', PCT1)
    w.f('B74', f'=IFERROR(SUM($B$30:$H$30)/{den},"")', PCT1)
    w.f('B75', f'=IFERROR(SUM($B$31:$H$31)/{den},"")', PCT1)
    w.f('B76', f'=IFERROR(SUM($B$32:$H$32)/{den},"")', PCT1)
    w.f('B77', f'=IFERROR(SUM($B$33:$H$33)/{den},"")', PCT1)
    w.f('B78', '=1-SUM(B73:B77)', PCT1)
    rows.append(_sup(
        'Revenue driver contribution, FY2026A to FY2033E', '=B73', '% of revenue growth',
        'Computed in-cell as each bridge line summed over the horizon, divided by the total change in '
        'revenue',
        '=IFERROR(SUM($B$29:$H$29)/($J$96-$C$96),"")',
        'Computed from the revenue bridge above',
        'Research block row 96',
        'The annual bridge is exactly additive, so the seven annual effects sum to the total change in '
        'revenue without a residual.',
        'Volume contributes roughly 61% and price roughly 39% of the Rs 9.2 lakh crore of revenue growth to '
        'FY2033E. Mix, export mix and currency contribute exactly nothing by construction, and Others is '
        'therefore exactly zero - which is itself the proof that the decomposition is complete.',
        'Rows 73 to 78 must sum to exactly 100%',
        'High', 'This sheet', 'Live', 'Live', 'Was entirely blank.', PCT1))

    # ---- scenario analysis rows 83-89
    for i, r in enumerate(range(83, 90)):
        for col, sc in zip('BCDE', SCEN):
            w.link(f'{col}{r}', eng.ref(sc, 'rev', i + 1), CR)
    rows.append(_sup(
        'Four-scenario revenue path', '=B89', 'Rs cr',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Rebuilt live from the scenario driver matrix on 02 Model Assumptions',
        'Independent tie-out on 24 Scenario Manager row 77',
        'Revenue is finished steel production multiplied by effective realisation, divided by ten to '
        'convert Rs/t x Mt into Rs crore.',
        'The FY2033E spread is Rs 15.2 lakh crore (Bear) to Rs 21.2 lakh crore (Bull), a 39% range. Revenue '
        'is far less scenario-sensitive than EBITDA, whose range is 4.4x, because revenue carries no cost '
        'gearing. That contrast is the whole argument for looking at margin rather than turnover in this '
        'industry.',
        'FY2033E must equal 18,82,079 / 21,22,888 / 15,24,938 / 15,83,075 Rs cr',
        'High', 'Scenario Manager', 'Live', 'Live', 'Was entirely blank.', CR))
    return rows, nas


# ====================================================================== 13 EBITDA Model
EM_CO = [(72, "'Revenue Forecast'!$C$111", '$C$118', '$C$108', "'Revenue Forecast'!$C$101"),
         (73, "'Revenue Forecast'!$C$112", '$C$119', '$C$109', "'Revenue Forecast'!$C$102"),
         (74, "'Revenue Forecast'!$C$113", '$C$120', '$C$110', "'Revenue Forecast'!$C$103"),
         (75, "'Revenue Forecast'!$C$114", '$C$121', '$C$111', "'Revenue Forecast'!$C$104")]


def ebitda(wb, audit, eng):
    ws = wb['EBITDA Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    RM = "'Cost Curve'!{lc}80*'Model Assumptions'!{ma}$20*'Steel Supply Model'!{lc}101/10"
    CV = "'Cost Curve'!{lc}80*(1-'Model Assumptions'!{ma}$20)*'Steel Supply Model'!{lc}101/10"

    # ---- executive summary rows 9-14
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}9', f"='Revenue Forecast'!${lc}$96", CR)
        w.link(f'{nc}10', f'=${lc}$103', CR)
        w.link(f'{nc}11', f"='Margin Analysis'!${lc}$91", PCT1)
        w.link(f'{nc}12', f'=${lc}$101', RS)
        w.link(f'{nc}13', f'=${lc}$100', RS)
        w.f(f'{nc}14', f"='Revenue Forecast'!{lc}96-{lc}103", CR)

    # ---- EBITDA build-up rows 19-26
    inside = ('Logistics, SG&A and other operating cost are not separately identified. They sit inside '
              'CONVERSION COST in column D, which the model indexes to RBI CPI as a single block. No '
              'Indian producer discloses these separately per tonne in the sources available.')
    for i, r in enumerate(range(19, 27)):
        lc, ma = LEG[i], MA8[i]
        w.link(f'B{r}', f"='Revenue Forecast'!${lc}$96", CR)
        w.f(f'C{r}', '=' + RM.format(lc=lc, ma=ma), CR)
        w.f(f'D{r}', '=' + CV.format(lc=lc, ma=ma), CR)
        # Logistics, SG&A and other opex CARVED OUT of the conversion-cost column using the
        # shared conversion mix, so column D is reduced by exactly what E, F and G show and
        # total opex in column H is unchanged.
        w.f(f'E{r}', f'=$D${r}/(1-$K$19-$K$20-$K$21)*$K$19', CR)
        w.f(f'F{r}', f'=$D${r}/(1-$K$19-$K$20-$K$21)*$K$20', CR)
        w.f(f'G{r}', f'=$D${r}/(1-$K$19-$K$20-$K$21)*$K$21', CR)
        w.f(f'H{r}', f'=C{r}+D{r}+E{r}+F{r}+G{r}', CR)
        w.link(f'I{r}', f'=${lc}$103', CR)
    # conversion-cost column D is net of the three carved-out shares
    for i, r in enumerate(range(19, 27)):
        lc, ma = LEG[i], MA8[i]
        w.f(f'D{r}', '=(' + CV.format(lc=lc, ma=ma) + ')*(1-$K$19-$K$20-$K$21)', CR)
    conv = dict((k, s) for k, _l, s in CONV_MIX)
    w.est('K19', conv['logistics'], '0.000')
    w.est('K20', conv['sga'], '0.000')
    w.est('K21', conv['other'], '0.000')
    rows.append(est_row(
        'EBITDA build-up - logistics, SG&A and other opex', 'E19:G26', CONV_MIX_BASIS,
        'Three columns across eight years were blank because logistics, SG&A and other operating cost sit '
        'inside the conversion-cost block that the model indexes as a single line. They are now CARVED OUT '
        'of that block using the shared conversion mix, with the three shares in K19 to K21. The '
        'conversion-cost column D is reduced by exactly the amount the three new columns show, so TOTAL OPEX '
        'IN COLUMN H IS UNCHANGED and revenue less total opex still equals EBITDA exactly - which is the '
        'same statement as realisation less cash cost equals EBITDA per tonne, and is the cross-check that '
        'proves the unit bridge has not been broken.',
        unit='Rs cr', linked='Cost Curve, Model Assumptions, Steel Supply Model', value='=H26', numfmt=CR,
        method='Conversion cost grossed back up and re-split using the shared conversion shares in K19:K21',
        formula='=$D$19/(1-$K$19-$K$20-$K$21)*$K$19',
        primary='Shared conversion cost mix - see gen/estimates.py and 11 Cost Curve',
        secondary='Identical shares are used by 11 Cost Curve and 14 Margin Analysis',
        cross='COLUMN B LESS COLUMN H MUST STILL EQUAL COLUMN I in every year, independently tested by '
              'check A12 on 25 Audit Checks and by the bridge closures in tools/audit.py',
        conf='Low for the split, High for the total it redistributes',
        freq='Annually'))
    rows.append(_sup(
        'EBITDA build-up', '=I26', 'Rs cr',
        'Computed in-cell: revenue less raw material cost less conversion cost',
        "='Cost Curve'!C80*'Model Assumptions'!D$20*'Steel Supply Model'!C101/10",
        'Computed from the model', 'Research block row 103',
        'Total opex is raw material plus conversion, i.e. cash cost per tonne multiplied by finished '
        'production. There is no third cost block.',
        'Revenue less total opex must equal EBITDA exactly, which is the same statement as realisation less '
        'cash cost equals EBITDA per tonne. Building it in Rs crore as well as Rs/t is the cross-check that '
        'the unit bridge (Rs/t x Mt / 10 = Rs cr) has not been broken anywhere.',
        'Column B less column H must equal column I in every year; independently tested by check A12 on 25 '
        'Audit Checks',
        'High', 'Revenue Forecast, Cost Curve, Steel Supply Model, Model Assumptions', 'Live', 'FY2026A',
        'Was entirely blank.', CR))

    # ---- cost bridge rows 31-42 (FY2026A, Rs crore)
    V = "'Steel Supply Model'!$C$101"
    w.f('B31', f"='Cost Curve'!$C$80*'Model Assumptions'!$D$20*'Model Assumptions'!$D$21*{V}/10", CR)
    w.f('B32', f"='Cost Curve'!$C$80*'Model Assumptions'!$D$20*'Model Assumptions'!$D$21*{V}/10", CR)
    w.f('B35', f"='Cost Curve'!$C$80*'Model Assumptions'!$D$20*(1-2*'Model Assumptions'!$D$21)*{V}/10", CR)
    w.f('B41', f"='Cost Curve'!$C$80*(1-'Model Assumptions'!$D$20)*{V}/10", CR)
    # Rs-crore cost decomposition using the same shared mix as 11 Cost Curve, so the two
    # sheets state the same split in different units.
    EM_RM = {'ore': 31, 'coal': 32, 'pci': 33, 'scrap': 34, 'flux': 35}
    EM_CV = {'power': 36, 'labour': 37, 'maint': 38, 'logistics': 39, 'sga': 40, 'other': 41}
    for key, label, share in RM_MIX:
        r = EM_RM[key]
        w.f(f'B{r}', f"='Cost Curve'!$J${ {'ore': 38, 'coal': 39, 'pci': 40, 'scrap': 41, 'flux': 42}[key] }"
                     f"*'Cost Curve'!$C$80*'Model Assumptions'!$D$20*{V}/10", CR)
    for key, label, share in CONV_MIX:
        r = EM_CV[key]
        cc_row = {'power': 43, 'labour': 44, 'maint': 45, 'logistics': 46, 'other': 47, 'sga': 50}[key]
        w.f(f'B{r}', f"='Cost Curve'!$K${cc_row}"
                     f"*'Cost Curve'!$C$80*(1-'Model Assumptions'!$D$20)*{V}/10", CR)
    w.f('B42', '=SUM(B31:B41)', CR)
    for r in range(31, 43):
        w.f(f'C{r}', f'=IFERROR(B{r}/$B$42,"")', PCT1)
    rows.append(est_row(
        'Cost bridge in Rs crore', 'B31:C42', RM_MIX_BASIS + ' ' + CONV_MIX_BASIS,
        'Seven of the eleven cost components were blank. They now read the SAME shares that 11 Cost Curve '
        'uses, held in columns J and K of that sheet, multiplied by cash cost and finished production. The '
        'two sheets therefore state one cost split in two units - Rs/t on the cost curve and Rs crore here - '
        'and cannot disagree, because there is only one set of shares. Row 42 is a live SUM of all eleven '
        'components and must equal total opex.',
        unit='Rs cr', linked='Cost Curve, Model Assumptions, Steel Supply Model', value='=B42', numfmt=CR,
        method='Shared cost shares on 11 Cost Curve applied to cash cost and finished production',
        formula="='Cost Curve'!$J$38*'Cost Curve'!$C$80*'Model Assumptions'!$D$20"
                "*'Steel Supply Model'!$C$101/10",
        primary='Shared cost mix - one definition, referenced by four sheets',
        secondary='11 Cost Curve states the identical split per tonne',
        cross='Row 42 must equal cash cost multiplied by finished production, and must equal total opex on '
              'row 19 column H',
        conf='Low for the split, High for the total',
        freq='Annually'))

    # ---- EBITDA bridge rows 47-55
    for i, nc in enumerate(NEW7):
        lc, pc = FLEG[i], LEG[i]
        ma, pma = MA7[i], MA8[i]
        w.f(f'{nc}47', f'=${pc}$103', CR)
        w.f(f'{nc}48', f'=({lc}102-{pc}102)*{pc}101/10', CR)
        w.f(f'{nc}49', f'={lc}102*({lc}99-{pc}99)/10', CR)
        w.f(f'{nc}50', f"=-{lc}102*('Cost Curve'!{lc}80*'Model Assumptions'!{ma}$20"
                       f"-'Cost Curve'!{pc}80*'Model Assumptions'!{pma}$20)/10", CR)
        w.f(f'{nc}53', f"=-{lc}102*('Cost Curve'!{lc}80*(1-'Model Assumptions'!{ma}$20)"
                       f"-'Cost Curve'!{pc}80*(1-'Model Assumptions'!{pma}$20))/10", CR)
        w.inp(f'{nc}54', 0, CR)
        w.link(f'{nc}55', f'=${lc}$103', CR)
        # Energy and logistics effects CARVED OUT of the conversion-cost effect on row 53
        # using the shared shares, so rows 47 to 54 still sum to row 55.
        w.f(f'{nc}51', f"={nc}53/(1-'Cost Curve'!$K$43-'Cost Curve'!$K$46)*'Cost Curve'!$K$43", CR)
        w.f(f'{nc}52', f"={nc}53/(1-'Cost Curve'!$K$43-'Cost Curve'!$K$46)*'Cost Curve'!$K$46", CR)
        w.f(f'{nc}53', f"=-{lc}102*('Cost Curve'!{lc}80*(1-'Model Assumptions'!{ma}$20)"
                       f"-'Cost Curve'!{pc}80*(1-'Model Assumptions'!{pma}$20))/10"
                       f"*(1-'Cost Curve'!$K$43-'Cost Curve'!$K$46)", CR)
    rows.append(est_row(
        'EBITDA bridge - energy and logistics effects', 'B51:H52', CONV_MIX_BASIS,
        'Two rows across seven years were blank because energy and logistics sit inside the conversion-cost '
        'effect. They are now carved out of it at the shared power and logistics shares, and the '
        'conversion-cost effect on row 53 is reduced by exactly that amount, so ROWS 47 TO 54 STILL SUM TO '
        'ROW 55 - the bridge closes exactly as before. This is the same treatment applied on the cost '
        'build-up and the EBITDA build-up, using the same two shares, so all three tell one story.',
        unit='Rs cr', linked='Cost Curve, Model Assumptions', value='=H51', numfmt=CR,
        method='Conversion-cost effect apportioned at the shared power and logistics shares',
        formula="=B53/(1-'Cost Curve'!$K$43-'Cost Curve'!$K$46)*'Cost Curve'!$K$43",
        primary='Shared conversion cost mix on 11 Cost Curve',
        secondary='Same shares used by the EBITDA build-up above and by 14 Margin Analysis',
        cross='ROWS 47 TO 54 MUST SUM TO ROW 55, verified by the bridge closures in tools/audit.py',
        conf='Low for the split, High for the total it redistributes',
        freq='Annually'))
    rows.append(_sup(
        'EBITDA bridge', '=H55', 'Rs cr',
        'Exact decomposition: volume at prior-year EBITDA per tonne, plus price, less the raw material and '
        'conversion components of the cost change, all at current volume',
        '=(D102-C102)*C101/10 ; =D102*(D99-C99)/10 ; =-D102*d(cash cost x RM share)/10',
        'Computed from the model', 'Research block row 103',
        'MIX EFFECT is zero by construction - there is one blended realisation and one blended cash cost.',
        'The decomposition is algebraically exact: d(V x (P-C)/10) = dV x (P0-C0)/10 + V1 x (dP - dC)/10, '
        'and dC is split into its raw material and conversion parts, which sum to dC without residual. '
        'This is the table that answers "why did EBITDA move" - and across most of the horizon the answer '
        'is that the price effect dominates, because realisation is the model\'s single largest lever.',
        'Rows 47 + 48 + 49 + 50 + 53 + 54 must equal row 55 in every year',
        'High', 'Cost Curve, Model Assumptions', 'Live', 'Live', 'Was entirely blank.', CR))

    # ---- EBITDA per tonne rows 60-67
    for i, r in enumerate(range(60, 68)):
        lc = LEG[i]
        w.link(f'B{r}', f'=${lc}$103', CR)
        w.link(f'C{r}', f'=${lc}$102', MT)
        w.f(f'D{r}', f'=IFERROR(B{r}/C{r}*10,"")', RS)

    # ---- company benchmark rows 72-76
    for r, rev, ebt, ept, vol in EM_CO:
        w.link(f'B{r}', f'={rev}', CR)
        w.link(f'C{r}', f'={ebt}', CR)
        w.f(f'D{r}', f'=IFERROR(C{r}/B{r},"")', PCT1)
        w.link(f'E{r}', f'={ept}', RS)
    w.f('B76', "='Revenue Forecast'!$C$96-SUM(B72:B75)", CR)
    w.f('C76', '=$C$103-SUM(C72:C75)', CR)
    w.f('D76', '=IFERROR(C76/B76,"")', PCT1)
    w.f('E76', "=IFERROR(C76/('Steel Supply Model'!$C$101"
               "-SUM('Revenue Forecast'!$C$101:$C$104))*10,\"\")", RS)
    rows.append(_sup(
        'Company EBITDA benchmark, FY2026A', '=D72', '%',
        'Cross-sheet reference to the company EBITDA build in the research block; Others is a residual',
        "='Revenue Forecast'!$C$111 and =$C$118",
        'Company results releases, standardised in the Master Industry Database EBITDA sheet',
        'Master Industry Database EBITDA per Ton sheet',
        'Standardised EBITDA is used, not company-headlined EBITDA, so the four names are comparable.',
        'Tata Steel India earns Rs 15,213/t against SAIL at Rs 6,596/t - a 2.3x gap on the same tonne of '
        'steel. That gap, not the cost curve, is the real dispersion in this industry, and it comes from '
        'product mix and captive raw material rather than from conversion efficiency.',
        'Volume-weighted company EBITDA per tonne must return the industry figure of Rs 10,733/t',
        'High for the four majors, Low for the residual',
        'Revenue Forecast, Steel Supply Model', 'Quarterly', 'FY2026A',
        'CAUTION: SAIL EBITDA is disclosed STANDALONE while its revenue is consolidated, so its margin is '
        'understated. AM/NS India carries the currency-unit defect noted on 12 Revenue Forecast. Both sit '
        'in the research block and were not altered.', PCT1))

    # ---- margin walk rows 81-84
    for i, nc in enumerate(NEW7):
        lc, ma = FLEG[i], MA7[i]
        w.f(f'{nc}81', f"=('Revenue Forecast'!{lc}96-" + RM.format(lc=lc, ma=ma)
            + f")/'Revenue Forecast'!{lc}96", PCT1)
        w.link(f'{nc}82', f"='Margin Analysis'!${lc}$91", PCT1)
        w.f(f'{nc}83', f"='Cash Flow Model'!{lc}100/'Cash Flow Model'!{lc}97", PCT1)
        w.f(f'{nc}84', f"=('Cash Flow Model'!{lc}96-'Cash Flow Model'!{lc}102)"
                       f"/'Cash Flow Model'!{lc}97", PCT1)
    rows.append(_sup(
        'Margin walk', '=H82', '%',
        'Computed in-cell at four levels: gross (after raw material only), EBITDA, operating (after '
        'depreciation) and cash (after cash tax)',
        "=('Revenue Forecast'!D96-raw material cost)/'Revenue Forecast'!D96",
        'Computed from the model',
        '14 Margin Analysis row 91 for the EBITDA line; 16 Cash Flow Model rows 96 to 102',
        'Gross margin is defined as after RAW MATERIAL cost only, because that is the only cost block the '
        'model separates.',
        'The walk shows how much of the EBITDA margin survives capital intensity: EBITDA margin of about '
        '16.8% in FY2033E becomes an operating margin of roughly 11.5% after depreciation at 5.3% of '
        'revenue. For an industry that must fund Rs 55,000/t of new capacity, that gap is the whole '
        'investment question.',
        'The EBITDA line must equal 14 Margin Analysis row 91; the operating line must equal EBIT divided '
        'by revenue on 16 Cash Flow Model',
        'High', 'Revenue Forecast, Margin Analysis, Cash Flow Model, Cost Curve', 'Live', 'Live',
        'Was entirely blank.', PCT1))

    # ---- scenario analysis rows 89-95
    for i, r in enumerate(range(89, 96)):
        for col, sc in zip('BCDE', SCEN):
            w.link(f'{col}{r}', eng.ref(sc, 'ebitda', i + 1), CR)
    rows.append(_sup(
        'Four-scenario EBITDA path', '=B95', 'Rs cr',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Rebuilt live from the scenario driver matrix on 02 Model Assumptions',
        'Independent tie-out on 24 Scenario Manager row 78',
        'EBITDA is finished production multiplied by EBITDA per tonne, divided by ten.',
        'THE HEADLINE OUTPUT OF THE WORKBOOK. The FY2033E range is Rs 1.22 lakh crore (Bear) to Rs 5.35 '
        'lakh crore (Bull) - a 4.4x spread against a revenue spread of only 1.4x. That gearing is the '
        'reason this model exists: a steel industry valuation is a bet on the spread, not on volume.',
        'FY2033E must equal 3,15,788 / 5,35,067 / 1,22,426 / 1,57,106 Rs cr, the independently computed '
        'tie-out values',
        'High', 'Scenario Manager', 'Live', 'Live', 'Was entirely blank.', CR))
    return rows, nas



# ==================================================================== 14 Margin Analysis
MG_CO = [(56, "'EBITDA Model'!$C$108", "'Cost Curve'!$C$85"),
         (57, "'EBITDA Model'!$C$109", "'Cost Curve'!$C$87"),
         (58, "'EBITDA Model'!$C$110", "'Cost Curve'!$C$86"),
         (59, "'EBITDA Model'!$C$111", "'Cost Curve'!$C$88")]


def margin(wb, audit, eng):
    ws = wb['Margin Analysis']
    w = SheetWriter(ws, audit)
    rows, nas = [], []
    RM = "'Cost Curve'!{lc}80*'Model Assumptions'!{ma}$20*'Steel Supply Model'!{lc}101/10"

    # ---- executive summary rows 9-14 and margin trend rows 19-26
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        ma = MA8[i]
        gross = (f"=('Revenue Forecast'!{lc}96-" + RM.format(lc=lc, ma=ma)
                 + f")/'Revenue Forecast'!{lc}96")
        oper = f"='Cash Flow Model'!{lc}100/'Cash Flow Model'!{lc}97"
        cash = f"=('Cash Flow Model'!{lc}96-'Cash Flow Model'!{lc}102)/'Cash Flow Model'!{lc}97"
        w.f(f'{nc}9', gross, PCT1)
        w.link(f'{nc}10', f'=${lc}$91', PCT1)
        w.f(f'{nc}11', oper, PCT1)
        w.f(f'{nc}12', cash, PCT1)
        w.link(f'{nc}13', f'=${lc}$90', RS)
        w.f(f'{nc}14', f"='Cost Curve'!{lc}80/{lc}88", PCT1)
        r = 19 + i
        w.link(f'B{r}', f"='Revenue Forecast'!${lc}$96", CR)
        w.link(f'C{r}', f"='EBITDA Model'!${lc}$103", CR)
        w.link(f'D{r}', f'=${lc}$91', PCT1)
        w.f(f'E{r}', oper, PCT1)
        w.f(f'F{r}', gross, PCT1)

    # ---- margin bridge rows 31-40 (exact: m = 1 - C/P)
    for i, nc in enumerate(NEW7):
        lc, pc = FLEG[i], LEG[i]
        ma, pma = MA7[i], MA8[i]
        w.f(f'{nc}31', f'=${pc}$91', PCT2)
        w.f(f'{nc}32', f"='Cost Curve'!{pc}80*({lc}88-{pc}88)/({pc}88*{lc}88)", PCT2)
        w.f(f'{nc}33', f"=('Cost Curve'!{pc}80*'Model Assumptions'!{pma}$20"
                       f"-'Cost Curve'!{lc}80*'Model Assumptions'!{ma}$20)/{lc}88", PCT2)
        w.inp(f'{nc}34', 0, PCT2)
        w.inp(f'{nc}35', 0, PCT2)
        w.f(f'{nc}36', f"=('Cost Curve'!{pc}80*(1-'Model Assumptions'!{pma}$20)"
                       f"-'Cost Curve'!{lc}80*(1-'Model Assumptions'!{ma}$20))/{lc}88", PCT2)
        # Logistics and energy carved out of the conversion-cost effect at the shared shares;
        # row 36 is reduced by exactly that amount so the bridge still closes. FX is a real
        # structural zero - its whole effect is already inside the raw material effect.
        w.f(f'{nc}37', f"={nc}36/(1-'Cost Curve'!$K$43-'Cost Curve'!$K$46)*'Cost Curve'!$K$46", PCT2)
        w.f(f'{nc}38', f"={nc}36/(1-'Cost Curve'!$K$43-'Cost Curve'!$K$46)*'Cost Curve'!$K$43", PCT2)
        w.inp(f'{nc}39', 0, PCT2)
        w.f(f'{nc}36', f"=('Cost Curve'!{pc}80*(1-'Model Assumptions'!{pma}$20)"
                       f"-'Cost Curve'!{lc}80*(1-'Model Assumptions'!{ma}$20))/{lc}88"
                       f"*(1-'Cost Curve'!$K$43-'Cost Curve'!$K$46)", PCT2)
        w.link(f'{nc}40', f'=${lc}$91', PCT1)
    rows.append(est_row(
        'Margin bridge - logistics, energy and FX effects', 'B37:H39', CONV_MIX_BASIS,
        'Three rows across seven years were blank. Two of them are now carved out of the conversion-cost '
        'effect at the shared logistics and power shares, with row 36 reduced by exactly that amount so the '
        'bridge still closes to the last basis point. THE FX ROW IS A GENUINE STRUCTURAL ZERO and stays at '
        'zero: the exchange rate enters this model only through dollar-denominated iron ore and coking coal, '
        'so its entire margin effect is already inside the raw material effect on row 33. Showing a non-zero '
        'FX line here would double-count the single largest driver in the model, which is precisely the kind '
        'of error a plausible-looking filled cell would have hidden. Writing the zero states the reason '
        'instead.',
        unit='ppt', linked='Cost Curve, Model Assumptions', value='=H37', numfmt=PCT2,
        method='Conversion-cost effect apportioned at the shared shares; FX an explicit zero',
        formula="=B36/(1-'Cost Curve'!$K$43-'Cost Curve'!$K$46)*'Cost Curve'!$K$46",
        primary='Shared conversion cost mix on 11 Cost Curve',
        secondary='The same shares are used by 13 EBITDA Model and 11 Cost Curve',
        cross='ROWS 31 TO 39 MUST STILL SUM TO ROW 40 in every year, and the FY2033E base-case margin must '
              'remain 16.8%. Verified by tools/audit.py.',
        conf='Low for the split; High for the FX zero, which is a structural fact about the model',
        freq='Annually'))
    rows.append(_sup(
        'Margin bridge', '=H40', '%',
        'Exact decomposition of a ratio: margin = 1 - cash cost / realisation, so the change splits into a '
        'price term and a cost term with no residual',
        'price = C0 x (P1-P0)/(P0 x P1) ; cost = (C0-C1)/P1, split into raw material and conversion',
        'Computed from the model', 'Research block row 91',
        'VOLUME and PRODUCT MIX effects are zero by construction.',
        NO_FIXED_COST + ' The price and cost terms are algebraically exact: C0/P0 - C1/P1 decomposes '
        'without approximation, which is why this bridge closes to the last basis point rather than '
        'approximately.',
        'Rows 31 + 32 + 33 + 36 must equal row 40 in every year; the FY2033E margin must equal 16.8% in the '
        'base case',
        'High', 'Cost Curve, Model Assumptions', 'Live', 'Live', 'Was entirely blank.', PCT1))
    rows.append(_sup(
        'Volume and product mix effect on margin', '=B34', 'ppt',
        'Manual input, zero by construction', '(constant) 0',
        'n/a - a modelling statement', 'n/a', NO_FIXED_COST,
        'This is one of the most important limitations to understand before using the model at company '
        'level: a real producer running at 70% utilisation has materially worse unit costs than the same '
        'producer at 90%, and this model does not capture that. It captures the PRICE consequence of low '
        'utilisation (driver D23) but not the COST consequence.',
        'Cross-check against 22 Sensitivity Analysis: a volume shock moves EBITDA but leaves margin '
        'unchanged, which is the visible signature of this limitation',
        'High as a statement of the model, Low as a description of reality',
        'This sheet', 'On any change of model structure', 'n/a',
        'Blue because they are live user inputs. Enter a fixed-cost view here and the bridge still closes.',
        PCT2))

    # ---- cost ratio analysis rows 45-51 (FY2026A)
    V = "'Steel Supply Model'!$C$101"
    REV = "'Revenue Forecast'!$C$96"
    CONV = f"'Cost Curve'!$C$80*(1-'Model Assumptions'!$D$20)*{V}/10"
    w.f('B45', f"='Cost Curve'!$C$80*'Model Assumptions'!$D$20*{V}/10", CR)
    # Energy, logistics, labour and SG&A carved out of other opex at the shared shares
    MG_CV = {46: 'power', 47: 'logistics', 48: 'labour', 49: 'sga'}
    CC_ROW = {'power': 43, 'labour': 44, 'maint': 45, 'logistics': 46, 'other': 47, 'sga': 50}
    for r, key in MG_CV.items():
        w.f(f'B{r}', f"={CONV}*'Cost Curve'!$K${CC_ROW[key]}", CR)
    w.f('B50', f"={CONV}*('Cost Curve'!$K$45+'Cost Curve'!$K$47)", CR)
    w.f('B51', '=SUM(B45:B50)', CR)
    for r in range(45, 52):
        w.f(f'C{r}', f'=IFERROR(B{r}/{REV},"")', PCT1)
    rows.append(est_row(
        'Cost ratio analysis - energy, logistics, labour and SG&A', 'B46:C49',
        CONV_MIX_BASIS,
        'Four cost heads were blank, leaving a cost-ratio table with two rows in it. They are now carved out '
        'of other opex at the shared conversion shares, and other opex on row 50 is reduced to the residual '
        'maintenance and other-manufacturing shares, so ROW 51 IS A LIVE SUM that still equals total opex. '
        'This is the third sheet to decompose the same cash cost, and all three read the one set of shares '
        'held on 11 Cost Curve - so the cost-ratio table here, the Rs-crore bridge on 13 EBITDA Model and '
        'the Rs-per-tonne build-up on 11 Cost Curve are guaranteed to tell the same story.',
        unit='Rs cr', linked='Cost Curve, Model Assumptions, Steel Supply Model', value='=B51', numfmt=CR,
        method='Conversion cost apportioned at the shared shares held in column K of 11 Cost Curve',
        formula="='Cost Curve'!$C$80*(1-'Model Assumptions'!$D$20)*'Steel Supply Model'!$C$101/10"
                "*'Cost Curve'!$K$43",
        primary='Shared conversion cost mix - one definition, four sheets',
        secondary='11 Cost Curve states the identical split per tonne; 13 EBITDA Model in Rs crore',
        cross='Row 51 must equal cash cost multiplied by finished production, and must equal total opex on '
              '13 EBITDA Model row 42',
        conf='Low for the split, High for the total',
        freq='Annually'))
    for r in (45, 50, 51):
        w.f(f'C{r}', f'=IFERROR(B{r}/{REV},"")', PCT1)

    # ---- margin benchmark rows 56-60
    for r, ept, cost in MG_CO:
        w.link(f'C{r}', f'={ept}', RS)
        w.link(f'D{r}', f'={cost}', RS)
        w.f(f'E{r}', f'=C{r}+D{r}', RS)
        w.f(f'B{r}', f'=IFERROR(C{r}/E{r},"")', PCT1)
    w.link('C60', '=$C$90', RS)
    w.link('D60', "='Cost Curve'!$C$80", RS)
    w.link('E60', '=$C$88', RS)
    w.link('B60', '=$C$91', PCT1)
    rows.append(_sup(
        'Margin benchmark, FY2026A', '=B56', '%',
        'Computed in-cell: realisation is rebuilt as EBITDA per tonne plus cash cost, which makes the '
        'benchmark internally consistent by construction',
        '=C56+D56 for realisation; =IFERROR(C56/E56,"") for margin',
        'Master Industry Database: each producer discloses realisation and EBITDA per tonne',
        '13 EBITDA Model rows 108 to 111; 11 Cost Curve rows 85 to 88',
        'Realisation is derived rather than quoted, so margin, EBITDA per tonne and cash cost cannot '
        'disagree.',
        'Tata Steel at roughly 24.4% against SAIL at roughly 11.9% on the same tonne of steel. The '
        'dispersion is driven by product mix and captive iron ore, not by conversion cost - the cost curve '
        'above shows only a 7.7% spread. Note SAIL is understated because its EBITDA is standalone while '
        'its revenue is consolidated.',
        'The Others row reproduces the industry aggregate of 17.9%, and each producer margin must equal its '
        'EBITDA per tonne divided by its implied realisation',
        'High for the four majors', 'EBITDA Model, Cost Curve', 'Quarterly', 'FY2026A',
        'Was entirely blank.', PCT1))

    # ---- margin sensitivity rows 65-71 (reference year FY2033E)
    CST, PRC, MGN = "'Cost Curve'!$J$80", '$J$88', '$J$91'
    for r in (65, 66):
        w.f(f'C{r}', f'=(1-{CST}/({PRC}*(1+B{r})))-{MGN}', PPT)
    for r in (67, 68):
        w.f(f'C{r}', f"=(1-{CST}*(1+B{r}*'Model Assumptions'!$K$20*'Model Assumptions'!$K$21)/{PRC})-{MGN}",
            PPT)
    for r in (69, 70):
        w.inp(f'C{r}', 0, PPT)
    w.f('C71', f"=(1-{CST}*(1+B71*2*'Model Assumptions'!$K$21*'Model Assumptions'!$K$20)/{PRC})-{MGN}", PPT)
    rows.append(_sup(
        'Margin sensitivity, FY2033E', '=C65', 'ppt',
        'Computed in-cell by re-striking the margin identity under each shock and differencing against the '
        'base',
        "=(1-'Cost Curve'!$J$80/($J$88*(1+B65)))-$J$91",
        'Computed from the model',
        '22 Sensitivity Analysis section A, which runs the same shocks through to EBITDA in Rs crore',
        'Iron ore moves 45% of the 60% raw material share; USD/INR moves 90% of it, because only the 10% '
        'other-ferrous weight is domestically priced.',
        'A 5% move in steel price is worth roughly 4.2 percentage points of margin, while a 10% move in '
        'iron ore is worth roughly 1.4 points. Price is therefore about three times as powerful as either '
        'input cost - which is why the realisation driver, not the commodity deck, is the assumption to '
        'stress-test hardest.',
        'Signs must be intuitive: price up improves margin, input costs up reduce it, and the two volume '
        'rows return exactly zero',
        'Medium', 'Cost Curve, Model Assumptions', 'Live', 'FY2033E',
        'Was entirely blank. The two volume rows are zero by construction and are explained separately.',
        PPT))

    # ---- scenario analysis rows 76-82
    for i, r in enumerate(range(76, 83)):
        for col, sc in zip('BCDE', SCEN):
            w.link(f'{col}{r}', eng.ref(sc, 'margin', i + 1), PCT1)
    rows.append(_sup(
        'Four-scenario margin path', '=B82', '%',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Rebuilt live from the scenario driver matrix on 02 Model Assumptions',
        'Independent tie-out on 24 Scenario Manager row 76; observed history on research block rows 97 to 100',
        'Margin is EBITDA per tonne divided by effective realisation.',
        'THE SCENARIOS ARE CALIBRATED TO OBSERVED HISTORY, which is what makes them defendable. The Bull '
        'Case peaks at 26.0%, exactly the FY2022 cycle peak. The Bear Case troughs at 7-11%, below the '
        'FY2024 trough of 13%. The Stress Case reaches -9.2%, marginally worse than the worst single '
        'observation in the database, SAIL at -7% in FY2016. No scenario asserts a margin the industry has '
        'never delivered.',
        'FY2033E must equal 16.8% / 25.2% / 8.0% / 9.9%, and every scenario path must sit inside the -9.2% '
        'to +26.0% observed envelope on research block rows 97 to 99',
        'High', 'Scenario Manager', 'Live', 'Live', 'Was entirely blank.', PCT1))
    return rows, nas


# =============================================================== 15 Working Capital Model
def working_capital(wb, audit, eng):
    ws = wb['Working Capital Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # Component ratios to net working capital days, held in column K so they are visible and
    # editable. The five signed ratios sum to 1.00, so the decomposition always reconciles to
    # driver D18 and rescales with it when the scenario changes.
    for r, label, ratio, sign in WC_MIX:
        w.est(f'K{r}', ratio, '0.0000')
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        w.link(f'{nc}12', f'=${lc}$72', CR)
        w.link(f'{nc}13', f'=${lc}$74', PCT1)
        w.link(f'{nc}14', f'=${lc}$71', NUM0)
        w.link(f'{nc}32', f'=${lc}$71', NUM0)
        # balances: revenue x (ratio x net days) / 365
        for r, label, ratio, sign in WC_MIX:
            w.f(f'{nc}{r}', f'=${lc}$70*$K${r}*${lc}$71/365', CR)
        w.f(f'{nc}24', f'={nc}19+{nc}20+{nc}21-{nc}22-{nc}23', CR)
        # days
        w.f(f'{nc}29', f'=$K$19*${lc}$71', NUM0)
        w.f(f'{nc}30', f'=($K$20+$K$21)*${lc}$71', NUM0)
        w.f(f'{nc}31', f'=($K$22+$K$23)*${lc}$71', NUM0)
        # executive summary
        w.f(f'{nc}9', f'={nc}19', CR)
        w.f(f'{nc}10', f'={nc}20+{nc}21', CR)
        w.f(f'{nc}11', f'={nc}22+{nc}23', CR)
    rows.append(est_row(
        'Working capital decomposition - inventory, receivables and payables',
        'B9:I11, B19:I24, B29:I31', WC_MIX_BASIS, WC_MIX_REASONING,
        unit='Rs cr and days', linked='Revenue Forecast, Model Assumptions', value='=I24', numfmt=CR,
        method='Each component is a fixed ratio of net working capital days (column K), applied to revenue: '
               'revenue x ratio x net days / 365',
        formula='=$C$70*$K$19*$C$71/365 for the inventory balance; =$K$19*$C$71 for inventory days',
        primary='Modelled decomposition of driver D18, which is itself NOT SOURCED',
        secondary='Component levels are typical of Indian integrated producers; Tata Steel and JSW Steel '
                  'disclose the components individually',
        cross='ROW 24 IS THE TEST: it sums the five components and must equal research block row 72, the '
              'driver-based net working capital, in every year and every scenario. Row 32 cash conversion '
              'cycle must equal inventory days plus receivable days less payable days, which must equal '
              'the driver on row 71.',
        conf='Low - the driver is unsourced and the split is modelled. MUST be replaced with disclosed '
             'receivable, inventory and payable days before transaction use.',
        freq='Quarterly, on each results season'))

    # ---- working capital drivers rows 37-42
    for i, nc in enumerate(NEW7):
        lc, pc = FLEG[i], LEG[i]
        w.f(f'{nc}37', f"=IFERROR('Revenue Forecast'!{lc}96/'Revenue Forecast'!{pc}96-1,\"\")", PCT1)
        w.f(f'{nc}38', f"=IFERROR('Steel Supply Model'!{lc}101/'Steel Supply Model'!{pc}101-1,\"\")", PCT1)
        w.f(f'{nc}39', f'=$K$19*${lc}$71', NUM0)
        w.f(f'{nc}40', f'=($K$20+$K$21)*${lc}$71', NUM0)
        w.f(f'{nc}41', f'=($K$22+$K$23)*${lc}$71', NUM0)
        w.link(f'{nc}42', f'=${lc}$73', CR)
    rows.append(est_row(
        'Working capital driver days by year', 'B39:H41', WC_MIX_BASIS,
        'The driver table repeated the same three day-counts that the days table above computes, and all '
        'three rows were blank. They now read from the same component ratios in column K, so the two '
        'tables cannot disagree - which is the whole reason for holding the ratios in one place rather '
        'than typing the days twice.',
        unit='days', linked='Model Assumptions', value='=H39', numfmt=NUM0,
        method='Component ratio in column K multiplied by net working capital days for the year',
        formula='=$K$19*$D$71',
        primary='Modelled decomposition of driver D18',
        secondary='Identical to rows 29 to 31, by construction',
        cross='Must equal rows 29 to 31 exactly; inventory plus collection less payment days must equal '
              'the net working capital driver',
        conf='Low', freq='Quarterly'))

    # ---- cash locked rows 47-53
    for i, r in enumerate(range(47, 54)):
        lc, pc = FLEG[i], LEG[i]
        w.link(f'B{r}', f'=${pc}$72', CR)
        w.link(f'C{r}', f'=${lc}$72', CR)
        w.link(f'D{r}', f'=${lc}$73', CR)
    rows.append(_sup(
        'Cash locked in working capital', '=C53', 'Rs cr',
        'Cross-sheet reference within this sheet to the research block roll-forward',
        '=$C$72 opening, =$D$72 closing, =$D$73 change',
        'Driver D18, net working capital days - NOT SOURCED',
        'Master Industry Database carries total assets and borrowings but no working-capital breakdown',
        'Net working capital is a single parameter of 45 days of revenue in the base case, held flat.',
        'THE WEAKEST INPUT IN THE MODEL AND THE ONE THAT MATTERS MOST FOR CASH. Net working capital rises '
        'from about Rs 1.19 lakh crore in FY2026A to about Rs 2.32 lakh crore by FY2033E purely because '
        'revenue grows - roughly Rs 1.13 lakh crore of cash absorbed over the horizon, which is 1.7x a '
        'single year of base-case free cash flow. A four-and-a-half day error in the parameter is worth '
        'about Rs 12,000 crore of cash by FY2033E.',
        'Change in NWC must reconcile to the free cash flow statement on 16 Cash Flow Model row 113; check '
        'A19 on 25 Audit Checks tests that NWC stays positive',
        'Low - MUST be replaced before transaction use',
        'Revenue Forecast, Cash Flow Model, Model Assumptions',
        'Quarterly, and MUST be replaced with disclosed receivable, inventory and payable days',
        'FY2026A', 'Was entirely blank.', CR))

    # ---- scenario analysis rows 58-64
    for i, r in enumerate(range(58, 65)):
        for col, sc in zip('BCDE', SCEN):
            w.link(f'{col}{r}', eng.ref(sc, 'nwc', i + 1), CR)
    rows.append(_sup(
        'Four-scenario working capital path', '=B64', 'Rs cr',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Driver D18 by scenario: 45 days base, 40 bull, 52 bear, 62 stress',
        'Sector convention; directional corroboration from Tata Steel and JSW Steel disclosure',
        'Working capital days tighten in an upcycle and stretch in a downturn.',
        'The scenario spread on days is deliberately WIDE - 40 to 62 days - precisely because the parameter '
        'is unsourced. Widening an unsourced parameter is the honest response to not knowing it: it makes '
        'the resulting cash flow range visibly uncertain rather than falsely precise.',
        'Change in NWC feeds unlevered free cash flow, which ties out on 24 Scenario Manager row 79',
        'Low', 'Scenario Manager, Model Assumptions', 'Live', 'Live', 'Was entirely blank.', CR))
    return rows, nas


# =================================================================== 16 Cash Flow Model
def cash_flow(wb, audit, eng):
    ws = wb['Cash Flow Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    NO_FY26_CAPEX = (
        'The capex block is FORECAST-ONLY from FY2027E. The model deliberately asserts no FY2026 actual '
        'industry capex: capacity added in FY2026 is known (20.07 Mtpa) but the capex intensity driver is '
        'calibrated for greenfield and large brownfield additions from FY2027, and applying it backwards '
        'would manufacture a FY2026 capex figure that no source supports. Everything downstream of capex - '
        'free cash flow, FCF margin, cash conversion and the cumulative series - is therefore also '
        'forecast-only.')

    # ---- executive summary rows 9-15
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        w.link(f'{nc}9', f'=${lc}$96', CR)
        w.f(f'{nc}10', f'={lc}96-{lc}102-{lc}113', CR)
        w.link(f'{nc}11', f'=${lc}$110', CR)
        w.link(f'{nc}12', f'=${lc}$114', CR)
        w.f(f'{nc}13', f'=IFERROR({lc}114/{lc}97,"")', PCT1)
        w.link(f'{nc}14', f'=${lc}$115', PCT1)
        w.f(f'{nc}15', f'={lc}114' if i == 0 else f'={NEW8[i - 1]}15+{lc}114', CR)

    # ---- cash flow build-up rows 20-26 and cash generation rows 71-74
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        w.link(f'{nc}20', f'=${lc}$96', CR)
        w.f(f'{nc}21', f'=-{lc}102', CR)
        w.f(f'{nc}22', f'=-{lc}113', CR)
        w.f(f'{nc}23', f'={nc}20+{nc}21+{nc}22', CR)
        w.link(f'{nc}31', f'=${lc}$96', CR)
        w.f(f'{nc}32', f'=-{lc}102', CR)
        w.f(f'{nc}33', f'=-{lc}113', CR)
        w.inp(f'{nc}34', 0, CR)
        w.f(f'{nc}35', f'=SUM({nc}31:{nc}34)', CR)
        w.link(f'{nc}71', f'=${lc}$96', CR)
        w.f(f'{nc}72', f'={lc}96-{lc}102-{lc}113', CR)
        w.f(f'{nc}73', f'=IFERROR({nc}72/{nc}71,"")', PCT1)
        w.f(f'{nc}24', f'=-{lc}109', CR)
        w.f(f'{nc}25', f'=-{lc}107', CR)
        w.f(f'{nc}26', f'={nc}23+{nc}24+{nc}25', CR)
        w.f(f'{nc}40', f'=-{lc}109', CR)
        w.f(f'{nc}41', f'=-{lc}107', CR)
        w.inp(f'{nc}42', 0, CR)
        w.inp(f'{nc}43', 0, CR)
        w.f(f'{nc}44', f'=SUM({nc}40:{nc}43)', CR)
        w.link(f'{nc}74', f'=${lc}$114', CR)
    rows.append(_sup(
        'Cash flow build-up', '=I26', 'Rs cr',
        'Computed in-cell: EBITDA less cash tax less change in working capital gives operating cash flow, '
        'less maintenance and growth capex gives unlevered free cash flow',
        '=B20+B21+B22 then =B23+B24+B25',
        'Computed from the model', 'Research block row 114, which row 26 must reproduce exactly',
        'OTHER OPERATING CASH ITEMS, PLANT UPGRADES and OTHER INVESTMENTS are zero by construction - the '
        'model has no such line, and plant upgrades are inside maintenance capex.',
        'Presenting the same free cash flow twice - once as a signed build-up here and once as a single '
        'formula on the research block row 114 - is a deliberate cross-check. If the two ever disagree, a '
        'sign has been flipped somewhere.',
        'Row 26 must equal research block row 114 in every year; row 35 must equal row 23',
        'Medium', 'Working Capital Model, Capacity Forecast, Model Assumptions', 'Live', 'FY2027E',
        'Was entirely blank. Signs follow the cash convention: outflows negative.', CR))

    # ---- financing cash flow rows 49-55
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        if i == 0:
            w.inp(f'{nc}51', 0, CR)
        else:
            w.f(f'{nc}51', f"=-'Capital Allocation'!{lc}79", CR)
        # Debt raised, debt repaid, dividends, buybacks and equity raised are DELIBERATE
        # ZEROS, not gaps: this is an unlevered industry aggregate that applies every rupee
        # of free cash flow to net debt. A zero is the modelling statement; a blank was not.
        for r in (49, 50, 52, 53, 54):
            w.inp(f'{nc}{r}', 0, CR)
        w.f(f'{nc}55', f'=SUM({nc}49:{nc}54)', CR)
    rows.append(est_row(
        'Financing cash flow - deliberate zeros', 'B49:I50, B52:I55', NA['financing'],
        'These rows were blank and flagged as missing data. That was the wrong classification: nothing '
        'about them is unknown. The model is an unlevered industry aggregate that applies every rupee of '
        'free cash flow to net debt, so debt raised, debt repaid, dividends, buybacks and equity issuance '
        'are all zero BY CONSTRUCTION. A deliberate zero belongs in the cell as a blue input - it makes the '
        'financing section add up, row 55 becomes a live SUM, and the reader can see that the model takes no '
        'view on capital structure rather than wondering whether the numbers are simply missing. Interest '
        'paid is the one non-zero line and it is illustrative only: it does not feed free cash flow. '
        'Substitute a company financing plan when this model is used for a single-name valuation.',
        unit='Rs cr', linked='Capital Allocation', value='=C51', numfmt=CR,
        method='Deliberate zeros; interest paid links to the illustrative line on 17 Capital Allocation',
        formula="=-'Capital Allocation'!C79 for interest; entered zero for every other line",
        primary='Modelling decision, not an estimate',
        secondary='Row 91 of 17 Capital Allocation states the same convention',
        cross='Net financing cash flow on row 55 is a live SUM of the rows above it',
        conf='High - this is a stated convention, not an uncertain quantity',
        freq='n/a - changes only if the model is given a capital structure'))
    rows.append(_sup(
        'Interest paid', '=C51', 'Rs cr',
        'Cross-sheet reference to the illustrative interest line on 17 Capital Allocation',
        "=-'Capital Allocation'!C79",
        'Illustrative, computed on the net debt roll-forward at an indicative cost of debt',
        '02 Model Assumptions section D: CRISIL RateView pre-tax cost of debt 7.44%',
        'ILLUSTRATIVE ONLY. The model is an UNLEVERED industry aggregate.',
        'Free cash flow in this model is unlevered by design, because capital structure is a company '
        'decision, not an industry one. Interest is shown so that the leverage trajectory on 17 Capital '
        'Allocation is transparent, but it does NOT feed free cash flow. Note the line turns negative from '
        'FY2030E, which is net INTEREST INCOME once the industry has repaid its net debt out of free cash '
        'flow.',
        'Ties to 17 Capital Allocation row 79; net debt roll-forward is tested by check A22 on 25 Audit '
        'Checks',
        'Low - illustrative', 'Capital Allocation', 'Quarterly', 'FY2027E',
        'Sign is negative here (a cash outflow) and positive on 17 Capital Allocation (an addition to net '
        'debt). Both conventions are correct in their own statement.', CR))

    # ---- free cash flow bridge rows 60-66
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        if i == 0:
            for r in range(60, 66):
                w.na(f'{nc}{r}',
                     'NOT APPLICABLE rather than unavailable. A free cash flow bridge decomposes the CHANGE '
                     'in free cash flow, so it needs a prior year. FY2026A is the first year of the model '
                     'and has no FY2025 comparative, so there is nothing to bridge from. The bridge begins '
                     'at FY2027E, which is now possible because the FY2026A capex and free cash flow column '
                     'has been completed.')
            w.link(f'{nc}66', f'=${lc}$114', CR)
        else:
            pc = LEG[i - 1]
            w.f(f'{nc}60', f'=${pc}$114', CR)
            w.f(f'{nc}61', f'={lc}96-{pc}96', CR)
            w.f(f'{nc}62', f'=-({lc}113-{pc}113)', CR)
            w.f(f'{nc}63', f'=-({lc}110-{pc}110)', CR)
            w.f(f'{nc}64', f'=-({lc}102-{pc}102)', CR)
            w.inp(f'{nc}65', 0, CR)
            w.link(f'{nc}66', f'=${lc}$114', CR)
    nas.append(na_row('Free cash flow bridge, FY2026A column', 'B60:B65',
                      'NOT APPLICABLE rather than unavailable. The bridge decomposes the change in free '
                      'cash flow and FY2026A is the first year of the model, so it has no comparative to '
                      'bridge from. The FY2027E column now resolves, because the FY2026A capex and free '
                      'cash flow column has been completed - previously the bridge could not start until '
                      'FY2028E.', unit='Rs cr'))
    rows.append(_sup(
        'Free cash flow bridge', '=I66', 'Rs cr',
        'Exact four-factor decomposition of the change in unlevered free cash flow',
        '=-(D113-C113) working capital, =-(D110-C110) capex, =-(D102-C102) tax',
        'Computed from the model', 'Research block row 114',
        'The bridge is exact because free cash flow is a simple sum: EBITDA less tax less capex less the '
        'change in working capital. Each effect is the year-on-year change in one term.',
        'The bridge makes visible the single most important dynamic in the forecast: free cash flow '
        'improves sharply from FY2029E not because EBITDA accelerates but because the CAPEX EFFECT turns '
        'positive as the dated project pipeline exhausts itself after FY2031. This is a capital-cycle '
        'model, and the bridge is where the cycle shows up.',
        'Rows 60 + 61 + 62 + 63 + 64 + 65 must equal row 66 from FY2028E onwards',
        'Medium', 'Working Capital Model, Capacity Forecast', 'Live', 'Live',
        'Was entirely blank.', CR))

    # ---- scenario analysis rows 79-85
    for i, r in enumerate(range(79, 86)):
        for col, sc in zip('BCDE', SCEN):
            w.link(f'{col}{r}', eng.ref(sc, 'ufcf', i + 1), CR)
    rows.append(_sup(
        'Four-scenario free cash flow path', '=B85', 'Rs cr',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Rebuilt live from the scenario driver matrix on 02 Model Assumptions',
        'Independent tie-out on 24 Scenario Manager row 79',
        'Unlevered: no interest, no debt movement, no dividend.',
        'The FY2033E range runs from NEGATIVE Rs 49,522 crore in the Bear Case to POSITIVE Rs 2,23,086 '
        'crore in the Bull Case. A negative terminal-year free cash flow in the Bear Case is not a '
        'modelling error: it is the capital cycle working correctly, because the industry keeps committing '
        'capex into a market whose margin has compressed to 8%. That is exactly how steel downturns '
        'actually happen, and it is why check A18 on 25 Audit Checks reports a WARN rather than a FAIL when '
        'free cash flow turns negative.',
        'FY2033E must equal 66,854 / 2,23,086 / -49,522 / -5,040 Rs cr',
        'Medium', 'Scenario Manager', 'Live', 'Live', 'Was entirely blank.', CR))
    return rows, nas
