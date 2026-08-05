"""Fills for 17 Capital Allocation, 18 Industry Cycle Model, 19 Trade Model and 20 ESG Model."""
from openpyxl.styles import Font, Alignment

from .style import (SheetWriter, PCT0, PCT1, PCT2, NUM0, NUM1, NUM2, NUM3, MT, MTS, RS, CR,
                    X1, X2, USD, TEXT, SCORE, PPT)
from .support import na_row
from .spec import NA

LEG = 'CDEFGHIJ'
FLEG = 'DEFGHIJ'
NEW8 = 'BCDEFGHI'
NEW7 = 'BCDEFGH'
MA8 = 'DEFGHIJK'
MA7 = 'EFGHIJK'
SCEN = ['Base Case', 'Bull Case', 'Bear Case', 'Stress Case']


def _sup(item, value, unit, method, formula, primary, secondary, assumption, reasoning,
         cross, conf, linked, freq, last, comments, numfmt=None):
    return dict(item=item, value=value, unit=unit, method=method, formula=formula, primary=primary,
                secondary=secondary, assumption=assumption, reasoning=reasoning, cross=cross,
                conf=conf, linked=linked, freq=freq, last=last, comments=comments, numfmt=numfmt)


# =================================================================== 17 Capital Allocation
NO_DISTRIB = (
    'The model does not forecast dividends, buybacks or M&A at industry level. Those are capital-structure '
    'and portfolio decisions taken company by company, and aggregating them would imply a single industry '
    'payout policy that does not exist. Every rupee of unlevered free cash flow is applied to net debt, '
    'which is the conservative and auditable treatment. Substitute a company payout policy when this model '
    'is used as the foundation for a single-name valuation.')


def capital_allocation(wb, audit, eng):
    ws = wb['Capital Allocation']
    w = SheetWriter(ws, audit)
    rows, nas = [], []
    NO_FY26 = ('The capex and free cash flow blocks are forecast-only from FY2027E, so no FY2026A capital '
               'allocation is asserted. See 16 Cash Flow Model for the reason.')

    # ---- executive summary rows 9-16
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        if i == 0:
            for r in (9, 10, 11, 12, 16):
                w.na(f'{nc}{r}', NO_FY26)
        else:
            w.link(f'{nc}9', f"='Cash Flow Model'!${lc}$114", CR)
            w.link(f'{nc}10', f"='Cash Flow Model'!${lc}$107", CR)
            w.link(f'{nc}11', f"='Cash Flow Model'!${lc}$109", CR)
            w.f(f'{nc}12', f'={lc}77-{lc}80', CR)
            w.f(f'{nc}16', f'={lc}78-{lc}79-({lc}77-{lc}80)', CR)
        for r in (13, 14, 15):
            w.na(f'{nc}{r}', NO_DISTRIB)
    nas.append(na_row('Dividends, share buybacks and M&A investment', 'B13:I15', NO_DISTRIB, unit='Rs cr'))
    nas.append(na_row('FY2026A capital allocation', 'B9:B12, B16', NO_FY26, unit='Rs cr'))

    # ---- capital allocation waterfall rows 21-31
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        w.f(f'{nc}21', f"='Cash Flow Model'!{lc}96-'Cash Flow Model'!{lc}102-'Cash Flow Model'!{lc}113", CR)
        w.f(f'{nc}22', f"=-'Cash Flow Model'!{lc}107", CR)
        w.f(f'{nc}23', f"=-'Cash Flow Model'!{lc}109", CR)
        w.f(f'{nc}24', f'=-({lc}77-{lc}80)', CR)
        w.f(f'{nc}25', f'=-{lc}79', CR)
        for r in (26, 27, 28):
            w.na(f'{nc}{r}', NO_DISTRIB)
        w.na(f'{nc}29', NA['esg_capex'])
        w.f(f'{nc}30', f'=SUM({nc}21:{nc}25)', CR)
        w.f(f'{nc}31', f'={nc}30', CR)
    nas.append(na_row('ESG and decarbonisation investment in the waterfall', 'B29:H29', NA['esg_capex'],
                      unit='Rs cr'))
    rows.append(_sup(
        'Capital allocation waterfall', '=H30', 'Rs cr',
        'Computed in-cell: operating cash flow less capex less debt service, with the residual as cash '
        'retained',
        '=SUM(B21:B25)',
        'Computed from the model',
        'Research block rows 77 to 80, the net debt roll-forward',
        'CASH RETAINED IS EXACTLY ZERO IN EVERY YEAR, and that is a mathematical consequence, not an '
        'assumption: net debt is defined as rolling forward by opening less unlevered free cash flow plus '
        'interest, so applying all free cash flow to debt leaves nothing over by construction.',
        'A zero residual is the proof that the waterfall is complete - no cash is created or destroyed '
        'anywhere in the statement. It also makes the model conservative: every rupee services debt, so the '
        'leverage trajectory is the best case for the balance sheet and the worst case for shareholders.',
        'Rows 21 to 25 must sum to zero in every year and in every scenario; check A22 on 25 Audit Checks '
        'tests the net debt roll-forward independently',
        'Medium', 'Cash Flow Model', 'Live', 'Live',
        'Was entirely blank. Signs follow the cash convention: outflows negative.', CR))

    # ---- capex allocation rows 36-42
    for i, nc in enumerate(NEW7):
        lc, ma, n7 = FLEG[i], MA7[i], NEW7[i]
        w.link(f'{nc}36', f"='Cash Flow Model'!${lc}$109", CR)
        w.f(f'{nc}37', f"='Capacity Forecast'!{n7}52*'Model Assumptions'!{ma}$23/10", CR)
        w.f(f'{nc}38', f"=('Capacity Forecast'!{n7}56-'Capacity Forecast'!{n7}52)"
                       f"*'Model Assumptions'!{ma}$23/10", CR)
        for r in (39, 40, 41):
            w.na(f'{nc}{r}', 'Technology, ESG and safety capex are not separately identified. The model has '
                             'one capex intensity parameter of Rs 55,000 per tonne of capacity added, which '
                             'is an all-in greenfield and large-brownfield number, plus maintenance capex '
                             'at 3% of revenue. No Indian producer publishes a capex split by category to '
                             'FY2033.')
        w.link(f'{nc}42', f"='Cash Flow Model'!${lc}$110", CR)
    nas.append(na_row('Technology, ESG and safety capex', 'B39:H41',
                      'The model has a single all-in capex intensity parameter plus maintenance capex; no '
                      'producer publishes a capex split by category.', unit='Rs cr'))
    rows.append(_sup(
        'Capex allocation', '=H42', 'Rs cr',
        'Computed in-cell: maintenance from the research block, brownfield from the dated tracker via the '
        'capacity bridge, greenfield as the residual of net additions',
        "='Capacity Forecast'!B52*'Model Assumptions'!E$23/10",
        'Dated project tracker on 07 Capacity Expansion Tracker; capex intensity driver D16',
        '06 Capacity Forecast capacity bridge rows 51 to 57',
        'GREENFIELD here is net capacity additions less brownfield, so it also carries the secondary and '
        'unattributed additions of about 11 Mtpa a year.',
        'Defining greenfield as the residual is what makes the three rows sum to total capex exactly. The '
        'alternative - a separate line for secondary additions - does not exist in this table, and leaving '
        'the difference unallocated would break the reconciliation.',
        'Rows 36 + 37 + 38 must equal row 42 in every year, which must equal 16 Cash Flow Model row 110',
        'Medium', 'Capacity Forecast, Cash Flow Model, Model Assumptions', 'Quarterly', 'FY2027E',
        'Was entirely blank.', CR))

    # ---- debt allocation rows 47-53
    for i, r in enumerate(range(47, 54)):
        lc = FLEG[i]
        w.link(f'B{r}', f'=${lc}$77', CR)
        w.f(f'C{r}', f'=MAX(0,{lc}77-{lc}80)', CR)
        w.f(f'D{r}', f'=MAX(0,{lc}80-{lc}77)', CR)
        w.link(f'E{r}', f'=${lc}$80', CR)
        w.link(f'F{r}', f'=${lc}$84', X2)
    rows.append(_sup(
        'Debt allocation and leverage trajectory', '=E53', 'Rs cr',
        'Cross-sheet reference within this sheet to the net debt roll-forward, with repayment and drawdown '
        'separated by MAX so the direction of the movement is explicit',
        '=MAX(0,D77-D80) repaid, =MAX(0,D80-D77) drawn',
        'FY2026A closing net debt of Rs 1,85,001 cr derived from the Master Industry Database',
        'Research block rows 77 to 85; 25 Audit Checks A22',
        'All free cash flow after interest is applied to net debt. No new debt is raised for capex.',
        'The base case deleverages the industry from 1.07x net debt to EBITDA to a NET CASH position by '
        'about FY2030E, which is the single most important balance-sheet message in the model. It also '
        'explains why the illustrative interest line turns negative from FY2030E - that is interest income, '
        'not a sign error. In the Bear Case the industry re-levers instead, because free cash flow turns '
        'negative while capex continues.',
        'Opening less repaid plus drawn must equal closing in every year; leverage must equal net debt '
        'divided by EBITDA on research block row 84',
        'Medium', 'EBITDA Model, Cash Flow Model', 'Quarterly', 'FY2026A',
        'Was entirely blank.', CR))

    # ---- strategic investments rows 58-63
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        w.link(f'{nc}58', f"='Cash Flow Model'!${lc}$107", CR)
        for r in (59, 60, 61, 62, 63):
            w.na(f'{nc}{r}', 'Downstream, renewable energy, digital, acquisition and overseas investment '
                             'are not modelled. The model forecasts Indian crude steel capacity only. '
                             'Acquisitions are explicitly excluded from the project tracker (they transfer '
                             'ownership of existing capacity rather than add new capacity), and no Indian '
                             'producer publishes a costed plan for the other four categories to FY2033.')
    nas.append(na_row('Strategic investment other than capacity expansion', 'B59:H63',
                      'The model forecasts Indian crude steel capacity only; the other categories are not '
                      'costed by any producer to FY2033.', unit='Rs cr'))

    # ---- scenario analysis rows 68-71 (cumulative FY2027E-FY2033E)
    for r, sc in zip(range(68, 72), SCEN):
        w.link(f'B{r}', f'=SUM({eng.rng(sc, "gcapex")})', CR)
        w.link(f'C{r}', f'=SUM({eng.rng(sc, "ufcf")})', CR)
        w.na(f'D{r}', NO_DISTRIB)
        w.inp(f'E{r}', 0, CR)
    nas.append(na_row('Dividends by scenario', 'D68:D71', NO_DISTRIB, unit='Rs cr'))
    rows.append(_sup(
        'Four-scenario capital allocation, cumulative FY2027E-FY2033E', '=C68', 'Rs cr',
        'Cross-sheet SUM across the seven forecast years of the scenario engine on 24 Scenario Manager',
        "=SUM('Scenario Manager'!$D$<row>:$J$<row>)",
        'Rebuilt live from the scenario driver matrix on 02 Model Assumptions',
        'Independent tie-out on 24 Scenario Manager row 79',
        'Cumulative rather than single-year, because capital allocation is a through-cycle question.',
        'The Bull Case generates roughly Rs 9.5 lakh crore of cumulative free cash flow against roughly Rs '
        '5.0 lakh crore of growth capex - comfortable self-funding. The Bear Case generates NEGATIVE '
        'cumulative free cash flow while still committing roughly Rs 3.5 lakh crore of growth capex, which '
        'must be debt-funded. That divergence, not the EBITDA number, is what determines whether the '
        'industry is investable through the cycle.',
        'Cash retained is exactly zero in every scenario by construction',
        'Medium', 'Scenario Manager', 'Live', 'Live', 'Was entirely blank.', CR))
    return rows, nas


# ================================================================= 18 Industry Cycle Model
# indicator -> (0-100 score formula template, band description)
CYCLE_SCORES = [
    (19, "=MEDIAN(0,100,('Macroeconomic Model'!{lc}101-0.04)/0.05*100)",
     'Real GDP growth on a 4.0% to 9.0% band'),
    (20, "=MEDIAN(0,100,('Steel Demand Model'!{lc}115-0.02)/0.10*100)",
     'Apparent consumption growth on a 2% to 12% band, which brackets the FY2021 contraction recovery and '
     'the FY2024 peak of 13.7%'),
    (21, "=MEDIAN(0,100,('Capacity Utilisation'!{lc}67-0.65)/0.27*100)",
     'Capacity utilisation on a 65% to 92% band, the floor of the observed range to the practical ceiling'),
    (22, "=MEDIAN(0,100,100-('Capacity Forecast'!{lc}101-'Steel Supply Model'!{lc}94)"
         "/'Capacity Forecast'!{lc}101/0.30*100)",
     'Spare capacity against demand-driven requirement, 0% to 30%, scored INVERSELY'),
    (23, "=MEDIAN(0,100,('Steel Price Forecast'!{lc}103-50000)/25000*100)",
     'Blended realisation on a Rs 50,000/t to Rs 75,000/t band'),
    (24, "=MEDIAN(0,100,('Margin Analysis'!{lc}91+0.10)/0.36*100)",
     'Industry EBITDA margin on a -10% to +26% band, which is exactly the observed envelope from SAIL '
     'FY2016 to the FY2022 cycle peak'),
    (25, "=MEDIAN(0,100,'Capacity Forecast'!{lc}100/25*100)",
     'Capacity additions on a 0 to 25 Mtpa band, against a demonstrated national delivery of about 20 Mtpa'),
    (26, "=MEDIAN(0,100,50-'Steel Supply Model'!{lc}91/'Steel Demand Model'!{lc}116*1000)",
     'Variation in stock as a share of consumption, centred on 50 - a destock scores above 50 because it '
     'signals a tight market'),
]
PHASE = ('=IF({c}>=80,"Overheating",IF({c}>=60,"Peak",IF({c}>=40,"Expansion",'
         'IF({c}>=20,"Recovery","Deep Downturn"))))')


def cycle(wb, audit, eng):
    ws = wb['Industry Cycle Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- scorecard rows 19-26 (row 27 already SUMPRODUCTs)
    for r, tmpl, _desc in CYCLE_SCORES:
        for cyc, lc in zip(LEG, LEG):
            w.f(f'{cyc}{r}', tmpl.format(lc=lc), SCORE)

    # ---- executive summary rows 9-14
    for nc, lc in zip(NEW8, LEG):
        w.f(f'{nc}10', f'=${lc}$27', SCORE)
        w.f(f'{nc}9', PHASE.format(c=f'{lc}27'), TEXT)
        w.f(f'{nc}11', f'=IF({lc}27>=80,"Elevated - late cycle risk",IF({lc}27>=60,"Moderate",'
                       f'IF({lc}27>=40,"Balanced",IF({lc}27>=20,"High","Severe"))))', TEXT)
        w.f(f'{nc}12', f"=IF('Capacity Utilisation'!{lc}69>0.03,\"Strong\","
                       f"IF('Capacity Utilisation'!{lc}69>0,\"Improving\","
                       f"IF('Capacity Utilisation'!{lc}69>-0.04,\"Limited\",\"Weak\")))", TEXT)
        w.f(f'{nc}13', f"=IF('Margin Analysis'!{lc}91>{lc}90+0.04,\"Above mid-cycle\","
                       f"IF('Margin Analysis'!{lc}91>{lc}90,\"At mid-cycle\","
                       f"IF('Margin Analysis'!{lc}91>{lc}90-0.05,\"Below mid-cycle\",\"Trough\")))", TEXT)
        w.f(f'{nc}14', f"=IF({lc}27>=60,\"Late - discipline required\","
                       f"IF({lc}27>=40,\"Selective\",IF({lc}27>=20,\"Accumulate\",\"Deep value\")))", TEXT)
    rows.append(_sup(
        'Cycle score and phase', '=I10', 'score 0-100',
        'Weighted SUMPRODUCT of eight indicator scores, each normalised onto a 0-100 band computed live '
        'from the model',
        '=SUMPRODUCT($B$19:$B$26,C19:C26); each indicator =MEDIAN(0,100,(x-floor)/(range)*100)',
        'Computed from the model. Every band is stated inside its own formula so it is auditable',
        'Research block rows 95 to 98, the observed cycle chronology from the Master Industry Database',
        'The BANDS are the modeller\'s framework; the INPUTS are all model outputs, and the margin band of '
        '-10% to +26% is taken directly from observed history rather than chosen.',
        'MEDIAN(0,100,x) is used deliberately as the clamp rather than MIN/MAX nesting: it is a single '
        'transparent function, it cannot return an out-of-range score, and it keeps each indicator formula '
        'readable in the formula bar. Anchoring the margin and utilisation bands on the observed range is '
        'what stops the scorecard from being arbitrary.',
        'Weights in column B sum to 1.00; the score must classify FY2026A as the recovery the research '
        'block documents on row 97',
        'Medium', 'Macroeconomic Model, Steel Demand Model, Capacity Utilisation, Capacity Forecast, '
        'Steel Price Forecast, Margin Analysis, Steel Supply Model',
        'Live', 'Live', 'Rows 19 to 26 and the whole executive summary were blank.', SCORE))

    # ---- leading indicators rows 41-47
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}41', f"='Macroeconomic Model'!${lc}$101", PCT1)
    for r in range(42, 48):
        w.narow(NEW8, r, NA['macro_sub'])
    nas.append(na_row('Leading indicators other than GDP', 'B42:I47', NA['macro_sub']))

    # ---- coincident indicators rows 52-56
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}52', f"='Steel Supply Model'!${lc}$101", MT)
        w.link(f'{nc}53', f"='Steel Demand Model'!${lc}$116", MT)
        w.link(f'{nc}54', f"='Capacity Utilisation'!${lc}$67", PCT1)
        w.link(f'{nc}55', f"='Steel Price Forecast'!${lc}$103", RS)
        w.link(f'{nc}56', f"='EBITDA Model'!${lc}$101", RS)

    # ---- lagging indicators rows 61-65
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        w.link(f'{nc}61', f"='Capacity Forecast'!${lc}$100", MT)
        w.link(f'{nc}62', f"='Capital Allocation'!${lc}$80", CR)
        if i:
            w.link(f'{nc}64', f"='Capacity Expansion Tracker'!{MA7[i - 1]}$99", MT)
        else:
            w.na(f'{nc}64', 'The dated project tracker covers FY2027E onwards only. Every project with a '
                            'FY2026 commissioning date is already inside the 220.4 Mtpa base.')
        w.na(f'{nc}63', 'Dividend payout is not modelled at industry level - see 17 Capital Allocation.')
        w.na(f'{nc}65', 'Industry employment is not carried in the Master Industry Database and no '
                        'consistent series was retrievable. Company BRSR filings disclose headcount but '
                        'were not obtained in this research cycle.')
    nas.append(na_row('Dividend payout and employment', 'B63:I63, B65:I65',
                      'Dividends are not modelled at industry level; employment was not retrievable.'))

    # ---- scenario analysis rows 70-71
    parts = [
        ("MEDIAN(0,100,({gdp}-0.04)/0.05*100)", 19),
        ("MEDIAN(0,100,({consg}-0.02)/0.10*100)", 20),
        ("MEDIAN(0,100,({util}-0.65)/0.27*100)", 21),
        ("MEDIAN(0,100,100-({cap}-{crudeu})/{cap}/0.30*100)", 22),
        ("MEDIAN(0,100,({real}-50000)/25000*100)", 23),
        ("MEDIAN(0,100,({margin}+0.10)/0.36*100)", 24),
        ("MEDIAN(0,100,{adds}/25*100)", 25),
        ("MEDIAN(0,100,50-{stock}/{cons}*1000)", 26),
    ]
    for col, sc in zip('BCDE', SCEN):
        terms = []
        for tmpl, wrow in parts:
            ref = {k: eng.cell(sc, k, 7) for k in
                   ('gdp', 'consg', 'util', 'cap', 'crudeu', 'real', 'margin', 'adds', 'stock', 'cons')}
            terms.append(f'$B${wrow}*' + tmpl.format(**ref))
        w.link(f'{col}70', '=' + '+'.join(terms), SCORE)
        w.f(f'{col}71', PHASE.format(c=f'{col}70'), TEXT)
    rows.append(_sup(
        'Four-scenario cycle score, FY2033E', '=B70', 'score 0-100',
        'The same eight-indicator weighted score, re-struck on the scenario engine outputs for FY2033E',
        '=$B$19*MEDIAN(0,100,(engine GDP-0.04)/0.05*100)+... for all eight indicators',
        'Rebuilt live from the scenario engine on 24 Scenario Manager',
        'The scorecard above, which uses identical bands and weights on the live model',
        'Identical bands and weights to the live scorecard, so the four scenarios and the live model are '
        'measured on exactly the same ruler.',
        'This is the table that answers the only question an investment committee really asks: where in the '
        'cycle does each scenario leave us in the terminal year? The Bull Case reaches Peak, the base case '
        'sits in Expansion, and both Bear and Stress land in Recovery or below - which is the correct '
        'signal that a terminal-year exit multiple should NOT be a peak multiple in those states.',
        'Must use the same weights as row 27 and the same bands as rows 19 to 26; the base-case column must '
        'be close to the live FY2033E score, and identical when the Base Case is selected',
        'Medium', 'Scenario Manager', 'Live', 'FY2033E', 'Was blank.', SCORE))

    # ---- valuation implications rows 76-80
    mult = {76: '$K$113', 77: '$K$112', 78: '$K$110', 79: '$K$111'}
    view = {76: 'Maximum operational and financial stress. Historically the best entry point, but only for '
                'balance sheets that survive it - the Stress Case takes industry margin to -9.2%.',
            77: 'Margin recovering off a trough with capacity discipline improving. Accumulate quality '
                'cost-curve positions.',
            78: 'The base case position for most of the horizon. Selective - returns come from cost '
                'position and product mix, not from the cycle.',
            79: 'Utilisation tight and margin at or above the FY2022 observed peak. Discipline required: '
                'this is when boards approve the capex that creates the next downcycle.',
            80: 'Not reached in any scenario in this model. Retained for completeness.'}
    for r in (76, 77, 78, 79):
        w.link(f'B{r}', f"='Model Assumptions'!{mult[r]}", X1)
    w.na('B80', 'No scenario in this model reaches an Overheating cycle score, so no exit multiple is '
                'asserted for that phase. The four multiples above are the model\'s own scenario exit '
                'multiples, driver D19.')
    for r in range(76, 81):
        for col in 'CD':
            w.na(f'{col}{r}', NA['share_price'])
        c = ws.cell(r, 5, view[r])
        c.font = Font(name='Calibri', sz=8, color='FF404040')
        c.alignment = Alignment(wrap_text=True, vertical='top')
    nas.append(na_row('Price/book and price/earnings by cycle phase', 'C76:D80', NA['share_price'],
                      unit='x'))
    rows.append(_sup(
        'Exit multiple by cycle phase', '=B78', 'x EV/EBITDA',
        'Cross-sheet reference to the four scenario values of driver D19 on 02 Model Assumptions',
        "='Model Assumptions'!$K$110 (Base 6.0x) .. $K$113 (Stress 4.0x)",
        'NOT SOURCED - market multiples were not obtained in this research cycle',
        'Standing limitation 6 on 25 Audit Checks',
        'The four scenario exit multiples are MAPPED onto the four cycle phases: Stress 4.0x to Deep '
        'Downturn, Bear 5.0x to Recovery, Base 6.0x to Expansion, Bull 7.0x to Peak.',
        'The mapping is not arbitrary: each scenario already implies a cycle position, so re-using its own '
        'exit multiple keeps the valuation internally consistent instead of introducing a fifth set of '
        'numbers. It also makes the circularity explicit - the multiple is an assumption, not an '
        'observation, and until traded multiples are entered on 23 Comparable Valuation it should be '
        'treated as a placeholder for judgement rather than evidence.',
        'CONTRADICTION FLAGGED: the DCF on 23 Comparable Valuation cell D135 hard-codes 7.0x while driver '
        'D19 base case is 6.0x. That cell sits inside the research block and was NOT altered. It must be '
        'repointed to Model Assumptions before transaction use.',
        'Low', 'Model Assumptions', 'Quarterly', 'n/a',
        'Overheating is deliberately left blank - no scenario reaches it.', X1))
    return rows, nas



# ========================================================================= 19 Trade Model
TRADE_HIST = [(19, 'FY2021', 4.752, 10.784), (20, 'FY2022', 4.669, 13.494),
              (21, 'FY2023', 6.022, 6.716), (22, 'FY2024', 8.320, 7.487),
              (23, 'FY2025', 9.551, 4.858), (24, 'FY2026', 6.524, 6.602)]
EXPORT_DEST = {30: ('Europe', 0.344, "Italy 16.2% + Belgium 10.9% + Spain 7.3%"),
               31: ('Middle East', 0.074, 'UAE 7.4%'),
               32: ('Southeast Asia', 0.117, 'Vietnam 11.7%')}
IMPORT_SRC = {30: ('China', 0.235), 31: ('Japan', 0.202), 32: ('South Korea', 0.354)}


def trade(wb, audit, eng):
    ws = wb['Trade Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- executive summary rows 9-14
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}9', f'=${lc}$76', MT)
        w.link(f'{nc}10', f'=${lc}$78', MT)
        w.link(f'{nc}11', f'=${lc}$85', MTS)
        w.link(f'{nc}12', f'=${lc}$83', PCT1)
        w.link(f'{nc}13', f'=${lc}$84', PCT1)
        w.f(f'{nc}14', f"=({lc}78-{lc}76)*'Steel Price Forecast'!{lc}103/10", CR)

    # ---- historical trade rows 19-24
    for r, fy, imp, exp in TRADE_HIST:
        w.db(f'B{r}', imp, NUM3, ref=f"='[Master Industry Database.xlsx]Imports & Exports'!$C${r + 7}")
        w.db(f'C{r}', exp, NUM3, ref=f"='[Master Industry Database.xlsx]Imports & Exports'!$D${r + 7}")
        w.f(f'D{r}', f'=C{r}-B{r}', MTS)
        if r == 19:
            for col in 'EF':
                w.na(f'{col}{r}', 'FY2020 is outside the six-year window imported here, so a FY2021 growth '
                                  'rate cannot be computed. The Master Industry Database carries FY2015 '
                                  'onwards - FY2020 imports were 6.768 Mt and exports 8.355 Mt.')
        else:
            w.f(f'E{r}', f'=IFERROR(B{r}/B{r - 1}-1,"")', PCT1)
            w.f(f'F{r}', f'=IFERROR(C{r}/C{r - 1}-1,"")', PCT1)
    rows.append(_sup(
        'Historical trade FY2021-FY2026', '=D24', 'Mt net',
        'Master Industry Database import (RED - to be reconnected as an external link)',
        "='[Master Industry Database.xlsx]Imports & Exports'!$C$26:$D$31",
        'Joint Plant Committee, published by the Ministry of Steel',
        'DGCIS; Master Industry Database source S02',
        'Finished steel only. Semi-finished exports are reported separately and are excluded.',
        'THE MOST VOLATILE SERIES IN THE ENTIRE MODEL. India swung from a record net export of +8.8 Mt in '
        'FY2022 to a net IMPORT of -4.7 Mt in FY2025, then back to a marginal net export of +0.078 Mt in '
        'FY2026 as the safeguard duty took effect. A 13.5 Mt swing in four years on a 160 Mt market is why '
        'net trade is a scenario driver (D07) rather than a forecast.',
        'FY2026 imports of 6.524 Mt are down 31.7% year on year and exports up 35.9%; both are consistent '
        'with the April-2026 safeguard duty of 12%',
        'High', 'Master Industry Database', 'Monthly, on each Ministry of Steel report',
        'FY2026 (provisional)', 'Six years of history were blank.', MTS))

    # ---- trade forecast rows 19-25 (columns I..L)
    for i, r in enumerate(range(19, 26)):
        lc = FLEG[i]
        w.link(f'I{r}', f'=${lc}$76', MT)
        w.link(f'J{r}', f'=${lc}$78', MT)
        w.f(f'K{r}', f'=I{r}-J{r}', MTS)
        w.f(f'L{r}', f"=(J{r}-I{r})*'Steel Price Forecast'!{lc}103/10", CR)

    # ---- export destination and import source rows 30-35
    for r, (name, share, note) in EXPORT_DEST.items():
        w.db(f'D{r}', share, PCT1, ref=f"='[Master Industry Database.xlsx]Imports & Exports'! section C "
                                      f"({note})")
        w.f(f'B{r}', f'=$C$78*D{r}', NUM3)
        w.f(f'C{r}', f'=$J$78*D{r}', NUM3)
    for r in (33, 34):
        w.narow('BCD', r, 'North American and African destination shares are not separately disclosed by the '
                          'Ministry of Steel. The top-ten destination list for FY2026 comprises Italy, '
                          'Vietnam, Belgium, UAE, Spain, Nepal and Taiwan; neither region appears. Any '
                          'volume to those regions sits inside Others on row 35.')
    w.f('D35', '=1-SUM(D30:D34)', PCT1)
    w.f('B35', '=$C$78*D35', NUM3)
    w.f('C35', '=$J$78*D35', NUM3)
    for r, (name, share) in IMPORT_SRC.items():
        w.db(f'I{r}', share, PCT1, ref=f"='[Master Industry Database.xlsx]Imports & Exports'!$F$52:$F$56 "
                                      f"({name})")
        w.f(f'G{r}', f'=$C$76*I{r}', NUM3)
        w.f(f'H{r}', f'=$J$76*I{r}', NUM3)
    for r in (33, 34):
        w.narow('GHI', r, 'Vietnam was 7.3% of imports in FY2025 but is NOT separately disclosed for '
                          'FY2026; Indonesia is not disclosed in either year. Korea, China and Japan '
                          'together were 79.1% of FY2026 imports, so both sit inside the Others residual '
                          'on row 35.')
    w.f('I35', '=1-SUM(I30:I34)', PCT1)
    w.f('G35', '=$C$76*I35', NUM3)
    w.f('H35', '=$J$76*I35', NUM3)
    nas.append(na_row('Export shares to North America and Africa; import shares from Vietnam and Indonesia',
                      'B33:D34 and G33:I34', NA['trade_region'], unit='% and Mt'))
    rows.append(_sup(
        'Trade shares by country and region', '=D30', '% of exports',
        'Shares imported from the Master Industry Database (RED); volumes computed as total trade multiplied '
        'by share; Others is a residual so the shares always sum to 100%',
        '=$C$78*D30 and =1-SUM(D30:D34)',
        'Ministry of Steel Monthly Economic Report; Master Industry Database source S01',
        'DGCIS trade statistics',
        'FY2026 SHARES ARE HELD CONSTANT to FY2033E. No country-level trade forecast is attempted.',
        'Holding shares flat is disclosed rather than dressed up as a forecast, and it makes the one '
        'genuinely important number visible: EUROPE IS 34.4% OF INDIAN STEEL EXPORTS. That single figure is '
        'what makes the EU CBAM calculation on 20 ESG Model material, and it is why the EU share is also '
        'carried as a live input there rather than re-keyed.',
        'Shares must sum to exactly 100% by construction; the EU share of 34.4% must equal the value used '
        'on 20 ESG Model research block row 106',
        'High for the shares, Low for holding them flat',
        'Steel Price Forecast, ESG Model', 'Monthly for shares', 'FY2026 (provisional)',
        'Was entirely blank.', PCT1))

    # ---- trade drivers rows 40-47
    for i, nc in enumerate(NEW7):
        lc, ma = FLEG[i], MA7[i]
        w.link(f'{nc}40', f"='Steel Demand Model'!${lc}$116", MT)
        w.link(f'{nc}41', f"='Steel Supply Model'!${lc}$101", MT)
        w.link(f'{nc}43', f"='Steel Price Forecast'!${lc}$103", RS)
        w.link(f'{nc}44', f"='Model Assumptions'!{ma}$10", NUM1)
        w.link(f'{nc}46', f'=${lc}$88', PCT1)
        w.na(f'{nc}42', 'No global steel price forecast is made. The model carries point observations for '
                        'China, South East Asia and CIS at 22-Jun-2026 but forecasting regional prices is '
                        'outside its scope - see 09 Steel Price Forecast rows 59 to 63.')
        w.na(f'{nc}45', 'Freight is inside conversion cost, which is indexed to RBI CPI as a single block, '
                        'so no separate freight path exists in the model.')
        w.na(f'{nc}47', 'Anti-dumping duties are product and origin specific and no consolidated forward '
                        'schedule exists. The SAFEGUARD duty, which does have a notified schedule, is on '
                        'row 46 and in the research block row 88.')
    nas.append(na_row('Global steel prices, freight costs and anti-dumping duties',
                      'B42:H42, B45:H45, B47:H47',
                      'No global price or freight path is modelled; anti-dumping duties have no consolidated '
                      'forward schedule. The notified safeguard duty IS modelled on row 46.'))
    rows.append(_sup(
        'Safeguard duty schedule', '=B46', '%',
        'Cross-sheet reference to the notified policy schedule in the research block',
        '=$D$88',
        'Government of India safeguard duty notification; Master Industry Database policy record P03',
        '19 Trade Model research block row 88',
        'A SCHEDULE, NOT A FORECAST. 12% to 20-Apr-2026, 11.5% from 21-Apr-2026, 11% thereafter, expiring '
        '20-Apr-2028. Zero from FY2029E.',
        'THE SINGLE LARGEST IDENTIFIABLE POLICY RISK IN THE MODEL, and the one place where a policy cliff '
        'is explicitly carried. The duty is currently defending a domestic premium of roughly Rs 9,000/t '
        'over Chinese FOB, about 20%. The model does NOT reduce realisation when it expires in FY2029E, '
        'because whether the premium survives depends on whether the duty is renewed - which is a '
        'judgement, not a forecast. That is disclosed, and check A23 on 25 Audit Checks flags the year.',
        'Must match the notified schedule exactly; the FY2029E step to zero is deliberate and is the year '
        'to stress-test',
        'High for the schedule, n/a for its renewal', 'This sheet',
        'On any Government notification', '21-Apr-2026', 'Was blank.', PCT1))

    # ---- price competitiveness rows 52-56
    w.db('B53', 48580, RS, ref="='[Master Industry Database.xlsx]Steel Prices'!$I$27")
    for i, nc in enumerate(NEW8):
        w.link(f'{nc}52', f"='Steel Price Forecast'!${nc}$9", RS)
        if i:
            w.f(f'{nc}53', '=$B$53', RS)
        w.f(f'{nc}54', f'={nc}52-{nc}53', RS)
        w.f(f'{nc}55', f'=IF({nc}54<0,"Competitive - India below China",'
                       f'IF({nc}54<5000,"Broadly at parity","Uncompetitive on price"))', TEXT)
        w.f(f'{nc}56', f'=IF({nc}54>7500,"High - imports attractive",'
                       f'IF({nc}54>3000,"Moderate","Low"))', TEXT)
    rows.append(_sup(
        'India-China HRC price competitiveness', '=B54', 'Rs/t',
        'Computed in-cell from the India HRC series on 09 Steel Price Forecast and a held-flat China print',
        '=B52-B53',
        'IDBI Capital regional price snapshot, week ended 22-Jun-2026: India Rs 58,200/t, China Rs 48,580/t',
        'Master Industry Database KPI K19 records the spread at Rs 9,620/t; import-parity check in Apr-2026',
        'THE CHINA PRICE IS HELD FLAT across the horizon. No Chinese price forecast is attempted.',
        'Holding China flat while India moves with the model means the spread widens mechanically, and that '
        'is the honest way to present the risk rather than hide it: if Indian realisation rises to Rs '
        '69,000/t by FY2033E while China stays where it is, the import-parity gap becomes indefensible '
        'without a duty. Read this table as a stress indicator on the realisation driver, not as a forecast '
        'of Chinese prices.',
        'The FY2026A spread must reproduce the Rs 9,620/t recorded in the database; the classification rows '
        'must flip to "Uncompetitive" as the spread widens',
        'Medium for FY2026A, Low for the held-flat forecast', 'Steel Price Forecast', 'Weekly',
        '22-Jun-2026', 'Was entirely blank.', RS))

    # ---- scenario analysis rows 61-63
    for col, sc in zip('BCDE', SCEN):
        w.link(f'{col}61', eng.ref(sc, 'imp', 7), MT)
        w.link(f'{col}62', eng.ref(sc, 'exp', 7), MT)
        w.f(f'{col}63', f'={col}62-{col}61', MTS)
    rows.append(_sup(
        'Four-scenario trade position, FY2033E', '=B63', 'Mt net',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Driver D07 net exports and D21 imports, by scenario, on 02 Model Assumptions',
        'Observed history: net +8.8 Mt (FY2022) to net -4.7 Mt (FY2025)',
        'Imports are the driver plus any shortfall routed from the capacity ceiling; exports are imports '
        'plus net exports.',
        'The FY2033E range runs from a net export of about +6.5 Mt in the Bear Case to a net IMPORT of about '
        '-5.0 Mt in the Stress Case. The Bear Case exports MORE because weak domestic demand pushes tonnes '
        'offshore - which is exactly what happened in FY2021 and FY2022. Every scenario value sits inside '
        'the observed FY2015-FY2026 range, which is the test that matters.',
        'FY2033E net trade must equal exports less imports; check A25 on 25 Audit Checks caps import '
        'penetration at 25%',
        'Medium', 'Scenario Manager, Model Assumptions', 'Live', 'Live', 'Was entirely blank.', MTS))
    return rows, nas


# =========================================================================== 20 ESG Model
def esg(wb, audit, eng):
    ws = wb['ESG Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []
    NO_FY26_CBAM = ('The CBAM calculation runs from FY2027E, which is when the EU definitive regime begins '
                    'charging. FY2026A carries no carbon cost because none was payable.')

    # ---- executive summary rows 9-14
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        w.link(f'{nc}10', f'=${lc}$102', NUM2)
        if i:
            w.link(f'{nc}11', f'=${lc}$112', CR)
        else:
            w.na(f'{nc}11', NO_FY26_CBAM)
        w.na(f'{nc}9', 'The overall ESG score cannot be struck because the social and governance pillars '
                       'have no sourced inputs - see rows 51 to 64. The ENVIRONMENTAL pillar IS computed, '
                       'on row 69, and scores zero: Indian average emission intensity of 2.55 tCO2e/tfs is '
                       'above the 2.2 threshold at which steel ceases to qualify as green under the '
                       'taxonomy notified on 23-Dec-2024. Cell C72 is a live formula and resolves as soon '
                       'as the social and governance scores are entered.')
        w.na(f'{nc}12', NA['esg_capex'])
        w.na(f'{nc}13', NA['esg'])
        w.na(f'{nc}14', NA['esg'])
    nas.append(na_row('ESG score, ESG capex, renewable energy share and rating trend', 'B9:I9, B12:I14',
                      'The social and governance pillars have no sourced inputs and no producer publishes a '
                      'costed decarbonisation plan to FY2033.'))

    # ---- environmental KPIs rows 19-26
    for nc, lc in zip(NEW8, LEG):
        w.f(f'{nc}19', f"={lc}102*'Steel Supply Model'!{lc}101", NUM1)
        w.link(f'{nc}20', f'=${lc}$102', NUM2)
        for r in range(21, 27):
            w.na(f'{nc}{r}', NA['esg'])
    nas.append(na_row('Renewable share, energy, water, water recycling, waste recycling and slag '
                      'utilisation', 'B21:I26', NA['esg']))
    rows.append(_sup(
        'Industry CO2 emissions', '=I19', 'Mt CO2e',
        'Computed in-cell: emission intensity multiplied by finished steel production',
        "=C102*'Steel Supply Model'!C101",
        'Emission intensity is a USER INPUT of 2.55 tCO2e/tfs, explicitly NOT SOURCED (research block row '
        '102); production is from the model',
        'Green steel taxonomy thresholds notified 23-Dec-2024 (research block rows 96 to 99)',
        'Indian average intensity of 2.55 tCO2e/tfs held flat across the horizon - no decarbonisation '
        'trajectory is assumed.',
        'Holding intensity flat is a deliberately conservative and transparent choice. Indian producers '
        'have announced net-zero ambitions but none has published a dated, costed intensity pathway, so '
        'assuming improvement would flatter both the CBAM cost and the green-steel classification. On these '
        'numbers industry emissions rise from about 410 Mt CO2e to roughly 685 Mt by FY2033E purely because '
        'volume grows - which is the honest picture and the reason this is a first-order transition risk.',
        'Intensity of 2.55 sits ABOVE the 2.2 tCO2e/tfs threshold for 3-star green steel, so on these '
        'inputs no Indian average production qualifies as green',
        'Low - the intensity input is not sourced', 'Steel Supply Model', 'Annually, on BRSR filings',
        'n/a - user input', 'Rows 19 and 20 were blank. EDIT ROW 102 BEFORE USE.', NUM1))

    # ---- carbon cost forecast rows 31-35
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        w.f(f'{nc}31', f'={lc}104*{lc}105', NUM0)
        w.f(f'{nc}32', f'={lc}110*{lc}111', NUM2)
        w.link(f'{nc}33', f'=${lc}$112', CR)
        w.inp(f'{nc}34', 0, CR)
        w.f(f'{nc}35', f'={nc}33-{nc}34', CR)
    rows.append(_sup(
        'CBAM carbon cost', '=H35', 'Rs cr',
        'Computed in-cell: EU-destined exports multiplied by the emission intensity excess over the EU '
        'benchmark, multiplied by the EU carbon price converted to rupees',
        '=D110*D111*D104*D105/10 (research block row 112)',
        'EU share of Indian exports 34.4% IS sourced (Master Industry Database). The carbon price, both '
        'emission intensities and EUR/INR are USER INPUTS and NOT SOURCED',
        '19 Trade Model export destination shares; green steel taxonomy notified 23-Dec-2024',
        'CARBON CREDITS ARE SET TO ZERO. India has no operating compliance carbon market that would '
        'generate CBAM-eligible credits, and the Carbon Credit Trading Scheme is not yet in force.',
        'This is the only transition cost the model quantifies, and it is quantified precisely because the '
        'exposure is concentrated and measurable: 34.4% of exports go to the EU, and the intensity gap of '
        '0.70 tCO2e/tfs against the EU benchmark is chargeable. On the user inputs as delivered the cost is '
        'roughly Rs 1,400-2,100 crore a year - about 0.5% of industry EBITDA, material but not existential. '
        'FOUR OF THE FIVE INPUTS ARE UNSOURCED, so treat the magnitude as an order of estimate.',
        'Row 31 multiplied by row 32 divided by ten must equal row 33; row 33 must equal research block row '
        '112',
        'Low', 'Trade Model, Steel Supply Model', 'Quarterly, and on any EU ETS price move', 'n/a',
        'Rows 31 to 35 were blank. Carbon credits are blue because zero is a live user input.', CR))

    # ---- ESG capex rows 40-46, social rows 51-55, governance rows 60-64
    for i, nc in enumerate(NEW7):
        for r in range(40, 47):
            w.na(f'{nc}{r}', NA['esg_capex'])
    for nc in NEW8:
        for r in range(51, 56):
            w.na(f'{nc}{r}', NA['esg'])
        for r in range(60, 65):
            w.na(f'{nc}{r}', NA['esg'])
    nas.append(na_row('ESG capex by category', 'B40:H46', NA['esg_capex'], unit='Rs cr'))
    nas.append(na_row('Social metrics - employees, LTIFR, training, diversity, community investment',
                      'B51:I55', NA['esg']))
    nas.append(na_row('Governance metrics - independent directors, board diversity, ESG-linked pay, '
                      'compliance', 'B60:I64', NA['esg']))

    # ---- ESG scorecard rows 69-72
    w.f('C69', '=MEDIAN(0,100,($C$99-$C$102)/($C$99-$C$96)*100)', SCORE)
    for r in (70, 71):
        w.na(f'C{r}', NA['esg'])
    w.f('C72', '=IF(COUNT(C69:C71)=3,SUMPRODUCT($B$69:$B$71,C69:C71),"")', SCORE)
    nas.append(na_row('Social and governance pillar scores', 'C70:C71', NA['esg'], unit='score 0-100'))
    rows.append(_sup(
        'Environmental pillar score', '=C69', 'score 0-100',
        'Computed in-cell: the position of Indian average emission intensity between the "not green" '
        'threshold and the 5-star threshold of the notified taxonomy, clamped to 0-100',
        '=MEDIAN(0,100,($C$99-$C$102)/($C$99-$C$96)*100)',
        'Green steel taxonomy thresholds notified 23-Dec-2024 - these ARE sourced (Master Industry Database '
        'policy record P15)',
        'Research block rows 96 to 99 and row 102',
        'Scored against the Government of India taxonomy rather than against an external ESG rating '
        'methodology.',
        'The score is ZERO, and that is the correct and important answer: at 2.55 tCO2e/tfs the Indian '
        'average is above the 2.2 threshold at which steel ceases to qualify as green at all. Using the '
        "Government's own notified taxonomy as the yardstick is far more defendable than an agency score, "
        "because the thresholds are law. For reference, Tata Steel's 0.75 Mtpa Ludhiana scrap EAF is the "
        'kind of asset that clears the 5-star threshold of 1.6.',
        'The overall score on row 72 stays blank until the social and governance pillars are entered, which '
        'is deliberate - a partial weighted score would be misleading',
        'Medium for the method, Low for the intensity input',
        'This sheet', 'On any taxonomy amendment', '23-Dec-2024', 'Was blank.', SCORE))

    # ---- ESG financial impact rows 77-81
    for i, nc in enumerate(NEW7):
        lc, ma = FLEG[i], MA7[i]
        w.na(f'{nc}77', NA['esg_capex'])
        w.link(f'{nc}78', f'=${lc}$112', CR)
        w.f(f'{nc}79', f'=-{lc}112', CR)
        w.f(f'{nc}80', f'=-{lc}112', CR)
        w.f(f'{nc}81', f"=-{lc}112*'Model Assumptions'!{ma}$26", CR)
    rows.append(_sup(
        'ESG financial impact', '=H81', 'Rs cr',
        'Computed in-cell: the carbon cost flows one-for-one to EBITDA and free cash flow, and is '
        'capitalised at the scenario exit multiple to give the valuation impact',
        "=-D112*'Model Assumptions'!E$26",
        'CBAM calculation in the research block; exit multiple driver D19',
        '23 Comparable Valuation for the enterprise value the impact is measured against',
        'The carbon cost is treated as a permanent operating cost, so capitalising it at the exit multiple '
        'is the correct valuation treatment. ESG CAPEX IS NOT INCLUDED because it is not known.',
        'THE VALUATION IMPACT IS THEREFORE UNDERSTATED, and knowingly so. It captures the carbon cost but '
        'not the capital required to avoid it, which is the larger number: a scrap-EAF or hydrogen-ready '
        'route change costs materially more per tonne than Rs 55,000. The direction of the omission is '
        'disclosed here rather than buried.',
        'EBITDA impact must equal the negative of the CBAM cost; valuation impact must equal that '
        'multiplied by the exit multiple in force for the selected scenario',
        'Low', 'Model Assumptions, Comparable Valuation', 'Quarterly', 'n/a',
        'Rows 78 to 81 were blank; ESG capex remains blank by necessity.', CR))

    # ---- scenario analysis rows 86-89
    for r, sc in zip(range(86, 90), SCEN):
        w.f(f'B{r}', '=$D$104*$D$105', NUM0)
        w.na(f'C{r}', NA['esg_capex'])
        w.link(f'D{r}', f'=-{eng.cell(sc, "exp", 7)}*$J$106*$J$111*$J$104*$J$105/10', CR)
    nas.append(na_row('ESG capex by scenario', 'C86:C89', NA['esg_capex'], unit='Rs cr'))
    rows.append(_sup(
        'Four-scenario CBAM impact, FY2033E', '=D86', 'Rs cr',
        'Computed in-cell on the scenario engine export volume, at the same unsourced carbon price and '
        'intensity inputs',
        "=-'Scenario Manager'!$J$<exports>*$J$106*$J$111*$J$104*$J$105/10",
        'Export volumes from the scenario engine; all price and intensity inputs are USER INPUTS',
        '19 Trade Model scenario trade position',
        'The carbon PRICE is identical in all four scenarios, because it is an unsourced user input and the '
        'model does not flex unsourced parameters to manufacture a scenario.',
        'The CBAM cost is therefore scenario-sensitive only through EXPORT VOLUME, which is the honest '
        'result: the Bear Case pushes more tonnes offshore and so carries the highest carbon cost, at '
        'roughly Rs 2,600 crore against Rs 1,400 crore in the base case. A user who wants to stress the '
        'carbon price should edit research block row 104, and every figure on this sheet and its valuation '
        'impact will move together.',
        'Export volumes must match 19 Trade Model row 62; the base-case column must reproduce the live '
        'FY2033E CBAM cost',
        'Low', 'Scenario Manager, Trade Model', 'Live', 'n/a',
        'Was blank. Identical carbon prices across scenarios are deliberate, not a copy error.', CR))
    return rows, nas
