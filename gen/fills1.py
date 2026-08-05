"""Fills for 01 Control Panel, 02b Model Calibration, 00 Database Import,
00H Forecast Horizon, 03 Macroeconomic Model and 04 Steel Demand Model."""
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Font, Alignment

from .style import (SheetWriter, PCT0, PCT1, PCT2, NUM0, NUM1, NUM2, MT, MTS, RS, CR,
                    X1, X2, SCORE, USD, TEXT, PPT)
from .support import na_row
from .spec import NA

LEG = 'CDEFGHIJ'      # legacy block: FY2026A .. FY2033E
NEW8 = 'BCDEFGHI'     # new table:    FY2026A .. FY2033E
NEW7 = 'BCDEFGH'      # new table:    FY2027E .. FY2033E
SCEN = ['Base Case', 'Bull Case', 'Bear Case', 'Stress Case']


def _pair(new_cols=NEW8, leg_cols=LEG):
    return list(zip(new_cols, leg_cols))


# ======================================================================= 01 Control Panel
NAV_TARGETS = {
    18: 'Database Import', 19: 'Model Assumptions', 20: 'Model Calibration',
    21: 'Macroeconomic Model', 22: 'Steel Demand Model', 23: 'Steel Supply Model',
    24: 'Capacity Forecast', 25: 'Steel Price Forecast', 26: 'Revenue Forecast',
    27: 'EBITDA Model', 28: 'Comparable Valuation', 29: 'Industry Dashboard',
    30: 'Audit Checks', 31: 'Sources',
}


def control_panel(wb, audit):
    ws = wb['Control Panel']
    w = SheetWriter(ws, audit)

    # workbook status: was six hard-coded ticks, now six live tests
    checks = [
        ('H6', '=IF(COUNT(\'Model Calibration\'!$D$10:$D$14)=5,"OK","CHECK")',
         'Database connected', 'Counts the five FY2026A industry calibration anchors imported from the '
         'Master Industry Database.'),
        ('H7', '=IF(COUNT(\'Model Assumptions\'!$E$8:$K$33)=182,"OK","CHECK")',
         'Assumptions complete', 'All 26 drivers must resolve to a number in all seven forecast years: '
         '26 x 7 = 182. Identical to check A28 on 25 Audit Checks.'),
        ('H8', '=IF(AND(ISNUMBER(\'Model Calibration\'!$D$19),ISNUMBER(\'Model Calibration\'!$D$21),'
               'ABS(\'Model Calibration\'!$D$19-\'Model Calibration\'!$D$21-\'Model Calibration\'!$D$20)<1),'
               '"OK","CHECK")',
         'Calibration complete', 'Realisation less cash cost must equal EBITDA per tonne at the FY2026A '
         'calibration anchors: 59,974 - 49,241 = 10,733.'),
        ('H9', '=IF(COUNT(\'EBITDA Model\'!$D$103:$J$103)=7,"OK","CHECK")',
         'Forecast built', 'All seven forecast years of industry EBITDA must be numeric.'),
        ('H10', '=IFERROR(IF(\'Audit Checks\'!$C$44="PASS","OK",\'Audit Checks\'!$C$44),"CHECK")',
         'Audit passed', 'Mirrors the OVERALL result of the thirty validations on 25 Audit Checks. WARN is '
         'expected in the Bear and Stress cases.'),
        ('H11', '=IF(COUNT(\'Industry Dashboard\'!$D$32:$J$32)=7,"OK","CHECK")',
         'Dashboard updated', 'The dashboard demand line must be live across all seven forecast years.'),
    ]
    rows = []
    for addr, formula, label, why in checks:
        w.f(addr, formula, TEXT)
        rows.append(dict(item=f'Workbook status - {label}', value=f'=IF({addr}="OK",1,0)', unit='flag',
                         method='Conditional test, live', formula=formula,
                         primary='Computed from the model itself', secondary='25 Audit Checks',
                         assumption='None. This is a test, not an assumption.', reasoning=why,
                         cross='Re-evaluates on every recalculation and on every scenario change',
                         conf='High', linked='Model Assumptions, Model Calibration, EBITDA Model, '
                                            'Audit Checks, Industry Dashboard',
                         freq='Live', last='Live', numfmt='0',
                         comments='Replaced a hard-coded tick, which could never fail.'))

    # quick navigation
    for r, target in NAV_TARGETS.items():
        c = ws[f'B{r}']
        c.value = 'Go \u2192'
        c.hyperlink = Hyperlink(ref=f'B{r}', location=f"'{target}'!A1")
        c.font = Font(name=c.font.name, sz=c.font.sz, b=c.font.b, u='single', color='FF0563C1')
    rows.append(dict(item='Quick navigation links', value=None, unit='n/a',
                     method='Internal workbook hyperlink', formula="location = 'Sheet Name'!A1",
                     primary='n/a', secondary='n/a',
                     assumption='None.', reasoning='The fourteen navigation cells were labelled but not '
                     'linked. Each now jumps to cell A1 of its target sheet.',
                     cross='Targets verified against the workbook sheet list', conf='High',
                     linked='All model sheets', freq='On any sheet rename', last='n/a',
                     comments='Hyperlinks only - no formatting was changed apart from the underline that '
                              'signals a link.'))
    return rows, []


# ================================================================== 02b Model Calibration
def model_calibration(wb, audit):
    ws = wb['Model Calibration']
    w = SheetWriter(ws, audit)
    rows = []

    w.link('D23', "='Revenue Forecast'!$C$96/'Revenue Forecast'!$C$94*10", RS)
    rows.append(dict(
        item='MC105 Revenue per tonne', value='=D23', unit='Rs/t',
        method='Computed in-cell from the two FY2026A anchors',
        formula="='Revenue Forecast'!$C$96/'Revenue Forecast'!$C$94*10",
        primary='Derived from the Master Industry Database via 12 Revenue Forecast',
        secondary='Cross-checks to MC101 average realisation of Rs 59,974/t',
        assumption='Revenue per tonne of finished steel PRODUCED, not sold.',
        reasoning='Rs 9,65,222 cr of industry revenue on 160.94 Mt of finished production. The unit bridge '
                  'is Rs cr x 10 / Mt = Rs/t.',
        cross='Must equal MC101 realisation, because industry revenue is built as production x realisation',
        conf='Medium', linked='Revenue Forecast', freq='Quarterly', last='FY2026A', numfmt=RS,
        comments='Was blank.'))

    # MC401 pointed at consumption rather than finished production
    w.f('C43', '=D10/D11', '0.0000"x"')
    rows.append(dict(
        item='MC401 Crude-to-finished ratio', value='=C43', unit='x',
        method='Computed in-cell', formula='=D10/D11',
        primary='Joint Plant Committee via the Master Industry Database',
        secondary='Driver D06 on 02 Model Assumptions; 05 Steel Supply Model row 93',
        assumption='Yield loss between crude and finished steel.',
        reasoning='DEFECT CORRECTED. The cell read =D10/D11 only after this build; as delivered it read '
                  '=D10/D12, which divided crude steel production by CONSUMPTION (163.74 Mt) instead of '
                  'finished steel PRODUCTION (160.94 Mt) and returned 1.0286x instead of 1.0465x.',
        cross='Must equal driver D06 and 05 Steel Supply Model row 93, both 1.04647x',
        conf='High', linked='This sheet', freq='Annually', last='FY2026A', numfmt='0.0000',
        comments='Correction to a new-table formula; the research block was not touched.'))

    w.f('C47', '=D36/D12', PCT1)
    rows.append(dict(
        item='MC405 Import dependency', value='=C47', unit='%',
        method='Computed in-cell', formula='=D36/D12',
        primary='Joint Plant Committee via the Master Industry Database',
        secondary='19 Trade Model row 83 import penetration',
        assumption='Imports measured against apparent consumption, the standard penetration definition.',
        reasoning='6.524 Mt of imports on 163.74 Mt of consumption is 4.0%, down from 6.3% in FY2025 as the '
                  'safeguard duty bit.',
        cross='Must equal 19 Trade Model row 83 FY2026A', conf='High', linked='Trade Model',
        freq='Monthly, on each Ministry of Steel report', last='FY2026A', numfmt=PCT1,
        comments='Was blank.'))

    # the five calibration checks were hard-coded ticks; make them arithmetic
    cchecks = [
        ('D52', '=IF(ABS($D$13*$D$14-$D$10)<0.05,"PASS","FAIL")', 'MC501 Production = Capacity x Utilisation',
         'Capacity 220.4 Mtpa multiplied by utilisation 76.42% must return crude production of 168.42 Mt.'),
        ('D53', '=IF(ABS($D$19*$D$11/10-\'Revenue Forecast\'!$C$96)<1,"PASS","FAIL")',
         'MC502 Revenue = Price x Volume',
         'Realisation Rs 59,974/t on 160.94 Mt of finished production must return Rs 9,65,222 cr. Proves '
         'the Rs/t to Rs crore bridge (divide by 10).'),
        ('D54', '=IF(ABS(($D$19-$D$21)-$D$20)<1,"PASS","FAIL")', 'MC503 EBITDA = Revenue less Cost',
         'Realisation less cash cost must equal EBITDA per tonne: 59,974 - 49,241 = 10,733.'),
        ('D55', '=IF(ABS($D$11-($D$12+($D$37-$D$36)+$D$38))<0.02,"PASS","FAIL")',
         'MC504 Consumption identity balances',
         'The Joint Plant Committee identity: finished production 160.94 = consumption 163.74 + net exports '
         '0.078 + stock variation -2.88.'),
        ('D56', '=IF(ABS(($D$37-$D$36)-\'Steel Supply Model\'!$C$90)<0.001,"PASS","FAIL")',
         'MC505 Import / export balance',
         'Exports 6.602 less imports 6.524 must equal the net export figure the supply model uses, +0.078 Mt.'),
    ]
    for addr, formula, label, why in cchecks:
        w.f(addr, formula, TEXT)
        rows.append(dict(item=label, value=f'=IF({addr}="PASS",1,0)', unit='flag',
                         method='Arithmetic reconciliation, live', formula=formula,
                         primary='Computed from the calibration anchors above',
                         secondary='25 Audit Checks A08, A11, A12, A13',
                         assumption='None. This is a reconciliation.', reasoning=why,
                         cross='Duplicated independently on 25 Audit Checks', conf='High',
                         linked='Revenue Forecast, Steel Supply Model', freq='Live', last='FY2026A',
                         numfmt='0', comments='Replaced a hard-coded tick.'))
    w.txt('C55', 'Finished production = Consumption + Net exports + Variation in stock')
    w.txt('C56', 'Exports less Imports = Net trade position')
    return rows, []


# =================================================================== 00 Database Import
def database_import(wb, audit, mdb_counts):
    ws = wb['Database Import']
    w = SheetWriter(ws, audit)
    rows = []
    for r in range(20, 33):
        sheet = ws[f'B{r}'].value
        if not sheet:
            continue
        n = mdb_counts.get(sheet)
        cur = ws[f'D{r}'].value
        if n is None:
            w.na(f'D{r}', f'Sheet {sheet!r} was not found in Master Industry Database.xlsx.')
        elif cur in (None, 'xxx', 'XXX'):
            w.db(f'D{r}', n, NUM0, ref=f"COUNTA('[Master Industry Database.xlsx]{sheet}')")
    rows.append(dict(
        item='Imported dataset record counts', value='=SUM(D20:D32)', unit='rows',
        method='Counted from Master Industry Database.xlsx at build time',
        formula="='[Master Industry Database.xlsx]<sheet>'! record count",
        primary='Master Industry Database.xlsx, 03-Aug-2026 vintage',
        secondary='Master Industry Database Contents sheet',
        assumption='A record is a row carrying an entity label and at least one numeric observation.',
        reasoning='The delivered workbook carried the placeholder "xxx" in twelve of the thirteen rows. '
                  'Counts are now the actual record counts of each source sheet.',
        cross='Company List row must equal 15, which is the stated coverage universe on the cover sheet',
        conf='High', linked='This sheet', freq='On each database refresh', last='03-Aug-2026', numfmt=NUM0,
        comments='RED fill marks every cell that should become an external link when the two workbooks are '
                 'reconnected.'))
    return rows, []


# ================================================================== 00H Forecast Horizon
def forecast_horizon(wb, audit):
    ws = wb['Forecast Horizon']
    w = SheetWriter(ws, audit)
    w.f('C20', '=IF(F20=MAX($F$20:$H$20),"SELECTED","Rejected")', TEXT)
    w.f('D20', '=IF(G20=MAX($F$20:$H$20),"SELECTED","Rejected")', TEXT)
    w.f('E20', '=IF(H20=MAX($F$20:$H$20),"SELECTED","Rejected")', TEXT)
    rows = [dict(
        item='Horizon decision', value='=G20', unit='weighted score',
        method='Live SUMPRODUCT of weights and criterion scores; verdict resolved by MAX',
        formula='=IF(G20=MAX($F$20:$H$20),"SELECTED","Rejected")',
        primary='Weighted scoring of five criteria, each with its evidence stated in column I',
        secondary='Master Industry Database project tracker (pipeline terminates FY2031)',
        assumption='The five criteria and their weights are the modeller\'s framework; the SCORES are '
                   'evidenced in column I.',
        reasoning='Seven years wins on four of the five criteria. The dated pipeline ends FY2031, so a '
                  'seven-year horizon captures 100% of commissioning plus a two-year ramp to a normalised '
                  'terminal year, which is what a DCF terminal value requires.',
        cross='A five-year horizon scores 2.4 and a ten-year 2.75 against 4.8 for seven years',
        conf='High', linked='This sheet', freq='Annually', last='03-Aug-2026', numfmt='0.00',
        comments='The verdict row was blank; it is now resolved by formula rather than asserted.')]
    return rows, []


# ================================================================ 03 Macroeconomic Model
def macro(wb, audit):
    ws = wb['Macroeconomic Model']
    w = SheetWriter(ws, audit)
    rows, na_groups = [], []

    # ---- executive summary rows 13-23 (B = FY2026A, C..I = FY2027E..FY2033E)
    # row 13 (GDP) already ships wired to 02 Model Assumptions row 8.
    w.link('B19', "='Model Assumptions'!$D$11", PCT1)
    for nc, ma in zip('CDEFGHI', 'EFGHIJK'):
        w.link(f'{nc}19', f"='Model Assumptions'!{ma}$11", PCT1)
    w.link('B22', "='Model Assumptions'!$D$10", NUM1)
    for nc, ma in zip('CDEFGHI', 'EFGHIJK'):
        w.link(f'{nc}22', f"='Model Assumptions'!{ma}$10", NUM1)

    w.inp('B21', 0.0525, PCT2)
    for nc in 'CDEFGHI':
        w.f(f'{nc}21', '=$B$21', PCT2)
    w.inp('B23', 79.0, USD)
    for nc in 'CDEFGHI':
        w.f(f'{nc}23', '=$B$23', USD)

    orange = {
        14: ('Industrial production growth (IIP)', NA['macro_sub']),
        15: ('Manufacturing GVA growth', NA['macro_sub']),
        16: ('Infrastructure growth', NA['macro_sub']),
        17: ('Construction growth', NA['macro_sub']),
        18: ('Auto production growth', NA['macro_sub']),
        20: ('WPI inflation', NA['macro_sub']),
    }
    for r, (label, why) in orange.items():
        w.narow(NEW8, r, why)
        na_groups.append(na_row(label, f'B{r}:I{r}', why, unit='%'))

    rows += [
        dict(item='Real GDP growth', value='=I13', unit='%',
             method='Cross-sheet reference to driver D01', formula="='Model Assumptions'!K8",
             primary='Reserve Bank of India, Monetary Policy Committee statement, 05-Jun-2026',
             secondary='IMF World Economic Outlook; Master Industry Database source S45',
             assumption='FY2027E adopts the RBI projection of 6.6% unchanged, then converges to a 6.5% '
                        'steady state by FY2032E.',
             reasoning='The RBI is the authoritative domestic forecaster and publishes an explicit FY2027 '
                       'number. Beyond FY2027 no institution publishes an India path, so the model '
                       'converges rather than extrapolates.',
             cross='FY2026A of 7.6% is the RBI estimate; Q4FY2026 actual was 7.7%',
             conf='High', linked='Model Assumptions', freq='On each RBI MPC statement (bi-monthly)',
             last='05-Jun-2026', numfmt=PCT1,
             comments='Single source of truth for every volume forecast in the workbook.'),
        dict(item='CPI inflation', value='=I19', unit='%',
             method='Cross-sheet reference to driver D04', formula="='Model Assumptions'!K$11",
             primary='Reserve Bank of India MPC, 05-Jun-2026',
             secondary='Master Industry Database; RBI target band 4% +/- 2%',
             assumption='FY2027E 5.1% as projected, converging to the 4.5% mid-point of the target band.',
             reasoning='The RBI raised its FY2027 projection by 50bps to 5.1% citing crude oil and supply '
                       'disruption. Convergence to the band mid-point is the RBI\'s own stated objective.',
             cross='Drives driver D12 conversion cost inflation and the discount rate',
             conf='High', linked='Model Assumptions', freq='On each RBI MPC statement', last='05-Jun-2026',
             numfmt=PCT1, comments='Was blank.'),
        dict(item='Policy repo rate', value='=B21', unit='%',
             method='Manual input, sourced observation held flat across the horizon',
             formula='=$B$21 (held flat)',
             primary='Reserve Bank of India MPC, 05-Jun-2026 - repo held at 5.25%',
             secondary='02 Model Assumptions section D records the same observation alongside the 10-year '
                       'G-sec of 6.833% at 31-Jul-2026',
             assumption='Held flat. No policy-rate path is forecast.',
             reasoning='The same convention the model already applies to the 10-year G-sec on row 108 of '
                       'the research block: a term-structure or policy-rate forecast is not attempted, so '
                       'the last observed print is carried forward and labelled as such.',
             cross='Consistent with the 6.833% 10-year G-sec used as the risk-free rate in the WACC build-up',
             conf='High for FY2026A, Low for the forecast years',
             linked='Model Assumptions', freq='On each RBI MPC statement (bi-monthly)', last='05-Jun-2026',
             numfmt=PCT2,
             comments='Not used downstream. Carried as a macro reference and for the interest-rate score below.'),
        dict(item='Brent crude', value='=B23', unit='US$/bbl',
             method='Manual input, sourced observation held flat across the horizon',
             formula='=$B$23 (held flat)',
             primary='IDBI Capital Commodity Price Update, week ended 22-Jun-2026 - Brent US$79/bbl',
             secondary='Master Industry Database source S21 input-cost deck',
             assumption='Held flat. No oil price forecast is attempted.',
             reasoning='Brent is not a driver of this model; it is carried because it is the RBI\'s stated '
                       'swing factor for both GDP and CPI. Holding a market print flat is disclosed rather '
                       'than dressed up as a forecast.',
             cross='The RBI cited energy prices as the reason for cutting FY2027 GDP to 6.6%',
             conf='High for the spot print, n/a for the forecast years', linked='This sheet',
             freq='Daily market data; refresh from Bloomberg or Refinitiv', last='22-Jun-2026', numfmt=USD,
             comments='Was blank.'),
        dict(item='USD/INR', value='=I22', unit='Rs/US$',
             method='Cross-sheet reference to driver D03', formula="='Model Assumptions'!K$10",
             primary='IDBI Capital, 22-Jun-2026 - spot 94, 52-week range 85 to 98',
             secondary="FY2026 average of 88.4 derived from Tata Steel's dual-currency disclosure",
             assumption='Gradual depreciation to 100.5 by FY2033E in the base case.',
             reasoning='The second largest single driver of FY2033E EBITDA. Iron ore and coking coal are '
                       'dollar-denominated while realisation is in rupees, so depreciation is '
                       'unambiguously margin-negative.',
             cross='22 Sensitivity Analysis row 98 quantifies a 10% adverse move at -23.7% of FY2033E EBITDA',
             conf='High', linked='Model Assumptions', freq='Quarterly, or on any material move',
             last='22-Jun-2026', numfmt=NUM1, comments='Was blank.'),
    ]

    # ---- economic scorecard rows 28-34
    score = [
        (28, '=MEDIAN(0,100,(\'Model Assumptions\'!$E$8-0.04)/0.05*100)',
         'Real GDP growth scored on a 4.0% to 9.0% band',
         'The band spans the FY2021 COVID contraction recovery floor and the strongest post-liberalisation '
         'prints. FY2027E of 6.6% scores 52.'),
        (33, '=MEDIAN(0,100,(0.09-\'Model Assumptions\'!$D$159)/0.04*100)',
         'Risk-free rate scored INVERSELY on a 5.0% to 9.0% band',
         'A lower yield is better for a capital-intensive industry. The 10-year G-sec of 6.833% scores 54.'),
        (34, '=MEDIAN(0,100,(0.08-\'Model Assumptions\'!$E$11)/0.06*100)',
         'CPI inflation scored INVERSELY on a 2.0% to 8.0% band',
         'The band is the RBI target band widened by one percentage point either side. FY2027E of 5.1% '
         'scores 48.'),
    ]
    for r, formula, method, why in score:
        w.f(f'C{r}', formula, SCORE)
        w.f(f'D{r}', f'=IF(C{r}>=60,"Strong",IF(C{r}>=40,"Neutral","Weak"))', TEXT)
        rows.append(dict(item=f'Economic scorecard - {ws[f"A{r}"].value}', value=f'=C{r}', unit='score 0-100',
                         method=method, formula=formula,
                         primary='Computed from the model drivers', secondary='02 Model Assumptions',
                         assumption='The scoring BANDS are the modeller\'s framework and are stated in the '
                                    'formula so they are auditable; the INPUTS are sourced.',
                         reasoning=why, cross='Weights in column B sum to 1.00 across rows 28 to 34',
                         conf='Medium', linked='Model Assumptions', freq='Live', last='Live', numfmt=SCORE,
                         comments='Was blank.'))
    for r in (29, 30, 31, 32):
        w.narow('CD', r, NA['macro_sub'])
    na_groups.append(na_row('Economic scorecard - infrastructure, manufacturing, construction and automobile',
                            'C29:D32', NA['macro_sub'], unit='score 0-100'))

    # ---- macro drivers of steel demand rows 39-46
    for r in range(39, 47):
        w.narow('CDEFGHI', r, NA['sector_split'])
    na_groups.append(na_row('Macro drivers of steel demand, by sector', 'C39:I46', NA['sector_split'],
                            unit='%'))

    # ---- leading indicators rows 51-57
    for r in range(51, 58):
        w.narow('BCD', r, NA['macro_sub'])
    na_groups.append(na_row('Leading indicators (PMI, core sector, infrastructure and government capex, '
                            'housing completions, vehicle production, railway capex)',
                            'B51:D57', NA['macro_sub']))

    # ---- exchange rate impact rows 62-66
    fx = [
        (62, '="Landed import cost moves 1:1 with the rupee. Cumulative depreciation FY2026A to FY2033E is "'
             '&TEXT(\'Model Assumptions\'!$K$10/$B$22-1,"0.0%")&", which is import-negative and therefore '
             'supportive of domestic realisation."'),
        (63, '="Rupee export realisation improves by the same "&TEXT(\'Model Assumptions\'!$K$10/$B$22-1,"0.0%")'
             '&", but India exports only "&TEXT(\'Trade Model\'!$C$84,"0.0%")&" of finished production, so the '
             'benefit is second-order."'),
        (64, '="Rs "&TEXT(\'Raw Material Forecast\'!$J$89*0.05,"#,##0")&"/dmt of additional cost per further 5% '
             'depreciation, at the FY2033E dollar iron ore price."'),
        (65, '="Rs "&TEXT(\'Raw Material Forecast\'!$J$90*0.05,"#,##0")&"/t of additional cost per further 5% '
             'depreciation, at the FY2033E dollar coking coal price."'),
        (66, '="Second largest driver of FY2033E EBITDA. A 10% adverse move is worth Rs "'
             '&TEXT(-\'Sensitivity Analysis\'!$E$98,"#,##0")&" cr, or "'
             '&TEXT(-\'Sensitivity Analysis\'!$F$98,"0.0%")&" - see 22 Sensitivity Analysis row 98."'),
    ]
    for r, formula in fx:
        w.f(f'B{r}', formula, TEXT)
        ws[f'B{r}'].alignment = Alignment(wrap_text=True, vertical='center')
    rows.append(dict(
        item='Exchange rate impact block', value="='Model Assumptions'!$K$10/$B$22-1", unit='%',
        method='Live text built with TEXT() around model cells - the numbers cannot go stale',
        formula='="Rs "&TEXT(\'Raw Material Forecast\'!$J$89*0.05,"#,##0")&"/dmt ..."',
        primary='Computed from the model', secondary='22 Sensitivity Analysis tornado row 98',
        assumption='A 5% incremental depreciation is used to size the cost sensitivity.',
        reasoning='India imports substantially all of its coking coal and prices iron ore off a dollar '
                  'benchmark, while selling in rupees. The exposure is asymmetric and negative.',
        cross='Consistent with the tornado ranking on 22 Sensitivity Analysis', conf='High',
        linked='Model Assumptions, Raw Material Forecast, Trade Model, Sensitivity Analysis',
        freq='Live', last='Live', numfmt=PCT1, comments='Was blank.'))

    # ---- commodity environment rows 71-74
    comm = [(71, '=$B$23', '=$I$23'),
            (72, "='Raw Material Forecast'!$C$85", "='Raw Material Forecast'!$J$85"),
            (73, "='Raw Material Forecast'!$C$86", "='Raw Material Forecast'!$J$86")]
    for r, cur, fc in comm:
        w.link(f'B{r}', cur, USD)
        w.link(f'C{r}', fc, USD)
        w.f(f'D{r}', f'=IF({fc.lstrip("=")}>{cur.lstrip("=")}*1.05,"Rising",'
                     f'IF({fc.lstrip("=")}<{cur.lstrip("=")}*0.95,"Falling","Broadly flat"))', TEXT)
    w.db('B74', 111.5, USD, ref="='[Master Industry Database.xlsx]Coking Coal Prices'!$I$28")
    w.narow('CD', 74, 'Thermal coal is carried as an energy-cost reference only and is deliberately NOT '
                      'forecast: the model indexes the non-raw-material element of cash cost to RBI CPI '
                      'rather than to a thermal coal path. The Master Industry Database warns explicitly '
                      'against using the thermal series as a coking coal proxy.')
    na_groups.append(na_row('Thermal coal forecast and trend', 'C74:D74',
                            'Thermal coal is an energy-cost reference only and is not forecast by this '
                            'model. Conversion cost is indexed to RBI CPI instead, which is sourced.',
                            unit='US$/t'))
    rows.append(dict(
        item='Commodity environment - iron ore and coking coal', value="='Raw Material Forecast'!$J$85",
        unit='US$', method='Cross-sheet reference to 10 Raw Material Forecast; trend resolved by formula',
        formula="=IF('Raw Material Forecast'!$J$85>'Raw Material Forecast'!$C$85*1.05,\"Rising\",...)",
        primary='World Bank Commodity Price Data (Pink Sheet) for iron ore; Ministry of Steel for coking coal',
        secondary='IDBI Capital weekly tracker; Master Industry Database sources S17, S01, S20, S21',
        assumption='Iron ore eases to US$90/dmt by FY2030E then firms; coking coal eases to US$190/t then firms.',
        reasoning='Iron ore has fallen for three consecutive fiscal years. The forecast reflects new Simandou '
                  'and Australian supply against decelerating Chinese output, then a modest firming as the '
                  'cost curve reasserts.',
        cross='FY2026 iron ore average of US$100.52/dmt is the mean of twelve World Bank monthly prints',
        conf='High for iron ore, Medium for coking coal', linked='Raw Material Forecast',
        freq='Monthly', last='Jun-2026', numfmt=USD, comments='Was blank.'))

    # ---- macro outlook narrative
    outlook = [
        (78, 'GROWTH. The RBI cut its FY2027 real GDP projection to 6.6% from 6.9% at the 05-Jun-2026 MPC, '
             'citing the West Asia conflict, energy prices and monsoon risk, against an estimated 7.6% for '
             'FY2026 and a Q4FY2026 actual of 7.7%. The base case adopts 6.6% unchanged and converges to a '
             '6.5% steady state; it does not extrapolate the FY2026 outturn, which was flattered by a low '
             'base and front-loaded government capex.'),
        (80, 'PRICES AND POLICY. CPI is projected at 5.1% for FY2027, 50bps above the previous projection, '
             'converging to the 4.5% mid-point of the RBI target band. The repo rate was held at 5.25% and '
             'the 10-year G-sec was 6.833% on 31-Jul-2026. Conversion cost inflation in the model is '
             'anchored on this CPI path rather than on a separately asserted wage or power series.'),
        (82, 'CURRENCY. USD/INR averaged 88.4 in FY2026 and was 94 spot on 22-Jun-2026 within an 85 to 98 '
             '52-week range. The base case takes the rupee to 100.5 by FY2033E. This is the second largest '
             'single driver of industry EBITDA in the model, because iron ore and coking coal are priced in '
             'dollars while realisation is in rupees.'),
        (84, 'TRANSMISSION TO STEEL. Macro reaches steel demand through exactly one channel in this model: '
             'real GDP growth multiplied by a demand elasticity of 1.15x tapering to 1.05x. That elasticity '
             'is the realised FY2015-FY2026 outturn, not an assertion. Everything downstream - supply, '
             'utilisation, price, cost, margin, cash flow and valuation - follows from that single link, '
             'which is why the GDP and elasticity rows are the two most important cells in the workbook.'),
        (86, 'WHAT IS NOT MODELLED. Industrial production, manufacturing GVA, sector-level infrastructure, '
             'construction and automotive growth, WPI and the leading-indicator block above are left blank '
             'and shaded orange. No institution publishes these to FY2033, and no FY2026 actual was '
             'retrievable in this research cycle. They are carried as a structure for the user to populate '
             'from MOSPI and RBI releases; nothing in the model depends on them.'),
    ]
    for r, text in outlook:
        c = ws.cell(r, 1, text)
        c.font = Font(name='Calibri', sz=9, color='FF404040')
        c.alignment = Alignment(wrap_text=False, vertical='top')
        ws.row_dimensions[r].height = 30
    return rows, na_groups


# ================================================================== 04 Steel Demand Model
def demand(wb, audit, eng):
    ws = wb['Steel Demand Model']
    w = SheetWriter(ws, audit)
    rows, na_groups = [], []

    # ---- executive summary rows 16-20
    for nc, lc in _pair():
        w.link(f'{nc}16', f'=${lc}$116', MT)
        w.link(f'{nc}17', f'=${lc}$115', PCT1)
        w.link(f'{nc}18', f'=${lc}$121', NUM1)
        w.f(f'{nc}19', f'=IFERROR({lc}116/({lc}116+\'Trade Model\'!{lc}78),"")', PCT1)
    w.link('B20', '=$K$116', PCT1)

    # ---- historical demand rows 25-30
    hist = [(25, 'FY2021', 94.89, -0.0527), (26, 'FY2022', 105.75, 0.1144),
            (27, 'FY2023', 119.89, 0.1337), (28, 'FY2024', 136.29, 0.1370),
            (29, 'FY2025', 152.13, 0.1162), (30, 'FY2026', 164.19, 0.0790)]
    for r, fy, mt, g in hist:
        w.db(f'B{r}', mt, MT, ref=f"='[Master Industry Database.xlsx]Demand & Consumption'!$C${r - 5 + 20}")
        w.db(f'C{r}', g, PCT1, ref=f"='[Master Industry Database.xlsx]Demand & Consumption'!$D${r - 5 + 20}")
        w.db(f'D{r}', 'JPC / Ministry of Steel', TEXT)
    rows.append(dict(
        item='Historical apparent consumption FY2021-FY2026', value='=B30', unit='Mt',
        method='Master Industry Database import (RED - to be reconnected as an external link)',
        formula="='[Master Industry Database.xlsx]Demand & Consumption'!$C$20:$C$31",
        primary='Joint Plant Committee, published by the Ministry of Steel',
        secondary='World Steel Association Short Range Outlook, April 2026',
        assumption='June-2026 JPC vintage (164.19 Mt for FY2026).',
        reasoning='Twelve unbroken fiscal years with a single contraction, FY2021, caused by the COVID '
                  'construction shutdown. Consumption compounded at 10.6% over FY2022-FY2026.',
        cross='The model forecasts off the April-2026 vintage of 163.74 Mt so the supply identity closes; '
              'the vintage difference of 0.45 Mt is carried as a memo on research block row 117',
        conf='High', linked='Master Industry Database', freq='Monthly, on each Ministry of Steel report',
        last='FY2026 (provisional)', numfmt=MT,
        comments='Six years of history were blank. RED fill marks the database-sourced cells.'))

    # ---- sector-wise demand rows 35-44, total row 45
    for r in range(35, 45):
        w.narow(NEW8, r, NA['sector_split'])
    for nc in 'CDEFGHI':
        w.f(f'{nc}45', f'=IF(COUNT({nc}35:{nc}44)=0,"",SUM({nc}35:{nc}44))', PCT0)
    na_groups.append(na_row('Sector-wise demand shares', 'B35:I44', NA['sector_split'], unit='% share'))

    # ---- demand drivers rows 50-56
    for nc, lc in zip('CDEFGHI', 'DEFGHIJ'):
        w.link(f'{nc}50', f"='Macroeconomic Model'!{lc}$101", PCT1)
        w.link(f'{nc}56', f"='Trade Model'!{lc}$78", MT)
    w.f('J50', '="Every volume forecast in the workbook resolves from this row through an elasticity of "'
                '&TEXT(\'Model Assumptions\'!$E$9,"0.00")&"x tapering to "'
                '&TEXT(\'Model Assumptions\'!$K$9,"0.00")&"x."', TEXT)
    w.f('J56', '="Exports are "&TEXT(\'Trade Model\'!$C$84,"0.0%")&" of finished production in FY2026A and "'
                '&TEXT(\'Trade Model\'!$J$84,"0.0%")&" in FY2033E. India is a domestic market; exports are '
                'the residual, not the driver."', TEXT)
    w.na('B50', 'A weighting scheme across demand drivers is not asserted. The model uses a single '
                'top-down channel - real GDP growth times elasticity - so a multi-driver weighting would '
                'be decorative and would imply a decomposition the model does not perform.')
    for r in range(51, 56):
        w.na(f'B{r}', 'A weighting scheme across demand drivers is not asserted - see the note on B50.')
        w.narow('CDEFGHI', r, NA['sector_split'])
        w.na(f'J{r}', NA['sector_split'])
    w.na('B56', 'A weighting scheme across demand drivers is not asserted - see the note on B50.')
    na_groups.append(na_row('Demand driver weights, and sector driver growth rates',
                            'B50:B56 and C51:I55', NA['sector_split']))

    # ---- per capita analysis rows 61-68
    for i, r in enumerate(range(61, 69)):
        lc = LEG[i]
        w.link(f'B{r}', f"=${lc}$119*1000", NUM1)
        w.link(f'C{r}', f'=${lc}$116', MT)
        w.f(f'D{r}', f'=IFERROR(C{r}/B{r}*1000,"")', NUM1)
    rows.append(dict(
        item='Per capita consumption build', value='=D68', unit='kg',
        method='Computed in-cell from the population and consumption rows of the research block',
        formula='=IFERROR(C68/B68*1000,"")',
        primary='Joint Plant Committee (115.7 kg for FY2026); population derived from the database',
        secondary='World Steel Association SSY-2025: world 215 kg, China 604 kg',
        assumption='Population growth of 0.85% per annum, which is NOT sourced and is flagged Low '
                   'confidence on research block row 120.',
        reasoning='Population is derived, not asserted: 163.74 Mt of consumption divided by the reported '
                  '115.7 kg per capita implies 1.4191 bn people. Reconciling in this direction keeps the '
                  'per capita row consistent with the JPC print.',
        cross='Must reproduce 115.7 kg in FY2026A; reaches roughly 179 kg by FY2033E, still below half the '
              'current world average',
        conf='Medium', linked='This sheet', freq='Annually', last='FY2026A', numfmt=NUM1,
        comments='Was blank. Population is shown in millions here and in billions on the research block.'))

    # ---- international comparison rows 73-78
    w.db('B73', 115.7, NUM1, ref="='[Master Industry Database.xlsx]Demand & Consumption'!$E$45")
    w.db('B74', 604.0, NUM1, ref="='[Master Industry Database.xlsx]Demand & Consumption'!$E$47")
    w.db('B78', 215.0, NUM1, ref="='[Master Industry Database.xlsx]Demand & Consumption'!$E$46")
    for r in (75, 76, 77):
        w.na(f'B{r}', NA['intl'])
    na_groups.append(na_row('Per capita consumption - Japan, South Korea and the USA', 'B75:B77',
                            NA['intl'], unit='kg'))
    rows.append(dict(
        item='International per capita comparison', value='=B73/B78', unit='x world average',
        method='Master Industry Database import (RED)',
        formula="='[Master Industry Database.xlsx]Demand & Consumption'!$E$45:$E$47",
        primary='World Steel Association, Steel Statistical Yearbook 2025 (CY2024 basis)',
        secondary='Joint Plant Committee for the India figure (FY2026 basis)',
        assumption='India FY2026 against China and world CY2024 - the periods differ and are labelled.',
        reasoning='India at 115.7 kg is 54% of the world average and 19% of China. This is the single '
                  'strongest structural argument for the demand elasticity above 1.0x that the model uses.',
        cross='Consistent with the 1.10x realised elasticity over FY2015-FY2026',
        conf='High', linked='Master Industry Database', freq='Annually, on the worldsteel yearbook',
        last='CY2024 / FY2026', numfmt=X2, comments='Japan, South Korea and USA were not retrievable.'))

    # ---- demand forecast rows 83-89 (column E already carries =B*C*D)
    for i, r in enumerate(range(83, 90)):
        lc = LEG[i + 1]
        prev = f'E{r - 1}' if r > 83 else '$C$116'
        w.f(f'B{r}', f'={prev}', MT)
        w.link(f'C{r}', f'=1+{lc}$115', '0.0000')
        w.inp(f'D{r}', 1.0, '0.0000')
    rows.append(dict(
        item='Demand forecast build', value='=E89', unit='Mt',
        method='Computed in-cell; base demand x macro adjustment x sector adjustment',
        formula='=B89*C89*D89, where C = 1 + implied consumption growth and D = 1.0000',
        primary='Derived from real GDP growth and the demand elasticity',
        secondary='Reconciles to research block row 116 to the last decimal',
        assumption='The SECTOR ADJUSTMENT is set to exactly 1.0000 in every year. That is a deliberate '
                   'modelling choice, not a placeholder.',
        reasoning='The model is a top-down GDP x elasticity build. No sector overlay is applied because '
                  'sector-level demand splits could not be sourced, and applying an unsourced overlay '
                  'would move the headline volume forecast without evidence. The column is live, so a user '
                  'who has a sector view can enter it and the whole model responds.',
        cross='E89 must equal research block J116 exactly; both give 268.74 Mt in the base case',
        conf='High for the build, n/a for the sector overlay', linked='This sheet', freq='Live',
        last='Live', numfmt=MT,
        comments='The multiplicative frame was empty except for column E. Column D is blue because it is '
                 'the one genuine user input in this block.'))

    # ---- scenario analysis rows 94-100
    for i, r in enumerate(range(94, 101)):
        for col, sc in zip('BCDE', SCEN):
            w.link(f'{col}{r}', eng.ref(sc, 'cons', i + 1), MT)
    rows.append(dict(
        item='Four-scenario consumption path', value='=B100', unit='Mt',
        method='Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        formula="='Scenario Manager'!$J$<engine row>",
        primary='Rebuilt live from the SCENARIO DRIVER MATRIX on 02 Model Assumptions rows 38 to 45',
        secondary='Independently computed tie-out values on 24 Scenario Manager rows 69 to 79',
        assumption='Consumption compounds at real GDP growth times elasticity in each scenario.',
        reasoning='Hard-coding a scenario table guarantees it drifts from the drivers. The engine rebuilds '
                  'the whole model for all four scenarios, so this table can never disagree with the '
                  'assumptions.',
        cross='FY2033E must equal 268.74 / 299.45 / 243.73 / 229.44 Mt, which are the independently '
              'computed values already on 24 Scenario Manager row 69',
        conf='High', linked='Scenario Manager, Model Assumptions', freq='Live', last='Live', numfmt=MT,
        comments='Demand is the least contested part of the model: even the stress case grows.'))
    return rows, na_groups
