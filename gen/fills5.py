"""Fills for 21 Scenario Manager, 22 Sensitivity Analysis, 23 Comparable Valuation,
24 Industry Dashboard and the workbook navigation."""
from openpyxl.styles import Font, Alignment

from .style import (SheetWriter, link_to, PCT0, PCT1, PCT2, NUM0, NUM1, NUM2, NUM3, MT, MTS, RS, CR,
                    X1, X2, USD, TEXT, SCORE, PPT)
from .support import na_row
from .spec import NA

LEG = 'CDEFGHIJ'
FLEG = 'DEFGHIJ'
NEW8 = 'BCDEFGHI'
NEW7 = 'BCDEFGH'
MA7 = 'EFGHIJK'
SCEN = ['Base Case', 'Bull Case', 'Bear Case', 'Stress Case']
LINKBLUE = 'FF0563C1'
SCOL = 'BCDE'          # Base / Bull / Bear / Stress on the new scenario tables
# Base-Case row of each driver in the scenario matrix on 02 Model Assumptions
DRV = dict(gdp=38, fx=46, cpi=50, ore=74, coal=78, tax=94, capexint=98, nwcdays=106,
           exitmult=110, wacc=114)


def _sup(item, value, unit, method, formula, primary, secondary, assumption, reasoning,
         cross, conf, linked, freq, last, comments, numfmt=None):
    return dict(item=item, value=value, unit=unit, method=method, formula=formula, primary=primary,
                secondary=secondary, assumption=assumption, reasoning=reasoning, cross=cross,
                conf=conf, linked=linked, freq=freq, last=last, comments=comments, numfmt=numfmt)


NO_PROB = ('No probability distribution over drivers is asserted. The four scenarios are discrete, '
           'internally coherent states of the world calibrated to observed history - they are not draws '
           'from a distribution, and attaching probabilities to them would imply a statistical basis that '
           'does not exist. Without probabilities a risk SCORE cannot be computed either. Use the impact '
           'column, which is measured, and apply your own house probabilities.')


# =================================================================== 21 Scenario Manager
def scenario_manager(wb, audit, eng):
    ws = wb['Scenario Manager']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- scenario dashboard rows 9-23: driver matrix FY2033E by scenario, plus the active value
    direct = {9: 'gdp', 13: 'ore', 14: 'coal', 17: 'cpi', 18: 'fx', 22: 'nwcdays', 23: 'tax'}
    active = {9: 8, 13: 17, 14: 18, 17: 11, 18: 10, 22: 25, 23: 22}
    fmt = {9: PCT1, 13: USD, 14: NUM0, 17: PCT1, 18: NUM1, 22: NUM0, 23: PCT2}
    for r, key in direct.items():
        for i, col in enumerate(SCOL):
            w.link(f'{col}{r}', f"='Model Assumptions'!$K${DRV[key] + i}", fmt[r])
        w.link(f'F{r}', f"='Model Assumptions'!$K${active[r]}", fmt[r])
    for i, (col, sc) in enumerate(zip(SCOL, SCEN)):
        w.link(f'{col}10', eng.ref(sc, 'consg', 7), PCT1)
        w.f(f'{col}11', f'=IFERROR({eng.cell(sc, "real", 7)}/{eng.cell(sc, "real", 6)}-1,"")', PCT1)
        w.link(f'{col}12', eng.ref(sc, 'util', 7), PCT1)
        w.link(f'{col}19', f"='Model Assumptions'!$K${DRV['wacc'] + i}", PCT1)
        w.link(f'{col}20', "='ESG Model'!$J$104", NUM0)
        w.f(f'{col}21', f"=IFERROR('Model Assumptions'!$K${DRV['capexint'] + i}"
                        f"/'Model Assumptions'!$J${DRV['capexint'] + i}-1,\"\")", PCT1)
    w.link('F10', "='Steel Demand Model'!$J$115", PCT1)
    w.f('F11', "=IFERROR('Steel Price Forecast'!$J$103/'Steel Price Forecast'!$I$103-1,\"\")", PCT1)
    w.link('F12', "='Capacity Utilisation'!$J$67", PCT1)
    w.link('F19', "='Model Assumptions'!$K$27", PCT1)
    w.link('F20', "='ESG Model'!$J$104", NUM0)
    w.f('F21', "=IFERROR('Model Assumptions'!$K$23/'Model Assumptions'!$J$23-1,\"\")", PCT1)
    for r in (15, 16):
        w.narow('BCDEF', r, 'Power cost and logistics inflation are not scenario drivers. Both sit inside '
                            'the CONVERSION COST block, which is indexed to RBI CPI through driver D12 - '
                            'and D12 IS a scenario driver, on row 17 above. Adding separate power and '
                            'freight paths would double-count the same inflation.')
    nas.append(na_row('Power cost and logistics inflation by scenario', 'B15:F16',
                      'Both are inside conversion cost, which is indexed to CPI through driver D12.',
                      unit='%'))
    rows.append(_sup(
        'Scenario dashboard', '=F9', 'various',
        'Cross-sheet reference to the FY2033E column of the SCENARIO DRIVER MATRIX on 02 Model Assumptions, '
        'with the Active Value column reading the resolved ACTIVE DRIVERS block',
        "='Model Assumptions'!$K$38 (Base) .. $K$41 (Stress); active =$K$8",
        'Each driver is sourced individually on 02 Model Assumptions column L and in 26 Sources',
        'The scenario engine below, which rebuilds the model from these same drivers',
        'The Active Value column must equal whichever scenario is selected on 01 Control Panel.',
        'This is the one-page answer to "what actually differs between the scenarios". Note that four of '
        'the fifteen rows are IDENTICAL across all four scenarios - carbon price, capex inflation, and the '
        'two blank inflation rows - because the model refuses to flex unsourced parameters to manufacture a '
        'scenario. That discipline is why the scenarios are defendable.',
        'The Active Value column must reproduce the selected scenario column exactly; check A27 on 25 Audit '
        'Checks tests the same thing on the tie-out block',
        'High', 'Model Assumptions, Scenario Manager engine, ESG Model', 'Live', 'Live',
        'Was entirely blank.', None))

    # ---- scenario summary rows 28-35 and comparison rows 40-46
    for col, sc in zip(SCOL, SCEN):
        w.link(f'{col}28', eng.ref(sc, 'rev', 7), CR)
        w.link(f'{col}29', eng.ref(sc, 'ebitda', 7), CR)
        w.link(f'{col}30', eng.ref(sc, 'margin', 7), PCT1)
        w.link(f'{col}31', eng.ref(sc, 'ufcf', 7), CR)
        w.link(f'{col}32', eng.ref(sc, 'nwc', 7), CR)
        w.link(f'{col}33', eng.ref(sc, 'util', 7), PCT1)
        w.f(f'{col}34', f'=AVERAGE({eng.rng(sc, "real")})', RS)
        w.link(f'{col}35', eng.ref(sc, 'ev', 7), CR)
        w.f(f'{col}40', f"=IFERROR(({eng.cell(sc, 'rev', 7)}/'Revenue Forecast'!$C$96)^(1/7)-1,\"\")", PCT1)
        w.f(f'{col}41', f"=IFERROR(({eng.cell(sc, 'ebitda', 7)}/'EBITDA Model'!$C$103)^(1/7)-1,\"\")", PCT1)
        w.f(f'{col}42', f"=IFERROR(({eng.cell(sc, 'ufcf', 7)}/{eng.cell(sc, 'ufcf', 1)})^(1/6)-1,\"\")",
            PCT1)
        w.f(f'{col}43', f'=MAX({eng.rng(sc, "margin")})', PCT1)
        w.f(f'{col}44', f'=MAX({eng.rng(sc, "util")})', PCT1)
        w.f(f'{col}45', f'=AVERAGE({eng.rng(sc, "roic")})', PCT1)
        w.link(f'{col}46', eng.ref(sc, 'exitmult', 7), X1)
    rows.append(_sup(
        'Scenario summary and comparison, FY2033E', '=B29', 'Rs cr',
        'Cross-sheet reference to the scenario engine; CAGRs computed against the FY2026A actual, peaks by '
        'MAX across the seven forecast years, ROIC as a true seven-year AVERAGE',
        "=IFERROR((engine EBITDA/'EBITDA Model'!$C$103)^(1/7)-1,\"\")",
        'Rebuilt live from the scenario driver matrix',
        'Independently computed tie-out values on rows 69 to 79 below',
        'Free cash flow CAGR is struck FY2027E to FY2033E, not off FY2026A, because FY2026A free cash flow '
        'is not modelled. It returns blank where either end is negative, which is correct rather than '
        'misleading.',
        'The comparison table is where the asymmetry of the investment case becomes visible: EBITDA CAGR '
        'runs from roughly -4.7% in the Bear Case to +17.5% in the Bull Case against a revenue CAGR range '
        'of only 6.7% to 11.9%. Average ROIC is the discipline test - it must be read against the WACC on '
        'row 19, and in the Bear and Stress cases it does not clear it.',
        'Every FY2033E figure must equal the independently computed value on rows 69 to 79; peak margin '
        'must sit inside the -9.2% to +26.0% observed envelope',
        'High', 'Scenario Manager engine, Revenue Forecast, EBITDA Model', 'Live', 'Live',
        'Was entirely blank.', CR))

    # ---- sensitivity matrix rows 51-58
    V, P, C_ = "'Steel Supply Model'!$J$101", "'Steel Price Forecast'!$J$103", "'Cost Curve'!$J$80"
    SH, WT = "'Model Assumptions'!$K$20", "'Model Assumptions'!$K$21"
    MULT = "'Model Assumptions'!$K$26"
    price_rows = (51, 52)
    for r in price_rows:
        w.f(f'C{r}', f"='Revenue Forecast'!$J$96*B{r}", CR)
        w.f(f'D{r}', f'={P}*B{r}*{V}/10', CR)
    for r in (53, 54):
        w.inp(f'C{r}', 0, CR)
        w.f(f'D{r}', f'=-{C_}*{SH}*{WT}*B{r}*{V}/10', CR)
    w.f('C55', "='Revenue Forecast'!$J$96*B55", CR)
    w.f('D55', "='EBITDA Model'!$J$103*B55", CR)
    w.f('C56', f"={V}*'Steel Price Forecast'!$J$97*'Model Assumptions'!$K$30*B56/10", CR)
    w.f('D56', f"='Steel Price Forecast'!$J$97*'Model Assumptions'!$K$30*B56*{V}/10", CR)
    w.inp('C57', 0, CR)
    w.f('D57', f'=-{C_}*{SH}*2*{WT}*B57*{V}/10', CR)
    w.inp('C58', 0, CR)
    w.inp('D58', 0, CR)
    for r in range(51, 58):
        w.f(f'E{r}', f'=D{r}*{MULT}', CR)
    w.f('E58', "=SUMPRODUCT('Comparable Valuation'!$D$127:$J$127,"
               "1/(1+'Model Assumptions'!$K$27+B58)^'Comparable Valuation'!$D$129:$J$129)"
               "+'Comparable Valuation'!$D$136/(1+'Model Assumptions'!$K$27+B58)^6.5"
               "-'Comparable Valuation'!$D$138", CR)
    rows.append(_sup(
        'Sensitivity matrix, FY2033E', '=D51', 'Rs cr',
        'Computed in-cell. Volume and price shocks flow through the revenue identity; input cost shocks '
        'through the basket weights; the WACC shock re-runs the discounted cash flow exactly',
        "=SUMPRODUCT('Comparable Valuation'!$D$127:$J$127,1/(1+WACC+shock)^discount periods)+TV/(1+WACC"
        "+shock)^6.5-base EV",
        'Computed from the model',
        '22 Sensitivity Analysis section A tornado, which ranks the same drivers',
        'Iron ore, coking coal and USD/INR have ZERO revenue impact - they are cost-side only. The WACC '
        'shock has zero revenue and zero EBITDA impact and affects enterprise value only.',
        'The WACC row is the only one that cannot be approximated, so it is not approximated: the formula '
        're-discounts all seven cash flows and the terminal value at the shocked rate and differences '
        'against the base enterprise value. A one percentage point rise in WACC costs materially more '
        'enterprise value than a 5% fall in the steel price, which is the argument for spending time on the '
        'equity risk premium and beta rather than on the commodity deck.',
        'Enterprise value impact must equal the EBITDA impact multiplied by the exit multiple on every row '
        'except WACC; signs must be intuitive throughout',
        'Medium', 'Revenue Forecast, EBITDA Model, Cost Curve, Comparable Valuation, Model Assumptions',
        'Live', 'FY2033E', 'Was entirely blank.', CR))
    return rows, nas


# ================================================================ 22 Sensitivity Analysis
# one-way ladder: row -> (variable key, shock value to write into column B or None to keep)
ONEWAY_SHOCKS = {27: 0.2, 28: 0.2, 29: 0.05, 30: 0.02, 31: 0.05, 32: 0.01, 33: 0.005}


def sensitivity(wb, audit, eng):
    ws = wb['Sensitivity Analysis']
    w = SheetWriter(ws, audit)
    rows, nas = [], []
    V, P, C_ = "'Steel Supply Model'!$J$101", "'Steel Price Forecast'!$J$103", "'Cost Curve'!$J$80"
    SH, WT = "'Model Assumptions'!$K$20", "'Model Assumptions'!$K$21"
    MULT, TAX = "'Model Assumptions'!$K$26", "'Model Assumptions'!$K$22"
    REV, EB, FCF = "'Revenue Forecast'!$J$96", "'EBITDA Model'!$J$103", "'Cash Flow Model'!$J$114"

    # ---- executive summary rows 9-14
    band = [(9, "'Revenue Forecast'!B89:E89", CR), (10, "'EBITDA Model'!B95:E95", CR),
            (11, "'Cash Flow Model'!B85:E85", CR), (12, "'Scenario Manager'!B35:E35", CR),
            (13, "'Model Assumptions'!$K$110:$K$113", X1), (14, "'Margin Analysis'!B82:E82", PCT1)]
    base = {9: REV, 10: EB, 11: FCF, 12: "'Comparable Valuation'!$D$138",
            13: "'Model Assumptions'!$K$26", 14: "'Margin Analysis'!$J$91"}
    for r, rng, f in band:
        w.link(f'B{r}', f'={base[r]}', f)
        w.f(f'C{r}', f'=MAX({rng})', f)
        w.f(f'D{r}', f'=MIN({rng})', f)
        w.link(f'E{r}', '=$A$97', TEXT)

    # ---- one-way sensitivity rows 19-33
    for r, v in ONEWAY_SHOCKS.items():
        w.inp(f'B{r}', v, PCT0)
    for r in range(19, 26):                      # steel price ladder
        if r == 22:
            w.f(f'C{r}', f'={REV}', CR)
            w.f(f'D{r}', f'={EB}', CR)
        else:
            w.f(f'C{r}', f'={REV}*(1+B{r})', CR)
            w.f(f'D{r}', f'={EB}+{P}*B{r}*{V}/10', CR)
    for r in (26, 27):                           # iron ore
        w.f(f'C{r}', f'={REV}', CR)
        w.f(f'D{r}', f'={EB}-{C_}*{SH}*{WT}*B{r}*{V}/10', CR)
    w.f('C28', f'={REV}', CR)                    # coking coal, same basket weight
    w.f('D28', f'={EB}-{C_}*{SH}*{WT}*B28*{V}/10', CR)
    w.f('C29', f'={REV}*(1+B29)', CR)            # demand
    w.f('D29', f'={EB}*(1+B29)', CR)
    w.f('C30', f"={REV}+{V}*'Steel Price Forecast'!$J$97*'Model Assumptions'!$K$30*B30/10", CR)
    w.f('D30', f"={EB}+'Steel Price Forecast'!$J$97*'Model Assumptions'!$K$30*B30*{V}/10", CR)
    w.f('C31', f'={REV}', CR)                    # USD/INR moves 90% of the basket
    w.f('D31', f'={EB}-{C_}*{SH}*2*{WT}*B31*{V}/10', CR)
    w.f('C32', f'={REV}', CR)                    # WACC: no operating impact
    w.f('D32', f'={EB}', CR)
    for r in range(19, 33):
        w.f(f'E{r}', f'={FCF}+(D{r}-{EB})*(1-{TAX})', CR)
        if r == 32:
            w.f('F32', "=SUMPRODUCT('Comparable Valuation'!$D$127:$J$127,"
                       "1/(1+'Model Assumptions'!$K$27+B32)^'Comparable Valuation'!$D$129:$J$129)"
                       "+'Comparable Valuation'!$D$136/(1+'Model Assumptions'!$K$27+B32)^6.5", CR)
        else:
            w.f(f'F{r}', f'=D{r}*{MULT}', CR)
    TG_NA = ('This model does NOT use a terminal growth rate. Terminal value is struck as an exit '
             'EV/EBITDA multiple (driver D19), which is the convention for a cyclical industry where a '
             'perpetuity growth assumption would imply the cycle has been abolished. A terminal growth '
             'sensitivity therefore has no meaning on the primary valuation. The two-way table on rows 39 '
             'to 43 DOES provide a perpetuity-based enterprise value across a WACC and growth grid, as an '
             'alternative cross-check.')
    w.narow('CDEF', 33, TG_NA)
    nas.append(na_row('One-way sensitivity to terminal growth', 'C33:F33', TG_NA, unit='Rs cr'))
    rows.append(_sup(
        'One-way sensitivity ladder, FY2033E', '=D24', 'Rs cr',
        'Computed in-cell as LEVELS, not deltas, so each row is a readable state of the world',
        '=EBITDA+realisation x shock x volume/10 for price; =EBITDA-cash cost x RM share x weight x shock x '
        'volume/10 for inputs',
        'Computed from the model',
        'Section A tornado below, which expresses the same shocks as percentages',
        'Free cash flow is re-struck by taxing the EBITDA delta at the effective rate; capex and working '
        'capital are held at base, so the free cash flow column understates the true sensitivity.',
        'Showing levels rather than deltas is deliberate: an investment committee reads "EBITDA is Rs '
        '1,27,580 crore if realisation is 10% lower" far more readily than "EBITDA falls by Rs 1,88,208 '
        'crore". The tornado immediately below gives the same information as percentages for ranking.',
        'The Base row must reproduce the live model exactly; the -10% steel price row must reproduce the '
        'tornado value of Rs 1,27,580 crore on row 97',
        'Medium', 'Revenue Forecast, EBITDA Model, Cash Flow Model, Cost Curve, Model Assumptions',
        'Live', 'FY2033E',
        'Column B carried "..." placeholders on rows 27 to 33; those are now explicit shock sizes and are '
        'blue live inputs.', CR))

    # ---- two-way: EV vs WACC and terminal growth, rows 39-43
    for r in range(39, 44):
        for col in 'BCDEF':
            w.f(f'{col}{r}',
                f"=IFERROR(SUMPRODUCT('Comparable Valuation'!$D$127:$J$127,"
                f"1/(1+{col}$38)^'Comparable Valuation'!$D$129:$J$129)"
                f"+'Comparable Valuation'!$J$127*(1+$A{r})/({col}$38-$A{r})/(1+{col}$38)^6.5,\"\")", CR)
    rows.append(_sup(
        'Two-way enterprise value: WACC against terminal growth', '=D41', 'Rs cr',
        'Computed in-cell: the seven explicit cash flows re-discounted at the column WACC, plus a Gordon '
        'growth terminal value at the row growth rate',
        "=SUMPRODUCT(FCF,1/(1+WACC)^period)+FY2033E FCF x (1+g)/(WACC-g)/(1+WACC)^6.5",
        'Free cash flows from 16 Cash Flow Model; discount convention from 23 Comparable Valuation row 129',
        '23 Comparable Valuation section C, which uses an exit multiple instead',
        'AN ALTERNATIVE TERMINAL VALUE. The primary valuation uses an exit multiple, not perpetuity growth.',
        'Presenting both is the point. A perpetuity growth terminal value and an exit multiple answer the '
        'same question with different assumptions, and the gap between them is a measure of how much of the '
        'answer is terminal-value judgement. Note that at the base WACC of 12% the exit-multiple approach '
        'and a 3% perpetuity give materially different answers, which is exactly the disclosure an '
        'investment committee needs.',
        'The mid-cell should be read against the exit-multiple enterprise value on 23 Comparable Valuation '
        'row 138; both use the same seven explicit cash flows',
        'Medium', 'Comparable Valuation, Cash Flow Model', 'Live', 'FY2033E', 'Was entirely blank.', CR))

    # ---- two-way: iron ore against steel price, rows 47-51
    def num(ref):
        return f'IF(ISNUMBER({ref}),{ref},0)'
    for r in range(47, 52):
        for col in 'BCDEF':
            w.f(f'{col}{r}',
                f'=({EB}+{P}*{num(f"{col}$46")}*{V}/10'
                f'-{C_}*{SH}*{WT}*{num(f"$A{r}")}*{V}/10)*{MULT}', CR)
    rows.append(_sup(
        'Two-way enterprise value: iron ore against steel price', '=D49', 'Rs cr',
        'Computed in-cell: EBITDA re-struck for both shocks simultaneously, capitalised at the exit multiple',
        '=(EBITDA+price effect-iron ore effect)*exit multiple, with IF(ISNUMBER()) guards so the "Base" '
        'label in the header does not break the arithmetic',
        'Computed from the model', 'Section A tornado',
        'Both shocks are applied to the FY2033E terminal year only, and the exit multiple is held constant.',
        'The table shows the asymmetry that matters: a 10% steel price move is worth roughly three times a '
        '10% iron ore move, because realisation applies to the whole tonne while iron ore applies to 27% of '
        'cash cost. An investment committee looking for the one variable to underwrite should read the '
        'columns, not the rows.',
        'The centre cell must equal base EBITDA multiplied by the exit multiple; the grid must be monotonic '
        'in both directions',
        'Medium', 'EBITDA Model, Cost Curve, Steel Price Forecast, Model Assumptions', 'Live', 'FY2033E',
        'Was entirely blank.', CR))

    # ---- downside / upside table rows 56-63
    pair = {56: (20, 24), 57: (27, 26), 58: (28, 28), 59: (29, 29), 60: (30, 30), 63: (31, 31)}
    for r, (dn, up) in pair.items():
        w.f(f'B{r}', f'=D{dn}-{EB}', CR)
        w.f(f'C{r}', f'=D{up}-{EB}' if dn != up else f'=-(D{up}-{EB})', CR)
    w.f('B61', f"=F32-'Comparable Valuation'!$D$138", CR)
    w.f('C61', f"=-(F32-'Comparable Valuation'!$D$138)", CR)
    w.narow('BC', 62, TG_NA)
    rows.append(_sup(
        'Downside and upside impact', '=B56', 'Rs cr',
        'Computed in-cell by differencing the one-way ladder above against the base case',
        '=D20-EBITDA for downside, =D24-EBITDA for upside',
        'Computed from the model', 'The one-way ladder on rows 19 to 32',
        'Where only one direction of shock is laddered above (iron ore, coking coal, demand, utilisation, '
        'USD/INR), the opposite direction is taken as its exact mirror image.',
        'Mirroring is legitimate here because every one of those channels is LINEAR in the model - cost is '
        'a fixed share of a fixed basket, so a +10% and a -10% move are equal and opposite. The steel price '
        'and WACC rows are not mirrored, because the ladder carries both directions explicitly and WACC '
        'discounting is convex.',
        'Downside and upside must be equal and opposite on the linear channels; the WACC row must reconcile '
        'to the enterprise value column of the ladder',
        'Medium', 'This sheet', 'Live', 'FY2033E', 'Was entirely blank.', CR))

    # ---- break-even analysis rows 68-72
    BE = "('Cash Flow Model'!$J$110+'Cash Flow Model'!$J$113+'Cash Flow Model'!$J$102)"
    w.f('B68', f"={C_}/'Steel Price Forecast'!$J$102", RS)
    w.f('B69', f'=IFERROR({BE}/{REV},"")', PCT1)
    w.f('B70', f"=IFERROR('Capacity Utilisation'!$J$67*{BE}/{EB},\"\")", PCT1)
    w.f('B71', f"=IFERROR((1+'Steel Demand Model'!$J$115)*{BE}/{EB}-1,\"\")", PCT1)
    w.f('B72', f"='Model Assumptions'!$K$17*(1+({P}-{C_})/({C_}*{SH}*{WT}))", USD)
    rows.append(_sup(
        'Break-even analysis, FY2033E', '=B68', 'Rs/t',
        'Computed in-cell by solving each identity for the value at which the relevant output reaches zero',
        "=cash cost/price adjustment factor for price; =(capex+dNWC+tax)/revenue for the free-cash-flow "
        "break-even margin",
        'Computed from the model',
        'Observed history: the worst single margin observation in the database is SAIL at -7% in FY2016',
        'Two DIFFERENT break-evens are shown and the distinction matters: row 68 is where EBITDA reaches '
        'zero, rows 69 to 71 are where FREE CASH FLOW reaches zero.',
        'The free-cash-flow break-even is the number that actually constrains the industry. EBITDA turns '
        'negative only if realisation falls to about Rs 57,400/t, roughly 17% below base - but free cash '
        'flow turns negative at a far higher price, because the industry is committing capex of about Rs '
        '2.4 lakh crore a year. That gap between the two break-evens IS the capital cycle.',
        'The iron ore break-even must lie far above any observed price, confirming that iron ore alone '
        'cannot break the industry; the price break-even must equal cash cost divided by the price '
        'adjustment factor',
        'Medium', 'Cash Flow Model, Cost Curve, Steel Price Forecast, Capacity Utilisation', 'Live',
        'FY2033E', 'Was entirely blank.', RS))

    # ---- risk ranking rows 77-83
    impact = {77: '$F$97', 78: '$F$101', 79: '$F$102', 80: '$F$99'}
    for r, ref in impact.items():
        w.f(f'B{r}', f'=ABS({ref})', PCT1)
    w.f('B81', f"=IFERROR('Trade Model'!$J$83-'Trade Model'!$C$83,\"\")", PCT1)
    w.f('B82', f"=IFERROR('ESG Model'!$J$112/{EB},\"\")", PCT1)
    w.f('B83', f"=IFERROR(ABS(F32-'Comparable Valuation'!$D$138)/'Comparable Valuation'!$D$138,\"\")", PCT1)
    for r in range(77, 84):
        w.na(f'C{r}', NO_PROB)
        w.na(f'D{r}', NO_PROB)
    nas.append(na_row('Risk probability and risk score', 'C77:D83', NO_PROB))
    rows.append(_sup(
        'Risk ranking - measured impact', '=B77', '% of FY2033E EBITDA',
        'Cross-sheet reference to the tornado in section A, plus three impacts computed directly',
        '=ABS($F$97) for steel price; =ESG CBAM cost/EBITDA for carbon; =ABS(EV at WACC+1% - base EV)/base '
        'EV for WACC',
        'Computed from the model; the trade-policy row uses the change in import penetration once the '
        'safeguard duty expires',
        'Section A tornado; 19 Trade Model row 83; 20 ESG Model row 112',
        'Impact is expressed as a percentage of the relevant base - EBITDA for operating risks, enterprise '
        'value for WACC.',
        'The ranking is unambiguous and it is the single most useful output of this sheet: realisation '
        '(59.6%), then USD/INR (23.7%), then GDP and elasticity (12.2% each), then iron ore (12.1%) and '
        'coking coal (11.6%). Carbon cost is under 1%. Anyone underwriting this industry should spend their '
        'diligence budget in that order.',
        'Impacts must tie to the tornado; the probability and score columns are deliberately blank',
        'Medium', 'Trade Model, ESG Model, Comparable Valuation', 'Live', 'FY2033E',
        'Was entirely blank.', PCT1))
    return rows, nas



# ================================================================ 23 Comparable Valuation
# new row -> legacy trading-comps row for the six names carried there
CV_LEGACY = {20: 108, 21: 109, 22: 111, 23: 110, 24: 112, 25: 113, 26: 114}
CV_REVENUE = {20: 111, 21: 112, 22: 114, 23: 113, 24: 115, 25: 116, 26: 117}
# supporting names: (row, revenue Rs cr, standardised EBITDA Rs cr) from the Master Database
CV_SUPPORT = [(27, 18552, 2333), (28, 5381, 1253), (29, 23079, 1802), (30, 4890, 172),
              (31, 4419, 716), (32, None, None), (33, 6889, 850), (34, 6133, None)]
NO_BOOK = ('Book value and shareholders equity were not sourced in this research cycle, so price to book '
           'cannot be computed for any name. This is standing limitation 6 on 25 Audit Checks. The Master '
           'Industry Database carries total assets and borrowings but not equity.')
NO_LISTED = ('Unlisted. RINL is wholly Government-owned and publishes no balance sheet or earnings; AM/NS '
             'India is an unlisted 60:40 joint venture reporting on a 100% US dollar basis. Neither has a '
             'share price, market capitalisation or earnings per share.')
NO_PRICE_SUPP = ('Share price, and therefore market capitalisation, enterprise value, earnings and every '
                 'trading multiple, were not sourced for the supporting coverage names. Revenue and '
                 'standardised EBITDA ARE available and are shown. Enter share prices and net debt to '
                 'complete these rows.')


def comparable_valuation(wb, audit, eng):
    ws = wb['Comparable Valuation']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- peer universe rows 20-34
    for r, lr in CV_LEGACY.items():
        rev = CV_REVENUE[r]
        if r <= 24:
            w.link(f'B{r}', f'=$H${lr}', CR)
            w.link(f'C{r}', f'=$I${lr}', CR)
            w.link(f'D{r}', f"='Revenue Forecast'!$C${rev}", CR)
            w.link(f'E{r}', f'=$B${lr}', CR)
            w.link(f'F{r}', f'=$D${lr}', CR)
            w.link(f'H{r}', f'=$C${lr}', CR)
        elif r == 25:
            w.link(f'D{r}', f"='Revenue Forecast'!$C${rev}", CR)
            w.link(f'E{r}', f'=$B${lr}', CR)
            for col in 'BCFH':
                w.na(f'{col}{r}', NO_LISTED)
        else:
            w.link(f'D{r}', f"='Revenue Forecast'!$C${rev}", CR)
            for col in 'BCEFH':
                w.na(f'{col}{r}', NO_LISTED)
        w.na(f'G{r}', NO_BOOK)
    for r, rev, ebt in CV_SUPPORT:
        if rev is None:
            for col in 'BCDEFGH':
                w.na(f'{col}{r}', 'Uttam Galva Steels was delisted on 08-Dec-2022 following the '
                                  'NCLT-approved resolution plan. No FY2026 financials are published.')
            continue
        w.db(f'D{r}', rev, CR, ref=f"='[Master Industry Database.xlsx]Revenue'!$L${20 + r - 21}")
        if ebt is None:
            w.na(f'E{r}', "Vedanta's FY2026 release does not disclose a separate EBITDA for ESL Steel, "
                          'which is now a subsidiary rather than a listed entity.')
        else:
            w.db(f'E{r}', ebt, CR, ref=f"='[Master Industry Database.xlsx]EBITDA'!$O${20 + r - 21}")
        for col in 'BCFH':
            w.na(f'{col}{r}', NO_PRICE_SUPP)
        w.na(f'G{r}', NO_BOOK)
    nas.append(na_row('Peer universe - book value for every name', 'G20:G34', NO_BOOK, unit='Rs cr'))
    nas.append(na_row('Peer universe - market capitalisation, enterprise value, earnings and net debt for '
                      'the supporting names and the two unlisted producers', 'B25:H34',
                      NO_PRICE_SUPP + ' ' + NO_LISTED, unit='Rs cr'))
    rows.append(_sup(
        'Peer universe, FY2026A', '=D20', 'Rs cr',
        'Market capitalisation and enterprise value linked to the trading-comps block in the research block, '
        'which derives them from a user-entered share price; revenue and EBITDA for the supporting names '
        'imported from the Master Industry Database (RED)',
        '=$H$108 market cap, =$I$108 enterprise value, =$B$108 EBITDA',
        'Company results releases and issuer press releases. FY2026 revenue for Tata Steel, JSW Steel and '
        'SAIL ties EXACTLY to the issuer press releases',
        'Master Industry Database Revenue section A and EBITDA section A (standardised series)',
        'Share prices on research block rows 108 to 112 are USER INPUTS, not sourced market data.',
        'Revenue and standardised EBITDA are available for 13 of the 15 names, which is enough to size the '
        'universe; share prices exist for only the five listed majors, and they are user inputs rather than '
        'sourced prints. Every multiple derived below therefore inherits that limitation, which is why the '
        'replacement-cost benchmark in section B and the discounted cash flow in section C carry more weight '
        'in this model than the trading comparables do.',
        'FY2026 revenue for the three largest names ties exactly to issuer releases; standardised EBITDA is '
        'used throughout rather than company-headlined EBITDA so the names are comparable',
        'High for revenue and EBITDA, Low for anything derived from a share price',
        'Revenue Forecast, Master Industry Database', 'Quarterly, on each results season', 'FY2026A',
        'Was entirely blank. SAIL EBITDA is standalone while its revenue is consolidated - see the warning '
        'in the research block.', CR))

    # ---- trading multiples rows 39-53 and statistics rows 54-59
    for r in range(39, 54):
        peer = r - 19
        lr = CV_LEGACY.get(peer)
        if lr and peer <= 24:
            w.link(f'B{r}', f'=$J${lr}', X2)
            w.f(f'C{r}', f'=IFERROR($I${lr}/D{peer},"")', X2)
            w.f(f'D{r}', f'=IFERROR($H${lr}/$D${lr},"")', X1)
            w.f(f'F{r}', f'=IFERROR($I${lr}/$K${lr}*10,"")', RS)
            w.f(f'G{r}', f"=IFERROR($I${lr}/'Revenue Forecast'!$C${CV_REVENUE[peer] - 10}*10,\"\")", RS)
        else:
            for col in 'BCDFG':
                w.na(f'{col}{r}', NO_PRICE_SUPP if peer > 26 else NO_LISTED)
        w.na(f'E{r}', NO_BOOK)
    stats = [(54, 'AVERAGE'), (55, 'MEDIAN'), (58, 'MAX'), (59, 'MIN')]
    for r, fn in stats:
        for col in 'BCDFG':
            w.f(f'{col}{r}', f'=IFERROR({fn}({col}39:{col}53),"")', X2 if col in 'BC' else RS)
        w.na(f'E{r}', NO_BOOK)
    for r, q in ((56, 0.25), (57, 0.75)):
        for col in 'BCDFG':
            w.f(f'{col}{r}', f'=IFERROR(PERCENTILE({col}39:{col}53,{q}),"")',
                X2 if col in 'BC' else RS)
        w.na(f'E{r}', NO_BOOK)
    nas.append(na_row('Price to book for every name and every statistic', 'E39:E59', NO_BOOK, unit='x'))
    nas.append(na_row('Trading multiples for the supporting and unlisted names', 'B44:G53',
                      NO_PRICE_SUPP, unit='x'))
    rows.append(_sup(
        'Trading multiples and peer statistics', '=B55', 'x EV/EBITDA',
        'Computed in-cell from the trading-comps block; statistics use AVERAGE, MEDIAN, PERCENTILE, MAX and '
        'MIN across the five populated rows',
        '=IFERROR(MEDIAN(B39:B53),"") ; =IFERROR(PERCENTILE(B39:B53,0.25),"")',
        'Derived from user-entered share prices and Master-Database FY2026 EBITDA and net debt',
        'Research block row 116 computes the same median independently',
        'Only the five listed majors populate, so every statistic is struck on a sample of five.',
        'A five-name sample is small but it is the entire listed carbon and stainless steel universe of any '
        'scale in India, so it is the relevant population rather than a sample of it. The median EV/EBITDA '
        'that results is the number to compare against the 6.0x exit multiple in driver D19 - and research '
        'block row 116 exists precisely to make that comparison, which is check A21 on 25 Audit Checks.',
        'The median must equal research block row 116; enterprise value per tonne of capacity must be '
        'sanity-checked against the replacement cost of Rs 55,000/t in section B',
        'Low - every multiple depends on a user-entered share price',
        'This sheet, Revenue Forecast', 'Daily for prices, quarterly for financials', 'FY2026A',
        'Was entirely blank.', X2))

    rows.append(_sup(
        'DEFECT FLAGGED - enterprise value per tonne of capacity', '=F39', 'Rs/t',
        'Computed in-cell in this table as EV (Rs cr) / capacity (Mt) x 10, which is the correct unit bridge',
        '=IFERROR($I$108/$K$108*10,"")',
        'Derived from the trading-comps block', 'Section B replacement cost benchmark',
        'Rs crore divided by million tonnes must be multiplied by TEN to give rupees per tonne, because '
        'Rs 1 crore = 10^7 rupees and 1 Mt = 10^6 tonnes.',
        'UNIT ERROR IN THE RESEARCH BLOCK. Cell L108 and the cells below it multiply by 10,000 instead of '
        '10, so the enterprise value per tonne they report is one thousand times too large. Those cells are '
        'inside the research block and were deliberately NOT altered. Every new table on this sheet uses '
        'the correct factor of 10, which is why this table and research block column L disagree by 1,000x. '
        'Correct L108 to =IFERROR(I108/K108*10,"") before transaction use.',
        'Tata Steel enterprise value per tonne should read roughly Rs 1,22,800/t, which sits sensibly above '
        'the Rs 55,000/t replacement cost; research block column L reports Rs 12.3 crore per tonne, which '
        'is obviously wrong on inspection',
        'High - this is arithmetic, not judgement', 'This sheet', 'n/a', 'FY2026A',
        'Found by the unit-bridge check in tools/audit.py.', RS))

    # ---- historical multiples rows 64-69
    for r in range(64, 70):
        w.narow('BCDE', r, NA['hist_multiple'])
    nas.append(na_row('Historical trading multiples FY2021-FY2026', 'B64:E69', NA['hist_multiple'],
                      unit='x'))

    # ---- valuation bridge rows 74-79
    PAT26 = "('EBITDA Model'!$C$103-'Cash Flow Model'!$C$99)*(1-'Cash Flow Model'!$C$101)"
    w.link('B74', '=$B$55', X2)
    w.link('C74', "='EBITDA Model'!$C$103", CR)
    w.f('D74', '=IFERROR(B74*C74,"")', CR)
    w.link('B75', '=$C$55', X2)
    w.link('C75', "='Revenue Forecast'!$C$96", CR)
    w.f('D75', '=IFERROR(B75*C75,"")', CR)
    w.link('B76', '=$D$55', X1)
    w.f('C76', f'={PAT26}', CR)
    w.f('D76', "=IFERROR(B76*C76+'Capital Allocation'!$C$80,\"\")", CR)
    for col in 'BCD':
        w.na(f'{col}77', NO_BOOK)
    w.link('B78', '=$F$55', RS)
    w.link('C78', "='Capacity Forecast'!$C$101", MT)
    w.f('D78', '=IFERROR(B78*C78/10,"")', CR)
    w.link('B79', '=$G$55', RS)
    w.link('C79', "='Steel Supply Model'!$C$101", MT)
    w.f('D79', '=IFERROR(B79*C79/10,"")', CR)
    nas.append(na_row('Valuation bridge - price to book method', 'B77:D77', NO_BOOK, unit='Rs cr'))

    # ---- implied enterprise valuation rows 84-89
    for r, src in ((84, 'D74'), (85, 'D75'), (86, 'D76'), (88, 'D78')):
        w.f(f'B{r}', f'=IFERROR(${src},"")', CR)
        w.f(f'C{r}', f"=IFERROR(B{r}/'Capacity Forecast'!$C$101*10,\"\")", RS)
        w.f(f'D{r}', '=IFERROR(B%d/$E$121,"")' % r, PCT1)
    for col in 'BCD':
        w.na(f'{col}87', NO_BOOK)
    w.f('B89', '=IFERROR(AVERAGE(B84:B88),"")', CR)
    w.f('C89', "=IFERROR(B89/'Capacity Forecast'!$C$101*10,\"\")", RS)
    w.f('D89', '=IFERROR(B89/$E$121,"")', PCT1)
    rows.append(_sup(
        'Implied enterprise valuation and the replacement-cost cross-check', '=B89', 'Rs cr',
        'Computed in-cell: each method applied to its FY2026A metric, then expressed per tonne of capacity '
        'and as a percentage of replacement cost',
        '=IFERROR(B89/$E$121,"") where E121 is 220.4 Mtpa at the capex intensity of Rs 55,000/t',
        'Peer median multiples from the block above; replacement cost from driver D16',
        'Section B row 121, the replacement cost of Indian capacity',
        'Applied to FY2026A metrics, not FY2033E, because a traded multiple is a current-basis multiple. '
        'Price to book cannot be computed.',
        'EXPRESSING EVERY METHOD AS A PERCENTAGE OF REPLACEMENT COST IS THE MOST USEFUL DISCIPLINE ON THIS '
        'SHEET. Replacement cost of Rs 12.1 lakh crore for 220.4 Mtpa is a hard economic bound: no rational '
        'buyer pays materially above it when they could build, and sustained trading below it is the '
        'signal that the industry will not add capacity. It is also the only valuation anchor here that '
        'does not depend on a user-entered share price.',
        'The average must be read against the discounted cash flow enterprise value in section C and the '
        'replacement cost in section B; all three should be within a defensible range of each other',
        'Low - inherits the share-price limitation', 'Capacity Forecast, Revenue Forecast, EBITDA Model',
        'Quarterly', 'FY2026A', 'Was entirely blank.', CR))

    # ---- scenario valuation rows 94-97 - uses driver D19, NOT the hard-coded 7.0x in the DCF
    for i, (r, sc) in enumerate(zip(range(94, 98), SCEN)):
        w.link(f'B{r}', f"='Model Assumptions'!$K${110 + i}", X1)
        w.link(f'C{r}', eng.ref(sc, 'ev', 7), CR)
    rows.append(_sup(
        'Scenario valuation, FY2033E', '=C94', 'Rs cr',
        'Cross-sheet reference: exit multiple from driver D19 by scenario, applied to the scenario engine '
        'FY2033E EBITDA',
        "='Model Assumptions'!$K$110 x 'Scenario Manager'!$J$<engine EBITDA row>",
        'Driver D19 exit EV/EBITDA multiple - NOT SOURCED, flagged Low confidence',
        'Section C discounted cash flow; section B replacement cost',
        'Uses the SCENARIO exit multiples of 6.0x base, 7.0x bull, 5.0x bear and 4.0x stress. Flexing the '
        'multiple with the scenario is deliberate: a terminal multiple should not be constant across states '
        'of the world.',
        'IMPORTANT CONTRADICTION TO RESOLVE. Cell D135 in the research block hard-codes the discounted cash '
        'flow exit multiple at 7.0x, while driver D19 base case is 6.0x. This table uses the DRIVER, so the '
        'two valuations on this sheet are struck on different multiples. D135 sits inside the research block '
        'and was deliberately NOT altered. Repoint D135 to \'Model Assumptions\'!$K$26 before transaction '
        'use and the whole sheet becomes internally consistent.',
        'The base-case enterprise value here should be compared with the discounted cash flow on row 138; '
        'the difference is almost entirely the 6.0x versus 7.0x multiple',
        'Low', 'Model Assumptions, Scenario Manager engine', 'Quarterly', 'FY2033E',
        'Was blank. The multiple range of 4.0x to 7.0x is unsourced and is the single largest swing factor '
        'in the valuation.', CR))

    # ---- executive summary rows 9-15
    for i, (col, sc) in enumerate(zip(SCOL, SCEN)):
        ev = f'$C${94 + i}'
        nd = f"('Capital Allocation'!$C$80-SUM({eng.rng(sc, 'ufcf')}))"
        w.f(f'{col}9', f'={ev}', CR)
        w.f(f'{col}10', f'={ev}-{nd}', CR)
        w.link(f'{col}11', f"='Model Assumptions'!$K${110 + i}", X1)
        w.f(f'{col}12', f'=IFERROR({ev}/{eng.cell(sc, "rev", 7)},"")', X2)
        w.f(f'{col}13', f'=IFERROR(({ev}-{nd})/{eng.cell(sc, "nopat", 7)},"")', X1)
        w.na(f'{col}14', NO_BOOK)
        w.f(f'{col}15', f'=IF({ev}<$E$121*0.8,"Attractive - below replacement cost",'
                        f'IF({ev}<$E$121*1.2,"Fair - around replacement cost",'
                        f'"Full - above replacement cost"))', TEXT)
    nas.append(na_row('Price to book by scenario', 'B14:E14', NO_BOOK, unit='x'))
    rows.append(_sup(
        'Valuation executive summary', '=B9', 'Rs cr',
        'Computed in-cell: enterprise value from the scenario valuation; equity value by deducting net debt '
        'rolled forward from FY2026A at cumulative unlevered free cash flow',
        "=$C$94-('Capital Allocation'!$C$80-SUM(engine UFCF FY2027E:FY2033E))",
        'Driver D19 exit multiple; FY2026A net debt of Rs 1,85,001 cr from the Master Industry Database',
        'Section C discounted cash flow; section B replacement cost',
        'The net debt roll-forward here EXCLUDES interest, so equity value is marginally overstated in the '
        'strong scenarios and understated in the weak ones.',
        'The recommendation row is deliberately mechanical - it compares enterprise value against '
        'replacement cost with an 80% and 120% band - because a recommendation that is a formula can be '
        'audited, whereas a recommendation that is text cannot. Read it as a valuation screen, not as '
        'investment advice: the exit multiple it rests on is unsourced.',
        'Enterprise value must equal the scenario valuation table; price to earnings uses NOPAT rather than '
        'net income because the model is unlevered, and that is disclosed',
        'Low', 'Model Assumptions, Scenario Manager engine, Capital Allocation', 'Quarterly', 'FY2033E',
        'Was entirely blank.', CR))
    return rows, nas


# ================================================================== 24 Industry Dashboard
DASH_NAV = {'B': 'Index', 'C': 'Control Panel', 'D': 'Model Assumptions',
            'E': 'Comparable Valuation', 'F': 'Sensitivity Analysis', 'G': 'Database Import',
            'H': 'Sources'}


def dashboard(wb, audit, eng):
    ws = wb['Industry Dashboard']
    w = SheetWriter(ws, audit)
    rows = []

    for col, target in DASH_NAV.items():
        link_to(ws, f'{col}7', target)

    # the leverage row shipped with FY2026A only and empty strings across the forecast
    for lc in FLEG:
        w.link(f'{lc}55', f"='Capital Allocation'!{lc}$84", X2)
    rows.append(_sup(
        'Net debt / EBITDA', '=J55', 'x',
        'Cross-sheet reference to 17 Capital Allocation',
        "='Capital Allocation'!D$84",
        'FY2026A net debt of Rs 1,85,001 cr derived from the Master Industry Database; EBITDA from the model',
        '17 Capital Allocation rows 80 to 85',
        'Net debt is rolled forward from unlevered free cash flow less illustrative interest. No new '
        'borrowing is assumed.',
        'DEFECT CORRECTED. This row shipped with a FY2026A value of 1.07x and EMPTY STRINGS in all seven '
        'forecast columns, so the dashboard silently under-reported the balance sheet. It is now live across '
        'the horizon and shows the industry deleveraging to a net cash position by about FY2030E in the base '
        'case.',
        'Must equal 17 Capital Allocation row 84 in every year; check A22 on 25 Audit Checks tests the net '
        'debt roll-forward',
        'Medium', 'Capital Allocation', 'Quarterly', 'FY2026A',
        'The only cell inside a research block that was written to, and only because it was blank.', X2))
    rows.append(_sup(
        'Dashboard navigation buttons', None, 'n/a',
        'Internal workbook hyperlinks', "location = 'Sheet Name'!A1",
        'n/a', 'n/a', 'None.',
        'The seven navigation labels in row 7 were text only. Each now jumps to its target sheet.',
        'Targets verified against the workbook sheet list', 'High', 'All model sheets',
        'On any sheet rename', 'n/a', 'Hyperlinks only.', None))
    return rows, []


# ============================================================================== navigation
def index_links(wb, audit):
    ws = wb['Index']
    for r in range(6, 40):
        name = ws[f'B{r}'].value
        if isinstance(name, str) and name in wb.sheetnames:
            link_to(ws, f'B{r}', name)
    return [_sup(
        'Sheet navigation hyperlinks', None, 'n/a', 'Internal workbook hyperlinks',
        "location = 'Sheet Name'!A1", 'n/a', 'n/a', 'None.',
        'The index instructed the reader to "click a sheet name to navigate" but the names carried no '
        'hyperlinks. All 31 are now linked.',
        'Every target verified against the workbook sheet list', 'High', 'All sheets',
        'On any sheet rename', 'n/a', 'Hyperlinks only - no formatting changed apart from the underline '
        'that signals a link.', None)], []
