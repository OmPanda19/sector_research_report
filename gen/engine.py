"""Scenario engine.

The live model resolves exactly ONE scenario at a time (whichever is selected on
01 Control Panel). Every sheet, however, carries a four-scenario comparison
table. Rather than hard-code those tables - which would guarantee they drift away
from the drivers - this module writes a single engine block that rebuilds the whole
model for all four scenarios directly off the SCENARIO DRIVER MATRIX on
02 Model Assumptions. Every scenario table in the workbook then links into this one
block, so there is still exactly one source of truth.

The engine reproduces the live model line for line, which is what makes it
auditable: for the active scenario its output must equal the live model, and its
FY2033E column must equal the independently computed tie-out values already
carried on 24 Scenario Manager rows 69 to 79.
"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SHEET = 'Scenario Manager'
COLS = 'DEFGHIJ'          # FY2027E .. FY2033E on the engine block
MACOL = 'EFGHIJK'         # FY2027E .. FY2033E on 02 Model Assumptions
LEGCOL = 'DEFGHIJ'        # FY2027E .. FY2033E on the legacy research blocks
A = 'C'                   # FY2026A column on the engine block

SCENARIOS = ['Base Case', 'Bull Case', 'Bear Case', 'Stress Case']
SOFF = {'Base Case': 0, 'Bull Case': 1, 'Bear Case': 2, 'Stress Case': 3}

# Base-Case row of each driver in the SCENARIO DRIVER MATRIX on 02 Model Assumptions
DRV = dict(gdp=38, elast=42, fx=46, cpi=50, deliver=54, ratio=58, netexp=62, stock=66,
           real=70, ore=74, coal=78, convinf=82, rmshare=86, oreweight=90, tax=94,
           capexint=98, maintcapex=102, nwcdays=106, exitmult=110, wacc=114, imports=118,
           secondary=122, pricesens=126, normutil=130, maxutil=134, da=138)


def drv(key):
    """Formula fragment for a scenario-matrix driver."""
    return lambda R, x, p: f"='Model Assumptions'!{{ma}}${DRV[key] + R.s}"


# name, label, unit, number format, FY2026A anchor, forecast template
# In templates: R = row resolver (R('name') -> row number), x = this column, p = previous column
M = [
    ('gdp', 'India real GDP growth', '%', '0.0%', "='Macroeconomic Model'!$C$101", drv('gdp')),
    ('elast', 'Steel demand elasticity to GDP', 'x', '0.00"x"', "='Steel Demand Model'!$C$114", drv('elast')),
    ('consg', 'Implied consumption growth', '%', '0.0%', "='Steel Demand Model'!$C$115",
     lambda R, x, p: f"={x}{R('gdp')}*{x}{R('elast')}"),
    ('cons', 'Apparent finished steel consumption', 'Mt', '#,##0.0', "='Steel Demand Model'!$C$116",
     lambda R, x, p: f"={p}{R('cons')}*(1+{x}{R('consg')})"),
    ('netexp', 'Net finished steel exports', 'Mt', '+#,##0.0;-#,##0.0;0.0', "='Steel Supply Model'!$C$90", drv('netexp')),
    ('stock', 'Variation in stock', 'Mt', '+#,##0.0;-#,##0.0;0.0', "='Steel Supply Model'!$C$91", drv('stock')),
    ('reqfin', 'Required finished steel production', 'Mt', '#,##0.0', 'DERIVE',
     lambda R, x, p: f"={x}{R('cons')}+{x}{R('netexp')}+{x}{R('stock')}"),
    ('ratio', 'Crude-to-finished steel ratio', 'x', '0.0000"x"', "='Steel Supply Model'!$C$93", drv('ratio')),
    ('crudeu', 'Unconstrained crude steel production', 'Mt', '#,##0.0', 'DERIVE',
     lambda R, x, p: f"={x}{R('reqfin')}*{x}{R('ratio')}"),
    ('announced', 'Gross announced capacity additions', 'Mtpa', '#,##0.0', None,
     lambda R, x, p: "='Capacity Expansion Tracker'!{ma}$99"),
    ('deliver', 'Pipeline delivery factor', '%', '0%', None, drv('deliver')),
    ('secondary', 'Secondary and unattributed additions', 'Mtpa', '#,##0.0', "='Capacity Forecast'!$C$100", drv('secondary')),
    ('adds', 'Total capacity additions', 'Mtpa', '#,##0.0', "='Capacity Forecast'!$C$100",
     lambda R, x, p: f"={x}{R('announced')}*{x}{R('deliver')}+{x}{R('secondary')}"),
    ('cap', 'Closing crude steel capacity', 'Mtpa', '#,##0.0', "='Capacity Forecast'!$C$101",
     lambda R, x, p: f"={p}{R('cap')}+{x}{R('adds')}"),
    ('maxutil', 'Maximum practical utilisation', '%', '0%', "='Model Assumptions'!$D$32", drv('maxutil')),
    ('crudemax', 'Maximum crude steel production', 'Mt', '#,##0.0', 'DERIVE',
     lambda R, x, p: f"={x}{R('cap')}*{x}{R('maxutil')}"),
    ('crude', 'Crude steel production', 'Mt', '#,##0.0', "='Steel Supply Model'!$C$99",
     lambda R, x, p: f"=MIN({x}{R('crudeu')},{x}{R('crudemax')})"),
    ('util', 'Capacity utilisation', '%', '0.0%', 'DERIVE',
     lambda R, x, p: f"={x}{R('crude')}/{x}{R('cap')}"),
    ('fin', 'Finished steel production', 'Mt', '#,##0.0', "='Steel Supply Model'!$C$101",
     lambda R, x, p: f"={x}{R('crude')}/{x}{R('ratio')}"),
    ('realdrv', 'Blended realisation - scenario driver', 'Rs/t', '#,##0', "='Steel Price Forecast'!$C$97", drv('real')),
    ('normutil', 'Normal capacity utilisation', '%', '0%', "='Model Assumptions'!$D$31", drv('normutil')),
    ('psens', 'Price sensitivity to utilisation', 'x', '0.00"x"', "='Model Assumptions'!$D$30", drv('pricesens')),
    ('padj', 'Price adjustment factor (lagged one year)', 'x', '0.000"x"', '=1',
     lambda R, x, p: f"=1+{x}{R('psens')}*({p}{R('util')}-{x}{R('normutil')})"),
    ('real', 'Blended realisation - effective', 'Rs/t', '#,##0', "='Steel Price Forecast'!$C$103",
     lambda R, x, p: f"={x}{R('realdrv')}*{x}{R('padj')}"),
    ('ore', 'Iron ore 62% Fe CFR China', 'US$/dmt', '#,##0.0', "='Raw Material Forecast'!$C$85", drv('ore')),
    ('coal', 'Premium HCC coking coal FOB Australia', 'US$/t', '#,##0', "='Raw Material Forecast'!$C$86", drv('coal')),
    ('fx', 'USD/INR', 'Rs/US$', '#,##0.0', "='Raw Material Forecast'!$C$87", drv('fx')),
    ('oreinr', 'Iron ore in rupees', 'Rs/dmt', '#,##0', "='Raw Material Forecast'!$C$89",
     lambda R, x, p: f"={x}{R('ore')}*{x}{R('fx')}"),
    ('coalinr', 'Coking coal in rupees', 'Rs/t', '#,##0', "='Raw Material Forecast'!$C$90",
     lambda R, x, p: f"={x}{R('coal')}*{x}{R('fx')}"),
    ('convinf', 'Conversion cost inflation', '%', '0.0%', None, drv('convinf')),
    ('convidx', 'Conversion cost index', 'x', '0.000"x"', '=1',
     lambda R, x, p: f"={p}{R('convidx')}*(1+{x}{R('convinf')})"),
    ('oreweight', 'Iron ore weight in the raw material basket', '%', '0%', "='Model Assumptions'!$D$21", drv('oreweight')),
    ('rmidx', 'Raw material cost index', 'x', '0.000"x"', '=1',
     lambda R, x, p: (f"={x}{R('oreweight')}*({x}{R('oreinr')}/${A}${R('oreinr')})"
                      f"+{x}{R('oreweight')}*({x}{R('coalinr')}/${A}${R('coalinr')})"
                      f"+(1-2*{x}{R('oreweight')})*{x}{R('convidx')}")),
    ('rmshare', 'Raw material share of cash cost', '%', '0%', "='Model Assumptions'!$D$20", drv('rmshare')),
    ('cost', 'Industry cash cost per tonne', 'Rs/t', '#,##0', "='Cost Curve'!$C$80",
     lambda R, x, p: (f"=${A}${R('cost')}*({x}{R('rmshare')}*{x}{R('rmidx')}"
                      f"+(1-{x}{R('rmshare')})*{x}{R('convidx')})")),
    ('ebitdat', 'Industry EBITDA per tonne', 'Rs/t', '#,##0', 'DERIVE',
     lambda R, x, p: f"={x}{R('real')}-{x}{R('cost')}"),
    ('rev', 'Industry revenue', 'Rs cr', '#,##0', "='Revenue Forecast'!$C$96",
     lambda R, x, p: f"={x}{R('fin')}*{x}{R('real')}/10"),
    ('ebitda', 'Industry EBITDA', 'Rs cr', '#,##0', 'DERIVE',
     lambda R, x, p: f"={x}{R('fin')}*{x}{R('ebitdat')}/10"),
    ('margin', 'Industry EBITDA margin', '%', '0.0%', 'DERIVE',
     lambda R, x, p: f"={x}{R('ebitdat')}/{x}{R('real')}"),
    ('dapct', 'Depreciation and amortisation', '% of revenue', '0.00%', "='Cash Flow Model'!$C$98", drv('da')),
    ('da', 'Depreciation and amortisation', 'Rs cr', '#,##0', 'DERIVE',
     lambda R, x, p: f"={x}{R('rev')}*{x}{R('dapct')}"),
    ('ebit', 'EBIT', 'Rs cr', '#,##0;(#,##0)', 'DERIVE',
     lambda R, x, p: f"={x}{R('ebitda')}-{x}{R('da')}"),
    ('tax', 'Effective tax rate', '%', '0.00%', "='Cash Flow Model'!$C$101", drv('tax')),
    ('cashtax', 'Cash tax', 'Rs cr', '#,##0', 'DERIVE',
     lambda R, x, p: f"=MAX(0,{x}{R('ebit')})*{x}{R('tax')}"),
    ('capexint', 'Capex intensity', 'Rs/t of capacity', '#,##0', "='Model Assumptions'!$D$23", drv('capexint')),
    ('gcapex', 'Growth capex', 'Rs cr', '#,##0', None,
     lambda R, x, p: f"={x}{R('adds')}*{x}{R('capexint')}/10"),
    ('mcapexp', 'Maintenance capex', '% of revenue', '0.0%', "='Model Assumptions'!$D$24", drv('maintcapex')),
    ('mcapex', 'Maintenance capex', 'Rs cr', '#,##0', None,
     lambda R, x, p: f"={x}{R('rev')}*{x}{R('mcapexp')}"),
    ('capex', 'Total capex', 'Rs cr', '#,##0', None,
     lambda R, x, p: f"={x}{R('gcapex')}+{x}{R('mcapex')}"),
    ('nwcd', 'Net working capital', 'days of revenue', '#,##0', "='Working Capital Model'!$C$71", drv('nwcdays')),
    ('nwc', 'Net working capital', 'Rs cr', '#,##0', "='Working Capital Model'!$C$72",
     lambda R, x, p: f"={x}{R('rev')}*{x}{R('nwcd')}/365"),
    ('dnwc', 'Change in net working capital', 'Rs cr', '#,##0;(#,##0)', None,
     lambda R, x, p: f"={x}{R('nwc')}-{p}{R('nwc')}"),
    ('ufcf', 'Unlevered free cash flow', 'Rs cr', '#,##0;(#,##0)', None,
     lambda R, x, p: f"={x}{R('ebitda')}-{x}{R('cashtax')}-{x}{R('capex')}-{x}{R('dnwc')}"),
    ('exitmult', 'Exit EV/EBITDA multiple', 'x', '0.0"x"', None, drv('exitmult')),
    ('wacc', 'WACC (nominal, INR)', '%', '0.0%', None, drv('wacc')),
    ('impdrv', 'Finished steel imports - driver', 'Mt', '#,##0.0', "='Trade Model'!$C$74", drv('imports')),
    ('short', 'Shortfall routed from the capacity ceiling', 'Mt', '#,##0.0', '=0',
     lambda R, x, p: f"=MAX(0,{x}{R('reqfin')}-{x}{R('fin')})"),
    ('imp', 'Total finished steel imports', 'Mt', '#,##0.0', 'DERIVE',
     lambda R, x, p: f"={x}{R('impdrv')}+{x}{R('short')}"),
    ('exp', 'Finished steel exports', 'Mt', '#,##0.0', "='Trade Model'!$C$78",
     lambda R, x, p: f"={x}{R('impdrv')}+{x}{R('netexp')}"),
    ('imppen', 'Import penetration', '% of consumption', '0.0%', 'DERIVE',
     lambda R, x, p: f"={x}{R('imp')}/{x}{R('cons')}"),
    ('percap', 'Per capita finished steel consumption', 'kg', '#,##0.0', "='Steel Demand Model'!$C$121",
     lambda R, x, p: f"={x}{R('cons')}/'Steel Demand Model'!{{leg}}$119"),
    ('ev', 'Enterprise value at the exit multiple (terminal-year proxy)', 'Rs cr', '#,##0', None,
     lambda R, x, p: f"={x}{R('ebitda')}*{x}{R('exitmult')}"),
    ('nopat', 'NOPAT', 'Rs cr', '#,##0;(#,##0)', None,
     lambda R, x, p: f"={x}{R('ebit')}*(1-{x}{R('tax')})"),
    ('ic', 'Invested capital (proxy)', 'Rs cr', '#,##0', None,
     lambda R, x, p: f"=220.4*{x}{R('capexint')}/10+{x}{R('nwc')}"),
    ('roic', 'ROIC (indicative)', '%', '0.0%', None,
     lambda R, x, p: f"=IFERROR({x}{R('nopat')}/{x}{R('ic')},\"\")"),

    # ---- cycle scorecard, one indicator per row.
    # These exist so that 18 Industry Cycle Model rows 70-71 can be a single SUMPRODUCT
    # against a range instead of one 550-character formula per scenario. Each indicator
    # is scored 0-100 on the same band as the live scorecard on that sheet, so the four
    # scenarios and the live model are measured on exactly the same ruler.
    ('cyc1', 'Cycle indicator 1 - real GDP growth (4% to 9%)', 'score', '0.0', 'DERIVE',
     lambda R, x, p: f"=MIN(100,MAX(0,({x}{R('gdp')}-4%)/5%*100))"),
    ('cyc2', 'Cycle indicator 2 - consumption growth (2% to 12%)', 'score', '0.0', 'DERIVE',
     lambda R, x, p: f"=MIN(100,MAX(0,({x}{R('consg')}-2%)/10%*100))"),
    ('cyc3', 'Cycle indicator 3 - capacity utilisation (65% to 92%)', 'score', '0.0', 'DERIVE',
     lambda R, x, p: f"=MIN(100,MAX(0,({x}{R('util')}-65%)/27%*100))"),
    ('cyc4', 'Cycle indicator 4 - spare capacity, scored inversely (0% to 30%)', 'score', '0.0', 'DERIVE',
     lambda R, x, p: (f"=MIN(100,MAX(0,100-({x}{R('cap')}-{x}{R('crudeu')})"
                      f"/{x}{R('cap')}/30%*100))")),
    ('cyc5', 'Cycle indicator 5 - blended realisation (Rs 50,000 to Rs 75,000/t)', 'score', '0.0', 'DERIVE',
     lambda R, x, p: f"=MIN(100,MAX(0,({x}{R('real')}-50000)/25000*100))"),
    ('cyc6', 'Cycle indicator 6 - EBITDA margin (-10% to +26%)', 'score', '0.0', 'DERIVE',
     lambda R, x, p: f"=MIN(100,MAX(0,({x}{R('margin')}+10%)/36%*100))"),
    ('cyc7', 'Cycle indicator 7 - capacity additions (0 to 25 Mtpa)', 'score', '0.0', 'DERIVE',
     lambda R, x, p: f"=MIN(100,MAX(0,{x}{R('adds')}/25*100))"),
    ('cyc8', 'Cycle indicator 8 - variation in stock, centred on 50', 'score', '0.0', 'DERIVE',
     lambda R, x, p: f"=MIN(100,MAX(0,50-{x}{R('stock')}/{x}{R('cons')}*1000))"),
    ('cycscore', 'CYCLE SCORE - weighted', 'score 0-100', '0.0', 'DERIVE',
     lambda R, x, p: (f"=SUMPRODUCT('Industry Cycle Model'!$B$19:$B$26,"
                      f"{x}{R('cyc1')}:{x}{R('cyc8')})")),
]

NAMES = [m[0] for m in M]
IX = {name: i for i, name in enumerate(NAMES)}
NMET = len(M)

NAVY = 'FF1F3864'
THIN = Side(style='thin', color='FFD9D9D9')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


class _R:
    """Row resolver passed into the formula templates."""

    def __init__(self, base, s):
        self.base = base
        self.s = s

    def __call__(self, name):
        return self.base + IX[name]


class Engine:
    """Addressing helper handed to the sheet fillers."""

    def __init__(self, base_rows):
        self.base = base_rows

    def row(self, scenario, name):
        return self.base[scenario] + IX[name]

    def cell(self, scenario, name, year_index, abs_=True):
        """year_index 0 = FY2026A, 1..7 = FY2027E..FY2033E."""
        col = A if year_index == 0 else COLS[year_index - 1]
        d = '$' if abs_ else ''
        return f"'{SHEET}'!{d}{col}{d}{self.row(scenario, name)}"

    def ref(self, scenario, name, year_index, abs_=True):
        return '=' + self.cell(scenario, name, year_index, abs_)

    def rng(self, scenario, name, y0=1, y1=7):
        r = self.row(scenario, name)
        return f"'{SHEET}'!${COLS[y0 - 1]}${r}:${COLS[y1 - 1]}${r}"


def build(wb, start_row=90):
    ws = wb[SHEET]
    r = start_row

    t = ws.cell(r, 1, 'SCENARIO ENGINE - FULL SEVEN-YEAR PATH FOR ALL FOUR SCENARIOS')
    t.font = Font(name='Calibri', sz=12, b=True, color='FFFFFFFF')
    for c in range(1, 12):
        ws.cell(r, c).fill = PatternFill('solid', fgColor=NAVY)
    ws.row_dimensions[r].height = 20

    n = ws.cell(r + 1, 1, (
        'The live model resolves one scenario at a time. This block rebuilds the model line for line for all four '
        'scenarios directly off the SCENARIO DRIVER MATRIX on 02 Model Assumptions rows 38 to 141, so that every '
        'four-scenario table in the workbook is a live link rather than a hard-coded number. It is also the '
        "model's own regression test: for whichever scenario is active the block reproduces the live model exactly, "
        'and its FY2033E column reproduces the independently computed tie-out values on rows 69 to 79 above. '
        'Nothing here is an input - every cell is a formula.'))
    n.font = Font(name='Calibri', sz=9, i=True, color='FF595959')
    n.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=11)
    ws.row_dimensions[r + 1].height = 42

    hdr = ['Line item', 'Unit', 'FY2026A', 'FY2027E', 'FY2028E', 'FY2029E', 'FY2030E',
           'FY2031E', 'FY2032E', 'FY2033E', 'FY26-FY33 CAGR']
    hr = r + 2
    for i, h in enumerate(hdr, start=1):
        c = ws.cell(hr, i, h)
        c.font = Font(name='Calibri', sz=9, b=True, color='FFFFFFFF')
        c.fill = PatternFill('solid', fgColor='FF2E5C8A')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BOX
    ws.row_dimensions[hr].height = 26

    base_rows = {}
    cur = hr + 1
    for name in SCENARIOS:
        s = SOFF[name]
        bn = ws.cell(cur, 1, name.upper())
        bn.font = Font(name='Calibri', sz=10, b=True, color='FF1F3864')
        for c in range(1, 12):
            ws.cell(cur, c).fill = PatternFill('solid', fgColor='FFDCE6F1')
            ws.cell(cur, c).border = BOX
        ws.row_dimensions[cur].height = 16
        cur += 1
        b = cur
        base_rows[name] = b
        R = _R(b, s)
        for i, (key, label, unit, fmt, anchor, tmpl) in enumerate(M):
            row = b + i
            lc = ws.cell(row, 1, label)
            lc.font = Font(name='Calibri', sz=9, color='FF404040')
            lc.alignment = Alignment(vertical='center', indent=1)
            lc.border = BOX
            uc = ws.cell(row, 2, unit)
            uc.font = Font(name='Calibri', sz=8, color='FF808080')
            uc.alignment = Alignment(horizontal='center', vertical='center')
            uc.border = BOX
            ac = ws.cell(row, 3)
            if anchor == 'DERIVE':
                ac.value = tmpl(R, A, A).replace('{ma}', 'D').replace('{leg}', 'C')
                ac.font = Font(name='Calibri', sz=9, color='FF000000')
            elif anchor is not None:
                ac.value = anchor
                ac.font = Font(name='Calibri', sz=9, color='FF008000')
            if ac.value is not None:
                ac.number_format = fmt
            ac.alignment = Alignment(horizontal='right', vertical='center')
            ac.border = BOX
            for j, col in enumerate(COLS):
                prev = A if j == 0 else COLS[j - 1]
                cc = ws.cell(row, 4 + j)
                cc.value = tmpl(R, col, prev).replace('{ma}', MACOL[j]).replace('{leg}', LEGCOL[j])
                cc.number_format = fmt
                cc.font = Font(name='Calibri', sz=9,
                               color='FF008000' if 'Model Assumptions' in str(cc.value) else 'FF000000')
                cc.alignment = Alignment(horizontal='right', vertical='center')
                cc.border = BOX
            gc = ws.cell(row, 11)
            if fmt in ('#,##0.0', '#,##0', '#,##0;(#,##0)') and ac.value is not None:
                gc.value = (f'=IFERROR(IF(AND(ISNUMBER(C{row}),ISNUMBER(J{row}),C{row}>0),'
                            f'(J{row}/C{row})^(1/7)-1,""),"")')
                gc.number_format = '0.0%'
            gc.font = Font(name='Calibri', sz=9, color='FF000000')
            gc.alignment = Alignment(horizontal='right', vertical='center')
            gc.border = BOX
            ws.row_dimensions[row].height = 14
        cur = b + NMET + 1

    for letter, w in (('J', 11), ('K', 13)):
        if letter not in ws.column_dimensions:
            ws.column_dimensions[letter].width = w

    return Engine(base_rows), cur
