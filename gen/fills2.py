"""Fills for 05 Steel Supply Model, 06 Capacity Forecast, 07 Capacity Expansion Tracker,
08 Capacity Utilisation, 09 Steel Price Forecast and 10 Raw Material Forecast."""
from openpyxl.styles import Font, Alignment

from .style import (SheetWriter, PCT0, PCT1, PCT2, NUM0, NUM1, NUM2, MT, MTS, RS, CR,
                    X1, X2, USD, TEXT)
from .support import na_row
from .spec import NA

LEG = 'CDEFGHIJ'
FLEG = 'DEFGHIJ'
NEW8 = 'BCDEFGHI'
NEW7 = 'BCDEFGH'
MACOL = 'EFGHIJK'
SCEN = ['Base Case', 'Bull Case', 'Bear Case', 'Stress Case']
TRK = "'Capacity Expansion Tracker'"


def _sup(item, value, unit, method, formula, primary, secondary, assumption, reasoning,
         cross, conf, linked, freq, last, comments, numfmt=None):
    return dict(item=item, value=value, unit=unit, method=method, formula=formula, primary=primary,
                secondary=secondary, assumption=assumption, reasoning=reasoning, cross=cross,
                conf=conf, linked=linked, freq=freq, last=last, comments=comments, numfmt=numfmt)


# ================================================================== 05 Steel Supply Model
def supply(wb, audit, eng):
    ws = wb['Steel Supply Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}9', f'=${lc}$99', MT)
        w.link(f'{nc}10', f'=${lc}$101', MT)
        w.link(f'{nc}11', f"='Capacity Forecast'!${lc}$101", MT)
        w.link(f'{nc}12', f"='Capacity Utilisation'!${lc}$67", PCT1)
        w.f(f'{nc}13', f"={lc}101+'Trade Model'!{lc}76-'Trade Model'!{lc}78", MT)
        w.link(f'{nc}14', f"='Trade Model'!${lc}$83", PCT1)
        w.link(f'{nc}15', f"='Trade Model'!${lc}$84", PCT1)
        w.f(f'{nc}16', f"={nc}13-'Steel Demand Model'!{lc}116", MTS)

    # ---- historical supply rows 21-26
    hist = [(21, 'FY2021', 103.545, 96.20), (22, 'FY2022', None, 113.60),
            (23, 'FY2023', 127.20, 123.20), (24, 'FY2024', 144.30, 139.15),
            (25, 'FY2025', 152.18, 146.69), (26, 'FY2026', 168.42, 161.74)]
    for r, fy, crude, fin in hist:
        if crude is None:
            w.na(f'B{r}', NA['crude_fy22'])
        else:
            w.db(f'B{r}', crude, MT, ref="='[Master Industry Database.xlsx]Production Volume'!$D$20:$H$20")
        w.db(f'C{r}', fin, MT, ref="='[Master Industry Database.xlsx]Production Volume'!$D$21:$H$21")
        if r == 21:
            w.na(f'D{r}', 'FY2020 finished steel production is outside the five-year window imported into '
                          'this table, so a FY2021 growth rate cannot be computed here. The Master Industry '
                          'Database carries FY2015 onwards if a longer series is required.')
        else:
            w.f(f'D{r}', f'=IFERROR(C{r}/C{r - 1}-1,"")', PCT1)
    nas.append(na_row('FY2022 crude steel production', 'B22', NA['crude_fy22'], unit='Mt'))
    rows.append(_sup(
        'Historical production FY2021-FY2026', '=C26', 'Mt',
        'Master Industry Database import (RED - to be reconnected as an external link)',
        "='[Master Industry Database.xlsx]Production Volume'!$D$20:$H$21",
        'Joint Plant Committee, published by the Ministry of Steel',
        'World Steel Association for the calendar-year cross-check (India CY2025 164.89 Mt)',
        'Growth is computed on FINISHED steel production, which is the complete series; the crude series '
        'has a gap at FY2022.',
        'Finished steel production compounded at 9.2% over FY2022-FY2026 while capacity compounded faster, '
        'which is precisely why utilisation fell from 80.4% in FY2024 to 76.4% in FY2026.',
        'FY2026 crude of 168.42 Mt and finished of 161.74 Mt reconcile to a 1.041x ratio on the June '
        'vintage and 1.0465x on the April vintage used by the model',
        'High', 'Master Industry Database', 'Monthly, on each Ministry of Steel report',
        'FY2026 (provisional)', 'Six years of history were blank.', MT))

    # ---- supply build-up rows 31-41
    w.db('B31', 152.18, MT, ref="='[Master Industry Database.xlsx]Production Volume'!$G$20")
    for i, nc in enumerate(NEW8):
        lc = LEG[i]
        if i:
            w.f(f'{nc}31', f'={NEW8[i - 1]}35', MT)
        w.link(f'{nc}32', f"='Capacity Forecast'!${lc}$100", MT)
        w.link(f'{nc}33', f"='Capacity Forecast'!${lc}$101", MT)
        w.link(f'{nc}34', f"='Capacity Utilisation'!${lc}$67", PCT1)
        w.link(f'{nc}35', f'=${lc}$99', MT)
        w.link(f'{nc}36', f'=${lc}$93', '0.0000"x"')
        w.link(f'{nc}37', f'=${lc}$101', MT)
        w.link(f'{nc}38', f"='Trade Model'!${lc}$76", MT)
        w.link(f'{nc}39', f"='Trade Model'!${lc}$78", MT)
        w.link(f'{nc}40', f'=${lc}$91', MTS)
        w.f(f'{nc}41', f'={nc}37+{nc}38-{nc}39', MT)

    # ---- supply / demand reconciliation rows 46-51 (48 and 52 already carry formulas)
    for i, nc in enumerate(NEW8):
        lc = LEG[i]
        w.f(f'{nc}46', f'={nc}37', MT)
        w.f(f'{nc}47', f'={nc}38', MT)
        w.link(f'{nc}49', f"='Steel Demand Model'!${lc}$116", MT)
        w.f(f'{nc}50', f'={nc}39', MT)
        w.f(f'{nc}51', f'={nc}40', MTS)
    rows.append(_sup(
        'Supply / demand balance check', '=B52', 'Mt',
        'Arithmetic reconciliation, live in every year',
        '=B48-B49-B50-B51 where B48 = domestic supply + imports',
        'Joint Plant Committee supply balance identity',
        '25 Audit Checks A07 and A08 test the same identity independently',
        'The Joint Plant Committee identity: finished production = apparent consumption + net exports + '
        'variation in stock.',
        'THE MOST IMPORTANT RECONCILIATION IN THE WORKBOOK. It must return zero in every year and in every '
        'scenario. FY2026A: 163.74 + 0.078 - 2.88 = 160.94, which is exactly the reported figure.',
        'Independently duplicated by checks A07 and A08 on 25 Audit Checks',
        'High', 'Steel Demand Model, Trade Model', 'Live', 'FY2026A', 'Was blank in every column.', '0.000'))

    # ---- capacity growth rows 57-64
    w.db('B57', 200.33, MT, ref="='[Master Industry Database.xlsx]Production Capacity'!$E$20")
    for i, r in enumerate(range(57, 65)):
        lc = LEG[i]
        if i:
            w.f(f'B{r}', f'=D{r - 1}', MT)
        w.link(f'C{r}', f"='Capacity Forecast'!${lc}$100", MT)
        w.link(f'D{r}', f"='Capacity Forecast'!${lc}$101", MT)
        w.f(f'E{r}', f'=IFERROR(D{r}/B{r}-1,"")', PCT1)

    # ---- import / export dependency rows 69-73
    for i, nc in enumerate(NEW8):
        lc = LEG[i]
        w.link(f'{nc}69', f"='Trade Model'!${lc}$76", MT)
        w.link(f'{nc}70', f"='Trade Model'!${lc}$78", MT)
        w.f(f'{nc}71', f'={nc}69-{nc}70', MTS)
        w.link(f'{nc}72', f"='Trade Model'!${lc}$83", PCT1)
        w.link(f'{nc}73', f"='Trade Model'!${lc}$84", PCT1)

    # ---- supply risks rows 78-83
    risk_text = {
        78: 'Iron ore availability is the least binding of the six. India is a net exporter of iron ore and '
            'the model prices ore off the World Bank CFR China benchmark, so a domestic availability shock '
            'would show up as a price event, which the Stress Case already carries at US$115/dmt.',
        79: 'Coal supply is the most material. India imports substantially all of its coking coal and the '
            'Master Industry Database records no fiscal-year average price (Conflict C01). The Stress Case '
            'takes premium HCC to US$300/t, which is below the October-2023 observed peak of US$354/t.',
        80: 'Logistics is not separately modelled. It sits inside conversion cost, which is indexed to RBI '
            'CPI. A freight shock would therefore be understated by this model.',
        81: 'Power availability is not separately modelled and sits inside conversion cost. India is the '
            "world's largest DRI producer at 60.3 Mt, and the DRI-EAF route is materially more exposed to "
            'power and thermal coal than the BF-BOF route.',
        82: 'Environmental regulation is quantified only through the CBAM calculation on 20 ESG Model and '
            'the green steel taxonomy thresholds. A domestic carbon price is not modelled because none is '
            'in force.',
        83: 'Plant shutdowns are not modelled. The maximum practical utilisation ceiling of 92% is the only '
            'allowance for maintenance and relining, and it is an unsourced indicative parameter.',
    }
    for r in range(78, 84):
        w.narow('BCDEFGH', r, NA['supply_risk'])
        c = ws.cell(r, 9, risk_text[r])
        c.font = Font(name='Calibri', sz=8, color='FF404040')
        c.alignment = Alignment(wrap_text=True, vertical='top')
    nas.append(na_row('Supply risk scores by year', 'B78:H83', NA['supply_risk'],
                      linked='This sheet'))
    return rows, nas


# ==================================================================== 06 Capacity Forecast
COMPANY_ROWS = {19: 'Tata Steel', 20: 'JSW Steel', 21: 'SAIL', 22: 'Jindal Steel',
                23: 'Jindal Stainless', 24: 'AM/NS India', 25: 'Shyam Metalics',
                26: 'APL Apollo', 27: 'GPIL'}
# India-basis crude steel capacity at FY2026 year-end, as already carried on 23 Comparable
# Valuation column K (which is itself sourced from the Master Industry Database).
CAP26 = {19: "='Comparable Valuation'!$K$108", 20: "='Comparable Valuation'!$K$109",
         21: "='Comparable Valuation'!$K$110", 22: "='Comparable Valuation'!$K$111",
         23: "='Comparable Valuation'!$K$112", 24: "='Comparable Valuation'!$K$113"}


def capacity(wb, audit, eng):
    ws = wb['Capacity Forecast']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- executive summary rows 9-14
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}9', f'=${lc}$101', MT)
        w.link(f'{nc}10', f'=${lc}$100', MT)
        w.f(f'{nc}11', f'=IFERROR({lc}100/({lc}101-{lc}100),"")', PCT1)
        w.link(f'{nc}12', f"='Capacity Utilisation'!${lc}$67", PCT1)
        w.f(f'{nc}13', f"={lc}101-'Steel Supply Model'!{lc}99", MT)
        w.f(f'{nc}14', f"={lc}101-'Steel Supply Model'!{lc}94", MT)

    # ---- company-wise capacity rows 19-28, total row 29 already sums
    for r, name in COMPANY_ROWS.items():
        if r in CAP26:
            w.link(f'B{r}', CAP26[r], MT)
            for i, nc in enumerate('CDEFGHI'):
                prev = NEW8[i]
                ma = MACOL[i]
                fy = 2027 + i
                w.f(f'{nc}{r}',
                    f'={prev}{r}+SUMIFS({TRK}!$F$77:$F$96,{TRK}!$B$77:$B$96,$A{r},'
                    f'{TRK}!$I$77:$I$96,"FY{fy}",{TRK}!$L$77:$L$96,"Yes")'
                    f"*'Model Assumptions'!{ma}$12", MT)
        else:
            w.narow(NEW8, r, (
                f'{name} does not disclose crude steel capacity. The Master Industry Database records '
                '"Data Not Publicly Available" for this name, and for APL Apollo and Godawari Power the '
                'relevant capacity is tube and downstream capacity, not crude steel. Any capacity these '
                'names do operate is inside the Others residual on row 28, so the industry total still '
                'reconciles to 220.4 Mtpa.'))
    for i, nc in enumerate(NEW8):
        lc = LEG[i]
        w.f(f'{nc}28', f'=${lc}$101-SUM({nc}19:{nc}27)', MT)
    nas.append(na_row('Company capacity - Shyam Metalics, APL Apollo, Godawari Power',
                      'B25:I27',
                      'Crude steel capacity is not disclosed by these names. APL Apollo is a tube converter '
                      'and Godawari Power does not publish a crude steel figure. Their capacity is carried '
                      'inside the Others residual on row 28 so that the industry total still reconciles.',
                      unit='Mtpa'))
    rows.append(_sup(
        'Company-wise capacity, India basis', '=B29', 'Mtpa',
        'FY2026A linked to the sourced company capacities; forecast years roll forward with a live SUMIFS '
        'against the dated project tracker, scaled by the pipeline delivery factor',
        '=B19+SUMIFS(tracker capacity, company, FY, included)*delivery factor',
        'Company results releases and investor presentations; Ministry of Steel Rajya Sabha answer for SAIL',
        'Master Industry Database Production Capacity section B; 07 Capacity Expansion Tracker',
        'Others (row 28) is a RESIDUAL, not an estimate: it is the industry total less the named producers.',
        'The Master Industry Database warns explicitly that section B must not be summed, because it mixes '
        'global group capacity, India entity capacity and stainless melt capacity. This table therefore '
        'uses India-basis crude steel capacity only, taken from the same figures used for the EV per tonne '
        'calculation on 23 Comparable Valuation, and plugs the difference into Others. Others contains '
        'RINL, NMDC Steel, ESL, the JSW-POSCO joint venture and the entire secondary sector, which the '
        'Joint Plant Committee shows produced 73.02 Mt in FY2026.',
        'Row 29 must equal research block row 101 in every year, because Others is defined as the residual',
        'High for the named producers, Medium for the residual',
        'Comparable Valuation, Capacity Expansion Tracker, Model Assumptions',
        'Quarterly, on results and project announcements', 'FY2026A', 'Was entirely blank.', MT))

    # ---- capacity expansion pipeline rows 36-46 (rows 34-35 are the owner's own entries)
    tracker_rows = [78, 79, 80, 81, 83, 84, 85, 86, 88, 89, 95]
    for j, tr in enumerate(tracker_rows):
        r = 36 + j
        w.link(f'A{r}', f'={TRK}!$C${tr}', TEXT)
        w.link(f'B{r}', f'={TRK}!$B${tr}', TEXT)
        w.link(f'C{r}', f'={TRK}!$E${tr}', TEXT)
        w.f(f'D{r}', f'=IF({TRK}!$I${tr}="","",'
                     f'"FY"&RIGHT({TRK}!$I${tr},2))', TEXT)
        w.link(f'E{r}', f'={TRK}!$I${tr}', TEXT)
        w.link(f'F{r}', f'={TRK}!$F${tr}', NUM2)
        w.link(f'G{r}', f'={TRK}!$H${tr}', TEXT)
    rows.append(_sup(
        'Capacity expansion pipeline', f'=SUM(F36:F46)', 'Mtpa',
        'Cross-sheet reference to the dated project register on 07 Capacity Expansion Tracker',
        "='Capacity Expansion Tracker'!$F$78 etc.",
        'Company results releases, investor presentations and exchange filings; Ministry of Steel '
        'Parliamentary answers for SAIL',
        'Master Industry Database sources S08 to S43',
        'Only projects flagged "Included in forecast = Yes" on the tracker appear here.',
        'The eleven projects listed are every dated, India, crude steel addition inside the FY2027-FY2033 '
        'horizon. Acquisitions, downstream processing, projects already inside the FY2026 base and projects '
        'beyond FY2033 are excluded on the tracker and therefore excluded here.',
        'The sum must equal 53.62 Mtpa, which is the total on tracker row 101',
        'High', 'Capacity Expansion Tracker', 'Quarterly', 'FY2026A',
        'Rows 34 and 35 are the model owner\'s own entries and were left untouched. Rows 36 to 46 are live '
        'links, so the pipeline can never disagree with the tracker.', NUM2))

    # ---- capacity bridge rows 51-57
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        fy = 2027 + i
        w.link(f'{nc}51', f'=${lc}$95', MT)
        w.f(f'{nc}52', f'=SUMIFS({TRK}!$F$77:$F$96,{TRK}!$E$77:$E$96,"Brownfield*",'
                       f'{TRK}!$I$77:$I$96,"FY{fy}",{TRK}!$L$77:$L$96,"Yes")*${lc}$97', MT)
        w.f(f'{nc}53', f'=${lc}$96*${lc}$97-{nc}52', MT)
        w.link(f'{nc}54', f'=${lc}$99', MT)
        w.inp(f'{nc}55', 0, MT)
        w.f(f'{nc}56', f'={nc}52+{nc}53+{nc}54-{nc}55', MT)
        w.f(f'{nc}57', f'={nc}51+{nc}56', MT)
    rows.append(_sup(
        'Capacity bridge', '=H57', 'Mtpa',
        'Computed in-cell; opening plus brownfield plus greenfield plus the secondary reconciliation item '
        'less closures',
        '=B51+B52+B53+B54-B55, closing = opening + net addition',
        'Dated project tracker plus the secondary and unattributed reconciliation item, driver D22',
        'Research block rows 95 to 101, which the bridge must reproduce exactly',
        'PERMANENT CLOSURES ARE SET TO ZERO. That is a stated model limitation, not an oversight: the '
        'model does not idle or close capacity in any scenario. Check A02 on 25 Audit Checks tests that '
        'capacity is monotonically non-decreasing.',
        'The DEBOTTLENECKING line carries driver D22, the secondary and unattributed additions of about 11 '
        'Mtpa a year. This is the essential reconciliation item: the dated tracker averages only 7.7 Mtpa '
        'a year against a demonstrated national delivery of roughly 20 Mtpa a year in FY2025 and FY2026. '
        'The balance is induction furnace, small EAF and debottlenecking capacity that is never announced '
        'as a project, which is why it is classified here rather than as greenfield. GREENFIELD is '
        'computed as total announced less brownfield, so it also absorbs the two undated aspirational '
        "programmes - Tata Steel's roadmap to 40 Mtpa and SAIL's ~35 Mt ambition.",
        'Row 57 must equal research block row 101 in every year and in every scenario',
        'Medium', 'Capacity Expansion Tracker, Model Assumptions', 'Quarterly', 'FY2026A',
        'Was entirely blank. Closures are blue because zero is a deliberate input, not a formula.', MT))

    # ---- capacity vs demand rows 62-68 and utilisation outlook rows 73-79
    for i in range(7):
        lc = FLEG[i]
        r1, r2 = 62 + i, 73 + i
        w.link(f'B{r1}', f'=${lc}$101', MT)
        w.link(f'C{r1}', f"='Steel Supply Model'!${lc}$94", MT)
        w.f(f'D{r1}', f'=B{r1}-C{r1}', MT)
        w.f(f'E{r1}', f'=IFERROR(B{r1}/C{r1},"")', X2)
        w.link(f'B{r2}', f'=${lc}$101', MT)
        w.link(f'C{r2}', f"='Steel Supply Model'!${lc}$99", MT)
        w.f(f'D{r2}', f'=IFERROR(C{r2}/B{r2},"")', PCT1)

    # ---- regional capacity rows 84-88
    for r in range(84, 89):
        w.narow('BC', r, NA['regional_cap'])
    nas.append(na_row('Regional capacity split', 'B84:C88', NA['regional_cap'], unit='Mtpa'))
    return rows, nas


# ========================================================= 07 Capacity Expansion Tracker
def tracker(wb, audit):
    ws = wb['Capacity Expansion Tracker']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    w.f('B9', '=COUNTA($A$77:$A$96)', NUM0)
    w.f('B10', '=SUM($F$77:$F$96)', NUM2)
    w.f('B11', '=SUMIFS($F$77:$F$96,$H$77:$H$96,"Under*")', NUM2)
    w.f('B12', '=SUMIFS($F$77:$F$96,$H$77:$H$96,"Commissioned*")'
               '+SUMIFS($F$77:$F$96,$H$77:$H$96,"Completed*")', NUM2)
    w.f('B13', "=$D$101*'Model Assumptions'!$E$12", NUM2)
    w.link('B14', "='Model Assumptions'!$E$12", PCT0)
    w.f('B15', '=IFERROR(SUM($F$77:$F$96)/COUNT($F$77:$F$96),"")', NUM2)
    rows.append(_sup(
        'Tracker executive summary', '=B13', 'Mtpa',
        'Live COUNTA, SUM and SUMIFS across the dated project register below',
        "=$D$101*'Model Assumptions'!$E$12",
        'The register itself, sourced project by project in column O',
        '06 Capacity Forecast rows 96 to 100',
        'The "average completion probability" is driver D05, the pipeline delivery factor of 85% in the '
        'base case. It is applied at portfolio level rather than project by project.',
        'A project-by-project probability would imply a precision the disclosures do not support: eleven of '
        'the twenty entries are described only as announced, proposed or aspirational. The delivery factor '
        'is instead calibrated top-down against a demonstrated national delivery of roughly 20 Mtpa a year '
        'in FY2025 and FY2026.',
        'Weighted expected capacity of 45.6 Mtpa against 53.62 Mtpa of gross announced additions',
        'Medium', 'Model Assumptions, Capacity Forecast', 'Quarterly, on results and announcements',
        'FY2026A', 'Was entirely blank.', NUM2))

    # ---- master project register rows 20-39 mirror the dated register at rows 77-96
    for j in range(20):
        r, tr = 20 + j, 77 + j
        w.link(f'B{r}', f'=$B${tr}', TEXT)
        w.link(f'C{r}', f'=$C${tr}', TEXT)
        w.link(f'D{r}', f'=$D${tr}', TEXT)
        w.f(f'E{r}', f'=IF(ISNUMBER(SEARCH("EAF",$E${tr})),"EAF",'
                     f'IF(ISNUMBER(SEARCH("BF",$E${tr})),"BF-BOF",'
                     f'IF(ISNUMBER(SEARCH("integrated",$E${tr})),"BF-BOF (integrated)",'
                     f'IF(ISNUMBER(SEARCH("melt",$G${tr})),"Stainless melt","Not specified"))))', TEXT)
        w.link(f'F{r}', f'=$E${tr}', TEXT)
        w.na(f'G{r}', 'Existing capacity at the specific plant or asset is not disclosed for most projects. '
                      'Where it is, it appears in the project description in column C - for example '
                      '"Vijayanagar BF-3 expansion 3.0 to 4.5 Mtpa" and "Hazira expansion 9 to 15 Mtpa".')
        w.link(f'H{r}', f'=$F${tr}', NUM2)
        w.na(f'I{r}', 'Total capacity on completion cannot be computed without the plant-level existing '
                      'capacity in column G, which is not disclosed.')
        w.na(f'J{r}', 'A clean announcement DATE is not disclosed for most projects. Where a date is '
                      'available it is embedded in the status text carried in column M - for example '
                      '"Announced 23-Mar-2026" and "Announced Apr-2026". Parsing a date out of free text '
                      'would be unreliable.')
        w.link(f'K{r}', f'=$I${tr}', TEXT)
        w.f(f'L{r}', f'=IF($L${tr}="Yes",\'Model Assumptions\'!$E$12,0)', PCT0)
        w.link(f'M{r}', f'=$H${tr}', TEXT)
        w.link(f'N{r}', f'=$J${tr}', CR)
        w.link(f'O{r}', f'=$M${tr}', TEXT)
    nas.append(na_row('Per-project existing capacity, total capacity on completion and announcement date',
                      'G20:G39, I20:I39, J20:J39',
                      'Not disclosed at project level. Where the information exists it is inside the '
                      'project description and status text, which are both carried live in this table.'))
    rows.append(_sup(
        'Master project register', '=SUM(H20:H39)', 'Mtpa',
        'Cross-sheet reference within this sheet to the dated register at rows 77 to 96; the steelmaking '
        'ROUTE is derived from the project type and capacity definition by formula, not asserted',
        '=IF(ISNUMBER(SEARCH("EAF",$E$77)),"EAF",IF(ISNUMBER(SEARCH("BF",$E$77)),"BF-BOF",...))',
        'Twenty projects, each cited individually in column O against a numbered source',
        'Master Industry Database sources S08, S09, S11, S12, S13, S14, S26, S35, S37, S43, S47',
        'Probability is the portfolio delivery factor applied to projects flagged as included, and zero '
        'for those excluded.',
        'The register above the research block was designed with a different column set from the register '
        'below it. Rather than maintain the same twenty projects twice, every cell here is a live link, so '
        'the two registers can never diverge.',
        'Total expansion capacity must equal 68.87 Mtpa across all twenty projects, of which 53.62 Mtpa is '
        'flagged as included in the forecast',
        'High', 'Model Assumptions', 'Quarterly', 'FY2026A', 'Was entirely blank.', NUM2))

    # ---- annual capacity addition rows 44-50
    for i in range(7):
        r = 44 + i
        fy = 2027 + i
        w.f(f'B{r}', f'=SUMIFS($F$77:$F$96,$E$77:$E$96,"Brownfield*",$I$77:$I$96,"FY{fy}",'
                     f'$L$77:$L$96,"Yes")', NUM2)
        w.f(f'C{r}', f'=SUMIFS($F$77:$F$96,$E$77:$E$96,"Greenfield*",$I$77:$I$96,"FY{fy}",'
                     f'$L$77:$L$96,"Yes")+SUMIFS($F$77:$F$96,$E$77:$E$96,"Programme",'
                     f'$I$77:$I$96,"FY{fy}",$L$77:$L$96,"Yes")', NUM2)
        w.f(f'D{r}', f'=SUMIFS($F$77:$F$96,$E$77:$E$96,"Debottlenecking*",$I$77:$I$96,"FY{fy}",'
                     f'$L$77:$L$96,"Yes")', NUM2)
        w.f(f'E{r}', f'=SUMIFS($F$77:$F$96,$E$77:$E$96,"Restart*",$I$77:$I$96,"FY{fy}",'
                     f'$L$77:$L$96,"Yes")', NUM2)
        w.f(f'F{r}', f'=SUM(B{r}:E{r})', NUM2)
    rows.append(_sup(
        'Annual capacity addition by project type', '=SUM(F44:F50)', 'Mtpa',
        'Live SUMIFS by project type and commissioning year',
        '=SUMIFS($F$77:$F$96,$E$77:$E$96,"Brownfield*",$I$77:$I$96,"FY2027",$L$77:$L$96,"Yes")',
        'The dated project register below',
        'Row 99 of this sheet, which is the same SUMIFS aggregated across all types',
        'The two undated aspirational programmes - Tata Steel to 40 Mtpa and SAIL to ~35 Mt - are '
        'classified as GREENFIELD because neither company has designated brownfield sites for them.',
        'Debottlenecking and restart return zero because no project in the register carries those types. '
        'They are live formulas rather than hard zeros, so they populate automatically if such a project '
        'is added.',
        'Column F must reconcile to row 99 year by year, and the seven-year total to 53.62 Mtpa on row 101',
        'High', 'This sheet', 'Quarterly', 'FY2026A',
        'Was entirely blank. Note this table is GROSS announced additions, before the 85% delivery factor.',
        NUM2))
    return rows, nas


# ================================================================= 08 Capacity Utilisation
CU_COMPANY = {19: 'Tata Steel', 20: 'JSW Steel', 21: 'SAIL', 22: 'Jindal Steel',
              23: 'Jindal Stainless', 24: 'AM/NS India', 25: 'Shyam Metalics'}
CU_CAP = {19: "='Comparable Valuation'!$K$108", 20: "='Comparable Valuation'!$K$109",
          21: "='Comparable Valuation'!$K$110", 22: "='Comparable Valuation'!$K$111",
          23: "='Comparable Valuation'!$K$112", 24: "='Comparable Valuation'!$K$113"}
CU_PROD = {19: (23.43, "Production Volume'!$J$33"), 20: (29.31, "Production Volume'!$J$37"),
           21: (19.43, "Production Volume'!$J$40"), 22: (9.25, "Production Volume'!$J$43"),
           24: (7.367, "Production Volume'!$J$46")}
BANDS = ('=IF({c}>0.90,"Capacity constrained",IF({c}>0.85,"Tight market",'
         'IF({c}>0.75,"Balanced market",IF({c}>0.65,"Weak market","Significant overcapacity"))))')


def utilisation(wb, audit, eng):
    ws = wb['Capacity Utilisation']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}9', f'=${lc}$66', MT)
        w.link(f'{nc}10', f'=${lc}$65', MT)
        w.link(f'{nc}11', f'=${lc}$67', PCT1)
        w.f(f'{nc}12', f'={lc}66-{lc}65', MT)
        w.f(f'{nc}13', f'=IFERROR(1-{lc}67,"")', PCT1)
        w.f(f'{nc}14', BANDS.format(c=f'{lc}67'), TEXT)

    # ---- industry utilisation rows 19-26 + average row 27
    for i, r in enumerate(range(19, 27)):
        lc = LEG[i]
        w.link(f'B{r}', f'=${lc}$66', MT)
        w.link(f'C{r}', f'=${lc}$65', MT)
        w.f(f'D{r}', f'=IFERROR(C{r}/B{r},"")', PCT1)
    w.f('B27', '=AVERAGE(B19:B26)', MT)
    w.f('C27', '=AVERAGE(C19:C26)', MT)
    w.f('D27', '=IFERROR(C27/B27,"")', PCT1)

    # ---- company-wise utilisation rows 19-26 (columns F..I) + average row 27
    for r, name in CU_COMPANY.items():
        if r in CU_CAP:
            w.link(f'G{r}', CU_CAP[r], MT)
        else:
            w.na(f'G{r}', 'Shyam Metalics does not disclose crude steel capacity; the Master Industry '
                          'Database records aggregate installed METAL capacity across intermediate and '
                          'finished products, which is not comparable.')
        if r in CU_PROD:
            v, ref = CU_PROD[r]
            w.db(f'H{r}', v, MT, ref=f"='[Master Industry Database.xlsx]{ref}")
        else:
            w.na(f'H{r}', 'Crude steel or melt production is not disclosed by this name. Jindal Stainless '
                          'discloses DELIVERIES of 2.57 mn t but not melt production, and the two are not '
                          'interchangeable. Shyam Metalics disclosed no volumes.')
        w.f(f'I{r}', f'=IFERROR(H{r}/G{r},"")', PCT1)
    w.f('G26', "=$C$66-SUM(G19:G25)", MT)
    w.f('H26', "=$C$65-SUM(H19:H25)", MT)
    w.f('I26', '=IFERROR(H26/G26,"")', PCT1)
    w.f('G27', '=SUM(G19:G26)', MT)
    w.f('H27', '=SUM(H19:H26)', MT)
    w.f('I27', '=IFERROR(H27/G27,"")', PCT1)
    nas.append(na_row('Company utilisation - Jindal Stainless production, Shyam Metalics capacity and '
                      'production', 'G25, H23, H25',
                      'Jindal Stainless discloses deliveries but not melt production. Shyam Metalics '
                      'disclosed neither capacity on a crude steel basis nor volumes. Both are absorbed '
                      'into the Others residual on row 26 so the industry total still reconciles to 76.4%.',
                      unit='Mt / Mtpa'))
    rows.append(_sup(
        'Company-wise utilisation, FY2026A', '=I27', '%',
        'Capacity linked to the sourced India-basis figures; production imported from the Master Industry '
        'Database; Others is a residual so the total reconciles',
        '=IFERROR(H19/G19,"") ; Others = $C$65-SUM(H19:H25)',
        'Company results releases; Joint Plant Committee CPSE series for SAIL and RINL',
        'Master Industry Database Production Volume section B and Production Capacity section B',
        'Producer-level utilisation uses each company\'s own reported basis: Tata Steel India proforma '
        'including NINL, JSW Steel India operations, SAIL standalone, Jindal Steel consolidated, AM/NS '
        'India 100% derived fiscal year.',
        'Reported utilisation above 100% is possible and does occur: SAIL\'s Rourkela and Durgapur plants '
        'averaged 103.8% and 101.6% of nameplate over FY2021-FY2025. That is why the national aggregate '
        'ceiling of 92% is applied to the industry and not to individual assets.',
        'Row 27 must return 76.4%, which is the industry figure on research block row 67',
        'Medium', 'Comparable Valuation, Master Industry Database', 'Quarterly', 'FY2026A',
        'Was entirely blank.', PCT1))

    # ---- utilisation bridge rows 32-38
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        pc = LEG[i]
        w.link(f'{nc}32', f'=${pc}$67', PCT1)
        w.f(f'{nc}33', f"=('Steel Demand Model'!{lc}116-'Steel Demand Model'!{pc}116)"
                       f"*'Steel Supply Model'!{lc}93/'Capacity Forecast'!{pc}101", PCT2)
        w.f(f'{nc}34', f"=-'Steel Supply Model'!{lc}99*'Capacity Forecast'!{lc}100"
                       f"/('Capacity Forecast'!{pc}101*'Capacity Forecast'!{lc}101)", PCT2)
        w.f(f'{nc}35', f"=('Steel Supply Model'!{lc}90-'Steel Supply Model'!{pc}90)"
                       f"*'Steel Supply Model'!{lc}93/'Capacity Forecast'!{pc}101", PCT2)
        w.f(f'{nc}37', f'=${lc}$67-${pc}$67', PCT2)
        w.f(f'{nc}36', f'={nc}37-{nc}33-{nc}34-{nc}35', PCT2)
        w.f(f'{nc}38', f'={nc}32+{nc}37', PCT1)
    rows.append(_sup(
        'Utilisation bridge', '=H38', '%',
        'Exact algebraic decomposition of the change in utilisation, with one residual line',
        'Demand = (dConsumption x ratio)/C0 ; Capacity = -P1 x dC/(C0 x C1) ; Export = (dNetExports x '
        'ratio)/C0 ; Import = residual',
        'Computed from the model',
        'Research block row 67, which the closing line must reproduce',
        'The decomposition is exact: utilisation = P/C, so dU = dP/C0 - P1 x dC/(C0 x C1). The demand, '
        'capacity and export lines are those terms; the IMPORT line is the residual.',
        'The residual on row 36 is not purely imports. It comprises the change in variation in stock, the '
        'change in the crude-to-finished ratio, and any shortfall routed to imports when the capacity '
        'ceiling binds. It is shown as a residual so that the bridge closes to the last basis point rather '
        'than approximately.',
        'Row 32 + row 37 must equal row 38, and row 38 must equal research block row 67, in every year',
        'High', 'Steel Demand Model, Steel Supply Model, Capacity Forecast', 'Live', 'Live',
        'Was entirely blank.', PCT1))

    # ---- scenario comparison rows 52-58
    for i, r in enumerate(range(52, 59)):
        for col, sc in zip('BCDE', SCEN):
            w.link(f'{col}{r}', eng.ref(sc, 'util', i + 1), PCT1)
    rows.append(_sup(
        'Four-scenario utilisation path', '=B58', '%',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Rebuilt live from the scenario driver matrix on 02 Model Assumptions',
        'Independent tie-out on 24 Scenario Manager row 72',
        'Utilisation is production divided by CLOSING capacity, which is conservative because capacity '
        'commissioned late in the year is fully in the denominator.',
        'Counter-intuitively utilisation is HIGHEST in the Stress Case, at 86.0% in FY2033E against 84.0% '
        'in the base case. That is because the delivery factor collapses to 55% and secondary additions to '
        '3 Mtpa, so capacity falls faster than demand. High utilisation is therefore not automatically '
        'good news: read it with the capacity and demand rows.',
        'FY2033E must equal 84.0% / 87.9% / 85.5% / 86.0%, the independently computed tie-out values',
        'High', 'Scenario Manager', 'Live', 'Live', 'Was entirely blank.', PCT1))
    return rows, nas



# ================================================================= 09 Steel Price Forecast
def price(wb, audit, eng):
    ws = wb['Steel Price Forecast']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- executive summary rows 9-14
    w.db('B9', 57700, RS, ref="='[Master Industry Database.xlsx]Steel Prices'!$I$22")
    for i, nc in enumerate('CDEFGHI'):
        lc = LEG[i + 1]
        w.f(f'{nc}9', f'=$B$9*{lc}103/$C$103', RS)
    w.narow(NEW8, 10, ('No India export HRC price on an FOB basis was retrievable. The Master Industry '
                       'Database carries a China FOB print of US$500/t and an import-parity observation '
                       'that Indian domestic HRC traded at a DISCOUNT of US$29-45/t to landed imports, but '
                       'neither is an India export price. Converting the domestic rupee price at USD/INR '
                       'would produce a domestic price in dollars, not an export price, and is not done.'))
    nas.append(na_row('Export HRC price', 'B10:I10',
                      'No India export HRC FOB assessment was retrievable. See the cell note.', unit='US$/t'))
    for i, nc in enumerate(NEW8):
        lc = LEG[i]
        if i == 0:
            w.na(f'{nc}11', 'A FY2025 blended realisation is not carried by the model, so a FY2026 growth '
                            'rate cannot be computed. The Master Industry Database records the FY2026 price '
                            'PATH - Rs 46,000/t in early Dec-2025 rising to Rs 57,700/t at end-Mar-2026 - '
                            'but no fiscal-year average for either year.')
            w.na(f'{nc}14', 'Requires a FY2025 realisation, which is not available - see row 11.')
        else:
            pc = LEG[i - 1]
            ma = MACOL[i - 1]
            w.f(f'{nc}11', f'=IFERROR({lc}103/{pc}103-1,"")', PCT1)
            w.f(f'{nc}14', f"=IFERROR((1+{nc}11)/(1+'Model Assumptions'!{ma}$11)-1,\"\")", PCT1)
        w.link(f'{nc}12', f'=${lc}$103', RS)
        w.f(f'{nc}13', f'=IFERROR({lc}103/$C$103*100,"")', NUM1)
    nas.append(na_row('FY2026A steel price growth and real price growth', 'B11, B14',
                      'A FY2025 fiscal-year average realisation does not exist in the public domain, so a '
                      'FY2026 growth rate cannot be struck.', unit='%'))

    # ---- product price series rows 19-25
    w.link('B19', '=$B$9', RS)
    w.db('B21', 59800, RS, ref="='[Master Industry Database.xlsx]Steel Prices'!$I$25")
    w.db('B24', 38850, RS, ref="='[Master Industry Database.xlsx]Steel Prices'!$I$26")
    w.link('B25', '=$C$103', RS)
    for i, nc in enumerate('CDEFGHI'):
        lc = LEG[i + 1]
        for r in (19, 21, 24):
            w.f(f'{nc}{r}', f'=$B${r}*{lc}103/$C$103', RS)
        w.link(f'{nc}25', f'=${lc}$103', RS)
    for r in (20, 22, 23):
        w.narow(NEW8, r, NA['product_price'])
    nas.append(na_row('CRC, wire rod and plate price series', 'B20:I20, B22:I23', NA['product_price'],
                      unit='Rs/t'))
    rows.append(_sup(
        'Product price series', '=I25', 'Rs/t',
        'FY2026A imported from the Master Industry Database (RED); forecast years indexed to the blended '
        'realisation so that no product price can move independently of the model',
        '=$B$19*D103/$C$103',
        'ICRA end-March 2026 assessment for HRC (Rs 57,700/t); ETInfra trade assessment for rebar; IDBI '
        'Capital for billet ex-Raipur',
        'BigMint rebar benchmark Rs 50,900/t at 31-Jul-2026; IDBI weekly HRC Rs 58,200/t at 22-Jun-2026',
        'PERIOD MISMATCH DISCLOSED: HRC is an end-March-2026 assessment, rebar a March-2026 trade '
        'assessment and billet a 22-Jun-2026 print. None is a FY2026 fiscal-year average, because none '
        'exists.',
        'The Master Industry Database records as a MATERIAL LIMITATION that no fiscal-year average HRC '
        'realisation is published for India. Product prices are therefore anchored on the best available '
        'point observation and indexed forward to the blended realisation, which is itself derived from '
        'company disclosure rather than from a price assessment. Indexing rather than forecasting each '
        'product separately guarantees the product table cannot contradict the revenue model.',
        'Row 25 must equal research block row 103 exactly; HRC of Rs 57,700/t sits 3.8% below the blended '
        'realisation of Rs 59,974/t, which is consistent with a product mix containing higher-value CRC, '
        'coated and stainless volumes',
        'Medium for HRC and billet, Low for rebar', 'Master Industry Database, this sheet',
        'Monthly for the price assessments', 'Mar-2026 to Jul-2026', 'Was entirely blank.', RS))

    # ---- price bridge rows 30-41 (column B is a WEIGHT column, C..I are FY2027E..FY2033E)
    for i, nc in enumerate('CDEFGHI'):
        lc = LEG[i + 1]
        pc = LEG[i]
        w.f(f'{nc}30', f'=${pc}$103', RS)
        w.f(f'{nc}33', f'={lc}97*({lc}102-1)', RS)
        w.f(f'{nc}40', f'={lc}103-{pc}103', RS)
        w.link(f'{nc}41', f'=${lc}$103', RS)
    bridge_na = ('The realisation driver in this model is a single calibrated scenario path, not a '
                 'multi-factor price build. It cannot be decomposed into demand, supply-gap, iron ore, '
                 'coking coal, energy, import-pressure, export-opportunity and inflation channels without '
                 'asserting weights that are not sourced, and any such decomposition would be circular '
                 'because the driver is calibrated to an observed margin envelope rather than built up from '
                 'these components. The one price channel the model DOES model explicitly - the lagged '
                 'capacity utilisation feedback - is populated on row 33. The equivalent decomposition on '
                 'the COST side is fully populated on 10 Raw Material Forecast and 11 Cost Curve.')
    for r in (31, 32, 34, 35, 36, 37, 38, 39):
        w.narow('CDEFGHI', r, bridge_na)
    for r in range(30, 42):
        w.na(f'B{r}', 'A weighting scheme across price drivers is not asserted, because the model does not '
                      'decompose realisation into weighted channels - see the note on the shaded cells to '
                      'the right.')
    nas.append(na_row('Price bridge - demand, supply gap, raw material, energy, trade and inflation effects',
                      'B31:I32, B34:I39 and the weight column', bridge_na, unit='Rs/t'))
    rows.append(_sup(
        'Price bridge - capacity utilisation effect', '=I33', 'Rs/t',
        'Computed in-cell as the realisation driver multiplied by the price adjustment factor less one',
        '=D97*(D102-1)',
        'Driver D23, price sensitivity to capacity utilisation - INDICATIVE and unsourced',
        'Research block rows 99 to 103',
        'The feedback is LAGGED one year deliberately, so that price responds to the PRIOR year utilisation '
        'and no circular reference is created.',
        'This is the mechanism that makes the workbook an integrated model rather than a set of independent '
        'forecasts: capacity additions depress utilisation, depressed utilisation depresses realisation, '
        'and depressed realisation compresses margin. It is also the single most important unsourced '
        'parameter in the model, which is why it is held at 0.40x across all four scenarios and quantified '
        'on 22 Sensitivity Analysis section E.',
        'Rows 30 + 40 must equal row 41 in every year; check A17 on 25 Audit Checks tests that the price '
        'adjustment factor stays inside 0.85x to 1.15x',
        'Low', 'Model Assumptions', 'Quarterly review', 'Live',
        'The eight channels the model does not decompose are shaded orange rather than filled with '
        'plausible-looking numbers.', RS))

    # ---- cost floor analysis rows 46-54
    w.f('B46', "='Cost Curve'!$C$80*'Model Assumptions'!$D$20*'Model Assumptions'!$D$21", RS)
    w.f('B47', "='Cost Curve'!$C$80*'Model Assumptions'!$D$20*'Model Assumptions'!$D$21", RS)
    w.f('B48', "='Cost Curve'!$C$80*'Model Assumptions'!$D$20*(1-2*'Model Assumptions'!$D$21)", RS)
    for r in (49, 50):
        w.na(f'B{r}', 'Power and logistics are not separately identified. They sit inside CONVERSION COST '
                      'on row 51, which the model indexes to RBI CPI as a single block. No Indian producer '
                      'discloses a power or freight cost per tonne of crude steel in the sources available.')
    w.f('B51', "='Cost Curve'!$C$80*(1-'Model Assumptions'!$D$20)", RS)
    w.link('B52', "='Cost Curve'!$C$80", RS)
    w.link('B53', "='EBITDA Model'!$C$101", RS)
    w.link('B54', '=$C$103', RS)
    nas.append(na_row('Cost floor - power and logistics', 'B49:B50',
                      'Inside conversion cost on row 51; not separately disclosed by any producer.',
                      unit='Rs/t'))
    rows.append(_sup(
        'Cost floor analysis, FY2026A', '=B52', 'Rs/t',
        'Computed in-cell by applying the raw material share and basket weights to the calibrated cash cost',
        "='Cost Curve'!$C$80*'Model Assumptions'!$D$20*'Model Assumptions'!$D$21",
        'Cash cost of Rs 49,241/t derived from the Master Industry Database as realisation less EBITDA per '
        'tonne, volume-weighted across the four majors that disclose both',
        'Drivers D13 (60% raw material share) and D14 (45% iron ore weight), both INDICATIVE',
        'The basket has exactly three components: iron ore 45%, coking coal 45%, other ferrous and fluxes '
        '10%. Finer decomposition is not sourced.',
        'Iron ore Rs 13,295/t plus coking coal Rs 13,295/t plus fluxes Rs 2,954/t plus conversion Rs '
        '19,696/t sums to Rs 49,241/t, which is the calibrated cash cost. The allocation is therefore '
        'internally exact even though the split itself is indicative.',
        'Rows 46 + 47 + 48 + 51 must equal row 52; row 52 + row 53 must equal row 54',
        'Low for the split, Medium for the total', 'Cost Curve, EBITDA Model, Model Assumptions',
        'Quarterly', 'FY2026A', 'Was entirely blank.', RS))

    # ---- international price comparison rows 59-63
    w.link('B59', '=$B$9', RS)
    for i, nc in enumerate('CDE'):
        lc = LEG[i + 1]
        w.f(f'{nc}59', f'=$B$9*{lc}103/$C$103', RS)
    w.db('B60', 48580, RS, ref="='[Master Industry Database.xlsx]Steel Prices'!$I$27")
    w.narow('CDE', 60, 'The model does not forecast Chinese steel prices. Doing so would require a view on '
                       'Chinese supply-side policy and property demand that is outside its scope.')
    for r in (61, 62, 63):
        w.narow('BCDE', r, ('No Japan, Europe or USA HRC assessment was retrievable. The Master Industry '
                            'Database carries India, China, South East Asia and CIS prints only - South '
                            'East Asia Rs 50,938/t and CIS Rs 51,410/t at 22-Jun-2026 - and those regions '
                            'are not the rows in this table.'))
    nas.append(na_row('International price comparison - Japan, Europe, USA, and the China forecast',
                      'B61:E63, C60:E60',
                      'Only India, China, South East Asia and CIS prints exist in the Master Industry '
                      'Database. Regional price forecasting is outside the scope of this model.',
                      unit='Rs/t'))
    rows.append(_sup(
        'India-China HRC spread', '=B59-B60', 'Rs/t',
        'Computed in-cell from two sourced point observations',
        '=B59-B60',
        'IDBI Capital regional snapshot, week ended 22-Jun-2026',
        'Master Industry Database KPI K19, which records the same spread at Rs 9,620/t',
        'Both prints are converted to rupees by the broker on the same date and exclude GST.',
        'The Indian premium over Chinese FOB of roughly Rs 9,000-9,600/t, about 20%, is exactly what the '
        'safeguard duty defends. The duty steps down from 12% to 11.5% from 21-Apr-2026 and to 11% before '
        'expiring on 20-Apr-2028, which is the single largest identifiable policy risk to FY2029E '
        'realisation.',
        'Consistent with the import-parity observation that domestic HRC traded at a US$29-45/t DISCOUNT to '
        'landed imports in Apr-2026, i.e. the duty was doing the work',
        'Medium', 'Master Industry Database', 'Weekly', '22-Jun-2026', 'Was blank.', RS))

    # ---- scenario analysis rows 68-74
    for i, r in enumerate(range(68, 75)):
        for col, sc in zip('BCDE', SCEN):
            w.link(f'{col}{r}', eng.ref(sc, 'real', i + 1), RS)
    rows.append(_sup(
        'Four-scenario realisation path', '=B74', 'Rs/t',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Rebuilt live from the scenario driver matrix on 02 Model Assumptions rows 70 to 73',
        'Independent tie-out on 24 Scenario Manager row 73',
        'Realisation is the DRIVER path multiplied by the lagged utilisation adjustment factor.',
        'Note that the Stress Case realisation of Rs 70,222/t in FY2033E sits ABOVE the Bear Case at Rs '
        '60,700/t. That is not an error: the Stress Case is a stagflationary shock in which the rupee '
        'collapses to 113 and coking coal reaches US$300/t, so nominal rupee prices rise while margins '
        'collapse. The Bear Case is a demand-led downturn in which raw materials fall with demand.',
        'FY2033E must equal 69,006 / 70,422 / 60,700 / 70,222, the independently computed tie-out values',
        'Medium', 'Scenario Manager', 'Live', 'Live', 'Was entirely blank.', RS))

    # ---- price elasticity rows 79-86
    w.link('B82', "='Model Assumptions'!$E$30", X2)
    elast_na = ('This model contains exactly one price elasticity - the lagged sensitivity of realisation to '
                'the capacity utilisation gap, driver D23, which is populated on row 82. Realisation is '
                'otherwise a calibrated scenario path, so no elasticity of PRICE to GDP, demand, the supply '
                'gap, iron ore, coking coal or the exchange rate exists in the model and none is asserted. '
                'The elasticities of EBITDA to each of those variables ARE computed, and are on 22 '
                'Sensitivity Analysis section A.')
    for r in (79, 80, 81, 83, 84, 85, 86):
        w.na(f'B{r}', elast_na)
    nas.append(na_row('Price elasticities other than to capacity utilisation', 'B79:B81, B83:B86',
                      elast_na, unit='x'))
    return rows, nas


# ================================================================ 10 Raw Material Forecast
def rawmat(wb, audit, eng):
    ws = wb['Raw Material Forecast']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- executive summary rows 9-16
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}9', f'=${lc}$85', USD)
        w.link(f'{nc}10', f'=${lc}$86', NUM0)
        ma = 'D' if lc == 'C' else MACOL[LEG.index(lc) - 1]
        w.f(f'{nc}16', f"='Cost Curve'!{lc}80*'Model Assumptions'!{ma}$20", RS)
    for r in (11, 13, 14):
        w.narow(NEW8, r, NA['raw_hist'])
    w.db('B12', 111.5, USD, ref="='[Master Industry Database.xlsx]Coking Coal Prices'!$I$28")
    w.narow('CDEFGHI', 12, 'Thermal coal is an energy-cost reference only. The model indexes the '
                           'non-raw-material element of cash cost to RBI CPI rather than to a thermal coal '
                           'path, so no thermal coal forecast is made. The Master Industry Database warns '
                           'explicitly against using the thermal series as a coking coal proxy.')
    w.db('B15', 123200, RS,
         ref='Ferro chrome Rs 1,23,200/t - IDBI Capital input-cost deck, week ended 22-Jun-2026 '
             '(Master Industry Database source S21)')
    w.narow('CDEFGHI', 15, 'Ferro alloys are inside the 10% "other ferrous and fluxes" weight of the raw '
                           'material basket and are not forecast separately.')
    nas.append(na_row('PCI coal, scrap and limestone price series', 'B11:I11, B13:I14', NA['raw_hist'],
                      unit='US$/t and Rs/t'))
    nas.append(na_row('Thermal coal and ferro alloy forecasts', 'C12:I12, C15:I15',
                      'Neither is forecast separately: thermal coal is an energy reference and ferro alloys '
                      'sit inside the 10% other-ferrous basket weight.', unit='US$/t and Rs/t'))

    # ---- historical prices rows 21-29 (B = FY2021 .. G = FY2026)
    ore_hist = {'B': 128.03, 'C': 155.52, 'D': 117.20, 'E': 119.91, 'F': 103.97, 'G': 100.52}
    for col, v in ore_hist.items():
        w.db(f'{col}21', v, USD, ref="='[Master Industry Database.xlsx]Iron Ore Prices'!$C$27:$C$32")
    thermal_hist = {'E': 145.1, 'F': 131.89, 'G': 111.5}
    for col, v in thermal_hist.items():
        w.db(f'{col}24', v, USD, ref="='[Master Industry Database.xlsx]Industry KPIs'!$F$35:$H$35")
    for col in 'BCD':
        w.na(f'{col}24', 'The Master Industry Database carries Australian thermal coal fiscal-year averages '
                         'from FY2024 only. Earlier years were not compiled in this research cycle.')
    coal_na = ('CRITICAL AND DELIBERATE. The Master Industry Database asserts NO fiscal-year average for '
               'premium hard coking coal, and records this as Conflict C01. The Ministry of Steel publishes '
               'the series only as a chart and the two available assessors disagree materially - the '
               'Ministry put March-2026 at US$225/t while IDBI Capital put it at US$186/t, a US$39/t gap on '
               'the same month. Point observations that ARE available: US$354/t Oct-2023 (cycle peak), '
               'US$175/t Mar-2025 (trough), US$246/t Feb-2026, US$225/t Mar-2026. Averaging them would '
               'manufacture a number the sources do not support.')
    for r in (22, 23, 25, 26, 27, 28, 29):
        w.narow('BCDEFG', r, coal_na if r == 22 else NA['raw_hist'])
    nas.append(na_row('Historical premium HCC coking coal fiscal-year averages', 'B22:G22', coal_na,
                      unit='US$/t'))
    nas.append(na_row('Historical PCI coal, scrap, limestone, dolomite, ferro manganese and ferro silicon',
                      'B23:G23, B25:G29', NA['raw_hist'], unit='US$/t and Rs/t'))
    rows.append(_sup(
        'Historical iron ore, FY2021-FY2026', '=G21', 'US$/dmt',
        'Master Industry Database import (RED - to be reconnected as an external link)',
        "='[Master Industry Database.xlsx]Iron Ore Prices'!$C$27:$C$32",
        'World Bank Commodity Price Data (Pink Sheet), monthly file updated 02-Jul-2026',
        'Calendar-year series on the same database sheet; Master Industry Database KPI K15',
        'Fiscal-year averages are the arithmetic mean of twelve monthly nominal prints, 62% Fe fines spot '
        'CFR China.',
        'The highest-quality price series available to this model: an unbroken monthly history from a '
        'primary multilateral source, with no vintage or assessor conflict. Three consecutive annual '
        'declines from the FY2022 peak of US$155.52/dmt to US$100.52/dmt in FY2026 have been a persistent '
        'tailwind for Indian margins.',
        'FY2016 low of US$52.18/dmt and FY2022 peak of US$155.52/dmt bracket the full observed range; the '
        'base case forecast of US$90-102/dmt sits comfortably inside it',
        'High', 'Master Industry Database', 'Monthly, on each Pink Sheet release', 'Jun-2026',
        'Six years of history were blank. Coking coal remains blank by design.', USD))

    # ---- forecast rows 34-41
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        w.link(f'{nc}34', f'=${lc}$85', USD)
        w.link(f'{nc}35', f'=${lc}$86', NUM0)
    for r in (36, 37, 38, 39, 40, 41):
        w.narow(NEW7, r, ('The model forecasts exactly two raw material prices - iron ore and premium HCC '
                          'coking coal - because those are the two for which a defensible starting point '
                          'exists. Everything else sits inside the 10% "other ferrous and fluxes" basket '
                          'weight or inside conversion cost, both of which are indexed rather than '
                          'forecast. Asserting a scrap, PCI, limestone, dolomite or ferro alloy path would '
                          'add apparent precision without adding information.'))
    nas.append(na_row('Forecast PCI coal, thermal coal, scrap, limestone, dolomite and ferro alloys',
                      'B36:H41',
                      'The model forecasts only iron ore and premium HCC coking coal; all other inputs are '
                      'indexed inside the basket or inside conversion cost.', unit='US$/t and Rs/t'))

    # ---- cost build-up rows 46-53
    w.link('C46', '=$D$89', RS)
    w.link('C47', '=$D$90', RS)
    w.f('D46', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*'Model Assumptions'!$E$21", RS)
    w.f('D47', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*'Model Assumptions'!$E$21", RS)
    w.f('D49', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*(1-2*'Model Assumptions'!$E$21)", RS)
    w.f('D53', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20", RS)
    for r in range(46, 54):
        w.na(f'B{r}', NA['consumption_coef'])
    for r in (48, 50, 51, 52):
        w.narow('CD', r, ('PCI coal, ferro alloys, electrodes and power are not separately identified in '
                          'the model. PCI and ferro alloys are inside the 10% other-ferrous basket weight '
                          'carried on row 49; electrodes and power are inside conversion cost. Point '
                          'observations that exist but are not fiscal-year averages: graphite electrode UHP '
                          'US$4,189/t and ferro chrome Rs 1,23,200/t, both at 22-Jun-2026.'))
    w.na('C53', 'Not applicable - row 53 is a total, not a price.')
    w.na('C49', 'No price series exists for the "other ferrous and fluxes" block. It is a residual 10% '
                'basket weight, not a traded commodity with a quoted price, so it is indexed rather '
                'than priced. The COST of the block is on D49 and is computed.')
    nas.append(na_row('Specific consumption coefficients per tonne of crude steel', 'B46:B53',
                      NA['consumption_coef'], unit='t/t'))
    nas.append(na_row('Cost build-up for PCI coal, ferro alloys, electrodes and power', 'C48:D48, C50:D52',
                      'Inside the 10% other-ferrous basket weight or inside conversion cost. Not separately '
                      'identified by any producer.', unit='Rs/t'))
    rows.append(_sup(
        'Raw material cost build-up, FY2027E', '=D53', 'Rs/t',
        'Computed in-cell by applying the raw material share and basket weights to the forecast cash cost',
        "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*'Model Assumptions'!$E$21",
        'Iron ore from the World Bank Pink Sheet; coking coal from the Ministry of Steel; both converted at '
        'the USD/INR driver',
        'Drivers D13 and D14; research block rows 89, 90 and 95',
        'The model is calibrated TOP-DOWN. It does not build cost from consumption coefficients, because '
        'those are plant-specific and were not disclosed; it starts from the observed FY2026 cash cost of '
        'Rs 49,241/t and indexes it.',
        'Top-down calibration is the more defensible method here. The FY2026 anchor is an OBSERVED number - '
        'realisation less EBITDA per tonne for four producers that disclose both - whereas a bottom-up '
        'build would require six or seven unsourced coefficients, each compounding the error.',
        'Rows 46 + 47 + 49 must equal row 53, and row 53 divided by the cash cost on 11 Cost Curve must '
        'equal driver D13 of 60%',
        'Medium for the total, Low for the split', 'Cost Curve, Model Assumptions', 'Quarterly', 'FY2027E',
        'Was entirely blank.', RS))

    # ---- import dependency rows 58-61
    imp_na = ('Material-level import dependency is published by the Ministry of Mines and DGCIS but is not '
              'carried in the Master Industry Database, which covers finished STEEL trade only. The steel '
              'trade shares that ARE available are on 19 Trade Model: imports Korea 35.4%, China 23.5% and '
              'Japan 20.2% of FY2026 volumes. Asserting an iron ore, coking coal, scrap or ferro alloy '
              'import share without a source would be an invention.')
    for r in range(58, 62):
        w.narow('BCD', r, imp_na)
    nas.append(na_row('Raw material import dependency and source countries', 'B58:D61', imp_na))

    # ---- scenario analysis rows 66-70
    for col, sc in zip('BCDE', SCEN):
        w.link(f'{col}66', eng.ref(sc, 'ore', 7), USD)
        w.link(f'{col}67', eng.ref(sc, 'coal', 7), NUM0)
    for r in (68, 69, 70):
        w.narow('BCDE', r, 'Scrap, thermal coal and ferro alloys are not scenario drivers in this model.')
    nas.append(na_row('Scenario analysis for scrap, thermal coal and ferro alloys', 'B68:E70',
                      'Not scenario drivers in this model.', unit='US$/t and Rs/t'))
    rows.append(_sup(
        'Four-scenario input cost path, FY2033E', '=B66', 'US$/dmt',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Scenario driver matrix on 02 Model Assumptions rows 74 to 81',
        'Observed historical range: iron ore US$52.18/dmt (FY2016) to US$155.52/dmt (FY2022); coking coal '
        'US$175/t (Mar-2025) to US$354/t (Oct-2023)',
        'The BEAR case has raw materials FALLING, because it is a demand-led downturn. Only the STRESS '
        'case has them rising.',
        'This is the single most important structural feature of the scenario design. A naive bear case '
        'that combines weak demand with high input costs is not a coherent state of the world: iron ore and '
        'coking coal are themselves demand-driven. Separating a demand-led downturn (Bear) from a '
        'stagflationary supply shock (Stress) is what allows the model to show a 7-11% margin in one and '
        'a -9.2% trough in the other.',
        'Every scenario value sits inside the observed historical range for both commodities',
        'High for iron ore, Medium for coking coal', 'Scenario Manager, Model Assumptions', 'Live', 'Live',
        'Was entirely blank.', USD))

    # ---- sensitivity rows 75-79
    w.f('B75', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*'Model Assumptions'!$E$21*0.1", RS)
    w.f('B76', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*'Model Assumptions'!$E$21*0.1", RS)
    w.f('B78', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*2*'Model Assumptions'!$E$21*0.05", RS)
    w.na('B77', 'Scrap is not in the raw material basket, so a scrap price move has no effect in this '
                'model. That is a stated limitation: the DRI-EAF and scrap-EAF routes are a growing share '
                'of Indian capacity and are not separately costed.')
    w.na('B79', 'Freight is inside conversion cost, which is indexed to RBI CPI as a single block. A '
                'freight-specific shock is therefore not separable and would be understated by this model.')
    nas.append(na_row('Sensitivity to scrap and freight', 'B77, B79',
                      'Scrap is not in the basket and freight is inside conversion cost; neither is '
                      'separable in this model.', unit='Rs/t'))
    rows.append(_sup(
        'Raw material cost sensitivity', '=B75', 'Rs/t',
        'Computed in-cell: cash cost x raw material share x basket weight x the shock',
        "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*'Model Assumptions'!$E$21*0.1",
        'Computed from the model',
        '22 Sensitivity Analysis section A tornado, which quantifies the same shocks at EBITDA level',
        'A 10% move in iron ore affects 45% of the 60% of cash cost that is raw material, i.e. 2.7% of '
        'cash cost. USD/INR affects 90% of the basket, because only the 10% other-ferrous weight is '
        'domestically priced.',
        'This is why USD/INR outranks both iron ore and coking coal individually in the tornado: the rupee '
        'moves BOTH dollar-denominated inputs at once.',
        'Consistent with the tornado on 22 Sensitivity Analysis: iron ore -12.1%, coking coal -11.6% and '
        'USD/INR -23.7% of FY2033E EBITDA for a 10% adverse move',
        'Medium', 'Cost Curve, Model Assumptions, Sensitivity Analysis', 'Live', 'Live',
        'Was entirely blank.', RS))
    return rows, nas
