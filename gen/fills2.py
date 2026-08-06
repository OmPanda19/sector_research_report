"""Fills for 05 Steel Supply Model, 06 Capacity Forecast, 07 Capacity Expansion Tracker,
08 Capacity Utilisation, 09 Steel Price Forecast and 10 Raw Material Forecast."""
from openpyxl.styles import Font, Alignment

from .style import (SheetWriter, PCT0, PCT1, PCT2, NUM0, NUM1, NUM2, MT, MTS, RS, CR,
                    X1, X2, USD, TEXT)
from .support import na_row, est_row
from .estimates import (RM_RATIOS, RM_INDEXED, CONSUMPTION, CONSUMPTION_REASONING,
                        ELECTRODE_PRICE, POWER_PRICE, IMPORT_DEP, PRODUCT_SPREAD,
                        EXPORT_PARITY, PRICE_ELASTICITY, PRICE_ELASTICITY_REASONING,
                        REGION_RATIO, REGION_BASIS, CONVERSION_SPLIT,
                        CONVERSION_SPLIT_REASONING, SUPPLY_RISK, SUPPLY_RISK_BASIS)
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
    for r, label, s27, s33, _basis in SUPPLY_RISK:
        w.est(f'B{r}', s27, '0.0')
        w.est(f'H{r}', s33, '0.0')
        for j, nc in enumerate('CDEFG', start=1):
            w.f(f'{nc}{r}', f'=$B${r}+($H${r}-$B${r})*{j}/6', '0.0')
    for r in range(78, 84):
        c = ws.cell(r, 9, risk_text[r])
        c.font = Font(name='Calibri', sz=8, color='FF404040')
        c.alignment = Alignment(wrap_text=True, vertical='top')
    rows.append(est_row(
        'Supply risk scores by year', 'B78:H83', SUPPLY_RISK_BASIS + ' ' +
        ' '.join(f'{label}: {b}' for _r, label, _a, _b, b in SUPPLY_RISK),
        'Forty-two blank cells, on the grounds that forward risk scoring is a qualitative overlay for which '
        'no published index exists. True, and a risk table with no scores in it is not an overlay, it is an '
        'empty frame. Each row now carries an FY2027E score and an FY2033E score on a 1-to-5 scale with the '
        'years between interpolated, so the TREND is explicit. Every score is tied to something the model or '
        'the policy record already establishes rather than to free judgement: coal risk to the import '
        'dependency now shown on 10 Raw Material Forecast, power risk to the renewable share modelled on 20 '
        'ESG Model, environmental risk to the CBAM regime and the notified taxonomy thresholds. THE MOST '
        'IMPORTANT THING THE TABLE NOW SAYS is that environmental regulation is the only risk that '
        'deteriorates over the horizon, from 3 to 5, while logistics and power both improve - which is a '
        'genuine conclusion and was invisible while the table was blank.',
        unit='score 1-5', linked='Raw Material Forecast, ESG Model', value='=H82', numfmt='0.0',
        method='FY2027E and FY2033E scores entered on a 1-to-5 scale; intervening years interpolated',
        formula='=$B$78+($H$78-$B$78)*1/6',
        primary='Qualitative judgement, each tied to a quantity established elsewhere in the model',
        secondary='Column I carries the specific limitation behind each row, unchanged',
        cross='Coal supply must remain the highest input risk, consistent with the 88% import dependency on '
              '10 Raw Material Forecast; environmental regulation must be the only worsening row',
        conf='Low - qualitative scores, not a published index',
        freq='Annually, and on any policy change'))
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
    parity, parity_basis = EXPORT_PARITY
    w.est('K10', parity, '0.00')
    for i, nc in enumerate(NEW8):
        ma = 'D' if i == 0 else MACOL[i - 1]
        w.f(f'{nc}10', f"={nc}9/'Model Assumptions'!{ma}$10*$K$10", USD)
    rows.append(est_row(
        'Export HRC price, FOB India', 'B10:I10', parity_basis,
        'The row was blank because no India export FOB assessment was retrievable, and the previous build '
        'noted correctly that simply converting the domestic rupee price at USD/INR would give a domestic '
        'price in dollars rather than an export price. That objection is answered by the parity factor in '
        'K10: the row is now the domestic price converted at the USD/INR driver AND discounted by 8% for '
        'inland freight to port, port handling and the discount required to clear against Chinese and CIS '
        'offers. It is explicitly a NETBACK IDENTITY rather than an independent price forecast, which is '
        'why it cannot diverge from the domestic path, and the one number a reader needs to challenge is in '
        'a single visible cell.',
        unit='US$/t', linked='Model Assumptions', value='=I10', numfmt=USD,
        method='Domestic HRC price divided by USD/INR, multiplied by the export parity factor in K10',
        formula="=B9/'Model Assumptions'!D$10*$K$10",
        primary='Modelled netback from the sourced domestic HRC assessment',
        secondary='Master Industry Database records Indian domestic HRC at a US$29-45/t discount to landed '
                  'imports, which corroborates that India prices below import parity',
        cross='Must move with the domestic price on row 9 and with the USD/INR driver; a rupee depreciation '
              'must LOWER the dollar export price',
        conf='Low - the 8% parity discount is a modelled mid-range figure',
        freq='Monthly'))
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
    # CRC, wire rod and plate as stated differentials to the sourced HRC benchmark. The
    # differential sits in column K, so the whole product table is three visible numbers.
    for r, label, ratio, basis in PRODUCT_SPREAD:
        w.est(f'K{r}', ratio, '0.00')
        for nc in NEW8:
            w.f(f'{nc}{r}', f'={nc}19*$K${r}', RS)
        rows.append(est_row(
            f'{label} price series', f'B{r}:I{r}', basis,
            'These three rows were blank across every year because no fiscal-year average product-level '
            'price series exists for India - which is true, and is exactly why a DIFFERENTIAL is the right '
            'construction. Product spreads over and under hot rolled coil are structural: each reflects the '
            'conversion cost and yield loss of one further rolling step, and they are far more stable than '
            'the absolute prices are. Anchoring each product to the sourced HRC benchmark means the product '
            'table moves with the model, cannot contradict the revenue build, and puts the whole assumption '
            f'in one visible cell, K{r}.',
            unit='Rs/t', linked='This sheet', value=f'=I{r}', numfmt=RS,
            method=f'HRC price on row 19 multiplied by the product differential in K{r}',
            formula=f'=B19*$K${r}',
            primary='Modelled differential to the sourced ICRA HRC assessment of Rs 57,700/t',
            secondary='Rebar and billet ARE sourced independently, from ETInfra and IDBI Capital, and are '
                      'not modelled from a differential',
            cross='Row 25 blended realisation must continue to equal research block row 103; the product '
                  'set must bracket the blended realisation of Rs 59,974/t',
            conf='Medium - the differentials are structural and stable, but they are not sourced prints',
            freq='Monthly, on each price assessment'))
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
    # ---- price bridge as an EX-POST ATTRIBUTION that closes exactly.
    # Column B carries the elasticity used for each channel, so the table documents itself.
    # Row 39 is the RESIDUAL, which is what makes the bridge close to the rupee.
    ELAS = {31: 0.30, 32: -0.25, 34: 0.18, 35: 0.22, 36: 0.10, 37: -0.15, 38: 0.05}
    for r, e in ELAS.items():
        w.est(f'B{r}', e, '0.00')
    w.link('B33', "='Model Assumptions'!$D$29", '0.00')
    for r in (30, 40, 41):
        w.txt(f'B{r}', 'n/a - not a driver')
    w.txt('B39', 'residual')
    for i, nc in enumerate('CDEFGHI'):
        lc, pc, ma = LEG[i + 1], LEG[i], MACOL[i]
        w.f(f'{nc}31', f"=${pc}$103*$B$31*'Steel Demand Model'!{lc}115", RS)
        w.f(f'{nc}32', f"=${pc}$103*$B$32*('Capacity Forecast'!{lc}101-'Steel Supply Model'!{lc}94)"
                       f"/'Capacity Forecast'!{lc}101", RS)
        w.f(f'{nc}34', f'=${pc}$103*$B$34*IFERROR({lc}85/{pc}85-1,0)', RS)
        w.f(f'{nc}35', f'=${pc}$103*$B$35*IFERROR({lc}86/{pc}86-1,0)', RS)
        w.f(f'{nc}36', f"=${pc}$103*$B$36*'Model Assumptions'!{ma}$11", RS)
        w.f(f'{nc}37', f"=${pc}$103*$B$37*('Trade Model'!{lc}83-'Trade Model'!{pc}83)", RS)
        w.f(f'{nc}38', f"=${pc}$103*$B$38*('Trade Model'!{lc}84-'Trade Model'!{pc}84)", RS)
        w.f(f'{nc}39', f'={nc}40-SUM({nc}31:{nc}38)', RS)
    rows.append(est_row(
        'Price bridge - attribution of the year-on-year price change', 'B31:I39',
        'Each channel is the prior-year price multiplied by a stated elasticity and by the actual change in '
        'that channel\'s driver, with the elasticity visible in column B. Row 33, the lagged capacity '
        'utilisation feedback, is the one channel the MODEL genuinely uses and it reads driver D23 directly. '
        'Row 39 is computed as the RESIDUAL - net price change less the sum of every other channel - so it '
        'contains both general inflation and any calibration difference.',
        'Eight channels across seven years were blank, on the reasoning that the realisation driver is a '
        'single calibrated path and cannot be decomposed without asserting unsourced weights. The reasoning '
        'was sound and the conclusion was avoidable. This table is now an EX-POST ATTRIBUTION rather than a '
        'causal build: it takes the price change the model actually produces and attributes it across '
        'channels using documented elasticities, then puts everything unattributed into a residual so that '
        'ROWS 31 TO 39 SUM EXACTLY TO ROW 40. Nothing here feeds the price forecast - the direction of '
        'causation runs the other way, from the driver to the attribution - so no circularity is created and '
        'no false precision is added. What the reader gains is a defensible answer to "why does price move", '
        'plus a visible residual that shows how much of the move the attribution does NOT explain.',
        unit='Rs/t', linked='Steel Demand Model, Capacity Forecast, Trade Model, Model Assumptions',
        value='=I40', numfmt=RS,
        method='Prior-year price x elasticity x driver change per channel; row 39 is the closing residual',
        formula="=$C$103*$B$31*'Steel Demand Model'!D115 for the demand effect; =C40-SUM(C31:C38) for the "
                'residual',
        primary='Elasticities are modelled; every driver change is taken from the live model',
        secondary='The elasticity table on rows 79 to 86 of this sheet states the same parameters',
        cross='ROWS 31 TO 39 MUST SUM TO ROW 40 EXACTLY, and rows 30 plus 40 must equal row 41. Both are '
              'verified by the bridge-closure checks in tools/audit.py.',
        conf='Low - the attribution is indicative; the total it attributes is the model\'s own output',
        freq='Live'))
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
    # Power and logistics CARVED OUT of conversion cost, not added to it: row 51 is reduced
    # by exactly these two amounts so rows 46-51 still sum to cash cost on row 52.
    w.link('B49', "='Raw Material Forecast'!$D$52", RS)
    w.est('K50', CONVERSION_SPLIT[1][2], PCT0)
    w.f('B50', "='Cost Curve'!$C$80*(1-'Model Assumptions'!$D$20)*$K$50", RS)
    w.f('B51', "='Cost Curve'!$C$80*(1-'Model Assumptions'!$D$20)-B49-B50", RS)
    rows.append(est_row(
        'Cost floor - power and logistics carve-out', 'B49:B50',
        CONVERSION_SPLIT[0][3] + ' ' + CONVERSION_SPLIT[1][3],
        CONVERSION_SPLIT_REASONING,
        unit='Rs/t', linked='Raw Material Forecast, Cost Curve, Model Assumptions', value='=B49', numfmt=RS,
        method='Power read from the bottom-up build-up on 10 Raw Material Forecast; logistics at the share '
               'of conversion cost in K50; row 51 reduced by both so the column still reconciles',
        formula="='Raw Material Forecast'!$D$52 for power; conversion cost x $K$50 for logistics",
        primary='Power: published consumption coefficient at a sourced industrial tariff. Logistics: '
                'modelled share of conversion cost.',
        secondary='No Indian producer discloses either per tonne of crude steel',
        cross='ROWS 46 + 47 + 48 + 49 + 50 + 51 MUST STILL EQUAL ROW 52, the calibrated cash cost, and row '
              '52 plus row 53 must equal row 54. Verified by tools/audit.py.',
        conf='Medium for power, Low for logistics',
        freq='Annually'))
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
    # China: sourced FY2026A print, then the India-China spread held constant in percentage
    # terms, because the model takes no view on Chinese supply-side policy.
    for i, nc in enumerate('CDE'):
        w.f(f'{nc}60', f'=$B$60*{nc}59/$B$59', RS)
    for r, label, ratio, basis in REGION_RATIO:
        w.est(f'K{r}', ratio, '0.00')
        w.f(f'B{r}', f'=$B$59*$K${r}', RS)
        for nc in 'CDE':
            w.f(f'{nc}{r}', f'={nc}59*$K${r}', RS)
    rows.append(est_row(
        'International price comparison - China forecast, Japan, Europe and USA', 'B61:E63, C60:E60',
        REGION_BASIS + ' ' + ' '.join(b for _r, _l, _x, b in REGION_RATIO),
        'Sixteen blank cells in a five-row comparison table, which left only the India row and a single '
        'China data point - not enough to compare anything. Every regional row is now a RATIO to the Indian '
        'price, so the table is complete without implying any forecast of foreign supply, and the ratio for '
        'each region sits in one visible cell in column K. The China row is treated differently and better: '
        'its FY2026A price is sourced, and the forecast holds the India-China spread constant in percentage '
        'terms rather than asserting a Chinese price path. The point of the block is to show that India is '
        'one of the lowest-priced large steel markets in the world - which is the context for both the '
        'safeguard duty and the import-parity ceiling on domestic realisation - and it can now do that.',
        unit='Rs/t', linked='This sheet', value='=B63', numfmt=RS,
        method='Indian domestic price multiplied by the regional ratio in column K; China holds the '
               'observed spread constant in percentage terms',
        formula='=$B$59*$K$61 for Japan; =$B$60*C59/$B$59 for the China forecast',
        primary='China FY2026A IS sourced (IDBI Capital regional snapshot). Japan, Europe and USA ratios '
                'are modelled.',
        secondary='Master Industry Database also carries South East Asia Rs 50,938/t and CIS Rs 51,410/t at '
                  '22-Jun-2026, which bracket the Indian price and corroborate the ordering',
        cross='The India-China spread must remain about Rs 9,000-9,600/t in FY2026A, matching Master '
              'Industry Database KPI K19',
        conf='Medium for China, Low for Japan, Europe and the USA',
        freq='Weekly for the assessments; annually for the ratios'))
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
    for r, label, value, basis in PRICE_ELASTICITY:
        if r == 82:
            continue                     # row 82 IS the model's driver and is already linked
        w.est(f'B{r}', value, X2)
    rows.append(est_row(
        'Price elasticity table', 'B79:B81, B83:B86',
        ' '.join(f'{label}: {b}' for _r, label, _v, b in PRICE_ELASTICITY),
        PRICE_ELASTICITY_REASONING + ' The elasticities stated here are the SAME parameters used by the '
        'price-bridge attribution on rows 31 to 39, so the two tables cannot disagree - which is the reason '
        'to state them once and reference them, rather than to type them twice.',
        unit='x', linked='Model Assumptions, this sheet', value='=B82', numfmt=X2,
        method='Stated elasticities; row 82 links to driver D23, the one elasticity the model uses',
        formula="='Model Assumptions'!$E$30 for row 82; entered values elsewhere",
        primary='Row 82 is driver D23. Every other row is a modelled parameter.',
        secondary='The elasticities of EBITDA to each of these variables ARE computed independently, on 22 '
                  'Sensitivity Analysis section A, which is the place to look for magnitudes that matter',
        cross='The GDP and demand rows are deliberately consistent with the product of the demand '
              'elasticity and the utilisation elasticity, rather than independent assertions. Row 82 must '
              'equal driver D23 exactly.',
        conf='High for row 82, which is the model driver; Low for the rest, which are presentational',
        freq='Quarterly review'))
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
    # ---- rows 11-15: every other input carried as a RATIO to a sourced benchmark, or
    # indexed to the sourced CPI path. Ratios and anchors live in column K.
    pci, pci_basis = RM_RATIOS['pci']
    scrap, scrap_basis = RM_RATIOS['scrap']
    w.est('K11', pci, '0.00')
    w.est('K13', scrap, '0.00')
    for nc, lc in zip(NEW8, LEG):
        w.f(f'{nc}11', f'={nc}10*$K$11', USD)
        w.f(f'{nc}13', f'={nc}10*$K$13', USD)
    for key, row, fmt in (('limestone', 14, RS), ('ferro', 15, RS), ('thermal', 12, USD)):
        anchor, basis = RM_INDEXED[key]
        if row == 15:
            w.db('B15', anchor, fmt,
                 ref='Ferro chrome Rs 1,23,200/t - IDBI Capital input-cost deck, week ended 22-Jun-2026 '
                     '(Master Industry Database source S21)')
        elif row == 12:
            w.db('B12', anchor, fmt, ref="='[Master Industry Database.xlsx]Coking Coal Prices'!$I$28")
        else:
            w.est(f'B{row}', anchor, fmt)
        for i, nc in enumerate('CDEFGHI'):
            prev = NEW8[i]          # NEW8 = 'BCDEFGHI', so prev is the column to the left
            w.f(f'{nc}{row}', f"={prev}{row}*(1+'Model Assumptions'!{MACOL[i]}$11)", fmt)
    rows.append(est_row(
        'PCI coal and scrap price series', 'B11:I11, B13:I13', pci_basis + ' ' + scrap_basis,
        'Both rows were blank across every year. They are now RATIOS to the premium hard coking coal price '
        'the model already forecasts, with the ratio in column K. That construction matters: PCI coal and '
        'scrap are economic substitutes for coking coal, so tying them to it is not a convenience - it is '
        'the actual relationship, and it means a coal-price scenario moves all three together instead of '
        'leaving two of them frozen.',
        unit='US$/t', linked='This sheet', value='=I11', numfmt=USD,
        method='Premium HCC coking coal price multiplied by the ratio in column K',
        formula='=B10*$K$11',
        primary='Modelled ratio to the sourced Ministry of Steel coking coal assessment',
        secondary='Master Industry Database Coking Coal Prices, Conflict C01 records the assessor spread',
        cross='Moves with the coking coal forecast on row 10, so the ratio is preserved in every scenario',
        conf='Low - the ratio is a modelled mid-band estimate',
        freq='Quarterly'))
    rows.append(est_row(
        'Limestone, ferro alloy and thermal coal price series', 'B12:I12, B14:I15',
        RM_INDEXED['limestone'][1] + ' ' + RM_INDEXED['ferro'][1] + ' ' + RM_INDEXED['thermal'][1],
        'Three rows blank across the forecast years. Each is now indexed forward to the sourced RBI CPI '
        'path from its FY2026A anchor. Two of the three anchors ARE sourced - ferro chrome and thermal coal '
        'both come from the Master Industry Database - so only the forward path is modelled. Indexing a '
        'domestic, non-traded, freight-and-royalty-priced input to domestic inflation is the standard '
        'treatment and is far more defensible than inventing a commodity path for limestone.',
        unit='US$/t and Rs/t', linked='Model Assumptions', value='=I14', numfmt=RS,
        method='FY2026A anchor indexed forward at the RBI CPI path (driver D04)',
        formula="=B14*(1+'Model Assumptions'!E$11)",
        primary='Sourced FY2026A anchors for ferro alloys and thermal coal; modelled anchor for limestone',
        secondary='Consistent with the conversion-cost treatment, which is also CPI-indexed',
        cross='Thermal coal remains a REFERENCE series and drives nothing; the database warns against '
              'using it as a coking coal proxy and this model does not',
        conf='Medium for the sourced anchors, Low for the forward paths',
        freq='Quarterly'))

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
    # rows 36-41 mirror the executive summary series, which are now all populated
    FC_SRC = {36: 11, 37: 12, 38: 13, 39: 14, 40: 14, 41: 15}
    FC_FMT = {36: USD, 37: USD, 38: USD, 39: RS, 40: RS, 41: RS}
    for r, src in FC_SRC.items():
        for i, nc in enumerate(NEW7):
            top = 'CDEFGHI'[i]
            if r == 40:      # dolomite: priced off limestone
                w.f(f'{nc}{r}', f'={top}{src}*$K$40', FC_FMT[r])
            else:
                w.link(f'{nc}{r}', f'={top}{src}', FC_FMT[r])
    w.est('K40', 1.15, '0.00')
    rows.append(est_row(
        'Forecast PCI coal, thermal coal, scrap, limestone, dolomite and ferro alloys', 'B36:H41',
        'Each row reads the corresponding series in the executive summary above, which is itself either a '
        'ratio to a sourced benchmark or a CPI-indexed anchor. Dolomite is priced at 115% of limestone, in '
        'column K: both are domestically mined bulk fluxes on similar freight economics, and dolomite '
        'carries a modest premium for the magnesia content that makes it useful as a converter flux.',
        'Six rows across seven years, blank, on the reasoning that asserting these paths "would add '
        'apparent precision without adding information". The concern was right - and the answer is not to '
        'leave the table empty but to make each row a visible RATIO or INDEX rather than an assertion. '
        'These rows now carry no more information than the two sourced benchmarks and the CPI path already '
        'contain, which is exactly the point: the table is complete, and it adds no false precision because '
        'every number in it is a stated transformation of something the model already had.',
        unit='US$/t and Rs/t', linked='This sheet, Model Assumptions', value='=H36', numfmt=USD,
        method='Cross-reference to the executive summary series above; dolomite at a stated ratio to '
               'limestone',
        formula='=C11 for PCI coal; =C14*$K$40 for dolomite',
        primary='Derived from the two sourced price benchmarks and the sourced CPI path',
        secondary='Master Industry Database Iron Ore Prices and Coking Coal Prices',
        cross='Must equal the executive summary rows 11 to 15 exactly',
        conf='Low', freq='Quarterly'))

    # ---- cost build-up rows 46-53
    w.link('C46', '=$D$89', RS)
    w.link('C47', '=$D$90', RS)
    w.f('D46', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*'Model Assumptions'!$E$21", RS)
    w.f('D47', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*'Model Assumptions'!$E$21", RS)
    w.f('D49', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20*(1-2*'Model Assumptions'!$E$21)", RS)
    w.f('D53', "='Cost Curve'!$D$80*'Model Assumptions'!$E$20", RS)
    # ---- bottom-up build-up: consumption coefficient x forecast price = cost per tonne.
    # THIS IS A CROSS-CHECK, NOT AN INPUT. The model still calibrates cash cost top-down
    # from observed realisation less observed EBITDA per tonne.
    elec, elec_basis = ELECTRODE_PRICE
    pwr, pwr_basis = POWER_PRICE
    for r, label, coeff, unit_, _basis in CONSUMPTION:
        w.est(f'B{r}', coeff, '0.000')
    # prices in Rs/t for each input, FY2027E basis (column C of the forecast block)
    w.f('C46', "=$C$34*'Model Assumptions'!$E$10", RS)        # iron ore US$/dmt -> Rs/t
    w.f('C47', "=$C$35*'Model Assumptions'!$E$10", RS)        # premium HCC US$/t -> Rs/t
    w.f('C48', "=$C$11*'Model Assumptions'!$E$10", RS)        # PCI coal
    w.link('C49', '=$C$14', RS)                               # fluxes: limestone
    w.link('C50', '=$C$15', RS)                               # ferro alloys
    w.est('C51', elec, RS)                                    # graphite electrodes
    w.est('C52', pwr, RS)                                     # power, Rs/MWh
    for r, label, coeff, unit_, _basis in CONSUMPTION:
        w.f(f'D{r}', f'=B{r}*C{r}', RS)
    w.f('D53', '=SUM(D46:D52)', RS)
    w.f('B53', '=SUM(B46:B52)', '0.000')
    w.f('C53', '=IFERROR(D53/B53,"")', RS)
    rows.append(est_row(
        'Bottom-up raw material cost build-up', 'B46:D53',
        ' '.join(b for _r, _l, _c, _u, b in CONSUMPTION) + ' ' + elec_basis + ' ' + pwr_basis,
        CONSUMPTION_REASONING,
        unit='t/t and Rs/t', linked='This sheet, Model Assumptions', value='=D53', numfmt=RS,
        method='Specific consumption coefficient multiplied by the forecast price of each input, summed',
        formula='=B46*C46 per line; =SUM(D46:D52) for the total',
        primary='worldsheet raw-material norms for ore and coke; sourced prices for every input',
        secondary='Indian coking coal demand of about 80 Mt on 152 Mt of crude steel independently '
                  'corroborates the coal coefficient',
        cross='COMPARE ROW 53 AGAINST THE TOP-DOWN NUMBER on row 16 of this sheet, which is cash cost from '
              '11 Cost Curve multiplied by the raw material share. The two are built from completely '
              'different evidence - one from engineering coefficients and commodity prices, the other from '
              'company realisation and EBITDA disclosure - so agreement between them is meaningful and a '
              'large gap is a signal to investigate. This comparison was impossible before.',
        conf='Medium for the coefficients, which are published engineering norms; Low for the electrode and '
             'power prices',
        freq='Annually for the coefficients, quarterly for the prices'))
    # ---- import dependency rows 58-61
    imp_na = ('Material-level import dependency is published by the Ministry of Mines and DGCIS but is not '
              'carried in the Master Industry Database, which covers finished STEEL trade only. The steel '
              'trade shares that ARE available are on 19 Trade Model: imports Korea 35.4%, China 23.5% and '
              'Japan 20.2% of FY2026 volumes. Asserting an iron ore, coking coal, scrap or ferro alloy '
              'import share without a source would be an invention.')
    for r, label, dom, countries, _basis in IMPORT_DEP:
        w.est(f'B{r}', dom, PCT0)
        w.f(f'C{r}', f'=1-B{r}', PCT0)
        w.txt(f'D{r}', countries)
    rows.append(est_row(
        'Raw material import dependency and source countries', 'B58:D61',
        ' '.join(b for _r, _l, _d, _c, b in IMPORT_DEP),
        'Twelve blank cells, on the grounds that material-level import dependency is not carried in the '
        'Master Industry Database. True, but the direction and rough magnitude of these four dependencies '
        'are not in doubt and they matter to the model: India is essentially self-sufficient in iron ore and '
        'imports the large majority of its coking coal, and THAT ASYMMETRY IS THE REASON USD/INR is the '
        'second largest driver of industry EBITDA in this workbook. Leaving the table blank hid the '
        'single most important structural fact about the industry\'s cost base. The imported share is a '
        'formula, one less the domestic share, so the two columns cannot fail to sum to 100%.',
        unit='%', linked='This sheet', value='=C59', numfmt=PCT0,
        method='Domestic share entered; imported share computed as one less the domestic share',
        formula='=1-B59',
        primary='Modelled: Ministry of Coal has recorded that domestic coking coal supply is insufficient '
                'to meet demand; India is a net exporter of iron ore and ferro chrome',
        secondary='Ministry of Mines and DGCIS publish the exact shares; 19 Trade Model carries finished '
                  'steel trade shares',
        cross='Columns B and C must sum to 100% on every row. The coking coal dependency must be consistent '
              'with the USD/INR sensitivity on 22 Sensitivity Analysis, which ranks it second',
        conf='Medium for coking coal and iron ore, where the direction is unambiguous; Low for scrap and '
             'ferro alloys',
        freq='Annually, on DGCIS trade data'))

    # ---- scenario analysis rows 66-70
    for col, sc in zip('BCDE', SCEN):
        w.link(f'{col}66', eng.ref(sc, 'ore', 7), USD)
        w.link(f'{col}67', eng.ref(sc, 'coal', 7), NUM0)
    for col, sc in zip('BCDE', SCEN):
        # scrap responds through its ratio to coking coal; thermal coal and ferro alloys
        # respond through the CPI-driven conversion-cost index, which IS a scenario driver
        w.link(f'{col}68', f'={eng.cell(sc, "coal", 7)}*$K$13', USD)
        w.link(f'{col}69', f'=$B$12*{eng.cell(sc, "convidx", 7)}', USD)
        w.link(f'{col}70', f'=$B$15*{eng.cell(sc, "convidx", 7)}', RS)
    rows.append(est_row(
        'Four-scenario scrap, thermal coal and ferro alloy prices, FY2033E', 'B68:E70',
        'Scrap responds through its stated ratio to the scenario coking coal price. Thermal coal and ferro '
        'alloys respond through the conversion-cost index, which is driven by the scenario CPI path.',
        'Twelve blank cells, on the reasoning that these three "are not scenario drivers in this model". '
        'That was true of the old construction and is no longer true of the new one: scrap is now a ratio to '
        'coking coal, which IS a scenario driver, and the two indexed inputs run off CPI, which is also a '
        'scenario driver. So all three now have a genuine, derived scenario response rather than a blank - '
        'and none of them required a new scenario assumption to be invented, which is the test of whether a '
        'scenario table is honest.',
        unit='US$/t and Rs/t', linked='Scenario Manager', value='=B68', numfmt=USD,
        method='Scenario engine coking coal multiplied by the scrap ratio; FY2026A anchors multiplied by '
               'the scenario conversion-cost index',
        formula="='Scenario Manager'!$J$<coal>*$K$13",
        primary='Derived from the scenario driver matrix on 02 Model Assumptions',
        secondary='Consistent with the live rows 11 to 15 on this sheet',
        cross='The base-case column must equal the live FY2033E value of each series',
        conf='Low', freq='Live'))
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
    # Both are exactly zero in this model, and a zero is the correct entry: it states the
    # limitation on the sheet instead of leaving the reader to wonder.
    w.inp('B77', 0, RS)
    w.inp('B79', 0, RS)
    rows.append(est_row(
        'Raw material cost sensitivity to scrap and freight - structural zeros', 'B77, B79',
        'Both are zero because neither input is separately identified in the cost structure. Scrap does not '
        'appear in the raw material basket, which is iron ore, coking coal and a residual other-ferrous '
        'weight. Freight sits inside conversion cost, which is indexed to RBI CPI as a single block.',
        'These two cells were blank and flagged as missing data, which misrepresented them: the impact is '
        'not unknown, it is ZERO, and the zero is itself the disclosure. Writing it in makes the limitation '
        'visible in the table where a reader is looking for it. Both zeros are real limitations and worth '
        'stating plainly: the scrap and DRI-electric routes are a growing share of Indian capacity and this '
        'model does not cost them separately, and a freight-specific shock would be understated because '
        'freight cannot be separated from conversion cost.',
        unit='Rs/t', linked='This sheet', value='=B77', numfmt=RS,
        method='Structural zero - the input is not separately identified in the cost build',
        formula='(entered zero)',
        primary='Consequence of the model\'s cost structure, not an estimate',
        secondary='The bottom-up build-up on rows 46 to 53 shows which inputs ARE separately costed',
        cross='Consistent with the basket weights: driver D14 splits raw material between iron ore, coking '
              'coal and a residual, with no scrap line',
        conf='High - the zero is exact given the model\'s structure',
        freq='n/a - changes only if scrap is added to the cost build'))
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
