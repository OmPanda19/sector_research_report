"""02 Model Assumptions - complete the FY2026A column of the ACTIVE DRIVERS block.

The scenario columns (FY2027E-FY2033E) already resolve by ScenID. Column D was the
only gap. Every FY2026A entry is either a live link to the cell in the workbook that
already asserts that actual - so there is still exactly one source of truth - or a
structural parameter that is held constant across years and scenarios and therefore
applies to FY2026A unchanged, or an ORANGE cell where FY2026A does not exist as a
concept.
"""
from .style import SheetWriter, PCT1, PCT2, NUM1, NUM2, NUM3, X1, X2, RS, MT, MTS, NUM0

SHEET = 'Model Assumptions'

# row, kind, payload, number format, note for the audit table
DRIVERS = [
    (8, 'link', "='Macroeconomic Model'!$C$101", PCT1,
     "RBI's own estimate of FY2026 real GDP growth, 7.6%, with Q4FY2026 actual 7.7%."),
    (9, 'link', "='Steel Demand Model'!$C$114", X2,
     'Realised FY2015-FY2026 elasticity of 1.10x - apparent consumption compounded at 7.13% '
     'against real GDP over the same window.'),
    (10, 'link', "='Macroeconomic Model'!$C$106", NUM1,
     "FY2026 average USD/INR of 88.4, derived from Tata Steel's dual-currency disclosure."),
    (11, 'link', "='Macroeconomic Model'!$C$103", PCT1,
     'CPI as carried on 03 Macroeconomic Model. Sourced from the RBI MPC statement of 05-Jun-2026.'),
    (12, 'na', 'The FY2026 announced pipeline is not observable. Every project in the dated tracker with a '
               'FY2026 commissioning date is already inside the 220.4 Mtpa base, so there is no FY2026 '
               'denominator against which a delivery factor could be struck. The parameter is forward-looking '
               'only.', PCT1, ''),
    (13, 'link', "='Steel Supply Model'!$C$93", '0.0000"x"',
     '168.42 Mt crude divided by 160.94 Mt finished on the April-2026 Joint Plant Committee vintage.'),
    (14, 'link', "='Steel Supply Model'!$C$90", MTS,
     'FY2026 actual net exports of +0.078 Mt: exports 6.602 Mt less imports 6.524 Mt.'),
    (15, 'link', "='Steel Supply Model'!$C$91", MTS,
     'FY2026 actual variation in stock of -2.88 Mt, a destock of about 1.8% of consumption.'),
    (16, 'link', "='Steel Price Forecast'!$C$97", RS,
     'Rs 59,974/t derived as the volume-weighted realisation of the four majors that disclose both '
     'realisation and EBITDA per tonne.'),
    (17, 'link', "='Raw Material Forecast'!$C$85", NUM1,
     'FY2026 average of US$100.52/dmt, the arithmetic mean of twelve World Bank monthly prints.'),
    (18, 'link', "='Raw Material Forecast'!$C$86", NUM0,
     "US$225/t, the Ministry of Steel's March-2026 assessment. No fiscal-year average is asserted "
     '(Conflict C01 in the Master Industry Database).'),
    (19, 'na', 'No Indian producer discloses a conversion-cost inflation rate, and no official series measures '
               'it. The parameter is anchored on RBI CPI for the forecast years only; asserting a FY2026 '
               'outturn would be an invention.', PCT1, ''),
    (20, 'inp', 0.60, '0%',
     'INDICATIVE STRUCTURAL PARAMETER, held constant across all four scenarios and all years precisely so '
     'that an unsourced input is never flexed to manufacture a scenario. It therefore applies to FY2026A '
     'unchanged. Quantified on 22 Sensitivity Analysis section E.'),
    (21, 'inp', 0.45, '0%',
     'INDICATIVE. Iron ore 45%, coking coal 45%, other ferrous and fluxes 10%. Held constant, so it '
     'applies to FY2026A unchanged.'),
    (22, 'inp', 0.2517, PCT2,
     'Indian statutory corporate rate under section 115BAA - 22% plus surcharge and cess - which was in '
     'force in FY2026 and is held across the horizon.'),
    (23, 'inp', 55000, RS,
     'Greenfield and large brownfield Indian crude steel capacity. Cross-checked against JSW JVML at '
     'Rs 52,000/t for 5 Mtpa and Tata NINL. Held constant, so it applies to FY2026A unchanged.'),
    (24, 'inp', 0.03, PCT1,
     'INDICATIVE sustaining capex excluding growth. Held constant across scenarios and years.'),
    (25, 'inp', 45, NUM0,
     'NOT SOURCED. A sector convention for an integrated Indian producer. The live model already uses 45 '
     "days as the FY2026A anchor on 15 Working Capital Model, so this restates the model's own base."),
    (26, 'na', 'An exit EV/EBITDA multiple is a terminal-value parameter applied to FY2033E. It has no '
               'FY2026 actual. A FY2026 TRADED multiple would be a different quantity and would require '
               'share prices, which were not sourced - see standing limitation 6 on 25 Audit Checks.',
     X1, ''),
    (27, 'link', '=$D$168', PCT1,
     'Live from the WACC build-up in section D below, so driver D20 and the build-up can never disagree. '
     'Check A20 on 25 Audit Checks tests exactly this.'),
    (28, 'link', "='Trade Model'!$C$74", NUM2,
     'FY2026 actual finished steel imports of 6.524 Mt, down 31.7% from 9.551 Mt in FY2025.'),
    (29, 'link', "='Capacity Forecast'!$C$100", NUM2,
     'FY2026 actual capacity addition of 20.07 Mtpa. All of it is unattributed for model purposes, because '
     'every dated FY2026 project is already inside the 220.4 Mtpa base.'),
    (30, 'inp', 0.40, X2,
     'INDICATIVE. The lagged price-utilisation feedback elasticity. Held constant across all four '
     'scenarios so that the scenarios differ on evidence, not on this parameter.'),
    (31, 'inp', 0.80, '0%',
     'Anchored on the FY2024 actual utilisation of 80.4%, the last year before capacity outran production. '
     'Held constant.'),
    (32, 'inp', 0.92, '0%',
     'INDICATIVE. Allows for maintenance and relining. Individual assets run above this; a national '
     'aggregate does not.'),
    (33, 'link', "='Cash Flow Model'!$C$98", PCT2,
     'FY2026A of 5.28%, derived as Rs 30,715 cr of depreciation and amortisation on Rs 5,81,000 cr of '
     'aggregate revenue for the disclosing names.'),
]


def fill(wb, audit):
    ws = wb[SHEET]
    w = SheetWriter(ws, audit)
    rows = []
    for row, kind, payload, fmt, note in DRIVERS:
        addr = f'D{row}'
        did = ws[f'A{row}'].value
        item = ws[f'B{row}'].value
        unit = ws[f'C{row}'].value
        if kind == 'link':
            w.link(addr, payload, fmt)
            method = 'Cross-sheet reference to the cell that asserts the actual'
            formula = payload
            conf = ws[f'M{row}'].value or 'High'
        elif kind == 'inp':
            w.inp(addr, payload, fmt)
            method = 'Manual input - structural parameter held constant across years and scenarios'
            formula = f'(constant) {payload}'
            conf = ws[f'M{row}'].value or 'Low'
        else:
            w.na(addr, payload)
            method = 'NOT POPULATED - orange fill'
            formula = '(blank by design)'
            conf = 'n/a - not applicable to FY2026A'
        rows.append(dict(
            item=f'{did} {item} - FY2026A anchor', value=f'=D{row}', unit=unit or '',
            method=method, formula=formula,
            primary=ws[f'L{row}'].value or 'See 26 Sources',
            secondary='02 Model Assumptions section C assumption register; Master Industry Database',
            assumption=note if kind != 'na' else 'None asserted. The cell is deliberately blank.',
            reasoning=note if kind == 'na' else (
                'Populated by reference rather than re-keyed, so the FY2026A actual exists in exactly one '
                'place in the workbook.'),
            cross='Driver D01-D26 forecast columns resolve by ScenID; check A28 on 25 Audit Checks counts 182 '
                  'resolved driver cells',
            conf=conf, linked=SHEET, freq='Annually, when the base year rolls forward',
            last='FY2026A', comments=f'ACTIVE DRIVERS row {row}.'))
    return rows
