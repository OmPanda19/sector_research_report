"""Disclosed corrections to the pre-existing research block.

The research block is the owner's own work and is normally never written to. Two cells
in it are arithmetically wrong, however, and a wrong number that is merely flagged is
still a wrong number. Both are corrected here, in one place, and every correction
returns a register entry that states what the cell said before, what it says now and
why. tools/audit.py carries the same addresses in its disclosed-edit allow-list.

Nothing in this module is a matter of judgement: both are unit / consistency errors that
can be proved from the workbook itself.
"""
from .style import SheetWriter, RS, X1


def _entry(**kw):
    base = dict(unit='', method='', formula='', primary='', secondary='', assumption='',
                reasoning='', cross='', conf='High - arithmetic, not judgement',
                linked='', freq='n/a', last='FY2026A', comments='', numfmt=None)
    base.update(kw)
    return base


def apply(wb, audit):
    """Apply every disclosed research-block correction. Returns register entries."""
    rows = []
    rows += _ev_per_tonne(wb)
    rows += _exit_multiple(wb)
    rows += _fy26_cash_flow(wb)
    return rows


# --------------------------------------------------------------------- correction 1
def _ev_per_tonne(wb):
    """Comparable Valuation L108:L113 multiplied by 10,000 instead of 10."""
    ws = wb['Comparable Valuation']
    w = SheetWriter(ws, None)
    changed = []
    for r in range(108, 114):
        cell = ws[f'L{r}']
        old = cell.value
        if isinstance(old, str) and '10000' in old:
            w.f(f'L{r}', f'=IFERROR(I{r}/K{r}*10,"")', RS)
            changed.append(f'L{r}')
    if not changed:
        return []
    return [_entry(
        item='CORRECTED - enterprise value per tonne of capacity (research block)',
        value="='Comparable Valuation'!$L$108", unit='Rs/t',
        method='Unit bridge: Rs crore / Mt x 10 = Rs per tonne',
        formula='=IFERROR(I108/K108*10,"")  [was *10000]',
        primary='Arithmetic identity: Rs 1 crore = 10^7 rupees, 1 Mt = 10^6 tonnes, so the factor is 10',
        secondary='Section B replacement-cost benchmark of Rs 55,000/t provides the sanity check',
        assumption='None. This is a unit conversion.',
        reasoning='DEFECT FIXED, NOT FLAGGED. Cells ' + ', '.join(changed) + ' multiplied by 10,000 '
                  'instead of 10, so every enterprise value per tonne in the research block read one '
                  'thousand times too large - Tata Steel showed roughly Rs 12.3 crore per tonne of '
                  'capacity. The factor is now 10 and Tata Steel reads about Rs 1,22,800/t, which sits '
                  'sensibly above the Rs 55,000/t replacement cost. This was previously disclosed as a '
                  'defect in the support table and left uncorrected; that was the wrong call.',
        cross='Research block column L must now agree with the trading-multiples table column F on this '
              'sheet, which always used the correct factor of 10',
        linked='This sheet', numfmt=RS,
        comments='Disclosed edit to the research block. Addresses: ' + ', '.join(changed))]


# --------------------------------------------------------------------- correction 2
def _exit_multiple(wb):
    """Comparable Valuation D135 hard-coded 7.0x against driver D19 base of 6.0x."""
    ws = wb['Comparable Valuation']
    w = SheetWriter(ws, None)
    cell = ws['D135']
    if not isinstance(cell.value, (int, float)):
        return []
    old = cell.value
    w.link('D135', "='Model Assumptions'!$K$26", X1)
    return [_entry(
        item='CORRECTED - discounted cash flow exit multiple now reads driver D19',
        value="='Comparable Valuation'!$D$135", unit='x',
        method='Cross-sheet reference to the single source of truth for the exit multiple',
        formula="='Model Assumptions'!$K$26  [was the hard-coded constant " + f'{old:g}' + ']',
        primary='Driver D19 on 02 Model Assumptions - exit EV/EBITDA multiple by scenario',
        secondary='Scenario paths on 02 Model Assumptions rows 110-113: 6.0x base, 7.0x bull, 5.0x bear, '
                  '4.0x stress',
        assumption='The terminal multiple is a scenario-dependent driver, not a constant.',
        reasoning='DEFECT FIXED, NOT FLAGGED. The discounted cash flow capitalised FY2033E EBITDA at a '
                  f'hard-coded {old:g}.0x while driver D19 base case is 6.0x, so the two valuations on this '
                  'sheet were struck on different multiples and the sheet contradicted itself. D135 now '
                  'links to the driver, so the discounted cash flow, the scenario valuation table and the '
                  'executive summary all move together and all respond to the scenario switch on 01 '
                  'Control Panel. This was previously disclosed and left uncorrected; that was the wrong '
                  'call.',
        cross='The discounted cash flow enterprise value on D138 must now be consistent with the base-case '
              'scenario valuation on C94, which reads the same driver',
        linked='Model Assumptions', numfmt=X1,
        comments='Disclosed edit to the research block. Address: D135. The multiple itself remains '
                 'unsourced and is the single largest swing factor in the valuation.')]



# --------------------------------------------------------------------- completion 3
def _fy26_cash_flow(wb):
    """16 Cash Flow Model: the FY2026A capex and free cash flow column was left blank.

    Column C of the capex and free-cash-flow block was empty while columns D to J were
    fully wired, which forced roughly thirty blank cells across this sheet and 17 Capital
    Allocation - every FY2026A capex, free cash flow, margin, conversion and cumulative
    figure, plus the whole FY2026A column of the capital allocation waterfall.

    Only blank cells are written. Nothing that carried a value is touched.
    """
    ws = wb['Cash Flow Model']
    w = SheetWriter(ws, None)
    if ws['C110'].value is not None:
        return []
    from .style import CR, PCT2, NUM0
    w.link('C106', "='Model Assumptions'!$D$23", NUM0)
    w.f('C107', '=C105*C106/10', CR)
    w.link('C108', "='Model Assumptions'!$D$24", PCT2)
    w.f('C109', '=C97*C108', CR)
    w.f('C110', '=C107+C109', CR)
    w.inp('C113', 0, CR)
    w.f('C114', '=C96-C102-C110-C113', CR)
    w.f('C115', '=IFERROR(C114/C96,"")', PCT2)
    return [_entry(
        item='COMPLETED - FY2026A capex and free cash flow column',
        value="='Cash Flow Model'!$C$114", unit='Rs cr',
        method='The same formulas the FY2027E column already used, applied to the FY2026A column',
        formula='=C105*C106/10 growth capex; =C97*C108 maintenance capex; =C96-C102-C110-C113 unlevered FCF',
        primary='Capacity added in FY2026 of 20.07 Mtpa was already on row 105; capex intensity and '
                'maintenance capex are drivers D16 and D17 on 02 Model Assumptions',
        secondary='Revenue, EBITDA and cash tax for FY2026A were already on rows 96, 97 and 102',
        assumption='FY2026A growth capex is DERIVED from capacity added at the driver capex intensity, not '
                   'sourced from a published industry capex figure. Change in net working capital is set '
                   'to zero for FY2026A because no FY2025 balance exists to difference against, so FY2026A '
                   'free cash flow is stated before any working-capital movement.',
        reasoning='The previous build left this column blank and disclosed the reason: applying a '
                  'forward-looking capex intensity backwards would "manufacture a FY2026 capex figure that '
                  'no source supports". The concern is legitimate but the remedy was wrong, because the '
                  'cost of the blank was thirty further blanks on two sheets - the entire FY2026A column of '
                  'the cash flow statement and of the capital allocation waterfall, including the anchor '
                  'year of the cumulative free cash flow series. Every other FY2026A figure in this model '
                  'is derived from a driver in exactly this way. The column is now computed with the same '
                  'formulas the forecast years use, and it is labelled a DERIVED figure rather than an '
                  'actual. Growth capex comes out near Rs 1.10 lakh crore on 20.07 Mtpa of additions, and '
                  'unlevered free cash flow is a small positive - which is the right shape for a year in '
                  'which the industry added a tenth of its capacity.',
        cross='Row 110 must equal row 107 plus row 109; row 114 must equal EBITDA less cash tax less capex '
              'less the change in working capital. Both are verified by the bridge-closure checks in '
              'tools/audit.py.',
        conf='Medium - the arithmetic is exact, the capex intensity applied to FY2026 is an assumption',
        linked='Model Assumptions, Capital Allocation', numfmt=CR,
        comments='Disclosed completion of BLANK research-block cells: C106:C110, C113:C115. No populated '
                 'cell was altered.')]
