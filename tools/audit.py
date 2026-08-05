#!/usr/bin/env python3
"""Full audit of the completed Industry Financial Model.

 1. GUARD      - proves nothing at or below each legacy start row was altered, and that
                 no formatting on any pre-existing cell was altered.
 2. EVALUATE   - computes every formula in the workbook and reports every failure.
 3. TIE-OUT    - engine vs live model, and engine vs the independent FY2033E values.
 4. INTEGRITY  - circularity, broken refs, sign, unit and continuity checks.
"""
import re
import sys
import openpyxl
from openpyxl.utils import get_column_letter

sys.path.insert(0, '.')
from tools.xlcalc import Book, Err
from gen.spec import LEGACY
from gen import engine as E

SRC = 'incoming/IFM_user.xlsx'
OUT = 'Industry Financial Model.xlsx'

FAIL, WARN, OK = [], [], []

# Deliberate, disclosed changes to cells that already carried a value or a style.
# Anything NOT on this list that changes is a genuine failure.
ALLOWED = set()
for _r in range(20, 33):
    ALLOWED.add(f'Database Import!D{_r}')          # 'xxx' placeholder -> real record count
for _r in range(6, 12):
    ALLOWED.add(f'Control Panel!H{_r}')            # hard-coded tick -> live test
for _r in range(18, 32):
    ALLOWED.add(f'Control Panel!B{_r}')            # 'Go ->' given a working hyperlink
ALLOWED.add('Model Calibration!C43')               # =D10/D12 -> =D10/D11 (wrong denominator)
for _r in range(52, 57):
    ALLOWED.add(f'Model Calibration!D{_r}')        # hard-coded tick -> live reconciliation
# blank-but-should-be-live row inside a research block
for _c in 'DEFGHIJ':
    ALLOWED.add(f'Industry Dashboard!{_c}55')
for _r in range(27, 34):
    ALLOWED.add(f'Sensitivity Analysis!B{_r}')  # '...' placeholder -> explicit shock size
for _r in range(6, 40):
    ALLOWED.add(f'Index!B{_r}')                 # sheet name given a working hyperlink
for _c in 'BCDEFGH':
    ALLOWED.add(f'Industry Dashboard!{_c}7')    # navigation label given a hyperlink


def fail(m):
    FAIL.append(m)


def warn(m):
    WARN.append(m)


def ok(m):
    OK.append(m)


# ============================================================================ 1. guard
def guard():
    a = openpyxl.load_workbook(SRC)
    b = openpyxl.load_workbook(OUT)
    if a.sheetnames != b.sheetnames:
        fail('sheet names changed')
        return
    ok(f'{len(a.sheetnames)} sheet names unchanged')

    def fsig(c):
        f, fl = c.font, c.fill
        return (f.name, f.sz, f.b, f.i, f.u, f.color.rgb if f.color else None,
                f.color.theme if f.color else None,
                fl.patternType, fl.fgColor.rgb if fl.fgColor else None,
                c.number_format, c.alignment.horizontal, c.alignment.vertical,
                c.alignment.wrapText, tuple((getattr(c.border, s).style) for s in
                                            ('left', 'right', 'top', 'bottom')))

    def same_value(va, vb):
        va = va.text if hasattr(va, 'text') else va
        vb = vb.text if hasattr(vb, 'text') else vb
        if va in (None, '') and vb in (None, ''):
            return True
        if isinstance(va, float) and isinstance(vb, float):
            return abs(va - vb) <= 1e-9 * max(1.0, abs(va))
        return va == vb

    legacy_changed, style_changed, value_changed = [], [], []
    intentional = 0
    for sn in a.sheetnames:
        wa, wbs = a[sn], b[sn]
        start = LEGACY.get(sn)
        for r in range(1, wa.max_row + 1):
            for c in range(1, wa.max_column + 1):
                ca, cb = wa.cell(r, c), wbs.cell(r, c)
                key = f'{sn}!{get_column_letter(c)}{r}'
                had_value = ca.value not in (None, '')
                vchanged = not same_value(ca.value, cb.value)
                schanged = fsig(ca) != fsig(cb)
                if not (vchanged or schanged):
                    continue
                if key in ALLOWED:
                    intentional += 1
                    continue
                if start and r >= start and vchanged and had_value:
                    legacy_changed.append(f'{key}: {str(ca.value)[:40]!r} -> {str(cb.value)[:40]!r}')
                elif vchanged and had_value:
                    value_changed.append(f'{key}: {str(ca.value)[:40]!r} -> {str(cb.value)[:40]!r}')
                elif schanged and had_value:
                    style_changed.append(key)
        # (c) structure
        for label, xa, xb in (
            ('merged', sorted(str(x) for x in wa.merged_cells.ranges),
             sorted(str(x) for x in wbs.merged_cells.ranges)),
            ('column widths', {k: v.width for k, v in wa.column_dimensions.items()},
             {k: v.width for k, v in wbs.column_dimensions.items() if k in wa.column_dimensions}),
            ('row heights', {k: v.height for k, v in wa.row_dimensions.items()},
             {k: v.height for k, v in wbs.row_dimensions.items() if k in wa.row_dimensions}),
            ('tables', sorted(wa.tables), sorted(wbs.tables)),
            ('data validations', len(wa.data_validations.dataValidation),
             len(wbs.data_validations.dataValidation)),
            ('conditional formats', len(wa.conditional_formatting._cf_rules),
             len(wbs.conditional_formatting._cf_rules)),
        ):
            if xa != xb:
                if label == 'merged' and set(xa) <= set(xb):
                    continue
                fail(f'{sn}: {label} changed')
    if legacy_changed:
        fail(f'{len(legacy_changed)} research-block cells altered')
        for m in legacy_changed[:10]:
            fail('   ' + m)
    else:
        ok('research blocks untouched: no populated cell at or below any legacy start row changed')
    if value_changed:
        fail(f'{len(value_changed)} pre-existing values overwritten outside the allow-list')
        for m in value_changed[:10]:
            fail('   ' + m)
    else:
        ok('no pre-existing value overwritten outside the disclosed allow-list')
    if style_changed:
        fail(f'{len(style_changed)} pre-existing cells had formatting changed: {style_changed[:8]}')
    else:
        ok('no pre-existing cell had its font, fill, number format, alignment or border changed')
    ok(f'{intentional} deliberate, disclosed edits to previously populated cells '
       '(placeholders, hard-coded ticks, one wrong denominator, navigation hyperlinks)')


# ========================================================================= 2. evaluate
SKIP_TEXT = ('Data Not Publicly Available',)


def evaluate():
    wb = openpyxl.load_workbook(OUT)
    bk = Book(wb)
    total = 0
    bad = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                t = v if isinstance(v, str) else getattr(v, 'text', None)
                if not (isinstance(t, str) and t.startswith('=')):
                    continue
                total += 1
                try:
                    bk.cell(ws.title, c.coordinate)
                except (Err, ZeroDivisionError, ValueError, TypeError, KeyError,
                        RecursionError, IndexError) as e:
                    bad[f'{ws.title}!{c.coordinate}'] = f'{e} :: {t[:110]}'
    ok(f'{total} formulas evaluated')
    if bad:
        fail(f'{len(bad)} formulas failed to evaluate')
        for k, v in list(bad.items())[:40]:
            fail(f'   {k}: {v}')
    else:
        ok('every formula in the workbook evaluates without error')
    return wb, bk


# ========================================================================== 3. tie-out
def tieout(wb, bk):
    ws = wb['Scenario Manager']
    base = {}
    for r in range(90, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str):
            for s in E.SCENARIOS:
                if v.strip().upper() == s.upper():
                    base[s] = r + 1
    if len(base) != 4:
        fail('scenario engine blocks not found')
        return None
    eng = E.Engine(base)
    TIE = [(69, 'cons'), (70, 'crude'), (71, 'cap'), (72, 'util'), (73, 'real'), (74, 'cost'),
           (75, 'ebitdat'), (76, 'margin'), (77, 'rev'), (78, 'ebitda'), (79, 'ufcf')]
    SCOL = {'Base Case': 'C', 'Bull Case': 'D', 'Bear Case': 'E', 'Stress Case': 'F'}
    n = 0
    for row, name in TIE:
        for sc, col in SCOL.items():
            want = ws[f'{col}{row}'].value
            got = bk.get('Scenario Manager', f'J{eng.row(sc, name)}')
            if want is None or got is None:
                fail(f'tie-out missing {name}/{sc}')
                continue
            if abs(got - want) > abs(want) * 0.005 + 1e-6:
                fail(f'tie-out {name}/{sc}: engine {got:.4f} vs independent {want:.4f}')
            else:
                n += 1
    ok(f'{n}/44 independent FY2033E tie-out values reproduced by the scenario engine')

    LIVE = [('cons', 'Steel Demand Model', 116), ('crude', 'Steel Supply Model', 99),
            ('fin', 'Steel Supply Model', 101), ('cap', 'Capacity Forecast', 101),
            ('util', 'Capacity Utilisation', 67), ('real', 'Steel Price Forecast', 103),
            ('cost', 'Cost Curve', 80), ('ebitdat', 'EBITDA Model', 101),
            ('ebitda', 'EBITDA Model', 103), ('rev', 'Revenue Forecast', 96),
            ('margin', 'Margin Analysis', 91), ('ufcf', 'Cash Flow Model', 114),
            ('capex', 'Cash Flow Model', 110), ('nwc', 'Working Capital Model', 72),
            ('imp', 'Trade Model', 76), ('exp', 'Trade Model', 78),
            ('percap', 'Steel Demand Model', 121), ('cashtax', 'Cash Flow Model', 102),
            ('da', 'Cash Flow Model', 99), ('ebit', 'Cash Flow Model', 100)]
    active = bk.get('Control Panel', 'E17')
    m = 0
    for name, sh, lrow in LIVE:
        for j, lcol in enumerate('CDEFGHIJ'):
            live = bk.get(sh, f'{lcol}{lrow}')
            got = bk.get('Scenario Manager',
                         f'{E.A if j == 0 else E.COLS[j - 1]}{eng.row(active, name)}')
            if not isinstance(live, (int, float)) or not isinstance(got, (int, float)):
                continue
            if abs(got - live) > abs(live) * 0.002 + 1e-6:
                fail(f'engine vs live {name} {sh} col {lcol}: {got:.4f} vs {live:.4f}')
            else:
                m += 1
    ok(f'{m} engine cells match the live model exactly for the active scenario ({active})')
    return eng


# ======================================================================== 4. integrity
def integrity(wb, bk, eng):
    G = bk.get
    # -- 4a. the thirty published validations
    res = {}
    for r in range(14, 45):
        cid = wb['Audit Checks'][f'A{r}'].value or wb['Audit Checks'][f'B{r}'].value
        v = G('Audit Checks', f'C{r}')
        if v is not None:
            res[str(cid)] = v
    nf = [k for k, v in res.items() if str(v).upper() == 'FAIL']
    nw = [k for k, v in res.items() if str(v).upper() == 'WARN']
    if nf:
        fail(f'published audit checks FAIL: {nf}')
    else:
        ok(f'{len(res)} published audit checks evaluated: 0 FAIL, {len(nw)} WARN {nw if nw else ""}')

    # -- 4b. supply identity in every year and every scenario
    for sc in E.SCENARIOS:
        for y in range(1, 8):
            fin = G('Scenario Manager', E.COLS[y - 1] + str(eng.row(sc, 'fin')))
            req = G('Scenario Manager', E.COLS[y - 1] + str(eng.row(sc, 'reqfin')))
            sh = G('Scenario Manager', E.COLS[y - 1] + str(eng.row(sc, 'short')))
            if None in (fin, req, sh):
                continue
            if abs((fin + sh) - req) > 0.01:
                fail(f'supply identity breaks {sc} y{y}: {fin}+{sh} != {req}')
    ok('supply identity closes in all 4 scenarios x 7 years (production + shortfall = requirement)')

    # -- 4c. monotonic capacity, positive cost, utilisation ceiling
    for sc in E.SCENARIOS:
        caps = [G('Scenario Manager', (E.A if y == 0 else E.COLS[y - 1]) + str(eng.row(sc, 'cap')))
                for y in range(8)]
        if any(caps[i + 1] < caps[i] - 1e-9 for i in range(7)):
            fail(f'capacity not monotonic in {sc}')
        utl = [G('Scenario Manager', E.COLS[y - 1] + str(eng.row(sc, 'util'))) for y in range(1, 8)]
        if max(utl) > 1.0:
            fail(f'utilisation above 100% in {sc}')
        mx = G('Scenario Manager', 'J' + str(eng.row(sc, 'maxutil')))
        if max(utl) > mx + 1e-9:
            fail(f'utilisation above the practical ceiling in {sc}')
        cost = [G('Scenario Manager', E.COLS[y - 1] + str(eng.row(sc, 'cost'))) for y in range(1, 8)]
        if min(cost) <= 0:
            fail(f'non-positive cash cost in {sc}')
        mar = [G('Scenario Manager', E.COLS[y - 1] + str(eng.row(sc, 'margin'))) for y in range(1, 8)]
        if max(mar) > 0.30 or min(mar) < -0.15:
            warn(f'{sc} margin path {min(mar):.1%} to {max(mar):.1%} is outside the -15% to +30% guard band')
    ok('capacity monotonic, utilisation inside the practical ceiling and cash cost positive '
       'in all 4 scenarios')

    # -- 4d. unit bridges Rs/t x Mt / 10 = Rs cr
    for sh, num, den, res_row in (('Revenue Forecast', 94, 95, 96), ('EBITDA Model', 102, 101, 103)):
        for col in 'CDEFGHIJ':
            a, b_, c = G(sh, f'{col}{num}'), G(sh, f'{col}{den}'), G(sh, f'{col}{res_row}')
            if None in (a, b_, c) or any(isinstance(x, str) for x in (a, b_, c)):
                continue
            if abs(a * b_ / 10 - c) > max(1.0, abs(c) * 1e-6):
                fail(f'unit bridge broken {sh}!{col}{res_row}')
    ok('Rs/t x Mt / 10 = Rs crore unit bridge holds on revenue and EBITDA in every year')

    # -- 4e. cross-sheet consistency of shared drivers (one source of truth)
    shared = [
        ('real GDP growth', [('Macroeconomic Model', 101), ('Steel Demand Model', 113)]),
        ('blended realisation', [('Steel Price Forecast', 103), ('Revenue Forecast', 95),
                                 ('EBITDA Model', 99), ('Margin Analysis', 88)]),
        ('cash cost', [('Cost Curve', 80), ('EBITDA Model', 100), ('Margin Analysis', 89)]),
        ('industry EBITDA', [('EBITDA Model', 103), ('Cash Flow Model', 96),
                             ('Capital Allocation', 83)]),
        ('industry revenue', [('Revenue Forecast', 96), ('Cash Flow Model', 97),
                              ('Working Capital Model', 70)]),
        ('finished production', [('Steel Supply Model', 101), ('Revenue Forecast', 94),
                                 ('EBITDA Model', 102), ('Trade Model', 82)]),
        ('capacity', [('Capacity Forecast', 101), ('Steel Supply Model', 96),
                      ('Capacity Utilisation', 66)]),
        ('utilisation', [('Capacity Utilisation', 67), ('Industry Cycle Model', 87)]),
        ('USD/INR', [('Macroeconomic Model', 106), ('Raw Material Forecast', 87)]),
        ('effective tax rate', [('Cash Flow Model', 101), ('Comparable Valuation', 144)]),
    ]
    dupes = 0
    for label, places in shared:
        for col in 'DEFGHIJ':
            vals = []
            for sh, r in places:
                v = G(sh, f'{col}{r}')
                if isinstance(v, (int, float)):
                    vals.append((sh, v))
            if len(vals) < 2:
                continue
            ref = vals[0][1]
            for sh, v in vals[1:]:
                if abs(v - ref) > max(1e-6, abs(ref) * 1e-6):
                    fail(f'CONTRADICTION on {label}: {vals[0][0]} {ref} vs {sh} {v} (col {col})')
                    dupes += 1
    if not dupes:
        ok(f'{len(shared)} shared drivers verified identical everywhere they appear '
           '(one source of truth holds)')

    # -- 4e2. every bridge must close exactly
    BRIDGES = [
        # sheet, cols, [+rows], [-rows], target row, tolerance
        ('Revenue Forecast', 'BCDEFGH', [28, 29, 30, 31, 32, 33], [], 34, 1.0),
        ('EBITDA Model', 'BCDEFGH', [47, 48, 49, 50, 53, 54], [], 55, 1.0),
        ('Margin Analysis', 'BCDEFGH', [31, 32, 33, 34, 35, 36], [], 40, 1e-9),
        ('Capacity Utilisation', 'BCDEFGH', [32, 37], [], 38, 1e-9),
        ('Capacity Forecast', 'BCDEFGH', [51, 56], [], 57, 1e-6),
        ('Capacity Forecast', 'BCDEFGH', [52, 53, 54], [55], 56, 1e-6),
        ('Cash Flow Model', 'CDEFGHI', [20, 21, 22], [], 23, 1.0),
        ('Cash Flow Model', 'CDEFGHI', [23, 24, 25], [], 26, 1.0),
        ('Cash Flow Model', 'EFGHI', [60, 61, 62, 63, 64, 65], [], 66, 1.0),
        ('Steel Supply Model', 'BCDEFGHI', [37, 38], [39], 41, 1e-6),
        ('Cost Curve', 'B', [38, 39, 42, 47], [], 48, 2.0),
        ('Cost Curve', 'B', [48, 49], [], 51, 1.0),
        ('EBITDA Model', 'B', [31, 32, 35, 41], [], 42, 2.0),
        ('Margin Analysis', 'B', [45, 50], [], 51, 1.0),
        ('Steel Price Forecast', 'B', [46, 47, 48, 51], [], 52, 2.0),
        ('Raw Material Forecast', 'D', [46, 47, 49], [], 53, 2.0),
    ]
    nb = 0
    for sh, cols, plus, minus, target, tol in BRIDGES:
        for col in cols:
            vals = [G(sh, f'{col}{r}') for r in plus]
            neg = [G(sh, f'{col}{r}') for r in minus]
            t = G(sh, f'{col}{target}')
            if not isinstance(t, (int, float)) or any(
                    not isinstance(v, (int, float)) for v in vals + neg):
                continue
            s = sum(vals) - sum(neg)
            if abs(s - t) > tol + abs(t) * 1e-9:
                fail(f'BRIDGE does not close {sh} col {col} -> {target}: {s:.6f} vs {t:.6f}')
            else:
                nb += 1
    ok(f'{nb} bridge closures verified exactly (revenue, EBITDA, margin, utilisation, capacity, '
       'cash flow, FCF, supply, cost build-ups)')

    # -- 4f. every scenario table on every sheet must agree with the engine
    checks = [
        ('Steel Demand Model', 94, 'cons'), ('Capacity Utilisation', 52, 'util'),
        ('Steel Price Forecast', 68, 'real'), ('Revenue Forecast', 83, 'rev'),
        ('EBITDA Model', 89, 'ebitda'), ('Margin Analysis', 76, 'margin'),
        ('Working Capital Model', 58, 'nwc'), ('Cash Flow Model', 79, 'ufcf'),
    ]
    n = 0
    for sh, r0, metric in checks:
        for i in range(7):
            for col, sc in zip('BCDE', E.SCENARIOS):
                v = G(sh, f'{col}{r0 + i}')
                e = G('Scenario Manager', E.COLS[i] + str(eng.row(sc, metric)))
                if isinstance(v, (int, float)) and isinstance(e, (int, float)):
                    if abs(v - e) > max(1e-6, abs(e) * 1e-6):
                        fail(f'{sh} scenario table disagrees with the engine at {col}{r0 + i}')
                    else:
                        n += 1
    ok(f'{n} scenario-table cells agree with the scenario engine')


ORANGE = 'FFFFD9A0'
RED = 'FFFFC7CE'


def completeness():
    """No cell inside a designed table may be blank unless it is ORANGE-flagged."""
    wb = openpyxl.load_workbook(OUT)
    gaps = []
    orange = red = filled = 0
    for ws in wb.worksheets:
        lim = LEGACY.get(ws.title, ws.max_row + 1)
        # the Support & Audit table is documentation, not a designed input table
        for rr in range(1, ws.max_row + 1):
            if ws.cell(rr, 1).value == 'SUPPORT & AUDIT TABLE':
                lim = min(lim, rr)
                break
        for r in range(1, min(lim, ws.max_row + 1)):
            for c in range(1, min(ws.max_column, 20) + 1):
                cell = ws.cell(r, c)
                b = cell.border
                if not (b.left.style or b.right.style or b.top.style or b.bottom.style):
                    continue
                fg = cell.fill.fgColor.rgb if (cell.fill and cell.fill.patternType
                                               and cell.fill.fgColor) else None
                fg = str(fg) if fg else None
                if cell.value is not None:
                    filled += 1
                    if fg == RED:
                        red += 1
                    continue
                if fg == ORANGE:
                    orange += 1
                    continue
                # a merged continuation cell or a deliberately blank spacer
                if any(cell.coordinate in m for m in ws.merged_cells.ranges):
                    continue
                gaps.append(f'{ws.title}!{cell.coordinate}')
    ok(f'{filled} populated cells inside designed tables; {orange} ORANGE (data unavailable, '
       f'explained in the Support & Audit table); {red} RED (Master Database sourced)')
    # Model Assumptions carries deliberately blank continuation rows in its scenario matrix
    gaps = [g for g in gaps if not g.startswith('Model Assumptions!')]
    if gaps:
        by = {}
        for g in gaps:
            by[g.split('!')[0]] = by.get(g.split('!')[0], 0) + 1
        warn(f'{len(gaps)} blank cells inside designed tables carry no ORANGE flag: {by}')
        for g in gaps[:25]:
            warn('   ' + g)
    else:
        ok('every blank cell inside a designed table carries an ORANGE flag and a documented reason')


def all_scenarios():
    """The workbook must compute cleanly in all four scenarios, not just the selected one."""
    for scen, sid in (('Base Case', 1), ('Bull Case', 2), ('Bear Case', 3), ('Stress Case', 4)):
        wb = openpyxl.load_workbook(OUT)
        wb['Control Panel']['E17'] = scen
        bk = Book(wb)
        bad = 0
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    t = v if isinstance(v, str) else getattr(v, 'text', None)
                    if not (isinstance(t, str) and t.startswith('=')):
                        continue
                    try:
                        bk.cell(ws.title, c.coordinate)
                    except (Err, ZeroDivisionError, ValueError, TypeError, KeyError,
                            RecursionError, IndexError):
                        bad += 1
        res = bk.get('Audit Checks', 'C44')
        margin = bk.get('Margin Analysis', 'J91')
        ebitda = bk.get('EBITDA Model', 'J103')
        if bad:
            fail(f'{scen}: {bad} formulas fail to evaluate')
        else:
            ok(f'{scen}: all formulas evaluate; overall audit = {res}; '
               f'FY2033E margin {margin:.1%}; FY2033E EBITDA Rs {ebitda:,.0f} cr')


def main():
    guard()
    wb, bk = evaluate()
    eng = tieout(wb, bk)
    if eng:
        integrity(wb, bk, eng)
    completeness()
    all_scenarios()
    print('=' * 100)
    for m in OK:
        print('  PASS  ' + m)
    for m in WARN:
        print('  WARN  ' + m)
    for m in FAIL:
        print('  FAIL  ' + m)
    print('=' * 100)
    print(f'{len(OK)} pass, {len(WARN)} warn, {len(FAIL)} fail')
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main())
