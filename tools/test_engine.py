#!/usr/bin/env python3
"""Prove the scenario engine reproduces (a) the live model and (b) the independent tie-out."""
import sys, openpyxl
sys.path.insert(0, '.')
from tools.xlcalc import Book
from gen import engine as E

wb = openpyxl.load_workbook('Industry Financial Model.xlsx')
bk = Book(wb)

# locate the engine block base rows by scanning for the scenario banners
ws = wb['Scenario Manager']
base = {}
for r in range(90, ws.max_row + 1):
    v = ws.cell(r, 1).value
    if isinstance(v, str) and v.strip().upper() in [s.upper() for s in E.SCENARIOS]:
        for s in E.SCENARIOS:
            if v.strip().upper() == s.upper():
                base[s] = r + 1
print('engine base rows:', base)
eng = E.Engine(base)

# ---------------------------------------------------------------- tie-out on FY2033E
TIE = [  # Scenario Manager row -> engine metric name
    (69, 'cons'), (70, 'crude'), (71, 'cap'), (72, 'util'), (73, 'real'),
    (74, 'cost'), (75, 'ebitdat'), (76, 'margin'), (77, 'rev'), (78, 'ebitda'), (79, 'ufcf'),
]
SCOL = {'Base Case': 'C', 'Bull Case': 'D', 'Bear Case': 'E', 'Stress Case': 'F'}
bad = 0
print('\n%-34s %-11s %16s %16s %9s' % ('metric', 'scenario', 'engine', 'tie-out (hard)', 'diff %'))
for row, name in TIE:
    for sc, col in SCOL.items():
        want = ws[f'{col}{row}'].value
        got = bk.get('Scenario Manager', f'J{eng.row(sc, name)}')
        if want is None or got is None:
            print('  MISSING', name, sc, want, got)
            bad += 1
            continue
        d = (got - want) / abs(want) if want else (got - want)
        flag = '' if abs(d) < 0.005 else '   <-- MISMATCH'
        if flag:
            bad += 1
        print('%-34s %-11s %16.4f %16.4f %8.3f%%%s' % (name, sc, got, want, d * 100, flag))

# ---------------------------------------------------------------- engine vs live model
LIVE = [  # (engine metric, live sheet, live row)
    ('cons', 'Steel Demand Model', 116), ('crude', 'Steel Supply Model', 99),
    ('fin', 'Steel Supply Model', 101), ('cap', 'Capacity Forecast', 101),
    ('util', 'Capacity Utilisation', 67), ('real', 'Steel Price Forecast', 103),
    ('cost', 'Cost Curve', 80), ('ebitdat', 'EBITDA Model', 101),
    ('ebitda', 'EBITDA Model', 103), ('rev', 'Revenue Forecast', 96),
    ('margin', 'Margin Analysis', 91), ('ufcf', 'Cash Flow Model', 114),
    ('capex', 'Cash Flow Model', 110), ('nwc', 'Working Capital Model', 72),
    ('imp', 'Trade Model', 76), ('exp', 'Trade Model', 78), ('percap', 'Steel Demand Model', 121),
]
active = bk.get('Control Panel', 'E17')
print(f'\nactive scenario = {active!r}; comparing engine block against the live model')
print('%-12s %-24s %-9s %16s %16s %9s' % ('metric', 'live sheet', 'year', 'engine', 'live', 'diff %'))
for name, sh, lrow in LIVE:
    for j, lcol in enumerate('CDEFGHIJ'):
        live = bk.get(sh, f'{lcol}{lrow}')
        got = bk.get('Scenario Manager', f'{E.A if j == 0 else E.COLS[j-1]}{eng.row(active, name)}')
        if live is None or got is None or isinstance(live, str) or isinstance(got, str):
            continue
        d = (got - live) / abs(live) if live else got - live
        if abs(d) > 0.002:
            bad += 1
            print('%-12s %-24s %-9s %16.4f %16.4f %8.3f%%   <-- MISMATCH'
                  % (name, sh, ['FY26A'] + [f'FY{27+i}E' for i in range(7)][:7][j-1] if False else
                     (['FY26A', 'FY27E', 'FY28E', 'FY29E', 'FY30E', 'FY31E', 'FY32E', 'FY33E'][j]),
                     got, live, d * 100))
print('\nevaluator errors:', len(bk.errors))
for k, v in list(bk.errors.items())[:25]:
    print('   ', k, '->', v)
print('\nRESULT:', 'PASS' if bad == 0 else f'{bad} MISMATCH(ES)')
