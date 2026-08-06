"""Probe table geometry: for each sheet, report the right-hand edge of each
contiguous band of content above the documentation block, so the reference-code
column can be placed immediately to the right of each table.

Run:  python tools/geom.py [workbook.xlsx]
"""
import sys
from collections import Counter

import openpyxl
from openpyxl.utils import get_column_letter

sys.path.insert(0, '.')

BANNERS = ('SUPPORT & AUDIT TABLE', 'REFERENCE REGISTER')


def doc_row(ws):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().upper() in BANNERS:
            return r
    return ws.max_row + 1


def right_edge(ws, r, limit):
    """Last column in row r carrying content or a border."""
    last = 0
    for c in range(1, limit + 1):
        cell = ws.cell(row=r, column=c)
        if cell.value is not None:
            last = c
            continue
        b = cell.border
        if b and (b.left.style or b.right.style or b.top.style or b.bottom.style):
            last = c
        f = cell.fill
        if f is not None and f.patternType == 'solid':
            rgb = getattr(f.fgColor, 'rgb', None)
            if isinstance(rgb, str) and rgb not in ('00000000', 'FFFFFFFF'):
                last = c
    return last


def main(path='Industry Financial Model.xlsx'):
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        stop = doc_row(ws)
        limit = min(ws.max_column, 20)
        edges = Counter()
        for r in range(1, min(stop, ws.max_row + 1)):
            e = right_edge(ws, r, limit)
            if e:
                edges[e] += 1
        pretty = ', '.join(f'{get_column_letter(c)}({n})'
                           for c, n in sorted(edges.items(), reverse=True))
        print(f'{ws.title:<30} docrow={stop:<5} maxcol={get_column_letter(ws.max_column):<3} edges: {pretty}')


if __name__ == '__main__':
    main(*sys.argv[1:])
