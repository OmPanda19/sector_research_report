"""Per-sheet inventory of flagged cells in the built workbook.

Reports, for every sheet, the blank-but-flagged (missing data) cells and the
external-workbook-sourced cells, collapsed into contiguous rectangular ranges so
the output is readable. Used to plan and then verify the elimination of blanks.

Run:  python tools/blanks.py [workbook.xlsx]
"""
import sys
from collections import defaultdict

import openpyxl
from openpyxl.utils import get_column_letter

BANNER = 'SUPPORT & AUDIT TABLE'
REGISTER = 'REFERENCE REGISTER'

sys.path.insert(0, '.')
from gen.spec import LEGACY  # noqa: E402
from gen.style import EXT_RGB as EXTERNAL, MISS_RGB as MISSING  # noqa: E402


def fill_of(cell):
    f = cell.fill
    if f is None or f.patternType != 'solid':
        return None
    rgb = getattr(f.fgColor, 'rgb', None)
    return rgb if isinstance(rgb, str) else None


def stop_row(ws):
    """First row that is documentation rather than model content."""
    stop = LEGACY.get(ws.title)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().upper() in (BANNER, REGISTER):
            stop = r if stop is None else min(stop, r)
            break
    return stop or ws.max_row + 1


def collapse(cells):
    """Collapse a set of (row, col) into a list of A1 range strings."""
    by_row = defaultdict(list)
    for r, c in cells:
        by_row[r].append(c)
    # horizontal runs per row
    runs = []
    for r in sorted(by_row):
        cols = sorted(by_row[r])
        start = prev = cols[0]
        for c in cols[1:]:
            if c == prev + 1:
                prev = c
                continue
            runs.append((r, start, prev))
            start = prev = c
        runs.append((r, start, prev))
    # merge identical column-runs on consecutive rows
    out = []
    runs.sort(key=lambda t: (t[1], t[2], t[0]))
    i = 0
    while i < len(runs):
        r0, c1, c2 = runs[i]
        r1 = r0
        while i + 1 < len(runs) and runs[i + 1][1] == c1 and runs[i + 1][2] == c2 \
                and runs[i + 1][0] == r1 + 1:
            i += 1
            r1 = runs[i][0]
        a = f'{get_column_letter(c1)}{r0}'
        b = f'{get_column_letter(c2)}{r1}'
        out.append(a if a == b else f'{a}:{b}')
        i += 1
    return out


def main(path='Industry Financial Model.xlsx'):
    wb = openpyxl.load_workbook(path)
    tot_blank = tot_red = tot_pop = 0
    print(f'{"sheet":<30}{"populated":>10}{"missing":>9}{"extwb":>7}   ranges')
    print('-' * 100)
    for ws in wb.worksheets:
        stop = stop_row(ws)
        blanks, reds, pop = set(), set(), 0
        for row in ws.iter_rows(min_row=1, max_row=min(stop - 1, ws.max_row)):
            for cell in row:
                f = fill_of(cell)
                if cell.value is not None:
                    pop += 1
                    if f == EXTERNAL:
                        reds.add((cell.row, cell.column))     # external-workbook sourced
                elif f == MISSING:
                    blanks.add((cell.row, cell.column))       # missing data
        tot_pop += pop
        tot_blank += len(blanks)
        tot_red += len(reds)
        if blanks or reds:
            rs = ', '.join(collapse(blanks)[:8])
            more = '' if len(collapse(blanks)) <= 8 else f' (+{len(collapse(blanks)) - 8} more)'
            print(f'{ws.title:<30}{pop:>10}{len(blanks):>9}{len(reds):>7}   {rs}{more}')
        else:
            print(f'{ws.title:<30}{pop:>10}{0:>9}{len(reds):>7}')
    print('-' * 100)
    print(f'{"TOTAL":<30}{tot_pop:>10}{tot_blank:>9}{tot_red:>7}')


if __name__ == '__main__':
    main(*sys.argv[1:])
