"""The reference register.

Replaces the old per-sheet Support & Audit tables. Every documented item on a model
sheet now carries a short reference code in a narrow "Ref" column appended to the right
of its table, and all of the detail behind that code lives in one place: sheet-wise
tables on 25 Audit Checks, keyed by code.

Two consequences worth stating:

  * the model sheets stay clean - a model sheet holds model numbers and a code, nothing
    else;
  * the code is the join key, so the detail can be read, sorted and extended without
    touching a single model sheet.

The code column is APPENDED to the right of each table rather than inserted on the left.
Inserting a column programmatically does not repair cross-sheet references, so an insert
would silently break the model; appending cannot. Each code cell is also a hyperlink into
its register table.
"""
import re

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from .spec import CODE, LEGACY

# ---------------------------------------------------------------- register columns
HDR = ['Code', 'Model Item', 'Value', 'Unit', 'Calculation Method', 'Formula Used',
       'Primary Source', 'Secondary Source', 'Assumption', 'Reasoning', 'Cross-check',
       'Confidence Level', 'Linked Sheet', 'Update Frequency', 'Last Available Date', 'Comments']

NAVY = 'FF1F3864'
MID = 'FF2E5C8A'
BAND = 'FFFAFAFA'
CODEFILL = 'FFEDF2F8'
THIN = Side(style='thin', color='FFBFBFBF')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

BANNER = 'REFERENCE REGISTER'
CODE_HDR = 'Ref'


# ==================================================================== formula requalifying
_STR = re.compile(r'"[^"]*"')
_RANGE = re.compile(r"(?<![A-Za-z0-9_!$:'])(\$?[A-Z]{1,3}\$?[0-9]{1,5}"
                    r"(?::\$?[A-Z]{1,3}\$?[0-9]{1,5})?)(?![A-Za-z0-9_(])")


def qualify(formula, sheet):
    """Rewrite bare cell references so the formula works from another sheet.

    '=IF(ISNUMBER(J116),J116,C116)' on 'Cost Curve' becomes
    "=IF(ISNUMBER('Cost Curve'!J116),'Cost Curve'!J116,'Cost Curve'!C116)".
    String literals are protected, and references that already carry a sheet name are
    left alone.
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    parts, last = [], 0
    for m in _STR.finditer(formula):          # protect quoted text
        parts.append((formula[last:m.start()], True))
        parts.append((m.group(0), False))
        last = m.end()
    parts.append((formula[last:], True))
    out = []
    for chunk, substitute in parts:
        if substitute:
            chunk = _RANGE.sub(lambda m: f"'{sheet}'!{m.group(1)}", chunk)
        out.append(chunk)
    return ''.join(out)


# ==================================================================== geometry
def _right_edge(ws, r, limit):
    """Last column in row r carrying content, a border or a design fill."""
    last = 0
    for c in range(1, limit + 1):
        cell = ws.cell(row=r, column=c)
        if cell.value is not None:
            last = c
            continue
        b = cell.border
        if b and (b.left.style or b.right.style or b.top.style or b.bottom.style):
            last = c
            continue
        f = cell.fill
        if f is not None and f.patternType == 'solid':
            rgb = getattr(f.fgColor, 'rgb', None)
            if isinstance(rgb, str) and rgb not in ('00000000', 'FFFFFFFF'):
                last = c
    return last


def code_columns(ws, stop_row):
    """Map every row above stop_row to the column its table's Ref column belongs in.

    Rows are grouped into bands separated by entirely empty rows; one Ref column is
    chosen per band, two columns clear of the widest row in that band, so codes line up
    down a table instead of stepping in and out.
    """
    limit = min(ws.max_column + 2, 24)
    edges = {}
    for r in range(1, stop_row):
        edges[r] = _right_edge(ws, r, limit)
    bands, cur = [], []
    for r in range(1, stop_row):
        if edges[r]:
            cur.append(r)
        elif cur:
            bands.append(cur)
            cur = []
    if cur:
        bands.append(cur)
    out = {}
    for band in bands:
        col = max(edges[r] for r in band) + 2
        for r in band:
            out[r] = col
    return out


def doc_start(ws):
    """First row of documentation, i.e. where model content stops."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().upper() in (BANNER, 'SUPPORT & AUDIT TABLE'):
            return r
    return ws.max_row + 1


# ==================================================================== anchors
_FIRSTREF = re.compile(r"(?<![A-Za-z0-9_!$'])\$?([A-Z]{1,3})\$?([0-9]{1,5})(?![A-Za-z0-9_(])")


def anchor_rows(entry, sheet):
    """Rows on the origin sheet that this entry describes.

    Explicit `anchor` wins. Otherwise the rows are inferred: from the address range
    recorded against a missing-data group, or from the first same-sheet cell reference
    in the entry's live Value formula.
    """
    a = entry.get('anchor')
    if a is not None:
        return [a] if isinstance(a, int) else list(a)

    addrs = entry.get('addrs')
    if addrs:
        rows = set()
        for part in str(addrs).replace(' ', '').split(','):
            m = re.match(r'^\$?[A-Z]{1,3}\$?(\d+)(?::\$?[A-Z]{1,3}\$?(\d+))?$', part)
            if m:
                lo = int(m.group(1))
                hi = int(m.group(2) or m.group(1))
                rows.update(range(min(lo, hi), max(lo, hi) + 1))
        if rows:
            return sorted(rows)

    val = entry.get('value')
    if isinstance(val, str) and val.startswith('='):
        stripped = _STR.sub('""', val)
        if "'" not in stripped and '!' not in stripped:
            m = _FIRSTREF.search(stripped)
            if m:
                return [int(m.group(2))]
    return []


# ==================================================================== stamping codes
def stamp(ws, assignments, register_row):
    """Write reference codes into the Ref column of each table on `ws`.

    `assignments` is a list of (code, [rows]). Returns the number of cells stamped.
    """
    stop = doc_start(ws)
    colmap = code_columns(ws, stop)
    if not colmap:
        return 0

    per_row = {}
    for code, rows in assignments:
        for r in rows:
            if r >= stop:
                continue
            per_row.setdefault(r, []).append(code)

    def band_top(row, col):
        """First row of the contiguous band that `row` belongs to."""
        t = row
        while t - 1 >= 1 and colmap.get(t - 1) == col:
            t -= 1
        return t

    n = 0
    header_done = set()
    for r, codes in sorted(per_row.items()):
        col = colmap.get(r)
        if not col:
            continue
        cell = ws.cell(row=r, column=col)
        if cell.value is not None or type(cell).__name__ == 'MergedCell':
            continue
        cell.value = ', '.join(dict.fromkeys(codes))
        cell.font = Font(name='Calibri', sz=8, b=True, color=NAVY, u='single')
        cell.fill = PatternFill('solid', fgColor=CODEFILL)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = BOX
        h = Hyperlink(ref=cell.coordinate, location=f"'Audit Checks'!A{register_row}",
                      tooltip='Open the reference register on 25 Audit Checks')
        cell.hyperlink = h
        ws._hyperlinks.append(h)
        n += 1
        # a "Ref" header on the first stamped row of each TABLE, not merely of each column -
        # a sheet routinely has several tables sharing one Ref column
        key = (band_top(r, col), col)
        if key not in header_done:
            header_done.add(key)
            _ref_header(ws, r, col, colmap)
        letter = get_column_letter(col)
        if letter not in ws.column_dimensions:
            ws.column_dimensions[letter].width = 7
    return n


def _ref_header(ws, first_row, col, colmap):
    """Label the Ref column at the top of its band."""
    r = first_row
    while r - 1 >= 1 and colmap.get(r - 1) == col:
        r -= 1
    # walk down to the first row that looks like data rather than a section title
    target = None
    for rr in range(r, first_row + 1):
        if ws.cell(row=rr, column=2).value is not None or ws.cell(row=rr, column=3).value is not None:
            target = rr
            break
    if target is None or target == first_row:
        target = r
    cell = ws.cell(row=target, column=col)
    if cell.value is not None or type(cell).__name__ == 'MergedCell':
        return
    cell.value = CODE_HDR
    cell.font = Font(name='Calibri', sz=8, b=True, color='FFFFFFFF')
    cell.fill = PatternFill('solid', fgColor=MID)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = BOX


# ==================================================================== the register itself
def build(wb, entries_by_sheet, intro=None, intros=None):
    """Allocate codes, stamp them on every sheet and write the register on Audit Checks.

    Returns (total_entries, {sheet: register_header_row}).
    """
    ac = wb['Audit Checks']
    start = doc_start(ac)
    start = max(start, ac.max_row + 3)

    order = [n for n in wb.sheetnames if entries_by_sheet.get(n)]

    # ---- pass 1: allocate codes and work out where each sheet's table will start
    alloc = {}          # sheet -> list of (code, entry)
    for name in order:
        prefix = CODE.get(name, 'XX')
        alloc[name] = [(f'{prefix}{i + 1:02d}', e)
                       for i, e in enumerate(entries_by_sheet[name])]

    # index block, then one table per sheet
    idx_rows = len(order) + 4
    cursor = start + idx_rows
    header_row = {}
    for name in order:
        # banner on `cursor`, intro beneath it, header beneath that
        header_row[name] = cursor + 2
        cursor += 3 + len(alloc[name]) + 2

    # ---- pass 2: stamp the codes onto the model sheets
    stamped = 0
    for name in order:
        ws = wb[name]
        assignments = [(code, anchor_rows(e, name)) for code, e in alloc[name]]
        assignments = [(c, rows) for c, rows in assignments if rows]
        stamped += stamp(ws, assignments, header_row[name])

    # ---- pass 3: write the register
    _index(ac, start, order, alloc, header_row, intro)
    for name in order:
        _table(ac, header_row[name] - 2, name, alloc[name], (intros or {}).get(name))

    _widths(ac)
    total = sum(len(v) for v in alloc.values())
    return total, header_row, stamped


def _index(ws, top, order, alloc, header_row, intro):
    t = ws.cell(top, 1, BANNER)
    t.font = Font(name='Calibri', sz=12, b=True, color='FFFFFFFF')
    t.fill = PatternFill('solid', fgColor=NAVY)
    t.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    for c in range(2, len(HDR) + 1):
        ws.cell(top, c).fill = PatternFill('solid', fgColor=NAVY)
    ws.row_dimensions[top].height = 20

    sub = ws.cell(top + 1, 1, intro or (
        'Every documented number in the workbook, keyed by the reference code that appears in the "Ref" '
        'column beside it on its own sheet. Sheets are listed in workbook order; click a code on any sheet '
        'to jump to its entry here, and click a sheet name below to jump to its table. The Value column is '
        'live, so the register re-states itself whenever the scenario on 01 Control Panel is changed. This '
        'register replaces the Support & Audit tables that previously sat at the bottom of every sheet.'))
    sub.font = Font(name='Calibri', sz=9, i=True, color='FF595959')
    sub.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=top + 1, start_column=1, end_row=top + 1, end_column=len(HDR))
    ws.row_dimensions[top + 1].height = 42

    h = ws.cell(top + 2, 1, 'Sheet')
    h2 = ws.cell(top + 2, 2, 'Codes')
    h3 = ws.cell(top + 2, 3, 'Entries')
    for c in (h, h2, h3):
        c.font = Font(name='Calibri', sz=9, b=True, color='FFFFFFFF')
        c.fill = PatternFill('solid', fgColor=MID)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = BOX

    for i, name in enumerate(order):
        r = top + 3 + i
        codes = alloc[name]
        a = ws.cell(r, 1, name)
        a.font = Font(name='Calibri', sz=9, color='FF0563C1', u='single')
        hl = Hyperlink(ref=a.coordinate, location=f"'Audit Checks'!A{header_row[name]}",
                       tooltip=f'Jump to the {name} register table')
        a.hyperlink = hl
        ws._hyperlinks.append(hl)
        b = ws.cell(r, 2, f'{codes[0][0]} - {codes[-1][0]}' if codes else '')
        c = ws.cell(r, 3, len(codes))
        for cc in (a, b, c):
            cc.border = BOX
            cc.alignment = Alignment(horizontal='left' if cc is a else 'center', vertical='center')
            if cc is not a:
                cc.font = Font(name='Calibri', sz=9, color='FF404040')


def _table(ws, top, name, entries, intro=None):
    codes = f'{entries[0][0]} - {entries[-1][0]}' if entries else ''
    banner = ws.cell(top, 1, f'{name}  -  {len(entries)} documented items  ({codes})')
    banner.font = Font(name='Calibri', sz=10, b=True, color='FFFFFFFF')
    banner.fill = PatternFill('solid', fgColor=NAVY)
    banner.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    for c in range(2, len(HDR) + 1):
        ws.cell(top, c).fill = PatternFill('solid', fgColor=NAVY)
    ws.row_dimensions[top].height = 18

    note = ws.cell(top + 1, 1, intro or (
        f'Every documented number on {name}. The Code column matches the Ref column on that sheet.'))
    note.font = Font(name='Calibri', sz=8, i=True, color='FF595959')
    note.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=top + 1, start_column=1, end_row=top + 1, end_column=len(HDR))
    ws.row_dimensions[top + 1].height = 24

    hr = top + 2
    for i, htext in enumerate(HDR, start=1):
        c = ws.cell(hr, i, htext)
        c.font = Font(name='Calibri', sz=9, b=True, color='FFFFFFFF')
        c.fill = PatternFill('solid', fgColor=MID)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BOX
    ws.row_dimensions[hr].height = 30

    for n, (code, d) in enumerate(entries):
        r = hr + 1 + n
        value = d.get('value')
        if isinstance(value, str) and value.startswith('='):
            value = qualify(value, name)
        vals = [code, d.get('item', ''), value, d.get('unit', ''), d.get('method', ''),
                d.get('formula', ''), d.get('primary', ''), d.get('secondary', ''),
                d.get('assumption', ''), d.get('reasoning', ''), d.get('cross', ''),
                d.get('conf', ''), d.get('linked', ''), d.get('freq', ''),
                d.get('last', ''), d.get('comments', '')]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(r, i)
            if i == 1:
                c.value = v
                c.font = Font(name='Calibri', sz=9, b=True, color=NAVY)
                c.fill = PatternFill('solid', fgColor=CODEFILL)
            elif i == 3 and isinstance(v, str) and v.startswith('='):
                c.value = v
                c.number_format = d.get('numfmt') or '#,##0.00'
                c.font = Font(name='Calibri', sz=9, b=True, color=NAVY)
            else:
                if isinstance(v, str) and v.startswith('='):
                    v = "'" + v              # 'Formula Used' stays text
                c.value = v if v != '' else None
                c.font = Font(name='Calibri', sz=8, color='FF404040')
            c.alignment = Alignment(wrap_text=True, vertical='top',
                                    horizontal='center' if i in (1, 4, 12, 15) else 'left')
            c.border = BOX
            if n % 2 and i != 1:
                c.fill = PatternFill('solid', fgColor=BAND)
        ws.row_dimensions[r].height = 46


WIDTHS = {'A': 10, 'B': 34, 'C': 14, 'D': 13, 'E': 30, 'F': 34, 'G': 30, 'H': 28,
          'I': 30, 'J': 52, 'K': 32, 'L': 16, 'M': 22, 'N': 24, 'O': 18, 'P': 30}


def _widths(ws):
    for letter, wid in WIDTHS.items():
        cur = ws.column_dimensions.get(letter)
        if cur is None or not cur.width or cur.width < wid:
            ws.column_dimensions[letter].width = wid
