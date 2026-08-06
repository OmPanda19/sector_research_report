"""Institutional modelling conventions for the Industry Financial Model.

Colour conventions mandated by the model owner:
    Blue font    = manual input
    Black font   = formula cell
    Green font   = link to another sheet
    Orange fill  = sourced from an external workbook (Master Industry Database.xlsx)
    Red fill     = missing data - no reliable source and no defensible estimate

Note on the orange/red pair: the owner's specification reads "External Workbook -
yellow/orange" and "Missing Data - Red", so external-workbook cells are ORANGE and
missing-data cells are RED. Both are defined once, here, in EXT_RGB / MISS_RGB; swap
those two strings to reverse the pair everywhere in the workbook and in the audit
tooling, which reads these same constants.

The user's existing design is FINAL. We therefore only ever:
  * write .value
  * set .number_format when the existing format is 'General' (i.e. the cell was
    never formatted because it was empty)
  * set .fill only for the EXTERNAL / MISSING conventions
  * set .font colour only where the convention carries information and the cell
    is not part of a coloured header band
"""
from openpyxl.styles import Font, PatternFill, Alignment

# ---------------------------------------------------------------- fills (single source of truth)
EXT_RGB = 'FFFFD9A0'    # orange - value sourced from an external workbook
MISS_RGB = 'FFFFC7CE'   # red    - data missing, nothing asserted

# ---------------------------------------------------------------- fonts
BLUE = 'FF0000CC'      # manual input
BLACK = 'FF000000'     # formula
GREEN = 'FF008000'     # cross-sheet link
DARKRED = 'FF9C0006'   # text on the missing-data fill
DARKORANGE = 'FF7F4F00'  # text on the external-workbook fill

# ---------------------------------------------------------------- fills
FILL_EXT = PatternFill('solid', fgColor=EXT_RGB)    # orange - external workbook
FILL_MISS = PatternFill('solid', fgColor=MISS_RGB)  # red    - missing data

# Legacy aliases: db() writes external-workbook cells, na() writes missing-data cells.
FILL_DB = FILL_EXT
FILL_NA = FILL_MISS

# ---------------------------------------------------------------- number formats
PCT0 = '0%'
PCT1 = '0.0%'
PCT2 = '0.00%'
PPT = '+0.0" ppt";-0.0" ppt";0.0" ppt"'
NUM0 = '#,##0'
NUM1 = '#,##0.0'
NUM2 = '#,##0.00'
NUM3 = '#,##0.000'
CR = '#,##0;(#,##0)'          # Rs crore
RS = '#,##0'                  # Rs per tonne
MT = '#,##0.0'                # million tonnes
MTS = '+#,##0.0;-#,##0.0;0.0'  # signed Mt
X1 = '0.0"x"'
X2 = '0.00"x"'
SCORE = '0.0'
USD = '#,##0.0'
TEXT = '@'
DATE = 'dd-mmm-yyyy'


LINKBLUE = 'FF0563C1'


def link_to(ws, addr, target_sheet, text=None):
    """Internal navigation hyperlink.

    openpyxl's Cell.hyperlink setter only registers a hyperlink on the worksheet when
    `target` is set, so a location-only (internal) link is silently dropped on save.
    The Hyperlink object is therefore appended to ws._hyperlinks explicitly.
    """
    from openpyxl.worksheet.hyperlink import Hyperlink
    c = ws[addr]
    if text is not None:
        c.value = text
    h = Hyperlink(ref=addr, location=f"'{target_sheet}'!A1",
                  tooltip=f'Go to {target_sheet}')
    c.hyperlink = h
    if h not in ws._hyperlinks:
        ws._hyperlinks.append(h)
    f = c.font
    c.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, u='single', color=LINKBLUE)
    return c


class SheetWriter:
    """Bound writer that records provenance for the Support & Audit table."""

    def __init__(self, ws, audit):
        self.ws = ws
        self.audit = audit          # list of dicts, filled by .support()
        self.na_reasons = {}

    # -------------------------------------------------- primitives
    def _cell(self, addr):
        return self.ws[addr]

    def _fmt(self, cell, fmt):
        if fmt and cell.number_format in ('General', '@'):
            cell.number_format = fmt

    def _font(self, cell, colour):
        """Recolour only if the cell is not sitting on a dark design fill."""
        f = cell.font
        fill = cell.fill
        dark = False
        if fill is not None and fill.patternType == 'solid' and fill.fgColor is not None:
            rgb = fill.fgColor.rgb
            if isinstance(rgb, str) and len(rgb) == 8:
                r, g, b = int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)
                dark = (0.299 * r + 0.587 * g + 0.114 * b) < 140
        if dark:
            return
        cell.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, u=f.u,
                         strike=f.strike, vertAlign=f.vertAlign, color=colour)

    # -------------------------------------------------- public API
    def f(self, addr, formula, fmt=None):
        """Formula cell (black font)."""
        c = self._cell(addr)
        c.value = formula if formula.startswith('=') else '=' + formula
        self._fmt(c, fmt)
        self._font(c, BLACK)
        return c

    def link(self, addr, formula, fmt=None):
        """Cross-sheet reference (green font)."""
        c = self._cell(addr)
        c.value = formula if formula.startswith('=') else '=' + formula
        self._fmt(c, fmt)
        self._font(c, GREEN)
        return c

    def inp(self, addr, value, fmt=None):
        """Manual input (blue font)."""
        c = self._cell(addr)
        c.value = value
        self._fmt(c, fmt)
        self._font(c, BLUE)
        return c

    def db(self, addr, value, fmt=None, ref=None):
        """Value sourced from Master Industry Database.xlsx -> ORANGE fill."""
        c = self._cell(addr)
        c.value = value
        self._fmt(c, fmt)
        c.fill = FILL_EXT
        self._font(c, DARKORANGE)
        if ref:
            self.ws.parent._db_refs = getattr(self.ws.parent, '_db_refs', {})
            self.ws.parent._db_refs[f'{self.ws.title}!{addr}'] = ref
        return c

    def na(self, addr, reason):
        """Data genuinely missing -> blank + RED fill.

        Reserved for cells where there is no source AND no defensible way to model the
        value. A cell that can be estimated from a documented driver must be modelled
        with est() instead, so the model stays complete.
        """
        c = self._cell(addr)
        c.value = None
        c.fill = FILL_MISS
        self.na_reasons[addr] = reason
        return c

    def est(self, addr, value, fmt=None):
        """Modelled estimate entered as an input (blue font, no fill).

        Used where no published series exists but the value is derivable from a stated
        assumption. The assumption, its basis and its confidence are recorded in the
        reference register, so the number is auditable rather than asserted.
        """
        c = self._cell(addr)
        c.value = value
        self._fmt(c, fmt)
        self._font(c, BLUE)
        return c

    def txt(self, addr, text, fmt=TEXT):
        c = self._cell(addr)
        c.value = text
        if c.number_format == 'General':
            c.number_format = fmt
        return c

    # -------------------------------------------------- row helpers
    def frow(self, cols, row, template, fmt=None, black=False):
        """Write a formula across `cols` (iterable of column letters).

        `template` is called with (column_letter, index) and must return a formula.
        """
        out = []
        for i, col in enumerate(cols):
            out.append((self.f if black else self.link)(f'{col}{row}', template(col, i), fmt))
        return out

    def narow(self, cols, row, reason):
        for col in cols:
            self.na(f'{col}{row}', reason)
