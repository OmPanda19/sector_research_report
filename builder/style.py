"""Institutional formatting engine for Master Industry Database.xlsx.

House style rules enforced here (single source of truth):
  * One typeface throughout: Calibri. Body 10pt, headers 10pt bold, titles 14/11pt bold.
  * Two structural colours only: NAVY for header bands, RULE_GREY for hairline rules.
    No decorative fills, no colour-coding of values, no banded rows.
  * No merged cells anywhere. Title blocks use a single cell with text; the adjacent
    cells are left genuinely empty so that filters, tables and Power Query all behave.
  * Numbers are right-aligned, text left-aligned, headers wrapped and top-aligned.
  * Every numeric column carries an explicit number format. No "General" on numbers.
  * Negative numbers in brackets, thousands separated, consistent decimals per unit.
"""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --------------------------------------------------------------------------------------
# Palette (deliberately minimal)
# --------------------------------------------------------------------------------------
NAVY = "1F3864"          # header bands
NAVY_LIGHT = "D9E1F2"    # sub-header / total emphasis, used sparingly
RULE_GREY = "BFBFBF"     # hairline rules
TEXT_DARK = "000000"
TEXT_MUTED = "595959"    # footnotes, metadata
WHITE = "FFFFFF"

FONT_NAME = "Calibri"

# --------------------------------------------------------------------------------------
# Number formats
# --------------------------------------------------------------------------------------
FMT_TEXT = "@"
FMT_INT = "#,##0;(#,##0)"
FMT_1DP = "#,##0.0;(#,##0.0)"
FMT_2DP = "#,##0.00;(#,##0.00)"
FMT_3DP = "#,##0.000;(#,##0.000)"
FMT_PCT_1 = "0.0%;(0.0%)"
FMT_PCT_2 = "0.00%;(0.00%)"
FMT_MULT = '#,##0.00"x"'
FMT_DATE = "DD-MMM-YYYY"

# Canonical format per unit string, so units and decimals can never drift apart.
UNIT_FORMAT = {
    "Mt": FMT_1DP,
    "Mtpa": FMT_1DP,
    "mn t": FMT_2DP,
    "kt": FMT_INT,
    "'000 t": FMT_INT,
    "kg": FMT_1DP,
    "Rs crore": FMT_INT,
    "Rs cr": FMT_INT,
    "Rs/t": FMT_INT,
    "Rs/tonne": FMT_INT,
    "US$/t": FMT_2DP,
    "US$/dmt": FMT_2DP,
    "US$/mt": FMT_2DP,
    "US$ m": FMT_INT,
    "US$ bn": FMT_2DP,
    "EUR m": FMT_INT,
    "GBP m": FMT_INT,
    "%": FMT_PCT_1,
    "x": FMT_MULT,
    "Rs": FMT_2DP,
    "Number": FMT_INT,
    "tCO2e/tfs": FMT_2DP,
}


def fmt_for_unit(unit, default=FMT_1DP):
    return UNIT_FORMAT.get((unit or "").strip(), default)


# --------------------------------------------------------------------------------------
# Reusable style atoms
# --------------------------------------------------------------------------------------
def font(size=10, bold=False, italic=False, colour=TEXT_DARK):
    return Font(name=FONT_NAME, size=size, bold=bold, italic=italic, color=colour)


THIN = Side(style="thin", color=RULE_GREY)
MEDIUM_NAVY = Side(style="thin", color=NAVY)

BORDER_HEADER = Border(bottom=MEDIUM_NAVY)
BORDER_BOTTOM = Border(bottom=THIN)
BORDER_TOP_RULE = Border(top=MEDIUM_NAVY)
BORDER_NONE = Border()

FILL_NAVY = PatternFill("solid", fgColor=NAVY)
FILL_NAVY_LIGHT = PatternFill("solid", fgColor=NAVY_LIGHT)

AL_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=False)
AL_LEFT_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
AL_RIGHT = Alignment(horizontal="right", vertical="top")
AL_CENTRE = Alignment(horizontal="center", vertical="top", wrap_text=True)
AL_HEADER = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
AL_HEADER_LEFT = Alignment(horizontal="left", vertical="bottom", wrap_text=True)


# --------------------------------------------------------------------------------------
# Sheet furniture
# --------------------------------------------------------------------------------------
def sheet_defaults(ws, tab_colour=NAVY):
    ws.sheet_properties.tabColor = tab_colour
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


def set_widths(ws, widths):
    """widths: iterable of (column_index, width)."""
    for idx, width in widths:
        ws.column_dimensions[get_column_letter(idx)].width = width


def put(ws, row, col, value, *, f=None, number_format=None, alignment=None,
        fill=None, border=None):
    cell = ws.cell(row=row, column=col)
    cell.value = value
    cell.font = f or font()
    if number_format:
        cell.number_format = number_format
    if alignment:
        cell.alignment = alignment
    elif isinstance(value, (int, float)):
        cell.alignment = AL_RIGHT
    else:
        cell.alignment = AL_LEFT
    if fill:
        cell.fill = fill
    if border:
        cell.border = border
    return cell


def title_block(ws, lines, start_row=1, width_col=1):
    """Write a stacked title/metadata block without merging cells.

    lines: list of (text, kind) where kind in {'title','subtitle','label','body'}.
    """
    row = start_row
    for text, kind in lines:
        if kind == "title":
            put(ws, row, width_col, text, f=font(14, bold=True, colour=NAVY))
            ws.row_dimensions[row].height = 19
        elif kind == "subtitle":
            put(ws, row, width_col, text, f=font(11, bold=True, colour=NAVY))
            ws.row_dimensions[row].height = 15
        elif kind == "label":
            put(ws, row, width_col, text, f=font(9, bold=True, colour=TEXT_MUTED))
        else:
            put(ws, row, width_col, text, f=font(9, colour=TEXT_MUTED),
                alignment=AL_LEFT_WRAP)
        row += 1
    return row


def header_row(ws, row, headers, start_col=1, height=30, left_align_cols=()):
    """Write a navy header band. headers: list of strings (must be unique, non-blank)."""
    ws.row_dimensions[row].height = height
    for i, text in enumerate(headers):
        col = start_col + i
        alignment = AL_HEADER_LEFT if col in left_align_cols else AL_HEADER
        put(ws, row, col, text, f=font(10, bold=True, colour=WHITE),
            alignment=alignment, fill=FILL_NAVY, border=BORDER_HEADER)
    return row + 1


def footnote(ws, row, col, text):
    put(ws, row, col, text, f=font(8, italic=True, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP)
    return row + 1


def add_table(ws, name, first_row, first_col, last_row, last_col):
    """Wrap a written range in an Excel Table so downstream ranges are dynamic.

    Excel Tables give: structured references for models, auto-expanding ranges when a
    new fiscal year row is appended, and native Power BI / Power Query pickup.
    """
    if last_row <= first_row:
        return None
    ref = "%s%d:%s%d" % (get_column_letter(first_col), first_row,
                         get_column_letter(last_col), last_row)
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1", showFirstColumn=False, showLastColumn=False,
        showRowStripes=False, showColumnStripes=False)
    ws.add_table(table)
    return table


def freeze(ws, cell_ref):
    ws.freeze_panes = cell_ref


def repeat_header(ws, row):
    ws.print_title_rows = "%d:%d" % (row, row)


NA = "Data Not Publicly Available"
