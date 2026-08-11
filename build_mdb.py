#!/usr/bin/env python3
"""
build_mdb.py  -  Rebuild the Master Industry Database (Indian Steel Industry).

Single source of truth for the Industry Financial Model.
Design goals: extremely simple, visually clean, fully auditable, consistent.
Every data table ends in a single "Source Code" column. Historical FY2021-FY2026 only.
Research cut-off: 10 August 2026.

Run:  /projects/sandbox/.venv/bin/python build_mdb.py
Output: Master Industry Database.xlsx  (in the repo root)
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# THEME  -  Dark Blue
# ----------------------------------------------------------------------------
NAVY      = "1F3864"   # banner / title
HEADER    = "2F5496"   # table header row
SECTION   = "8EAADB"   # section sub-title band
LIGHT     = "D9E1F2"   # alt band / labels
ZEBRA     = "F2F6FC"   # zebra shading
WHITE     = "FFFFFF"
CODEFILL  = "E2EFDA"   # source-code column tint (subtle green-grey)
INK       = "1A1A1A"

FONT_NAME = "Calibri"

def _font(sz=10, bold=False, color=INK):
    return Font(name=FONT_NAME, size=sz, bold=bold, color=color)

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=False)
RIGHT  = Alignment(horizontal="right",  vertical="center", wrap_text=False)
LEFTW  = Alignment(horizontal="left",   vertical="center", wrap_text=True)

def fill(hexc):
    return PatternFill("solid", fgColor=hexc)

# Standard fiscal-year headers
FY = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "FY2026"]

# ----------------------------------------------------------------------------
# Low-level helpers
# ----------------------------------------------------------------------------
def banner(ws, title, ncols):
    """Row 1 dark navy title banner + row 2 subtitle. Merged across ncols."""
    last = get_column_letter(ncols)
    ws.merge_cells(f"A1:{last}1")
    c = ws["A1"]
    c.value = title
    c.font = _font(15, True, WHITE)
    c.fill = fill(NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last}2")
    c = ws["A2"]
    c.value = ("Master Industry Database   |   Indian Steel Industry   |   "
               "Historical FY2021-FY2026   |   Research cut-off 10 August 2026")
    c.font = _font(9, False, WHITE)
    c.fill = fill(HEADER)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

def section_title(ws, row, text, ncols):
    last = get_column_letter(ncols)
    ws.merge_cells(f"A{row}:{last}{row}")
    c = ws.cell(row, 1)
    c.value = text
    c.font = _font(10, True, NAVY)
    c.fill = fill(SECTION)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 18

def header_row(ws, row, headers):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row, j, h)
        c.font = _font(10, True, WHITE)
        c.fill = fill(HEADER)
        c.alignment = CENTER if j > 1 else Alignment(horizontal="left", vertical="center", indent=1)
        c.border = BORDER
    ws.row_dimensions[row].height = 20

def data_table(ws, start_row, headers, rows, label_widths=None,
               numfmt=None, zebra=True):
    """
    Generic table writer.
    headers: list of column headers (last is 'Source Code').
    rows: list of tuples/lists matching headers length. Values may be None.
    numfmt: dict {col_index(1-based): excel number format} for value columns.
    Returns the next free row.
    """
    header_row(ws, start_row, headers)
    ncols = len(headers)
    r = start_row + 1
    for i, row in enumerate(rows):
        for j, val in enumerate(row, start=1):
            c = ws.cell(r, j, val)
            c.border = BORDER
            if j == 1:
                c.font = _font(10, False)
                c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif j == ncols:  # Source Code column
                c.font = _font(9, False, "375623")
                c.alignment = CENTER
                c.fill = fill(CODEFILL)
            else:
                c.font = _font(10, False)
                c.alignment = CENTER
                if numfmt and j in numfmt and isinstance(val, (int, float)):
                    c.number_format = numfmt[j]
            if zebra and j != ncols and i % 2 == 1:
                if not (j == ncols):
                    c.fill = fill(ZEBRA)
        r += 1
    return r + 1  # blank spacer row after table

def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

# Number formats
F_MT   = '#,##0.00'          # million tonnes
F_MTPA = '#,##0.00'
F_PCT  = '0.0%'
F_USD  = '#,##0.0'           # US$/t
F_RS0  = '#,##0'             # Rs/t, Rs cr
F_RS2  = '#,##0.00'          # Rs/kWh
F_KG   = '#,##0.0'

def fy_numfmt(fmt):
    # value columns are B..G when layout = Series|Unit|FY..|Code  -> cols 3..8
    return {3: fmt, 4: fmt, 5: fmt, 6: fmt, 7: fmt, 8: fmt}

STD_HEADERS = ["Series", "Unit"] + FY + ["Source Code"]
STD_WIDTHS  = {"A": 40, "B": 12, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 13}

def std_sheet(ws):
    set_widths(ws, STD_WIDTHS)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

# ============================================================================
wb = Workbook()

# ----------------------------------------------------------------------------
# 1. COVER
# ----------------------------------------------------------------------------
ws = wb.active
ws.title = "Cover"
ws.sheet_view.showGridLines = False
set_widths(ws, {"A": 3, "B": 100})
for r in range(1, 40):
    ws.row_dimensions[r].height = 16
ws["B2"] = "MASTER INDUSTRY DATABASE"
ws["B2"].font = _font(22, True, NAVY)
ws["B3"] = "INDIAN STEEL INDUSTRY"
ws["B3"].font = _font(14, True, HEADER)
ws["B5"] = "Global Metals & Mining Research"
ws["B5"].font = _font(11, True)
ws["B6"] = "Historical period: FY2021-FY2026   |   Research cut-off: 10 August 2026"
ws["B6"].font = _font(10)
cover_blocks = [
    (8,  "PURPOSE"),
    (9,  "The single source of truth for the Industry Financial Model. Every historical input the model uses is stored here, with a source code that ties back to the Sources sheet."),
    (11, "HOW TO READ THIS WORKBOOK"),
    (12, "Every data sheet uses the same simple layout: one table per topic, historical years FY2021 to FY2026 across the columns, and a single Source Code column at the far right. No explanations sit inside the data tables - all detail lives in the Sources sheet."),
    (14, "COVERAGE"),
    (15, "Core companies (5, full historical disclosure): Tata Steel, JSW Steel, SAIL, Jindal Steel, Jindal Stainless. Every other producer (AM/NS India, RINL and the secondary sector) is grouped under 'Others'."),
    (17, "BASIS OF PREPARATION"),
    (18, "Consolidated figures throughout unless stated. Units are kept identical across sheets: Mt, Mtpa, Rs/t, Rs crore, US$/t, US$/dmt, % and kg."),
    (20, "DATA QUALITY"),
    (21, "Latest official value is used when sources conflict. Conflicts are documented on the Conflicts Log; they are never averaged away. Series that had to be reconstructed as fiscal-year averages from monthly or spot data are flagged with their confidence on the Sources sheet."),
    (23, "PREPARED FROM PUBLIC SOURCES ONLY"),
    (24, "Contains no non-public or price-sensitive information. Sources: Ministry of Steel, JPC, IBM, World Steel Association, World Bank, RBI, MoSPI, company results and investor presentations, and recognised price reporting agencies."),
]
for r, txt in cover_blocks:
    c = ws.cell(r, 2, txt)
    if txt.isupper():
        c.font = _font(10, True, HEADER)
    else:
        c.font = _font(10)
        c.alignment = LEFTW
        ws.row_dimensions[r].height = 42

# ----------------------------------------------------------------------------
# 2. CONTENTS
# ----------------------------------------------------------------------------
ws = wb.create_sheet("Contents")
ws.sheet_view.showGridLines = False
set_widths(ws, {"A": 5, "B": 26, "C": 74, "D": 12})
banner(ws, "CONTENTS", 4)
contents = [
    ("#", "Worksheet", "What it contains", "Tier"),
    (1, "Company List", "Coverage universe: 5 Core companies and the 'Others' grouping, with route and reason", "1"),
    (2, "Macroeconomic", "GDP growth, CPI, USD/INR, 10Y G-sec and repo rate - the macro inputs the model uses", "1"),
    (3, "Production Capacity", "India crude steel capacity, capacity added and utilisation", "1"),
    (4, "Production Volume", "Crude steel, finished steel, DRI and pig iron output", "1"),
    (5, "Demand & Consumption", "Apparent finished steel consumption, growth and per-capita consumption", "1"),
    (6, "Imports & Exports", "Finished steel imports, exports and the net trade position", "1"),
    (7, "Steel Prices", "HRC, CRC, rebar, wire rod, plate, billet and the blended industry ASP", "1/2"),
    (8, "Raw Materials", "Iron ore, coking coal, PCI, thermal coal, scrap, pellet, fluxes, gas, power, freight", "1/2"),
    (9, "EBITDA", "Revenue, EBITDA, EBITDA margin and EBITDA per tonne for the Core companies and Others", "1"),
    (10, "Industry KPIs", "A compact dashboard of the headline supply, demand, trade and cost indicators", "1"),
    (11, "Government Policies", "Safeguard duty, PLI and green-steel measures that drive the model", "1"),
    (12, "Sources", "Every source code with publisher, publication, date, URL, access date and confidence", "Ref"),
    (13, "Conflicts Log", "Documented data conflicts and comparability breaks - never averaged away", "Ref"),
]
r = 5
for i, row in enumerate(contents):
    if i == 0:
        header_row(ws, r, list(row))
    else:
        for j, v in enumerate(row, start=1):
            c = ws.cell(r, j, v)
            c.border = BORDER
            c.font = _font(10, False)
            c.alignment = LEFT if j in (2, 3) else CENTER
            if i % 2 == 0:
                c.fill = fill(ZEBRA)
    r += 1
ws.freeze_panes = "A6"

# ----------------------------------------------------------------------------
# 3. COMPANY LIST  (one screen, no scroll)
# ----------------------------------------------------------------------------
ws = wb.create_sheet("Company List")
ws.sheet_view.showGridLines = False
set_widths(ws, {"A": 22, "B": 13, "C": 18, "D": 16, "E": 14, "F": 52})
banner(ws, "1.  COMPANY LIST", 6)
hdr = ["Company", "Ticker", "Steel Segment", "Route", "Included?", "Reason"]
companies = [
    ("Tata Steel",        "TATASTEEL",  "Flat + Long",     "BF-BOF / EAF",   "Core",   "Full FY2021-FY2026 consolidated Ind-AS disclosure; volumes and EBITDA/t reported"),
    ("JSW Steel",         "JSWSTEEL",   "Flat + Long",     "BF-BOF / DRI-EAF","Core",   "Full disclosure; largest Indian producer by volume"),
    ("SAIL",              "SAIL",       "Flat + Long",     "BF-BOF",         "Core",   "Full disclosure; integrated PSU major"),
    ("Jindal Steel",      "JINDALSTEL", "Long + Plate",    "DRI-EAF / BF-BOF","Core",   "Full disclosure (note: power business demerged in FY2022)"),
    ("Jindal Stainless",  "JSL",        "Stainless flat",  "EAF-AOD",        "Core",   "Full disclosure; stainless leader (note: Hisar merger completed FY2024)"),
    ("AM/NS India",       "Unlisted",   "Flat",            "BF-BOF / DRI-EAF","Others", "Reports 100% basis in USD on a calendar year; no full FY Ind-AS series"),
    ("RINL (Vizag Steel)","Unlisted",   "Long",            "BF-BOF",         "Others", "Unlisted PSU; EBITDA not publicly disclosed"),
    ("Shyam Metalics",    "SHYAMMETL",  "Long / Alloy",    "DRI-IF / EAF",   "Others", "Secondary producer; limited comparable operating series"),
    ("Godawari Power",    "GPIL",       "Long / Pellet",   "DRI-EAF",        "Others", "Secondary producer; steel volumes not disclosed"),
    ("APL Apollo",        "APLAPOLLO",  "Structural tubes","ERW converter",  "Others", "Downstream tube converter, not a crude steel producer"),
    ("Kirloskar Ferrous", "KIRLFER",    "Pig iron / Cast", "Mini blast furnace","Others","Pig iron, castings and tubes - not a crude steel producer"),
    ("Others - secondary","-",          "Long / Flat",     "Mixed",          "Others", "The wider secondary and mid-tier sector (~73 Mt of finished output in FY2026)"),
]
r = 4
header_row(ws, r, hdr); r += 1
for i, row in enumerate(companies):
    for j, v in enumerate(row, start=1):
        c = ws.cell(r, j, v)
        c.border = BORDER
        c.alignment = LEFT if j in (1, 3, 4, 6) else CENTER
        if j == 5:  # Included?
            if v == "Core":
                c.font = _font(10, True, "375623"); c.fill = fill("E2EFDA")
            else:
                c.font = _font(10, False, "833C00"); c.fill = fill("FCE4D6")
        else:
            c.font = _font(10, False)
            if i % 2 == 1:
                c.fill = fill(ZEBRA)
    ws.row_dimensions[r].height = 15
    r += 1
r += 1
ws.cell(r, 1, "Core = individual analysis in the model. Others = grouped throughout the model. AM/NS India and RINL are large producers but lack a consistent FY2021-FY2026 disclosed series, so they sit in Others.")
ws.merge_cells(f"A{r}:F{r}")
ws.cell(r,1).font = _font(9, False, "595959"); ws.cell(r,1).alignment = LEFTW
ws.row_dimensions[r].height = 28
ws.freeze_panes = "A3"

# ----------------------------------------------------------------------------
# Generic builder for standard "Series|Unit|FY..|Code" sheets
# ----------------------------------------------------------------------------
def build_std(sheet_name, title, sections):
    """
    sections: list of (section_title, rows) where each row is
              (series, unit, v21, v22, v23, v24, v25, v26, code, fmt)
              fmt is an excel number format string applied to the value cols.
    """
    ws = wb.create_sheet(sheet_name)
    std_sheet(ws)
    banner(ws, title, len(STD_HEADERS))
    r = 4
    for sec_title, rows in sections:
        if sec_title:
            section_title(ws, r, sec_title, len(STD_HEADERS)); r += 1
        header_row(ws, r, STD_HEADERS); r += 1
        for i, row in enumerate(rows):
            series, unit, *vals, code, fmt = row
            cells = [series, unit] + vals + [code]
            for j, v in enumerate(cells, start=1):
                c = ws.cell(r, j, v)
                c.border = BORDER
                if j == 1:
                    c.font = _font(10, False); c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                elif j == len(cells):
                    c.font = _font(9, False, "375623"); c.alignment = CENTER; c.fill = fill(CODEFILL)
                else:
                    c.font = _font(10, False); c.alignment = CENTER
                    if j >= 3 and j <= 8 and isinstance(v, (int, float)):
                        c.number_format = fmt
                if i % 2 == 1 and 1 <= j <= 8:
                    c.fill = fill(ZEBRA)
            r += 1
        r += 1
    return ws

# ----------------------------------------------------------------------------
# 2. MACROECONOMIC   (cross-check addition the IFM Macro Model needs)
# ----------------------------------------------------------------------------
build_std("Macroeconomic", "2.  MACROECONOMIC", [
    ("SECTION A - MACRO DRIVERS USED BY THE MODEL", [
        ("India real GDP growth", "%", -0.058, 0.097, 0.070, 0.082, 0.065, 0.076, "S52", F_PCT),
        ("CPI inflation (average)", "%", 0.062, 0.055, 0.067, 0.054, 0.046, 0.037, "S51", F_PCT),
        ("USD/INR (fiscal-year average)", "Rs/US$", 74.2, 74.5, 80.4, 82.8, 84.6, 86.5, "S51", F_USD),
        ("10-year G-sec yield (year-end)", "%", 0.0618, 0.0684, 0.0731, 0.0706, 0.0658, 0.0660, "S51", F_PCT),
        ("RBI repo rate (year-end)", "%", 0.040, 0.040, 0.065, 0.065, 0.0625, 0.060, "S51", F_PCT),
    ]),
])

# ----------------------------------------------------------------------------
# 3. PRODUCTION CAPACITY
# ----------------------------------------------------------------------------
build_std("Production Capacity", "3.  PRODUCTION CAPACITY", [
    ("SECTION A - INDIA CRUDE STEEL CAPACITY", [
        ("India crude steel capacity", "Mtpa", 142.29, 154.06, 161.30, 179.51, 200.33, 220.40, "S02", F_MTPA),
        ("Capacity added during the year", "Mtpa", None, 11.77, 7.24, 18.21, 20.82, 20.07, "S02", F_MTPA),
        ("Crude steel production", "Mt", 103.54, 120.29, 127.20, 144.30, 152.18, 168.42, "S01", F_MT),
        ("Capacity utilisation (crude steel)", "%", 0.728, 0.781, 0.789, 0.804, 0.760, 0.764, "S01", F_PCT),
    ]),
])

# ----------------------------------------------------------------------------
# 4. PRODUCTION VOLUME
# ----------------------------------------------------------------------------
build_std("Production Volume", "4.  PRODUCTION VOLUME", [
    ("SECTION A - INDIA OUTPUT BY PRODUCT", [
        ("Crude steel production", "Mt", 103.54, 120.29, 127.20, 144.30, 152.18, 168.42, "S01", F_MT),
        ("Total finished steel production", "Mt", 96.20, 113.60, 123.20, 139.15, 146.69, 161.74, "S02", F_MT),
        ("Sponge iron (DRI) production", "Mt", 34.90, 39.20, 43.62, 51.56, 55.76, 60.30, "S02", F_MT),
        ("Pig iron production", "Mt", 5.00, 6.26, 5.86, 7.36, 8.33, 8.37, "S02", F_MT),
    ]),
])

# ----------------------------------------------------------------------------
# 5. DEMAND & CONSUMPTION
# ----------------------------------------------------------------------------
build_std("Demand & Consumption", "5.  DEMAND & CONSUMPTION", [
    ("SECTION A - APPARENT FINISHED STEEL CONSUMPTION", [
        ("Apparent finished steel consumption", "Mt", 94.89, 105.75, 119.86, 136.29, 152.13, 164.19, "S02", F_MT),
        ("Consumption growth", "%", -0.053, 0.114, 0.133, 0.137, 0.116, 0.079, "S02", F_PCT),
        ("Per-capita finished steel consumption", "kg", 67.9, 75.1, 84.7, 97.7, 108.3, 115.7, "S03", F_KG),
    ]),
])

# ----------------------------------------------------------------------------
# 6. IMPORTS & EXPORTS
# ----------------------------------------------------------------------------
build_std("Imports & Exports", "6.  IMPORTS & EXPORTS", [
    ("SECTION A - FINISHED STEEL TRADE (VOLUME)", [
        ("Finished steel imports", "Mt", 4.75, 4.67, 6.02, 8.32, 9.551, 6.524, "S02", F_MT),
        ("Finished steel exports", "Mt", 10.78, 13.49, 6.72, 7.487, 4.858, 6.602, "S02", F_MT),
        ("Net trade position (exports - imports)", "Mt", 6.03, 8.82, 0.70, -0.833, -4.693, 0.078, "S02", F_MT),
        ("Import penetration (imports / consumption)", "%", 0.050, 0.044, 0.050, 0.061, 0.063, 0.040, "S02", F_PCT),
    ]),
])

# ----------------------------------------------------------------------------
# 7. STEEL PRICES
# ----------------------------------------------------------------------------
build_std("Steel Prices", "7.  STEEL PRICES", [
    ("SECTION A - DOMESTIC STEEL PRICES (FISCAL-YEAR AVERAGE, EX-GST)", [
        ("HRC (hot rolled coil)", "Rs/t", 46500, 63000, 57500, 54500, 50000, 52000, "S48", F_RS0),
        ("CRC (cold rolled coil)", "Rs/t", 53000, 70000, 64000, 61000, 56500, 58500, "S48", F_RS0),
        ("Rebar (TMT)", "Rs/t", 45000, 57000, 56000, 53000, 49500, 51500, "S48", F_RS0),
        ("Wire rod", "Rs/t", 47000, 59000, 58000, 55000, 51500, 53500, "S48", F_RS0),
        ("Plate", "Rs/t", 50500, 67000, 61500, 58500, 54000, 56000, "S48", F_RS0),
        ("Billet (semi-finished)", "Rs/t", 38000, 47000, 45000, 43000, 40000, 42000, "S48", F_RS0),
        ("Blended industry ASP (realisation)", "Rs/t", 48000, 63000, 58000, 55000, 52000, 59974, "S48", F_RS0),
    ]),
])

# ----------------------------------------------------------------------------
# 8. RAW MATERIALS
# ----------------------------------------------------------------------------
build_std("Raw Materials", "8.  RAW MATERIALS", [
    ("SECTION A - IRON-BEARING MATERIALS", [
        ("Iron ore (62% Fe fines, CFR China)", "US$/dmt", 128.03, 155.52, 117.20, 119.91, 103.97, 100.52, "S17", F_USD),
        ("Iron ore pellet (domestic, Fe 63-64%)", "Rs/t", 9000, 14000, 11000, 10500, 9500, 9000, "S50", F_RS0),
    ]),
    ("SECTION B - CARBON & REDUCTANTS", [
        ("Premium hard coking coal (FOB Australia)", "US$/t", 130, 340, 300, 296, 240, 205, "S56", F_USD),
        ("PCI coal (FOB Australia)", "US$/t", 81, 211, 186, 184, 149, 127, "S56", F_USD),
        ("Thermal coal (FOB Newcastle, 6000 kcal)", "US$/t", 68, 170, 330, 145.1, 131.89, 111.5, "S17", F_USD),
        ("Steel scrap (shredded, CFR Nhava Sheva)", "US$/t", 430, 530, 450, 400, 380, 365, "S49", F_USD),
    ]),
    ("SECTION C - FLUXES, ALLOYS & ENERGY", [
        ("Limestone (delivered)", "Rs/t", 1450, 1520, 1620, 1700, 1760, 1800, "S50", F_RS0),
        ("Dolomite (delivered)", "Rs/t", 1668, 1748, 1863, 1955, 2024, 2070, "S50", F_RS0),
        ("Ferro manganese", "Rs/t", 75000, 105000, 95000, 90000, 92000, 95000, "S48", F_RS0),
        ("Ferro silicon", "Rs/t", 90000, 140000, 120000, 110000, 105000, 108000, "S48", F_RS0),
        ("Natural gas (domestic APM)", "US$/mmBtu", 2.39, 2.90, 7.01, 6.50, 6.50, 6.75, "S53", F_RS2),
        ("Power (industrial tariff)", "Rs/kWh", 7.0, 7.2, 7.6, 7.9, 8.1, 8.3, "S54", F_RS2),
        ("Freight / logistics (per t of steel)", "Rs/t", 2800, 2950, 3100, 3250, 3350, 3400, "S48", F_RS0),
    ]),
])

# ----------------------------------------------------------------------------
# 9. EBITDA  (Revenue / EBITDA / Margin / EBITDA per tonne)
# ----------------------------------------------------------------------------
ws = wb.create_sheet("EBITDA")
std_sheet(ws)
banner(ws, "9.  EBITDA", len(STD_HEADERS))
r = 4

CORE = ["Tata Steel", "JSW Steel", "SAIL", "Jindal Steel", "Jindal Stainless"]

rev = {
    "Tata Steel":       [156294, 243959, 243353, 229171, 218543, 232140],
    "JSW Steel":        [79839, 146371, 165960, 175006, 168824, 185470],
    "SAIL":             [69114, 103477, 104448, 105378, 102479, 110811],
    "Jindal Steel":     [38933, 51166, 53212, 50354, 50129, 53225],
    "Jindal Stainless": [11679, 32733, 35697, 38562, 39312, 42955],
}
ebitda = {
    "Tata Steel":       [30892, 63490, 32300, 22248, 25298, 34352],
    "JSW Steel":        [20141, 39114, 18470, 28157, 22725, 29464],
    "SAIL":             [15583, 21363, 8038, 11149, 10690, 12000],
    "Jindal Steel":     [15254, 15559, 9942, 10202, 9488, 9644],
    "Jindal Stainless": [1364, 5090, 3586, 4511, 4469, 5560],
}
codes = {"Tata Steel": "S08", "JSW Steel": "S09", "SAIL": "S10",
         "Jindal Steel": "S11", "Jindal Stainless": "S12"}
ebitda_t = {  # FY25, FY26 disclosed only
    "Tata Steel": (13983, 15213), "JSW Steel": (8661, 10064), "SAIL": (6576, 6596),
    "Jindal Steel": (11712, 10482), "Jindal Stainless": (19622, 21634),
}

# --- Derived India industry totals so that Top 5 + Others = Total ----------
finished = [96.20, 113.60, 123.20, 139.15, 146.69, 161.74]   # Mt finished production (S02)
asp      = [48000, 63000, 58000, 55000, 52000, 59974]        # Rs/t blended realisation (S48)
ind_eb_t = [9000, 17000, 8500, 9500, 9500, 10733]            # Rs/t industry EBITDA per tonne

def _r(x): return int(round(x))
total_rev = [_r(finished[i] * asp[i] / 10) for i in range(6)]
total_eb  = [_r(finished[i] * ind_eb_t[i] / 10) for i in range(6)]
top5_rev  = [_r(sum(rev[n][i] for n in CORE)) for i in range(6)]
top5_eb   = [_r(sum(ebitda[n][i] for n in CORE)) for i in range(6)]
rev["Others (India, residual)"]    = [total_rev[i] - top5_rev[i] for i in range(6)]
rev["Total Industry (India)"]      = total_rev
ebitda["Others (India, residual)"] = [total_eb[i] - top5_eb[i] for i in range(6)]
ebitda["Total Industry (India)"]   = total_eb
codes["Others (India, residual)"] = "S57"
codes["Total Industry (India)"]   = "S57"

order_rev = CORE + ["Others (India, residual)", "Total Industry (India)"]

def company_table(ws, r, title, dct, fmt, codes, order):
    section_title(ws, r, title, len(STD_HEADERS)); r += 1
    hdr = ["Company / Segment", "Unit"] + FY + ["Source Code"]
    header_row(ws, r, hdr); r += 1
    for i, name in enumerate(order):
        vals = dct[name]
        unit = "%" if fmt == F_PCT else ("Rs/t" if "PER TONNE" in title else "Rs crore")
        is_total = name.startswith("Total Industry")
        cells = [name, unit] + list(vals) + [codes[name]]
        for j, v in enumerate(cells, start=1):
            c = ws.cell(r, j, v)
            c.border = BORDER
            if j == 1:
                c.font = _font(10, is_total); c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif j == len(cells):
                c.font = _font(9, False, "375623"); c.alignment = CENTER; c.fill = fill(CODEFILL)
            else:
                c.font = _font(10, is_total); c.alignment = CENTER
                if 3 <= j <= 8 and isinstance(v, (int, float)):
                    c.number_format = fmt
            if is_total and j <= 8:
                c.fill = fill(LIGHT)
            elif i % 2 == 1 and 1 <= j <= 8:
                c.fill = fill(ZEBRA)
        r += 1
    return r + 1

# Revenue
r = company_table(ws, r, "SECTION A - REVENUE FROM OPERATIONS (Rs crore)", rev, F_RS0, codes, order_rev)
# EBITDA
r = company_table(ws, r, "SECTION B - EBITDA (Rs crore)", ebitda, F_RS0, codes, order_rev)

# Margin
margin = {}
for name in order_rev:
    margin[name] = [round(ebitda[name][i] / rev[name][i], 4) if rev[name][i] else None for i in range(6)]
r = company_table(ws, r, "SECTION C - EBITDA MARGIN (%)", margin, F_PCT, codes, order_rev)

# EBITDA per tonne
section_title(ws, r, "SECTION D - EBITDA PER TONNE (Rs/t)", len(STD_HEADERS)); r += 1
hdr = ["Company / Segment", "Unit"] + FY + ["Source Code"]
header_row(ws, r, hdr); r += 1
etrows = [(n, (None, None, None, None, ebitda_t[n][0], ebitda_t[n][1]), codes[n]) for n in CORE]
etrows.append(("Industry EBITDA per tonne (India)", tuple(ind_eb_t), "S57"))
for i, (name, vals, code) in enumerate(etrows):
    is_total = name.startswith("Industry")
    cells = [name, "Rs/t"] + list(vals) + [code]
    for j, v in enumerate(cells, start=1):
        c = ws.cell(r, j, v)
        c.border = BORDER
        if j == 1:
            c.font = _font(10, is_total); c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        elif j == len(cells):
            c.font = _font(9, False, "375623"); c.alignment = CENTER; c.fill = fill(CODEFILL)
        else:
            c.font = _font(10, is_total); c.alignment = CENTER
            if 3 <= j <= 8 and isinstance(v, (int, float)):
                c.number_format = F_RS0
        if is_total and j <= 8:
            c.fill = fill(LIGHT)
        elif i % 2 == 1 and 1 <= j <= 8:
            c.fill = fill(ZEBRA)
    r += 1
r += 1
note = ("BASIS: Core 5 revenue and EBITDA are as reported (consolidated). 'Total Industry (India)' is DERIVED - India finished "
        "steel production x blended realisation for revenue, and x industry EBITDA per tonne for EBITDA (code S57) - because no single "
        "official industry-wide revenue or EBITDA figure is published. 'Others (India, residual)' = Total Industry minus the Top 5, so "
        "the Core 5 plus Others always tie exactly to the Total. Tata and JSW report consolidated global figures, so treat the residual as "
        "indicative. EBITDA per tonne is company-disclosed for FY2025-FY2026 only; comparability notes (Jindal Steel power demerger FY2022; "
        "Jindal Stainless Hisar merger FY2024) are on the Conflicts Log.")
ws.merge_cells(f"A{r}:I{r}")
ws.cell(r, 1, note).font = _font(9, False, "595959")
ws.cell(r, 1).alignment = LEFTW
ws.row_dimensions[r].height = 64

# ----------------------------------------------------------------------------
# 10. INDUSTRY KPIs
# ----------------------------------------------------------------------------
build_std("Industry KPIs", "10.  INDUSTRY KPIs", [
    ("SECTION A - HEADLINE INDICATORS", [
        ("Crude steel production", "Mt", 103.54, 120.29, 127.20, 144.30, 152.18, 168.42, "S01", F_MT),
        ("Crude steel capacity", "Mtpa", 142.29, 154.06, 161.30, 179.51, 200.33, 220.40, "S02", F_MTPA),
        ("Capacity utilisation", "%", 0.728, 0.781, 0.789, 0.804, 0.760, 0.764, "S01", F_PCT),
        ("Finished steel production", "Mt", 96.20, 113.60, 123.20, 139.15, 146.69, 161.74, "S02", F_MT),
        ("Apparent finished steel consumption", "Mt", 94.89, 105.75, 119.86, 136.29, 152.13, 164.19, "S02", F_MT),
        ("Finished steel imports", "Mt", 4.75, 4.67, 6.02, 8.32, 9.551, 6.524, "S02", F_MT),
        ("Finished steel exports", "Mt", 10.78, 13.49, 6.72, 7.487, 4.858, 6.602, "S02", F_MT),
        ("Iron ore (62% Fe CFR China)", "US$/dmt", 128.03, 155.52, 117.20, 119.91, 103.97, 100.52, "S17", F_USD),
        ("Blended industry ASP", "Rs/t", 48000, 63000, 58000, 55000, 52000, 59974, "S48", F_RS0),
    ]),
])

# ----------------------------------------------------------------------------
# 11. GOVERNMENT POLICIES
# ----------------------------------------------------------------------------
ws = wb.create_sheet("Government Policies")
ws.sheet_view.showGridLines = False
set_widths(ws, {"A": 34, "B": 16, "C": 16, "D": 60, "E": 13})
banner(ws, "11.  GOVERNMENT POLICIES", 5)
r = 4
hdr = ["Measure", "Effective", "Value / Rate", "What it does", "Source Code"]
pol = [
    ("Safeguard duty on flat steel", "21-Apr-2025", "12% -> 11.5% -> 11%", "Definitive safeguard duty on non-alloy/alloy flat products; 12% year 1, 11.5% year 2 (from 21-Apr-2026), 11% year 3; expires 20-Apr-2028", "S27"),
    ("PLI Specialty Steel", "To Mar-2026", "Rs 23,827 cr realised", "Production-Linked Incentive; Rs 23,827 cr investment realised against Rs 55,993 cr committed", "S24"),
    ("Green steel taxonomy", "23-Dec-2024", "2.2 tCO2e/tfs", "Steel below 2.2 tCO2e per tonne of finished steel qualifies as 'green'; star-rating bands apply", "S25"),
    ("National Steel Policy target", "By FY2031", "300 Mtpa", "Policy path to 300 Mtpa crude steel capacity and ~160 kg per-capita consumption", "S03"),
]
header_row(ws, r, hdr); r += 1
for i, row in enumerate(pol):
    for j, v in enumerate(row, start=1):
        c = ws.cell(r, j, v)
        c.border = BORDER
        c.alignment = LEFTW if j in (1, 4) else CENTER
        if j == 5:
            c.font = _font(9, False, "375623"); c.fill = fill(CODEFILL)
        else:
            c.font = _font(10, False)
            if i % 2 == 1:
                c.fill = fill(ZEBRA)
    ws.row_dimensions[r].height = 30
    r += 1
ws.freeze_panes = "A3"

# ----------------------------------------------------------------------------
# 12. SOURCES
# ----------------------------------------------------------------------------
ws = wb.create_sheet("Sources")
ws.sheet_view.showGridLines = False
set_widths(ws, {"A": 8, "B": 42, "C": 30, "D": 13, "E": 52, "F": 12, "G": 11})
banner(ws, "12.  SOURCES", 7)
r = 4
hdr = ["Code", "Title / Publication", "Organisation", "Date", "URL", "Access Date", "Confidence"]
sources = [
    ("S01", "Monthly Economic Report / An Overview of Steel Sector", "Ministry of Steel (JPC)", "Jul-2026", "https://steel.gov.in", "10-Aug-2026", "High"),
    ("S02", "Development of Indian Steel Sector since 2014-15", "Ministry of Steel (JPC)", "Jun-2026", "https://steel.gov.in", "10-Aug-2026", "High"),
    ("S03", "An Overview of Steel Sector (capacity, per-capita)", "Ministry of Steel", "Jul-2026", "https://steel.gov.in", "10-Aug-2026", "High"),
    ("S07", "Monthly crude steel production statistics", "World Steel Association", "Apr-2026", "https://worldsteel.org", "10-Aug-2026", "High"),
    ("S08", "4QFY2026 results press release & investor presentation", "Tata Steel Limited", "15-May-2026", "https://tatasteel.com", "10-Aug-2026", "High"),
    ("S09", "Q4 FY26 investor / analyst presentation", "JSW Steel Limited", "14-May-2026", "https://jswsteel.in", "10-Aug-2026", "High"),
    ("S10", "Audited results & press release FY2026", "Steel Authority of India (SAIL)", "15-May-2026", "https://sail.co.in", "10-Aug-2026", "High"),
    ("S11", "Q4 FY26 investor presentation", "Jindal Steel Limited", "02-May-2026", "https://jindalsteelpower.com", "10-Aug-2026", "High"),
    ("S12", "4QFY26 earnings call & results", "Jindal Stainless Limited", "09-May-2026", "https://jindalstainless.com", "10-Aug-2026", "High"),
    ("S17", "Commodity Price Data ('Pink Sheet') - monthly", "World Bank", "02-Jul-2026", "https://worldbank.org/commodities", "10-Aug-2026", "High"),
    ("S18", "Commodity Price Data ('Pink Sheet') - annual", "World Bank", "03-Mar-2026", "https://worldbank.org/commodities", "10-Aug-2026", "High"),
    ("S19", "Indian steel sector research (HRC price path)", "ICRA Limited", "Apr-2026", "https://icra.in", "10-Aug-2026", "Medium"),
    ("S22", "India steel index - BF rebar benchmark", "BigMint (formerly SteelMint)", "31-Jul-2026", "https://bigmint.co", "10-Aug-2026", "Medium"),
    ("S23", "Iron ore price revision intimations (Bailadila grades)", "NMDC Limited", "2026 (various)", "https://nmdc.co.in", "10-Aug-2026", "High"),
    ("S24", "Lok Sabha Q5098 - PLI Specialty Steel", "Ministry of Steel", "24-Mar-2026", "https://sansad.in", "10-Aug-2026", "High"),
    ("S25", "Rajya Sabha Q3350 - Green Steel Taxonomy", "Ministry of Steel", "20-Mar-2026", "https://steel.gov.in", "10-Aug-2026", "High"),
    ("S26", "Rajya Sabha Q3982 - SAIL plant-wise capacity", "Ministry of Steel", "27-Mar-2026", "https://steel.gov.in", "10-Aug-2026", "High"),
    ("S27", "Safeguard duty notification (flat steel)", "Ministry of Finance", "30-Dec-2025", "https://cbic.gov.in", "10-Aug-2026", "High"),
    ("S30", "Aggregated exchange filings (secondary producers)", "NSE/BSE filings & aggregators", "FY2022-FY2026", "https://nseindia.com", "10-Aug-2026", "Medium"),
    ("S48", "Domestic steel & ferro-alloy price assessments (FY averages reconstructed from monthly assessments)", "SteelMint / BigMint / MEPS", "FY2021-FY2026", "https://bigmint.co", "10-Aug-2026", "Medium"),
    ("S49", "Steel scrap shredded index, CFR Nhava Sheva", "Fastmarkets / SteelMint", "FY2021-FY2026", "https://fastmarkets.com", "10-Aug-2026", "Medium"),
    ("S50", "Iron ore pellet & flux domestic price assessments", "NMDC / SteelMint / IBM", "FY2021-FY2026", "https://nmdc.co.in", "10-Aug-2026", "Medium"),
    ("S51", "Database on Indian Economy - policy rates, CPI, FX", "Reserve Bank of India", "FY2021-FY2026", "https://rbi.org.in", "10-Aug-2026", "High"),
    ("S52", "National Accounts Statistics - real GDP growth", "MoSPI, Government of India", "FY2021-FY2026", "https://mospi.gov.in", "10-Aug-2026", "High"),
    ("S53", "Domestic natural gas (APM) price ceiling", "PPAC / MoPNG", "FY2021-FY2026", "https://ppac.gov.in", "10-Aug-2026", "Medium"),
    ("S54", "Average industrial power tariff", "Central Electricity Authority", "FY2021-FY2026", "https://cea.nic.in", "10-Aug-2026", "Low"),
    ("S56", "Premium HCC & PCI FOB Australia (FY averages reconstructed)", "World Bank / SteelMint / Platts", "FY2021-FY2026", "https://worldbank.org/commodities", "10-Aug-2026", "Medium"),
    ("S57", "Derived India industry total = finished steel production x blended realisation (revenue) and x industry EBITDA/t (EBITDA); Others = Total minus Top 5", "Derived (JPC production x S48 realisation)", "FY2021-FY2026", "https://steel.gov.in", "10-Aug-2026", "Low"),
]
header_row(ws, r, hdr); r += 1
for i, row in enumerate(sources):
    for j, v in enumerate(row, start=1):
        c = ws.cell(r, j, v)
        c.border = BORDER
        c.alignment = LEFTW if j in (2, 3, 5) else CENTER
        if j == 1:
            c.font = _font(9, True, "375623")
        else:
            c.font = _font(9, False)
        if i % 2 == 1:
            c.fill = fill(ZEBRA)
    ws.row_dimensions[r].height = 26
    r += 1
ws.freeze_panes = "A3"

# ----------------------------------------------------------------------------
# 13. CONFLICTS LOG
# ----------------------------------------------------------------------------
ws = wb.create_sheet("Conflicts Log")
ws.sheet_view.showGridLines = False
set_widths(ws, {"A": 8, "B": 34, "C": 70, "D": 40})
banner(ws, "13.  CONFLICTS LOG", 4)
r = 4
hdr = ["ID", "Topic", "The conflict", "How it is handled here"]
conf = [
    ("C01", "Coking coal FY average", "Ministry of Steel (~US$225/t, Mar-2026) and IDBI (US$186/t, Mar-2026) differ materially for the same month.", "No single official FY average exists; the FY series is a reconstructed average flagged Medium (S56)."),
    ("C07", "Domestic HRC spot price", "Mar-2026 HRC quoted at Rs 55,900, 57,700 and 59,500/t by different assessors.", "Fiscal-year averages are used (not spot), reconstructed from monthly assessments (S48, Medium)."),
    ("C10", "Jindal Steel FY2021 basis", "FY2021 consolidated figures include the power business that was demerged in FY2022.", "FY2021 shown as reported; comparability break noted. Later years exclude power."),
    ("C11", "Jindal Stainless FY2021 basis", "FY2021 pre-dates the Jindal Stainless Hisar merger completed in FY2024.", "FY2021 (Rs 11,679 cr) is the pre-merger entity; break noted; FY2024+ is merged."),
    ("C12", "SAIL EBITDA basis", "SAIL headline EBITDA is standalone; the standardised series is on a consolidated basis.", "Standardised (consolidated) EBITDA used for comparability (S10)."),
    ("C14", "Reconstructed FY averages", "Several raw-material and steel-price series are not published as clean official FY averages.", "Presented as fiscal-year averages with explicit Medium/Low confidence on the Sources sheet; never asserted as official."),
]
header_row(ws, r, hdr); r += 1
for i, row in enumerate(conf):
    for j, v in enumerate(row, start=1):
        c = ws.cell(r, j, v)
        c.border = BORDER
        c.alignment = LEFTW if j in (2, 3, 4) else CENTER
        if j == 1:
            c.font = _font(9, True, "833C00")
        else:
            c.font = _font(9, False)
        if i % 2 == 1:
            c.fill = fill(ZEBRA)
    ws.row_dimensions[r].height = 40
    r += 1
ws.freeze_panes = "A3"

# ----------------------------------------------------------------------------
wb.save("Master Industry Database.xlsx")
print("Saved Master Industry Database.xlsx with", len(wb.sheetnames), "sheets:")
for s in wb.sheetnames:
    print("  -", s)
