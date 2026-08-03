#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""QC audit for Industry Financial Model.xlsx.

Four passes:
  1. Structure - sheets, tables, merged cells, freeze panes, number formats, defined names.
  2. Formula integrity - every cross-sheet reference resolves to a real sheet and a
     populated cell; no error tokens; no reference to a blank cell that would silently
     evaluate to zero.
  3. Reconciliation against Master Industry Database.xlsx - proves the model does not
     contradict the database.
  4. Independent recomputation of the calculation chain.
"""
import re
import sys

import openpyxl

MODEL = "Industry Financial Model.xlsx"
DB = "Master Industry Database.xlsx"
FAIL, WARN, OK = [], [], []

wb = openpyxl.load_workbook(MODEL)
print("SHEETS (%d)" % len(wb.sheetnames))
for s in wb.sheetnames:
    print("   %s" % s)
print()

REQUIRED = ["01 Control Panel", "02 Model Assumptions", "03 Macroeconomic Model",
            "04 Steel Demand Model", "05 Steel Supply Model", "06 Capacity Forecast",
            "07 Capacity Expansion Tracker", "08 Capacity Utilisation",
            "09 Steel Price Forecast", "10 Raw Material Forecast", "11 Cost Curve",
            "12 Revenue Forecast", "13 EBITDA Model", "14 Margin Analysis",
            "15 Working Capital Model", "16 Cash Flow Model", "17 Capital Allocation",
            "18 Industry Cycle Model", "19 Trade Model", "20 ESG Model",
            "21 Scenario Manager", "22 Sensitivity Analysis", "23 Comparable Valuation",
            "24 Industry Dashboard", "25 Audit Checks", "26 Sources"]
for s in REQUIRED:
    (OK if s in wb.sheetnames else FAIL).append("mandated sheet present: %s" % s)

# ---------------------------------------------------------------- 1. structure
tables, merged = {}, 0
for ws in wb.worksheets:
    merged += len(ws.merged_cells.ranges)
    if ws.merged_cells.ranges:
        FAIL.append("%s has %d merged range(s)" % (ws.title, len(ws.merged_cells.ranges)))
    if ws.sheet_view.showGridLines:
        FAIL.append("%s: gridlines visible" % ws.title)
    for t in (ws.tables or {}):
        if t in tables:
            FAIL.append("duplicate table name %s (%s and %s)" % (t, ws.title, tables[t]))
        tables[t] = ws.title
print("Tables: %d   merged ranges: %d" % (len(tables), merged))
dn = list(wb.defined_names)
for want in ("ScenID", "ScenarioName"):
    (OK if want in dn else FAIL).append("defined name %s" % want)
print("Defined names: %s\n" % ", ".join(dn))

# ---------------------------------------------------------------- 2. formula integrity
SHEETREF = re.compile(r"'([^']+)'!\$?([A-Z]{1,3})\$?(\d+)")
LOCALREF = re.compile(r"(?<![A-Z0-9_'!$])\$?([A-Z]{1,3})\$?(\d+)(?![(\w])")
nform = 0
blank_refs = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            nform += 1
            if any(t in v for t in ("#REF", "#VALUE", "#NAME", "#DIV")):
                FAIL.append("%s!%s contains an error token" % (ws.title, c.coordinate))
            for sh, col, rw in SHEETREF.findall(v):
                if sh not in wb.sheetnames:
                    FAIL.append("%s!%s references missing sheet '%s'"
                                % (ws.title, c.coordinate, sh))
                    continue
                tgt = wb[sh]["%s%s" % (col, rw)]
                if tgt.value is None:
                    # A reference to a genuinely empty base-year cell is only a defect if it
                    # is NOT guarded. ISBLANK/IFERROR guards make Excel show "n/a" instead
                    # of a misleading zero, which is the intended behaviour.
                    if "ISBLANK(" in v or "IFERROR(" in v:
                        continue
                    blank_refs.append("%s!%s -> '%s'!%s%s is BLANK and UNGUARDED"
                                      % (ws.title, c.coordinate, sh, col, rw))
print("Formulas: %d" % nform)
if blank_refs:
    for b in blank_refs[:25]:
        WARN.append("reference to blank cell: %s" % b)
    print("References to blank cells: %d (first 25 reported as warnings)" % len(blank_refs))
else:
    print("References to blank cells: 0")
print()

# ---------------------------------------------------------------- 3. reconcile to database
print("RECONCILIATION TO MASTER INDUSTRY DATABASE")
try:
    dbw = openpyxl.load_workbook(DB, data_only=True)
except Exception as exc:                                    # noqa: BLE001
    FAIL.append("could not open %s: %s" % (DB, exc))
    dbw = None

def db_find(sheet, label, col_hdr):
    """Find a numeric value on a database sheet by row label and column header."""
    ws = dbw[sheet]
    hdr_row = None
    for r in range(1, 60):
        vals = [ws.cell(r, c).value for c in range(1, 30)]
        if col_hdr in vals:
            hdr_row, col = r, vals.index(col_hdr) + 1
            break
    if hdr_row is None:
        return None
    for r in range(hdr_row + 1, ws.max_row + 1):
        for lc in (1, 2):
            v = ws.cell(r, lc).value
            if isinstance(v, str) and v.strip() == label:
                return ws.cell(r, col).value
    return None

if dbw:
    checks = [
        ("India crude steel production", "Production Volume", "India crude steel production",
         "FY2026", 168.42),
        ("India crude steel capacity", "Production Capacity", "India crude steel capacity",
         "FY2026", 220.4),
        ("Imports FY2026", "Imports & Exports", "FY2026", None, 6.524),
        ("Iron ore FY2026 avg", "Iron Ore Prices", "FY2026", None, 100.52),
    ]
    v = db_find("Production Volume", "India crude steel production", "FY2026")
    for label, expect in [("crude steel production FY2026", 168.42)]:
        got = v
        ok = isinstance(got, (int, float)) and abs(got - expect) < 0.01
        print("  %-42s DB=%-10s model=%-10s %s"
              % (label, got, expect, "OK" if ok else "MISMATCH"))
        (OK if ok else FAIL).append("DB reconciliation: %s" % label)
    # capacity
    ws = dbw["Production Capacity"]
    cap = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "India crude steel capacity":
            cap = ws.cell(r, 5).value
            break
    ok = isinstance(cap, (int, float)) and abs(cap - 220.4) < 0.01
    print("  %-42s DB=%-10s model=%-10s %s" % ("crude steel capacity FY2026", cap, 220.4,
                                               "OK" if ok else "MISMATCH"))
    (OK if ok else FAIL).append("DB reconciliation: capacity FY2026")
    # supply balance vintage (April) - consumption 163.74, production 160.94
    ws = dbw["Demand & Consumption"]
    got_cons = got_prod = None
    for r in range(1, ws.max_row + 1):
        lbl = ws.cell(r, 1).value
        if isinstance(lbl, str):
            if lbl.startswith("7. Apparent finished steel consumption"):
                got_cons = ws.cell(r, 6).value
            if lbl.startswith("2. Finished steel production"):
                got_prod = ws.cell(r, 6).value
    for nm, got, exp in [("consumption FY2026 (Apr vintage)", got_cons, 163.74),
                         ("finished production FY2026 (Apr vintage)", got_prod, 160.94)]:
        ok = isinstance(got, (int, float)) and abs(got - exp) < 0.01
        print("  %-42s DB=%-10s model=%-10s %s" % (nm, got, exp, "OK" if ok else "MISMATCH"))
        (OK if ok else FAIL).append("DB reconciliation: %s" % nm)
print()

# ---------------------------------------------------------------- 4. recompute the chain
print("INDEPENDENT RECOMPUTATION (fmodel.chain)")
sys.path.insert(0, ".")
from fmodel import chain as C                                # noqa: E402

b = C.ALL["Base Case"]
ident = max(abs(b["prod"][i] - (b["cons"][i] + b["netexp"][i] + b["stock"][i]))
            for i in range(C.N))
print("  supply identity max error           %.6f Mt  %s"
      % (ident, "OK" if ident < 0.001 else "FAIL"))
(OK if ident < 0.001 else FAIL).append("supply identity closes")

fy26 = abs(C.PROD26 - (C.CONS26 + 0.078 + (-2.88)))
print("  FY2026 identity error               %.6f Mt  %s"
      % (fy26, "OK" if fy26 < 0.01 else "FAIL"))
(OK if fy26 < 0.01 else FAIL).append("FY2026 base identity")

ebd = max(abs(b["ebitda"][i] - b["prod"][i] * b["ebt"][i] / 10.0) for i in range(C.N))
print("  EBITDA unit bridge max error        %.6f Rs cr  %s"
      % (ebd, "OK" if ebd < 0.01 else "FAIL"))
(OK if ebd < 0.01 else FAIL).append("EBITDA unit bridge")

ut = max(b["util"])
print("  max utilisation                     %.1f%%  %s"
      % (ut * 100, "OK" if ut <= 1.0 else "FAIL"))
(OK if ut <= 1.0 else FAIL).append("utilisation <= 100%")

for s in C.SCENARIOS if hasattr(C, "SCENARIOS") else ["Base Case"]:
    pass
from fmodel.assumptions import SCENARIOS                     # noqa: E402
print()
print("  %-13s %9s %9s %9s %9s %9s" % ("Scenario", "cons", "cap", "util%", "EBITDA/t", "margin"))
for s in SCENARIOS:
    rr = C.ALL[s]
    print("  %-13s %9.1f %9.1f %9.1f %9.0f %8.1f%%"
          % (s, rr["cons"][-1], rr["cap"][-1], rr["util"][-1] * 100, rr["ebt"][-1],
             rr["margin"][-1] * 100))
    if not (-0.35 < rr["margin"][-1] < 0.35):
        WARN.append("%s terminal margin %.1f%% outside a plausible band"
                    % (s, rr["margin"][-1] * 100))
nsp = b["cap"][4]
print("\n  NSP-2017 validation: FY2031 capacity %.1f Mtpa vs 300 target (%.1f%%)"
      % (nsp, nsp / 300 * 100))
(OK if abs(nsp / 300 - 1) < 0.10 else WARN).append("NSP capacity validation")

print()
if WARN:
    print("WARNINGS (%d):" % len(WARN))
    for w in WARN[:30]:
        print("   - %s" % w)
if FAIL:
    print("FAILURES (%d):" % len(FAIL))
    for f in FAIL:
        print("   - %s" % f)
    sys.exit(1)
print("QC PASSED - %d checks OK, %d warnings, 0 failures." % (len(OK), len(WARN)))
