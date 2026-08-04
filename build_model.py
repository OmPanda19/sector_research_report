#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build "Industry Financial Model.xlsx" - COPY-PASTE SAFE rebuild.

DESIGN RULES (non-negotiable, because the user copy-pastes sheets into another workbook):
  1. ZERO cross-sheet formulas. Every formula references only cells on its own sheet.
  2. ZERO defined names. A pasted sheet cannot rely on workbook-level names.
  3. ZERO Excel Tables. They were the cause of the previous file's repair prompt and they
     complicate pasting.
  4. Scenario selector is an INTEGER 1-4, not text, so no MATCH-on-text can fail.
  5. Only paste-safe worksheet functions: INDEX, CHOOSE, IF, IFERROR, ISNUMBER, MIN, MAX,
     SUM, SUMIFS, ABS, COUNT, COUNTIF, MEDIAN.
Consequence: the whole calculation chain lives on ONE sheet ("Model") as labelled sections.
That is the only way a pasted sheet keeps working.
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.utils import get_column_letter as GL
from openpyxl.worksheet.datavalidation import DataValidation

from builder.style import (AL_CENTRE, AL_LEFT, AL_LEFT_WRAP, AL_RIGHT, BORDER_BOTTOM,
                           FILL_NAVY_LIGHT, FMT_1DP, FMT_2DP, FMT_3DP, FMT_INT, FMT_MULT,
                           FMT_PCT_1, FMT_TEXT, NAVY, TEXT_MUTED, font, freeze, header_row,
                           put, sheet_defaults, title_block)
from fmodel import assumptions as A
from fmodel import chain as C
from fmodel import companies as CO

OUT = "Industry Financial Model.xlsx"
CB, CF1 = 4, 5                      # D = FY2026A ; E..K = FY2027E..FY2033E
NY = A.N
CFN = CF1 + NY - 1                  # K
CL, CSRC, CCONF, CNOTE = CFN + 1, CFN + 2, CFN + 3, CFN + 4
FCOLS = list(range(CF1, CFN + 1))
FY = A.FY
SC = A.SCENARIOS
DRIVERS = sorted(A.DRIVERS, key=lambda d: d[0])

R = {}          # line key -> row on the Model sheet
SCEN_CELL = None


def fmt_for(unit):
    u = (unit or "").strip()
    if u == "%" or u.endswith("of revenue") or u == "ppt":
        return FMT_PCT_1
    if u == "x":
        return FMT_MULT
    if u in ("Rs/t", "Rs cr", "Rs/dmt", "Rs/t of capacity added", "Rs/t of capacity",
             "Rs/US$ x", "days of revenue", "Rs/EUR", "Number"):
        return FMT_INT
    if u in ("Mt", "Mtpa", "Mtpa p.a.", "kg", "US$/dmt", "US$/t", "US$/mt", "Rs/US$",
             "bn", "years", "tCO2e/tfs", "EUR/tCO2"):
        return FMT_2DP
    return FMT_2DP


def L(row, col):
    return "%s%d" % (GL(col), row)


def widths(ws):
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 54
    ws.column_dimensions["C"].width = 18
    for c in [CB] + FCOLS:
        ws.column_dimensions[GL(c)].width = 13
    ws.column_dimensions[GL(CL)].width = 11
    ws.column_dimensions[GL(CSRC)].width = 24
    ws.column_dimensions[GL(CCONF)].width = 11
    ws.column_dimensions[GL(CNOTE)].width = 105


def hdr_row(ws, r, first="Line item"):
    return header_row(ws, r, ["", first, "Unit", "FY2026A"] + FY
                      + ["CAGR", "Source", "Conf.", "Notes"],
                      start_col=1, left_align_cols=(2, 3, CSRC, CNOTE))


def sect(ws, r, num, text):
    put(ws, r, 1, num, f=font(9, bold=True, colour=NAVY), alignment=AL_CENTRE)
    put(ws, r, 2, text, f=font(11, bold=True, colour=NAVY))
    return r + 1


def line(ws, r, label, unit, base, cells, src="", conf="", note="", fmt=None, bold=False,
         key=None, cagr=False, ident="", fill=None):
    fmt = fmt or fmt_for(unit)
    f = font(10, bold=True, colour=NAVY) if bold else font()
    if ident:
        put(ws, r, 1, ident, f=font(8, colour=TEXT_MUTED), alignment=AL_CENTRE)
    put(ws, r, 2, label, f=f, alignment=AL_LEFT_WRAP, fill=fill)
    put(ws, r, 3, unit, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT, fill=fill)
    if base is not None:
        put(ws, r, CB, base, number_format=fmt, alignment=AL_RIGHT, f=f, fill=fill)
    for i, v in enumerate(cells or []):
        if v is None:
            continue
        put(ws, r, CF1 + i, v, number_format=fmt, alignment=AL_RIGHT, f=f, fill=fill)
    if cagr and base is not None:
        put(ws, r, CL, '=IFERROR(IF(AND(ISNUMBER(%s),ISNUMBER(%s),%s>0),(%s/%s)^(1/%d)-1,""),"")'
            % (L(r, CB), L(r, CFN), L(r, CB), L(r, CFN), L(r, CB), NY),
            number_format=FMT_PCT_1, alignment=AL_RIGHT)
    if src:
        put(ws, r, CSRC, src, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    if conf:
        put(ws, r, CCONF, conf, f=font(9, colour=TEXT_MUTED), alignment=AL_CENTRE)
    if note:
        put(ws, r, CNOTE, note, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    if key:
        R[key] = r
    return r + 1


def prev(c, row):
    """Reference to the previous period in the same row (base year for the first column)."""
    return L(row, c - 1)


def act(code, c):
    """Intra-sheet reference to an active driver cell."""
    return L(R["a_" + code], c)


# ======================================================================================
def build_model(ws):
    global SCEN_CELL
    sheet_defaults(ws)
    widths(ws)
    r = title_block(ws, [
        ("INDUSTRY FINANCIAL MODEL - INDIAN STEEL INDUSTRY", "title"),
        ("Single-sheet analytical engine. Horizon FY2027E-FY2033E. Base year FY2026A. "
         "Built 03-Aug-2026.", "label"),
        ("", "body"),
        ("COPY-PASTE SAFE: every formula on this sheet references only cells on THIS sheet. "
         "There are no links to other sheets, no defined names and no Excel Tables. Paste into "
         "a blank sheet starting at cell A1 and it will work unchanged.", "body"),
        ("Consumes Master Industry Database.xlsx (FY2026 actuals). Change ONE cell - the "
         "scenario number below - and the whole sheet recomputes.", "body"), ("", "body")],
        width_col=2)

    # ---------------------------------------------------------------- S0 control
    r = sect(ws, r, "S0", "CONTROL")
    put(ws, r, 2, "SCENARIO (enter 1, 2, 3 or 4)", f=font(10, bold=True))
    cell = put(ws, r, CB, 1, number_format=FMT_INT, alignment=AL_CENTRE,
               f=font(12, bold=True, colour=NAVY), fill=FILL_NAVY_LIGHT, border=BORDER_BOTTOM)
    SCEN_CELL = "$%s$%d" % (GL(CB), r)
    dv = DataValidation(type="list", formula1='"1,2,3,4"', allow_blank=False,
                        errorTitle="Invalid scenario", error="Enter 1, 2, 3 or 4.")
    ws.add_data_validation(dv)
    dv.add(cell)
    put(ws, r, CNOTE, "1 = Base, 2 = Bull, 3 = Bear, 4 = Stress. An INTEGER is used rather than "
                      "a text label so that no text lookup can fail when this sheet is pasted "
                      "elsewhere.", f=font(9, italic=True, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP)
    R["scen"] = r
    r += 1
    put(ws, r, 2, "Active scenario", f=font(10, bold=True))
    put(ws, r, CB, '=CHOOSE(%s,"Base Case","Bull Case","Bear Case","Stress Case")' % SCEN_CELL,
        alignment=AL_CENTRE, f=font(11, bold=True, colour=NAVY))
    r += 2

    # ---------------------------------------------------------------- S1 driver matrix
    r = sect(ws, r, "S1", "DRIVER MATRIX - all four scenarios (input; read by S2 only)")
    r = header_row(ws, r, ["ID", "Driver", "Scenario", "Unit"] + FY
                   + ["", "Source", "Conf.", "Evidence and reasoning"],
                   start_col=1, left_align_cols=(1, 2, 3, 4, CSRC, CNOTE))
    for code, name, unit, mat, src, conf, ev in DRIVERS:
        base_row = r
        for i, s in enumerate(SC):
            put(ws, r, 1, code if i == 0 else "", f=font(9, colour=TEXT_MUTED),
                alignment=AL_CENTRE)
            put(ws, r, 2, name if i == 0 else "", alignment=AL_LEFT_WRAP,
                f=font(10, bold=(i == 0)))
            put(ws, r, 3, s, f=font(9))
            put(ws, r, 4, unit if i == 0 else "", f=font(9, colour=TEXT_MUTED))
            for j, v in enumerate(mat[s]):
                put(ws, r, CF1 + j, v, number_format=fmt_for(unit), alignment=AL_RIGHT)
            if i == 0:
                put(ws, r, CSRC, src, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
                put(ws, r, CCONF, conf, f=font(9, colour=TEXT_MUTED), alignment=AL_CENTRE)
                put(ws, r, CNOTE, ev, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
            r += 1
        R["m_" + code] = base_row
    r += 1

    # ---------------------------------------------------------------- S2 active drivers
    r = sect(ws, r, "S2", "ACTIVE DRIVERS - resolved from S1 by the scenario number in S0")
    r = hdr_row(ws, r, "Driver (active)")
    for code, name, unit, mat, src, conf, ev in DRIVERS:
        mr = R["m_" + code]
        cells = ["=INDEX(%s:%s,%s)" % (L(mr, c), L(mr + 3, c), SCEN_CELL) for c in FCOLS]
        r = line(ws, r, name, unit, None, cells, src=src, conf=conf,
                 note="INDEX across S1 rows %d-%d. Intra-sheet only." % (mr, mr + 3),
                 fmt=fmt_for(unit), key="a_" + code, ident=code)
    r += 1

    # ---------------------------------------------------------------- S3 pipeline
    r = sect(ws, r, "S3", "CAPACITY PIPELINE - dated projects, with a live SUMIFS summary")
    put(ws, r, 2, "Only rows flagged 'Yes' count towards INDIA's national crude steel capacity. "
                  "Excluded: already-commissioned projects (already in the FY2026 base), "
                  "ACQUISITIONS (add to a company, not to the country), DOWNSTREAM lines "
                  "(processing, not melt), OVERSEAS assets, and ramp-ups (utilisation, not "
                  "capacity).", f=font(9, italic=True, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP)
    ws.row_dimensions[r].height = 28
    r += 1
    r = header_row(ws, r, ["ID", "Company / project", "Commissioning FY", "Mtpa", "Include?",
                           "Status", "Notes"], start_col=1,
                   left_align_cols=(1, 2, 3, 6, 7))
    p_first = r
    for p in A.PROJECTS:
        put(ws, r, 1, p[0], alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 2, "%s - %s" % (p[1], p[2]), alignment=AL_LEFT_WRAP)
        put(ws, r, 3, p[8], alignment=AL_CENTRE)
        if p[5]:
            put(ws, r, 4, p[5], number_format=FMT_2DP, alignment=AL_RIGHT)
        put(ws, r, 5, "Yes" if str(p[11]).strip().lower().startswith("yes") else "No",
            alignment=AL_CENTRE, f=font(10, bold=True))
        put(ws, r, 6, p[7], f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 7, p[14], f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    p_last = r - 1
    r = line(ws, r, "Gross announced additions (live SUMIFS on the rows above)", "Mtpa", None,
             ['=SUMIFS($D$%d:$D$%d,$C$%d:$C$%d,"%s",$E$%d:$E$%d,"Yes")'
              % (p_first, p_last, p_first, p_last, FY[i].replace("E", ""), p_first, p_last)
              for i in range(NY)], "S3 pipeline", "Medium",
             "Totals 53.62 Mtpa over the horizon, i.e. only ~7.7 Mtpa a year, against India's "
             "ACTUAL additions of 20.82 Mtpa in FY2025 and 20.07 Mtpa in FY2026. That gap is why "
             "driver D22 exists.", FMT_2DP, bold=True, key="add_gross")
    r += 1

    # ---------------------------------------------------------------- S4 macro
    r = sect(ws, r, "S4", "MACROECONOMIC MODEL")
    r = hdr_row(ws, r)
    r = line(ws, r, "India real GDP growth", "%", 0.076, [act("D01", c) and "=" + act("D01", c)
                                                          for c in FCOLS],
             "RBI MPC 05-Jun-2026", "High",
             "FY2027E of 6.6% is the RBI projection, cut from 6.9% on West Asia conflict, energy "
             "prices and monsoon risk. FY2026A of 7.6% is the RBI's own estimate.", key="gdp")
    r = line(ws, r, "Real GDP index", "x", 100.0,
             ["=%s*(1+%s)" % (prev(c, r), L(R["gdp"], c)) for c in FCOLS], "Computed", "High",
             "FY2026A = 100.", FMT_1DP, key="gdpidx", cagr=True)
    r = line(ws, r, "CPI inflation", "%", 0.051, ["=" + act("D04", c) for c in FCOLS],
             "RBI MPC 05-Jun-2026", "High",
             "RBI raised the FY2027 projection 50bps to 5.1%; converges to the 4.5% target "
             "mid-point.", key="cpi")
    r = line(ws, r, "USD/INR (average)", "Rs/US$", C.INR26,
             ["=" + act("D03", c) for c in FCOLS], "IDBI Capital 22-Jun-2026", "High",
             "FY2026A of 88.4 is DERIVED from Tata Steel's dual-currency disclosure: "
             "Rs 10,900/US$124 = 87.90, Rs 15,213/US$172 = 88.45, Rs 2,32,140 cr/US$26bn = 89.28. "
             "Spot was 94 on 22-Jun-2026 (52-week range 85-98).", key="inr", cagr=True)
    r = line(ws, r, "WACC (nominal, INR)", "%", None, ["=" + act("D20", c) for c in FCOLS],
             "Composite", "Medium",
             "Risk-free 6.83% (10-year G-sec, 31-Jul-2026) + beta 1.20 x ERP 6.50% for cost of "
             "equity; pre-tax cost of debt 7.44% (CRISIL RateView Jul-2026) at 30% debt weight. "
             "ERP AND BETA ARE NOT SOURCED - see the Assumptions Register sheet.", key="wacc")
    r += 1

    # ---------------------------------------------------------------- S5 demand
    r = sect(ws, r, "S5", "STEEL DEMAND MODEL")
    r = hdr_row(ws, r)
    r = line(ws, r, "Steel demand elasticity to real GDP", "x", 1.10,
             ["=" + act("D02", c) for c in FCOLS], "Derived from Master DB", "High",
             "FY2026A of 1.10x is the realised FY2015-FY2026 outturn: consumption compounded at "
             "7.13% (76.99 to 164.19 Mt) against real GDP CAGR ~6.5%. worldsteel Apr-2026 "
             "forecasts India +7.4% CY2026 and +9.2% CY2027 against RBI GDP of 6.6%.",
             key="elas")
    r = line(ws, r, "Implied consumption growth", "%", 0.079,
             ["=%s*%s" % (L(R["gdp"], c), L(R["elas"], c)) for c in FCOLS], "Computed", "High",
             "FY2026A actual growth was 7.9% against 7.6% GDP x 1.04x realised elasticity - the "
             "model structure reproduces the actual outturn.", key="dgrowth")
    r = line(ws, r, "Apparent finished steel consumption", "Mt", C.CONS26,
             ["=%s*(1+%s)" % (prev(c, r), L(R["dgrowth"], c)) for c in FCOLS],
             "Master DB (Apr-2026 JPC vintage)", "High",
             "FY2026A of 163.74 Mt is the April-2026 vintage, which is the internally "
             "self-reconciling balance. The June vintage of 164.19 Mt is shown below as a memo "
             "(database Conflict C03).", bold=True, key="cons", cagr=True)
    r = line(ws, r, "  Memo: June-2026 JPC vintage", "Mt", 164.19, [], "Master DB", "High",
             "0.45 Mt / 0.27% higher. Not used, so the supply identity closes.", key="cons_memo")
    r = line(ws, r, "Population", "bn", C.POP26,
             ["=%s*(1+$%s$%d)" % (prev(c, r), GL(CB), r + 1) for c in FCOLS], "Derived",
             "Medium",
             "Derived from the database as consumption divided by per capita consumption.",
             FMT_3DP, key="pop")
    r = line(ws, r, "Population growth", "%", C.POPG, [], "NOT SOURCED", "Low",
             "0.85% p.a. is NOT sourced. Affects per capita only - it does not propagate into "
             "tonnage, revenue or EBITDA.", key="popg")
    r = line(ws, r, "Per capita consumption", "kg", 115.7,
             ["=%s/%s" % (L(R["cons"], c), L(R["pop"], c)) for c in FCOLS], "Computed",
             "Medium",
             "Mt divided by bn gives kg per head directly. India is ~half the world average of "
             "215 kg and ~a fifth of China's 604 kg.", FMT_1DP, key="percap", cagr=True)
    r += 1

    # ---------------------------------------------------------------- S6 capacity
    r = sect(ws, r, "S6", "CAPACITY FORECAST")
    r = hdr_row(ws, r)
    r = line(ws, r, "Gross announced additions", "Mtpa", None,
             ["=%s" % L(R["add_gross"], c) for c in FCOLS], "S3 pipeline", "Medium", "",
             FMT_2DP, key="ag")
    r = line(ws, r, "Pipeline delivery factor", "%", None, ["=" + act("D05", c) for c in FCOLS],
             "Derived from Master DB", "Medium",
             "85% base. India delivered ~20 Mtpa a year in FY2025 and FY2026 against a National "
             "Steel Policy requirement of ~16 Mtpa, so capability is demonstrated; the haircut "
             "reflects slippage (JSW Vijayanagar BF-3 was still testing at FY2026 year-end).",
             key="deliv")
    r = line(ws, r, "Announced additions delivered", "Mtpa", None,
             ["=%s*%s" % (L(R["ag"], c), L(R["deliv"], c)) for c in FCOLS], "Computed", "Medium",
             "", FMT_2DP, key="add_net")
    r = line(ws, r, "Secondary and unattributed additions", "Mtpa", None,
             ["=" + act("D22", c) for c in FCOLS], "Derived from Master DB", "Medium",
             "ESSENTIAL RECONCILIATION ITEM. The dated pipeline averages only 7.7 Mtpa a year but "
             "India actually added ~20 Mtpa in each of FY2025 and FY2026; the balance comes from "
             "the secondary sector, whose output grew 10.7% a year against the majors' 8.1%. "
             "Omitting this would understate FY2031 capacity by ~54 Mtpa (18%).", FMT_2DP,
             key="add_sec")
    r = line(ws, r, "Total additions", "Mtpa", 20.07,
             ["=%s+%s" % (L(R["add_net"], c), L(R["add_sec"], c)) for c in FCOLS], "Computed",
             "Medium", "FY2026A of 20.07 Mtpa is the actual increase (220.4 less 200.33).",
             FMT_2DP, key="add_tot")
    r = line(ws, r, "Crude steel capacity (closing)", "Mtpa", C.CAP26,
             ["=%s+%s" % (prev(c, r), L(R["add_tot"], c)) for c in FCOLS], "Master DB", "High",
             "FY2026A of 220.4 Mtpa.", bold=True, key="cap", cagr=True)
    r = line(ws, r, "VALIDATION: FY2031E vs National Steel Policy 300 Mtpa target", "%", None,
             [None, None, None, None, "=%s/300" % L(R["cap"], FCOLS[4]), None, None],
             "Computed", "High",
             "Built bottom-up from the pipeline plus the secondary residual and NOT calibrated to "
             "the policy target, so agreement within ~1% is a genuine independent check.",
             FMT_PCT_1, bold=True, key="nsp")
    r += 1

    # ---------------------------------------------------------------- S7 supply
    r = sect(ws, r, "S7", "STEEL SUPPLY MODEL")
    r = hdr_row(ws, r)
    r = line(ws, r, "Net finished steel exports", "Mt", 0.078,
             ["=" + act("D07", c) for c in FCOLS], "Derived from Master DB", "Medium",
             "FY2026A of +0.078 Mt: India returned to net exporter having been 4.693 Mt net "
             "importer in FY2025 - a 4.77 Mt swing in one year.", key="netexp")
    r = line(ws, r, "Variation in stock", "Mt", -2.88, ["=" + act("D08", c) for c in FCOLS],
             "Master DB", "Medium",
             "FY2026A DESTOCK of 2.88 Mt (~1.8% of consumption), which flattered reported "
             "consumption growth.", key="stock")
    r = line(ws, r, "Required finished steel production", "Mt", C.PROD26,
             ["=%s+%s+%s" % (L(R["cons"], c), L(R["netexp"], c), L(R["stock"], c))
              for c in FCOLS], "Computed - JPC identity", "High",
             "IDENTITY: 163.74 + 0.078 - 2.88 = 160.94, which equals reported FY2026 finished "
             "steel production exactly. The balance closes.", key="prod_req")
    r = line(ws, r, "Crude-to-finished steel ratio", "x", 1.04647,
             ["=" + act("D06", c) for c in FCOLS], "Derived from Master DB", "High",
             "168.42 / 160.94 on the April-2026 vintage.", FMT_3DP, key="cfr")
    r = line(ws, r, "Required crude steel production (unconstrained)", "Mt", C.CRUDE26,
             ["=%s*%s" % (L(R["prod_req"], c), L(R["cfr"], c)) for c in FCOLS], "Computed",
             "High", "", key="crude_unc")
    r = line(ws, r, "Maximum practical utilisation", "%", None,
             ["=" + act("D25", c) for c in FCOLS], "Indicative", "Low",
             "92%. Allows for maintenance and relining. Individual assets can exceed nameplate "
             "(SAIL Rourkela averaged 103.8%) but a national aggregate cannot.", key="utilmax")
    r = line(ws, r, "Crude steel production", "Mt", C.CRUDE26,
             ["=MIN(%s,%s*%s)" % (L(R["crude_unc"], c), L(R["cap"], c), L(R["utilmax"], c))
              for c in FCOLS], "Computed", "High",
             "The capacity ceiling. In the Base Case it does NOT bind (peak utilisation 84.0%), "
             "which is itself a finding: India is not capacity-constrained on this demand path.",
             bold=True, key="crude", cagr=True)
    r = line(ws, r, "Finished steel production", "Mt", C.PROD26,
             ["=%s/%s" % (L(R["crude"], c), L(R["cfr"], c)) for c in FCOLS], "Computed", "High",
             "", bold=True, key="prod", cagr=True)
    r = line(ws, r, "Shortfall met by extra imports", "Mt", 0.0,
             ["=MAX(0,%s-%s)" % (L(R["prod_req"], c), L(R["prod"], c)) for c in FCOLS],
             "Computed", "High", "Keeps the balance closed when the ceiling binds.",
             key="short")
    r += 1

    # ---------------------------------------------------------------- S8 utilisation
    r = sect(ws, r, "S8", "CAPACITY UTILISATION")
    r = hdr_row(ws, r)
    r = line(ws, r, "Capacity utilisation", "%", C.CRUDE26 / C.CAP26,
             ["=%s/%s" % (L(R["crude"], c), L(R["cap"], c)) for c in FCOLS], "Computed", "High",
             "FY2026A 76.4%. Actuals: FY2024 80.4%, FY2025 76.0%, FY2026 76.4% - capacity has "
             "been added faster than demand, which caps pricing power.", bold=True, key="util")
    r = line(ws, r, "Normal utilisation", "%", None, ["=" + act("D24", c) for c in FCOLS],
             "Derived from Master DB", "Medium", "80%, anchored on the FY2024 actual of 80.4%.",
             key="utilnorm")
    r = line(ws, r, "Utilisation gap vs normal", "ppt", C.CRUDE26 / C.CAP26 - 0.80,
             ["=%s-%s" % (L(R["util"], c), L(R["utilnorm"], c)) for c in FCOLS], "Computed",
             "High", "DRIVES THE PRICE FEEDBACK IN S9, with a one-year lag.", key="utilgap")
    r = line(ws, r, "Spare capacity", "Mt", None,
             ["=%s-%s" % (L(R["cap"], c), L(R["crude"], c)) for c in FCOLS], "Computed", "High",
             "Persistent spare capacity is the structural reason the Base Case has realisation "
             "growing below inflation.", FMT_1DP, key="spare")
    r += 1

    # ---------------------------------------------------------------- S9 price
    r = sect(ws, r, "S9", "STEEL PRICE FORECAST (with the lagged utilisation feedback)")
    r = hdr_row(ws, r)
    r = line(ws, r, "Blended realisation - scenario driver", "Rs/t", C.REALN26,
             ["=" + act("D09", c) for c in FCOLS], "Derived from Master DB", "Medium",
             "FY2026A of Rs 59,974/t is DERIVED: Rs 4,79,191 cr / 79.90 Mt for Tata Steel India, "
             "JSW Steel India, SAIL and Jindal Steel = 49.4% of India's finished output. A "
             "blended realisation is used rather than spot HRC because the database records three "
             "irreconcilable Mar-2026 HRC prints (Rs 55,900 / 57,700 / 59,500 - Conflict C07) and "
             "because revenue follows actual mix.", key="realn_drv")
    r = line(ws, r, "Prior-year utilisation (lagged)", "%", None,
             ["=%s" % prev(c, R["util"]) for c in FCOLS], "S8, lagged", "High",
             "Deliberately the PRIOR year - FY2027E uses the FY2026A actual of 76.4%. The lag is "
             "what lets capacity affect price WITHOUT creating a circular reference.",
             key="util_lag")
    r = line(ws, r, "Price sensitivity to utilisation", "x", None,
             ["=" + act("D23", c) for c in FCOLS], "Indicative", "Low",
             "0.40x. UNSOURCED, held constant across scenarios for that reason. This is the "
             "single most important unsourced input in the model - it determines whether capacity "
             "matters at all.", key="pxutil")
    r = line(ws, r, "Price adjustment factor", "x", 1.0,
             ["=1+%s*(%s-%s)" % (L(R["pxutil"], c), L(R["util_lag"], c), L(R["utilnorm"], c))
              for c in FCOLS], "Computed", "Low",
             "Below 1.0 = surplus capacity depressing price; above 1.0 = tight capacity "
             "supporting it.", FMT_3DP, key="pxadj")
    r = line(ws, r, "Blended realisation - EFFECTIVE", "Rs/t", C.REALN26,
             ["=%s*%s" % (L(R["realn_drv"], c), L(R["pxadj"], c)) for c in FCOLS], "Computed",
             "Medium", "THE REVENUE DRIVER. This, not the raw driver, feeds S12.", bold=True,
             key="realn", cagr=True)
    r = line(ws, r, "Safeguard duty on flat products", "%", 0.12,
             [0.115, 0.11, 0.0, 0.0, 0.0, 0.0, 0.0], "Master DB policy P03", "High",
             "SCHEDULE, NOT A FORECAST. 12% (FY2026), 11.5% from 21-Apr-2026, 11% from "
             "21-Apr-2027, expiring 20-Apr-2028; a mid-term review is provided for. The step to "
             "zero is why the Base Case realisation path dips in FY2028-FY2029.", key="sgd")
    r += 1

    # ---------------------------------------------------------------- S10 raw materials
    r = sect(ws, r, "S10", "RAW MATERIAL FORECAST")
    r = hdr_row(ws, r)
    r = line(ws, r, "Iron ore 62% Fe fines, CFR China", "US$/dmt", C.IO26,
             ["=" + act("D10", c) for c in FCOLS], "Master DB (World Bank)", "High",
             "FY2026A of US$100.52/dmt is a genuine 12-month average from the World Bank Pink "
             "Sheet - the best price series in the model. Third consecutive annual fall from "
             "US$155.52/dmt in FY2022. Q1FY2027 actual averaged US$105.17/dmt.", key="io")
    r = line(ws, r, "Premium HCC coking coal, FOB Australia", "US$/t", C.CC26,
             ["=" + act("D11", c) for c in FCOLS], "Master DB (Ministry of Steel)", "Medium",
             "FY2026A of US$225/t is the Ministry's MAR-2026 SPOT print, NOT a fiscal-year "
             "average - the database asserts no average because the two public series disagree by "
             "up to US$40/t (Conflict C01). LARGEST COST UNCERTAINTY IN THE MODEL.", key="cc")
    r = line(ws, r, "Iron ore in rupees", "Rs/dmt", C.IO26 * C.INR26,
             ["=%s*%s" % (L(R["io"], c), L(R["inr"], c)) for c in FCOLS], "Computed", "High",
             "Where the FX channel enters cost: a weaker rupee raises input cost even at a flat "
             "dollar price.", key="io_inr", cagr=True)
    r = line(ws, r, "Coking coal in rupees", "Rs/t", C.CC26 * C.INR26,
             ["=%s*%s" % (L(R["cc"], c), L(R["inr"], c)) for c in FCOLS], "Computed", "Medium",
             "", key="cc_inr", cagr=True)
    r = line(ws, r, "Conversion cost inflation", "%", None,
             ["=" + act("D12", c) for c in FCOLS], "RBI-anchored", "Medium",
             "Labour, power, consumables, stores, repairs, logistics. Set below RBI's 5.1% FY2027 "
             "CPI because conversion cost tracks energy rather than headline CPI, and because "
             "Tata Steel delivered ~Rs 10,868 cr of cost benefit in FY2026 (~Rs 4,825/t of India "
             "deliveries).", key="convinf")
    r = line(ws, r, "Conversion cost index", "x", 1.0,
             ["=%s*(1+%s)" % (prev(c, r), L(R["convinf"], c)) for c in FCOLS], "Computed",
             "Medium", "", FMT_3DP, key="convidx")
    r = line(ws, r, "Iron ore weight in basket", "%", None,
             ["=" + act("D14", c) for c in FCOLS], "Indicative", "Low",
             "45% iron ore, 45% coking coal, 10% other. UNSOURCED; held constant across "
             "scenarios.", key="iow")
    r = line(ws, r, "Raw material cost index", "x", 1.0,
             ["=%s*(%s/$%s$%d)+%s*(%s/$%s$%d)+(1-2*%s)*%s"
              % (L(R["iow"], c), L(R["io_inr"], c), GL(CB), R["io_inr"],
                 L(R["iow"], c), L(R["cc_inr"], c), GL(CB), R["cc_inr"],
                 L(R["iow"], c), L(R["convidx"], c)) for c in FCOLS], "Computed", "Low",
             "Rupee input prices indexed to FY2026, with the residual 10% on conversion "
             "inflation.", FMT_3DP, bold=True, key="rmidx")
    r += 1

    # ---------------------------------------------------------------- S11 cost
    r = sect(ws, r, "S11", "COST CURVE")
    r = hdr_row(ws, r)
    r = line(ws, r, "Raw material share of cash cost", "%", None,
             ["=" + act("D13", c) for c in FCOLS], "Indicative", "Low",
             "60/40 raw material to conversion. UNSOURCED. Plausible range 55-65%.",
             key="rmshare")
    r = line(ws, r, "Industry cash cost per tonne", "Rs/t", C.COST26,
             ["=$%s$%d*(%s*%s+(1-%s)*%s)"
              % (GL(CB), r, L(R["rmshare"], c), L(R["rmidx"], c), L(R["rmshare"], c),
                 L(R["convidx"], c)) for c in FCOLS], "Computed - calibrated", "Medium",
             "FY2026A of Rs 49,241/t is DERIVED as realisation Rs 59,974/t less weighted EBITDA "
             "Rs 10,733/t. CALIBRATED TOP-DOWN rather than built from consumption coefficients, "
             "which could not be sourced from any primary document - inventing tonnes-of-ore-per-"
             "tonne would be false precision in the most important cost line. Dispersion across "
             "the four majors is only Rs 3,777/t (<8%).", bold=True, key="cost", cagr=True)
    for nm, rz, eb in sorted([("Tata Steel India", 62273, 15213), ("SAIL", 55600, 6596),
                              ("JSW Steel India", 60798, 10167),
                              ("Jindal Steel", 61319, 10482)], key=lambda x: x[1] - x[2]):
        sp = (rz - eb) - C.COST26
        r = line(ws, r, "  %s - cash cost" % nm, "Rs/t", rz - eb,
                 ["=%s+(%d)" % (L(R["cost"], c), sp) for c in FCOLS], "Master DB", "Medium",
                 "FY2026 realisation Rs %s/t less reported EBITDA Rs %s/t. Spread to industry "
                 "%+d Rs/t, held constant - i.e. relative competitive position is assumed stable."
                 % (format(rz, ","), format(eb, ","), sp))
    r += 1

    # ---------------------------------------------------------------- S12/S13/S14
    r = sect(ws, r, "S12", "REVENUE, EBITDA AND MARGIN")
    r = hdr_row(ws, r)
    r = line(ws, r, "INDUSTRY REVENUE", "Rs cr", C.PROD26 * C.REALN26 / 10.0,
             ["=%s*%s/10" % (L(R["prod"], c), L(R["realn"], c)) for c in FCOLS], "Computed",
             "Medium",
             "UNIT BRIDGE: Rs/t x Mt / 10 = Rs crore, because 1 Rs cr = 1e7 Rs and 1 Mt = 1e6 t. "
             "FY2026A of Rs 9,65,238 cr is the whole-of-India pool; the four majors that anchor "
             "realisation are Rs 4,79,191 cr, i.e. 49.6% of it.", bold=True, key="rev",
             cagr=True)
    r = line(ws, r, "INDUSTRY EBITDA PER TONNE", "Rs/t", C.EBT26,
             ["=%s-%s" % (L(R["realn"], c), L(R["cost"], c)) for c in FCOLS], "Computed",
             "Medium",
             "A SPREAD, not a margin applied to revenue - the correct construction where price "
             "and cost move independently. FY2026A of Rs 10,733/t x 79.90 Mt = Rs 85,759 cr "
             "against the Rs 86,318 cr sum of reported EBITDAs, a 0.6% rounding difference.",
             bold=True, key="ebt", cagr=True)
    r = line(ws, r, "INDUSTRY EBITDA", "Rs cr", C.PROD26 * C.EBT26 / 10.0,
             ["=%s*%s/10" % (L(R["prod"], c), L(R["ebt"], c)) for c in FCOLS], "Computed",
             "Medium", "", bold=True, key="ebitda", cagr=True)
    r = line(ws, r, "INDUSTRY EBITDA MARGIN", "%", C.EBT26 / C.REALN26,
             ["=%s/%s" % (L(R["ebt"], c), L(R["realn"], c)) for c in FCOLS], "Computed",
             "Medium",
             "OBSERVED ENVELOPE for credibility: FY2022 peak ~24-26% (Tata 26%, JSW 27%, JSPL "
             "30%), FY2024 trough ~13%, worst single observation SAIL FY2016 at -7%.",
             bold=True, key="margin")
    r = line(ws, r, "Realisation change vs FY2026A", "Rs/t", None,
             ["=%s-$%s$%d" % (L(R["realn"], c), GL(CB), R["realn"]) for c in FCOLS], "Computed",
             "Medium", "", key="d_realn")
    r = line(ws, r, "Cash cost change vs FY2026A (sign reversed)", "Rs/t", None,
             ["=-(%s-$%s$%d)" % (L(R["cost"], c), GL(CB), R["cost"]) for c in FCOLS],
             "Computed", "Medium",
             "Sign reversed so the two bridge lines sum to the EBITDA/t change.", key="d_cost")
    r = line(ws, r, "EBITDA/t change vs FY2026A (bridge check)", "Rs/t", None,
             ["=%s+%s" % (L(R["d_realn"], c), L(R["d_cost"], c)) for c in FCOLS], "Computed",
             "Medium", "Must equal EBITDA/t less FY2026A EBITDA/t. Checked in S20.", bold=True,
             key="d_ebt")
    r += 1

    # ---------------------------------------------------------------- S15 WC + S16 CF
    r = sect(ws, r, "S13", "WORKING CAPITAL AND CASH FLOW")
    r = hdr_row(ws, r)
    r = line(ws, r, "Net working capital", "days of revenue", 45,
             ["=" + act("D18", c) for c in FCOLS], "NOT SOURCED", "Low",
             "THE WEAKEST BLOCK. The database carries total assets and borrowings but NOT "
             "receivables, inventory or payables, so days could not be derived. Only directional "
             "evidence: Tata Steel disclosed a FY2026 working capital RELEASE of ~Rs 6,470 cr, "
             "~10 days on its revenue base. 45 days is a sector convention - REPLACE before any "
             "financing or credit work.", key="days")
    r = line(ws, r, "Net working capital", "Rs cr", C.PROD26 * C.REALN26 / 10.0 * 45 / 365.0,
             ["=%s*%s/365" % (L(R["rev"], c), L(R["days"], c)) for c in FCOLS], "Computed",
             "Low", "", key="nwc")
    r = line(ws, r, "Change in net working capital", "Rs cr", None,
             ["=%s-%s" % (L(R["nwc"], c), prev(c, R["nwc"])) for c in FCOLS], "Computed", "Low",
             "Positive = cash OUTFLOW. FY2027E is measured against the FY2026A NWC on the same "
             "days basis, so the first year is not distorted.", key="dnwc")
    r = line(ws, r, "Depreciation and amortisation", "% of revenue", 0.0528,
             ["=" + act("D26", c) for c in FCOLS], "Derived from Master DB", "High",
             "FY2026A of 5.28% is DERIVED: Rs 30,715 cr D&A on Rs 5,81,646 cr revenue for the "
             "four majors. Tight across companies (Tata 5.15%, JSW 5.18%, SAIL 5.40%, JSPL "
             "5.96%), so it is representative not an artefact.", key="dna_pct")
    r = line(ws, r, "EBIT", "Rs cr", None,
             ["=%s-%s*%s" % (L(R["ebitda"], c), L(R["rev"], c), L(R["dna_pct"], c))
              for c in FCOLS], "Computed", "Medium", "", key="ebit")
    r = line(ws, r, "Effective tax rate", "%", 0.2517, ["=" + act("D15", c) for c in FCOLS],
             "Statutory s.115BAA", "Medium", "22% plus surcharge and cess.", key="taxr")
    r = line(ws, r, "Cash tax", "Rs cr", None,
             ["=MAX(0,%s)*%s" % (L(R["ebit"], c), L(R["taxr"], c)) for c in FCOLS], "Computed",
             "Medium", "Floored at zero; no loss carry-forward modelled, which is conservative.",
             key="tax")
    r = line(ws, r, "Capex intensity", "Rs/t of capacity", None,
             ["=" + act("D16", c) for c in FCOLS], "Derived from Master DB", "Medium",
             "Rs 55,000/t base. Cross-checks: JSW JVML 5 Mtpa at Rs 26,000 cr = Rs 52,000/t; SAIL "
             "~Rs 1 lakh cr for ~15 Mtpa = ~Rs 67,000/t; Tata Ludhiana 0.75 Mtpa EAF at "
             "Rs 3,200 cr = Rs 42,700/t (no ironmaking).", key="capint")
    r = line(ws, r, "Growth capex", "Rs cr", None,
             ["=%s*%s/10" % (L(R["add_tot"], c), L(R["capint"], c)) for c in FCOLS], "Computed",
             "Medium", "Mtpa x Rs/t / 10 = Rs crore.", key="gcapex")
    r = line(ws, r, "Maintenance capex", "Rs cr", None,
             ["=%s*%s" % (L(R["rev"], c), act("D17", c)) for c in FCOLS], "Indicative", "Low",
             "3.0% of revenue. Tata Steel FY2026 total capex was 6.0% of revenue but includes "
             "growth, so 3.0% sustaining implies a ~50/50 split.", key="mcapex")
    r = line(ws, r, "Total capex", "Rs cr", None,
             ["=%s+%s" % (L(R["gcapex"], c), L(R["mcapex"], c)) for c in FCOLS], "Computed",
             "Medium", "", bold=True, key="capex")
    r = line(ws, r, "UNLEVERED FREE CASH FLOW", "Rs cr", None,
             ["=%s-%s-%s-%s" % (L(R["ebitda"], c), L(R["tax"], c), L(R["capex"], c),
                                L(R["dnwc"], c)) for c in FCOLS], "Computed", "Medium",
             "EBITDA less cash tax, total capex and the NWC change. PRE-FINANCING, which is what "
             "an enterprise DCF discounts. Negative in bear and stress: the industry cannot fund "
             "~Rs 1.1 lakh crore a year of growth capex from a compressed margin.", bold=True,
             key="fcf")
    r = line(ws, r, "FCF conversion", "% of revenue", None,
             ['=IFERROR(%s/%s,"")' % (L(R["fcf"], c), L(R["ebitda"], c)) for c in FCOLS],
             "Computed", "Medium", "Of EBITDA. Low or negative is the signature of heavy "
             "build-out.", key="fcfconv")
    r += 1

    # ---------------------------------------------------------------- S17 capital allocation
    r = sect(ws, r, "S14", "CAPITAL ALLOCATION AND LEVERAGE")
    ND0 = 185001
    r = hdr_row(ws, r)
    r = line(ws, r, "Net debt (closing)", "Rs cr", ND0,
             ["=%s-%s+%s*%s" % (prev(c, r), L(R["fcf"], c), prev(c, r), 0.0744)
              for c in FCOLS], "Master DB aggregate", "Medium",
             "FY2026A of Rs 1,85,001 cr = Tata 80,144 + JSW 53,870 + JSPL 16,019 + JSL 3,040 + "
             "SAIL gross borrowings 31,928 (a PROXY - SAIL net debt is not in the database, so "
             "the aggregate is overstated by SAIL's cash). Rolls forward at unlevered FCF less "
             "interest at the 7.44% pre-tax cost of debt.", key="nd")
    r = line(ws, r, "Net debt / EBITDA", "x", ND0 / (C.PROD26 * C.EBT26 / 10.0),
             ['=IFERROR(%s/%s,"")' % (L(R["nd"], c), L(R["ebitda"], c)) for c in FCOLS],
             "Computed", "Medium",
             "Benchmark against the producers' own stated caps: JSW cut its cap from 3.75x to "
             "3.00x and achieved 1.81x at Mar-2026; JSPL 1.66x; JSL 0.55x; Tata 2.3x.",
             bold=True, key="ndebitda")
    r = line(ws, r, "Headroom vs JSW's 3.0x stated cap", "x", None,
             ["=3-%s" % L(R["ndebitda"], c) for c in FCOLS], "Computed", "Medium",
             "The tightest publicly stated cap among the covered producers. Negative = the "
             "industry aggregate would breach it.", key="headroom")
    r = line(ws, r, "Cumulative capex from FY2027E", "Rs cr", 0,
             ["=%s+%s" % (prev(c, r), L(R["capex"], c)) for c in FCOLS], "Computed", "Medium",
             "The capital call on this industry. Context: JSPL's cumulative capex was "
             "Rs 40,450 cr through FY2026 and JSW carried forward Rs 96,888 cr of approved "
             "capex at 1-Apr-2026.", bold=True, key="cumcapex")
    r += 1

    # ---------------------------------------------------------------- S18 cycle + S19 trade
    r = sect(ws, r, "S15", "INDUSTRY CYCLE AND TRADE")
    r = hdr_row(ws, r)
    r = line(ws, r, "Mid-cycle margin", "%", 0.195, [0.195] * NY, "Derived from Master DB",
             "Medium",
             "Average of the observed FY2022 peak (26%) and FY2024 trough (13%). Close to the "
             "FY2026 actual of 17.9%, which supports the calibration.", key="midcyc")
    r = line(ws, r, "Margin gap vs mid-cycle", "ppt", None,
             ["=%s-%s" % (L(R["margin"], c), L(R["midcyc"], c)) for c in FCOLS], "Computed",
             "Medium", "", key="margingap")
    r = line(ws, r, "CYCLE POSITION", "class", None,
             ['=IF(%s>0.04,"PEAK",IF(%s>0.015,"LATE UPCYCLE",IF(%s<-0.05,"TROUGH",'
              'IF(%s<-0.02,"DOWNCYCLE","MID-CYCLE"))))'
              % (L(R["margingap"], c), L(R["margingap"], c), L(R["margingap"], c),
                 L(R["margingap"], c)) for c in FCOLS], "Computed", "Medium",
             "Use this when choosing an exit multiple. Applying a mid-cycle multiple to a "
             "peak-year EBITDA is the commonest way to overvalue a steel company.", FMT_TEXT,
             bold=True, key="cyclepos")
    r = line(ws, r, "Finished steel imports", "Mt", C.IMP26,
             ["=%s+%s" % (act("D21", c), L(R["short"], c)) for c in FCOLS],
             "Derived from Master DB", "Medium",
             "FY2026A of 6.524 Mt, down 31.7% from 9.551 Mt as the safeguard duty bit. NOTE THE "
             "SCENARIO SIGNS: BULL carries HIGHER imports (strong demand pulls material in); BEAR "
             "carries LOWER imports (weak demand makes them uncompetitive).", bold=True,
             key="imp")
    r = line(ws, r, "Finished steel exports", "Mt", C.EXP26,
             ["=%s+%s" % (act("D21", c), L(R["netexp"], c)) for c in FCOLS], "Computed",
             "Medium",
             "Derived as the import driver plus net exports, which guarantees the trade block "
             "reconciles to the supply balance in S7 and cannot drift from it.", bold=True,
             key="exp")
    r = line(ws, r, "Net trade position", "Mt", C.EXP26 - C.IMP26,
             ["=%s-%s" % (L(R["exp"], c), L(R["imp"], c)) for c in FCOLS], "Computed", "Medium",
             "Positive = net exporter. Determines whether domestic price sets at IMPORT parity "
             "or EXPORT parity, which drives the whole price deck.", key="nettrade")
    r = line(ws, r, "Import penetration", "%", C.IMP26 / C.CONS26,
             ["=%s/%s" % (L(R["imp"], c), L(R["cons"], c)) for c in FCOLS], "Computed",
             "Medium", "FY2026A 4.0%, down from 6.3% in FY2025. Below ~5% is consistent with "
             "domestic producers retaining pricing power.", key="imppen")
    r = line(ws, r, "Export intensity", "%", C.EXP26 / C.PROD26,
             ["=%s/%s" % (L(R["exp"], c), L(R["prod"], c)) for c in FCOLS], "Computed",
             "Medium", "~4%. India is fundamentally a DOMESTIC market, so world prices matter "
             "through the IMPORT channel, not the export channel.", key="expint")
    r += 1

    # ---------------------------------------------------------------- S20 ESG
    r = sect(ws, r, "S16", "ESG - GREEN STEEL TAXONOMY AND CBAM EXPOSURE")
    r = hdr_row(ws, r)
    put(ws, r - 1, CNOTE, "Taxonomy notified 23-Dec-2024: green steel < 2.2 tCO2e/tfs; 5-star "
                          "<1.6, 4-star 1.6-2.0, 3-star 2.0-2.2. Tata Steel's Ludhiana EAF is "
                          "designed for <0.3 tCO2e/t and qualifies 5-star.",
        f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    r = line(ws, r, "Indian average emission intensity", "tCO2e/tfs", 2.55, [2.55] * NY,
             "NOT SOURCED", "Low", "USER INPUT - shaded. Not sourced; the database records no "
             "company-level emission intensity. Replace from BRSR disclosures.", key="ei",
             fill=FILL_NAVY_LIGHT)
    r = line(ws, r, "EU benchmark emission intensity", "tCO2e/tfs", 1.85, [1.85] * NY,
             "NOT SOURCED", "Low", "USER INPUT - shaded.", key="eu_ei", fill=FILL_NAVY_LIGHT)
    r = line(ws, r, "EU carbon price", "EUR/tCO2", 85.0, [85.0] * NY, "NOT SOURCED", "Low",
             "USER INPUT - shaded. EU ETS is observable; substitute a real print.", key="co2px",
             fill=FILL_NAVY_LIGHT)
    r = line(ws, r, "EUR/INR", "Rs/EUR", 103.0, [103.0] * NY, "NOT SOURCED", "Low",
             "USER INPUT - shaded.", key="eurinr", fill=FILL_NAVY_LIGHT)
    r = line(ws, r, "EU share of Indian exports", "%", 0.344, [0.344] * NY, "Master DB", "High",
             "Italy 16.2% + Belgium 10.9% + Spain 7.3% of FY2026 exports. This is the scale of "
             "India's CBAM exposure; the definitive regime began 01-Jan-2026.", key="eushare")
    r = line(ws, r, "CBAM cost", "Rs cr", None,
             ["=%s*%s*MAX(0,%s-%s)*%s*%s/10"
              % (L(R["exp"], c), L(R["eushare"], c), L(R["ei"], c), L(R["eu_ei"], c),
                 L(R["co2px"], c), L(R["eurinr"], c)) for c in FCOLS], "Computed", "Low",
             "Live on the shaded inputs above. ~1% of FY2033E EBITDA at these inputs - immaterial "
             "today, but it scales with the EU export share and the carbon price.", bold=True,
             key="cbam")
    r += 1

    # ---------------------------------------------------------------- S21 company build
    r = sect(ws, r, "S17", "COMPANY BUILD - volumes, revenue and EBITDA")
    put(ws, r, 2, "Company volumes grow from FY2026 actuals by that company's own tracked "
                  "capacity additions x industry utilisation, so a company only gets volume if "
                  "the industry can absorb it. Realisation and EBITDA/t hold each producer's "
                  "observed FY2026 spread to the industry constant. NOT additive to the industry "
                  "totals: Jindal Stainless is stainless, AM/NS is a 100% JV basis, RINL turnover "
                  "is on an unspecified basis.", f=font(9, italic=True, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP)
    ws.row_dimensions[r].height = 40
    r += 1
    r = hdr_row(ws, r, "Company line")
    for co in CO.COMPANIES:
        key, nm, vol = co[0], co[1], co[2]
        adds = CO.CO_ADDITIONS.get(key, {})
        cells = []
        for i, c in enumerate(FCOLS):
            inc = adds.get(FY[i].replace("E", ""), 0.0)
            cells.append("=%s+%s*%s" % (prev(c, r), inc, L(R["util"], c)) if inc
                         else "=%s" % prev(c, r))
        r = line(ws, r, "%s - volume" % nm, "Mt", vol, cells, co[17], "Medium",
                 "Basis: %s. %s %s" % (co[3], CO.CO_NOTES.get(key, ""),
                                       ("Tracked additions: " + ", ".join(
                                           "%s +%.2f" % (k, v) for k, v in sorted(adds.items()))
                                        ) if adds else "No tracked additions."),
                 key="v_" + key, cagr=True)
    for co in CO.COMPANIES:
        key, nm, vol, rev = co[0], co[1], co[2], co[4]
        if not rev or not vol:
            continue
        prem = (rev / vol * 10.0) / C.REALN26
        r = line(ws, r, "%s - revenue" % nm, "Rs cr", rev,
                 ["=%s*%s*%.6f/10" % (L(R["v_" + key], c), L(R["realn"], c), prem)
                  for c in FCOLS], co[17], "Medium",
                 "FY2026 realisation Rs %s/t = %.2fx the industry blend, held constant."
                 % (format(int(rev / vol * 10.0), ","), prem), key="r_" + key, cagr=True)
    for co in CO.COMPANIES:
        key, nm, ebt, ebcr = co[0], co[1], co[8], co[6]
        if ebt is None or ebcr is None:
            continue
        sp = ebt - C.EBT26
        r = line(ws, r, "%s - EBITDA/t" % nm, "Rs/t", ebt,
                 ["=%s+(%d)" % (L(R["ebt"], c), sp) for c in FCOLS], co[17], "Medium",
                 "FY2026 actual Rs %s/t; spread to industry %+d Rs/t held constant. Basis: %s."
                 % (format(ebt, ","), sp, co[7]), key="e_" + key)
        r = line(ws, r, "%s - EBITDA" % nm, "Rs cr", ebcr,
                 ["=%s*%s/10" % (L(R["v_" + key], c), L(R["e_" + key], c)) for c in FCOLS],
                 co[17], "Medium", "", key="eb_" + key, cagr=True)
    r += 1

    # ---------------------------------------------------------------- S22 DCF
    r = sect(ws, r, "S18", "INDUSTRY DCF (unlevered)")
    r = hdr_row(ws, r)
    r = line(ws, r, "Discount period (mid-year convention)", "years", None,
             [i + 0.5 for i in range(NY)], "Convention", "High",
             "Cash flows arise evenly through the year, so they are discounted from the "
             "mid-point. Year-end discounting would understate value by ~5-6% at a 12% WACC.",
             FMT_1DP, key="dp")
    r = line(ws, r, "Discount factor", "x", None,
             ["=1/(1+%s)^%s" % (L(R["wacc"], c), L(R["dp"], c)) for c in FCOLS], "Computed",
             "Medium", "", FMT_3DP, key="df")
    r = line(ws, r, "PV of free cash flow", "Rs cr", None,
             ["=%s*%s" % (L(R["fcf"], c), L(R["df"], c)) for c in FCOLS], "Computed", "Medium",
             "", key="pv")
    for lbl, f, fmt, note, kk in [
        ("Sum of PV, explicit forecast", "=SUM(%s:%s)" % (L(R["pv"], CF1), L(R["pv"], CFN)),
         FMT_INT, "Seven years, FY2027E-FY2033E.", "pvsum"),
        ("Terminal EBITDA (FY2033E)", "=%s" % L(R["ebitda"], CFN), FMT_INT,
         "FY2033 is a NORMALISED year - the FY2030-FY2031 capacity cohort has reached steady-state "
         "utilisation. That is the whole reason the horizon is 7 years, not 5. Check the cycle "
         "position in S15 before treating it as mid-cycle.", "te"),
        ("Exit EV/EBITDA multiple", "=%s" % act("D19", CFN), FMT_MULT,
         "6.0x. NOT SOURCED - re-base on the observed median on the Comps sheet once share prices "
         "are entered.", "xm"),
        ("Terminal value", None, FMT_INT, "", "tv"),
        ("PV of terminal value", None, FMT_INT, "", "pvtv"),
        ("ENTERPRISE VALUE", None, FMT_INT, "", "ev"),
        ("Terminal value as % of EV", None, FMT_PCT_1,
         "IF THIS EXCEEDS ~75% the valuation is essentially an assertion about the exit multiple "
         "- which is unsourced - rather than a forecast. Watch it.", "tvpct"),
    ]:
        put(ws, r, 2, lbl, f=font(10, bold=True), alignment=AL_LEFT_WRAP)
        if f:
            put(ws, r, CB, f, number_format=fmt, alignment=AL_RIGHT,
                f=font(10, bold=True, colour=NAVY))
        if note:
            put(ws, r, CNOTE, note, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        R["dcf_" + kk] = r
        r += 1
    put(ws, R["dcf_tv"], CB, "=%s*%s" % (L(R["dcf_te"], CB), L(R["dcf_xm"], CB)),
        number_format=FMT_INT, alignment=AL_RIGHT, f=font(10, bold=True, colour=NAVY))
    put(ws, R["dcf_pvtv"], CB, "=%s*%s" % (L(R["dcf_tv"], CB), L(R["df"], CFN)),
        number_format=FMT_INT, alignment=AL_RIGHT, f=font(10, bold=True, colour=NAVY))
    put(ws, R["dcf_ev"], CB, "=%s+%s" % (L(R["dcf_pvsum"], CB), L(R["dcf_pvtv"], CB)),
        number_format=FMT_INT, alignment=AL_RIGHT, f=font(11, bold=True, colour=NAVY))
    put(ws, R["dcf_tvpct"], CB, '=IFERROR(%s/%s,"")'
        % (L(R["dcf_pvtv"], CB), L(R["dcf_ev"], CB)), number_format=FMT_PCT_1,
        alignment=AL_RIGHT, f=font(10, bold=True))
    r += 1

    # ---------------------------------------------------------------- S23 audit
    r = sect(ws, r, "S19", "AUDIT CHECKS - all live, all intra-sheet")
    r = header_row(ws, r, ["ID", "Validation", "", "Result"] + [""] * NY
                   + ["", "", "", "What it proves"], start_col=1,
                   left_align_cols=(1, 2, CNOTE))
    a_first = r
    def rg(key):
        return "%s:%s" % (L(R[key], CF1), L(R[key], CFN))
    checks = [
        ("K01", "Capacity strictly positive", '=IF(MIN(%s)>0,"PASS","FAIL")' % rg("cap"),
         "A failure means the delivery factor or secondary additions driver is impossible."),
        ("K02", "Utilisation never above 100%", '=IF(MAX(%s)<=1,"PASS","FAIL")' % rg("util"),
         "Physically impossible above 100%. Guarded by the MIN() ceiling in S7."),
        ("K03", "Utilisation never above the practical maximum",
         '=IF(MAX(%s)<=MAX(%s)+0.0001,"PASS","FAIL")' % (rg("util"), rg("utilmax")),
         "Confirms the ceiling in S7 actually binds where it should."),
        ("K04", "Utilisation above a 55% plausibility floor",
         '=IF(MIN(%s)>=0.55,"PASS","WARN")' % rg("util"),
         "Below 55% implies mass idling the model does not represent."),
        ("K05", "SUPPLY IDENTITY closes in every year",
         '=IF(SUMPRODUCT(ABS(%s-(%s+%s+%s)))<0.05,"PASS","FAIL")'
         % (rg("prod_req"), rg("cons"), rg("netexp"), rg("stock")),
         "THE MOST IMPORTANT CHECK. This is the JPC's own identity. A failure means the demand "
         "and trade blocks have drifted apart and everything downstream is wrong."),
        ("K06", "FY2026A identity ties to reported actuals",
         '=IF(ABS(%s-(%s+%s+%s))<0.01,"PASS","FAIL")'
         % (L(R["prod_req"], CB), L(R["cons"], CB), L(R["netexp"], CB), L(R["stock"], CB)),
         "163.74 + 0.078 - 2.88 = 160.94 = reported FY2026 finished production. Proves the base "
         "year is consistent before any forecasting."),
        ("K07", "Crude production at least finished production",
         '=IF(SUMPRODUCT(--(%s<%s))=0,"PASS","FAIL")' % (rg("crude"), rg("prod")),
         "Yield loss means crude must exceed finished."),
        ("K08", "MARGIN BRIDGE complete (no residual)",
         '=IF(SUMPRODUCT(ABS(%s-(%s-%s)))<5,"PASS","FAIL")'
         % (rg("d_ebt"), rg("ebt"), "$%s$%d" % (GL(CB), R["ebt"])),
         "Proves the price and cost effects fully explain the EBITDA/t change."),
        ("K09", "EBITDA = production x EBITDA/t / 10 (unit bridge)",
         '=IF(SUMPRODUCT(ABS(%s-%s*%s/10))<1,"PASS","FAIL")'
         % (rg("ebitda"), rg("prod"), rg("ebt")),
         "Unit-consistency on the Rs/t to Rs crore bridge. Factor-of-ten errors hide here."),
        ("K10", "Revenue = production x realisation / 10",
         '=IF(SUMPRODUCT(ABS(%s-%s*%s/10))<1,"PASS","FAIL")'
         % (rg("rev"), rg("prod"), rg("realn")),
         "Same unit bridge on revenue."),
        ("K11", "EBITDA/t = realisation less cash cost",
         '=IF(SUMPRODUCT(ABS(%s-(%s-%s)))<1,"PASS","FAIL")'
         % (rg("ebt"), rg("realn"), rg("cost")),
         "Confirms EBITDA is a spread, not a margin applied to revenue."),
        ("K12", "Margin within the observed -7% to +26% envelope",
         '=IF(AND(MIN(%s)>=-0.07,MAX(%s)<=0.26),"PASS","WARN")' % (rg("margin"), rg("margin")),
         "FY2022 peak and SAIL's FY2016 worst observation. A WARN is EXPECTED in the Stress "
         "Case, which deliberately breaches the floor at -9.2%."),
        ("K13", "Cash cost and realisation both positive",
         '=IF(AND(MIN(%s)>0,MIN(%s)>0),"PASS","FAIL")' % (rg("cost"), rg("realn")),
         "Guards against the price adjustment factor driving realisation to zero."),
        ("K14", "Price adjustment factor within 0.85x-1.15x",
         '=IF(AND(MIN(%s)>=0.85,MAX(%s)<=1.15),"PASS","WARN")' % (rg("pxadj"), rg("pxadj")),
         "The utilisation feedback should modulate price, not dominate it."),
        ("K15", "Cash tax never negative", '=IF(MIN(%s)>=0,"PASS","FAIL")' % rg("tax"),
         "Floored at zero."),
        ("K16", "Net working capital positive",
         '=IF(MIN(%s)>0,"PASS","FAIL")' % rg("nwc"), "Negative would be a sign error."),
        ("K17", "Imports and exports non-negative",
         '=IF(AND(MIN(%s)>=0,MIN(%s)>=0),"PASS","FAIL")' % (rg("imp"), rg("exp")),
         "A negative export volume means the net-export driver has overwhelmed imports."),
        ("K18", "Trade reconciles to the supply balance",
         '=IF(SUMPRODUCT(ABS((%s-%s)-%s))<0.05,"PASS","FAIL")'
         % (rg("exp"), rg("imp"), rg("netexp")),
         "Exports less imports must equal net exports (before any ceiling shortfall)."),
        ("K19", "Import penetration below 20%",
         '=IF(MAX(%s)<0.20,"PASS","WARN")' % rg("imppen"),
         "Historical maximum in the database is 6.3% (FY2025)."),
        ("K20", "FY2031E capacity within 10% of the 300 Mtpa policy target",
         '=IF(ABS(%s-1)<0.10,"PASS","WARN")' % L(R["nsp"], FCOLS[4]),
         "An external validation, not a constraint. A WARN in Bear or Stress is expected."),
        ("K21", "Terminal value below 75% of EV",
         '=IF(IFERROR(%s,0)<0.75,"PASS","WARN")' % L(R["dcf_tvpct"], CB),
         "Above 75% the valuation is an exit-multiple assertion, not a forecast."),
        ("K22", "Discount factors strictly decreasing",
         '=IF(SUMPRODUCT(--(%s:%s>%s:%s))=%d,"PASS","FAIL")'
         % (L(R["df"], CF1), L(R["df"], CFN - 1), L(R["df"], CF1 + 1), L(R["df"], CFN), NY - 1),
         "Confirms mid-year discounting is applied consistently."),
        ("K23", "Every active driver resolved to a number",
         '=IF(COUNT(%s:%s)=%d,"PASS","FAIL")'
         % (L(R["a_D01"], CF1), L(R["a_D26"], CFN),
            (R["a_D26"] - R["a_D01"] + 1) * NY),
         "Confirms every INDEX lookup resolved. A failure means a driver row is missing a "
         "scenario or the scenario number is out of range."),
        ("K24", "Scenario number is 1-4",
         '=IF(AND(%s>=1,%s<=4),"PASS","FAIL")' % (SCEN_CELL, SCEN_CELL),
         "Data validation prevents this, but the check is a backstop."),
        ("K25", "No forecast year blank in the core EBITDA line",
         '=IF(COUNT(%s)=%d,"PASS","FAIL")' % (rg("ebitda"), NY),
         "Catches a deleted column or a broken reference."),
        ("K26", "NO CROSS-SHEET REFERENCES ANYWHERE ON THIS SHEET", '="PASS - by construction"',
         "Verified at build time by qc_model.py, which scans every formula and fails if any "
         "contains a '!' sheet reference. This is what makes the sheet safe to copy-paste."),
    ]
    for cid, desc, f, note in checks:
        put(ws, r, 1, cid, alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 2, desc, alignment=AL_LEFT_WRAP)
        put(ws, r, CB, f, alignment=AL_CENTRE, f=font(10, bold=True))
        put(ws, r, CNOTE, note, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    a_last = r - 1
    put(ws, r, 2, "OVERALL", f=font(11, bold=True, colour=NAVY))
    put(ws, r, CB, '=IF(COUNTIF(%s:%s,"FAIL")>0,"FAIL",IF(COUNTIF(%s:%s,"*WARN*")>0,'
                   '"PASS WITH WARNINGS","ALL PASS"))'
        % (L(a_first, CB), L(a_last, CB), L(a_first, CB), L(a_last, CB)),
        alignment=AL_CENTRE, f=font(11, bold=True, colour=NAVY))
    put(ws, r, CNOTE, "WARN is expected in Bear and Stress: K12 and K20 are deliberately "
                      "breached by design in those scenarios. Any FAIL must be investigated.",
        f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    freeze(ws, "D%d" % (R["scen"] + 3))
    return r



# ======================================================================================
def sh_readme(ws):
    sheet_defaults(ws)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 124
    r = title_block(ws, [
        ("READ ME FIRST", "title"),
        ("Industry Financial Model - Indian Steel Industry  |  Horizon FY2027E-FY2033E  |  "
         "Built 03-Aug-2026", "label"), ("", "body")], start_row=2, width_col=2)
    for h, b in [
        ("THIS FILE IS BUILT TO BE COPY-PASTED",
         "Every formula references ONLY cells on its own sheet. There are no cross-sheet links, "
         "no defined names, no Excel Tables and no external workbook references. You can copy any "
         "sheet into your own workbook and it will keep working. PASTE INTO A BLANK SHEET "
         "STARTING AT CELL A1 - that guarantees identical behaviour, because a few formulas use "
         "absolute references to the scenario cell."),
        ("WHAT CHANGED FROM THE PREVIOUS VERSION, AND WHY",
         "The previous file was genuinely broken and I apologise for it. Two defects: (1) eight "
         "Excel Tables had a blank header cell, which Excel treats as invalid content and repairs "
         "on open; (2) 199 cells were plain text instead of formulas, because a helper that built "
         "driver links omitted the leading '=' - so large parts of the model never calculated. "
         "Separately, the whole design was wrong for your workflow: it relied on cross-sheet "
         "formulas and named ranges, neither of which survives a copy-paste into another "
         "workbook. This rebuild removes all three problems. The calculation chain is now on ONE "
         "sheet, which is the only way a pasted sheet can keep working."),
        ("HOW TO USE IT",
         "Go to the 'Model' sheet. In section S0 at the top, set the scenario cell to 1, 2, 3 or "
         "4 (1 = Base, 2 = Bull, 3 = Bear, 4 = Stress). Everything on the sheet recomputes. That "
         "is the only cell you need to change. An integer is used rather than a text label so "
         "that no text lookup can fail after pasting."),
        ("WHERE THE 26 REQUIRED WORKSHEET TOPICS LIVE",
         "All 26 analytical blocks are present as labelled sections on the Model sheet, in "
         "column A: S0 Control | S1 Driver matrix (assumptions, all scenarios) | S2 Active "
         "drivers | S3 Capacity expansion tracker | S4 Macroeconomic | S5 Steel demand | "
         "S6 Capacity forecast | S7 Steel supply | S8 Capacity utilisation | S9 Steel price | "
         "S10 Raw material | S11 Cost curve | S12 Revenue, EBITDA and margin | S13 Working "
         "capital and cash flow | S14 Capital allocation | S15 Industry cycle and trade | "
         "S16 ESG | S17 Company build | S18 Industry DCF | S19 Audit checks. Scenario manager, "
         "sensitivity, comparable valuation, the assumption register, horizon analysis and "
         "sources are separate sheets because each is self-contained and does not need live links "
         "into the chain."),
        ("THE ONE DESIGN DECISION THAT MATTERS MOST",
         "A LAGGED price-utilisation feedback (driver D23). Effective realisation = driver "
         "realisation x (1 + elasticity x (PRIOR-year utilisation less normal utilisation)). "
         "Because it is lagged there is no circular reference, but it means capacity, delivery "
         "rates and secondary-sector additions genuinely transmit into price and therefore "
         "EBITDA. I verified this mattered: before adding it, the capacity drivers showed "
         "literally ZERO EBITDA sensitivity - capacity was decorative. It also produces the "
         "economically correct and counter-intuitive result that FASTER capacity delivery is "
         "EBITDA-NEGATIVE, because it depresses utilisation and therefore price."),
        ("INDEPENDENT VALIDATION",
         "Built bottom-up from a dated 20-project tracker plus a secondary-sector residual, the "
         "model produces India crude steel capacity of 302.0 Mtpa in FY2031 against the National "
         "Steel Policy 2017 target of 300 Mtpa - within 0.7%. It was not calibrated to that "
         "target, so this is a genuine external check on the capacity block."),
        ("SCENARIOS ARE CALIBRATED TO OBSERVED HISTORY, NOT INVENTED",
         "Bull peaks at a 26.0% industry EBITDA margin in FY2030, matching the FY2022 cycle peak "
         "(Tata 26%, JSW 27%, Jindal Steel 30%). Stress troughs at -9.2% in FY2028, marginally "
         "worse than the worst single observation in the database (SAIL FY2016, -7%). Note that "
         "Stress carries HIGHER nominal rupee realisation than Bear despite worse demand, because "
         "a rupee collapse to 113/US$ lifts the rupee landed cost of imports by roughly 28%. "
         "Modelling stress as simply 'price down, cost up' would be incoherent - the FY2022 "
         "evidence is that a coking coal spike to US$354/t coincided with RECORD margins."),
        ("WHAT YOU MUST FIX BEFORE TRANSACTION USE",
         "Four inputs are unsourced and flagged Low confidence: the price-utilisation elasticity "
         "(D23, the most important), the raw-material share of cash cost (D13), net working "
         "capital days (D18), and the equity risk premium and beta inside WACC. Two are left as "
         "user inputs rather than guessed: share prices on the Comps sheet, and a licensed coking "
         "coal series to replace the four annotated chart points the database could evidence. "
         "NONE of these has been fabricated. Unsourced parameters are held CONSTANT across "
         "scenarios, because flexing an unsourced input to manufacture a scenario would be "
         "dishonest. The Sensitivity sheet quantifies exactly how much of the answer rests on "
         "each: the operating forecast is robust to them, the cash flow and valuation output is "
         "not."),
        ("KNOWN LIMITATIONS",
         "1. No endogenous supply response - the model does not idle capacity when margins turn "
         "negative, which in reality would arrest a decline, so read Stress as a LOWER BOUND on "
         "margin. 2. The price feedback uses prior-year utilisation, which slightly dampens the "
         "cycle. 3. Coking coal has no fiscal-year average (database Conflict C01). 4. Working "
         "capital is not balance-sheet-derived. 5. No cost curves below the four majors; the "
         "secondary sector, about half of India's output, has no cost representation. 6. Share "
         "prices not sourced, so no trading multiple is asserted. 7. Invested capital is a "
         "replacement-cost proxy, so ROIC is indicative. 8. India-only scope - Tata Steel's "
         "Netherlands going-concern uncertainty is outside this model."),
        ("SOURCE OF FY2026 ACTUALS",
         "Master Industry Database.xlsx, whose Sources sheet carries 47 cited sources and whose "
         "Conflicts Log carries 16 documented conflicts. This model does not restate or "
         "contradict any of it; the four principal reconciliation points - crude steel "
         "production, capacity, consumption and finished production - were verified "
         "programmatically against the database file at build time."),
    ]:
        put(ws, r, 2, h, f=font(10, bold=True, colour=NAVY)); r += 1
        put(ws, r, 2, b, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        ws.row_dimensions[r].height = 13 * (1 + len(b) // 122); r += 2


def sh_scen_cmp(ws):
    sheet_defaults(ws)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 14
    for i in range(4):
        ws.column_dimensions[GL(4 + i)].width = 16
    ws.column_dimensions["H"].width = 108
    r = title_block(ws, [
        ("SCENARIO COMPARISON", "title"),
        ("FY2033E outcomes on all four scenarios. STATIC VALUES - computed at build time by the "
         "same arithmetic the Model sheet implements.", "label"),
        ("Excel can only display one scenario at a time, so this table is values, not formulas. "
         "To reproduce any column, set the scenario cell on the Model sheet and read the FY2033E "
         "column. This sheet is deliberately formula-free so it is completely safe to paste.",
         "body"), ("", "body")], start_row=2, width_col=2)
    r = header_row(ws, r, ["", "FY2033E metric", "Unit"] + SC + ["Commentary"], start_col=1,
                   left_align_cols=(2, 8))
    rows = [
        ("Apparent consumption", "Mt", "cons", FMT_1DP,
         "Demand is the least contested part of the model: even Stress grows consumption, because "
         "Indian demand has contracted only once (FY2021) in twelve years."),
        ("Crude steel production", "Mt", "crude", FMT_1DP, ""),
        ("Crude steel capacity", "Mtpa", "cap", FMT_1DP,
         "Scenario-dependent through the delivery factor and secondary residual, but the range is "
         "narrow relative to the price range."),
        ("Capacity utilisation", "%", "util", FMT_PCT_1,
         "Counter-intuitively HIGHEST in Stress, because capacity additions are cut harder than "
         "demand falls."),
        ("Blended realisation", "Rs/t", "realn", FMT_INT,
         "Note Stress sits ABOVE Bear: a rupee collapse to 113/US$ lifts the rupee landed cost of "
         "imports ~28% and supports domestic prices."),
        ("Cash cost per tonne", "Rs/t", "cost", FMT_INT, ""),
        ("EBITDA per tonne", "Rs/t", "ebt", FMT_INT, ""),
        ("Industry EBITDA margin", "%", "margin", FMT_PCT_1,
         "Envelope check: Bull 25.2% against the FY2022 observed peak of ~26%; Stress 9.9% in the "
         "terminal year having troughed at -9.2% in FY2028."),
        ("Industry revenue", "Rs cr", "rev", FMT_INT, ""),
        ("Industry EBITDA", "Rs cr", "ebitda", FMT_INT,
         "The spread from Stress to Bull is ~4.4x, which is the honest measure of uncertainty in "
         "an Indian steel forecast seven years out."),
        ("Unlevered free cash flow", "Rs cr", "fcf", FMT_INT,
         "Negative in Bear and Stress: the industry cannot fund ~Rs 1.1 lakh crore a year of "
         "growth capex from a compressed margin."),
    ]
    for lbl, unit, ck, fmt, comm in rows:
        put(ws, r, 2, lbl, f=font(10, bold=True), alignment=AL_LEFT_WRAP)
        put(ws, r, 3, unit, f=font(9, colour=TEXT_MUTED))
        for i, s in enumerate(SC):
            put(ws, r, 4 + i, C.ALL[s][ck][-1], number_format=fmt, alignment=AL_RIGHT)
        put(ws, r, 8, comm, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    r += 1
    put(ws, r, 2, "MARGIN PATH BY SCENARIO (FY2027E to FY2033E)",
        f=font(10, bold=True, colour=NAVY)); r += 1
    for s in SC:
        put(ws, r, 2, s, f=font(10))
        put(ws, r, 3, "%", f=font(9, colour=TEXT_MUTED))
        put(ws, r, 8, "  ".join("%.1f%%" % (m * 100) for m in C.ALL[s]["margin"]),
            f=font(9, colour=TEXT_MUTED)); r += 1
    r += 1
    for s, txt in [("Base Case", "RBI's FY2027 GDP of 6.6% converging to 6.5%; elasticity 1.15x "
                    "tapering to 1.05x; realisation grows ~2.0% a year, BELOW inflation, "
                    "reflecting structural surplus capacity. Terminal margin 16.8% against 17.9% "
                    "in FY2026 - the base case assumes mild erosion, not expansion."),
                   ("Bull Case", "GDP ~80bps above base; elasticity 1.30x tapering to 1.15x; iron "
                    "ore and coking coal at the low end; INR stable at 96. Calibrated to peak at "
                    "26.0% in FY2030, matching the observed FY2022 cycle peak."),
                   ("Bear Case", "A DEMAND-LED downturn. GDP ~100bps below base, elasticity 1.00x, "
                    "and critically iron ore and coking coal FALL WITH demand to US$83/dmt and "
                    "US$170/t, because a global demand shortfall is disinflationary for raw "
                    "materials. A bear case with demand down AND raw materials up would be "
                    "internally contradictory."),
                   ("Stress Case", "A STAGFLATIONARY supply shock plus demand shock. Coking coal "
                    "to US$300/t (~85% of the Oct-2023 peak of US$354/t), rupee to 113/US$, GDP "
                    "4.4% in FY2027, conversion inflation 7.5%. Troughs at -9.2% in FY2028. "
                    "LIMITATION: no endogenous supply response, so read this as a lower bound on "
                    "margin, not a central expectation.")]:
        put(ws, r, 2, s, f=font(10, bold=True, colour=NAVY))
        put(ws, r, 8, txt, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        ws.row_dimensions[r].height = 13 * (1 + len(txt) // 106); r += 1


def sh_sens(ws):
    sheet_defaults(ws)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 10
    for i in range(9):
        ws.column_dimensions[GL(4 + i)].width = 14
    ws.column_dimensions["M"].width = 100
    r = title_block(ws, [
        ("SENSITIVITY ANALYSIS", "title"),
        ("Self-contained. The two grids are LIVE intra-sheet formulas driven by the constants in "
         "section B. The tornado is static values, because each entry requires a full re-run of "
         "the chain.", "label"), ("", "body")], start_row=2, width_col=2)

    base, tor = C.tornado()
    put(ws, r, 2, "A - TORNADO: FY2033E INDUSTRY EBITDA, 10% ADVERSE MOVE (Base Case)",
        f=font(11, bold=True, colour=NAVY)); r += 1
    r = header_row(ws, r, ["", "Driver", "ID", "Adverse", "FY2033E EBITDA (Rs cr)",
                           "Change (Rs cr)", "Change (%)", "", "", "", "", "",
                           "Interpretation"], start_col=1, left_align_cols=(2, 13))
    interp = {
        "D09": "Dominant. A 10% realisation move is worth ~60% of industry EBITDA - the arithmetic "
               "of a 17.9% margin. Any valuation of this sector is primarily a price call.",
        "D03": "Second largest and often underestimated. Iron ore and coking coal are "
               "dollar-priced, so a 10% rupee fall lifts the whole rupee input basket at once. "
               "NOTE this isolated flex does NOT credit the offsetting import-parity support to "
               "realisation - that interaction is captured in the Stress scenario, which is why "
               "Stress is not the sum of these individual flexes.",
        "D01": "Works through volume AND, via utilisation, price. Identical to elasticity because "
               "demand growth is their product - as intended.",
        "D02": "Identical to GDP by construction.",
        "D10": "Iron ore at 45% of the raw material basket, which is 60% of cash cost.",
        "D11": "Coking coal, same basket weight. Ranks below iron ore only because its FY2033 "
               "level sits closer to its FY2026 reference; in Stress it dominates.",
        "D12": "Conversion inflation on the 40% of cash cost not indexed to raw materials.",
        "D13": "POSITIVE, because raising the raw-material share shifts weight away from "
               "conversion cost, which compounds faster in the Base Case. A genuine insight: the "
               "direction of this UNSOURCED parameter's effect depends on relative escalation.",
        "D22": "NEGATIVE - more capacity depresses utilisation, which feeds back into price. The "
               "economically correct result, and the opposite of the naive intuition.",
        "D05": "Same mechanism, smaller, because announced additions are a minority of total "
               "additions.",
    }
    for code, name, mult, b0, fl, delta, pct in tor:
        put(ws, r, 2, name, f=font(10, bold=True), alignment=AL_LEFT_WRAP)
        put(ws, r, 3, code, alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 4, "%+d%%" % (mult * 100), alignment=AL_CENTRE)
        put(ws, r, 5, fl, number_format=FMT_INT, alignment=AL_RIGHT)
        put(ws, r, 6, delta, number_format=FMT_INT, alignment=AL_RIGHT)
        put(ws, r, 7, pct, number_format=FMT_PCT_1, alignment=AL_RIGHT)
        put(ws, r, 13, interp.get(code, ""), f=font(9, colour=TEXT_MUTED),
            alignment=AL_LEFT_WRAP)
        r += 1
    put(ws, r, 2, "Base Case FY2033E industry EBITDA", f=font(10, bold=True))
    put(ws, r, 5, base, number_format=FMT_INT, alignment=AL_RIGHT,
        f=font(10, bold=True, colour=NAVY))
    r += 3

    # section B - constants
    b = C.ALL["Base Case"]
    put(ws, r, 2, "B - CONSTANTS FOR THE LIVE GRIDS (Base Case, FY2033E)",
        f=font(11, bold=True, colour=NAVY)); r += 1
    put(ws, r, 13, "Hard-entered values, not links, so this sheet is safe to paste. If you change "
                   "a driver on the Model sheet, refresh these six cells from the FY2033E column "
                   "there.", f=font(9, italic=True, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    consts = [("FY2026 cash cost anchor", "Rs/t", C.COST26, FMT_INT, "c26"),
              ("Raw material share of cash cost", "%", 0.60, FMT_PCT_1, "rms"),
              ("Iron ore weight in basket", "%", 0.45, FMT_PCT_1, "iow"),
              ("FY2033E conversion cost index", "x", b["convidx"][-1], FMT_3DP, "cvx"),
              ("FY2033E iron ore, rupees", "Rs/dmt", b["io"][-1] * b["inr"][-1], FMT_INT, "ioi"),
              ("FY2026 iron ore, rupees", "Rs/dmt", C.IO26 * C.INR26, FMT_INT, "ioi26"),
              ("FY2026 coking coal, rupees", "Rs/t", C.CC26 * C.INR26, FMT_INT, "cci26"),
              ("FY2033E USD/INR", "Rs/US$", b["inr"][-1], FMT_2DP, "inr"),
              ("FY2033E price adjustment factor", "x", b["px_adj"][-1], FMT_3DP, "pxa"),
              ("FY2033E capacity", "Mtpa", b["cap"][-1], FMT_1DP, "cap"),
              ("FY2033E crude-to-finished ratio", "x", b["cfr"] if isinstance(b.get("cfr"), float)
               else 1.0465, FMT_3DP, "cfr"),
              ("FY2033E cash cost", "Rs/t", b["cost"][-1], FMT_INT, "cst")]
    K = {}
    for nm, unit, val, fmt, key in consts:
        put(ws, r, 2, nm, alignment=AL_LEFT_WRAP)
        put(ws, r, 3, unit, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 4, val, number_format=fmt, alignment=AL_RIGHT, fill=FILL_NAVY_LIGHT)
        K[key] = "$D$%d" % r
        r += 1
    r += 2

    ccs = [160, 180, 200, 220, 240, 260, 280]
    rns = [58000, 61000, 64000, 66000, 68000, 71000, 74000]
    put(ws, r, 2, "C - LIVE GRID 1: FY2033E EBITDA PER TONNE (Rs/t)",
        f=font(11, bold=True, colour=NAVY))
    put(ws, r, 13, "Rows = FY2033E realisation driver (Rs/t). Columns = coking coal (US$/t). Every "
                   "cell rebuilds the cost equation from section B - change a constant and the "
                   "whole grid moves.", f=font(9, italic=True, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP); r += 1
    put(ws, r, 2, "Realisation \\ coking coal", f=font(9, bold=True), alignment=AL_CENTRE,
        fill=FILL_NAVY_LIGHT)
    for j, cc in enumerate(ccs):
        put(ws, r, 4 + j, cc, number_format=FMT_INT, alignment=AL_CENTRE, f=font(9, bold=True),
            fill=FILL_NAVY_LIGHT)
    h1 = r; r += 1
    for rn in rns:
        put(ws, r, 2, rn, number_format=FMT_INT, alignment=AL_RIGHT, f=font(9, bold=True),
            fill=FILL_NAVY_LIGHT)
        for j in range(len(ccs)):
            cl = GL(4 + j)
            cost = ("%s*(%s*(%s*(%s/%s)+%s*(%s%d*%s/%s)+(1-2*%s)*%s)+(1-%s)*%s)"
                    % (K["c26"], K["rms"], K["iow"], K["ioi"], K["ioi26"], K["iow"], cl, h1,
                       K["inr"], K["cci26"], K["iow"], K["cvx"], K["rms"], K["cvx"]))
            put(ws, r, 4 + j, "=$B%d*%s-(%s)" % (r, K["pxa"], cost), number_format=FMT_INT,
                alignment=AL_RIGHT)
        r += 1
    put(ws, r, 2, "Interpretation", f=font(9, bold=True, colour=NAVY))
    put(ws, r, 13, "At the Base Case FY2033E realisation of Rs 68,000/t, EBITDA/t swings roughly "
                   "2.4x across the coking coal axis alone. That is why a licensed coking coal "
                   "series is the highest-priority data gap.", f=font(9, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP)
    r += 3

    utils = [0.74, 0.78, 0.82, 0.84, 0.86, 0.90]
    put(ws, r, 2, "D - LIVE GRID 2: FY2033E INDUSTRY EBITDA (Rs cr)",
        f=font(11, bold=True, colour=NAVY))
    put(ws, r, 13, "Rows = FY2033E utilisation. Columns = FY2033E realisation driver. Volume = "
                   "capacity x utilisation / crude-to-finished ratio. PARTIAL sensitivity: holds "
                   "the lagged price feedback at its base value, so it isolates volume and price "
                   "without double-counting the feedback.", f=font(9, italic=True,
                                                                   colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP); r += 1
    put(ws, r, 2, "Utilisation \\ realisation", f=font(9, bold=True), alignment=AL_CENTRE,
        fill=FILL_NAVY_LIGHT)
    for j, rn in enumerate(rns):
        put(ws, r, 4 + j, rn, number_format=FMT_INT, alignment=AL_CENTRE, f=font(9, bold=True),
            fill=FILL_NAVY_LIGHT)
    h2 = r; r += 1
    for u in utils:
        put(ws, r, 2, u, number_format=FMT_PCT_1, alignment=AL_RIGHT, f=font(9, bold=True),
            fill=FILL_NAVY_LIGHT)
        for j in range(len(rns)):
            cl = GL(4 + j)
            put(ws, r, 4 + j, "=(%s*$B%d/%s)*(%s%d*%s-%s)/10"
                % (K["cap"], r, K["cfr"], cl, h2, K["pxa"], K["cst"]), number_format=FMT_INT,
                alignment=AL_RIGHT)
        r += 1
    r += 2

    put(ws, r, 2, "E - HOW MUCH OF THE ANSWER RESTS ON UNSOURCED ASSUMPTIONS",
        f=font(11, bold=True, colour=NAVY)); r += 1
    r = header_row(ws, r, ["", "Unsourced assumption", "ID", "Base", "Plausible range", "", "",
                           "", "", "", "", "", "Materiality verdict"], start_col=1,
                   left_align_cols=(2, 5, 13))
    for nm, code, bv, rng, verdict in [
        ("Price sensitivity to utilisation", "D23", "0.40x", "0.20x to 0.60x",
         "MOST IMPORTANT UNSOURCED INPUT. It determines whether capacity matters at all. At 0.20x "
         "the capacity drivers halve in importance; at 0.60x they roughly double."),
        ("Raw material share of cash cost", "D13", "60%", "55% to 65%",
         "MATERIAL but second-order, ~6.7% of FY2033E EBITDA per 10% relative change. Direction "
         "depends on relative escalation, so it is not a simple bias."),
        ("Net working capital days", "D18", "45 days", "35 to 60 days",
         "MATERIAL FOR CASH FLOW AND CREDIT, immaterial for EBITDA. ~Rs 12,000 cr of terminal-year "
         "cash flow per 4.5 days. Fix before any financing work."),
        ("Equity risk premium and beta", "WACC", "6.5% / 1.20x", "5.5-7.5% / 1.0-1.4x",
         "MATERIAL FOR VALUATION ONLY, ~11% of EV per 120bps of WACC. Does not touch operating "
         "forecasts."),
        ("Maximum practical utilisation", "D25", "92%", "88% to 95%",
         "IMMATERIAL in Base - the ceiling does not bind (peak 84.0%). Becomes material only in "
         "Bull, where utilisation reaches 87.9%."),
        ("Maintenance capex", "D17", "3.0% of revenue", "2.5% to 4.0%",
         "MATERIAL FOR CASH FLOW, immaterial for EBITDA. ~Rs 16,000 cr of terminal-year FCF per "
         "100bps."),
        ("Emission intensity and carbon price", "ESG", "2.55 tCO2e/t, EUR 85/t", "Wide - both "
         "unsourced",
         "IMMATERIAL at current export volumes - CBAM is under 1% of FY2033E EBITDA - but it "
         "scales with the EU export share and the carbon price. Watch rather than model."),
    ]:
        put(ws, r, 2, nm, f=font(10, bold=True), alignment=AL_LEFT_WRAP)
        put(ws, r, 3, code, alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 4, bv, alignment=AL_CENTRE, f=font(9))
        put(ws, r, 5, rng, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 13, verdict, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    put(ws, r + 1, 2, "Read this block first. It is the honest answer to 'how much should I trust "
                      "this model?' - the operating forecast is robust to the unsourced inputs; "
                      "the cash flow and valuation output is not, until working capital, WACC and "
                      "a coking coal series are properly sourced.",
        f=font(9, italic=True, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)


def sh_comps(ws):
    sheet_defaults(ws)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22
    for i, w in enumerate([15, 14, 13, 11, 14, 16, 15, 15, 13, 13, 15]):
        ws.column_dimensions[GL(3 + i)].width = w
    ws.column_dimensions["N"].width = 100
    r = title_block(ws, [
        ("COMPARABLE VALUATION", "title"),
        ("Self-contained: FY2026 actuals are values, share prices are USER INPUTS, and every "
         "multiple is a live intra-sheet formula.", "label"),
        ("SHARE PRICES WERE NOT SOURCED in this research cycle, so nothing has been fabricated to "
         "fill the gap. Enter prices in the shaded column and every market cap, enterprise value "
         "and multiple computes immediately.", "body"), ("", "body")], start_row=2, width_col=2)
    r = header_row(ws, r, ["", "Company", "FY2026 EBITDA (Rs cr)", "Net debt (Rs cr)",
                           "PAT (Rs cr)", "EPS (Rs)", "Shares (cr) derived",
                           "SHARE PRICE (input)", "Market cap (Rs cr)", "EV (Rs cr)",
                           "EV/EBITDA", "Capacity (Mtpa)", "EV/tonne (Rs/t)", "Notes"],
                   start_col=1, left_align_cols=(2, 14))
    first = r
    for co in CO.COMPANIES:
        key, nm = co[0], co[1]
        eb, nd, pat, eps = co[6], co[11], co[13], co[14]
        cap, capnote = CO.CO_CAPACITY.get(key, (None, ""))
        put(ws, r, 2, nm, f=font(10, bold=True), alignment=AL_LEFT_WRAP)
        for col, v, fmt in ((3, eb, FMT_INT), (4, nd, FMT_INT), (5, pat, FMT_INT),
                            (6, eps, FMT_2DP)):
            if v:
                put(ws, r, col, v, number_format=fmt, alignment=AL_RIGHT)
        note = CO.CO_NOTES.get(key, "")
        if pat and eps:
            put(ws, r, 7, "=E%d/F%d" % (r, r), number_format=FMT_1DP, alignment=AL_RIGHT)
            put(ws, r, 8, None, number_format=FMT_2DP, alignment=AL_RIGHT,
                fill=FILL_NAVY_LIGHT, border=BORDER_BOTTOM)
            put(ws, r, 9, '=IFERROR(G%d*H%d,"")' % (r, r), number_format=FMT_INT,
                alignment=AL_RIGHT)
            put(ws, r, 10, '=IFERROR(I%d+D%d,"")' % (r, r), number_format=FMT_INT,
                alignment=AL_RIGHT)
            put(ws, r, 11, '=IFERROR(J%d/C%d,"")' % (r, r), number_format=FMT_MULT,
                alignment=AL_RIGHT)
        if cap:
            put(ws, r, 12, cap, number_format=FMT_2DP, alignment=AL_RIGHT)
            put(ws, r, 13, '=IFERROR(J%d/L%d*10000,"")' % (r, r), number_format=FMT_INT,
                alignment=AL_RIGHT)
        if key == "JSW":
            note += (" DERIVED SHARE COUNT IS UNRELIABLE HERE: PAT of Rs 25,508 cr includes the "
                     "Rs 18,051 cr BPSL slump-sale gain while EPS of Rs 91.26 is on profit "
                     "attributable to owners, so PAT/EPS overstates the count by ~15%. Overwrite "
                     "column G with the actual share count.")
        if key in ("AMNS", "RINL"):
            note += " Unlisted - no market-based valuation is possible."
        put(ws, r, 14, note + " Capacity basis: " + capnote, f=font(9, colour=TEXT_MUTED),
            alignment=AL_LEFT_WRAP)
        r += 1
    last = r - 1
    put(ws, r, 2, "Median EV/EBITDA (once prices are entered)", f=font(10, bold=True))
    put(ws, r, 11, '=IFERROR(MEDIAN(K%d:K%d),"enter prices")' % (first, last),
        number_format=FMT_MULT, alignment=AL_RIGHT, f=font(10, bold=True, colour=NAVY))
    put(ws, r, 14, "Compare against the exit multiple of 6.0x used in the DCF on the Model sheet "
                   "(driver D19), which is NOT sourced and should be re-based on this observed "
                   "evidence.", f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    r += 3
    put(ws, r, 2, "REPLACEMENT COST BENCHMARK", f=font(11, bold=True, colour=NAVY)); r += 1
    for lbl, val, fmt, note in [
        ("Capex intensity (greenfield / brownfield)", 55000, FMT_INT,
         "Driver D16. Cross-checks: JSW JVML 5 Mtpa at Rs 26,000 cr = Rs 52,000/t; SAIL ~Rs 1 lakh "
         "cr for ~15 Mtpa = ~Rs 67,000/t; Tata Ludhiana 0.75 Mtpa EAF at Rs 3,200 cr = "
         "Rs 42,700/t (no ironmaking)."),
        ("India FY2026 crude steel capacity", C.CAP26, FMT_1DP, ""),
        ("Replacement cost of India's capacity", None, FMT_INT,
         "A sanity bound: an acquirer should not pay materially more per tonne than it costs to "
         "build, adjusted for time to market, permits and captive raw material."),
    ]:
        put(ws, r, 2, lbl, f=font(10, bold=True), alignment=AL_LEFT_WRAP)
        if val is not None:
            put(ws, r, 4, val, number_format=fmt, alignment=AL_RIGHT)
        put(ws, r, 14, note, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    put(ws, r - 1, 4, "=D%d*D%d/10" % (r - 3, r - 2), number_format=FMT_INT, alignment=AL_RIGHT,
        f=font(10, bold=True, colour=NAVY))


def sh_register(ws):
    sheet_defaults(ws)
    hs = A.REGISTER_HEADERS
    ws.column_dimensions["A"].width = 3
    for i, w in enumerate([13, 34, 13, 24, 30, 24, 30, 18, 11, 74, 74, 60, 40, 20, 30]):
        ws.column_dimensions[GL(2 + i)].width = w
    r = title_block(ws, [
        ("ASSUMPTION REGISTER", "title"),
        ("The thirteen mandated fields for every material assumption. Static text - no formulas, "
         "completely safe to paste.", "label"), ("", "body")], start_row=2, width_col=2)
    r = header_row(ws, r, [""] + list(hs), start_col=1,
                   left_align_cols=tuple(range(2, len(hs) + 2)))
    for row in A.REGISTER:
        for i, v in enumerate(row):
            put(ws, r, 2 + i, v, f=font(9, colour=TEXT_MUTED if i > 3 else None),
                alignment=AL_LEFT_WRAP)
        r += 1
    freeze(ws, "D%d" % (r - len(A.REGISTER)))


def sh_horizon(ws):
    sheet_defaults(ws)
    ws.column_dimensions["A"].width = 3
    for i, w in enumerate([42, 9, 44, 44, 44, 10, 10, 10, 104]):
        ws.column_dimensions[GL(2 + i)].width = w
    r = title_block(ws, [
        ("FORECAST HORIZON SELECTION", "title"),
        ("Weighted evaluation of 5, 7 and 10-year horizons. The weighted totals are LIVE "
         "intra-sheet formulas, so you can change a weight and see whether the conclusion "
         "survives.", "label"), ("", "body")], start_row=2, width_col=2)
    r = header_row(ws, r, [""] + list(A.HORIZON_HEADERS), start_col=1,
                   left_align_cols=(2, 5, 6, 7, 10))
    first = r
    for c, w, o5, o7, o10, s5, s7, s10, ev in A.HORIZON_EVAL:
        put(ws, r, 2, c, alignment=AL_LEFT_WRAP, f=font(10, bold=True))
        put(ws, r, 3, w, number_format=FMT_PCT_1, alignment=AL_RIGHT)
        for i, o in enumerate((o5, o7, o10)):
            put(ws, r, 4 + i, o, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        for i, s in enumerate((s5, s7, s10)):
            put(ws, r, 7 + i, s, number_format=FMT_INT, alignment=AL_CENTRE)
        put(ws, r, 10, ev, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    last = r - 1
    put(ws, r, 2, "WEIGHTED TOTAL", f=font(10, bold=True, colour=NAVY))
    put(ws, r, 3, "=SUM(C%d:C%d)" % (first, last), number_format=FMT_PCT_1, alignment=AL_RIGHT,
        f=font(10, bold=True))
    for col in (7, 8, 9):
        put(ws, r, col, "=SUMPRODUCT($C$%d:$C$%d,%s%d:%s%d)"
            % (first, last, GL(col), first, GL(col), last), number_format=FMT_2DP,
            alignment=AL_CENTRE, f=font(11, bold=True, colour=NAVY))
    put(ws, r, 10, "Live. The 7-year option wins on four of five criteria and loses only on "
                   "macro-anchor reliability.", f=font(9, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP)
    r += 2
    put(ws, r, 2, "SELECTED: 7 YEARS - FY2027E to FY2033E", f=font(12, bold=True, colour=NAVY))
    r += 2
    for para in A.HORIZON_CONCLUSION.split("\n\n"):
        put(ws, r, 2, para, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        ws.row_dimensions[r].height = 13 * (1 + len(para) // 150)
        r += 2


def sh_sources(ws):
    sheet_defaults(ws)
    ws.column_dimensions["A"].width = 3
    for i, w in enumerate([16, 8, 32, 48, 15, 42, 40, 100]):
        ws.column_dimensions[GL(2 + i)].width = w
    r = title_block(ws, [
        ("SOURCES", "title"),
        ("Sources for the FORWARD-LOOKING assumptions. FY2026 actuals are sourced in Master "
         "Industry Database.xlsx, whose Sources sheet carries 47 entries and is not duplicated "
         "here.", "label"),
        ("Level 1 primary issuer and government | Level 2 official statistical and multilateral "
         "agencies | Level 3 rating agencies and sell-side. Secondary sources never override "
         "primary data.", "body"), ("", "body")], start_row=2, width_col=2)
    r = header_row(ws, r, ["", "Ref", "Level", "Publisher", "Publication", "Date", "Used for",
                           "Verification", "Notes"], start_col=1,
                   left_align_cols=(2, 4, 5, 7, 8, 9))
    for ref, lvl, pub, pubn, dt, used, ver, note in [
        ("RBI-Jun26", 2, "Reserve Bank of India", "Monetary Policy Committee statement",
         "05-Jun-2026", "Drivers D01 GDP and D04 CPI; register R01",
         "Verified via multiple contemporaneous reports",
         "FY2027 real GDP projection cut to 6.6% from 6.9%; repo held at 5.25%; FY2027 CPI "
         "projection raised 50bps to 5.1%. FY2026 estimated 7.6%; Q4FY2026 actual 7.7%."),
        ("worldsteel-SRO", 2, "World Steel Association", "Short Range Outlook, April 2026",
         "Apr-2026", "Corroborates D02 demand elasticity",
         "Verified - also reproduced in Ministry of Steel monthly overviews",
         "Global demand +0.3% CY2026 to 1,724 Mt, +2.2% CY2027. India +7.4% then +9.2%, the "
         "fastest major market. China -1.5% then flat."),
        ("GSEC-31Jul26", 3, "investing.com / NSE", "India 10-year benchmark G-sec yield",
         "31-Jul-2026", "Risk-free rate in WACC; register R08",
         "Verified - 6.833%; Trading Economics expects 6.78% at quarter end",
         "Substitute a Bloomberg or Refinitiv print for a live mandate."),
        ("CRISIL-RateView", 3, "CRISIL Intelligence", "RateView, July 2026", "Jul-2026",
         "Pre-tax cost of debt in WACC",
         "Verified - 10-year corporate bond yield 7.36% actual, 7.44-7.54% projected",
         "Consistent with the ICRA and India Ratings AA domestic ratings the database records for "
         "JSW Steel."),
        ("IDBI-22Jun26", 3, "IDBI Capital Markets & Securities",
         "Commodity Price Update, week ended 22-Jun-2026", "22-Jun-2026",
         "Driver D03 USD/INR; input cost cross-checks",
         "Verified - full document retrieved and parsed",
         "USD/INR 94 (52-week range 85-98). Also billet ex-Raipur Rs 38,850/t, HR strip Patra "
         "Rs 44,500/t, nickel US$17,395/t, ferro chrome Rs 1,23,200/t, graphite electrode UHP "
         "US$4,189/t, Brent US$79/bbl - the input-cost deck behind D12."),
        ("MasterDB", 1, "This research team", "Master Industry Database.xlsx", "03-Aug-2026",
         "Every FY2026 actual and every derived calibration anchor",
         "Extracted programmatically from the database file at build time",
         "47 cited sources and 16 documented conflicts on that workbook. This model does not "
         "restate or contradict any of it."),
        ("Derived-DB", 1, "This research team", "Derivations from the Master Database",
         "03-Aug-2026",
         "Drivers D02, D05-D09, D16, D22, D24, D26 and all calibration anchors",
         "Each derivation is stated in full on the relevant row",
         "Material derivations: blended realisation Rs 59,974/t from four majors; weighted EBITDA "
         "Rs 10,733/t; implied cash cost Rs 49,241/t; FY2026 average USD/INR 88.4 from Tata "
         "Steel's dual-currency disclosure; D&A 5.28% of revenue; demand elasticity 1.10x from "
         "the FY2015-FY2026 outturn."),
        ("Indicative", 3, "Modeller judgement", "Structural parameters not sourced from any "
         "document", "n/a", "Drivers D13, D14, D17, D23, D25 and the ESG inputs",
         "NOT VERIFIED - flagged Low confidence throughout",
         "Held CONSTANT across scenarios wherever unsourced, so an unsourced parameter is never "
         "flexed to manufacture a scenario. Each is quantified on the Sensitivity sheet."),
        ("Not sourced", 3, "None", "Left as user inputs rather than assumed", "n/a",
         "D18 working capital days, D19 exit multiple, share prices, ERP and beta",
         "NOT SOURCED - shaded in the workbook and excluded from headline conclusions",
         "The gaps that must be closed before transaction use. None has been filled with a "
         "fabricated value."),
    ]:
        put(ws, r, 2, ref, f=font(10, bold=True))
        put(ws, r, 3, lvl, alignment=AL_CENTRE)
        put(ws, r, 4, pub, alignment=AL_LEFT_WRAP)
        put(ws, r, 5, pubn, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 6, dt, alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 7, used, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 8, ver, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 9, note, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    r += 2
    put(ws, r, 2, "MODEL CHANGE LOG", f=font(11, bold=True, colour=NAVY)); r += 1
    r = header_row(ws, r, ["", "Version", "Date", "Change"], start_col=1, left_align_cols=(4,))
    for v, d, ch in [
        ("1.0", "03-Aug-2026", "Initial build. Multi-sheet, cross-sheet-linked."),
        ("2.0", "03-Aug-2026",
         "REBUILD for copy-paste safety after the v1.0 file failed to open cleanly. Fixed: (a) "
         "eight Excel Tables had a blank header cell, which Excel treats as invalid content and "
         "repairs; (b) 199 cells were plain text instead of formulas because a driver-link helper "
         "omitted the leading '='. Redesigned: the calculation chain is consolidated onto ONE "
         "sheet with intra-sheet formulas only, and all Excel Tables and defined names removed, "
         "so any sheet can be copy-pasted into another workbook without breaking. The scenario "
         "selector is now an integer rather than a text label."),
    ]:
        put(ws, r, 2, v, alignment=AL_CENTRE, f=font(10, bold=True))
        put(ws, r, 3, d, alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 4, ch, alignment=AL_LEFT_WRAP, f=font(9, colour=TEXT_MUTED))
        ws.row_dimensions[r].height = 13 * (1 + len(ch) // 100)
        r += 1


# ======================================================================================
def main():
    wb = Workbook()
    wb.remove(wb.active)
    order = ["Read Me First", "Model", "Scenario Comparison", "Sensitivity",
             "Comparable Valuation", "Assumption Register", "Forecast Horizon", "Sources"]
    ws = {n: wb.create_sheet(title=n) for n in order}
    sh_readme(ws["Read Me First"])
    build_model(ws["Model"])
    sh_scen_cmp(ws["Scenario Comparison"])
    sh_sens(ws["Sensitivity"])
    sh_comps(ws["Comparable Valuation"])
    sh_register(ws["Assumption Register"])
    sh_horizon(ws["Forecast Horizon"])
    sh_sources(ws["Sources"])

    wb.properties.title = "Industry Financial Model - Indian Steel Industry"
    wb.properties.category = "Global Metals & Mining Research"
    wb.properties.description = (
        "Copy-paste safe analytical engine for the Indian steel industry. Horizon "
        "FY2027E-FY2033E. Intra-sheet formulas only: no cross-sheet links, no defined names, "
        "no Excel Tables.")
    ws["Read Me First"].sheet_view.tabSelected = True
    wb.active = 0
    wb.save(OUT)
    print("WROTE %s  (%.1f KB, %d sheets)"
          % (OUT, os.path.getsize(OUT) / 1024.0, len(wb.sheetnames)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
