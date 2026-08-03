#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build "Industry Financial Model.xlsx" - Indian steel industry analytical engine.

Consumes Master Industry Database.xlsx. Run:  python3 build_financial_model.py
"""
import os
import sys

import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as GL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from builder.style import (AL_CENTRE, AL_LEFT, AL_LEFT_WRAP, AL_RIGHT, BORDER_BOTTOM,
                           BORDER_TOP_RULE, FILL_NAVY_LIGHT, FMT_1DP, FMT_2DP, FMT_3DP,
                           FMT_INT, FMT_MULT, FMT_PCT_1, FMT_TEXT, NA, NAVY, TEXT_MUTED,
                           add_table, font, footnote, freeze, header_row, put,
                           repeat_header, set_widths, sheet_defaults, title_block)
from fmodel import assumptions as A
from fmodel import chain as C
from fmodel import companies as CO

OUT = "Industry Financial Model.xlsx"
DB = "Master Industry Database.xlsx"

CB, CF1, CFN, CL = 4, 5, 11, 12          # D base, E..K forecast, L CAGR
CSRC, CCONF, CNOTE = 13, 14, 15          # M source, N confidence, O notes
FCOLS = list(range(CF1, CFN + 1))
FY = A.FY
NY = A.N
SC = A.SCENARIOS

RW = {}
_ORDER = sorted(A.DRIVERS, key=lambda d: d[0])


def K(ws, key, row):
    RW[(ws.title, key)] = row
    return row


def R(sheet, key):
    return RW[(sheet, key)]


def cr(sheet, key, col):
    """Absolute-row reference to a tracked line on another sheet."""
    return "'%s'!%s$%d" % (sheet, GL(col), R(sheet, key))


# --------------------------------------------------------------------------------------
def widths(ws, note=95):
    set_widths(ws, [(1, 7), (2, 52), (3, 17), (CB, 13)]
               + [(c, 13) for c in FCOLS]
               + [(CL, 12), (CSRC, 22), (CCONF, 12), (CNOTE, note)])


def meta(ws, num, name, objective, method, units, deps, note=95):
    widths(ws, note)
    lines = [("%s  %s" % (num, name) if num else name, "title"),
             ("Industry Financial Model  |  Indian Steel Industry  |  Horizon FY2027E-FY2033E  |  "
              "Base year FY2026A  |  Built 03-Aug-2026", "label"), ("", "body"),
             ("OBJECTIVE", "label"), (objective, "body"),
             ("METHODOLOGY", "label"), (method, "body"),
             ("UNITS", "label"), (units, "body"),
             ("DEPENDENCIES", "label"), (deps, "body"), ("", "body")]
    return title_block(ws, lines, width_col=2)


def hdr(ws, r, extra_left=("Source", "Conf.", "Notes / methodology")):
    hs = ["", "Line item", "Unit", "FY2026A"] + FY + ["CAGR"] + list(extra_left)
    return header_row(ws, r, hs, start_col=1, left_align_cols=(2, 3, CSRC, CNOTE))


def line(ws, r, label, unit, base, cells, src="", conf="", note="", fmt=FMT_1DP,
         bold=False, key=None, cagr=False, ident="", fill=None):
    f = font(10, bold=bold, colour=NAVY if bold else None) if bold else font()
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
        put(ws, r, CL, '=IFERROR(IF(AND(ISNUMBER(%s%d),ISNUMBER(%s%d),%s%d>0),'
                       '(%s%d/%s%d)^(1/%d)-1,""),"")'
            % (GL(CB), r, GL(CFN), r, GL(CB), r, GL(CFN), r, GL(CB), r, NY),
            number_format=FMT_PCT_1, alignment=AL_RIGHT)
    if src:
        put(ws, r, CSRC, src, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    if conf:
        put(ws, r, CCONF, conf, f=font(9, colour=TEXT_MUTED), alignment=AL_CENTRE)
    if note:
        put(ws, r, CNOTE, note, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    if key:
        K(ws, key, r)
    return r + 1


def sect(ws, r, text):
    put(ws, r, 2, text, f=font(10, bold=True, colour=NAVY))
    return r + 1


def dv_row(code):
    return R("02 Model Assumptions", "act_" + code)


def dv(code, col):
    return "'02 Model Assumptions'!%s$%d" % (GL(col), dv_row(code))


# ======================================================================================
def sh_cover(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 3), (2, 120)])
    r = title_block(ws, [
        ("INDUSTRY FINANCIAL MODEL", "title"),
        ("INDIAN STEEL INDUSTRY - ANALYTICAL ENGINE", "subtitle"), ("", "body"),
        ("Global Metals & Mining Research", "label"),
        ("Forecast horizon FY2027E-FY2033E (7 years)  |  Base year FY2026A  |  "
         "Built 03 August 2026", "body"), ("", "body")], start_row=2, width_col=2)
    for h, b in [
        ("WHAT THIS MODEL IS",
         "An analytical engine, not a database. It consumes FY2026 actuals from Master Industry "
         "Database.xlsx and forecasts Indian steel demand, supply, capacity, utilisation, "
         "prices, costs, revenue, EBITDA, working capital, cash flow and trade to FY2033, under "
         "four fully integrated scenarios. It is the parent model for single-name DCF, trading "
         "and transaction comparables, EV per tonne, replacement cost and ROIC work on Tata "
         "Steel, JSW Steel, SAIL, Jindal Steel, Jindal Stainless and AM/NS India."),
        ("HOW SCENARIOS WORK",
         "Change one cell - the scenario selector on 01 Control Panel - and the entire workbook "
         "re-computes. The selector drives a named range ScenID, which every driver on 02 Model "
         "Assumptions resolves through an INDEX across its four scenario rows. No sheet contains "
         "a scenario-specific hardcode. There is nothing else to change."),
        ("THE ONE THING THAT MAKES THIS AN ENGINE RATHER THAN A SPREADSHEET",
         "A lagged price-utilisation feedback (driver D23). Effective realisation is adjusted by "
         "an elasticity applied to the gap between the PRIOR year's capacity utilisation and "
         "normal utilisation. Because it is lagged there is no circular reference, but it means "
         "capacity, pipeline delivery rates and secondary-sector additions genuinely transmit "
         "into price and therefore EBITDA. Without it, capacity would be decorative and the "
         "capacity drivers would show literally zero EBITDA sensitivity. It also produces the "
         "economically correct and counter-intuitive result that FASTER capacity delivery is "
         "EBITDA-NEGATIVE, because it depresses utilisation and therefore price."),
        ("INDEPENDENT VALIDATION OF THE CAPACITY BLOCK",
         "Built bottom-up from a dated 20-project tracker plus a secondary-sector residual, the "
         "model produces India crude steel capacity of 302.0 Mtpa in FY2031 against the National "
         "Steel Policy 2017 target of 300 Mtpa - within 0.7%. The model was not calibrated to "
         "that target, so this is a genuine external check."),
        ("SCENARIOS ARE CALIBRATED, NOT INVENTED",
         "The four paths are calibrated to margin envelopes actually observed in the database "
         "rather than set independently. Bull peaks at a 26.0% industry EBITDA margin in FY2030, "
         "matching the FY2022 cycle peak (Tata 26%, JSW 27%, Jindal Steel 30%). Stress troughs "
         "at -9.2% in FY2028, marginally worse than the worst single observation in the database "
         "(SAIL FY2016, -7%). Critically, the stress case carries HIGHER nominal rupee "
         "realisation than the bear case despite worse demand, because a rupee collapse to "
         "113/US$ lifts the rupee landed cost of imports by roughly 28% and supports domestic "
         "prices. Modelling stress as simply 'price down, cost up' would be incoherent: the "
         "FY2022 evidence is that a coking coal spike coincided with RECORD margins."),
        ("WHAT YOU MUST FIX BEFORE TRANSACTION USE",
         "Four inputs are unsourced and are flagged Low confidence throughout: the equity risk "
         "premium and beta inside WACC (R08), net working capital days (R09), the raw-material "
         "share of cash cost (R10), and the price-utilisation elasticity (D23). Two are not "
         "sourced at all and are user inputs: share prices on 23 Comparable Valuation, and a "
         "licensed coking coal series to replace the four annotated chart points the database "
         "was able to evidence. None of these has been fabricated; each is isolated, flagged, "
         "and given a sensitivity axis so you can see exactly how much of the answer depends "
         "on it."),
        ("DISCIPLINE",
         "No assumption is hardcoded inside a formula. Every forecast cell references either 02 "
         "Model Assumptions or 00 Database Import. Column alignment is identical on every "
         "forecast sheet - D is FY2026A and E to K are FY2027E to FY2033E without exception - so "
         "any cell can reference the same column on any other sheet. 25 Audit Checks runs 30 "
         "live validations including a tie-out that proves the Excel formulas reproduce the "
         "intended arithmetic."),
    ]:
        put(ws, r, 2, h, f=font(10, bold=True, colour=NAVY)); r += 1
        put(ws, r, 2, b, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        ws.row_dimensions[r].height = 13 * (1 + len(b) // 118); r += 2
    put(ws, r, 2, "Public sources only. No non-public or price-sensitive information.",
        f=font(8, italic=True, colour=TEXT_MUTED))


def sh_contents(ws, specs):
    sheet_defaults(ws)
    set_widths(ws, [(1, 6), (2, 32), (3, 104), (4, 16)])
    r = title_block(ws, [("CONTENTS", "title"),
                         ("Click a sheet name to navigate.", "body"), ("", "body")], width_col=2)
    r = header_row(ws, r, ["#", "Worksheet", "Purpose and principal content", "Type"],
                   left_align_cols=(2, 3))
    for n, nm, sc, tp in specs:
        put(ws, r, 1, n, alignment=AL_CENTRE)
        c = put(ws, r, 2, nm, f=font(10, bold=True, colour=NAVY))
        c.hyperlink = "#'%s'!A1" % nm
        put(ws, r, 3, sc, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 4, tp, alignment=AL_CENTRE); r += 1
    freeze(ws, "A%d" % (r - len(specs)))


# ======================================================================================
def sh_import(ws):
    r = meta(ws, "00", "Database Import",
             "The single controlled boundary between Master Industry Database.xlsx and this "
             "model. Every FY2026 actual the model uses enters here, once, with its exact "
             "location in the database recorded. Nothing downstream references the database "
             "directly.",
             "DESIGN CHOICE, STATED EXPLICITLY. Values are imported as static numbers with full "
             "provenance rather than as live external-workbook links. Live links to another "
             "file are an operational risk in a model that will be circulated - they break on "
             "move or rename, they prompt on open, and they silently return stale values if the "
             "source is unavailable. Instead each row records the database sheet and the exact "
             "line it came from, and the 'Refresh formula' column contains the external-link "
             "formula as TEXT, ready to paste if a live link is wanted. The refresh procedure is "
             "at the foot of this sheet. Every value here was extracted programmatically from "
             "the database file at build time, so a transcription error is structurally "
             "impossible.",
             "As stated per row.",
             "Source: Master Industry Database.xlsx. Consumed by every forecast sheet.", 100)
    hs = ["ID", "Imported item", "Unit", "FY2026A value", "Master Database location",
          "Source ID in database", "Refresh formula (paste to activate a live link)",
          "Confidence", "Notes"]
    set_widths(ws, [(1, 7), (2, 46), (3, 13), (4, 15), (5, 62), (6, 15), (7, 62), (8, 11),
                    (9, 100)])
    r = header_row(ws, r, hs, left_align_cols=(2, 5, 6, 7, 9))
    first = r
    for aid, name, unit, val, loc, sid in A.ANCHORS:
        put(ws, r, 1, aid, alignment=AL_CENTRE)
        put(ws, r, 2, name, alignment=AL_LEFT_WRAP)
        put(ws, r, 3, unit, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 4, val, number_format=FMT_3DP if unit in ("Mt", "Mtpa", "x") else FMT_INT,
            alignment=AL_RIGHT, f=font(10, bold=True))
        put(ws, r, 5, loc.split(". VINTAGE")[0].split(". June")[0],
            f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 6, sid, f=font(9, colour=TEXT_MUTED), alignment=AL_CENTRE)
        put(ws, r, 7, "='[%s]%s'!<cell>" % (DB, loc.split(" > ")[1] if " > " in loc else "Sheet"),
            f=font(8, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 8, "High", alignment=AL_CENTRE)
        put(ws, r, 9, loc if "VINTAGE" in loc or "June" in loc or "Derived" in loc else "",
            f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        K(ws, aid, r); r += 1
    add_table(ws, "tbl_Import", first - 1, 1, r - 1, 9)
    freeze(ws, "B%d" % first)
    r += 1
    for t in [
        "REFRESH PROCEDURE. (1) Open both workbooks. (2) For each row, read the value at the "
        "stated Master Database location. (3) Compare against column D. (4) If different, "
        "overwrite column D and record the change in the model change log on 26 Sources. "
        "(5) Re-run 25 Audit Checks. Nothing else needs to change: every downstream sheet "
        "references this sheet, so a single correction propagates.",
        "DERIVED IMPORTS. A09 to A13 and A08 are DERIVED from database figures rather than "
        "lifted from a single cell, and their derivation is stated in the Notes column. They are "
        "imported here rather than computed downstream so that the calibration of the entire "
        "model sits in one auditable place.",
        "VINTAGE. The model uses the April-2026 Joint Plant Committee vintage for the supply "
        "balance because that balance is internally self-reconciling - production plus imports "
        "less exports less stock change equals consumption - which a forecast identity requires. "
        "The database's June-2026 vintage figures (finished production 161.74 Mt, consumption "
        "164.19 Mt) are carried as reconciliation memos. This is the resolution the database "
        "itself recommends under its Conflict C03.",
    ]:
        put(ws, r, 2, t, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        ws.row_dimensions[r].height = 13 * (1 + len(t) // 150); r += 2


def sh_horizon(ws):
    r = meta(ws, "0H", "Forecast Horizon",
             "Evaluates 5-year, 7-year and 10-year forecast horizons against weighted criteria "
             "and selects one. The selected horizon is applied identically on every forecast "
             "sheet in the workbook.",
             "Five criteria, weighted. Each option is scored 1 to 5 against each criterion. The "
             "weighted total is a live formula, not a hardcoded result, so a reviewer who "
             "disagrees with a weight can change it and see whether the conclusion survives.",
             "Scores 1 (worst) to 5 (best). Weights sum to 1.00.",
             "Determines the column structure of every forecast sheet.", 110)
    set_widths(ws, [(1, 7), (2, 44), (3, 9), (4, 46), (5, 46), (6, 46), (7, 10), (8, 10),
                    (9, 10), (10, 110)])
    r = header_row(ws, r, A.HORIZON_HEADERS, left_align_cols=(1, 4, 5, 6, 9))
    first = r
    for c, w, o5, o7, o10, s5, s7, s10, ev in A.HORIZON_EVAL:
        put(ws, r, 1, c, alignment=AL_LEFT_WRAP)
        put(ws, r, 2, w, number_format=FMT_PCT_1, alignment=AL_RIGHT)
        for i, o in enumerate((o5, o7, o10)):
            put(ws, r, 3 + i, o, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        for i, s in enumerate((s5, s7, s10)):
            put(ws, r, 6 + i, s, number_format=FMT_INT, alignment=AL_CENTRE)
        put(ws, r, 9, ev, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    last = r - 1
    put(ws, r, 1, "WEIGHTED TOTAL", f=font(10, bold=True, colour=NAVY))
    put(ws, r, 2, "=SUM(B%d:B%d)" % (first, last), number_format=FMT_PCT_1, alignment=AL_RIGHT,
        f=font(10, bold=True))
    for i, col in enumerate((6, 7, 8)):
        put(ws, r, col, "=SUMPRODUCT($B$%d:$B$%d,%s%d:%s%d)"
            % (first, last, GL(col), first, GL(col), last), number_format=FMT_2DP,
            alignment=AL_CENTRE, f=font(10, bold=True, colour=NAVY))
    put(ws, r, 9, "Live formula. The 7-year option wins on four of five criteria and ties on "
                  "none; it loses only on macro-anchor reliability, where the 5-year option is "
                  "stronger.", f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    add_table(ws, "tbl_Horizon", first - 1, 1, last, 9)
    r += 2
    put(ws, r, 1, "SELECTED: 7 YEARS - FY2027E to FY2033E", f=font(12, bold=True, colour=NAVY))
    r += 2
    for para in A.HORIZON_CONCLUSION.split("\n\n"):
        put(ws, r, 1, para, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        ws.row_dimensions[r].height = 13 * (1 + len(para) // 150); r += 2


def sh_control(ws):
    r = meta(ws, "01", "Control Panel",
             "The only cell in this workbook a user needs to change. The scenario selector "
             "drives every forecast on every sheet.",
             "The selector is a validated dropdown. It resolves to the named range ScenID via "
             "MATCH against the scenario list below. Every driver on 02 Model Assumptions uses "
             "INDEX with ScenID to pick its active row. Consequently no sheet anywhere in the "
             "workbook contains a scenario-specific hardcode.",
             "As stated per row.",
             "Drives 02 Model Assumptions and therefore every sheet from 03 to 24.", 100)
    put(ws, r, 2, "SCENARIO SELECTOR", f=font(11, bold=True, colour=NAVY)); r += 1
    put(ws, r, 2, "Active scenario", f=font(10, bold=True))
    c = put(ws, r, CB, "Base Case", f=font(11, bold=True, colour=NAVY), alignment=AL_CENTRE,
            fill=FILL_NAVY_LIGHT, border=BORDER_BOTTOM)
    sel_row = r
    d = DataValidation(type="list", formula1='"%s"' % ",".join(SC), allow_blank=False,
                       showDropDown=False, errorTitle="Invalid scenario",
                       error="Choose one of: %s" % ", ".join(SC))
    ws.add_data_validation(d); d.add(c)
    put(ws, r, CNOTE, "CHANGE THIS CELL ONLY. Dropdown-validated. Everything else recomputes.",
        f=font(9, italic=True, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP); r += 1
    put(ws, r, 2, "Scenario index (ScenID)", f=font(10))
    put(ws, r, CB, "=MATCH(%s$%d,%s$%d:%s$%d,0)" % (GL(CB), sel_row, GL(CB), r + 2, GL(CB), r + 1 + len(SC)),
        number_format=FMT_INT, alignment=AL_CENTRE, f=font(10, bold=True))
    id_row = r
    put(ws, r, CNOTE, "Named range ScenID. Resolves the selector to 1-4 for the INDEX lookups "
                      "on 02 Model Assumptions.", f=font(9, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP); r += 2
    put(ws, r, 2, "Scenario list (do not reorder - ScenID depends on this order)",
        f=font(9, bold=True, colour=TEXT_MUTED)); r += 1
    list_first = r
    desc = {"Base Case": "Central case. RBI FY2027 GDP of 6.6%, realisation calibrated to a "
                         "FY2033 industry margin of 16.8%.",
            "Bull Case": "Demand-led upcycle. Peaks at a 26.0% industry margin in FY2030, "
                         "matching the FY2022 cycle peak.",
            "Bear Case": "Demand-led downturn. Raw materials FALL with demand, so margins "
                         "compress to 7-11% rather than collapsing.",
            "Stress Case": "Stagflationary supply shock plus demand shock. Coking coal to "
                           "US$300/t, INR to 113. Troughs at -9.2% in FY2028."}
    for i, s in enumerate(SC):
        put(ws, r, CB, s, alignment=AL_CENTRE, f=font(10))
        put(ws, r, 2, "  %d." % (i + 1), f=font(9, colour=TEXT_MUTED))
        put(ws, r, CNOTE, desc[s], f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP); r += 1
    K(ws, "scen_first", list_first)
    r += 1
    put(ws, r, 2, "LIVE MODEL OUTPUT - ACTIVE SCENARIO", f=font(11, bold=True, colour=NAVY)); r += 1
    r = hdr(ws, r)
    out = [("India finished steel consumption", "Mt", "04 Steel Demand Model", "cons", FMT_1DP),
           ("India crude steel production", "Mt", "05 Steel Supply Model", "crude", FMT_1DP),
           ("India crude steel capacity", "Mtpa", "06 Capacity Forecast", "cap", FMT_1DP),
           ("Capacity utilisation", "%", "08 Capacity Utilisation", "util", FMT_PCT_1),
           ("Blended realisation", "Rs/t", "09 Steel Price Forecast", "realn", FMT_INT),
           ("Cash cost per tonne", "Rs/t", "11 Cost Curve", "cost", FMT_INT),
           ("EBITDA per tonne", "Rs/t", "13 EBITDA Model", "ebt", FMT_INT),
           ("Industry EBITDA margin", "%", "14 Margin Analysis", "margin", FMT_PCT_1),
           ("Industry revenue", "Rs cr", "12 Revenue Forecast", "rev", FMT_INT),
           ("Industry EBITDA", "Rs cr", "13 EBITDA Model", "ebitda", FMT_INT),
           ("Unlevered free cash flow", "Rs cr", "16 Cash Flow Model", "fcf", FMT_INT)]
    for lbl, u, sheet, key, fmt in out:
        bl = "=IF(ISBLANK('%s'!%s$%d),\"n/a\",'%s'!%s$%d)" % (
            sheet, GL(CB), R(sheet, key), sheet, GL(CB), R(sheet, key))
        r = line(ws, r, lbl, u, bl,
                 ["='%s'!%s$%d" % (sheet, GL(c), R(sheet, key)) for c in FCOLS],
                 src=sheet.split()[0], conf="", note="", fmt=fmt, bold=True, cagr=True)
    footnote(ws, r, 2, "Every figure above is a live link. Change the selector and this panel, "
                       "and every other sheet, recomputes.")
    return sel_row, id_row


# ======================================================================================
def sh_assump(ws):
    r = meta(ws, "02", "Model Assumptions",
             "Every driver in the model, on all four scenarios, with the full evidence base. "
             "Section B resolves the active scenario and is the ONLY block the rest of the "
             "workbook reads.",
             "Section A holds the raw scenario matrix: four rows per driver, one per scenario, "
             "in a fixed order. Section B holds one row per driver whose forecast cells are "
             "=INDEX(<the four scenario cells in this column>, ScenID). Downstream sheets "
             "reference Section B only. Section C is the assumption register carrying the "
             "mandated thirteen fields. Section D shows the WACC build-up cell by cell rather "
             "than as a single number, so a reviewer can substitute house assumptions in one "
             "place.",
             "As stated per row. Percentages are stored as decimals.",
             "Reads ScenID from 01 Control Panel. Read by every sheet from 03 to 24.", 110)

    r = sect(ws, r, "SECTION A - SCENARIO DRIVER MATRIX (raw input; not read downstream)")
    hs = ["ID", "Driver", "Unit", "Scenario"] + FY + ["", "Source", "Conf.", "Evidence and reasoning"]
    r = header_row(ws, r, hs, left_align_cols=(1, 2, 4, 13, CNOTE))
    a_first = r
    for code, name, unit, mat, src, conf, ev in _ORDER:
        base_row = r
        for i, s in enumerate(SC):
            put(ws, r, 1, code if i == 0 else "", f=font(9, colour=TEXT_MUTED),
                alignment=AL_CENTRE)
            put(ws, r, 2, name if i == 0 else "", alignment=AL_LEFT_WRAP,
                f=font(10, bold=(i == 0)))
            put(ws, r, 3, unit if i == 0 else "", f=font(9, colour=TEXT_MUTED))
            put(ws, r, 4, s, f=font(9), alignment=AL_LEFT)
            fmt = (FMT_PCT_1 if unit == "%" else FMT_MULT if unit == "x"
                   else FMT_INT if unit in ("Rs/t", "Rs/US$", "Rs/t of capacity added") else FMT_2DP)
            if unit == "Rs/US$":
                fmt = FMT_2DP
            for j, v in enumerate(mat[s]):
                put(ws, r, CF1 + j, v, number_format=fmt, alignment=AL_RIGHT)
            if i == 0:
                put(ws, r, CSRC, src, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
                put(ws, r, CCONF, conf, f=font(9, colour=TEXT_MUTED), alignment=AL_CENTRE)
                put(ws, r, CNOTE, ev, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
            r += 1
        K(ws, "mat_" + code, base_row)
    add_table(ws, "tbl_DriverMatrix", a_first - 1, 1, r - 1, CNOTE)
    r += 1

    r = sect(ws, r, "SECTION B - ACTIVE DRIVERS (resolved by ScenID; THIS is what the model reads)")
    r = hdr(ws, r, extra_left=("Source", "Conf.", "Notes"))
    b_first = r
    for code, name, unit, mat, src, conf, ev in _ORDER:
        mr = R(ws.title, "mat_" + code)
        fmt = (FMT_PCT_1 if unit == "%" else FMT_MULT if unit == "x"
               else FMT_INT if unit in ("Rs/t", "Rs/t of capacity added") else FMT_2DP)
        cells = ["=INDEX(%s%d:%s%d,ScenID)" % (GL(c), mr, GL(c), mr + 3) for c in FCOLS]
        r = line(ws, r, name, unit, None, cells, src=src, conf=conf,
                 note="Resolves Section A rows %d-%d by ScenID." % (mr, mr + 3), fmt=fmt,
                 key="act_" + code, ident=code)
    add_table(ws, "tbl_ActiveDrivers", b_first - 1, 1, r - 1, CNOTE)
    r += 1

    r = sect(ws, r, "SECTION C - ASSUMPTION REGISTER (thirteen mandated fields)")
    set_widths(ws, [(CNOTE, 110)])
    rh = A.REGISTER_HEADERS
    r2 = header_row(ws, r, rh, left_align_cols=tuple(range(1, len(rh) + 1)))
    c_first = r2
    for row in A.REGISTER:
        for i, v in enumerate(row):
            put(ws, r2, 1 + i, v, f=font(9, colour=TEXT_MUTED if i > 3 else None),
                alignment=AL_LEFT_WRAP)
        r2 += 1
    add_table(ws, "tbl_Register", c_first - 1, 1, r2 - 1, len(rh))
    r = r2 + 1

    r = sect(ws, r, "SECTION D - WACC BUILD-UP (every component shown and editable)")
    hs = ["", "Component", "Unit", "Value", "Source", "Publication date", "Confidence",
          "Evidence and caveats"]
    r = header_row(ws, r, hs, left_align_cols=(2, 5, 6, 8))
    w_first = r
    wrow = {}
    for nm, unit, val, src, pub, conf, ev in A.WACC_BUILD:
        put(ws, r, 2, nm, alignment=AL_LEFT_WRAP, f=font(10, bold=(val is None)))
        put(ws, r, 3, unit, f=font(9, colour=TEXT_MUTED))
        wrow[nm] = r
        if val is not None:
            put(ws, r, 4, val, number_format=FMT_MULT if unit == "x" else FMT_PCT_1,
                alignment=AL_RIGHT, fill=FILL_NAVY_LIGHT if conf == "Low" else None)
        put(ws, r, 5, src, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 6, pub, f=font(9, colour=TEXT_MUTED), alignment=AL_CENTRE)
        put(ws, r, 7, conf, f=font(9, colour=TEXT_MUTED), alignment=AL_CENTRE)
        put(ws, r, 8, ev, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    for nm, f in [("Cost of equity", "=D%d+D%d*D%d" % (wrow["Risk-free rate (10-year G-sec)"],
                                                       wrow["Equity beta (sector, levered)"],
                                                       wrow["Equity risk premium"])),
                  ("Post-tax cost of debt", "=D%d*(1-D%d)" % (wrow["Pre-tax cost of debt"],
                                                              wrow["Effective tax rate"])),
                  ("Target equity weight", "=1-D%d" % wrow["Target debt weight"]),
                  ("WACC (nominal, INR)", "=D%d*D%d+D%d*D%d" % (
                      wrow["Target equity weight"], wrow["Cost of equity"],
                      wrow["Target debt weight"], wrow["Post-tax cost of debt"]))]:
        put(ws, wrow[nm], 4, f, number_format=FMT_PCT_1, alignment=AL_RIGHT,
            f=font(10, bold=True, colour=NAVY))
    K(ws, "wacc_calc", wrow["WACC (nominal, INR)"])
    add_table(ws, "tbl_WACC", w_first - 1, 2, r - 1, 8)
    footnote(ws, r + 1, 2,
             "Shaded value cells are UNSOURCED modeller assumptions and are the weakest inputs "
             "in the model. The computed WACC here is a cross-check on driver D20, which is the "
             "value actually used downstream; 25 Audit Checks compares the two.")



# ======================================================================================
def P(c):
    """Previous column letter for a forecast column (base year for the first)."""
    return GL(c - 1)


def sh_macro(ws):
    r = meta(ws, "03", "Macroeconomic Model",
             "The macro spine: real GDP, inflation, the exchange rate and interest rates. Every "
             "demand and cost forecast in the workbook is anchored here.",
             "Real GDP growth and CPI are taken from RBI's June-2026 Monetary Policy Committee "
             "projections for FY2027 and converge to a steady state thereafter, because no "
             "Indian official body publishes a real GDP path beyond roughly five years. USD/INR "
             "is anchored on the 22-Jun-2026 spot of 94. Index rows compound the growth rates so "
             "that cumulative effects are visible rather than implied.",
             "Percentages as decimals. Indices FY2026A = 100. USD/INR in rupees per dollar.",
             "Reads 02 Model Assumptions. Read by 04 Steel Demand, 10 Raw Material, 16 Cash Flow.")
    r = hdr(ws, r)
    r = sect(ws, r, "GROWTH AND PRICES")
    r = line(ws, r, "India real GDP growth", "%", 0.076, [dv("D01", c) for c in FCOLS],
             "RBI MPC 05-Jun-2026", "High",
             "FY2026A of 7.6% is the RBI's own estimate for FY2026. FY2027E of 6.6% is the RBI "
             "projection, revised down from 6.9% on West Asia conflict, energy prices and "
             "monsoon risk. Beyond FY2027 the path converges to 6.5%.",
             FMT_PCT_1, key="gdp")
    r = line(ws, r, "Real GDP index", "FY2026A=100", 100.0,
             ["=%s%d*(1+%s%d)" % (P(c), r, GL(c), R(ws.title, "gdp")) for c in FCOLS],
             "Computed", "High", "Compounds the growth rate above.", FMT_1DP, key="gdpidx",
             cagr=True)
    r = line(ws, r, "CPI inflation", "%", 0.051, [dv("D04", c) for c in FCOLS],
             "RBI MPC 05-Jun-2026", "High",
             "RBI raised its FY2027 CPI projection by 50bps to 5.1%. Converges to the 4.5% "
             "mid-point of the 4% plus or minus 2% target band.", FMT_PCT_1, key="cpi")
    r = line(ws, r, "Nominal GDP index", "FY2026A=100", 100.0,
             ["=%s%d*(1+%s%d)*(1+%s%d)" % (P(c), r, GL(c), R(ws.title, "gdp"), GL(c),
                                           R(ws.title, "cpi")) for c in FCOLS],
             "Computed", "High", "Real growth compounded with inflation.", FMT_1DP,
             key="ngdpidx", cagr=True)
    r += 1
    r = sect(ws, r, "EXCHANGE RATE AND INTEREST RATES")
    r = line(ws, r, "USD/INR (average)", "Rs/US$", C.INR26, [dv("D03", c) for c in FCOLS],
             "IDBI Capital 22-Jun-2026", "High",
             "FY2026A of 88.4 is DERIVED from Tata Steel's dual-currency disclosure in the "
             "database - consolidated EBITDA/t of Rs 10,900 against US$124 gives 87.90, India "
             "EBITDA/t of Rs 15,213 against US$172 gives 88.45, and group turnover of "
             "Rs 2,32,140 cr against US$26bn gives 89.28. Spot was 94 on 22-Jun-2026 with a "
             "52-week range of 85 to 98.", FMT_2DP, key="inr", cagr=True)
    r = line(ws, r, "INR depreciation", "%", None,
             ["=%s%d/%s%d-1" % (GL(c), R(ws.title, "inr"), P(c), R(ws.title, "inr"))
              for c in FCOLS], "Computed", "High",
             "Positive equals depreciation, which raises rupee raw-material cost but also raises "
             "the rupee landed cost of imports and therefore supports domestic realisations.",
             FMT_PCT_1, key="inrdep")
    wr = R("02 Model Assumptions", "wacc_calc")
    r = line(ws, r, "10-year G-sec yield", "%", 0.0683,
             ["='02 Model Assumptions'!$D$%d" % (wr - 9)] * NY,
             "investing.com 31-Jul-2026", "High",
             "6.833% on 31-Jul-2026. Trading Economics expects 6.78% at quarter end. Held flat "
             "across the horizon - a term-structure forecast is not attempted.", FMT_PCT_1,
             key="gsec")
    r = line(ws, r, "WACC (nominal, INR)", "%", None, [dv("D20", c) for c in FCOLS],
             "Composite - see 02 Section D", "Medium",
             "Used for discounting on 23 Comparable Valuation. Cross-checked against the "
             "build-up on 02 Section D by 25 Audit Checks.", FMT_PCT_1, key="wacc")
    footnote(ws, r + 1, 2, "Macro drivers are scenario-dependent: the stress case applies a "
                           "220bps FY2027 GDP shock and takes USD/INR beyond its 52-week low.")


def sh_demand(ws):
    r = meta(ws, "04", "Steel Demand Model",
             "Forecasts India's apparent finished steel consumption, the demand anchor for the "
             "whole model.",
             "Consumption growth equals real GDP growth multiplied by a steel demand elasticity. "
             "The elasticity is not assumed - it is calibrated to realised history: consumption "
             "compounded at 7.13% from 76.99 Mt in FY2015 to 164.19 Mt in FY2026 against real "
             "GDP CAGR of roughly 6.3-6.7%, giving 1.06x to 1.13x. It is set at 1.15x near term, "
             "consistent with worldsteel's Apr-2026 forecast of +7.4% in CY2026 and +9.2% in "
             "CY2027 against RBI GDP of 6.6%, then tapers to 1.05x on maturing intensity.",
             "Mt. Population in billions. Per capita in kg.",
             "Reads 03 Macroeconomic Model and 02 Model Assumptions. Read by 05 Steel Supply.")
    r = hdr(ws, r)
    r = line(ws, r, "India real GDP growth", "%", 0.076,
             [cr_f("03 Macroeconomic Model", "gdp", c) for c in FCOLS], "03 Macro", "High",
             "Linked, not re-entered.", FMT_PCT_1, key="gdp")
    r = line(ws, r, "Steel demand elasticity to GDP", "x", 1.10, [dv("D02", c) for c in FCOLS],
             "Derived from Master DB", "High",
             "FY2026A of 1.10x is the realised FY2015-FY2026 average, the midpoint of the "
             "1.06x-1.13x range depending on the GDP deflator used.", FMT_MULT, key="elas")
    r = line(ws, r, "Implied consumption growth", "%", 0.079,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "gdp"), GL(c), R(ws.title, "elas"))
              for c in FCOLS], "Computed", "High",
             "FY2026A of 7.9% is the actual reported growth, against 7.6% GDP times 1.04x "
             "realised elasticity - the model's structure reproduces the actual outturn.",
             FMT_PCT_1, key="dgrowth")
    r = line(ws, r, "Apparent finished steel consumption", "Mt", C.CONS26,
             ["=%s%d*(1+%s%d)" % (P(c), r, GL(c), R(ws.title, "dgrowth")) for c in FCOLS],
             "00 Database Import A03", "High",
             "FY2026A of 163.74 Mt is the April-2026 JPC vintage, which is the internally "
             "self-reconciling balance. The June-2026 vintage of 164.19 Mt is a memo below.",
             FMT_2DP, bold=True, key="cons", cagr=True)
    r = line(ws, r, "  Memo: June-2026 JPC vintage, FY2026 actual", "Mt", 164.19, [],
             "Master DB", "High",
             "Carried so the vintage difference is visible rather than buried - Master DB "
             "Conflict C03. The 0.45 Mt difference is 0.27%.", FMT_2DP)
    r += 1
    r = sect(ws, r, "DEMAND INTENSITY")
    r = line(ws, r, "Population", "bn", C.POP26,
             ["=%s%d*(1+$D$%d)" % (P(c), r, r + 1) for c in FCOLS], "Derived", "Medium",
             "DERIVED from the database: consumption of 163.74 Mt divided by per capita "
             "consumption of 115.7 kg gives 1.4152 bn; the 1.4191 bn shown uses the June-vintage "
             "consumption of 164.19 Mt against 115.7 kg, which is the internally consistent "
             "pairing the database publishes.", FMT_3DP, key="pop")
    r = line(ws, r, "Population growth", "%", C.POPG, [], "NOT SOURCED", "Low",
             "0.85% per annum is NOT sourced in this research cycle. Replace with a UN World "
             "Population Prospects series. Affects per capita consumption only, not tonnage, so "
             "it does not propagate into revenue or EBITDA.", FMT_PCT_1, key="popg")
    r = line(ws, r, "Per capita finished steel consumption", "kg", 115.7,
             ["=%s%d/%s%d" % (GL(c), R(ws.title, "cons"), GL(c), R(ws.title, "pop"))
              for c in FCOLS], "Computed", "Medium",
             "Consumption in Mt divided by population in bn gives kg per head directly. India "
             "is at roughly half the world average of 215 kg and about a fifth of China's "
             "604 kg.", FMT_1DP, key="percap", cagr=True)
    footnote(ws, r + 1, 2, "Demand is the ONLY block in this model that is not capacity-aware. "
                           "The capacity ceiling is applied on 05 Steel Supply, so consumption "
                           "here is unconstrained demand.")


def cr_f(sheet, key, col):
    return "='%s'!%s$%d" % (sheet, GL(col), R(sheet, key))


def sh_tracker(ws):
    r = meta(ws, "07", "Capacity Expansion Tracker",
             "The dated project pipeline that drives the capacity forecast. Twenty projects, "
             "each with its capacity, status, expected commissioning year and whether it counts "
             "towards INDIA's national crude steel capacity.",
             "Every project is taken from the Master Database. The 'Included in forecast' flag is "
             "the critical field and encodes four exclusions that are routinely got wrong: "
             "already-commissioned projects are in the FY2026 base and must not be added again; "
             "ACQUISITIONS add to a company but nothing to national capacity; DOWNSTREAM lines "
             "(HRAP, CRAP, cold rolling) are processing capacity, not melt capacity; and OVERSEAS "
             "assets are outside India. Ramp-ups are utilisation, not capacity. 06 Capacity "
             "Forecast reads this sheet by SUMIFS on the flag and the year, so changing a flag or "
             "a date here immediately changes the forecast.",
             "Mtpa. Capex in Rs crore.",
             "Source: Master Industry Database. Read by 06 Capacity Forecast and 12 Revenue.", 105)
    set_widths(ws, [(1, 8), (2, 19), (3, 40), (4, 20), (5, 22), (6, 12), (7, 20), (8, 30),
                    (9, 15), (10, 12), (11, 13), (12, 34), (13, 10), (14, 10), (15, 105)])
    r = header_row(ws, r, A.PROJECT_HEADERS, left_align_cols=(1, 2, 3, 4, 5, 7, 8, 12, 15))
    first = r
    for p in A.PROJECTS:
        for i, v in enumerate(p):
            if v is None:
                continue
            if i == 10:
                continue
            put(ws, r, 1 + i, v,
                number_format=FMT_2DP if i == 5 else FMT_INT if i == 9 else FMT_TEXT,
                alignment=AL_RIGHT if i in (5, 9) else AL_LEFT_WRAP,
                f=font(9, colour=TEXT_MUTED) if i in (14,) else font())
        if p[5] and p[9]:
            put(ws, r, 11, "=J%d/F%d*10000" % (r, r), number_format=FMT_INT, alignment=AL_RIGHT)
        r += 1
    last = r - 1
    add_table(ws, "tbl_Projects", first - 1, 1, last, 15)
    freeze(ws, "C%d" % first)
    K(ws, "proj_first", first); K(ws, "proj_last", last)
    r += 1
    r = sect(ws, r, "SUMMARY - GROSS ANNOUNCED INDIA CRUDE STEEL ADDITIONS BY YEAR (live SUMIFS)")
    put(ws, r, 2, "Included projects only", f=font(9, colour=TEXT_MUTED))
    put(ws, r, 3, "Mtpa", f=font(9, colour=TEXT_MUTED))
    for i, c in enumerate(FCOLS):
        put(ws, r, c, '=SUMIFS($F$%d:$F$%d,$I$%d:$I$%d,"%s",$L$%d:$L$%d,"Yes")'
            % (first, last, first, last, FY[i].replace("E", ""), first, last),
            number_format=FMT_2DP, alignment=AL_RIGHT, f=font(10, bold=True))
    K(ws, "add_gross", r)
    put(ws, r, CNOTE, "Live SUMIFS across the table above. Matches on the commissioning year AND "
                      "on the inclusion flag being exactly 'Yes'.", f=font(9, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP)
    r += 2
    put(ws, r, 2, "Total gross announced additions, FY2027E-FY2033E", f=font(10, bold=True))
    put(ws, r, 3, "Mtpa", f=font(9, colour=TEXT_MUTED))
    put(ws, r, CB, "=SUM(%s%d:%s%d)" % (GL(CF1), R(ws.title, "add_gross"), GL(CFN),
                                        R(ws.title, "add_gross")),
        number_format=FMT_2DP, alignment=AL_RIGHT, f=font(10, bold=True, colour=NAVY))
    put(ws, r, CNOTE, "53.62 Mtpa of individually announced projects. Note this averages only "
                      "7.7 Mtpa a year against India's ACTUAL 20.8 and 20.1 Mtpa additions in "
                      "FY2025 and FY2026, which is precisely why driver D22 exists - see 06.",
        f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)


def sh_capacity(ws):
    r = meta(ws, "06", "Capacity Forecast",
             "Forecasts India's crude steel capacity from the dated project pipeline plus a "
             "secondary-sector residual, and validates the result against the National Steel "
             "Policy 2017 target.",
             "Capacity equals prior-year capacity plus announced additions delivered plus "
             "secondary and unattributed additions. Announced additions come live from the "
             "project tracker by SUMIFS and are multiplied by a delivery factor of 85% "
             "reflecting observed slippage. The secondary residual exists because the tracker "
             "captures only individually announced projects averaging 7.7 Mtpa a year, whereas "
             "India actually added 20.8 Mtpa in FY2025 and 20.1 Mtpa in FY2026 - the difference "
             "comes from the secondary sector, whose output the database shows growing at 10.7% "
             "a year, faster than the integrated majors' 8.1%.",
             "Mtpa.",
             "Reads 07 Capacity Expansion Tracker and 02 Model Assumptions. Read by 08 "
             "Utilisation, 05 Steel Supply, 17 Capital Allocation.")
    r = hdr(ws, r)
    r = line(ws, r, "Opening capacity", "Mtpa", None,
             ["=%s%d" % (P(c), r + 6) for c in FCOLS], "Computed", "High",
             "Prior-year closing capacity.", FMT_1DP, key="cap_open")
    r = line(ws, r, "Gross announced additions (from tracker)", "Mtpa", None,
             [cr_f("07 Capacity Expansion Tracker", "add_gross", c) for c in FCOLS],
             "07 Tracker (SUMIFS)", "Medium",
             "Live from the project tracker. Excludes acquisitions, downstream lines, overseas "
             "assets and already-commissioned projects.", FMT_2DP, key="add_gross")
    r = line(ws, r, "Pipeline delivery factor", "%", None, [dv("D05", c) for c in FCOLS],
             "Derived from Master DB", "Medium",
             "85% in the base case. India delivered ~20 Mtpa a year in FY2025 and FY2026 against "
             "a National Steel Policy requirement of ~16 Mtpa a year, so delivery capability is "
             "demonstrated; the 15% haircut reflects project-level slippage such as JSW "
             "Vijayanagar BF-3 still testing at FY2026 year-end.", FMT_PCT_1, key="deliv")
    r = line(ws, r, "Announced additions delivered", "Mtpa", None,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "add_gross"), GL(c), R(ws.title, "deliv"))
              for c in FCOLS], "Computed", "Medium", "", FMT_2DP, key="add_net")
    r = line(ws, r, "Secondary and unattributed additions", "Mtpa", None,
             [dv("D22", c) for c in FCOLS], "Derived from Master DB", "Medium",
             "ESSENTIAL RECONCILIATION ITEM. Omitting this would understate FY2031 capacity by "
             "about 54 Mtpa, or 18%.", FMT_2DP, key="add_sec")
    r = line(ws, r, "Total additions", "Mtpa", 20.07,
             ["=%s%d+%s%d" % (GL(c), R(ws.title, "add_net"), GL(c), R(ws.title, "add_sec"))
              for c in FCOLS], "Computed", "Medium",
             "FY2026A of 20.07 Mtpa is the actual year-on-year capacity increase from the "
             "database (220.4 less 200.33).", FMT_2DP, key="add_tot")
    r = line(ws, r, "Closing crude steel capacity", "Mtpa", C.CAP26,
             ["=%s%d+%s%d" % (P(c), r, GL(c), R(ws.title, "add_tot")) for c in FCOLS],
             "00 Database Import A04", "High",
             "FY2026A of 220.4 Mtpa is the database figure (June-2026 Ministry vintage; the "
             "March vintage said 220.3 - Conflict C04).", FMT_1DP, bold=True, key="cap",
             cagr=True)
    r += 1
    r = sect(ws, r, "VALIDATION AGAINST NATIONAL STEEL POLICY 2017")
    r = line(ws, r, "National Steel Policy 2017 target", "Mtpa", None,
             [None, None, None, None, 300.0, None, None], "Master DB policy P10", "High",
             "300 Mtpa of crude steel capacity by FY2031.", FMT_1DP, key="nsp")
    r = line(ws, r, "Model capacity in FY2031E", "Mtpa", None,
             [None, None, None, None, "=%s%d" % (GL(FCOLS[4]), R(ws.title, "cap")), None, None],
             "Computed", "High", "", FMT_1DP, key="nsp_model")
    r = line(ws, r, "Model as % of policy target", "%", None,
             [None, None, None, None,
              "=%s%d/%s%d" % (GL(FCOLS[4]), R(ws.title, "nsp_model"), GL(FCOLS[4]),
                              R(ws.title, "nsp")), None, None],
             "Computed", "High",
             "The model was built bottom-up from the project tracker plus the secondary residual "
             "and was NOT calibrated to the policy target, so agreement within roughly 1% is a "
             "genuine independent check on the capacity block.", FMT_PCT_1, bold=True,
             key="nsp_pct")
    footnote(ws, r + 1, 2, "Sign warning: more capacity is EBITDA-NEGATIVE in this model, because "
                           "it depresses utilisation which feeds back into price via driver D23. "
                           "See 09 Steel Price Forecast.")


def sh_supply(ws):
    r = meta(ws, "05", "Steel Supply Model",
             "Converts demand into required domestic production, applies a practical capacity "
             "ceiling, and routes any shortfall to imports so that the supply balance always "
             "closes.",
             "The identity is the Joint Plant Committee's own: consumption equals production plus "
             "imports less exports adjusted for the change in stocks. Rearranged, required "
             "finished production equals consumption plus net exports plus the stock change. "
             "Crude steel equals finished production multiplied by the crude-to-finished ratio of "
             "1.0465, which is the FY2026 actual on the April-2026 vintage. Production is then "
             "capped at capacity times maximum practical utilisation of 92%; where demand would "
             "exceed that, the shortfall is met by additional imports rather than by impossible "
             "domestic output.",
             "Mt. Ratio in times.",
             "Reads 04 Steel Demand, 06 Capacity Forecast, 02 Model Assumptions. Read by 09, 11, "
             "12, 13, 19.")
    r = hdr(ws, r)
    r = line(ws, r, "Apparent finished steel consumption", "Mt", C.CONS26,
             [cr_f("04 Steel Demand Model", "cons", c) for c in FCOLS], "04 Demand", "High",
             "", FMT_2DP, key="cons")
    r = line(ws, r, "Net finished steel exports", "Mt", 0.078, [dv("D07", c) for c in FCOLS],
             "Derived from Master DB", "Medium",
             "FY2026A of +0.078 Mt - India returned to net exporter status having been 4.693 Mt "
             "net importer in FY2025, a 4.77 Mt swing in one year.", FMT_2DP, key="netexp")
    r = line(ws, r, "Variation in stock", "Mt", -2.88, [dv("D08", c) for c in FCOLS],
             "Master DB supply balance", "Medium",
             "FY2026A of -2.88 Mt, i.e. a DESTOCK of about 1.8% of consumption which flattered "
             "reported consumption growth. Base normalises to a small restock.", FMT_2DP,
             key="stock")
    r = line(ws, r, "Required finished steel production", "Mt", C.PROD26,
             ["=%s%d+%s%d+%s%d" % (GL(c), R(ws.title, "cons"), GL(c), R(ws.title, "netexp"),
                                   GL(c), R(ws.title, "stock")) for c in FCOLS],
             "Computed - JPC identity", "High",
             "IDENTITY CHECK: 163.74 + 0.078 + (-2.88) = 160.94, which equals the reported "
             "FY2026 finished steel production exactly. The balance closes.", FMT_2DP,
             key="prod_req")
    r = line(ws, r, "Crude-to-finished steel ratio", "x", 1.04647, [dv("D06", c) for c in FCOLS],
             "Derived from Master DB", "High",
             "168.42 / 160.94 on the April-2026 vintage. On the June vintage it would be 1.0413, "
             "a 0.5% difference.", FMT_MULT, key="cfr")
    r = line(ws, r, "Required crude steel production (unconstrained)", "Mt", C.CRUDE26,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "prod_req"), GL(c), R(ws.title, "cfr"))
              for c in FCOLS], "Computed", "High", "", FMT_2DP, key="crude_uncon")
    r += 1
    r = sect(ws, r, "CAPACITY CEILING")
    r = line(ws, r, "Crude steel capacity", "Mtpa", C.CAP26,
             [cr_f("06 Capacity Forecast", "cap", c) for c in FCOLS], "06 Capacity", "High", "",
             FMT_1DP, key="cap")
    r = line(ws, r, "Maximum practical utilisation", "%", None, [dv("D25", c) for c in FCOLS],
             "Indicative", "Low",
             "92%. Allows for maintenance and relining. Individual assets can exceed nameplate - "
             "the database records SAIL Rourkela at 103.8% - but a national aggregate cannot.",
             FMT_PCT_1, key="utilmax")
    r = line(ws, r, "Maximum crude steel production", "Mt", None,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "cap"), GL(c), R(ws.title, "utilmax"))
              for c in FCOLS], "Computed", "Low", "", FMT_2DP, key="crude_max")
    r = line(ws, r, "Crude steel production", "Mt", C.CRUDE26,
             ["=MIN(%s%d,%s%d)" % (GL(c), R(ws.title, "crude_uncon"), GL(c),
                                   R(ws.title, "crude_max")) for c in FCOLS],
             "Computed", "High",
             "The binding constraint. In the base case the ceiling does NOT bind - peak "
             "utilisation is 84.0% in FY2033 - which is itself an important finding: India is "
             "not capacity-constrained on this demand path.", FMT_2DP, bold=True, key="crude",
             cagr=True)
    r = line(ws, r, "Capacity ceiling binding?", "flag", None,
             ['=IF(%s%d>%s%d+0.001,"BINDING","no")' % (GL(c), R(ws.title, "crude_uncon"),
                                                       GL(c), R(ws.title, "crude_max"))
              for c in FCOLS], "Computed", "High",
             "Flags any year where domestic demand cannot be met from domestic capacity.",
             FMT_TEXT, key="bind")
    r = line(ws, r, "Finished steel production", "Mt", C.PROD26,
             ["=%s%d/%s%d" % (GL(c), R(ws.title, "crude"), GL(c), R(ws.title, "cfr"))
              for c in FCOLS], "Computed", "High", "", FMT_2DP, bold=True, key="prod",
             cagr=True)
    r = line(ws, r, "Shortfall met by additional imports", "Mt", 0.0,
             ["=MAX(0,%s%d-%s%d)" % (GL(c), R(ws.title, "prod_req"), GL(c), R(ws.title, "prod"))
              for c in FCOLS], "Computed", "High",
             "Keeps the balance closed when the ceiling binds. Flows through to 19 Trade Model.",
             FMT_2DP, key="shortfall")
    footnote(ws, r + 1, 2, "The model does NOT endogenise supply response - it does not idle "
                           "capacity when margins turn negative. This is why the stress case "
                           "should be read as a lower bound on margin rather than a central "
                           "expectation, and is a documented limitation on 25 Audit Checks.")


def sh_util(ws):
    r = meta(ws, "08", "Capacity Utilisation",
             "The bridge between the capacity and supply blocks, and the input to the price "
             "feedback that makes this model integrated.",
             "Utilisation equals crude steel production divided by closing capacity. The "
             "utilisation GAP against normal utilisation of 80% is the quantity that drives the "
             "lagged price adjustment on 09 Steel Price Forecast. Normal utilisation of 80% is "
             "anchored on actual Indian utilisation from the database - 80.4% in FY2024, 76.0% "
             "in FY2025 and 76.4% in FY2026 - with FY2024's recent high adopted as normal.",
             "Percentages as decimals. Mt for headroom.",
             "Reads 05 Steel Supply and 06 Capacity Forecast. Read by 09 Steel Price and 18 Cycle.")
    r = hdr(ws, r)
    r = line(ws, r, "Crude steel production", "Mt", C.CRUDE26,
             [cr_f("05 Steel Supply Model", "crude", c) for c in FCOLS], "05 Supply", "High", "",
             FMT_2DP, key="crude")
    r = line(ws, r, "Crude steel capacity", "Mtpa", C.CAP26,
             [cr_f("06 Capacity Forecast", "cap", c) for c in FCOLS], "06 Capacity", "High", "",
             FMT_1DP, key="cap")
    r = line(ws, r, "Capacity utilisation", "%", C.CRUDE26 / C.CAP26,
             ["=%s%d/%s%d" % (GL(c), R(ws.title, "crude"), GL(c), R(ws.title, "cap"))
              for c in FCOLS], "Computed", "High",
             "FY2026A of 76.4%. Uses closing capacity as the denominator, so it understates true "
             "utilisation in heavy commissioning years - the same convention as the database.",
             FMT_PCT_1, bold=True, key="util")
    r = line(ws, r, "Normal utilisation", "%", None, [dv("D24", c) for c in FCOLS],
             "Derived from Master DB", "Medium",
             "80%, anchored on the FY2024 actual of 80.4%.", FMT_PCT_1, key="utilnorm")
    r = line(ws, r, "Utilisation gap vs normal", "ppt", None,
             ["=%s%d-%s%d" % (GL(c), R(ws.title, "util"), GL(c), R(ws.title, "utilnorm"))
              for c in FCOLS], "Computed", "High",
             "THE DRIVER OF THE PRICE FEEDBACK. Positive means tight capacity supporting price; "
             "negative means surplus capacity depressing it. Used with a one-year lag on 09.",
             FMT_PCT_1, key="utilgap")
    r = line(ws, r, "Spare capacity", "Mt", None,
             ["=%s%d-%s%d" % (GL(c), R(ws.title, "cap"), GL(c), R(ws.title, "crude"))
              for c in FCOLS], "Computed", "High",
             "Idle capacity. Large and persistent spare capacity is the structural reason the "
             "base case forecasts realisation growing below inflation.", FMT_1DP, key="spare")
    footnote(ws, r + 1, 2, "FY2026 utilisation of 76.4% against 80.4% in FY2024 is the single "
                           "most important structural fact in the Indian steel investment case: "
                           "capacity has been added faster than demand, which caps domestic "
                           "pricing power and is why a safeguard duty was needed.")



# ======================================================================================
def sh_price(ws):
    r = meta(ws, "09", "Steel Price Forecast",
             "Forecasts blended industry realisation, including the lagged utilisation feedback "
             "that connects the capacity block to earnings.",
             "The base-year anchor is a DERIVED blended realisation of Rs 59,974/t: the "
             "volume-weighted revenue per tonne of Tata Steel India, JSW Steel India, SAIL and "
             "Jindal Steel from the database - Rs 4,79,191 cr over 79.90 Mt, covering 49.4% of "
             "India's finished steel production. A blended realisation is used rather than a spot "
             "HRC quote for two reasons: the database documents three irreconcilable Mar-2026 "
             "HRC assessments spanning Rs 55,900 to Rs 59,500/t (Conflict C07), and revenue is "
             "driven by actual product mix, not by HRC alone. Effective realisation then equals "
             "the scenario driver multiplied by (1 + elasticity x prior-year utilisation gap).",
             "Rs per tonne. Percentages as decimals.",
             "Reads 02 Model Assumptions and 08 Capacity Utilisation. Read by 12 Revenue, 13 "
             "EBITDA, 14 Margin.")
    r = hdr(ws, r)
    r = sect(ws, r, "DRIVER REALISATION (before the utilisation feedback)")
    r = line(ws, r, "Blended realisation - scenario driver", "Rs/t", C.REALN26,
             [dv("D09", c) for c in FCOLS], "Derived from Master DB", "Medium",
             "FY2026A of Rs 59,974/t is derived, not quoted. Component realisations: Tata Steel "
             "India Rs 62,273/t, JSW Steel India Rs 60,798/t, Jindal Steel Rs 61,319/t, SAIL "
             "Rs 55,600/t. Jindal Stainless is EXCLUDED at roughly Rs 1,67,000/t because "
             "stainless is not comparable.", FMT_INT, key="realn_drv")
    r += 1
    r = sect(ws, r, "UTILISATION FEEDBACK (lagged one year - no circular reference)")
    ul = R("08 Capacity Utilisation", "util")
    r = line(ws, r, "Prior-year capacity utilisation", "%", None,
             ["='08 Capacity Utilisation'!%s$%d" % (P(c), ul) for c in FCOLS],
             "08 Utilisation (lagged)", "High",
             "Deliberately the PRIOR year. FY2027E uses the FY2026A actual of 76.4%. The lag is "
             "what allows capacity to affect price without creating a circular reference.",
             FMT_PCT_1, key="util_lag")
    r = line(ws, r, "Normal utilisation", "%", None, [dv("D24", c) for c in FCOLS],
             "Derived from Master DB", "Medium", "", FMT_PCT_1, key="utilnorm")
    r = line(ws, r, "Price sensitivity to utilisation", "x", None, [dv("D23", c) for c in FCOLS],
             "Indicative", "Low",
             "0.40x. INDICATIVE and unsourced, held constant across scenarios for that reason, "
             "and given a sensitivity axis on 22.", FMT_MULT, key="pxutil")
    r = line(ws, r, "Price adjustment factor", "x", 1.0,
             ["=1+%s%d*(%s%d-%s%d)" % (GL(c), R(ws.title, "pxutil"), GL(c),
                                       R(ws.title, "util_lag"), GL(c), R(ws.title, "utilnorm"))
              for c in FCOLS], "Computed", "Low",
             "Below 1.0 means surplus capacity is depressing realisation; above 1.0 means tight "
             "capacity is supporting it.", FMT_MULT, key="px_adj")
    r = line(ws, r, "Blended realisation - effective", "Rs/t", C.REALN26,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "realn_drv"), GL(c), R(ws.title, "px_adj"))
              for c in FCOLS], "Computed", "Medium",
             "THE REVENUE DRIVER. This, not the raw driver above, is what 12 Revenue reads.",
             FMT_INT, bold=True, key="realn", cagr=True)
    r += 1
    r = sect(ws, r, "REFERENCE PRICES AND POLICY")
    r = line(ws, r, "Memo: domestic HRC, end-Mar-2026 (ICRA)", "Rs/t", 57700, [],
             "Master DB S19", "High",
             "The database's recommended HRC reference. Shown as a memo only - the model runs on "
             "blended realisation. IDBI put Mar-2026 at Rs 59,500/t and ETInfra at Rs 55,900/t "
             "(Conflict C07).", FMT_INT)
    r = line(ws, r, "Memo: HRC 52-week range to Jun-2026", "Rs/t", None,
             ["Rs 45,700 - 59,600/t"] + [None] * 6, "Master DB S21", "High",
             "A ~30% range. The single largest source of FY2027 earnings uncertainty.", FMT_TEXT)
    r = line(ws, r, "Safeguard duty on flat products", "%", 0.12,
             [0.115, 0.11, 0.0, 0.0, 0.0, 0.0, 0.0], "Master DB policy P03", "High",
             "SCHEDULE, NOT A FORECAST. 12% for 21-Apr-2025 to 20-Apr-2026 (FY2026), 11.5% from "
             "21-Apr-2026 (FY2027), 11% from 21-Apr-2027 (FY2028), expiring 20-Apr-2028. Zero "
             "thereafter unless extended - a mid-term review is provided for. The FY2029 step to "
             "zero is why the base case realisation path dips in FY2028-FY2029.", FMT_PCT_1,
             key="sgd")
    footnote(ws, r + 1, 2, "The duty covers FLAT products only, which is why the database records "
                           "primary rebar rising 25% against HRC's 18% between the Dec-2025 "
                           "quarter and Mar-2026.")


def sh_rawmat(ws):
    r = meta(ws, "10", "Raw Material Forecast",
             "Forecasts the input basket - iron ore, coking coal and the exchange rate - and "
             "builds the raw-material cost index that drives cash cost.",
             "Dollar prices are converted to rupees at the forecast exchange rate, then indexed "
             "to FY2026. The basket weights iron ore and coking coal at 45% each with 10% other, "
             "which is INDICATIVE and unsourced. The resulting index, together with a conversion "
             "cost index compounding inflation, drives cash cost on 11 Cost Curve.",
             "US dollars per tonne as quoted; rupees per tonne after conversion. Indices "
             "FY2026A = 1.00.",
             "Reads 02 Model Assumptions and 03 Macroeconomic Model. Read by 11 Cost Curve.")
    r = hdr(ws, r)
    r = sect(ws, r, "DOLLAR PRICES")
    r = line(ws, r, "Iron ore 62% Fe fines, CFR China", "US$/dmt", C.IO26,
             [dv("D10", c) for c in FCOLS], "Master DB S17 (World Bank)", "High",
             "FY2026A of US$100.52/dmt is a genuine 12-month average from the World Bank Pink "
             "Sheet - the highest-quality price series in the model. Third consecutive annual "
             "decline from US$155.52/dmt in FY2022. Q1FY2027 actual averaged US$105.17/dmt.",
             FMT_2DP, key="io")
    r = line(ws, r, "Premium HCC coking coal, FOB Australia", "US$/t", C.CC26,
             [dv("D11", c) for c in FCOLS], "Master DB S01 (Ministry of Steel)", "Medium",
             "FY2026A of US$225/t is the Ministry of Steel's MAR-2026 SPOT observation, NOT a "
             "fiscal-year average - the database deliberately asserts no average because the two "
             "public series disagree by up to US$40/t (Conflict C01). THE LARGEST COST "
             "UNCERTAINTY IN THE MODEL. Replace with a licensed PRA series before transaction "
             "use.", FMT_2DP, key="cc")
    r = line(ws, r, "USD/INR", "Rs/US$", C.INR26,
             [cr_f("03 Macroeconomic Model", "inr", c) for c in FCOLS], "03 Macro", "High", "",
             FMT_2DP, key="inr")
    r += 1
    r = sect(ws, r, "RUPEE-CONVERTED PRICES")
    r = line(ws, r, "Iron ore in rupees", "Rs/dmt", C.IO26 * C.INR26,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "io"), GL(c), R(ws.title, "inr"))
              for c in FCOLS], "Computed", "High",
             "This is where the FX channel enters cost: a weaker rupee raises input cost even if "
             "the dollar price is flat.", FMT_INT, key="io_inr", cagr=True)
    r = line(ws, r, "Coking coal in rupees", "Rs/t", C.CC26 * C.INR26,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "cc"), GL(c), R(ws.title, "inr"))
              for c in FCOLS], "Computed", "Medium", "", FMT_INT, key="cc_inr", cagr=True)
    r += 1
    r = sect(ws, r, "COST INDICES")
    r = line(ws, r, "Conversion cost inflation", "%", None, [dv("D12", c) for c in FCOLS],
             "RBI-anchored", "Medium",
             "Applies to labour, power, consumables, stores, repairs and logistics. Set below "
             "RBI's 5.1% FY2027 CPI because conversion cost tracks energy rather than headline "
             "CPI and because the majors are delivering cost programmes - Tata Steel recorded "
             "roughly Rs 10,868 cr of benefit in FY2026, about Rs 4,825 per tonne of India "
             "deliveries.", FMT_PCT_1, key="convinf")
    r = line(ws, r, "Conversion cost index", "x", 1.0,
             ["=%s%d*(1+%s%d)" % (P(c), r, GL(c), R(ws.title, "convinf")) for c in FCOLS],
             "Computed", "Medium", "Compounds the inflation rate above.", FMT_3DP, key="convidx")
    r = line(ws, r, "Iron ore weight in basket", "%", None, [dv("D14", c) for c in FCOLS],
             "Indicative", "Low",
             "45% iron ore, 45% coking coal, 10% other ferrous and fluxes. INDICATIVE and "
             "unsourced; held constant across scenarios and given a sensitivity axis on 22.",
             FMT_PCT_1, key="iow")
    r = line(ws, r, "Raw material cost index", "x", 1.0,
             ["=%s%d*(%s%d/$D$%d)+%s%d*(%s%d/$D$%d)+(1-2*%s%d)*%s%d"
              % (GL(c), R(ws.title, "iow"), GL(c), R(ws.title, "io_inr"), R(ws.title, "io_inr"),
                 GL(c), R(ws.title, "iow"), GL(c), R(ws.title, "cc_inr"), R(ws.title, "cc_inr"),
                 GL(c), R(ws.title, "iow"), GL(c), R(ws.title, "convidx")) for c in FCOLS],
             "Computed", "Low",
             "Weighted index of rupee input prices relative to FY2026, with the residual 10% "
             "indexed to conversion inflation.", FMT_3DP, bold=True, key="rmidx")
    footnote(ws, r + 1, 2, "The Australian THERMAL coal series in the database (FY2026 average "
                           "US$111.50/mt) is deliberately NOT used here. It is thermal, not "
                           "metallurgical, coal - see database Conflict C15.")


def sh_cost(ws):
    r = meta(ws, "11", "Cost Curve",
             "Forecasts industry cash cost per tonne and positions each covered producer on the "
             "cost curve.",
             "CALIBRATED TOP-DOWN, NOT BUILT BOTTOM-UP. The FY2026 anchor of Rs 49,241/t is "
             "derived as blended realisation of Rs 59,974/t less weighted EBITDA of Rs 10,733/t, "
             "both from the database. It is forecast by splitting cash cost into a raw-material "
             "component indexed to the input basket and a conversion component indexed to "
             "inflation. A bottom-up build with consumption coefficients - tonnes of ore and coal "
             "per tonne of steel - was deliberately NOT attempted because those coefficients "
             "could not be sourced from any primary document, and inventing them would create "
             "false precision in the single most important cost line in the model.",
             "Rs per tonne.",
             "Reads 09 Steel Price and 10 Raw Material. Read by 13 EBITDA and 14 Margin.")
    r = hdr(ws, r)
    r = line(ws, r, "Raw material share of cash cost", "%", None, [dv("D13", c) for c in FCOLS],
             "Indicative", "Low",
             "60/40 raw material to conversion. INDICATIVE and unsourced. Plausible range 55-65%; "
             "a 5ppt change alters FY2033 EBITDA/t by roughly Rs 600/t.", FMT_PCT_1,
             key="rmshare")
    r = line(ws, r, "Raw material cost index", "x", 1.0,
             [cr_f("10 Raw Material Forecast", "rmidx", c) for c in FCOLS], "10 Raw Material",
             "Low", "", FMT_3DP, key="rmidx")
    r = line(ws, r, "Conversion cost index", "x", 1.0,
             [cr_f("10 Raw Material Forecast", "convidx", c) for c in FCOLS],
             "10 Raw Material", "Medium", "", FMT_3DP, key="convidx")
    r = line(ws, r, "Industry cash cost per tonne", "Rs/t", C.COST26,
             ["=$D$%d*(%s%d*%s%d+(1-%s%d)*%s%d)"
              % (r, GL(c), R(ws.title, "rmshare"), GL(c), R(ws.title, "rmidx"), GL(c),
                 R(ws.title, "rmshare"), GL(c), R(ws.title, "convidx")) for c in FCOLS],
             "Computed - calibrated", "Medium",
             "FY2026A of Rs 49,241/t is derived from the database. Dispersion across the four "
             "majors is only Rs 3,777/t, under 8%, which supports the weighted figure being "
             "representative rather than an artefact of one company.", FMT_INT, bold=True,
             key="cost", cagr=True)
    r += 1
    r = sect(ws, r, "PRODUCER COST CURVE - FY2026A ACTUAL AND FORECAST")
    put(ws, r, 2, "Each producer's FY2026 implied cash cost is realisation less reported "
                  "EBITDA/t. The forecast holds each producer's SPREAD to the industry constant, "
                  "which assumes relative competitive position is stable - a strong assumption, "
                  "stated explicitly.", f=font(9, italic=True, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP)
    ws.row_dimensions[r].height = 26
    r += 1
    r = hdr(ws, r)
    curve = [("Tata Steel", 62273, 15213), ("SAIL", 55600, 6596), ("JSW Steel", 60798, 10167),
             ("Jindal Steel", 61319, 10482)]
    curve = sorted(curve, key=lambda x: x[1] - x[2])
    for nm, rz, eb in curve:
        c26 = rz - eb
        spread = c26 - C.COST26
        r = line(ws, r, "%s - cash cost" % nm, "Rs/t", c26,
                 ["=%s%d+%s" % (GL(c), R(ws.title, "cost"), _n(spread)) for c in FCOLS],
                 "Master DB", "Medium",
                 "FY2026 realisation Rs %s/t less reported EBITDA Rs %s/t. Spread to industry "
                 "%+d Rs/t, held constant in the forecast."
                 % (format(rz, ","), format(eb, ","), spread), FMT_INT)
    footnote(ws, r + 1, 2, "Cost-curve position is the primary determinant of who survives a "
                           "downturn. On FY2026 actuals Tata Steel India sits roughly Rs 2,180/t "
                           "BELOW the industry weighted cash cost and Jindal Steel roughly "
                           "Rs 1,600/t above it - a Rs 3,780/t spread between best and worst of "
                           "the four, equal to about 35% of industry EBITDA per tonne.")


def _n(v):
    return ("(%d)" % v) if v < 0 else str(int(v))


def sh_revenue(ws):
    r = meta(ws, "12", "Revenue Forecast",
             "Industry revenue and a company-level revenue build for the covered producers.",
             "Industry revenue equals finished steel production multiplied by effective blended "
             "realisation. The unit bridge is explicit: Rs/t multiplied by Mt divided by 10 gives "
             "Rs crore. Company volumes grow from their FY2026 actual by their own tracked "
             "capacity additions multiplied by the industry utilisation rate, so a company only "
             "gets volume if the industry can absorb it. Company realisation moves with the "
             "industry index, preserving each producer's observed FY2026 mix premium or discount.",
             "Rs crore. Volumes in Mt.",
             "Reads 05 Steel Supply, 09 Steel Price, 07 Tracker. Read by 13, 14, 15, 16, 23.")
    r = hdr(ws, r)
    r = line(ws, r, "Finished steel production", "Mt", C.PROD26,
             [cr_f("05 Steel Supply Model", "prod", c) for c in FCOLS], "05 Supply", "High", "",
             FMT_2DP, key="prod")
    r = line(ws, r, "Blended realisation - effective", "Rs/t", C.REALN26,
             [cr_f("09 Steel Price Forecast", "realn", c) for c in FCOLS], "09 Price", "Medium",
             "", FMT_INT, key="realn")
    r = line(ws, r, "INDUSTRY REVENUE", "Rs cr", C.PROD26 * C.REALN26 / 10.0,
             ["=%s%d*%s%d/10" % (GL(c), R(ws.title, "prod"), GL(c), R(ws.title, "realn"))
              for c in FCOLS], "Computed", "Medium",
             "Unit bridge: Rs/t x Mt / 10 = Rs crore, because 1 Rs crore = 1e7 Rs and 1 Mt = 1e6 "
             "tonnes. FY2026A of Rs 9,65,238 cr is the whole-of-India finished steel revenue "
             "pool, against Rs 4,79,191 cr for the four majors that anchor realisation - i.e. "
             "the calibration sample is 49.6% of the pool by revenue.", FMT_INT, bold=True,
             key="rev", cagr=True)
    r += 2
    r = sect(ws, r, "COMPANY VOLUME BUILD")
    r = hdr(ws, r)
    ur = R("08 Capacity Utilisation", "util")
    for co in CO.COMPANIES:
        key, nm, vol = co[0], co[1], co[2]
        adds = CO.CO_ADDITIONS.get(key, {})
        cells = []
        for i, c in enumerate(FCOLS):
            inc = adds.get(FY[i].replace("E", ""), 0.0)
            prev = "%s%d" % (P(c), r)
            if inc:
                cells.append("=%s+%s*'08 Capacity Utilisation'!%s$%d" % (prev, inc, GL(c), ur))
            else:
                cells.append("=%s" % prev)
        r = line(ws, r, "%s - volume" % nm, "Mt", vol, cells, co[17], "Medium",
                 "%s. Basis: %s. %s" % (CO.CO_NOTES.get(key, ""), co[3],
                                        ("Tracked additions: " + ", ".join(
                                            "%s +%.2f Mtpa" % (k, v) for k, v in
                                            sorted(adds.items()))) if adds else
                                        "No tracked additions - volume growth comes from "
                                        "utilising capacity already built."),
                 FMT_2DP, key="vol_" + key, cagr=True)
    r += 1
    r = sect(ws, r, "COMPANY REVENUE BUILD")
    r = hdr(ws, r)
    for co in CO.COMPANIES:
        key, nm, vol, _, rev = co[0], co[1], co[2], co[3], co[4]
        if not rev or not vol:
            continue
        r26 = rev / vol * 10.0
        prem = r26 / C.REALN26
        cells = ["=%s%d*%s%d*%.6f/10" % (GL(c), R(ws.title, "vol_" + key), GL(c),
                                         R(ws.title, "realn"), prem) for c in FCOLS]
        r = line(ws, r, "%s - revenue" % nm, "Rs cr", rev, cells, co[17], "Medium",
                 "FY2026 realisation Rs %s/t = %.2fx the industry blend, held constant. Basis: "
                 "%s." % (format(int(r26), ","), prem, co[5]), FMT_INT, key="rev_" + key,
                 cagr=True)
    footnote(ws, r + 1, 2, "Company revenues are NOT additive to industry revenue: Jindal "
                           "Stainless is stainless, AM/NS India is on a 100% JV basis, RINL "
                           "turnover is on an unspecified basis, and the six together are only "
                           "part of India's output. Use industry revenue for the market and "
                           "company revenue for relative work.")


def sh_ebitda(ws):
    r = meta(ws, "13", "EBITDA Model",
             "Industry and company EBITDA, the primary valuation input.",
             "Industry EBITDA per tonne equals effective realisation less cash cost - a spread, "
             "not a margin applied to revenue, which is the correct construction for a commodity "
             "where price and cost move independently. Company EBITDA per tonne holds each "
             "producer's FY2026 observed spread to the industry constant. The FY2026 weighted "
             "anchor of Rs 10,733/t reconciles to the sum of reported EBITDAs within 0.6%.",
             "Rs per tonne and Rs crore.",
             "Reads 09 Price, 11 Cost, 12 Revenue. Read by 14, 16, 17, 21, 22, 23, 24.")
    r = hdr(ws, r)
    r = line(ws, r, "Blended realisation - effective", "Rs/t", C.REALN26,
             [cr_f("09 Steel Price Forecast", "realn", c) for c in FCOLS], "09 Price", "Medium",
             "", FMT_INT, key="realn")
    r = line(ws, r, "Industry cash cost per tonne", "Rs/t", C.COST26,
             [cr_f("11 Cost Curve", "cost", c) for c in FCOLS], "11 Cost Curve", "Medium", "",
             FMT_INT, key="cost")
    r = line(ws, r, "INDUSTRY EBITDA PER TONNE", "Rs/t", C.EBT26,
             ["=%s%d-%s%d" % (GL(c), R(ws.title, "realn"), GL(c), R(ws.title, "cost"))
              for c in FCOLS], "Computed", "Medium",
             "FY2026A of Rs 10,733/t is the volume-weighted actual across the four majors. "
             "Reconciliation: Rs 10,733/t x 79.90 Mt = Rs 85,759 cr against the Rs 86,318 cr sum "
             "of reported EBITDAs, a 0.6% rounding difference.", FMT_INT, bold=True, key="ebt",
             cagr=True)
    r = line(ws, r, "Finished steel production", "Mt", C.PROD26,
             [cr_f("05 Steel Supply Model", "prod", c) for c in FCOLS], "05 Supply", "High", "",
             FMT_2DP, key="prod")
    r = line(ws, r, "INDUSTRY EBITDA", "Rs cr", C.PROD26 * C.EBT26 / 10.0,
             ["=%s%d*%s%d/10" % (GL(c), R(ws.title, "prod"), GL(c), R(ws.title, "ebt"))
              for c in FCOLS], "Computed", "Medium", "", FMT_INT, bold=True, key="ebitda",
             cagr=True)
    r += 2
    r = sect(ws, r, "COMPANY EBITDA PER TONNE (FY2026 spread to industry held constant)")
    r = hdr(ws, r)
    for co in CO.COMPANIES:
        key, nm, ebt = co[0], co[1], co[8]
        if ebt is None:
            r = line(ws, r, "%s - EBITDA per tonne" % nm, "Rs/t", None, [NA] * NY, co[17],
                     "n/a", CO.CO_NOTES.get(key, ""), FMT_TEXT)
            continue
        spread = ebt - C.EBT26
        r = line(ws, r, "%s - EBITDA per tonne" % nm, "Rs/t", ebt,
                 ["=%s%d+%s" % (GL(c), R(ws.title, "ebt"), _n(spread)) for c in FCOLS],
                 co[17], "Medium",
                 "FY2026 actual Rs %s/t, a %+d Rs/t spread to the industry weighted average, "
                 "held constant. Basis: %s." % (format(ebt, ","), spread, co[7]),
                 FMT_INT, key="ebt_" + key)
    r += 1
    r = sect(ws, r, "COMPANY EBITDA")
    r = hdr(ws, r)
    for co in CO.COMPANIES:
        key, nm, ebt, eb_cr = co[0], co[1], co[8], co[6]
        if ebt is None or eb_cr is None:
            r = line(ws, r, "%s - EBITDA" % nm, "Rs cr", None, [NA] * NY, co[17], "n/a",
                     CO.CO_NOTES.get(key, ""), FMT_TEXT)
            continue
        r = line(ws, r, "%s - EBITDA" % nm, "Rs cr", eb_cr,
                 ["=%s%d*%s%d/10" % (GL(c), R("12 Revenue Forecast", "vol_" + key),
                                     GL(c), R(ws.title, "ebt_" + key)) for c in FCOLS],
                 co[17], "Medium", "Volume from 12 Revenue x EBITDA/t above.", FMT_INT,
                 key="eb_" + key, cagr=True)
    footnote(ws, r + 1, 2, "SAIL's EBITDA is on a STANDALONE basis while its revenue is "
                           "consolidated - database Conflict C09. Never pair a consolidated "
                           "enterprise value with SAIL's EBITDA headline without adjusting.")


def sh_margin(ws):
    r = meta(ws, "14", "Margin Analysis",
             "Decomposes margin into its price and cost components and benchmarks the forecast "
             "against observed cycle peaks and troughs.",
             "Margin is EBITDA per tonne divided by realisation. The observed envelope from the "
             "database is used as a credibility test on every scenario: FY2022 cycle peak was "
             "roughly 24-26% at industry level (Tata 26%, JSW 27%, SAIL 21%, Jindal Steel 30%), "
             "FY2024 trough roughly 12-14%, and the worst single observation was SAIL at -7% in "
             "FY2016. Any forecast outside that envelope requires explicit justification.",
             "Percentages as decimals. Rs per tonne.",
             "Reads 09 Price, 11 Cost, 13 EBITDA. Read by 18 Cycle, 21 Scenario, 24 Dashboard.")
    r = hdr(ws, r)
    r = line(ws, r, "Realisation", "Rs/t", C.REALN26,
             [cr_f("09 Steel Price Forecast", "realn", c) for c in FCOLS], "09 Price", "Medium",
             "", FMT_INT, key="realn")
    r = line(ws, r, "Cash cost", "Rs/t", C.COST26,
             [cr_f("11 Cost Curve", "cost", c) for c in FCOLS], "11 Cost", "Medium", "",
             FMT_INT, key="cost")
    r = line(ws, r, "EBITDA per tonne", "Rs/t", C.EBT26,
             [cr_f("13 EBITDA Model", "ebt", c) for c in FCOLS], "13 EBITDA", "Medium", "",
             FMT_INT, key="ebt")
    r = line(ws, r, "INDUSTRY EBITDA MARGIN", "%", C.EBT26 / C.REALN26,
             ["=%s%d/%s%d" % (GL(c), R(ws.title, "ebt"), GL(c), R(ws.title, "realn"))
              for c in FCOLS], "Computed", "Medium",
             "FY2026A of 17.9% is the volume-weighted actual across the four majors.",
             FMT_PCT_1, bold=True, key="margin")
    r += 1
    r = sect(ws, r, "MARGIN BRIDGE - CHANGE VS FY2026A")
    r = line(ws, r, "Realisation change vs FY2026A", "Rs/t", None,
             ["=%s%d-$D$%d" % (GL(c), R(ws.title, "realn"), R(ws.title, "realn"))
              for c in FCOLS], "Computed", "Medium", "", FMT_INT, key="d_realn")
    r = line(ws, r, "Cash cost change vs FY2026A", "Rs/t", None,
             ["=-(%s%d-$D$%d)" % (GL(c), R(ws.title, "cost"), R(ws.title, "cost"))
              for c in FCOLS], "Computed", "Medium",
             "Shown with the sign reversed so that the two bridge lines sum to the EBITDA/t "
             "change - a negative value means cost inflation is eroding margin.", FMT_INT,
             key="d_cost")
    r = line(ws, r, "EBITDA per tonne change vs FY2026A", "Rs/t", None,
             ["=%s%d+%s%d" % (GL(c), R(ws.title, "d_realn"), GL(c), R(ws.title, "d_cost"))
              for c in FCOLS], "Computed", "Medium",
             "BRIDGE CHECK: must equal EBITDA/t less the FY2026A EBITDA/t. Verified on 25 Audit "
             "Checks.", FMT_INT, bold=True, key="d_ebt")
    r += 1
    r = sect(ws, r, "CREDIBILITY BENCHMARKS FROM OBSERVED HISTORY (Master Database)")
    for lbl, val, note in [
        ("Observed cycle PEAK margin, FY2022", 0.26,
         "Standardised EBITDA margins FY2022: Tata Steel 26%, JSW Steel 27%, SAIL 21%, Jindal "
         "Steel 30%. The bull case is calibrated to peak at 26.0% in FY2030."),
        ("Observed cycle TROUGH margin, FY2024", 0.13,
         "FY2024: Tata Steel 10%, JSW Steel 16%, SAIL 11%, Jindal Steel 20%."),
        ("Worst single observation, SAIL FY2016", -0.07,
         "The floor reference. The stress case troughs at -9.2% in FY2028, marginally worse, "
         "which is deliberate: a stress case should breach the worst observed outcome."),
        ("FY2026A actual industry margin", C.EBT26 / C.REALN26,
         "Volume-weighted across the four majors."),
    ]:
        r = line(ws, r, lbl, "%", val, [], "Master DB", "High", note, FMT_PCT_1)
    footnote(ws, r + 1, 2, "If a scenario produces a margin outside the -7% to +26% observed "
                           "envelope, 25 Audit Checks flags it. That is a warning, not "
                           "necessarily an error - but it must be justified.")



# ======================================================================================
def sh_wc(ws):
    r = meta(ws, "15", "Working Capital Model",
             "Forecasts net working capital and the cash flow it absorbs or releases.",
             "NWC is modelled as days of revenue. THIS IS THE WEAKEST BLOCK IN THE MODEL and is "
             "flagged as such: the Master Database carries total assets and borrowings but not "
             "receivables, inventory or payables, so working capital days could NOT be derived "
             "from it. The only directional evidence available is Tata Steel's disclosed FY2026 "
             "working capital RELEASE of about Rs 6,470 cr, which on Rs 2,32,140 cr of revenue is "
             "roughly 10 days of tightening. 45 days is a sector convention and must be replaced "
             "with balance-sheet-derived days before transaction use.",
             "Days. Rs crore.",
             "Reads 12 Revenue and 02 Model Assumptions. Read by 16 Cash Flow.")
    r = hdr(ws, r)
    r = line(ws, r, "Industry revenue", "Rs cr", C.PROD26 * C.REALN26 / 10.0,
             [cr_f("12 Revenue Forecast", "rev", c) for c in FCOLS], "12 Revenue", "Medium", "",
             FMT_INT, key="rev")
    r = line(ws, r, "Net working capital", "days of revenue", 45,
             [dv("D18", c) for c in FCOLS], "NOT SOURCED", "Low",
             "45 days base, 40 bull, 52 bear, 62 stress. A working capital BUILD in a downturn is "
             "a first-order cash flow risk, not a second-order one, which is why the stress flex "
             "is large.", FMT_INT, key="days")
    r = line(ws, r, "Net working capital", "Rs cr", C.PROD26 * C.REALN26 / 10.0 * 45 / 365.0,
             ["=%s%d*%s%d/365" % (GL(c), R(ws.title, "rev"), GL(c), R(ws.title, "days"))
              for c in FCOLS], "Computed", "Low", "", FMT_INT, key="nwc")
    r = line(ws, r, "Change in net working capital", "Rs cr", None,
             ["=%s%d-%s%d" % (GL(c), R(ws.title, "nwc"), P(c), R(ws.title, "nwc"))
              for c in FCOLS], "Computed", "Low",
             "Positive equals a cash OUTFLOW (working capital absorbed). Deducted on 16 Cash "
             "Flow. FY2027E is measured against the FY2026A NWC computed on the same days basis, "
             "so the first year is not distorted by a basis change.", FMT_INT, bold=True,
             key="dnwc")
    r = line(ws, r, "NWC as % of revenue", "%", 45 / 365.0,
             ["=%s%d/%s%d" % (GL(c), R(ws.title, "nwc"), GL(c), R(ws.title, "rev"))
              for c in FCOLS], "Computed", "Low", "", FMT_PCT_1, key="nwc_pct")
    footnote(ws, r + 1, 2, "ACTION REQUIRED. Derive receivable, inventory and payable days from "
                           "each company's balance sheet and replace driver D18 before this model "
                           "is used to support a financing or transaction.")


def sh_cf(ws):
    r = meta(ws, "16", "Cash Flow Model",
             "Unlevered free cash flow for the industry - the input a DCF requires.",
             "UNLEVERED by design: EBITDA less cash tax less total capex less the change in "
             "working capital. Interest is deliberately excluded and sits on 17 Capital "
             "Allocation instead, because an enterprise DCF discounts pre-financing cash flow. "
             "Cash tax is computed on EBIT, not EBITDA, using depreciation of 5.3% of revenue "
             "DERIVED from the database - FY2026 D&A of Rs 30,715 cr on Rs 5,81,646 cr of revenue "
             "for the four majors, a ratio that is tight across companies at 5.15% to 5.96%. Tax "
             "is floored at zero on negative EBIT with no loss carry-forward modelled, which is "
             "conservative.",
             "Rs crore. Percentages as decimals.",
             "Reads 12 Revenue, 13 EBITDA, 15 Working Capital, 06 Capacity, 02 Assumptions. Read "
             "by 17, 21, 22, 23, 24.")
    r = hdr(ws, r)
    r = line(ws, r, "Industry EBITDA", "Rs cr", C.PROD26 * C.EBT26 / 10.0,
             [cr_f("13 EBITDA Model", "ebitda", c) for c in FCOLS], "13 EBITDA", "Medium", "",
             FMT_INT, key="ebitda", bold=True)
    r = line(ws, r, "Industry revenue", "Rs cr", C.PROD26 * C.REALN26 / 10.0,
             [cr_f("12 Revenue Forecast", "rev", c) for c in FCOLS], "12 Revenue", "Medium", "",
             FMT_INT, key="rev")
    r = line(ws, r, "Depreciation and amortisation", "% of revenue", 0.0528,
             [dv("D26", c) for c in FCOLS], "Derived from Master DB", "High",
             "FY2026A of 5.28% is derived: Rs 30,715 cr of D&A on Rs 5,81,646 cr of consolidated "
             "revenue for the four majors. Tight across companies - Tata 5.15%, JSW 5.18%, SAIL "
             "5.40%, Jindal Steel 5.96%.", FMT_PCT_1, key="dna_pct")
    r = line(ws, r, "Depreciation and amortisation", "Rs cr", None,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "rev"), GL(c), R(ws.title, "dna_pct"))
              for c in FCOLS], "Computed", "High", "", FMT_INT, key="dna")
    r = line(ws, r, "EBIT", "Rs cr", None,
             ["=%s%d-%s%d" % (GL(c), R(ws.title, "ebitda"), GL(c), R(ws.title, "dna"))
              for c in FCOLS], "Computed", "Medium", "", FMT_INT, key="ebit")
    r = line(ws, r, "Effective tax rate", "%", 0.2517, [dv("D15", c) for c in FCOLS],
             "Statutory 115BAA", "Medium",
             "22% plus surcharge and cess. Company-specific effective rates differ because of MAT "
             "credits and deferred tax - the database flags Kirloskar Ferrous FY2026 net profit "
             "exceeding PBT on a write-back (Conflict C12).", FMT_PCT_1, key="taxr")
    r = line(ws, r, "Cash tax", "Rs cr", None,
             ["=MAX(0,%s%d)*%s%d" % (GL(c), R(ws.title, "ebit"), GL(c), R(ws.title, "taxr"))
              for c in FCOLS], "Computed", "Medium",
             "Floored at zero. No loss carry-forward is modelled, which understates cash flow in "
             "any recovery following a loss year - conservative and disclosed.", FMT_INT,
             key="tax")
    r += 1
    r = sect(ws, r, "CAPITAL EXPENDITURE")
    r = line(ws, r, "Capacity added", "Mtpa", 20.07,
             [cr_f("06 Capacity Forecast", "add_tot", c) for c in FCOLS], "06 Capacity",
             "Medium", "", FMT_2DP, key="add")
    r = line(ws, r, "Capex intensity", "Rs/t of capacity", None, [dv("D16", c) for c in FCOLS],
             "Derived from Master DB", "Medium",
             "Rs 55,000/t base. Cross-checks: JSW JVML 5 Mtpa at Rs 26,000 cr = Rs 52,000/t; "
             "SAIL roughly Rs 1 lakh cr for about 15 Mtpa = roughly Rs 67,000/t; Tata Ludhiana "
             "0.75 Mtpa EAF at Rs 3,200 cr = Rs 42,700/t (lower, no ironmaking).", FMT_INT,
             key="capint")
    r = line(ws, r, "Growth capex", "Rs cr", None,
             ["=%s%d*%s%d/10" % (GL(c), R(ws.title, "add"), GL(c), R(ws.title, "capint"))
              for c in FCOLS], "Computed", "Medium",
             "Unit bridge: Mtpa x Rs/t / 10 = Rs crore.", FMT_INT, key="gcapex")
    r = line(ws, r, "Maintenance capex", "% of revenue", None, [dv("D17", c) for c in FCOLS],
             "Indicative", "Low",
             "3.0% of revenue. Cross-check: Tata Steel FY2026 total capex of Rs 14,026 cr on "
             "Rs 2,32,140 cr revenue is 6.0%, but that includes growth capex, so 3.0% sustaining "
             "implies a roughly 50/50 split.", FMT_PCT_1, key="mcapex_pct")
    r = line(ws, r, "Maintenance capex", "Rs cr", None,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "rev"), GL(c), R(ws.title, "mcapex_pct"))
              for c in FCOLS], "Computed", "Low", "", FMT_INT, key="mcapex")
    r = line(ws, r, "Total capex", "Rs cr", None,
             ["=%s%d+%s%d" % (GL(c), R(ws.title, "gcapex"), GL(c), R(ws.title, "mcapex"))
              for c in FCOLS], "Computed", "Medium", "", FMT_INT, key="capex", bold=True)
    r += 1
    r = sect(ws, r, "FREE CASH FLOW")
    r = line(ws, r, "Change in net working capital", "Rs cr", None,
             [cr_f("15 Working Capital Model", "dnwc", c) for c in FCOLS], "15 WC", "Low", "",
             FMT_INT, key="dnwc")
    r = line(ws, r, "UNLEVERED FREE CASH FLOW", "Rs cr", None,
             ["=%s%d-%s%d-%s%d-%s%d" % (GL(c), R(ws.title, "ebitda"), GL(c), R(ws.title, "tax"),
                                        GL(c), R(ws.title, "capex"), GL(c), R(ws.title, "dnwc"))
              for c in FCOLS], "Computed", "Medium",
             "EBITDA less cash tax less total capex less the change in working capital. "
             "Pre-financing, which is what an enterprise DCF discounts.", FMT_INT, bold=True,
             key="fcf")
    r = line(ws, r, "FCF conversion", "% of EBITDA", None,
             ["=IFERROR(%s%d/%s%d,\"\")" % (GL(c), R(ws.title, "fcf"), GL(c),
                                            R(ws.title, "ebitda")) for c in FCOLS],
             "Computed", "Medium",
             "Low or negative conversion in the early years is the signature of an industry in "
             "heavy build-out - roughly 20 Mtpa a year of capacity at Rs 55,000/t is about "
             "Rs 1.1 lakh crore of growth capex annually.", FMT_PCT_1, key="fcfconv")
    footnote(ws, r + 1, 2, "This is an INDUSTRY aggregate. For a single-name DCF, substitute that "
                           "company's volume, realisation spread, cost spread, capex programme "
                           "and working capital, all of which are already isolated as separate "
                           "lines in this workbook.")


def sh_capalloc(ws):
    r = meta(ws, "17", "Capital Allocation",
             "Traces the industry balance sheet: capex funding, the net debt path and leverage "
             "against the covered producers' own stated caps.",
             "Net debt rolls forward from an FY2026 industry aggregate of Rs 1,85,001 cr - the sum "
             "of Tata Steel Rs 80,144 cr, JSW Steel Rs 53,870 cr, Jindal Steel Rs 16,019 cr, "
             "Jindal Stainless Rs 3,040 cr and SAIL's gross borrowings of Rs 31,928 cr used as a "
             "PROXY because SAIL's net debt is not disclosed in the database. Unlevered FCF is "
             "applied to debt reduction with no dividend or buyback modelled at industry level.",
             "Rs crore. Leverage in times.",
             "Reads 16 Cash Flow and 13 EBITDA. Read by 23 Comparable Valuation and 25 Audit.")
    ND0 = 185001
    r = hdr(ws, r)
    r = line(ws, r, "Opening net debt", "Rs cr", None,
             ["=%s%d" % (P(c), r + 3) for c in FCOLS], "Computed", "Medium", "", FMT_INT,
             key="nd_open")
    r = line(ws, r, "Unlevered free cash flow", "Rs cr", None,
             [cr_f("16 Cash Flow Model", "fcf", c) for c in FCOLS], "16 Cash Flow", "Medium", "",
             FMT_INT, key="fcf")
    r = line(ws, r, "Interest paid (illustrative)", "Rs cr", None,
             ["=%s%d*'02 Model Assumptions'!$D$%d"
              % (GL(c), R(ws.title, "nd_open"), R("02 Model Assumptions", "wacc_calc") - 5)
              for c in FCOLS], "Computed", "Medium",
             "Opening net debt at the pre-tax cost of debt of 7.44% from the WACC build-up. "
             "Illustrative: actual coupons differ by issuer and by instrument.", FMT_INT,
             key="int")
    r = line(ws, r, "Closing net debt", "Rs cr", ND0,
             ["=%s%d-%s%d+%s%d" % (GL(c), R(ws.title, "nd_open"), GL(c), R(ws.title, "fcf"),
                                   GL(c), R(ws.title, "int")) for c in FCOLS],
             "Master DB (aggregate)", "Medium",
             "FY2026A of Rs 1,85,001 cr. Note SAIL is a gross-borrowings proxy, so the aggregate "
             "is overstated by SAIL's cash balance, which the database does not disclose.",
             FMT_INT, bold=True, key="nd")
    r += 1
    r = sect(ws, r, "LEVERAGE")
    r = line(ws, r, "Industry EBITDA", "Rs cr", C.PROD26 * C.EBT26 / 10.0,
             [cr_f("13 EBITDA Model", "ebitda", c) for c in FCOLS], "13 EBITDA", "Medium", "",
             FMT_INT, key="ebitda")
    r = line(ws, r, "Net debt / EBITDA", "x", ND0 / (C.PROD26 * C.EBT26 / 10.0),
             ["=IFERROR(%s%d/%s%d,\"\")" % (GL(c), R(ws.title, "nd"), GL(c),
                                            R(ws.title, "ebitda")) for c in FCOLS],
             "Computed", "Medium",
             "Benchmark against the producers' own stated caps from the database: JSW Steel "
             "revised its cap down from 3.75x to 3.00x and achieved 1.81x at Mar-2026; Jindal "
             "Steel 1.66x; Jindal Stainless 0.55x; Tata Steel 2.3x.", FMT_MULT, bold=True,
             key="nd_ebitda")
    r = line(ws, r, "Leverage covenant headroom vs 3.0x", "x", None,
             ["=3-%s%d" % (GL(c), R(ws.title, "nd_ebitda")) for c in FCOLS], "Computed",
             "Medium",
             "Against JSW Steel's revised stated cap of 3.00x, the tightest publicly stated cap "
             "among the covered producers. Negative means the industry aggregate would breach it.",
             FMT_MULT, key="headroom")
    r += 1
    r = sect(ws, r, "CUMULATIVE CAPITAL DEPLOYED")
    r = line(ws, r, "Total capex", "Rs cr", None,
             [cr_f("16 Cash Flow Model", "capex", c) for c in FCOLS], "16 Cash Flow", "Medium",
             "", FMT_INT, key="capex")
    r = line(ws, r, "Cumulative capex FY2027E onwards", "Rs cr", 0,
             ["=%s%d+%s%d" % (P(c), r, GL(c), R(ws.title, "capex")) for c in FCOLS],
             "Computed", "Medium",
             "The scale of the capital call on this industry over the horizon. For context the "
             "database records Jindal Steel's cumulative capex at Rs 40,450 cr through FY2026 and "
             "JSW Steel carrying forward Rs 96,888 cr of approved capex at 1-Apr-2026.",
             FMT_INT, bold=True, key="cumcapex")
    footnote(ws, r + 1, 2, "No dividend or buyback is modelled at industry level. The database "
                           "records FY2026 dividends of Rs 4.00 per share at Tata Steel, Rs 7.10 "
                           "at JSW Steel, Rs 2.35 at SAIL, Rs 2.00 at Jindal Steel and Rs 4.00 "
                           "at Jindal Stainless; add them in single-name models.")


def sh_cycle(ws):
    r = meta(ws, "18", "Industry Cycle Model",
             "Positions each forecast year in the industry cycle, so that a valuation is not "
             "struck on a peak or trough year by accident.",
             "Two independent cycle indicators are used: the utilisation gap against normal "
             "utilisation of 80%, and the margin gap against the mid-cycle margin. Mid-cycle "
             "margin is computed as the average of the observed FY2022 peak of 26% and the FY2024 "
             "trough of 13%, giving 19.5% - close to the FY2026 actual of 17.9%, which is "
             "reassuring. A composite score classifies each year.",
             "Percentages as decimals and percentage points.",
             "Reads 08 Utilisation and 14 Margin. Read by 23 Valuation and 24 Dashboard.")
    r = hdr(ws, r)
    r = line(ws, r, "Capacity utilisation", "%", C.CRUDE26 / C.CAP26,
             [cr_f("08 Capacity Utilisation", "util", c) for c in FCOLS], "08 Utilisation",
             "High", "", FMT_PCT_1, key="util")
    r = line(ws, r, "Utilisation gap vs normal (80%)", "ppt", C.CRUDE26 / C.CAP26 - 0.80,
             [cr_f("08 Capacity Utilisation", "utilgap", c) for c in FCOLS], "08 Utilisation",
             "High", "", FMT_PCT_1, key="utilgap")
    r = line(ws, r, "Industry EBITDA margin", "%", C.EBT26 / C.REALN26,
             [cr_f("14 Margin Analysis", "margin", c) for c in FCOLS], "14 Margin", "Medium", "",
             FMT_PCT_1, key="margin")
    r = line(ws, r, "Mid-cycle margin", "%", 0.195, [0.195] * NY, "Derived from Master DB",
             "Medium",
             "Average of the observed FY2022 peak (26%) and FY2024 trough (13%). Close to the "
             "FY2026 actual of 17.9%, which supports the calibration.", FMT_PCT_1, key="midcyc")
    r = line(ws, r, "Margin gap vs mid-cycle", "ppt", None,
             ["=%s%d-%s%d" % (GL(c), R(ws.title, "margin"), GL(c), R(ws.title, "midcyc"))
              for c in FCOLS], "Computed", "Medium", "", FMT_PCT_1, key="margingap")
    r = line(ws, r, "CYCLE POSITION", "classification", None,
             ['=IF(%s%d>0.04,"PEAK",IF(%s%d>0.015,"LATE UPCYCLE",'
              'IF(%s%d<-0.05,"TROUGH",IF(%s%d<-0.02,"DOWNCYCLE","MID-CYCLE"))))'
              % (GL(c), R(ws.title, "margingap"), GL(c), R(ws.title, "margingap"),
                 GL(c), R(ws.title, "margingap"), GL(c), R(ws.title, "margingap"))
              for c in FCOLS], "Computed", "Medium",
             "Classified on the margin gap: above +4ppt peak, +1.5 to +4ppt late upcycle, -2 to "
             "+1.5ppt mid-cycle, -2 to -5ppt downcycle, below -5ppt trough.", FMT_TEXT,
             bold=True, key="cyclepos")
    r += 1
    r = sect(ws, r, "OBSERVED CYCLE CHRONOLOGY (Master Database)")
    for lbl, val, note in [
        ("FY2022 - cycle peak", 0.26, "Industry margin roughly 24-26%. Coking coal spiked but was "
         "passed through, so margins were at RECORD highs - the key evidence that a cost spike "
         "is not automatically margin-negative."),
        ("FY2024 - cycle trough", 0.13, "Four years peak to trough."),
        ("FY2026 - recovery", C.EBT26 / C.REALN26, "Aided by the safeguard duty, which lifted "
         "domestic HRC roughly 25% between early Dec-2025 and end-Mar-2026."),
        ("Implied full cycle length", 7.5, "Roughly 7 to 8 years on a 4-year peak-to-trough leg. "
         "This is the quantitative basis for the 7-year horizon: the forecast average is "
         "approximately cycle-neutral."),
    ]:
        r = line(ws, r, lbl, "%" if val < 1 else "years", val, [], "Master DB", "High", note,
                 FMT_PCT_1 if val < 1 else FMT_1DP)
    footnote(ws, r + 1, 2, "Use the cycle position row when selecting an exit multiple. Applying a "
                           "mid-cycle multiple to a peak-year EBITDA is the most common way to "
                           "overvalue a steel company.")


def sh_trade(ws):
    r = meta(ws, "19", "Trade Model",
             "Forecasts imports, exports and the net trade position, and the penetration ratios "
             "that determine whether domestic prices set at import or export parity.",
             "Imports are a scenario driver plus any shortfall routed from 05 Steel Supply when "
             "the capacity ceiling binds. Exports are derived as the import driver plus net "
             "exports, so the two trade drivers are internally consistent by construction and "
             "cannot drift apart. Scenario signs are deliberately counter-intuitive and are "
             "explained on the rows.",
             "Mt. Percentages as decimals.",
             "Reads 02 Assumptions, 04 Demand, 05 Supply. Read by 24 Dashboard.")
    r = hdr(ws, r)
    r = line(ws, r, "Finished steel imports - driver", "Mt", C.IMP26,
             [dv("D21", c) for c in FCOLS], "Derived from Master DB", "Medium",
             "FY2026A of 6.524 Mt, down 31.7% from 9.551 Mt as the safeguard duty bit. NOTE THE "
             "SCENARIO SIGNS: the BULL case carries HIGHER imports because strong demand pulls "
             "material in; the BEAR case carries LOWER imports because weak demand makes imports "
             "uncompetitive.", FMT_2DP, key="imp_drv")
    r = line(ws, r, "Shortfall routed from supply ceiling", "Mt", 0.0,
             [cr_f("05 Steel Supply Model", "shortfall", c) for c in FCOLS], "05 Supply",
             "High", "Non-zero only where the capacity ceiling binds.", FMT_2DP, key="short")
    r = line(ws, r, "Total finished steel imports", "Mt", C.IMP26,
             ["=%s%d+%s%d" % (GL(c), R(ws.title, "imp_drv"), GL(c), R(ws.title, "short"))
              for c in FCOLS], "Computed", "Medium", "", FMT_2DP, bold=True, key="imp")
    r = line(ws, r, "Net finished steel exports", "Mt", 0.078,
             [cr_f("05 Steel Supply Model", "netexp", c) for c in FCOLS], "05 Supply", "Medium",
             "", FMT_2DP, key="netexp")
    r = line(ws, r, "Finished steel exports", "Mt", C.EXP26,
             ["=%s%d+%s%d" % (GL(c), R(ws.title, "imp_drv"), GL(c), R(ws.title, "netexp"))
              for c in FCOLS], "Computed", "Medium",
             "Derived as the import driver plus net exports, which guarantees the trade block "
             "reconciles to the supply balance on 05.", FMT_2DP, bold=True, key="exp")
    r += 1
    r = sect(ws, r, "TRADE INTENSITY")
    r = line(ws, r, "Apparent consumption", "Mt", C.CONS26,
             [cr_f("04 Steel Demand Model", "cons", c) for c in FCOLS], "04 Demand", "High", "",
             FMT_2DP, key="cons")
    r = line(ws, r, "Finished steel production", "Mt", C.PROD26,
             [cr_f("05 Steel Supply Model", "prod", c) for c in FCOLS], "05 Supply", "High", "",
             FMT_2DP, key="prod")
    r = line(ws, r, "Import penetration", "% of consumption", C.IMP26 / C.CONS26,
             ["=%s%d/%s%d" % (GL(c), R(ws.title, "imp"), GL(c), R(ws.title, "cons"))
              for c in FCOLS], "Computed", "Medium",
             "FY2026A of 4.0%, down from 6.3% in FY2025. Below roughly 5% is generally consistent "
             "with domestic producers retaining pricing power.", FMT_PCT_1, key="imppen")
    r = line(ws, r, "Export intensity", "% of production", C.EXP26 / C.PROD26,
             ["=%s%d/%s%d" % (GL(c), R(ws.title, "exp"), GL(c), R(ws.title, "prod"))
              for c in FCOLS], "Computed", "Medium",
             "FY2026A of 4.1%. India is fundamentally a domestic market, so global steel prices "
             "matter through the IMPORT channel rather than the export channel.", FMT_PCT_1,
             key="expint")
    r = line(ws, r, "Net trade position", "Mt", C.EXP26 - C.IMP26,
             ["=%s%d-%s%d" % (GL(c), R(ws.title, "exp"), GL(c), R(ws.title, "imp"))
              for c in FCOLS], "Computed", "Medium",
             "Positive equals net exporter. Determines whether the domestic price sets at import "
             "parity (net importer) or export parity (net exporter) - which drives the entire "
             "price deck.", FMT_2DP, bold=True, key="nettrade")
    r += 1
    r = sect(ws, r, "POLICY CONTEXT")
    r = line(ws, r, "Safeguard duty on flat products", "%", 0.12,
             [cr_f("09 Steel Price Forecast", "sgd", c) for c in FCOLS], "Master DB P03", "High",
             "Schedule, not a forecast. Expires 20-Apr-2028 subject to a mid-term review. Its "
             "expiry is the single largest policy risk to the FY2029 onwards price path.",
             FMT_PCT_1, key="sgd")
    footnote(ws, r + 1, 2, "EU destinations - Italy, Belgium and Spain - were 34.4% of India's "
                           "FY2026 exports, so the EU CBAM definitive regime that began "
                           "01-Jan-2026 bears on roughly a third of the export book. Quantified "
                           "on 20 ESG Model.")


def sh_esg(ws):
    r = meta(ws, "20", "ESG Model",
             "Quantifies the industry's position against India's Green Steel Taxonomy and its "
             "exposure to the EU Carbon Border Adjustment Mechanism.",
             "Emission intensity is benchmarked against the taxonomy thresholds notified on "
             "23-Dec-2024. CBAM exposure is computed as EU-destined exports multiplied by the "
             "excess of Indian emission intensity over an EU benchmark, multiplied by a carbon "
             "price. THE CARBON PRICE AND THE INDIAN EMISSION INTENSITY ARE NOT SOURCED and are "
             "user inputs - the calculation is a live framework, not an assertion.",
             "tCO2e per tonne of finished steel. EUR per tonne of CO2. Rs crore.",
             "Reads 19 Trade Model. Inputs are user-editable on this sheet.")
    r = hdr(ws, r)
    r = sect(ws, r, "GREEN STEEL TAXONOMY THRESHOLDS (notified 23-Dec-2024)")
    for lbl, v, note in [
        ("5-star green steel", 1.6, "Below 1.6 tCO2e/tfs. Tata Steel's 0.75 Mtpa Ludhiana scrap "
         "EAF, commissioned Mar-2026, is designed for under 0.3 tCO2e/t and therefore qualifies."),
        ("4-star green steel", 2.0, "1.6 to 2.0 tCO2e/tfs."),
        ("3-star green steel", 2.2, "2.0 to 2.2 tCO2e/tfs. This is also the headline definition "
         "of green steel."),
        ("Not classified as green steel", 2.2, "Above 2.2 tCO2e/tfs. Conventional Indian BF-BOF "
         "production sits well above this."),
    ]:
        r = line(ws, r, lbl, "tCO2e/tfs", v, [], "Master DB P15", "High", note, FMT_2DP)
    r += 1
    r = sect(ws, r, "USER INPUTS - NOT SOURCED, EDIT BEFORE USE")
    r = line(ws, r, "Indian average emission intensity", "tCO2e/tfs", 2.55, [2.55] * NY,
             "NOT SOURCED", "Low",
             "USER INPUT. Not sourced in this research cycle. Indian BF-BOF is generally "
             "understood to sit well above the 2.2 threshold, and the database records no "
             "company-level emission intensity disclosure. Replace with company Business "
             "Responsibility and Sustainability Report data.", FMT_2DP, key="ei",
             fill=FILL_NAVY_LIGHT)
    r = line(ws, r, "EU benchmark emission intensity", "tCO2e/tfs", 1.85, [1.85] * NY,
             "NOT SOURCED", "Low", "USER INPUT. Not sourced.", FMT_2DP, key="eu_ei",
             fill=FILL_NAVY_LIGHT)
    r = line(ws, r, "EU carbon price", "EUR/tCO2", 85.0, [85.0] * NY, "NOT SOURCED", "Low",
             "USER INPUT. Not sourced. EU ETS prices are observable and should be substituted.",
             FMT_1DP, key="co2px", fill=FILL_NAVY_LIGHT)
    r = line(ws, r, "EUR/INR", "Rs/EUR", 103.0, [103.0] * NY, "NOT SOURCED", "Low",
             "USER INPUT. Not sourced. Implied cross from the derived USD/INR of 88.4 at a "
             "EUR/USD of roughly 1.17.", FMT_1DP, key="eurinr", fill=FILL_NAVY_LIGHT)
    r = line(ws, r, "EU share of Indian exports", "%", 0.344, [0.344] * NY, "Master DB", "High",
             "Italy 16.2%, Belgium 10.9% and Spain 7.3% of FY2026 finished steel exports.",
             FMT_PCT_1, key="eushare")
    r += 1
    r = sect(ws, r, "CBAM EXPOSURE (live calculation on the inputs above)")
    r = line(ws, r, "Finished steel exports", "Mt", C.EXP26,
             [cr_f("19 Trade Model", "exp", c) for c in FCOLS], "19 Trade", "Medium", "",
             FMT_2DP, key="exp")
    r = line(ws, r, "EU-destined exports", "Mt", None,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "exp"), GL(c), R(ws.title, "eushare"))
              for c in FCOLS], "Computed", "Medium", "", FMT_2DP, key="eu_exp")
    r = line(ws, r, "Emission intensity excess vs EU benchmark", "tCO2e/tfs", None,
             ["=MAX(0,%s%d-%s%d)" % (GL(c), R(ws.title, "ei"), GL(c), R(ws.title, "eu_ei"))
              for c in FCOLS], "Computed", "Low", "", FMT_2DP, key="excess")
    r = line(ws, r, "CBAM cost", "Rs cr", None,
             ["=%s%d*%s%d*%s%d*%s%d/10" % (GL(c), R(ws.title, "eu_exp"), GL(c),
                                           R(ws.title, "excess"), GL(c), R(ws.title, "co2px"),
                                           GL(c), R(ws.title, "eurinr")) for c in FCOLS],
             "Computed", "Low",
             "Mt x tCO2e/t x EUR/tCO2 x Rs/EUR / 10 = Rs crore. Live on the user inputs above.",
             FMT_INT, bold=True, key="cbam")
    r = line(ws, r, "CBAM cost per tonne of EU-destined export", "Rs/t", None,
             ["=IFERROR(%s%d/%s%d*10,\"\")" % (GL(c), R(ws.title, "cbam"), GL(c),
                                               R(ws.title, "eu_exp")) for c in FCOLS],
             "Computed", "Low",
             "Compare against industry EBITDA per tonne on 13: this is the margin that CBAM "
             "removes from EU-destined tonnes unless emission intensity falls.", FMT_INT,
             key="cbam_t")
    footnote(ws, r + 1, 2, "Also unmodelled: the proposed 37% government procurement share for "
                           "green steel (a PROPOSAL, not an enacted mandate) and the Tata Steel "
                           "Netherlands going-concern uncertainty, which is company-specific and "
                           "outside this India model.")



# ======================================================================================
def sh_scenario(ws):
    r = meta(ws, "21", "Scenario Manager",
             "Compares all four scenarios side by side, and proves that the live workbook "
             "reproduces the intended arithmetic for whichever scenario is selected.",
             "Excel can only display one scenario at a time. The comparison block below was "
             "computed at BUILD TIME by fmodel/chain.py, which implements exactly the same "
             "arithmetic as the workbook formulas. The 'Live (active scenario)' column is a "
             "genuine link to the model. The tie-out row uses INDEX on ScenID to pick the "
             "pre-computed column matching the active scenario and compares it against the live "
             "value: if the Excel formulas and the intended arithmetic ever diverge, it fails. "
             "That is a real regression test, not a decoration.",
             "As stated per row. Terminal year FY2033E unless stated.",
             "Reads 01 Control Panel, 05, 06, 08, 09, 11, 13, 14, 16. Read by 25 Audit Checks.",
             105)
    hs = ["", "FY2033E metric", "Unit"] + SC + ["Live (active scenario)", "Tie-out", "Commentary"]
    set_widths(ws, [(1, 7), (2, 44), (3, 13), (4, 15), (5, 15), (6, 15), (7, 15), (8, 18),
                    (9, 13), (10, 105)])
    r = header_row(ws, r, hs, left_align_cols=(2, 10))
    first = r
    K = FCOLS[-1]
    rows = [
        ("Apparent consumption", "Mt", "cons", "05 Steel Supply Model", "cons", FMT_1DP,
         "Demand is the least contested part of the model: even the stress case grows "
         "consumption, because Indian demand has contracted only once (FY2021) in twelve years."),
        ("Crude steel production", "Mt", "crude", "05 Steel Supply Model", "crude", FMT_1DP, ""),
        ("Crude steel capacity", "Mtpa", "cap", "06 Capacity Forecast", "cap", FMT_1DP,
         "Capacity is scenario-dependent through the delivery factor and the secondary "
         "residual, but the range is narrow relative to the price range."),
        ("Capacity utilisation", "%", "util", "08 Capacity Utilisation", "util", FMT_PCT_1,
         "Counter-intuitively HIGHEST in the stress case, because capacity additions are cut "
         "harder than demand falls."),
        ("Blended realisation", "Rs/t", "realn", "09 Steel Price Forecast", "realn", FMT_INT,
         "Note the stress case sits ABOVE the bear case: a rupee collapse to 113/US$ lifts the "
         "rupee landed cost of imports by roughly 28% and supports domestic prices."),
        ("Cash cost per tonne", "Rs/t", "cost", "11 Cost Curve", "cost", FMT_INT, ""),
        ("EBITDA per tonne", "Rs/t", "ebt", "13 EBITDA Model", "ebt", FMT_INT, ""),
        ("Industry EBITDA margin", "%", "margin", "14 Margin Analysis", "margin", FMT_PCT_1,
         "Envelope check: bull 25.2% against the FY2022 observed peak of roughly 26%; stress "
         "9.9% in the terminal year having troughed at -9.2% in FY2028."),
        ("Industry revenue", "Rs cr", "rev", "12 Revenue Forecast", "rev", FMT_INT, ""),
        ("Industry EBITDA", "Rs cr", "ebitda", "13 EBITDA Model", "ebitda", FMT_INT,
         "The spread from stress to bull is roughly 4.4x, which is the honest measure of "
         "uncertainty in an Indian steel forecast seven years out."),
        ("Unlevered free cash flow", "Rs cr", "fcf", "16 Cash Flow Model", "fcf", FMT_INT,
         "Negative in bear and stress: the industry cannot fund roughly Rs 1.1 lakh crore a "
         "year of growth capex from a compressed margin."),
    ]
    tie_rows = []
    for lbl, unit, ck, sheet, key, fmt, comm in rows:
        put(ws, r, 2, lbl, alignment=AL_LEFT_WRAP, f=font(10, bold=True))
        put(ws, r, 3, unit, f=font(9, colour=TEXT_MUTED))
        for i, s in enumerate(SC):
            put(ws, r, 4 + i, C.ALL[s][ck][-1], number_format=fmt, alignment=AL_RIGHT)
        put(ws, r, 8, "='%s'!%s$%d" % (sheet, GL(K), R(sheet, key)), number_format=fmt,
            alignment=AL_RIGHT, f=font(10, bold=True, colour=NAVY))
        put(ws, r, 9, '=IF(ABS(H%d-INDEX(D%d:G%d,ScenID))<=ABS(H%d)*0.005+0.01,"OK","FAIL")'
            % (r, r, r, r), alignment=AL_CENTRE, f=font(10, bold=True))
        put(ws, r, 10, comm, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        tie_rows.append(r)
        r += 1
    add_table(ws, "tbl_Scenarios", first - 1, 1, r - 1, 10)
    K2 = r
    put(ws, r, 2, "TIE-OUT SUMMARY", f=font(11, bold=True, colour=NAVY))
    put(ws, r, 9, '=IF(COUNTIF(I%d:I%d,"FAIL")=0,"ALL OK","%d CHECK(S) FAILED")'
        % (tie_rows[0], tie_rows[-1], 0), alignment=AL_CENTRE,
        f=font(11, bold=True, colour=NAVY))
    put(ws, r, 10, "Tolerance is 0.5% plus an absolute floor, to absorb floating-point and "
                   "rounding differences between Excel and the Python mirror. Any genuine "
                   "formula error would exceed it by orders of magnitude.",
        f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    _K(ws, "tie_summary", r)
    r += 2
    r = sect(ws, r, "SCENARIO DEFINITIONS")
    for s, txt in [
        ("Base Case", "RBI's June-2026 FY2027 GDP projection of 6.6% converging to 6.5%. "
         "Elasticity 1.15x tapering to 1.05x. Realisation grows at roughly 2.0% a year, BELOW "
         "expected inflation, reflecting structural surplus capacity. Iron ore continues its "
         "decline to US$93/dmt; coking coal eases to US$200/t. Terminal margin 16.8% against "
         "17.9% in FY2026 - i.e. the base case assumes mild margin erosion, not expansion."),
        ("Bull Case", "GDP roughly 80bps above base, elasticity 1.30x tapering to 1.15x, iron ore "
         "and coking coal at the low end, INR stable at 96. Calibrated to peak at a 26.0% "
         "industry margin in FY2030, matching the FY2022 observed cycle peak. This is a genuine "
         "ceiling, not an extrapolation."),
        ("Bear Case", "A DEMAND-LED downturn. GDP roughly 100bps below base and elasticity at "
         "1.00x, but critically iron ore and coking coal FALL WITH demand - to US$83/dmt and "
         "US$170/t - because a global demand shortfall is disinflationary for raw materials. "
         "Modelling a bear case with demand down AND raw materials up would be internally "
         "contradictory. Margins compress to 7-11% rather than collapsing."),
        ("Stress Case", "A STAGFLATIONARY supply shock plus demand shock. Coking coal spikes to "
         "US$300/t - roughly 85% of the Oct-2023 peak of US$354/t - the rupee collapses to "
         "113/US$, GDP falls to 4.4% in FY2027 and conversion cost inflation hits 7.5%. "
         "Realisation is HIGHER than the bear case in nominal rupees because of the import-parity "
         "effect of the currency, yet margin troughs at -9.2% in FY2028, marginally worse than "
         "the worst observation in the database (SAIL, FY2016, -7%). IMPORTANT LIMITATION: the "
         "model does not endogenise supply response, so it does not idle capacity when margins "
         "turn negative. Read the stress case as a lower bound on margin, not a central "
         "expectation.")]:
        put(ws, r, 2, s, f=font(10, bold=True, colour=NAVY))
        put(ws, r, 4, txt, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        ws.row_dimensions[r].height = 13 * (1 + len(txt) // 120)
        r += 1


def _K(ws, key, row):
    RW[(ws.title, key)] = row
    return row


def sh_sens(ws):
    r = meta(ws, "22", "Sensitivity Analysis",
             "Isolates how much of the answer depends on each assumption, including the four that "
             "are unsourced.",
             "Three constructs. (1) A tornado ranking every driver by the FY2033 industry EBITDA "
             "impact of a 10% ADVERSE move - the direction that reduces EBITDA, which is +10% for "
             "cost drivers and -10% for revenue and volume drivers; getting that direction wrong "
             "is the commonest tornado error, so it is stated per row. Tornado values are "
             "computed by the Python mirror because each requires a full chain re-run. (2) Two "
             "LIVE two-variable tables that recompute in Excel. (3) A dedicated block quantifying "
             "the unsourced assumptions.",
             "Rs crore, Rs per tonne, percentages as decimals.",
             "Reads the whole model. Read by 24 Dashboard.", 100)
    base, tor = C.tornado()
    r = sect(ws, r, "A - TORNADO: FY2033E INDUSTRY EBITDA, 10% ADVERSE MOVE (Base Case)")
    hs = ["", "Driver", "ID", "Adverse direction", "FY2033E EBITDA (Rs cr)", "Change (Rs cr)",
          "Change (%)", "Interpretation"]
    set_widths(ws, [(1, 7), (2, 42), (3, 8), (4, 18), (5, 20), (6, 18), (7, 12), (8, 100)])
    r = header_row(ws, r, hs, left_align_cols=(2, 8))
    first = r
    interp = {
        "D09": "By far the dominant driver. A 10% realisation move is worth roughly 60% of "
               "industry EBITDA, which is the arithmetic of a 17.9% margin: a 10% price move is "
               "a 56% move in the margin. Any valuation of this sector is primarily a price call.",
        "D03": "The second largest, and frequently underestimated. Iron ore and coking coal are "
               "dollar-priced, so a 10% rupee depreciation raises rupee input cost across the "
               "whole basket simultaneously. Note the model does NOT credit the offsetting "
               "import-parity support to realisation in this isolated flex - that interaction is "
               "captured in the stress scenario, which is why the stress case is not simply the "
               "sum of these individual flexes.",
        "D01": "Works through both volume and, via utilisation, price. The fact that GDP and "
               "elasticity show identical impacts confirms the model treats them as a pure "
               "product, as intended.",
        "D02": "Identical to GDP by construction - demand growth is their product.",
        "D10": "Iron ore at 45% of the raw material basket, which is 60% of cash cost.",
        "D11": "Coking coal, same basket weight. Note it ranks BELOW iron ore only because the "
               "FY2033 base level is closer to its FY2026 reference; in the stress case coking "
               "coal is the dominant cost driver.",
        "D12": "Conversion cost inflation on the 40% of cash cost not indexed to raw materials.",
        "D13": "POSITIVE because raising the raw-material share shifts weight away from "
               "conversion cost, which is compounding faster than the raw material basket in the "
               "base case. This is a genuine model insight: the direction of this UNSOURCED "
               "parameter's effect depends on the relative escalation of the two components.",
        "D22": "NEGATIVE - more capacity is EBITDA-negative because it depresses utilisation, "
               "which feeds back into price via D23. This is the economically correct result and "
               "the opposite of the naive intuition that more capacity means more earnings.",
        "D05": "Same mechanism as D22, smaller because announced additions are a minority of "
               "total additions.",
    }
    for code, name, mult, b0, fl, delta, pct in tor:
        put(ws, r, 2, name, alignment=AL_LEFT_WRAP, f=font(10, bold=True))
        put(ws, r, 3, code, alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 4, "%+d%%" % (mult * 100), alignment=AL_CENTRE)
        put(ws, r, 5, fl, number_format=FMT_INT, alignment=AL_RIGHT)
        put(ws, r, 6, delta, number_format=FMT_INT, alignment=AL_RIGHT)
        put(ws, r, 7, pct, number_format=FMT_PCT_1, alignment=AL_RIGHT)
        put(ws, r, 8, interp.get(code, ""), f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    add_table(ws, "tbl_Tornado", first - 1, 1, r - 1, 8)
    put(ws, r, 2, "Base Case FY2033E industry EBITDA", f=font(10, bold=True))
    put(ws, r, 5, base, number_format=FMT_INT, alignment=AL_RIGHT,
        f=font(10, bold=True, colour=NAVY))
    put(ws, r, 8, "Computed by the Python mirror. Cross-checked against the live workbook by the "
                  "tie-out on 21 Scenario Manager.", f=font(9, colour=TEXT_MUTED),
        alignment=AL_LEFT_WRAP)
    r += 3

    # ---- helper constants for the live tables
    r = sect(ws, r, "B - HELPER CONSTANTS FOR THE LIVE TABLES (all linked, none re-entered)")
    KL = FCOLS[-1]
    helpers = [
        ("FY2026 cash cost anchor", C.COST26, "Rs/t", None, FMT_INT, "c26"),
        ("Raw material share of cash cost", None, "%", dv("D13", KL), FMT_PCT_1, "rmshare"),
        ("Iron ore weight in basket", None, "%", dv("D14", KL), FMT_PCT_1, "iow"),
        ("FY2033E conversion cost index", None, "x",
         cr_f("10 Raw Material Forecast", "convidx", KL), FMT_3DP, "convidx"),
        ("FY2033E iron ore in rupees", None, "Rs/dmt",
         cr_f("10 Raw Material Forecast", "io_inr", KL), FMT_INT, "io_inr"),
        ("FY2026 iron ore in rupees", C.IO26 * C.INR26, "Rs/dmt", None, FMT_INT, "io_inr26"),
        ("FY2026 coking coal in rupees", C.CC26 * C.INR26, "Rs/t", None, FMT_INT, "cc_inr26"),
        ("FY2033E USD/INR", None, "Rs/US$", cr_f("10 Raw Material Forecast", "inr", KL),
         FMT_2DP, "inr"),
        ("FY2033E price adjustment factor", None, "x",
         cr_f("09 Steel Price Forecast", "px_adj", KL), FMT_3DP, "pxadj"),
        ("FY2033E capacity", None, "Mtpa", cr_f("06 Capacity Forecast", "cap", KL), FMT_1DP,
         "cap"),
        ("FY2033E crude-to-finished ratio", None, "x", cr_f("05 Steel Supply Model", "cfr", KL),
         FMT_MULT, "cfr"),
        ("FY2033E cash cost", None, "Rs/t", cr_f("11 Cost Curve", "cost", KL), FMT_INT, "cost"),
    ]
    for lbl, val, unit, f, fmt, key in helpers:
        put(ws, r, 2, lbl, alignment=AL_LEFT_WRAP)
        put(ws, r, 3, unit, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 4, f if f else val, number_format=fmt, alignment=AL_RIGHT)
        _K(ws, "h_" + key, r)
        r += 1
    r += 2

    def H(key):
        return "$D$%d" % R(ws.title, "h_" + key)

    # ---- Table 1: realisation x coking coal -> FY2033 EBITDA/t (LIVE)
    r = sect(ws, r, "C - LIVE TABLE 1: FY2033E EBITDA PER TONNE (Rs/t) - realisation vs coking coal")
    put(ws, r, 2, "Rows are the FY2033E realisation driver (Rs/t); columns are coking coal "
                  "(US$/t). Every cell is a live formula rebuilding the cost equation from the "
                  "helper constants above - change any driver and the whole grid moves.",
        f=font(9, italic=True, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    ws.row_dimensions[r].height = 26
    r += 1
    ccs = [160, 180, 200, 220, 240, 260, 280]
    rns = [58000, 61000, 64000, 66000, 68000, 71000, 74000, 77000]
    put(ws, r, 2, "Realisation \\ Coking coal", f=font(9, bold=True, colour=WHITE) if False
        else font(9, bold=True), alignment=AL_CENTRE, fill=FILL_NAVY_LIGHT)
    for j, cc in enumerate(ccs):
        put(ws, r, 4 + j, cc, number_format=FMT_INT, alignment=AL_CENTRE,
            f=font(9, bold=True), fill=FILL_NAVY_LIGHT)
    hdr_row = r
    r += 1
    t1_first = r
    for rn in rns:
        put(ws, r, 2, rn, number_format=FMT_INT, alignment=AL_RIGHT, f=font(9, bold=True),
            fill=FILL_NAVY_LIGHT)
        for j in range(len(ccs)):
            cl = GL(4 + j)
            cost = ("%s*(%s*(%s*(%s/%s)+%s*(%s%d*%s/%s)+(1-2*%s)*%s)+(1-%s)*%s)"
                    % (H("c26"), H("rmshare"), H("iow"), H("io_inr"), H("io_inr26"), H("iow"),
                       cl, hdr_row, H("inr"), H("cc_inr26"), H("iow"), H("convidx"),
                       H("rmshare"), H("convidx")))
            put(ws, r, 4 + j, "=$B%d*%s-(%s)" % (r, H("pxadj"), cost), number_format=FMT_INT,
                alignment=AL_RIGHT)
        r += 1
    t1_last = r - 1
    put(ws, r, 2, "Interpretation", f=font(9, bold=True, colour=NAVY))
    put(ws, r, 4, "At the Base Case FY2033E realisation driver of Rs 68,000/t, EBITDA per tonne "
                  "swings from roughly Rs 14,600/t at US$160/t coking coal to roughly Rs 6,100/t "
                  "at US$280/t - a 2.4x range on the coking coal axis alone. This is why a "
                  "licensed coking coal series is the highest-priority data gap in the model.",
        f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    ws.row_dimensions[r].height = 26
    r += 3

    # ---- Table 2: utilisation x realisation -> FY2033 industry EBITDA (LIVE)
    r = sect(ws, r, "D - LIVE TABLE 2: FY2033E INDUSTRY EBITDA (Rs cr) - utilisation vs realisation")
    put(ws, r, 2, "Rows are FY2033E capacity utilisation; columns are the FY2033E realisation "
                  "driver (Rs/t). Volume is derived as capacity x utilisation / the "
                  "crude-to-finished ratio. This is a PARTIAL sensitivity: it holds the lagged "
                  "price feedback at its base value, so it isolates the volume and price effects "
                  "and does not double-count the feedback.",
        f=font(9, italic=True, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    ws.row_dimensions[r].height = 26
    r += 1
    utils = [0.74, 0.78, 0.82, 0.84, 0.86, 0.90]
    put(ws, r, 2, "Utilisation \\ Realisation", f=font(9, bold=True), alignment=AL_CENTRE,
        fill=FILL_NAVY_LIGHT)
    for j, rn in enumerate(rns):
        put(ws, r, 4 + j, rn, number_format=FMT_INT, alignment=AL_CENTRE, f=font(9, bold=True),
            fill=FILL_NAVY_LIGHT)
    h2 = r
    r += 1
    t2_first = r
    for u in utils:
        put(ws, r, 2, u, number_format=FMT_PCT_1, alignment=AL_RIGHT, f=font(9, bold=True),
            fill=FILL_NAVY_LIGHT)
        for j in range(len(rns)):
            cl = GL(4 + j)
            put(ws, r, 4 + j,
                "=(%s*$B%d/%s)*(%s%d*%s-%s)/10" % (H("cap"), r, H("cfr"), cl, h2, H("pxadj"),
                                                   H("cost")),
                number_format=FMT_INT, alignment=AL_RIGHT)
        r += 1
    t2_last = r - 1
    r += 2

    r = sect(ws, r, "E - HOW MUCH OF THE ANSWER RESTS ON UNSOURCED ASSUMPTIONS")
    hs = ["", "Unsourced assumption", "ID", "Base value", "Plausible range",
          "FY2033E EBITDA impact across the range", "Materiality verdict"]
    r = header_row(ws, r, hs, left_align_cols=(2, 5, 6, 7))
    first = r
    for nm, code, base_v, rng, impact, verdict in [
        ("Raw material share of cash cost", "D13", "60%", "55% to 65%",
         "Roughly Rs 21,000 cr per 10% relative change, i.e. about 6.7% of FY2033E EBITDA",
         "MATERIAL but second-order. The direction depends on relative escalation of raw material "
         "versus conversion cost, so it is not a simple bias."),
        ("Price sensitivity to utilisation", "D23", "0.40x", "0.20x to 0.60x",
         "Sets the whole magnitude of the capacity feedback. At 0.20x the capacity drivers halve "
         "in importance; at 0.60x they roughly double.",
         "MATERIAL. This parameter determines whether capacity matters at all. It is the single "
         "most important unsourced input in the model."),
        ("Net working capital days", "D18", "45 days", "35 to 60 days",
         "Roughly Rs 12,000 cr of terminal-year cash flow per 4.5 days",
         "MATERIAL FOR CASH FLOW AND CREDIT, immaterial for EBITDA. Must be fixed before any "
         "financing work."),
        ("Equity risk premium and beta (within WACC)", "R08", "6.5% and 1.20x", "5.5-7.5%, 1.0-1.4x",
         "Roughly 11% of enterprise value per 120bps of WACC",
         "MATERIAL FOR VALUATION ONLY. Does not touch operating forecasts. Substitute house "
         "assumptions on 02 Section D."),
        ("Maximum practical utilisation", "D25", "92%", "88% to 95%",
         "Zero in the Base Case - the ceiling does not bind, peak utilisation is 84.0%",
         "IMMATERIAL in the base case; becomes material only in the bull case where utilisation "
         "reaches 87.9%."),
        ("Maintenance capex", "D17", "3.0% of revenue", "2.5% to 4.0%",
         "Roughly Rs 16,000 cr of terminal-year free cash flow per 100bps",
         "MATERIAL FOR CASH FLOW, immaterial for EBITDA."),
        ("Indian emission intensity and EU carbon price", "20 ESG", "2.55 tCO2e/t, EUR 85/t",
         "Wide - both unsourced",
         "CBAM cost is roughly Rs 2,100 cr at the base inputs, under 1% of FY2033E EBITDA",
         "IMMATERIAL at current export volumes, but scales with the EU export share and the "
         "carbon price. Watch rather than model."),
    ]:
        put(ws, r, 2, nm, alignment=AL_LEFT_WRAP, f=font(10, bold=True))
        put(ws, r, 3, code, alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 4, base_v, alignment=AL_CENTRE)
        put(ws, r, 5, rng, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 6, impact, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 7, verdict, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    add_table(ws, "tbl_Unsourced", first - 1, 1, r - 1, 7)
    set_widths(ws, [(5, 26), (6, 52), (7, 62)])
    footnote(ws, r + 1, 2, "Read this block first. It is the honest answer to 'how much should I "
                           "trust this model?' - the operating forecast is robust to the "
                           "unsourced inputs; the cash flow and valuation output is not, until "
                           "working capital, WACC and a coking coal series are properly sourced.")


WHITE = "FFFFFF"



# ======================================================================================
def sh_comps(ws):
    r = meta(ws, "23", "Comparable Valuation",
             "Valuation output: trading comparables, EV per tonne, replacement cost, ROIC and a "
             "DCF built off the industry free cash flow.",
             "SHARE PRICES WERE NOT SOURCED in this research cycle, so they are USER INPUTS "
             "(shaded). Every multiple, market capitalisation and enterprise value on this sheet "
             "is a live formula that computes the moment prices are entered - nothing has been "
             "fabricated to fill the gap. Share counts are DERIVED as profit after tax divided by "
             "earnings per share, which reconciles well for SAIL, Jindal Steel, Jindal Stainless "
             "and Tata Steel but NOT for JSW Steel, where EPS is struck on profit attributable to "
             "owners while PAT includes an Rs 18,051 cr exceptional gain; that row is flagged and "
             "the cell is editable.",
             "Rs crore, Rs per share, times, Rs per tonne.",
             "Reads 13 EBITDA, 16 Cash Flow, 17 Capital Allocation, 02 Assumptions.", 95)
    set_widths(ws, [(1, 7), (2, 22), (3, 13), (4, 13), (5, 13), (6, 14), (7, 14), (8, 14),
                    (9, 14), (10, 12), (11, 12), (12, 13), (13, 13), (14, 12), (15, 95)])
    r = sect(ws, r, "A - TRADING COMPARABLES (enter share prices in the shaded column)")
    hs = ["", "Company", "FY2026 EBITDA (Rs cr)", "Net debt (Rs cr)", "PAT (Rs cr)", "EPS (Rs)",
          "Shares (cr) - derived", "Share price (Rs) - USER INPUT", "Market cap (Rs cr)",
          "EV (Rs cr)", "EV/EBITDA (x)", "Capacity (Mtpa)", "EV per tonne (Rs/t)", "Conf.",
          "Notes"]
    r = header_row(ws, r, hs, left_align_cols=(2, 15))
    first = r
    for co in CO.COMPANIES:
        key, nm = co[0], co[1]
        eb, nd, pat, eps = co[6], co[11], co[13], co[14]
        cap, capnote = CO.CO_CAPACITY.get(key, (None, ""))
        put(ws, r, 2, nm, f=font(10, bold=True), alignment=AL_LEFT_WRAP)
        if eb:
            put(ws, r, 3, eb, number_format=FMT_INT, alignment=AL_RIGHT)
        if nd:
            put(ws, r, 4, nd, number_format=FMT_INT, alignment=AL_RIGHT)
        if pat:
            put(ws, r, 5, pat, number_format=FMT_INT, alignment=AL_RIGHT)
        if eps:
            put(ws, r, 6, eps, number_format=FMT_2DP, alignment=AL_RIGHT)
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
        put(ws, r, 14, "Medium" if eps else "n/a", alignment=AL_CENTRE,
            f=font(9, colour=TEXT_MUTED))
        note = CO.CO_NOTES.get(key, "")
        if key == "JSW":
            note += (" DERIVED SHARE COUNT IS UNRELIABLE for JSW: PAT of Rs 25,508 cr includes "
                     "the Rs 18,051 cr BPSL slump-sale gain while EPS of Rs 91.26 is on profit "
                     "attributable to owners, so PAT/EPS overstates the share count by roughly "
                     "15%. Overwrite column G with the actual count.")
        if key in ("AMNS", "RINL"):
            note += " Unlisted - no market-based valuation is possible."
        put(ws, r, 15, note + " Capacity basis: " + capnote, f=font(9, colour=TEXT_MUTED),
            alignment=AL_LEFT_WRAP)
        r += 1
    add_table(ws, "tbl_Comps", first - 1, 1, r - 1, 15)
    r += 1
    put(ws, r, 2, "Median EV/EBITDA (listed, once prices entered)", f=font(10, bold=True))
    put(ws, r, 11, '=IFERROR(MEDIAN(K%d:K%d),"enter prices")' % (first, r - 2),
        number_format=FMT_MULT, alignment=AL_RIGHT, f=font(10, bold=True, colour=NAVY))
    put(ws, r, 15, "Live. Compare against the exit multiple driver D19 of 6.0x, which is NOT "
                   "sourced and should be re-based on this observed evidence once prices are "
                   "entered.", f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    r += 3

    r = sect(ws, r, "B - REPLACEMENT COST AND EV PER TONNE BENCHMARK")
    for lbl, val, fmt, note in [
        ("Greenfield / brownfield capex intensity", None, FMT_INT,
         "From driver D16 - Rs 55,000/t base. Cross-checked against JSW JVML at Rs 52,000/t, SAIL "
         "at roughly Rs 67,000/t and Tata Ludhiana EAF at Rs 42,700/t."),
        ("Replacement cost of India's FY2026 capacity", None, FMT_INT,
         "220.4 Mtpa at the capex intensity above. A useful sanity bound: an acquirer should not "
         "pay materially more per tonne than it costs to build, adjusted for time to market, "
         "permits and captive raw material."),
        ("Implied EV per tonne at the exit multiple", None, FMT_INT,
         "FY2033E industry EBITDA at the exit multiple D19, divided by FY2033E capacity. Compare "
         "against replacement cost: an implied EV/t far above build cost signals either scarcity "
         "value or overvaluation."),
    ]:
        put(ws, r, 2, lbl, alignment=AL_LEFT_WRAP, f=font(10, bold=True))
        put(ws, r, 15, note, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        _K(ws, "rc_%d" % r, r)
        r += 1
    ci = dv("D16", FCOLS[-1])
    put(ws, r - 3, 4, ci, number_format=FMT_INT, alignment=AL_RIGHT)
    put(ws, r - 2, 4, "=%.1f*D%d/10" % (C.CAP26, r - 3), number_format=FMT_INT,
        alignment=AL_RIGHT, f=font(10, bold=True))
    put(ws, r - 1, 4, "=IFERROR('13 EBITDA Model'!%s$%d*%s/('06 Capacity Forecast'!%s$%d)*10,\"\")"
        % (GL(FCOLS[-1]), R("13 EBITDA Model", "ebitda"), dv("D19", FCOLS[-1]).replace("'02", "'02"),
           GL(FCOLS[-1]), R("06 Capacity Forecast", "cap")),
        number_format=FMT_INT, alignment=AL_RIGHT, f=font(10, bold=True))
    r += 2

    r = sect(ws, r, "C - INDUSTRY DCF (unlevered, off 16 Cash Flow Model)")
    r = hdr(ws, r)
    fr = R("16 Cash Flow Model", "fcf")
    r = line(ws, r, "Unlevered free cash flow", "Rs cr", None,
             ["='16 Cash Flow Model'!%s$%d" % (GL(c), fr) for c in FCOLS], "16 Cash Flow",
             "Medium", "", FMT_INT, key="fcf")
    r = line(ws, r, "WACC", "%", None, [dv("D20", c) for c in FCOLS], "02 Assumptions", "Medium",
             "Cross-checked against the build-up on 02 Section D by 25 Audit Checks.", FMT_PCT_1,
             key="wacc")
    r = line(ws, r, "Discount period", "years", None, [i + 0.5 for i in range(NY)],
             "Convention", "High",
             "MID-YEAR CONVENTION: cash flows are assumed to arise evenly through the year, so "
             "they are discounted from the mid-point. Using year-end would understate value by "
             "roughly 5-6% at a 12% WACC.", FMT_1DP, key="dp")
    r = line(ws, r, "Discount factor", "x", None,
             ["=1/(1+%s%d)^%s%d" % (GL(c), R(ws.title, "wacc"), GL(c), R(ws.title, "dp"))
              for c in FCOLS], "Computed", "Medium", "", FMT_3DP, key="df")
    r = line(ws, r, "Present value of FCF", "Rs cr", None,
             ["=%s%d*%s%d" % (GL(c), R(ws.title, "fcf"), GL(c), R(ws.title, "df"))
              for c in FCOLS], "Computed", "Medium", "", FMT_INT, key="pv")
    pvr = R(ws.title, "pv")
    r += 1
    for kk, lbl, f, fmt, note in [
        ("pvsum", "Sum of PV of explicit forecast FCF",
         "=SUM(%s%d:%s%d)" % (GL(CF1), pvr, GL(CFN), pvr),
         FMT_INT, "Seven years, FY2027E to FY2033E."),
        ("termebitda", "Terminal EBITDA (FY2033E)", "='13 EBITDA Model'!%s$%d"
         % (GL(FCOLS[-1]), R("13 EBITDA Model", "ebitda")), FMT_INT,
         "FY2033 is a NORMALISED year - the FY2030-FY2031 capacity cohort has reached steady-state "
         "utilisation - which is the whole reason the horizon is 7 years rather than 5. Check the "
         "cycle position on 18 before accepting it as mid-cycle."),
        ("exitmult", "Exit EV/EBITDA multiple", "=%s" % dv("D19", FCOLS[-1]), FMT_MULT,
         "6.0x base. NOT SOURCED - re-base on the observed median in section A once share prices "
         "are entered."),
        ("tv", "Terminal value", None, FMT_INT, "Terminal EBITDA times the exit multiple."),
        ("pvtv", "PV of terminal value", None, FMT_INT,
         "Discounted at the terminal-year discount factor."),
        ("ev", "ENTERPRISE VALUE", None, FMT_INT, "Sum of the explicit PV and the discounted "
         "terminal value."),
        ("tvpct", "Terminal value as % of EV", None, FMT_PCT_1,
         "IF THIS EXCEEDS ROUGHLY 75% the valuation is essentially a terminal-multiple assertion "
         "rather than a forecast, and the exit multiple - which is unsourced - dominates. Watch "
         "this number."),
    ]:
        put(ws, r, 2, lbl, alignment=AL_LEFT_WRAP, f=font(10, bold=True))
        if f:
            put(ws, r, CB, f, number_format=fmt, alignment=AL_RIGHT,
                f=font(10, bold=True, colour=NAVY))
        put(ws, r, CNOTE, note, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        _K(ws, "dcf_" + kk, r)
        r += 1
    b = R(ws.title, "dcf_pvsum")
    put(ws, b + 3, CB, "=D%d*D%d" % (b + 1, b + 2), number_format=FMT_INT, alignment=AL_RIGHT,
        f=font(10, bold=True, colour=NAVY))
    put(ws, b + 4, CB, "=D%d*%s%d" % (b + 3, GL(FCOLS[-1]), R(ws.title, "df")),
        number_format=FMT_INT, alignment=AL_RIGHT, f=font(10, bold=True, colour=NAVY))
    put(ws, b + 5, CB, "=D%d+D%d" % (b, b + 4), number_format=FMT_INT, alignment=AL_RIGHT,
        f=font(11, bold=True, colour=NAVY))
    put(ws, b + 6, CB, '=IFERROR(D%d/D%d,"")' % (b + 4, b + 5), number_format=FMT_PCT_1,
        alignment=AL_RIGHT, f=font(10, bold=True))
    r += 1
    r = sect(ws, r, "D - RETURNS")
    r = hdr(ws, r)
    r = line(ws, r, "EBIT", "Rs cr", None,
             ["='16 Cash Flow Model'!%s$%d" % (GL(c), R("16 Cash Flow Model", "ebit"))
              for c in FCOLS], "16 Cash Flow", "Medium", "", FMT_INT, key="ebit")
    r = line(ws, r, "Effective tax rate", "%", None, [dv("D15", c) for c in FCOLS],
             "02 Assumptions", "Medium", "", FMT_PCT_1, key="taxr")
    r = line(ws, r, "NOPAT", "Rs cr", None,
             ["=%s%d*(1-%s%d)" % (GL(c), R(ws.title, "ebit"), GL(c), R(ws.title, "taxr"))
              for c in FCOLS], "Computed", "Medium", "", FMT_INT, key="nopat")
    r = line(ws, r, "Invested capital (proxy)", "Rs cr", None,
             ["=%.1f*'02 Model Assumptions'!%s$%d/10+'15 Working Capital Model'!%s$%d"
              % (C.CAP26, GL(c), dv_row("D16"), GL(c), R("15 Working Capital Model", "nwc"))
              for c in FCOLS], "Computed - PROXY", "Low",
             "PROXY ONLY: FY2026 capacity at replacement cost plus net working capital. The "
             "database does not carry net fixed assets, so a true invested-capital base could not "
             "be built. A replacement-cost proxy understates ROIC for older, depreciated assets "
             "and is stated as indicative.", FMT_INT, key="ic")
    r = line(ws, r, "ROIC (indicative)", "%", None,
             ["=IFERROR(%s%d/%s%d,\"\")" % (GL(c), R(ws.title, "nopat"), GL(c), R(ws.title, "ic"))
              for c in FCOLS], "Computed", "Low", "", FMT_PCT_1, bold=True, key="roic")
    r = line(ws, r, "ROIC less WACC (value spread)", "ppt", None,
             ["=%s%d-%s%d" % (GL(c), R(ws.title, "roic"), GL(c), R(ws.title, "wacc"))
              for c in FCOLS], "Computed", "Low",
             "THE VALUE-CREATION TEST. A persistently negative spread means the industry destroys "
             "value by adding capacity, however strong the volume growth - which is the central "
             "strategic question for Indian steel and the reason capacity discipline matters more "
             "than capacity ambition.", FMT_PCT_1, bold=True, key="spread")
    footnote(ws, r + 1, 2, "This is an INDUSTRY DCF. It is a sanity frame and a source of "
                           "mid-cycle multiples, not a substitute for a single-name model. Every "
                           "line needed for a single-name build - volume, realisation spread, cost "
                           "spread, capex, working capital - is isolated elsewhere in this "
                           "workbook.")


def sh_dash(ws):
    r = meta(ws, "24", "Industry Dashboard",
             "One-page summary of the active scenario for an investment committee or board pack.",
             "Every figure is a live link. Nothing on this sheet is entered.",
             "As stated per row.", "Reads the whole model.", 95)
    put(ws, r, 2, "ACTIVE SCENARIO", f=font(11, bold=True, colour=NAVY))
    put(ws, r, CB, "='01 Control Panel'!$D$%d" % R("01 Control Panel", "sel"),
        f=font(12, bold=True, colour=NAVY), alignment=AL_CENTRE, fill=FILL_NAVY_LIGHT)
    r += 2
    r = hdr(ws, r)
    for lbl, unit, sheet, key, fmt in [
        ("DEMAND AND SUPPLY", "", None, None, None),
        ("Apparent finished steel consumption", "Mt", "04 Steel Demand Model", "cons", FMT_1DP),
        ("Consumption growth", "%", "04 Steel Demand Model", "dgrowth", FMT_PCT_1),
        ("Per capita consumption", "kg", "04 Steel Demand Model", "percap", FMT_1DP),
        ("Crude steel production", "Mt", "05 Steel Supply Model", "crude", FMT_1DP),
        ("Crude steel capacity", "Mtpa", "06 Capacity Forecast", "cap", FMT_1DP),
        ("Capacity utilisation", "%", "08 Capacity Utilisation", "util", FMT_PCT_1),
        ("PRICES AND COSTS", "", None, None, None),
        ("Blended realisation", "Rs/t", "09 Steel Price Forecast", "realn", FMT_INT),
        ("Iron ore 62% Fe CFR China", "US$/dmt", "10 Raw Material Forecast", "io", FMT_2DP),
        ("Premium HCC coking coal", "US$/t", "10 Raw Material Forecast", "cc", FMT_2DP),
        ("USD/INR", "Rs/US$", "10 Raw Material Forecast", "inr", FMT_2DP),
        ("Cash cost per tonne", "Rs/t", "11 Cost Curve", "cost", FMT_INT),
        ("PROFITABILITY", "", None, None, None),
        ("EBITDA per tonne", "Rs/t", "13 EBITDA Model", "ebt", FMT_INT),
        ("Industry EBITDA margin", "%", "14 Margin Analysis", "margin", FMT_PCT_1),
        ("Industry revenue", "Rs cr", "12 Revenue Forecast", "rev", FMT_INT),
        ("Industry EBITDA", "Rs cr", "13 EBITDA Model", "ebitda", FMT_INT),
        ("Cycle position", "class", "18 Industry Cycle Model", "cyclepos", FMT_TEXT),
        ("CASH AND BALANCE SHEET", "", None, None, None),
        ("Total capex", "Rs cr", "16 Cash Flow Model", "capex", FMT_INT),
        ("Unlevered free cash flow", "Rs cr", "16 Cash Flow Model", "fcf", FMT_INT),
        ("FCF conversion", "% of EBITDA", "16 Cash Flow Model", "fcfconv", FMT_PCT_1),
        ("Net debt", "Rs cr", "17 Capital Allocation", "nd", FMT_INT),
        ("Net debt / EBITDA", "x", "17 Capital Allocation", "nd_ebitda", FMT_MULT),
        ("TRADE", "", None, None, None),
        ("Finished steel imports", "Mt", "19 Trade Model", "imp", FMT_2DP),
        ("Finished steel exports", "Mt", "19 Trade Model", "exp", FMT_2DP),
        ("Net trade position", "Mt", "19 Trade Model", "nettrade", FMT_2DP),
        ("Import penetration", "% of consumption", "19 Trade Model", "imppen", FMT_PCT_1),
        ("RETURNS", "", None, None, None),
        ("ROIC (indicative)", "%", "23 Comparable Valuation", "roic", FMT_PCT_1),
        ("ROIC less WACC", "ppt", "23 Comparable Valuation", "spread", FMT_PCT_1),
    ]:
        if sheet is None:
            r = sect(ws, r, lbl)
            continue
        base = "=IF(ISBLANK('%s'!%s$%d),\"n/a\",'%s'!%s$%d)" % (
            sheet, GL(CB), R(sheet, key), sheet, GL(CB), R(sheet, key))
        r = line(ws, r, lbl, unit, base,
                 ["='%s'!%s$%d" % (sheet, GL(c), R(sheet, key)) for c in FCOLS],
                 sheet.split()[0], "", "", fmt, cagr=(fmt in (FMT_INT, FMT_1DP, FMT_2DP)))
    footnote(ws, r + 1, 2, "Change the scenario on 01 Control Panel and this page re-states "
                           "entirely. Read alongside 22 Sensitivity section E, which states how "
                           "much of this rests on unsourced inputs.")


def sh_audit(ws):
    r = meta(ws, "25", "Audit Checks",
             "Thirty live validations. Every one is a formula, so they re-run automatically "
             "whenever a driver or the scenario changes.",
             "Checks are grouped: structural integrity, identity reconciliation, economic "
             "plausibility, cross-sheet consistency, and known model limitations. The limitations "
             "block is not a set of tests - it is a standing disclosure of what this model does "
             "NOT do, which is as important to a reviewer as what it does.",
             "PASS, FAIL or WARN.", "Reads every sheet.", 100)
    set_widths(ws, [(1, 7), (2, 13), (3, 56), (4, 14), (5, 100)])
    r = header_row(ws, r, ["", "Check ID", "Validation", "Result", "What it proves / what to do "
                                                                  "if it fails"],
                   left_align_cols=(3, 5))
    first = r
    F1, FN = GL(CF1), GL(CFN)

    def rng(sheet, key):
        return "'%s'!%s$%d:%s$%d" % (sheet, F1, R(sheet, key), FN, R(sheet, key))

    checks = [
        ("A01", "No negative capacity in any forecast year",
         '=IF(MIN(%s)>0,"PASS","FAIL")' % rng("06 Capacity Forecast", "cap"),
         "Capacity must be strictly positive. A failure means the delivery factor or secondary "
         "additions driver has been set to an impossible value."),
        ("A02", "Capacity is monotonically non-decreasing",
         '=IF(SUMPRODUCT(--(%s<0))=0,"PASS","FAIL")'
         % ("'06 Capacity Forecast'!%s$%d:%s$%d" % (GL(CF1), R("06 Capacity Forecast", "add_tot"),
                                                    GL(CFN), R("06 Capacity Forecast", "add_tot"))),
         "The model does not permit capacity closures. If a scenario should include closures, "
         "add them as negative entries in the project tracker."),
        ("A03", "Capacity utilisation never exceeds 100%",
         '=IF(MAX(%s)<=1,"PASS","FAIL")' % rng("08 Capacity Utilisation", "util"),
         "Physically impossible above 100%. Guarded by the ceiling on 05, so a failure indicates "
         "the ceiling has been bypassed."),
        ("A04", "Capacity utilisation never exceeds the practical maximum",
         '=IF(MAX(%s)<=MAX(%s)+0.0001,"PASS","FAIL")'
         % (rng("08 Capacity Utilisation", "util"), rng("05 Steel Supply Model", "utilmax")),
         "Confirms the MIN() ceiling on 05 Steel Supply is actually binding where it should."),
        ("A05", "Capacity utilisation is above a plausible floor of 55%",
         '=IF(MIN(%s)>=0.55,"PASS","WARN")' % rng("08 Capacity Utilisation", "util"),
         "Below 55% would imply mass idling that the model does not represent. A WARN means the "
         "capacity path is far ahead of the demand path."),
        ("A06", "Consumption is positive and growing in every year",
         '=IF(AND(MIN(%s)>0,SUMPRODUCT(--(%s<0))=0),"PASS","WARN")'
         % (rng("04 Steel Demand Model", "cons"), rng("04 Steel Demand Model", "dgrowth")),
         "Indian consumption has contracted only once in twelve years (FY2021, -5.27%). A WARN "
         "flags a scenario assuming contraction, which needs justification."),
        ("A07", "SUPPLY IDENTITY: production equals consumption plus net exports plus stock change",
         '=IF(SUMPRODUCT(ABS(%s-(%s+%s+%s)))<0.05,"PASS","FAIL")'
         % (rng("05 Steel Supply Model", "prod_req"), rng("05 Steel Supply Model", "cons"),
            rng("05 Steel Supply Model", "netexp"), rng("05 Steel Supply Model", "stock")),
         "THE MOST IMPORTANT CHECK ON THE SHEET. This is the Joint Plant Committee's own identity. "
         "A failure means the demand and trade blocks have drifted apart and every downstream "
         "number is wrong."),
        ("A08", "FY2026A supply identity reconciles to reported actuals",
         '=IF(ABS(%.2f-(%.2f+%.3f+%.2f))<0.01,"PASS","FAIL")'
         % (C.PROD26, C.CONS26, 0.078, -2.88),
         "163.74 + 0.078 - 2.88 = 160.94, which equals reported FY2026 finished steel production. "
         "Proves the base year is internally consistent before any forecasting begins."),
        ("A09", "Crude steel production is at least finished steel production",
         '=IF(SUMPRODUCT(--(%s<%s))=0,"PASS","FAIL")'
         % (rng("05 Steel Supply Model", "crude"), rng("05 Steel Supply Model", "prod")),
         "Yield loss means crude must exceed finished. A failure means the crude-to-finished "
         "ratio has been set below 1.0."),
        ("A10", "MARGIN BRIDGE: realisation change plus cost change equals EBITDA/t change",
         '=IF(SUMPRODUCT(ABS(%s-(%s-%.0f)))<5,"PASS","FAIL")'
         % (rng("14 Margin Analysis", "d_ebt"), rng("14 Margin Analysis", "ebt"), C.EBT26),
         "Proves the margin decomposition on 14 is arithmetically complete - no unexplained "
         "residual between the price and cost effects."),
        ("A11", "EBITDA equals production times EBITDA per tonne divided by 10",
         '=IF(SUMPRODUCT(ABS(%s-%s*%s/10))<1,"PASS","FAIL")'
         % (rng("13 EBITDA Model", "ebitda"), rng("13 EBITDA Model", "prod"),
            rng("13 EBITDA Model", "ebt")),
         "Unit-consistency check on the Rs/t to Rs crore bridge. This is where factor-of-ten "
         "errors hide."),
        ("A12", "Revenue equals production times realisation divided by 10",
         '=IF(SUMPRODUCT(ABS(%s-%s*%s/10))<1,"PASS","FAIL")'
         % (rng("12 Revenue Forecast", "rev"), rng("12 Revenue Forecast", "prod"),
            rng("12 Revenue Forecast", "realn")),
         "Same unit bridge on the revenue line."),
        ("A13", "EBITDA per tonne equals realisation less cash cost",
         '=IF(SUMPRODUCT(ABS(%s-(%s-%s)))<1,"PASS","FAIL")'
         % (rng("13 EBITDA Model", "ebt"), rng("13 EBITDA Model", "realn"),
            rng("13 EBITDA Model", "cost")),
         "Confirms EBITDA is built as a spread, not as a margin applied to revenue."),
        ("A14", "Industry EBITDA margin lies within the observed envelope of -7% to +26%",
         '=IF(AND(MIN(%s)>=-0.07,MAX(%s)<=0.26),"PASS","WARN")'
         % (rng("14 Margin Analysis", "margin"), rng("14 Margin Analysis", "margin")),
         "The envelope is the FY2022 observed peak and the SAIL FY2016 worst observation. A WARN "
         "is expected in the Stress Case, which deliberately breaches the floor at -9.2%. Outside "
         "the envelope is not automatically wrong, but it must be justified."),
        ("A15", "Cash cost is positive in every year",
         '=IF(MIN(%s)>0,"PASS","FAIL")' % rng("11 Cost Curve", "cost"),
         "A non-positive cash cost is impossible."),
        ("A16", "Realisation is positive in every year",
         '=IF(MIN(%s)>0,"PASS","FAIL")' % rng("09 Steel Price Forecast", "realn"),
         "Guards against a price adjustment factor driving realisation to or below zero."),
        ("A17", "Price adjustment factor stays within a plausible 0.85x to 1.15x band",
         '=IF(AND(MIN(%s)>=0.85,MAX(%s)<=1.15),"PASS","WARN")'
         % (rng("09 Steel Price Forecast", "px_adj"), rng("09 Steel Price Forecast", "px_adj")),
         "The utilisation feedback should modulate price, not dominate it. A WARN means the "
         "elasticity D23 or the utilisation path is extreme."),
        ("A18", "Cash tax is never negative",
         '=IF(MIN(%s)>=0,"PASS","FAIL")' % rng("16 Cash Flow Model", "tax"),
         "Tax is floored at zero on negative EBIT. No loss carry-forward is modelled, which is "
         "conservative and disclosed."),
        ("A19", "Net working capital is positive in every year",
         '=IF(MIN(%s)>0,"PASS","FAIL")' % rng("15 Working Capital Model", "nwc"),
         "Steel producers carry positive working capital. Negative would indicate a sign error in "
         "the days driver."),
        ("A20", "WACC used downstream matches the WACC build-up on 02 Section D",
         '=IF(ABS(%s-\'02 Model Assumptions\'!$D$%d)<0.005,"PASS","WARN")'
         % (dv("D20", CF1).replace("=", ""), R("02 Model Assumptions", "wacc_calc")),
         "Driver D20 is what the model discounts at; Section D is the transparent build-up. If "
         "they diverge, one of them has been edited without the other - reconcile before using "
         "any valuation output."),
        ("A21", "Terminal value is less than 75% of enterprise value",
         '=IF(IFERROR(\'23 Comparable Valuation\'!$D$%d,0)<0.75,"PASS","WARN")'
         % R("23 Comparable Valuation", "dcf_tvpct"),
         "Above 75% the valuation is essentially an assertion about the exit multiple - which is "
         "UNSOURCED - rather than a forecast. Shorten the horizon or re-base the multiple."),
        ("A22", "Discount factors are strictly decreasing",
         '=IF(SUMPRODUCT(--(\'23 Comparable Valuation\'!%s$%d:%s$%d>'
         '\'23 Comparable Valuation\'!%s$%d:%s$%d))=%d,"PASS","FAIL")'
         % (F1, R("23 Comparable Valuation", "df"), GL(CFN - 1),
            R("23 Comparable Valuation", "df"), GL(CF1 + 1),
            R("23 Comparable Valuation", "df"), FN, R("23 Comparable Valuation", "df"), NY - 1),
         "Confirms the mid-year discounting convention is applied consistently."),
        ("A23", "Imports and exports are both non-negative",
         '=IF(AND(MIN(%s)>=0,MIN(%s)>=0),"PASS","FAIL")'
         % (rng("19 Trade Model", "imp"), rng("19 Trade Model", "exp")),
         "A negative export volume would mean the net export driver has overwhelmed the import "
         "driver - the two must be set coherently."),
        ("A24", "Trade block reconciles to the supply balance",
         '=IF(SUMPRODUCT(ABS((%s-%s)-%s))<0.05,"PASS","FAIL")'
         % (rng("19 Trade Model", "exp"), rng("19 Trade Model", "imp_drv"),
            rng("19 Trade Model", "netexp")),
         "Exports less the import driver must equal net exports by construction. Proves 19 has "
         "not drifted from 05."),
        ("A25", "Import penetration stays below 20% of consumption",
         '=IF(MAX(%s)<0.20,"PASS","WARN")' % rng("19 Trade Model", "imppen"),
         "The historical maximum in the database is 6.3% (FY2025). Above 20% would imply a "
         "collapse in domestic competitiveness that the price deck does not reflect."),
        ("A26", "FY2031 capacity is within 10% of the National Steel Policy 300 Mtpa target",
         '=IF(ABS(\'06 Capacity Forecast\'!%s$%d-1)<0.10,"PASS","WARN")'
         % (GL(FCOLS[4]), R("06 Capacity Forecast", "nsp_pct")),
         "An external validation, not a constraint. The capacity block was built bottom-up and "
         "was not calibrated to the policy target, so close agreement is evidence the build is "
         "sensible. A WARN in the Bear or Stress case is expected and correct."),
        ("A27", "Scenario tie-out: live workbook matches the independent Python mirror",
         "='21 Scenario Manager'!$I$%d" % R("21 Scenario Manager", "tie_summary"),
         "THE REGRESSION TEST. Compares eleven live FY2033 outputs against values computed "
         "independently at build time by fmodel/chain.py. A failure means an Excel formula does "
         "not implement the intended arithmetic - stop and investigate before using any output."),
        ("A28", "Every active driver resolves to a number for the selected scenario",
         '=IF(COUNT(\'02 Model Assumptions\'!%s$%d:%s$%d)=%d,"PASS","FAIL")'
         % (F1, R("02 Model Assumptions", "act_D01"), FN,
            R("02 Model Assumptions", "act_D26"),
            (R("02 Model Assumptions", "act_D26") - R("02 Model Assumptions", "act_D01") + 1) * NY),
         "Confirms every INDEX lookup against ScenID has resolved. A failure means a driver row "
         "is missing a scenario or ScenID is out of range."),
        ("A29", "ScenID resolves to a valid index between 1 and 4",
         '=IF(AND(ScenID>=1,ScenID<=4),"PASS","FAIL")',
         "If the scenario cell is edited to free text outside the four permitted values, MATCH "
         "returns an error and the entire workbook fails. Data validation prevents this, but the "
         "check is retained as a backstop."),
        ("A30", "No forecast column is entirely blank",
         '=IF(COUNT(\'13 EBITDA Model\'!%s$%d:%s$%d)=%d,"PASS","FAIL")'
         % (F1, R("13 EBITDA Model", "ebitda"), FN, R("13 EBITDA Model", "ebitda"), NY),
         "Catches a broken link or a deleted column in the core EBITDA line."),
    ]
    for cid, _, f, note in [(c[0], c[1], c[2], c[3]) for c in checks]:
        put(ws, r, 2, cid, alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 3, [c[1] for c in checks if c[0] == cid][0], alignment=AL_LEFT_WRAP)
        put(ws, r, 4, f, alignment=AL_CENTRE, f=font(10, bold=True))
        put(ws, r, 5, note, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        r += 1
    last = r - 1
    add_table(ws, "tbl_Audit", first - 1, 1, last, 5)
    put(ws, r, 3, "OVERALL", f=font(11, bold=True, colour=NAVY))
    put(ws, r, 4, '=IF(COUNTIF(D%d:D%d,"FAIL")>0,"FAIL",IF(COUNTIF(D%d:D%d,"WARN")>0,'
                  '"PASS WITH WARNINGS","ALL PASS"))' % (first, last, first, last),
        alignment=AL_CENTRE, f=font(11, bold=True, colour=NAVY))
    put(ws, r, 5, "WARN results are expected in the Bear and Stress cases: A14 (margin envelope) "
                  "and A26 (policy target) are deliberately breached by design in those "
                  "scenarios. Any FAIL must be investigated before the model is used.",
        f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    r += 3
    r = sect(ws, r, "KNOWN LIMITATIONS - STANDING DISCLOSURE, NOT TESTS")
    for t in [
        "1. NO ENDOGENOUS SUPPLY RESPONSE. The model does not idle capacity or cut production when "
        "margins turn negative, which in reality would arrest a decline. The Stress Case must "
        "therefore be read as a lower bound on margin, not a central expectation.",
        "2. NO CIRCULARITY BETWEEN PRICE AND CURRENT-YEAR UTILISATION. The price feedback uses "
        "PRIOR-year utilisation. This is a deliberate simplification to avoid an iterative solve; "
        "it slightly dampens the cycle relative to reality.",
        "3. COKING COAL HAS NO FISCAL-YEAR AVERAGE. The FY2026 reference of US$225/t is a Mar-2026 "
        "SPOT observation. The database asserts no average because the two public series disagree "
        "by up to US$40/t. A licensed price-reporting-agency series is the single highest-priority "
        "data gap.",
        "4. WORKING CAPITAL IS NOT DERIVED FROM BALANCE SHEETS. The database does not carry "
        "receivables, inventory or payables. 45 days is a sector convention. Fix before any "
        "financing or credit work.",
        "5. NO COMPANY-SPECIFIC COST CURVES BELOW THE FOUR MAJORS. Cost positions for Jindal "
        "Stainless, AM/NS India and RINL are not built, and the secondary sector - roughly half "
        "of India's output - has no cost representation at all.",
        "6. SHARE PRICES ARE NOT SOURCED, so no trading multiple is asserted. The comparables "
        "framework computes the moment prices are entered.",
        "7. INVESTED CAPITAL IS A REPLACEMENT-COST PROXY, so ROIC is indicative only and "
        "understates returns on older depreciated assets.",
        "8. SINGLE-COUNTRY SCOPE. Tata Steel's Netherlands and UK operations, including the "
        "disclosed going-concern uncertainty at Tata Steel Netherlands, are entirely outside this "
        "India model and must be handled in the single-name build.",
    ]:
        put(ws, r, 3, t, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        ws.row_dimensions[r].height = 13 * (1 + len(t) // 95)
        r += 1


def sh_sources(ws):
    r = meta(ws, "26", "Sources",
             "Sources used for the forward-looking assumptions in this model. FY2026 actuals are "
             "sourced in Master Industry Database.xlsx, whose Sources sheet carries 47 entries "
             "and is not duplicated here.",
             "Level 1 primary issuer and government, Level 2 official statistical and multilateral "
             "agencies, Level 3 rating agencies and sell-side. Secondary sources never override "
             "primary data.",
             "Not applicable.", "Referenced by 02 Model Assumptions.", 100)
    hs = ["", "Ref", "Level", "Publisher", "Publication", "Date", "Used for", "Verification",
          "Notes"]
    set_widths(ws, [(1, 7), (2, 15), (3, 8), (4, 34), (5, 52), (6, 16), (7, 44), (8, 40),
                    (9, 100)])
    r = header_row(ws, r, hs, left_align_cols=(2, 4, 5, 7, 8, 9))
    first = r
    srcs = [
        ("RBI-Jun26", 2, "Reserve Bank of India",
         "Monetary Policy Committee statement", "05-Jun-2026",
         "Drivers D01 real GDP growth and D04 CPI inflation; assumption register R01",
         "Verified via multiple contemporaneous reports of the MPC statement",
         "FY2027 real GDP projection cut to 6.6% from 6.9%; repo held at 5.25%; FY2027 CPI "
         "projection raised 50bps to 5.1%. FY2026 growth estimated 7.6%; Q4FY2026 actual 7.7%."),
        ("worldsteel-SRO", 2, "World Steel Association",
         "Short Range Outlook, April 2026", "Apr-2026",
         "Corroborates driver D02 steel demand elasticity",
         "Verified - also reproduced in Ministry of Steel monthly overviews",
         "Global demand +0.3% in CY2026 to 1,724 Mt and +2.2% in CY2027 to 1,762 Mt. India +7.4% "
         "and +9.2%, the fastest major market. China -1.5% then flat."),
        ("GSEC-31Jul26", 3, "investing.com / NSE",
         "India 10-year benchmark government bond yield", "31-Jul-2026",
         "Risk-free rate in the WACC build-up; assumption register R08",
         "Verified - 6.833%; Trading Economics expects 6.78% at quarter end",
         "Market data. A Bloomberg or Refinitiv print should be substituted for a live mandate."),
        ("CRISIL-RateView", 3, "CRISIL Intelligence",
         "RateView, July 2026", "Jul-2026",
         "Pre-tax cost of debt in the WACC build-up",
         "Verified - 10-year corporate bond yield 7.36% actual, 7.44-7.54% projected",
         "Consistent with the ICRA and India Ratings AA domestic ratings the database records for "
         "JSW Steel."),
        ("IDBI-22Jun26", 3, "IDBI Capital Markets & Securities",
         "Commodity Price Update, week ended 22-Jun-2026", "22-Jun-2026",
         "Driver D03 USD/INR; input cost cross-checks",
         "Verified - full document retrieved and parsed",
         "USD/INR 94 with a 52-week range of 85 to 98. Also billet ex-Raipur Rs 38,850/t, HR "
         "strip Patra Rs 44,500/t, nickel US$17,395/t, ferro chrome Rs 1,23,200/t, graphite "
         "electrode UHP US$4,189/t, Brent US$79/bbl - the input-cost deck behind D12."),
        ("MasterDB", 1, "This research team",
         "Master Industry Database.xlsx", "03-Aug-2026",
         "Every FY2026 actual and every derived calibration anchor",
         "Values extracted programmatically from the database file at build time",
         "47 underlying sources with full citations on that workbook's Sources sheet, and 16 "
         "documented conflicts on its Conflicts Log. This model does not restate or contradict "
         "any of it."),
        ("Derived-DB", 1, "This research team",
         "Derivations from Master Industry Database", "03-Aug-2026",
         "Drivers D02, D05, D06, D07, D08, D09, D16, D22, D24, D26 and anchors A08 to A13",
         "Each derivation is stated in full on the relevant row",
         "The material derivations are: blended realisation of Rs 59,974/t from four majors; "
         "weighted EBITDA of Rs 10,733/t; implied cash cost of Rs 49,241/t; FY2026 average "
         "USD/INR of 88.4 from Tata Steel's dual-currency disclosure; D&A at 5.28% of revenue; "
         "demand elasticity of 1.10x from the FY2015-FY2026 outturn."),
        ("Indicative", 3, "Modeller judgement",
         "Structural parameters not sourced from any document", "n/a",
         "Drivers D13, D14, D17, D23, D25 and the ESG inputs on sheet 20",
         "NOT VERIFIED - explicitly flagged Low confidence throughout",
         "Held constant across scenarios wherever unsourced, so that an unsourced parameter is "
         "never flexed to manufacture a scenario. Each is quantified on 22 Sensitivity section E."),
        ("Not sourced", 3, "None",
         "Inputs left as user inputs rather than assumed", "n/a",
         "Driver D18 working capital days, D19 exit multiple, share prices on 23, equity risk "
         "premium and beta on 02 Section D",
         "NOT SOURCED - shaded in the workbook and excluded from headline conclusions",
         "These are the four gaps that must be closed before transaction use. None has been "
         "filled with a fabricated value."),
    ]
    for ref, lvl, pub, pubn, dt, used, ver, note in srcs:
        put(ws, r, 2, ref, alignment=AL_LEFT, f=font(10, bold=True))
        put(ws, r, 3, lvl, alignment=AL_CENTRE)
        put(ws, r, 4, pub, alignment=AL_LEFT_WRAP)
        put(ws, r, 5, pubn, alignment=AL_LEFT_WRAP, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 6, dt, alignment=AL_CENTRE, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 7, used, alignment=AL_LEFT_WRAP, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 8, ver, alignment=AL_LEFT_WRAP, f=font(9, colour=TEXT_MUTED))
        put(ws, r, 9, note, alignment=AL_LEFT_WRAP, f=font(9, colour=TEXT_MUTED))
        r += 1
    add_table(ws, "tbl_ModelSources", first - 1, 1, r - 1, 9)
    r += 2
    r = sect(ws, r, "MODEL CHANGE LOG")
    r = header_row(ws, r, ["", "Version", "Date", "Change", "Author"],
                   left_align_cols=(4,))
    put(ws, r, 2, "1.0", alignment=AL_CENTRE)
    put(ws, r, 3, "03-Aug-2026", alignment=AL_CENTRE)
    put(ws, r, 4, "Initial build. Horizon FY2027E-FY2033E selected on the weighted analysis on "
                  "0H. Four scenarios calibrated to observed margin envelopes. Lagged "
                  "price-utilisation feedback introduced so that capacity transmits to earnings.",
        alignment=AL_LEFT_WRAP, f=font(9, colour=TEXT_MUTED))
    put(ws, r, 5, "Head of Metals & Mining Research", alignment=AL_LEFT_WRAP,
        f=font(9, colour=TEXT_MUTED))


# ======================================================================================
PLAN = [
    ("", "Cover"), ("", "Contents"), ("", "00 Database Import"), ("", "0H Forecast Horizon"),
    ("", "01 Control Panel"), ("", "02 Model Assumptions"), ("", "03 Macroeconomic Model"),
    ("", "04 Steel Demand Model"), ("", "05 Steel Supply Model"), ("", "06 Capacity Forecast"),
    ("", "07 Capacity Expansion Tracker"), ("", "08 Capacity Utilisation"),
    ("", "09 Steel Price Forecast"), ("", "10 Raw Material Forecast"), ("", "11 Cost Curve"),
    ("", "12 Revenue Forecast"), ("", "13 EBITDA Model"), ("", "14 Margin Analysis"),
    ("", "15 Working Capital Model"), ("", "16 Cash Flow Model"), ("", "17 Capital Allocation"),
    ("", "18 Industry Cycle Model"), ("", "19 Trade Model"), ("", "20 ESG Model"),
    ("", "21 Scenario Manager"), ("", "22 Sensitivity Analysis"), ("", "23 Comparable Valuation"),
    ("", "24 Industry Dashboard"), ("", "25 Audit Checks"), ("", "26 Sources"),
]

CONTENTS = [
    ("00", "00 Database Import", "The controlled boundary with Master Industry Database.xlsx: 18 "
     "FY2026 anchors with exact provenance and a documented refresh procedure.", "Input"),
    ("0H", "0H Forecast Horizon", "Weighted evaluation of 5, 7 and 10-year horizons; 7 years "
     "selected. Live weighted-score formulas.", "Documentation"),
    ("01", "01 Control Panel", "The scenario selector - the only cell a user changes - plus a "
     "live output panel.", "Control"),
    ("02", "02 Model Assumptions", "26 drivers on 4 scenarios, the active resolved block, the "
     "13-field assumption register and the WACC build-up.", "Input"),
    ("03", "03 Macroeconomic Model", "Real GDP, CPI, USD/INR, G-sec yield and WACC.", "Calc"),
    ("04", "04 Steel Demand Model", "Consumption via GDP times a history-calibrated elasticity; "
     "per capita intensity.", "Calc"),
    ("05", "05 Steel Supply Model", "The JPC supply identity, the capacity ceiling and the "
     "shortfall routed to imports.", "Calc"),
    ("06", "06 Capacity Forecast", "Capacity from the project tracker plus a secondary residual; "
     "validated against the 300 Mtpa policy target.", "Calc"),
    ("07", "07 Capacity Expansion Tracker", "20 dated projects with live SUMIFS feeding 06; "
     "encodes the acquisition, downstream and overseas exclusions.", "Input"),
    ("08", "08 Capacity Utilisation", "Utilisation and the utilisation gap that drives the price "
     "feedback.", "Calc"),
    ("09", "09 Steel Price Forecast", "Blended realisation with the lagged utilisation feedback; "
     "safeguard duty schedule.", "Calc"),
    ("10", "10 Raw Material Forecast", "Iron ore, coking coal, FX and the raw material cost "
     "index.", "Calc"),
    ("11", "11 Cost Curve", "Industry cash cost, calibrated top-down, and the four-producer cost "
     "curve.", "Calc"),
    ("12", "12 Revenue Forecast", "Industry and company revenue with an explicit unit bridge.",
     "Calc"),
    ("13", "13 EBITDA Model", "EBITDA as a price-cost spread, industry and company.", "Calc"),
    ("14", "14 Margin Analysis", "Margin, the price-cost bridge and observed credibility "
     "benchmarks.", "Calc"),
    ("15", "15 Working Capital Model", "NWC on days of revenue. The weakest block - flagged.",
     "Calc"),
    ("16", "16 Cash Flow Model", "Unlevered FCF: EBITDA less cash tax, capex and the NWC change.",
     "Calc"),
    ("17", "17 Capital Allocation", "Net debt path, leverage and covenant headroom.", "Calc"),
    ("18", "18 Industry Cycle Model", "Cycle position on utilisation and margin gaps.", "Calc"),
    ("19", "19 Trade Model", "Imports, exports, penetration and export intensity.", "Calc"),
    ("20", "20 ESG Model", "Green Steel Taxonomy thresholds and a live CBAM exposure "
     "calculation.", "Calc"),
    ("21", "21 Scenario Manager", "All four scenarios side by side plus an 11-metric tie-out "
     "against the independent Python mirror.", "Output"),
    ("22", "22 Sensitivity Analysis", "Tornado, two LIVE two-variable tables, and a block "
     "quantifying every unsourced assumption.", "Output"),
    ("23", "23 Comparable Valuation", "Trading comps (share price = user input), EV per tonne, "
     "replacement cost, industry DCF and ROIC.", "Output"),
    ("24", "24 Industry Dashboard", "One-page live summary of the active scenario.", "Output"),
    ("25", "25 Audit Checks", "30 live validations plus a standing disclosure of 8 known "
     "limitations.", "Check"),
    ("26", "26 Sources", "Sources for the forward-looking assumptions, plus the model change "
     "log.", "Documentation"),
]


def main():
    if not os.path.exists(DB):
        print("WARNING: %s not found in the working directory. Import provenance still written, "
              "but values could not be re-verified against the database file." % DB)
    wb = Workbook()
    wb.remove(wb.active)
    made = {}
    for _, nm in PLAN:
        made[nm] = wb.create_sheet(title=nm)
    for nm in made:
        sheet_defaults(made[nm])

    sh_import(made["00 Database Import"])
    sh_horizon(made["0H Forecast Horizon"])
    # assumptions must precede any sheet that references active driver rows
    sh_assump(made["02 Model Assumptions"])
    sh_macro(made["03 Macroeconomic Model"])
    sh_demand(made["04 Steel Demand Model"])
    sh_tracker(made["07 Capacity Expansion Tracker"])
    sh_capacity(made["06 Capacity Forecast"])
    sh_supply(made["05 Steel Supply Model"])
    sh_util(made["08 Capacity Utilisation"])
    sh_price(made["09 Steel Price Forecast"])
    sh_rawmat(made["10 Raw Material Forecast"])
    sh_cost(made["11 Cost Curve"])
    sh_revenue(made["12 Revenue Forecast"])
    sh_ebitda(made["13 EBITDA Model"])
    sh_margin(made["14 Margin Analysis"])
    sh_wc(made["15 Working Capital Model"])
    sh_cf(made["16 Cash Flow Model"])
    sh_capalloc(made["17 Capital Allocation"])
    sh_cycle(made["18 Industry Cycle Model"])
    sh_trade(made["19 Trade Model"])
    sh_esg(made["20 ESG Model"])
    sel, sid = sh_control(made["01 Control Panel"])
    _K(made["01 Control Panel"], "sel", sel)
    sh_scenario(made["21 Scenario Manager"])
    sh_comps(made["23 Comparable Valuation"])
    sh_sens(made["22 Sensitivity Analysis"])
    sh_dash(made["24 Industry Dashboard"])
    sh_audit(made["25 Audit Checks"])
    sh_sources(made["26 Sources"])
    sh_cover(made["Cover"])
    sh_contents(made["Contents"], CONTENTS)

    wb.defined_names["ScenID"] = DefinedName(
        "ScenID", attr_text="'01 Control Panel'!$D$%d" % sid)
    wb.defined_names["ScenarioName"] = DefinedName(
        "ScenarioName", attr_text="'01 Control Panel'!$D$%d" % sel)

    wb.properties.title = "Industry Financial Model - Indian Steel Industry"
    wb.properties.category = "Global Metals & Mining Research"
    wb.properties.description = (
        "Analytical engine for the Indian steel industry. Horizon FY2027E-FY2033E. Consumes "
        "Master Industry Database.xlsx. Four integrated scenarios driven by a single selector.")
    made["Cover"].sheet_view.tabSelected = True
    wb.active = 0
    wb.save(OUT)
    print("WROTE %s  (%.1f KB, %d sheets)"
          % (OUT, os.path.getsize(OUT) / 1024.0, len(wb.sheetnames)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
