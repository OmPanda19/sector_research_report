# -*- coding: utf-8 -*-
"""Sheet writers for Master Industry Database.xlsx."""
from openpyxl.utils import get_column_letter

from . import data as D
from . import data_market as M
from . import sources as S
from .style import (AL_CENTRE, AL_LEFT, AL_LEFT_WRAP, AL_RIGHT, BORDER_BOTTOM,
                    BORDER_TOP_RULE, FILL_NAVY_LIGHT, FMT_1DP, FMT_2DP, FMT_3DP, FMT_INT,
                    FMT_MULT, FMT_PCT_1, FMT_TEXT, NA, NAVY, TEXT_MUTED, add_table, font,
                    footnote, freeze, header_row, put, repeat_header, set_widths,
                    sheet_defaults, title_block)

CUTOFF = D.CUTOFF
WIDE, MED, NARROW, NOTE_W = 34, 18, 12, 90


def _num_fmt(header):
    """Infer a number format from a column header. Keeps units and decimals aligned."""
    h = header.lower()
    if "%" in h or "margin" in h or "cagr" in h or "share" in h or "utilisation" in h \
            or "growth" in h or "penetration" in h or "intensity" in h or "rate in force" in h:
        return FMT_PCT_1
    if "rs crore" in h or "rs cr" in h or "crore" in h or "us$ m" in h or "'000 t" in h \
            or "rs/t" in h or "rs/tonne" in h or "number" in h:
        return FMT_INT
    if "us$/" in h or "kg" in h or "tco2e" in h:
        return FMT_2DP
    if "mn t" in h or "(mt)" in h or "mtpa" in h:
        return FMT_3DP
    if "x)" in h:
        return FMT_MULT
    return FMT_1DP


def block(ws, row, headers, rows, table_name, *, col_fmts=None, left_cols=(1,),
          note_cols=(), start_col=1):
    """Write a header band plus data rows, wrap it in an Excel Table, return next row."""
    hrow = row
    row = header_row(ws, row, headers, start_col=start_col,
                     left_align_cols=tuple(start_col + c - 1 for c in left_cols))
    first_data = row
    for r in rows:
        for i, val in enumerate(r):
            col = start_col + i
            hdr = headers[i] if i < len(headers) else ""
            if val is None:
                put(ws, row, col, None)
                continue
            if isinstance(val, str):
                al = AL_LEFT_WRAP if (i + 1) in note_cols else AL_LEFT
                f = font(9, colour=TEXT_MUTED) if (i + 1) in note_cols else font()
                put(ws, row, col, val, f=f, alignment=al, number_format=FMT_TEXT)
            elif isinstance(val, str) or not isinstance(val, (int, float)):
                put(ws, row, col, val, number_format=FMT_TEXT)
            else:
                nf = (col_fmts or {}).get(i + 1) or _num_fmt(hdr)
                put(ws, row, col, val, number_format=nf, alignment=AL_RIGHT)
        row += 1
    add_table(ws, table_name, hrow, start_col, row - 1, start_col + len(headers) - 1)
    freeze(ws, "%s%d" % (get_column_letter(start_col + max(left_cols)), first_data))
    repeat_header(ws, hrow)
    return row + 1


def meta(ws, sheet_no, name, objective, columns_note, units, coverage, frequency,
         hierarchy, extra=()):
    lines = [
        ("%s. %s" % (sheet_no, name), "title"),
        ("Master Industry Database  |  Indian Steel Industry  |  Data cutoff: %s" % CUTOFF,
         "label"),
        ("", "body"),
        ("OBJECTIVE", "label"), (objective, "body"),
        ("UNITS AND CURRENCY", "label"), (units, "body"),
        ("COVERAGE AND HISTORICAL DEPTH", "label"), (coverage, "body"),
        ("DATA FREQUENCY", "label"), (frequency, "body"),
        ("SOURCE HIERARCHY APPLIED", "label"), (hierarchy, "body"),
        ("COLUMN RATIONALE", "label"), (columns_note, "body"),
    ]
    for e in extra:
        lines.append(("", "body"))
        lines.append((e, "body"))
    lines.append(("", "body"))
    return title_block(ws, lines)


HIER_GOV = ("Tier 1 only: Ministry of Steel and Joint Plant Committee publications, Parliamentary "
            "answers, and the World Steel Association as reproduced by the Ministry. No secondary "
            "source is relied upon for any figure on this sheet.")
HIER_CO = ("Tier 1 first: issuer results releases, exchange-filed investor presentations, audited "
           "financial results and exchange-filed earnings call transcripts. Tier 3 aggregators of "
           "exchange filings are used only to extend the series backwards and every FY2026 figure "
           "so sourced has been reconciled to a Tier 1 or Tier 2 source or to the sum of the four "
           "reported quarters.")
HIER_PX = ("Tier 1 first: World Bank Commodity Price Data and Ministry of Steel published "
           "assessments. Tier 2 rating agency and sell-side trackers are used only where no "
           "official series exists, and are labelled with the assessor on every row.")


# ======================================================================================
def cover(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 3), (2, 118)])
    r = title_block(ws, [
        ("MASTER INDUSTRY DATABASE", "title"),
        ("INDIAN STEEL INDUSTRY", "subtitle"),
        ("", "body"),
        ("Global Metals & Mining Research", "label"),
        ("Data cutoff: 03 August 2026  |  Fiscal year convention: FY2026 = 1 April 2025 to 31 March 2026",
         "body"),
        ("", "body"),
    ], start_row=2, width_col=2)

    sections = [
        ("PURPOSE",
         "Single source of truth for financial models, DCF and comparable valuation, sector "
         "dashboards, quarterly updates and initiation reports on the Indian steel industry. "
         "Structured so that every figure is traceable to one document and so that a new fiscal "
         "year is added by appending a single column or row, without restructuring."),
        ("COVERAGE",
         "15 companies. Core Coverage (7): Tata Steel, JSW Steel, SAIL, Jindal Steel, AM/NS India, "
         "Jindal Stainless, RINL. Supporting Coverage (8): Shyam Metalics, Godawari Power & Ispat, "
         "APL Apollo Tubes, Mukand, Gallantt, Uttam Galva Steels, Electrosteel Steels (ESL Steel), "
         "Kirloskar Ferrous Industries. Twelve are listed, one is an unlisted joint venture, one is "
         "an unlisted state-owned enterprise and two are delisted."),
        ("BASIS OF PREPARATION",
         "Consolidated financial statements are used throughout unless stated otherwise on the row. "
         "Where an issuer's own headline differs from the comparable statutory measure, BOTH are "
         "carried in separately labelled columns. EBITDA is presented on two bases because the peer "
         "set does not define it consistently: as reported by the company, and standardised as "
         "revenue from operations less operating expenses excluding other income, finance cost, "
         "depreciation and amortisation, exceptional items and share of profit of joint ventures."),
        ("DATA INTEGRITY STANDARD",
         "No figure has been estimated, interpolated or inferred. Where a figure is not in the "
         "public domain it is recorded as \"Data Not Publicly Available\" with an explanation, "
         "rather than filled. Figures this workbook has calculated are marked Methodology = "
         "\"Computed\" or \"Derived\"; figures lifted from a document are marked \"As reported\"."),
        ("CONFLICTS",
         "Sixteen conflicts and definitional traps were identified during research and are set out "
         "in full on the Conflicts Log, each with the competing figures, the reason for the "
         "difference, the resolution adopted and the residual risk. Three are material to "
         "modelling: C01 coking coal benchmark divergence of up to US$40/t; C02 the Jindal Steel "
         "revenue basis; C16 the AM/NS India fiscal year mismatch."),
        ("THE \"DATA NOT PUBLICLY AVAILABLE\" SENTINEL - IMPORTANT FOR DOWNSTREAM TOOLS",
         "Where a figure is not in the public domain the cell carries the text \"Data Not "
         "Publicly Available\" rather than being left blank, so that a human reader can see the "
         "difference between a gap that was investigated and a gap that was overlooked. The "
         "consequence is that some fiscal-year columns are mixed-type. This is safe in Excel - "
         "SUM, AVERAGE and the like ignore text, and every calculated column in this workbook is "
         "ISNUMBER-guarded inside IFERROR so no formula returns an error. When loading into Power "
         "BI or Power Query, replace \"Data Not Publicly Available\" with null before setting the "
         "column type to decimal. Do not simply force the type, which would silently produce "
         "errors."),
        ("KNOWN LIMITATIONS - READ BEFORE USE",
         "1. No fiscal-year average is asserted for coking coal; the two available public series "
         "disagree materially and neither is published as a complete monthly history. A licensed "
         "price-reporting-agency series is required before coking coal is used in a cost model. "
         "2. Sales volumes were not sourced for six of the eight Supporting Coverage names, so "
         "EBITDA per tonne is not computed for them. "
         "3. RINL and ESL Steel have no publicly available FY2026 statutory financials. "
         "4. Company capacity is not sourced for several Supporting names. "
         "5. Exchange scrip codes, ISINs and vendor identifiers are deliberately not asserted "
         "because they were not independently verified in this research cycle."),
        ("HOW TO NAVIGATE",
         "Use the Contents sheet. Every data sheet opens with a block stating its objective, units, "
         "coverage, frequency, source hierarchy and column rationale. Every data range is an Excel "
         "Table so that downstream models and Power Query bind to an auto-expanding range rather "
         "than a fixed one. Header rows are frozen and repeat on print."),
    ]
    for head, body in sections:
        put(ws, r, 2, head, f=font(10, bold=True, colour=NAVY))
        r += 1
        put(ws, r, 2, body, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        ws.row_dimensions[r].height = 14 * (1 + len(body) // 118)
        r += 2
    put(ws, r, 2,
        "Prepared from public sources only. Contains no non-public or price-sensitive information.",
        f=font(8, italic=True, colour=TEXT_MUTED))


def contents(ws, sheet_specs):
    sheet_defaults(ws)
    set_widths(ws, [(1, 6), (2, 30), (3, 96), (4, 14)])
    r = meta(ws, "", "CONTENTS",
             "Navigation index. Click a sheet name to jump to it.",
             "Sheet number, name, scope and the highest source tier relied upon.",
             "Not applicable.", "All 14 mandated worksheets plus Cover, Contents, Data Dictionary "
             "and Conflicts Log.", "Not applicable.", "Not applicable.")
    r = header_row(ws, r, ["#", "Worksheet", "Scope and primary content", "Top Source Tier"],
                   left_align_cols=(2, 3))
    for num, name, scope, tier in sheet_specs:
        put(ws, r, 1, num, alignment=AL_CENTRE)
        c = put(ws, r, 2, name, f=font(10, bold=True, colour=NAVY))
        c.hyperlink = "#'%s'!A1" % name
        put(ws, r, 3, scope, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
        put(ws, r, 4, tier, alignment=AL_CENTRE)
        r += 1
    freeze(ws, "A%d" % (r - len(sheet_specs)))


# ======================================================================================
def company_list(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 5), (2, 46), (3, 20), (4, 13), (5, 13), (6, 26), (7, 30), (8, 34),
                    (9, 46), (10, 34), (11, 32), (12, 11), (13, 11), (14, 26), (15, 24),
                    (16, 12), (17, 11), (18, NOTE_W)])
    r = meta(ws, 1, "Company List",
             "Defines the coverage universe and fixes, for each company, the identifiers, "
             "ownership, steelmaking route, assets, consolidation basis, reporting currency and "
             "latest reported period. Every other sheet inherits the basis stated here, so this "
             "sheet is the control record for the workbook.",
             "Coverage Tier separates the seven Core names, which carry a full financial build-out, "
             "from the eight Supporting names included for market structure and read-across. "
             "Primary Steelmaking Route is carried because route determines cost-curve position, "
             "coking coal versus scrap and gas exposure, and carbon intensity - the three largest "
             "valuation drivers in this sector. Consolidation Basis Used prevents the classic error "
             "of mixing standalone and consolidated figures across sheets. Latest Reported Period "
             "makes explicit where FY2026 was unavailable. Asset Detail Verification is stated "
             "because asset-level detail for the Supporting names is indicative and has not been "
             "traced to an annual report.",
             "Not applicable - descriptive sheet. Capacity figures are in Mtpa.",
             "15 companies: 7 Core Coverage and 8 Supporting Coverage. Point-in-time as at the data cutoff.",
             "Point-in-time; reviewed each quarter.", HIER_CO,
             extra=["IDENTIFIERS: NSE symbols are stated because they were verified. BSE scrip "
                    "codes, ISINs and Bloomberg or Refinitiv identifiers are deliberately NOT "
                    "asserted because they were not independently verified in this research cycle. "
                    "Populate them in the identifier-mapping step rather than inferring them."])
    r = block(ws, r, D.COMPANY_HEADERS, D.COMPANIES, "tbl_CompanyList",
              left_cols=(1, 2), note_cols=(18,),
              col_fmts={1: FMT_INT})
    footnote(ws, r, 1,
             "APL Apollo Tubes and Uttam Galva Steels are converters, not steelmakers, and are "
             "excluded by design from crude steel capacity, crude steel volume and EBITDA per "
             "tonne of crude steel comparisons. Kirloskar Ferrous is predominantly a pig iron, "
             "castings and tubes producer.")


def production_capacity(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 34), (2, 13), (3, 46), (4, 10), (5, 16), (6, 20), (7, 20), (8, 30),
                    (9, 18), (10, 14), (11, 22), (12, 11), (13, NOTE_W)])
    r = meta(ws, 2, "Production Capacity",
             "Establishes installed and pipeline capacity at industry and company level. Capacity "
             "is the denominator for utilisation, the anchor for growth-capex modelling and the "
             "measure against which the National Steel Policy 2017 target of 300 Mtpa is tracked.",
             "Capacity Definition is the most important column on this sheet: companies publish "
             "crude steel, liquid steel, hot metal, melt and aggregate installed metal capacity "
             "and these are not interchangeable. Approval Status separates in-place earnings "
             "capacity from optionality that still requires capital and time, so that a Board-"
             "approved tranche is never summed with an aspirational target.",
             "Mtpa (million tonnes per annum). Utilisation in per cent.",
             "Industry: FY2024-FY2026 plus the FY2031 policy target. Company: capacity at FY2026 "
             "year-end plus pipeline. All 15 companies are listed even where capacity is not "
             "publicly sourced.",
             "Annual, at year-end. Reviewed each quarter for new project approvals.", HIER_CO)

    put(ws, r, 1, "SECTION A - INDIA INDUSTRY CAPACITY", f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, D.CAPACITY_INDUSTRY_HEADERS, D.CAPACITY_INDUSTRY,
              "tbl_CapacityIndustry", left_cols=(1,), note_cols=(11,),
              col_fmts={3: FMT_1DP, 4: FMT_1DP, 5: FMT_1DP, 6: FMT_1DP})
    put(ws, r, 1, "SECTION B - COMPANY CAPACITY", f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, D.CAPACITY_COMPANY_HEADERS, D.CAPACITY_COMPANY,
              "tbl_CapacityCompany", left_cols=(1,), note_cols=(13,),
              col_fmts={5: FMT_3DP, 6: FMT_3DP, 7: FMT_3DP})
    put(ws, r, 1, "SECTION C - SAIL PLANT-WISE CAPACITY AND UTILISATION",
        f=font(11, bold=True, colour=NAVY))
    sail = [
        ("Bhilai Steel Plant (BSP)", "Chhattisgarh", 6.16, 0.832),
        ("Bokaro Steel Plant (BSL)", "Jharkhand", 4.60, 0.878),
        ("Rourkela Steel Plant (RSP)", "Odisha", 3.80, 1.038),
        ("IISCO Steel Plant (ISP), Burnpur", "West Bengal", 2.50, 0.916),
        ("Durgapur Steel Plant (DSP)", "West Bengal", 2.20, 1.016),
        ("Alloy Steels Plant (ASP), Durgapur", "West Bengal", 0.234, 0.398),
        ("Salem Steel Plant (SSP)", "Tamil Nadu", 0.180, 0.730),
    ]
    hdrs = ["SAIL Plant", "State", "Capacity (Mtpa)",
            "Average Capacity Utilisation FY2021-FY2025 (%)", "Source ID", "Confidence", "Notes"]
    rows = [(n, st, cap, ut, "S26", "High", "") for n, st, cap, ut in sail]
    rows.append(("TOTAL", "", 19.676, 0.908, "S26", "High",
                 "Plant capacities sum to 19.674 Mtpa against the stated total of 19.676 Mtpa - a "
                 "rounding difference of 0.002 Mtpa. Rourkela and Durgapur ran ABOVE nameplate on "
                 "a five-year average, while the Alloy Steels Plant averaged 39.8% and Salem 73%, "
                 "i.e. SAIL's utilisation problem is concentrated in its two small specialty "
                 "plants, not its five integrated works. Board-approved additions of 7.62 Mtpa "
                 "comprise IISCO greenfield 4.08, Bokaro brownfield 2.65 and Durgapur brownfield 0.89."))
    r = block(ws, r + 1, hdrs, rows, "tbl_CapacitySAIL", left_cols=(1,), note_cols=(7,),
              col_fmts={3: FMT_3DP, 4: FMT_PCT_1})
    footnote(ws, r, 1,
             "Do NOT sum Section B. It mixes crude steel, melt and aggregate installed metal "
             "capacity definitions across companies, and Tata Steel's 35 Mtpa is global while "
             "other figures are India-only. Use Section A for any India total.")


def production_volume(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 42), (2, 46), (3, 10), (4, 34), (5, 12), (6, 12), (7, 12), (8, 12),
                    (9, 12), (10, 18), (11, 34), (12, 11), (13, NOTE_W)])
    r = meta(ws, 3, "Production Volume",
             "Records crude steel, finished steel, DRI and pig iron output at industry level, and "
             "production and sales or delivery volumes at company level. Volume is the quantity "
             "line in every revenue build and the denominator for EBITDA per tonne.",
             "Production and sales are shown as separate rows, never blended, because they differ "
             "by inventory movement and because different issuers strike EBITDA per tonne on "
             "different denominators. Reporting Basis distinguishes consolidated global figures "
             "from India-only and from 100%-basis joint venture disclosure. Period Type is carried "
             "because the World Steel Association rows are CALENDAR years while all others are "
             "Indian fiscal years - conflating the two is a common and material error (Conflict C06).",
             "Mt and mn t (million tonnes). Both denote the same unit; the label follows the source.",
             "Industry: FY2022-FY2026 where published. Company: FY2025-FY2026, the periods for "
             "which primary volume disclosure was obtained. All 15 companies are listed.",
             "Annual (fiscal year). Industry data is also published monthly by the Joint Plant Committee.",
             HIER_GOV + " Company rows follow the issuer hierarchy.")
    put(ws, r, 1, "SECTION A - INDIA AND WORLD INDUSTRY VOLUMES",
        f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, D.VOLUME_INDUSTRY_HEADERS, D.VOLUME_INDUSTRY, "tbl_VolumeIndustry",
              left_cols=(1,), note_cols=(13,),
              col_fmts={3: FMT_2DP, 4: FMT_2DP, 5: FMT_2DP, 6: FMT_2DP, 7: FMT_2DP})
    put(ws, r, 1, "SECTION B - COMPANY VOLUMES", f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, D.VOLUME_COMPANY_HEADERS, D.VOLUME_COMPANY, "tbl_VolumeCompany",
              left_cols=(1, 2), note_cols=(13,),
              col_fmts={5: FMT_3DP, 6: FMT_3DP, 7: FMT_3DP, 8: FMT_3DP, 9: FMT_3DP})
    footnote(ws, r, 1,
             "Tata Steel's consolidated production is a composite of crude steel for India, liquid "
             "steel for the UK and Netherlands and saleable steel for South East Asia, and is not "
             "a like-for-like crude steel figure. JSW Steel's 'Combined operations' includes JSW "
             "JFE Steel Ltd for 27-31 March 2026 only. AM/NS India fiscal-year rows are derived by "
             "summing calendar quarters - see Conflict C16.")


def revenue(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 24), (2, 13), (3, 26), (4, 11), (5, 10), (6, 40),
                    (7, 13), (8, 13), (9, 13), (10, 13), (11, 13), (12, 20), (13, 12),
                    (14, 40), (15, 40), (16, 11), (17, NOTE_W)])
    r = meta(ws, 4, "Revenue",
             "Consolidated revenue from operations for the coverage universe on a consistent, "
             "peer-comparable basis, with each issuer's own headline measure carried separately "
             "where it differs.",
             "Revenue from Operations is the only construct comparable across the peer set: "
             "statutory, net of GST, per the audited or filed statement. Company-Headline Revenue "
             "exists because Jindal Steel leads with a 'Gross Revenue' measure that includes GST "
             "and other income; carrying both prevents an analyst reconciling to the wrong number "
             "(Conflict C02). The CAGR column is a live formula, not a hard-coded value, so it "
             "recalculates when a year is appended.",
             "Rs crore (Rs 10 million) for Indian reporters; US$ million for AM/NS India, which "
             "reports in dollars. The two are shown in separate blocks and are never summed.",
             "FY2022 to FY2026 for all 15 companies. FY2026 is the latest audited or filed year.",
             "Annual (fiscal year ending 31 March), except AM/NS India which is calendar.", HIER_CO)

    put(ws, r, 1, "SECTION A - INDIAN RUPEE REPORTERS", f=font(11, bold=True, colour=NAVY))
    hrow = r + 1
    r = block(ws, hrow, D.REVENUE_HEADERS, D.REVENUE, "tbl_Revenue",
              left_cols=(1, 3), note_cols=(17,))
    # live CAGR formulas in column M (13)
    for i in range(len(D.REVENUE)):
        rr = hrow + 1 + i
        put(ws, rr, 13,
            '=IFERROR(IF(AND(ISNUMBER(G%d),ISNUMBER(K%d),G%d>0),(K%d/G%d)^(1/4)-1,""),"")'
            % (rr, rr, rr, rr, rr), number_format=FMT_PCT_1, alignment=AL_RIGHT)
    put(ws, r, 1, "SECTION B - US DOLLAR REPORTER (AM/NS INDIA)",
        f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, D.REVENUE_HEADERS, D.REVENUE_USD, "tbl_RevenueUSD",
              left_cols=(1, 3), note_cols=(17,))
    footnote(ws, r, 1,
             "FY2026 revenue for Tata Steel, JSW Steel and SAIL ties exactly to the issuer results "
             "release. Jindal Steel and Jindal Stainless FY2026 revenue was verified by summing the "
             "four reported quarters. Do not add Section A and Section B: different currencies and, "
             "for AM/NS India, a different twelve-month period.")


def ebitda(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 24), (2, 13), (3, 11), (4, 10), (5, 46), (6, 56), (7, 20),
                    (8, 14), (9, 14), (10, 14), (11, 14), (12, 14), (13, 14), (14, 14),
                    (15, 20), (16, 15), (17, 40), (18, 34), (19, 11), (20, NOTE_W)])
    r = meta(ws, 5, "EBITDA",
             "EBITDA on two bases for every company: as reported by the issuer, and standardised "
             "so that the peer set is genuinely comparable. This dual presentation is not "
             "redundancy - it is required, because the companies in this universe do not define "
             "EBITDA the same way.",
             "EBITDA - As Reported is what the market quotes and what guidance is given against, so "
             "it must be present. EBITDA - Standardised is revenue from operations less operating "
             "expenses, excluding other income, finance cost, depreciation and amortisation, "
             "exceptional items and share of profit of joint ventures. 'Other Income Included In "
             "Reported EBITDA?' evidences the reconciliation between the two and turns an "
             "unexplained variance into a documented definitional difference. Shyam Metalics "
             "includes other income (Rs 2,333 cr plus Rs 204 cr equals the reported Rs 2,537 cr, "
             "exactly); Jindal Stainless does not (Rs 5,560 cr reported against Rs 5,561 cr "
             "standardised). The memo revenue column is present so that the margin is a live "
             "in-sheet formula rather than a hard-coded number.",
             "Rs crore for Indian reporters; US$ million for AM/NS India. Margin in per cent.",
             "Reported: FY2025-FY2026, the periods for which issuer disclosure was obtained. "
             "Standardised: FY2022-FY2026. All 15 companies.",
             "Annual (fiscal year), except AM/NS India which is calendar.", HIER_CO)

    rev_lookup = {row[0]: row[10] for row in D.REVENUE}
    hdrs = D.EBITDA_HEADERS[:14] + ["FY2026 Revenue from Operations (memo)"] \
        + D.EBITDA_HEADERS[14:]
    rows = []
    for row in D.EBITDA:
        rows.append(tuple(row[:14]) + (rev_lookup.get(row[0]),) + tuple(row[14:]))
    put(ws, r, 1, "SECTION A - INDIAN RUPEE REPORTERS", f=font(11, bold=True, colour=NAVY))
    hrow = r + 1
    r = block(ws, hrow, hdrs, rows, "tbl_EBITDA", left_cols=(1, 5, 6), note_cols=(20,),
              col_fmts={8: FMT_INT, 9: FMT_INT, 10: FMT_INT, 11: FMT_INT, 12: FMT_INT,
                        13: FMT_INT, 14: FMT_INT, 15: FMT_INT})
    for i in range(len(rows)):
        rr = hrow + 1 + i
        put(ws, rr, 16,
            '=IFERROR(IF(AND(ISNUMBER(N%d),ISNUMBER(O%d),O%d>0),N%d/O%d,""),"")'
            % (rr, rr, rr, rr, rr), number_format=FMT_PCT_1, alignment=AL_RIGHT)
    put(ws, r, 1, "SECTION B - US DOLLAR REPORTER (AM/NS INDIA)",
        f=font(11, bold=True, colour=NAVY))
    rows_usd = [tuple(x[:14]) + (None,) + tuple(x[14:]) for x in D.EBITDA_USD]
    r = block(ws, r + 1, hdrs, rows_usd, "tbl_EBITDAUSD", left_cols=(1, 5, 6), note_cols=(20,))
    footnote(ws, r, 1,
             "SAIL requires particular care: Rs 13,146 cr is the FY2026 STANDALONE headline, "
             "Rs 12,000 cr is standardised consolidated EBITDA excluding other income, and "
             "Rs 12,894 cr is consolidated EBITDA including other income. All three reconcile - "
             "see Conflict C09. Never pair a consolidated enterprise value with the standalone "
             "headline. Mukand FY2026 profit is not operating in nature - see Conflict C11.")


def ebitda_per_ton(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 24), (2, 42), (3, 10), (4, 10), (5, 40), (6, 13), (7, 13), (8, 13),
                    (9, 13), (10, 24), (11, 18), (12, 46), (13, 11), (14, NOTE_W)])
    r = meta(ws, 6, "EBITDA per Ton",
             "The sector's primary cross-cycle profitability metric, struck per tonne so that it is "
             "independent of scale and, where quoted in dollars, of currency. Presented with the "
             "volume denominator stated on every row.",
             "Volume Denominator is the critical column. EBITDA per tonne moves by several hundred "
             "rupees depending on whether it is struck on crude steel production, deliveries or "
             "shipments: Tata Steel uses deliveries, JSW Steel saleable sales, Jindal Steel sales "
             "volume. A per-tonne figure quoted without its denominator is not interpretable. "
             "Quarterly columns are carried because this metric is the standard quarterly "
             "surprise variable. The FY2027 guidance column is present because Jindal Stainless "
             "has guided BELOW its FY2026 outcome, which is an explicit management signal of "
             "expected margin normalisation.",
             "Rs/t for Indian reporters; US$/t for AM/NS India. Not summable across companies.",
             "FY2025-FY2026 annual, plus Q3FY2026 and Q4FY2026 where disclosed, plus FY2027 "
             "guidance where given. All 15 companies are listed; the metric is computed only where "
             "both EBITDA and a volume denominator are available.",
             "Annual and quarterly.", HIER_CO)
    r = block(ws, r, D.EBITDA_PT_HEADERS, D.EBITDA_PT, "tbl_EBITDAPerTon",
              left_cols=(1, 2, 5, 10), note_cols=(14,),
              col_fmts={6: FMT_2DP, 7: FMT_2DP, 8: FMT_2DP, 9: FMT_2DP})
    footnote(ws, r, 1,
             "Rows marked Methodology = 'Computed' were calculated in this workbook from the "
             "issuer's own EBITDA and volume; rows marked 'As reported' are the issuer's own "
             "published per-tonne figure. EBITDA per tonne is deliberately NOT presented for "
             "APL Apollo, a tube converter, and is not computable for six Supporting names whose "
             "volumes were not sourced. No per-tonne figure has been inferred.")


def steel_prices(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 40), (2, 42), (3, 20), (4, 52), (5, 10), (6, 10), (7, 26),
                    (8, 13), (9, 13), (10, 13), (11, 16), (12, 12), (13, 8), (14, 22),
                    (15, 11), (16, NOTE_W)])
    r = meta(ws, 7, "Steel Prices",
             "Domestic and international steel price observations for FY2026 and FY2027 to date, "
             "each carrying its assessment basis, so that realisation assumptions can be built on "
             "a stated and reproducible reference rather than an unattributed number.",
             "Assessment Basis is as important as the value itself. Indian HRC quotes differ by up "
             "to Rs 4,000/t purely on basis - ex-works against ex-Mumbai trade, GST-inclusive "
             "against exclusive, month-end against monthly average, primary mill against trade. "
             "52-week high and low are carried because the FY2026 range of Rs 45,700-59,600/t is a "
             "~30% swing and is the largest single source of FY2027 earnings uncertainty. Source "
             "Tier is shown on every row because no official Indian HRC series exists in the public "
             "domain and every domestic price here is Tier 2.",
             "Rs/t and US$/t as labelled per row. GST treatment is stated in Assessment Basis.",
             "FY2026 price path plus a June 2026 cross-regional snapshot and a July 2026 rebar "
             "benchmark. No continuous monthly history is asserted.",
             "Point-in-time and monthly observations. The Ministry of Steel publishes a month-end "
             "Mumbai retail series, registered here but published as a chart only.",
             HIER_PX)
    r = block(ws, r, M.STEEL_PRICE_HEADERS, M.STEEL_PRICES, "tbl_SteelPrices",
              left_cols=(1, 2, 3, 4, 7), note_cols=(16,),
              col_fmts={8: FMT_INT, 9: FMT_INT, 10: FMT_INT, 13: FMT_INT})
    r = footnote(ws, r, 1,
                 "MATERIAL LIMITATION: no fiscal-year average HRC realisation can be constructed "
                 "from these point observations. Where an average realisation is required, use the "
                 "issuer's own disclosed revenue divided by its own disclosed volume, or its "
                 "disclosed EBITDA per tonne, both of which are on the Revenue and EBITDA per Ton "
                 "sheets. See Conflict C07 on the March 2026 level.")
    footnote(ws, r, 1,
             "The safeguard duty covers FLAT products only. This is why primary rebar rose 25% "
             "against HRC's 18% between the December 2025 quarter and March 2026: flat supply "
             "tightened while long demand was lifted by construction activity.")


def iron_ore_prices(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 40), (2, 46), (3, 46), (4, 44), (5, 12), (6, 10), (7, 24),
                    (8, 14), (9, 12), (10, 12), (11, 18), (12, 8), (13, 11), (14, NOTE_W)])
    r = meta(ws, 8, "Iron Ore Prices",
             "International and Indian domestic iron ore price series, including a complete "
             "eleven-year fiscal-year history of the 62% Fe CFR China benchmark computed from World "
             "Bank monthly data, and NMDC's notified ex-mine prices.",
             "Grade and Specification is carried on every row because 62% Fe fines, 65.5% Fe lump "
             "and pellet are different products with different prices and different netbacks; "
             "quoting one against another understates or overstates delivered cost materially. "
             "Basis is carried because NMDC's notified prices are ex-mine and EXCLUDE royalty, DMF, "
             "NMET and cess - a gross-up of roughly 24% to a taxes-inclusive figure. The fiscal-"
             "year block is provided because Indian modelling requires April-March averages, which "
             "no publisher provides for this benchmark.",
             "US$/dmt for the international benchmark (dry metric tonne, 62% Fe, CFR China). Rs/t "
             "for domestic. US$/mt for the thermal coal reference column.",
             "International benchmark: FY2016-FY2026 fiscal-year averages, CY2015-CY2025 calendar "
             "averages, monthly April 2025 to June 2026, and quarterly. Domestic: dated "
             "notifications and assessments in FY2026 and FY2027 to date.",
             "Monthly for the World Bank series; per-notification for NMDC; weekly for broker "
             "assessments.", HIER_PX,
             extra=["UNIT WARNING (Conflict C14): the World Bank Pink Sheet column header for this "
                    "series reads '($/dmtu)', a legacy label from the era of annual contract "
                    "pricing. The Pink Sheet's own Description sheet states the spot series is in "
                    "US dollars per DRY TON for 62% Fe fines CFR China. Taking the header at face "
                    "value would misstate iron ore by a factor of roughly 62."])

    put(ws, r, 1, "SECTION A - INTERNATIONAL BENCHMARK, FISCAL-YEAR AVERAGES",
        f=font(11, bold=True, colour=NAVY))
    rows = [(fy, v, n, "Annual (FY) average", "S17",
             "Computed: arithmetic mean of the World Bank monthly nominal USD series over "
             "April to March", "High",
             "Iron ore (any origin) fines, spot, CFR China, 62% Fe, per the World Bank Description "
             "sheet." if fy == "FY2016" else
             ("Third consecutive annual decline. A persistent tailwind for non-integrated "
              "producers and a headwind for ore-integrated names." if fy == "FY2026" else
              ("Cycle peak of the period." if fy == "FY2022" else "")))
            for fy, v, n in M.IRON_ORE_FY]
    r = block(ws, r + 1, M.IRON_ORE_FY_HEADERS, rows, "tbl_IronOreFY",
              left_cols=(1,), note_cols=(8,), col_fmts={2: FMT_2DP, 3: FMT_INT})

    put(ws, r, 1, "SECTION B - INTERNATIONAL BENCHMARK, CALENDAR-YEAR AVERAGES",
        f=font(11, bold=True, colour=NAVY))
    rows = [(cy, v, "Calendar Year average", "S18", "As reported", "High",
             "Provided for cross-country comparison and for reconciling to non-Indian research. "
             "The annual Pink Sheet file is a March 2026 vintage and therefore stops at CY2025."
             if cy == "2025" else "") for cy, v in M.IRON_ORE_CY]
    r = block(ws, r + 1, M.IRON_ORE_CY_HEADERS, rows, "tbl_IronOreCY",
              left_cols=(1,), note_cols=(7,), col_fmts={2: FMT_2DP})

    put(ws, r, 1, "SECTION C - MONTHLY SERIES, APRIL 2025 TO JUNE 2026",
        f=font(11, bold=True, colour=NAVY))
    rows = [(lbl, io, tc, fy, fq, "Month", "S17", "High")
            for lbl, io, tc, fy, fq in M.IRON_ORE_MONTHLY]
    r = block(ws, r + 1, M.IRON_ORE_MONTHLY_HEADERS, rows, "tbl_IronOreMonthly",
              left_cols=(1,), col_fmts={2: FMT_2DP, 3: FMT_2DP})

    put(ws, r, 1, "SECTION D - QUARTERLY AVERAGES", f=font(11, bold=True, colour=NAVY))
    rows = [(q, io, tc, "Quarter average", "S17",
             "Computed: arithmetic mean of the World Bank monthly nominal USD series", "High")
            for q, io, tc in M.IRON_ORE_QTR]
    r = block(ws, r + 1, M.IRON_ORE_QTR_HEADERS, rows, "tbl_IronOreQtr",
              left_cols=(1,), col_fmts={2: FMT_2DP, 3: FMT_2DP})

    put(ws, r, 1, "SECTION E - INDIAN DOMESTIC PRICES AND NOTIFICATIONS",
        f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, M.IRON_ORE_DOMESTIC_HEADERS, M.IRON_ORE_DOMESTIC,
              "tbl_IronOreDomestic", left_cols=(1, 2, 3, 4, 7), note_cols=(13,),
              col_fmts={8: FMT_2DP})
    footnote(ws, r, 1,
             "The Australian thermal coal column in Sections C and D is included for energy-cost "
             "reference only and is NOT coking coal - see the Coking Coal Prices sheet and "
             "Conflict C15. The FY2026 NMDC notification series was not fully reconstructed within "
             "the research window; only the dated notifications verified are shown, and earlier "
             "months are deliberately left blank rather than interpolated.")


def coking_coal_prices(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 52), (2, 26), (3, 46), (4, 22), (5, 10), (6, 10), (7, 24),
                    (8, 12), (9, 13), (10, 13), (11, 16), (12, 12), (13, 8), (14, 11),
                    (15, NOTE_W)])
    r = meta(ws, 9, "Coking Coal Prices",
             "Premium hard coking coal observations from the two series available in the public "
             "domain, presented separately and never blended, plus an Australian thermal coal "
             "reference that is explicitly excluded from any coking coal average.",
             "Benchmark and Assessor are carried on every row because 'coking coal FOB Australia' "
             "is not one price: it varies by grade (premium low-vol hard coking coal against "
             "mid-vol, PCI or semi-soft), by index provider, by delivery term (FOB Australia "
             "against CFR India) and by assessment convention (spot against monthly average "
             "against quarterly settlement). Source Tier is carried because the Ministry of Steel "
             "series is Tier 1 and the broker series is Tier 2, and they disagree by up to US$40/t.",
             "US$/t for coking coal; US$/mt for the thermal coal reference. All nominal USD.",
             "Ministry of Steel narrative observations from October 2023 to March 2026; broker "
             "observations for February 2026, March 2026 and June 2026. No continuous monthly "
             "history is available publicly.",
             "Monthly and weekly point observations.", HIER_PX,
             extra=["CRITICAL - " + M.COKING_COAL_NOTE])
    r = block(ws, r, M.COKING_COAL_HEADERS, M.COKING_COAL, "tbl_CokingCoal",
              left_cols=(1, 2, 3, 4, 7), note_cols=(15,),
              col_fmts={8: FMT_2DP, 9: FMT_2DP, 10: FMT_2DP})
    r = footnote(ws, r, 1,
                 "CONFLICT C01 IS UNRESOLVED IN THE PUBLIC DOMAIN AND IS MATERIAL. For March 2026 "
                 "the Ministry of Steel assesses premium HCC FOB Australia at ~US$225/t while "
                 "IDBI Capital assesses coking coal FOB Australia at US$186/t. Decisively, the "
                 "broker series' 52-week high of US$217/t as at June 2026 is BELOW the Ministry's "
                 "February 2026 level of US$246/t, so the two are demonstrably different "
                 "assessments and must not be averaged together.")
    footnote(ws, r, 1,
             "Coking coal is the dominant swing factor in BF-BOF conversion cost. The rise from a "
             "~US$175/t trough in March 2025 to ~US$246/t in February 2026 is the principal reason "
             "Q4FY2026 margin commentary was cost-focused across Tata Steel, SAIL and Jindal Steel, "
             "and it is why Jindal Steel's EBITDA per tonne fell 11% despite 14% volume growth.")


def demand(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 30), (2, 22), (3, 13), (4, 22), (5, 20), (6, 20), (7, 20),
                    (8, 18), (9, 12), (10, 11), (11, NOTE_W)])
    r = meta(ws, 10, "Demand & Consumption",
             "Apparent finished steel consumption for India over twelve fiscal years, the full "
             "supply-demand balance that derives it, per capita intensity against world and China, "
             "and the latest quarterly reading.",
             "Apparent consumption is production plus imports less exports adjusted for the change "
             "in stocks - it is an APPARENT, not end-use, measure and the balance is shown in full "
             "so a user can see exactly how it is derived and how sensitive it is to the stock "
             "line. Consumption As % Of Production is a live formula, not a hard-coded value. Per "
             "capita consumption is carried because it is the standard cross-country intensity "
             "comparator and the quantitative core of the long-run India demand thesis.",
             "Mt (million tonnes). Per capita in kg. Growth in per cent.",
             "FY2015 to FY2026, twelve fiscal years, plus Q1 FY2027. FY2026 is provisional.",
             "Annual (fiscal year); also published monthly by the Joint Plant Committee.", HIER_GOV)

    put(ws, r, 1, "SECTION A - LONG-RUN CONSUMPTION SERIES", f=font(11, bold=True, colour=NAVY))
    hrow = r + 1
    rows = [(fy, cons, g, fp, cp, cap, None, "Annual (FY)", "S02", "High", note)
            for fy, cons, g, fp, cp, cap, note in M.DEMAND]
    r = block(ws, hrow, M.DEMAND_HEADERS, rows, "tbl_Demand", left_cols=(1,), note_cols=(11,),
              col_fmts={2: FMT_2DP, 3: FMT_PCT_1, 4: FMT_2DP, 5: FMT_2DP, 6: FMT_1DP,
                        7: FMT_PCT_1})
    for i in range(len(rows)):
        rr = hrow + 1 + i
        put(ws, rr, 7,
            '=IFERROR(IF(AND(ISNUMBER(B%d),ISNUMBER(D%d),D%d>0),B%d/D%d,""),"")'
            % (rr, rr, rr, rr, rr), number_format=FMT_PCT_1, alignment=AL_RIGHT)

    put(ws, r, 1, "SECTION B - FINISHED STEEL SUPPLY-DEMAND BALANCE (APRIL 2026 JPC VINTAGE)",
        f=font(11, bold=True, colour=NAVY))
    rows = [(ln, u, a, b, c, d, "Annual (FY)", "S01", "High", note)
            for ln, u, a, b, c, d, note in M.DEMAND_SUPPLY_BALANCE]
    r = block(ws, r + 1, M.DEMAND_SUPPLY_BALANCE_HEADERS, rows, "tbl_SupplyBalance",
              left_cols=(1,), note_cols=(10,),
              col_fmts={3: FMT_2DP, 4: FMT_2DP, 5: FMT_2DP, 6: FMT_2DP})

    put(ws, r, 1, "SECTION C - PER CAPITA CONSUMPTION", f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, M.PER_CAPITA_HEADERS, M.PER_CAPITA, "tbl_PerCapita",
              left_cols=(1,), note_cols=(8,), col_fmts={4: FMT_1DP})

    put(ws, r, 1, "SECTION D - LATEST QUARTER", f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, M.QUARTERLY_DEMAND_HEADERS, M.QUARTERLY_DEMAND, "tbl_QtrDemand",
              left_cols=(1,), note_cols=(9,),
              col_fmts={2: FMT_2DP, 3: FMT_PCT_1, 4: FMT_2DP, 5: FMT_PCT_1})
    footnote(ws, r, 1,
             "Sections A and B are on DIFFERENT JPC vintages and will not tie for FY2026: the June "
             "2026 vintage in Section A shows consumption of 164.19 Mt while the internally "
             "consistent April 2026 balance in Section B shows 163.74 Mt. This is deliberate and "
             "is documented as Conflict C03. Note also that the stock line was negative in both "
             "FY2025 and FY2026, so destocking flattered reported consumption growth.")


def trade(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 16), (2, 30), (3, 10), (4, 16), (5, 18), (6, 16), (7, 20),
                    (8, 20), (9, 24), (10, 18), (11, 12), (12, 11), (13, NOTE_W)])
    r = meta(ws, 11, "Imports & Exports",
             "India's finished steel trade over twelve fiscal years by volume and value, product "
             "mix for FY2025 and FY2026, source and destination country shares, and the monthly "
             "net-trade turning points around the safeguard duty.",
             "Net Trade is a live formula, not a hard-coded value, and matters because it "
             "determines whether domestic prices set at import parity or export parity - which "
             "drives the entire domestic price deck. Product-level detail is carried because the "
             "safeguard duty covers flat products only, so the trade response is concentrated in "
             "specific HS categories rather than spread evenly. Country shares are carried because "
             "Korea and Japan supply under free trade agreements, which is central to the policy "
             "debate, and because EU destinations quantify India's CBAM exposure.",
             "Mt (million tonnes) for finished steel; '000 tonnes for semis and pig iron; "
             "Rs crore for values. Shares in per cent.",
             "FY2015 to FY2026 for the annual series. Product and country detail for FY2025 and "
             "FY2026. Monthly turning points for March 2025 to March 2026.",
             "Annual, with monthly detail published by the Joint Plant Committee.", HIER_GOV)

    put(ws, r, 1, "SECTION A - ANNUAL TRADE, FY2015 TO FY2026", f=font(11, bold=True, colour=NAVY))
    hrow = r + 1
    rows = [(fy, imp, exp, None, semis, pi, pe, iv, ev, "Annual (FY)", "S02", "High", note)
            for fy, imp, exp, semis, pi, pe, iv, ev, note in M.TRADE_ANNUAL]
    r = block(ws, hrow, M.TRADE_ANNUAL_HEADERS, rows, "tbl_TradeAnnual",
              left_cols=(1,), note_cols=(13,),
              col_fmts={2: FMT_3DP, 3: FMT_3DP, 4: FMT_3DP, 5: FMT_INT, 6: FMT_INT,
                        7: FMT_INT, 8: FMT_INT, 9: FMT_INT})
    for i in range(len(rows)):
        rr = hrow + 1 + i
        put(ws, rr, 4,
            '=IFERROR(IF(AND(ISNUMBER(B%d),ISNUMBER(C%d)),C%d-B%d,""),"")' % (rr, rr, rr, rr),
            number_format=FMT_3DP, alignment=AL_RIGHT)

    put(ws, r, 1, "SECTION B - PRODUCT MIX", f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, M.TRADE_PRODUCT_HEADERS, M.TRADE_PRODUCT, "tbl_TradeProduct",
              left_cols=(1, 2), note_cols=(11,),
              col_fmts={4: FMT_2DP, 5: FMT_2DP, 6: FMT_PCT_1, 7: FMT_PCT_1})

    put(ws, r, 1, "SECTION C - SOURCE AND DESTINATION COUNTRY SHARES",
        f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, M.TRADE_COUNTRY_HEADERS, M.TRADE_COUNTRY, "tbl_TradeCountry",
              left_cols=(1, 2), note_cols=(8,), col_fmts={4: FMT_PCT_1, 5: FMT_PCT_1})

    put(ws, r, 1, "SECTION D - MONTHLY NET TRADE TURNING POINTS",
        f=font(11, bold=True, colour=NAVY))
    r = block(ws, r + 1, M.TRADE_MONTHLY_HEADERS, M.TRADE_MONTHLY, "tbl_TradeMonthly",
              left_cols=(1,), note_cols=(8,),
              col_fmts={2: FMT_2DP, 3: FMT_2DP, 4: FMT_2DP})
    put(ws, r, 1, M.TRADE_NARRATIVE, f=font(9, colour=TEXT_MUTED), alignment=AL_LEFT_WRAP)
    ws.row_dimensions[r].height = 42
    footnote(ws, r + 2, 1,
             "Section A finished steel volumes are 'total finished steel' including alloy and "
             "stainless. FY2026 is provisional. Section B percentages are changes on FY2025, not "
             "shares, except the TOTAL rows.")


def policies(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 9), (2, 52), (3, 22), (4, 40), (5, 30), (6, 34), (7, 26), (8, 34),
                    (9, 30), (10, 62), (11, 52), (12, 78), (13, 14), (14, 11), (15, NOTE_W)])
    r = meta(ws, 12, "Government Policies",
             "The policy and regulatory stack bearing on Indian steel, dated and scoped, covering "
             "trade remedies, standards and non-tariff measures, industrial policy, fiscal "
             "measures, decarbonisation, state support and the overseas regimes that affect Indian "
             "producers' export books and foreign operations.",
             "Announced date, Effective From and Effective To are separate columns because the "
             "single most consequential policy event of FY2026 - the safeguard duty - had a "
             "provisional measure that lapsed, a final recommendation, and a definitive "
             "notification whose year-one window is back-dated over the gap. Collapsing these into "
             "one date destroys the ability to explain the FY2026 price path. Status At Cutoff "
             "distinguishes measures in force from those superseded. Assessed Sector Impact states "
             "direction and magnitude rather than describing the measure again, because a database "
             "that only restates the gazette is of no use to an investment committee.",
             "Rates in per cent; quanta in Rs crore, Mt or Mtpa as stated in the terms column.",
             "20 measures and events, from the National Steel Policy 2017 through to July 2026. "
             "Includes two overseas regimes (EU CBAM, UK quotas), one overseas enforcement action "
             "(Netherlands) and one geopolitical event, because all four are cited in issuers' own "
             "filed FY2026 disclosures.",
             "Event-driven. Reviewed each quarter and on any gazette notification.",
             "Tier 1 throughout for the Indian measures: gazette notifications, DGTR "
             "determinations, Parliamentary answers and Ministry of Steel orders. Issuer results "
             "releases are used where an issuer is the primary source for an overseas measure's "
             "effect on it.")
    r = block(ws, r, M.POLICY_HEADERS, M.POLICIES, "tbl_Policies",
              left_cols=(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), note_cols=(10, 11, 12, 15),
              col_fmts={})
    footnote(ws, r, 1,
             "P01, P02 and P03 must be read as a single sequence - see Conflict C08. The year-two "
             "safeguard rate of 11.5% has applied since 21 April 2026 and a mid-term review is "
             "provided for: this is the key FY2027 policy monitorable. P07, the suspension of the "
             "stainless steel Quality Control Order, is the clearest live downside risk to a "
             "specific issuer in the universe and is corroborated in Jindal Stainless's own "
             "exchange-filed earnings call. P18, the Netherlands enforcement action, carries a "
             "disclosed material uncertainty related to going concern at Tata Steel Netherlands "
             "and is the largest single identifiable downside in the coverage universe.")


def kpis(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 9), (2, 52), (3, 18), (4, 13), (5, 16), (6, 16), (7, 16),
                    (8, 18), (9, 24), (10, 30), (11, 14), (12, 52), (13, 11), (14, NOTE_W)])
    r = meta(ws, 13, "Industry KPIs",
             "A single dashboard-ready indicator set spanning supply, demand, trade, input costs, "
             "realisations, policy and coverage, so that downstream tools bind to one sheet rather "
             "than scraping every sheet in the workbook.",
             "Every KPI carries a stable KPI ID so a dashboard or Power BI model can reference it "
             "without depending on row position. Latest Available and Latest Period are separate "
             "from the fiscal-year columns because several indicators have a more recent reading "
             "than FY2026 - the most recent industry datapoint in this workbook is Q1 FY2027 "
             "consumption. 'Why This KPI Matters' is carried because an indicator without a stated "
             "investment rationale invites misuse.",
             "As labelled per row. Percentages are stored as decimals and formatted as per cent.",
             "FY2024 to FY2026 plus the latest available reading. 24 indicators.",
             "Mixed - annual, quarterly, monthly and point-in-time, stated per row.",
             "Inherits the source of each underlying sheet; Source ID is carried on every row.")
    r = block(ws, r, M.KPI_HEADERS, M.KPIS, "tbl_KPIs",
              left_cols=(1, 2, 3, 9, 10, 12), note_cols=(12, 14),
              col_fmts={})
    footnote(ws, r, 1,
             "K16, premium HCC coking coal, deliberately shows no fiscal-year values: no defensible "
             "fiscal-year average exists in the public domain - see Conflict C01. K03 uses year-end "
             "capacity as the denominator and therefore understates true utilisation in years of "
             "heavy commissioning. K21 is a coverage aggregate, not an industry total: it omits "
             "AM/NS India, RINL and the large secondary sector.")


def sources_sheet(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 10), (2, 7), (3, 40), (4, 52), (5, 66), (6, 34), (7, 56), (8, 34),
                    (9, 96), (10, 60), (11, 14), (12, NOTE_W)])
    r = meta(ws, 14, "Sources",
             "The authority record for the workbook. Every figure on every sheet carries a Source "
             "ID that resolves to exactly one row here, giving publisher, publication, date, "
             "section or page, period covered, URL and verification status.",
             "Tier classifies the source: 1 primary (issuer, government, official statistical "
             "agency), 2 secondary (rating agency, sell-side, price reporting agency), 3 tertiary "
             "(aggregator of filings, general press). Verification Status is explicit about what "
             "was actually done - whether a document was retrieved and parsed in full, or a figure "
             "was taken from a search snippet, or a value was reconciled arithmetically. This "
             "distinction is what allows a reviewer to decide how much weight a figure carries. "
             "Section or Page Reference is given so a figure can be re-found rather than re-hunted.",
             "Not applicable - reference sheet.",
             "47 sources. Source IDs are stable and must never be re-pointed at a different "
             "document; if a document is superseded, add a new ID and retire the old one.",
             "Updated whenever a figure is added or refreshed.",
             "Not applicable - this sheet defines the hierarchy used elsewhere.")
    r = block(ws, r, S.SOURCE_HEADERS, S.SOURCES, "tbl_Sources",
              left_cols=(1, 3, 4, 5, 6, 7, 8, 9, 10, 11), note_cols=(9, 10, 12),
              col_fmts={2: FMT_INT})
    footnote(ws, r, 1,
             "Tier 3 sources are used only to extend series backwards and to cross-check. No "
             "FY2026 figure in this workbook rests on a Tier 3 source alone: each was reconciled "
             "to a Tier 1 or Tier 2 source or to the sum of the four reported quarters.")


def dictionary_sheet(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 26), (2, 42), (3, 96), (4, 26), (5, 110)])
    r = meta(ws, 15, "Data Dictionary",
             "Defines every column used in the workbook and states why it exists, so that a "
             "reviewer who has never seen this file can interpret any cell without asking.",
             "Rows tagged 'ALL SHEETS' define the global metadata convention applied to every data "
             "row: Source ID, Confidence, Reporting Basis, Currency, Unit, Fiscal Year, Period "
             "Type, Last Updated, Notes and Methodology. Sheet-specific rows follow.",
             "Not applicable - reference sheet.",
             "All columns on all 14 mandated worksheets.",
             "Updated whenever a column is added or redefined.",
             "Not applicable.")
    r = block(ws, r, S.DICTIONARY_HEADERS, S.DICTIONARY, "tbl_Dictionary",
              left_cols=(1, 2, 3, 4, 5), note_cols=(3, 5))
    footnote(ws, r, 1,
             "Methodology values are controlled: 'As reported' means lifted from a document "
             "unchanged; 'Computed' means calculated in this workbook from figures on the same or "
             "another sheet; 'Derived' means constructed by combining periods or bases, as with "
             "AM/NS India's fiscal-year equivalent; 'Cross-checked' means sourced from one place "
             "and independently reconciled to another.")


def conflicts_sheet(ws):
    sheet_defaults(ws)
    set_widths(ws, [(1, 10), (2, 46), (3, 34), (4, 78), (5, 78), (6, 66), (7, 110),
                    (8, 110), (9, 66), (10, 96)])
    r = meta(ws, 16, "Conflicts Log",
             "Every conflict, vintage difference and definitional trap identified during research, "
             "with the competing figures and their sources, the reason for the difference, the "
             "resolution adopted in this workbook, the figure carried forward, and the residual "
             "risk that remains after resolution.",
             "Figures A, B and C are stated in full with their Source IDs so a reviewer can "
             "re-derive the judgement rather than accept it. 'Why the figures differ' is separated "
             "from 'Resolution adopted' because the reason is often more useful than the choice: "
             "several entries here are not conflicts at all but definitional differences that "
             "reconcile exactly once the basis is fixed, and knowing which is which prevents a "
             "modeller chasing a phantom error. Residual Risk is carried because resolving a "
             "conflict on today's information does not eliminate the exposure.",
             "As stated per row.",
             "16 entries. Three are material to modelling: C01 coking coal, C02 Jindal Steel "
             "revenue basis, C16 AM/NS India fiscal year mismatch. Two are units and naming traps "
             "that would cause large errors if missed: C14 and C15.",
             "Updated whenever a conflict is identified or resolved.",
             "Not applicable - this sheet adjudicates between sources.")
    r = block(ws, r, S.CONFLICT_HEADERS, S.CONFLICTS, "tbl_Conflicts",
              left_cols=tuple(range(1, 11)), note_cols=(4, 5, 6, 7, 8, 9, 10))
    footnote(ws, r, 1,
             "C06, C10, C14, C15 and C16 are NOT disagreements between publishers. They are "
             "definitional or unit differences that reconcile exactly once the basis is stated, "
             "and they are logged here precisely because each is a well-known way to lose money in "
             "this sector.")
