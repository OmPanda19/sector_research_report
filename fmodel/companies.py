# -*- coding: utf-8 -*-
"""Company FY2026 actuals, lifted from Master Industry Database.xlsx.

BASIS DISCIPLINE. The industry model is an INDIA model, so volume, realisation and
EBITDA-per-tonne use the India segment where an issuer reports one (Tata Steel, JSW Steel).
Valuation inputs - net debt, profit, EPS - are CONSOLIDATED because that is what an
enterprise value corresponds to. Every field states which basis it is on. Mixing the two
is the single most common error in Indian steel comparables.
"""
from .chain import INR26

# (key, name, vol_Mt, vol_basis, rev_cr, rev_basis, ebitda_cr, ebitda_basis, ebitda_t,
#  consol_rev, consol_ebitda, net_debt, net_debt_note, pat, eps, dep, listed, src)
COMPANIES = [
    ("TATA", "Tata Steel", 22.53, "India deliveries", 140302, "India", 34272, "India reported",
     15213, 232140, 34848, 80144, "Consolidated net debt", 10886, 8.65, 11955, True, "S08"),
    ("JSW", "JSW Steel", 28.76, "India steel sales", 174854, "India operations", 29240,
     "India reported", 10167, 185470, 29821, 53870, "Consolidated net debt, excludes leases "
     "and revenue acceptances", 25508, 91.26, 9601, True, "S09"),
    ("SAIL", "SAIL", 19.93, "Standalone sales volume", 110811, "Consolidated", 13146,
     "STANDALONE reported", 6596, 110811, 12000, 31928,
     "PROXY - gross borrowings; SAIL net debt is not disclosed in the database", 3373, 8.17,
     5988, True, "S10"),
    ("JSPL", "Jindal Steel", 8.68, "Consolidated steel sales", 53225, "Consolidated", 9660,
     "Consolidated reported (adj 9,099 plus 561 FX)", 10482, 53225, 9660, 16019,
     "Consolidated net debt", 3361, 33.01, 3171, True, "S11"),
    ("JSL", "Jindal Stainless", 2.57, "Consolidated deliveries", 42955, "Consolidated", 5560,
     "Consolidated reported", 21634, 42955, 5560, 3040, "Consolidated net debt", 3185, 38.74,
     1060, True, "S12"),
    ("AMNS", "AM/NS India", 7.975, "Derived FY2026 shipments (sum of calendar quarters)",
     round(6196 * INR26 / 1000.0), "100% basis, USD converted at the derived FY2026 average "
     "USD/INR of 88.4", round(778 * INR26 / 1000.0), "100% basis, USD converted",
     round(778 * INR26 / 1000.0 / 7.975 * 10), round(6196 * INR26 / 1000.0),
     round(778 * INR26 / 1000.0), None,
     "Not disclosed - AM/NS India does not publish a balance sheet publicly", None, None, None,
     False, "S13/S14"),
    ("RINL", "RINL", 4.4, "Standalone sales volume (issuer rounded)", 22300,
     "Turnover, basis not specified by issuer", None, "Not publicly available", None, 22300,
     None, None, "Not disclosed - unlisted, wholly Government-owned", None, None, None, False,
     "S16"),
]

# Company capacity at FY2026 and its own tracked additions, for the company volume build.
# Only India crude steel capacity is relevant to the India volume build.
CO_CAPACITY = {
    "TATA": (26.0, "INDICATIVE - Tata Steel discloses 35 Mtpa GLOBAL group crude steel "
                   "capacity. The India component is not separately disclosed in the database; "
                   "26 Mtpa is inferred from FY2026 India crude steel production of 23.43 Mt at "
                   "roughly 90% utilisation and is flagged Low confidence."),
    "JSW": (31.9, "JSW Steel India existing capacity, disclosed. Cross-checked: FY2026 India "
                  "production of 29.25 Mt at the stated 92% utilisation implies 31.8 Mtpa."),
    "SAIL": (19.676, "SAIL crude steel operating capacity per Ministry of Steel, Rajya Sabha "
                     "USQ 3982 answered 27-Mar-2026."),
    "JSPL": (15.6, "Jindal Steel total steelmaking capacity entering FY2027, disclosed."),
    "JSL": (3.0, "Jindal Stainless INDIA melt capacity. Group melt capacity is 4.2 Mtpa "
                 "including 1.2 Mtpa in Indonesia, which is excluded from the India build."),
    "AMNS": (9.0, "AM/NS India Hazira capacity, expanding to 15 Mtpa."),
    "RINL": (None, "Not verified - a 7.3 Mtpa liquid steel nameplate is widely cited but is not "
                   "asserted in the database."),
}

# Company-attributed additions from the project tracker, by company key and forecast FY.
# Derived from fmodel.assumptions.PROJECTS. Acquisitions ARE included here (they add to the
# company even though they add nothing to national capacity) and are flagged.
CO_ADDITIONS = {
    "TATA": {"FY2030": 4.80, "FY2033": 5.00},
    "JSW": {"FY2027": 2.40, "FY2030": 5.00, "FY2031": 6.00},   # incl. BMM Ispat 0.90 in FY2027
    "SAIL": {"FY2029": 0.89, "FY2030": 2.65, "FY2031": 4.08, "FY2033": 8.00},
    "JSPL": {},
    "JSL": {},
    "AMNS": {"FY2029": 6.00, "FY2032": 8.20},
    "RINL": {},
}

CO_NOTES = {
    "TATA": "India segment used for volume and margin. Netherlands carries a disclosed material "
            "uncertainty related to going concern (Master DB, policy P18) - the single largest "
            "identifiable downside in the universe and NOT captured in this industry model, "
            "which forecasts India only.",
    "JSW": "FY2027 additions include BMM Ispat (0.90 Mtpa), an acquisition of existing capacity: "
           "it adds to JSW but NOT to national capacity. BPSL was deconsolidated from "
           "27-Mar-2026, so FY2026 to FY2027 company volumes are not like-for-like.",
    "SAIL": "EBITDA per tonne is struck on the STANDALONE EBITDA headline over standalone sales "
            "volume. Do not pair a consolidated enterprise value with this EBITDA - see Master "
            "DB Conflict C09.",
    "JSPL": "Capacity rose from 9.6 to 15.6 Mtpa in FY2026, so FY2027 volume growth comes from "
            "utilising capacity already built rather than from new additions. Company guidance "
            "for FY2027 is 11.0-11.5 Mt of production against 9.25 Mt in FY2026.",
    "JSL": "STAINLESS. Realisation of roughly Rs 1,67,000/t and EBITDA per tonne of Rs 21,634 "
           "are not comparable with carbon steel and this company is EXCLUDED from the blended "
           "industry realisation anchor. Management guided FY2027 EBITDA per tonne to "
           "Rs 18,000-20,000, i.e. BELOW the FY2026 outcome.",
    "AMNS": "Unlisted 60:40 joint venture. All figures are 100% basis, USD, calendar quarters, "
            "converted at the derived FY2026 average USD/INR of 88.4. The fiscal-year figures "
            "are themselves derived by summing calendar quarters - Master DB Conflict C16.",
    "RINL": "Unlisted and wholly Government-owned. No EBITDA, balance sheet or valuation data is "
            "publicly available, so RINL appears in the volume build only.",
}
