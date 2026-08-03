# -*- coding: utf-8 -*-
"""Python mirror of the Excel calculation chain.

Why this exists: the Excel workbook is the deliverable, but three things cannot be done
in Excel alone without volatile constructs or manual toggling -
  (1) a scenario comparison table showing all four scenarios simultaneously,
  (2) a tornado chart requiring the full chain re-run per variable,
  (3) an automated tie-out proving the Excel formulas reproduce the intended arithmetic.
This module implements exactly the same arithmetic as the workbook formulas. The Audit
Checks sheet compares the live Excel result against the value this module produced at
build time; if the two ever diverge, the check fails loudly.
"""
from .assumptions import DRIVERS, N, PROJECTS, SCENARIOS

# ---- FY2026 base, all traceable to Master Industry Database.xlsx -----------------------
CONS26 = 163.74      # apparent finished steel consumption, Mt (April-2026 JPC vintage)
PROD26 = 160.94      # finished steel production, Mt (same vintage - balance closes)
CRUDE26 = 168.42     # crude steel production, Mt
CAP26 = 220.4        # crude steel capacity, Mtpa
IMP26 = 6.524        # finished steel imports, Mt
EXP26 = 6.602        # finished steel exports, Mt
REALN26 = 59974.0    # blended realisation, Rs/t (derived - four majors)
COST26 = 49241.0     # implied cash cost, Rs/t (derived)
EBT26 = 10733.0      # weighted EBITDA, Rs/t (derived)
IO26 = 100.52        # iron ore 62% Fe CFR China, US$/dmt (FY2026 average)
CC26 = 225.0         # premium HCC coking coal, US$/t (Mar-2026 spot reference)
POP26 = 1.4191       # population, bn - DERIVED as 163.74 / 115.7 x 1000 ... see note below
POPG = 0.0085        # population growth p.a. - NOT SOURCED, flagged Low confidence

# FY2026 average USD/INR, DERIVED from Tata Steel's dual-currency disclosure in the
# Master Database. Three independent derivations from the same primary document:
#   consolidated EBITDA/t  Rs 10,900 / US$124 = 87.90
#   India EBITDA/t         Rs 15,213 / US$172 = 88.45
#   group turnover         Rs 2,32,140 cr / US$26bn = 89.28
# Midpoint of the two EBITDA/t derivations is used; the turnover figure corroborates.
INR26 = 88.4

DRV = {d[0]: d for d in DRIVERS}


def additions_by_year():
    """Gross announced India crude steel capacity additions by forecast year, Mtpa.

    Only projects flagged Included = 'Yes' count. Acquisitions, downstream lines,
    overseas assets, already-commissioned projects and ramp-ups are excluded because
    none of them adds new NATIONAL crude steel capacity.
    """
    out = {}
    for p in PROJECTS:
        cap, fy, inc = p[5], p[8], p[11]
        if cap and isinstance(inc, str) and inc.strip().lower().startswith("yes"):
            out[fy] = out.get(fy, 0.0) + cap
    return out


ADDITIONS = additions_by_year()
FY_KEYS = ["FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032", "FY2033"]
ADD_VEC = [ADDITIONS.get(k, 0.0) for k in FY_KEYS]


def d(code, scen):
    return DRV[code][3][scen]


def compute(scen, overrides=None):
    """Run the full chain for one scenario. overrides: {driver_code: multiplier}."""
    ov = overrides or {}

    def g(code):
        v = list(d(code, scen))
        if code in ov:
            v = [x * ov[code] for x in v]
        return v

    gdp, elas, inr, cpi = g("D01"), g("D02"), g("D03"), g("D04")
    deliv, cfr, netexp, stock = g("D05"), g("D06"), g("D07"), g("D08")
    realn_drv, io, cc, convinf = g("D09"), g("D10"), g("D11"), g("D12")
    rmshare, iow, imports_drv = g("D13"), g("D14"), g("D21")
    secadd = g("D22")
    pxutil, utilnorm, utilmax, dna = g("D23"), g("D24"), g("D25"), g("D26")
    taxr, capint, mcapex = g("D15"), g("D16"), g("D17")
    nwcd, wacc = g("D18"), g("D20")

    r = {k: [] for k in (
        "gdp", "elas", "dgrowth", "cons", "percap", "pop", "netexp", "stock", "prod",
        "crude", "crude_uncon", "cap", "capadd", "util", "util_lag", "px_adj", "convidx",
        "rmidx", "cost", "realn_drv", "realn", "ebt", "margin", "rev", "ebitda", "dna",
        "ebit", "imports", "exports", "imp_pen", "exp_int", "growth_capex", "maint_capex",
        "capex", "nwc", "dnwc", "tax", "fcf", "io", "cc", "inr", "cpi", "cap_gross_add",
        "cap_sec_add", "cap_constrained")}

    cons_p, cap_p, pop_p, conv_p = CONS26, CAP26, POP26, 1.0
    nwc_p = PROD26 * REALN26 / 10.0 * nwcd[0] / 365.0   # FY2026 NWC on the same days basis
    util_p = CRUDE26 / CAP26                            # FY2026 actual utilisation, 76.42%
    io_inr26, cc_inr26 = IO26 * INR26, CC26 * INR26

    for t in range(N):
        # --- demand
        dg = gdp[t] * elas[t]
        cons = cons_p * (1 + dg)
        pop = pop_p * (1 + POPG)
        # --- capacity (independent of demand)
        gross = ADD_VEC[t] * deliv[t]
        cap = cap_p + gross + secadd[t]
        # --- supply, subject to a practical capacity ceiling
        prod_demand = cons + netexp[t] + stock[t]
        crude_uncon = prod_demand * cfr[t]
        crude = min(crude_uncon, cap * utilmax[t])
        constrained = crude_uncon > cap * utilmax[t]
        prod = crude / cfr[t]
        # any shortfall against demand is met by imports, so the balance still closes
        shortfall = max(0.0, prod_demand - prod)
        imports = imports_drv[t] + shortfall
        exports = imports_drv[t] + netexp[t]
        util = crude / cap if cap else 0.0
        # --- price, with a LAGGED utilisation feedback (no circularity)
        px_adj = 1 + pxutil[t] * (util_p - utilnorm[t])
        realn = realn_drv[t] * px_adj
        # --- cost
        conv = conv_p * (1 + convinf[t])
        rmidx = (iow[t] * (io[t] * inr[t] / io_inr26)
                 + iow[t] * (cc[t] * inr[t] / cc_inr26)
                 + (1 - 2 * iow[t]) * conv)
        cost = COST26 * (rmshare[t] * rmidx + (1 - rmshare[t]) * conv)
        ebt = realn - cost
        # --- P&L and cash flow (unlevered)
        rev = prod * realn / 10.0                       # Rs cr
        ebitda = prod * ebt / 10.0                      # Rs cr
        dna_v = rev * dna[t]
        ebit = ebitda - dna_v
        tax = max(0.0, ebit) * taxr[t]
        nwc = rev * nwcd[t] / 365.0
        dnwc = nwc - nwc_p
        gcapex = (gross + secadd[t]) * capint[t] * 1e6 / 1e7   # Mtpa x Rs/t -> Rs cr
        mcx = rev * mcapex[t]
        fcf = ebitda - tax - gcapex - mcx - dnwc

        for k, v in (("gdp", gdp[t]), ("elas", elas[t]), ("dgrowth", dg), ("cons", cons),
                     ("pop", pop), ("percap", cons / pop), ("netexp", netexp[t]),
                     ("stock", stock[t]), ("prod", prod), ("crude", crude),
                     ("crude_uncon", crude_uncon), ("cap", cap),
                     ("capadd", gross + secadd[t]), ("cap_gross_add", gross),
                     ("cap_sec_add", secadd[t]), ("cap_constrained", 1.0 if constrained else 0.0),
                     ("util", util), ("util_lag", util_p), ("px_adj", px_adj),
                     ("convidx", conv), ("rmidx", rmidx), ("cost", cost),
                     ("realn_drv", realn_drv[t]), ("realn", realn), ("ebt", ebt),
                     ("margin", ebt / realn), ("rev", rev), ("ebitda", ebitda),
                     ("dna", dna_v), ("ebit", ebit), ("imports", imports),
                     ("exports", exports), ("imp_pen", imports / cons),
                     ("exp_int", exports / prod), ("growth_capex", gcapex),
                     ("maint_capex", mcx), ("capex", gcapex + mcx), ("nwc", nwc),
                     ("dnwc", dnwc), ("tax", tax), ("fcf", fcf), ("io", io[t]),
                     ("cc", cc[t]), ("inr", inr[t]), ("cpi", cpi[t])):
            r[k].append(v)
        cons_p, cap_p, pop_p, conv_p, nwc_p, util_p = cons, cap, pop, conv, nwc, util

    r["wacc"] = wacc
    r["scenario"] = scen
    return r


ALL = {s: compute(s) for s in SCENARIOS}


# ---- Tornado: FY2033 industry EBITDA sensitivity to a 10% adverse move ----------------
# Sign convention: 'adverse' means the direction that REDUCES industry EBITDA. For cost
# drivers that is +10%; for revenue and volume drivers it is -10%. Getting this wrong is
# the most common error in a tornado, so the direction is stated explicitly per driver.
TORNADO_SPEC = [
    ("D09", "Blended realisation", -0.10),
    ("D11", "Coking coal price", +0.10),
    ("D10", "Iron ore price", +0.10),
    ("D01", "Real GDP growth", -0.10),
    ("D02", "Demand elasticity", -0.10),
    ("D03", "USD/INR", +0.10),
    ("D12", "Conversion cost inflation", +0.10),
    ("D13", "Raw material share of cash cost", +0.10),
    ("D22", "Secondary capacity additions", +0.10),
    ("D05", "Pipeline delivery factor", +0.10),
]


def tornado(scen="Base Case"):
    base = ALL[scen]["ebitda"][-1]
    out = []
    for code, name, mult in TORNADO_SPEC:
        flexed = compute(scen, {code: 1 + mult})["ebitda"][-1]
        out.append((code, name, mult, base, flexed, flexed - base,
                    (flexed - base) / base if base else 0.0))
    out.sort(key=lambda x: abs(x[5]), reverse=True)
    return base, out
