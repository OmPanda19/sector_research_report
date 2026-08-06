"""Fills for 17 Capital Allocation, 18 Industry Cycle Model, 19 Trade Model and 20 ESG Model."""
from openpyxl.styles import Font, Alignment

from .style import (SheetWriter, PCT0, PCT1, PCT2, NUM0, NUM1, NUM2, NUM3, MT, MTS, RS, CR,
                    X1, X2, USD, TEXT, SCORE, PPT)
from .support import na_row, est_row
from .spec import NA
from .estimates import (ESG_ENV, ESG_SOCIAL, ESG_GOV, ESG_CAPEX_SHARE, ESG_CAPEX_SHARE_BASIS,
                        ESG_CAPEX_MIX, LABOUR_PRODUCTIVITY, EMPLOYEES_FY26, CSR_RATE,
                        MACRO_BETA)

LEG = 'CDEFGHIJ'
FLEG = 'DEFGHIJ'
NEW8 = 'BCDEFGHI'
NEW7 = 'BCDEFGH'
MA8 = 'DEFGHIJK'
MA7 = 'EFGHIJK'
SCEN = ['Base Case', 'Bull Case', 'Bear Case', 'Stress Case']


def _sup(item, value, unit, method, formula, primary, secondary, assumption, reasoning,
         cross, conf, linked, freq, last, comments, numfmt=None):
    return dict(item=item, value=value, unit=unit, method=method, formula=formula, primary=primary,
                secondary=secondary, assumption=assumption, reasoning=reasoning, cross=cross,
                conf=conf, linked=linked, freq=freq, last=last, comments=comments, numfmt=numfmt)


# =================================================================== 17 Capital Allocation
NO_DISTRIB = (
    'The model does not forecast dividends, buybacks or M&A at industry level. Those are capital-structure '
    'and portfolio decisions taken company by company, and aggregating them would imply a single industry '
    'payout policy that does not exist. Every rupee of unlevered free cash flow is applied to net debt, '
    'which is the conservative and auditable treatment. Substitute a company payout policy when this model '
    'is used as the foundation for a single-name valuation.')


def capital_allocation(wb, audit, eng):
    ws = wb['Capital Allocation']
    w = SheetWriter(ws, audit)
    rows, nas = [], []
    NO_FY26 = ('The capex and free cash flow blocks are forecast-only from FY2027E, so no FY2026A capital '
               'allocation is asserted. See 16 Cash Flow Model for the reason.')

    # ---- executive summary rows 9-16
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        w.link(f'{nc}9', f"='Cash Flow Model'!${lc}$114", CR)
        w.link(f'{nc}10', f"='Cash Flow Model'!${lc}$107", CR)
        w.link(f'{nc}11', f"='Cash Flow Model'!${lc}$109", CR)
        if i == 0:
            # FY2026A has no net-debt roll-forward: row 77 opening net debt starts at FY2027E.
            w.f(f'{nc}12', '=0', CR)
            w.f(f'{nc}16', f'={nc}9-{nc}10-{nc}11', CR)
        else:
            w.f(f'{nc}12', f'={lc}77-{lc}80', CR)
            w.f(f'{nc}16', f'={lc}78-{lc}79-({lc}77-{lc}80)', CR)
        # Dividends, buybacks and M&A are DELIBERATE ZEROS, not gaps - see the register.
        for r in (13, 14, 15):
            w.inp(f'{nc}{r}', 0, CR)
    rows.append(est_row(
        'Dividends, share buybacks and M&A investment - deliberate zeros', 'B13:I15', NO_DISTRIB,
        'Three rows across eight years, blank and flagged as missing data. Nothing about them is unknown: '
        'this model applies every rupee of unlevered free cash flow to net debt, so distributions and '
        'acquisitions are zero BY CONSTRUCTION. Writing the zeros in makes the waterfall add up and makes '
        'the modelling decision visible on the sheet instead of looking like an omission. The same three '
        'lines appear in the waterfall on rows 26 to 28 and in the scenario table on column D, and all of '
        'them are now consistent zeros. Substitute a company payout policy when this model is used as the '
        'foundation for a single-name valuation.',
        unit='Rs cr', linked='This sheet', value='=B13', numfmt=CR,
        method='Deliberate zero - no distribution is modelled at industry level',
        formula='(entered zero)',
        primary='Modelling decision, not an estimate',
        secondary='Row 91 of this sheet states the same convention',
        cross='Consistent with the financing section of 16 Cash Flow Model, which is also zero',
        conf='High - this is a stated convention, not an uncertain quantity',
        freq='n/a - changes only if the model is given a payout policy'))

    # ---- capital allocation waterfall rows 21-31
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        w.f(f'{nc}21', f"='Cash Flow Model'!{lc}96-'Cash Flow Model'!{lc}102-'Cash Flow Model'!{lc}113", CR)
        w.f(f'{nc}22', f"=-'Cash Flow Model'!{lc}107", CR)
        w.f(f'{nc}23', f"=-'Cash Flow Model'!{lc}109", CR)
        w.f(f'{nc}24', f'=-({lc}77-{lc}80)', CR)
        w.f(f'{nc}25', f'=-{lc}79', CR)
        for r in (26, 27, 28):
            w.inp(f'{nc}{r}', 0, CR)
        # ESG investment is now modelled on 20 ESG Model row 46 and is shown here as a
        # memorandum: it sits INSIDE total capex, so it is not deducted again in the sum.
        w.link(f'{nc}29', f"='ESG Model'!{nc}46", CR)
        w.f(f'{nc}30', f'=SUM({nc}21:{nc}25)', CR)
        w.f(f'{nc}31', f'={nc}30', CR)
    rows.append(est_row(
        'ESG and decarbonisation investment in the waterfall', 'B29:H29',
        'Reads the ESG capex total modelled on 20 ESG Model row 46, which is itself a share of the total '
        'capex this model forecasts.',
        'This line was blank because ESG capex was not modelled anywhere. It is now populated as a '
        'MEMORANDUM line: decarbonisation spend sits inside the total capex already deducted on rows 22 and '
        '23, so it must NOT be subtracted again in the cash-retained sum on row 30 - doing so would '
        'double-count it and break the zero residual that proves the waterfall is complete. The line is '
        'shown because a reader wants to know how much of the capex is transition spend, not because it is '
        'an additional call on cash.',
        unit='Rs cr', linked='ESG Model', value='=H29', numfmt=CR,
        method='Cross-sheet reference to the modelled ESG capex total',
        formula="='ESG Model'!B46",
        primary='Modelled: share of forecast total capex - see the ESG Model register entries',
        secondary='Ministry of Steel net-zero-by-2070 roadmap; EU CBAM definitive regime',
        cross='EXCLUDED from the row 30 sum by design, because it is already inside rows 22 and 23. Row 30 '
              'must still be exactly zero in every year.',
        conf='Low - inherits the ESG capex share assumption',
        freq='Annually'))
    rows.append(_sup(
        'Capital allocation waterfall', '=H30', 'Rs cr',
        'Computed in-cell: operating cash flow less capex less debt service, with the residual as cash '
        'retained',
        '=SUM(B21:B25)',
        'Computed from the model',
        'Research block rows 77 to 80, the net debt roll-forward',
        'CASH RETAINED IS EXACTLY ZERO IN EVERY YEAR, and that is a mathematical consequence, not an '
        'assumption: net debt is defined as rolling forward by opening less unlevered free cash flow plus '
        'interest, so applying all free cash flow to debt leaves nothing over by construction.',
        'A zero residual is the proof that the waterfall is complete - no cash is created or destroyed '
        'anywhere in the statement. It also makes the model conservative: every rupee services debt, so the '
        'leverage trajectory is the best case for the balance sheet and the worst case for shareholders.',
        'Rows 21 to 25 must sum to zero in every year and in every scenario; check A22 on 25 Audit Checks '
        'tests the net debt roll-forward independently',
        'Medium', 'Cash Flow Model', 'Live', 'Live',
        'Was entirely blank. Signs follow the cash convention: outflows negative.', CR))

    # ---- capex allocation rows 36-42
    for i, nc in enumerate(NEW7):
        lc, ma, n7 = FLEG[i], MA7[i], NEW7[i]
        w.link(f'{nc}36', f"='Cash Flow Model'!${lc}$109", CR)
        w.f(f'{nc}37', f"='Capacity Forecast'!{n7}52*'Model Assumptions'!{ma}$23/10", CR)
        w.f(f'{nc}38', f"=('Capacity Forecast'!{n7}56-'Capacity Forecast'!{n7}52)"
                       f"*'Model Assumptions'!{ma}$23/10", CR)
        # Technology, ESG and safety are MEMORANDUM lines carved out of maintenance capex,
        # not additions to it: rows 36 to 38 already sum to total capex on row 42.
        w.est(f'J39', 0.08, PCT0)
        w.est(f'J40', 0.00, PCT0)
        w.est(f'J41', 0.03, PCT0)
        w.f(f'{nc}39', f"='Cash Flow Model'!${lc}$109*$J$39", CR)
        w.link(f'{nc}40', f"='ESG Model'!{nc}46", CR)
        w.f(f'{nc}41', f"='Cash Flow Model'!${lc}$109*$J$41", CR)
        w.link(f'{nc}42', f"='Cash Flow Model'!${lc}$110", CR)
    rows.append(est_row(
        'Technology, ESG and safety capex', 'B39:H41',
        'Technology capex modelled at 8% of maintenance capex and safety capex at 3%, with the shares in '
        'column J. ESG capex reads the modelled total on 20 ESG Model row 46. Digitalisation and automation '
        'programmes at the Indian majors run at single-digit percentages of sustaining capital expenditure, '
        'and statutory safety spend is smaller again.',
        'These three rows were blank because the model carries one all-in capex intensity and no producer '
        'publishes a capex split by category. They are now MEMORANDUM lines: rows 36 to 38 already sum '
        'exactly to total capex on row 42, so these three carve out slices of that same spend rather than '
        'adding to it. That distinction is why they can be populated without breaking the reconciliation - '
        'and it is stated here rather than left for a reader to infer, because a category table that '
        'silently double-counted would be worse than one left blank.',
        unit='Rs cr', linked='Cash Flow Model, ESG Model', value='=H39', numfmt=CR,
        method='Stated share of maintenance capex for technology and safety; ESG reads the ESG Model total',
        formula="='Cash Flow Model'!$D$109*$J$39",
        primary='Modelled shares of sustaining capital expenditure',
        secondary='ESG capex is modelled independently on 20 ESG Model as a share of total capex',
        cross='MEMORANDUM ONLY - excluded from the row 42 total, which must continue to equal rows 36 plus '
              '37 plus 38 and to tie to 16 Cash Flow Model row 110',
        conf='Low - no producer publishes a capex split by category',
        freq='Annually'))
    rows.append(_sup(
        'Capex allocation', '=H42', 'Rs cr',
        'Computed in-cell: maintenance from the research block, brownfield from the dated tracker via the '
        'capacity bridge, greenfield as the residual of net additions',
        "='Capacity Forecast'!B52*'Model Assumptions'!E$23/10",
        'Dated project tracker on 07 Capacity Expansion Tracker; capex intensity driver D16',
        '06 Capacity Forecast capacity bridge rows 51 to 57',
        'GREENFIELD here is net capacity additions less brownfield, so it also carries the secondary and '
        'unattributed additions of about 11 Mtpa a year.',
        'Defining greenfield as the residual is what makes the three rows sum to total capex exactly. The '
        'alternative - a separate line for secondary additions - does not exist in this table, and leaving '
        'the difference unallocated would break the reconciliation.',
        'Rows 36 + 37 + 38 must equal row 42 in every year, which must equal 16 Cash Flow Model row 110',
        'Medium', 'Capacity Forecast, Cash Flow Model, Model Assumptions', 'Quarterly', 'FY2027E',
        'Was entirely blank.', CR))

    # ---- debt allocation rows 47-53
    for i, r in enumerate(range(47, 54)):
        lc = FLEG[i]
        w.link(f'B{r}', f'=${lc}$77', CR)
        w.f(f'C{r}', f'=MAX(0,{lc}77-{lc}80)', CR)
        w.f(f'D{r}', f'=MAX(0,{lc}80-{lc}77)', CR)
        w.link(f'E{r}', f'=${lc}$80', CR)
        w.link(f'F{r}', f'=${lc}$84', X2)
    rows.append(_sup(
        'Debt allocation and leverage trajectory', '=E53', 'Rs cr',
        'Cross-sheet reference within this sheet to the net debt roll-forward, with repayment and drawdown '
        'separated by MAX so the direction of the movement is explicit',
        '=MAX(0,D77-D80) repaid, =MAX(0,D80-D77) drawn',
        'FY2026A closing net debt of Rs 1,85,001 cr derived from the Master Industry Database',
        'Research block rows 77 to 85; 25 Audit Checks A22',
        'All free cash flow after interest is applied to net debt. No new debt is raised for capex.',
        'The base case deleverages the industry from 1.07x net debt to EBITDA to a NET CASH position by '
        'about FY2030E, which is the single most important balance-sheet message in the model. It also '
        'explains why the illustrative interest line turns negative from FY2030E - that is interest income, '
        'not a sign error. In the Bear Case the industry re-levers instead, because free cash flow turns '
        'negative while capex continues.',
        'Opening less repaid plus drawn must equal closing in every year; leverage must equal net debt '
        'divided by EBITDA on research block row 84',
        'Medium', 'EBITDA Model, Cash Flow Model', 'Quarterly', 'FY2026A',
        'Was entirely blank.', CR))

    # ---- strategic investments rows 58-63
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        w.link(f'{nc}58', f"='Cash Flow Model'!${lc}$107", CR)
        # Renewable energy IS modelled, on 20 ESG Model row 41. The other four categories are
        # deliberate zeros: this model forecasts Indian crude steel capacity and nothing else.
        w.inp(f'{nc}59', 0, CR)
        w.link(f'{nc}60', f"='ESG Model'!{nc}41", CR)
        for r in (61, 62, 63):
            w.inp(f'{nc}{r}', 0, CR)
    rows.append(est_row(
        'Strategic investment other than capacity expansion', 'B59:H63',
        'Renewable energy investment reads the renewable category of the ESG capex block on 20 ESG Model row '
        '41. Downstream products, digital transformation, acquisitions and overseas expansion are '
        'deliberate zeros.',
        'Five rows across seven years, blank. Two different things were being conflated. Renewable energy '
        'spend IS now modelled, inside the ESG capex block, so this line links to it. The other four are '
        'not unavailable - they are OUT OF SCOPE, and a zero says that far more clearly than a blank does. '
        'The model forecasts Indian crude steel capacity only; acquisitions are explicitly excluded from '
        'the project tracker because they transfer ownership of existing capacity rather than add new '
        'capacity, and downstream, digital and overseas programmes are company portfolio choices that no '
        'producer costs to FY2033.',
        unit='Rs cr', linked='ESG Model, Cash Flow Model', value='=H60', numfmt=CR,
        method='Renewable energy links to the ESG capex block; the other four categories are deliberate '
               'zeros',
        formula="='ESG Model'!B41 for renewable energy; entered zero for the rest",
        primary='Modelled for renewable energy; a scope decision for the other four',
        secondary='07 Capacity Expansion Tracker states the exclusion of acquisitions',
        cross='Row 58 capacity expansion must equal growth capex on 16 Cash Flow Model row 107',
        conf='Low for renewable energy, High for the zeros - they are a scope statement',
        freq='Annually'))

    # ---- scenario analysis rows 68-71 (cumulative FY2027E-FY2033E)
    for r, sc in zip(range(68, 72), SCEN):
        w.link(f'B{r}', f'=SUM({eng.rng(sc, "gcapex")})', CR)
        w.link(f'C{r}', f'=SUM({eng.rng(sc, "ufcf")})', CR)
        w.inp(f'D{r}', 0, CR)
        w.inp(f'E{r}', 0, CR)
    rows.append(_sup(
        'Four-scenario capital allocation, cumulative FY2027E-FY2033E', '=C68', 'Rs cr',
        'Cross-sheet SUM across the seven forecast years of the scenario engine on 24 Scenario Manager',
        "=SUM('Scenario Manager'!$D$<row>:$J$<row>)",
        'Rebuilt live from the scenario driver matrix on 02 Model Assumptions',
        'Independent tie-out on 24 Scenario Manager row 79',
        'Cumulative rather than single-year, because capital allocation is a through-cycle question.',
        'The Bull Case generates roughly Rs 9.5 lakh crore of cumulative free cash flow against roughly Rs '
        '5.0 lakh crore of growth capex - comfortable self-funding. The Bear Case generates NEGATIVE '
        'cumulative free cash flow while still committing roughly Rs 3.5 lakh crore of growth capex, which '
        'must be debt-funded. That divergence, not the EBITDA number, is what determines whether the '
        'industry is investable through the cycle.',
        'Cash retained is exactly zero in every scenario by construction',
        'Medium', 'Scenario Manager', 'Live', 'Live', 'Was entirely blank.', CR))
    return rows, nas


# ================================================================= 18 Industry Cycle Model
# indicator -> (0-100 score formula template, band description)
CYCLE_SCORES = [
    (19, "=MIN(100,MAX(0,('Macroeconomic Model'!{lc}101-4%)/5%*100))",
     'Real GDP growth on a 4.0% to 9.0% band'),
    (20, "=MIN(100,MAX(0,('Steel Demand Model'!{lc}115-2%)/10%*100))",
     'Apparent consumption growth on a 2% to 12% band, which brackets the FY2021 contraction recovery and '
     'the FY2024 peak of 13.7%'),
    (21, "=MIN(100,MAX(0,('Capacity Utilisation'!{lc}67-65%)/27%*100))",
     'Capacity utilisation on a 65% to 92% band, the floor of the observed range to the practical ceiling'),
    (22, "=MIN(100,MAX(0,100-('Capacity Forecast'!{lc}101-'Steel Supply Model'!{lc}94)"
         "/'Capacity Forecast'!{lc}101/30%*100))",
     'Spare capacity against demand-driven requirement, 0% to 30%, scored INVERSELY'),
    (23, "=MIN(100,MAX(0,('Steel Price Forecast'!{lc}103-50000)/25000*100))",
     'Blended realisation on a Rs 50,000/t to Rs 75,000/t band'),
    (24, "=MIN(100,MAX(0,('Margin Analysis'!{lc}91+10%)/36%*100))",
     'Industry EBITDA margin on a -10% to +26% band, which is exactly the observed envelope from SAIL '
     'FY2016 to the FY2022 cycle peak'),
    (25, "=MIN(100,MAX(0,'Capacity Forecast'!{lc}100/25*100))",
     'Capacity additions on a 0 to 25 Mtpa band, against a demonstrated national delivery of about 20 Mtpa'),
    (26, "=MIN(100,MAX(0,50-'Steel Supply Model'!{lc}91/'Steel Demand Model'!{lc}116*1000))",
     'Variation in stock as a share of consumption, centred on 50 - a destock scores above 50 because it '
     'signals a tight market'),
]
PHASE = ('=IF({c}>=80,"Overheating",IF({c}>=60,"Peak",IF({c}>=40,"Expansion",'
         'IF({c}>=20,"Recovery","Deep Downturn"))))')


def cycle(wb, audit, eng):
    ws = wb['Industry Cycle Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- scorecard rows 19-26 (row 27 already SUMPRODUCTs)
    for r, tmpl, _desc in CYCLE_SCORES:
        for cyc, lc in zip(LEG, LEG):
            w.f(f'{cyc}{r}', tmpl.format(lc=lc), SCORE)

    # ---- executive summary rows 9-14
    for nc, lc in zip(NEW8, LEG):
        w.f(f'{nc}10', f'=${lc}$27', SCORE)
        w.f(f'{nc}9', PHASE.format(c=f'{lc}27'), TEXT)
        w.f(f'{nc}11', f'=IF({lc}27>=80,"Elevated - late cycle risk",IF({lc}27>=60,"Moderate",'
                       f'IF({lc}27>=40,"Balanced",IF({lc}27>=20,"High","Severe"))))', TEXT)
        w.f(f'{nc}12', f"=IF('Capacity Utilisation'!{lc}69>0.03,\"Strong\","
                       f"IF('Capacity Utilisation'!{lc}69>0,\"Improving\","
                       f"IF('Capacity Utilisation'!{lc}69>-0.04,\"Limited\",\"Weak\")))", TEXT)
        w.f(f'{nc}13', f"=IF('Margin Analysis'!{lc}91>{lc}90+0.04,\"Above mid-cycle\","
                       f"IF('Margin Analysis'!{lc}91>{lc}90,\"At mid-cycle\","
                       f"IF('Margin Analysis'!{lc}91>{lc}90-0.05,\"Below mid-cycle\",\"Trough\")))", TEXT)
        w.f(f'{nc}14', f"=IF({lc}27>=60,\"Late - discipline required\","
                       f"IF({lc}27>=40,\"Selective\",IF({lc}27>=20,\"Accumulate\",\"Deep value\")))", TEXT)
    rows.append(_sup(
        'Cycle score and phase', '=I10', 'score 0-100',
        'Weighted SUMPRODUCT of eight indicator scores, each normalised onto a 0-100 band computed live '
        'from the model',
        '=SUMPRODUCT($B$19:$B$26,C19:C26); each indicator =MIN(100,MAX(0,(x-floor)/band*100))',
        'Computed from the model. Every band is stated inside its own formula so it is auditable',
        'Research block rows 95 to 98, the observed cycle chronology from the Master Industry Database',
        'The BANDS are the modeller\'s framework; the INPUTS are all model outputs, and the margin band of '
        '-10% to +26% is taken directly from observed history rather than chosen.',
        'SIMPLIFIED. The clamp was written as MEDIAN(0,100,x), which is compact but reads as a statistic '
        'rather than as a limit and makes a reader stop and work out why a median of three numbers bounds a '
        'score. It is now MIN(100,MAX(0,x)) - "not below nought, not above a hundred", in the order you read '
        'it. Band edges are written as percentages rather than decimals so they match the units of the cells '
        'they are compared against. Anchoring the margin and utilisation bands on the observed range is what '
        'stops the scorecard from being arbitrary.',
        'Weights in column B sum to 1.00; the score must classify FY2026A as the recovery the research '
        'block documents on row 97',
        'Medium', 'Macroeconomic Model, Steel Demand Model, Capacity Utilisation, Capacity Forecast, '
        'Steel Price Forecast, Margin Analysis, Steel Supply Model',
        'Live', 'Live', 'Rows 19 to 26 and the whole executive summary were blank.', SCORE))

    # ---- leading indicators rows 41-47
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}41', f"='Macroeconomic Model'!${lc}$101", PCT1)
    # rows 42-47 now read the modelled macro sub-series on 03 Macroeconomic Model
    CYC_LEAD = {42: 14, 43: 15, 44: 16, 45: 17, 46: 18, 47: 20}
    for r, src in CYC_LEAD.items():
        for nc in NEW8:
            w.link(f'{nc}{r}', f"='Macroeconomic Model'!{nc}{src}", PCT1)
    rows.append(est_row(
        'Leading indicators other than GDP', 'B42:I47',
        'Each row links to the corresponding series on 03 Macroeconomic Model, where it is modelled as a '
        'documented beta to real GDP growth or, for WPI, as a spread to CPI.',
        'Six rows across eight years were blank here for the same reason they were blank on the macro sheet: '
        'no institution publishes an India path for these series to FY2033. Now that the macro sheet models '
        'them off the sourced GDP and CPI paths, this table simply reads them. That is the right '
        'relationship - the cycle sheet should never hold its own copy of a macro series - and it means a '
        'change to a beta on the macro sheet propagates here automatically.',
        unit='%', linked='Macroeconomic Model', value='=I42', numfmt=PCT1,
        method='Cross-sheet link to the modelled macro sub-series',
        formula="='Macroeconomic Model'!B14",
        primary='Modelled on 03 Macroeconomic Model as betas to the sourced RBI real GDP path',
        secondary='MOSPI and RBI publish all of these as actuals',
        cross='Must equal 03 Macroeconomic Model rows 14 to 18 and row 20 exactly',
        conf='Medium - modelled from a sourced driver', freq='Monthly, as each series prints'))

    # ---- coincident indicators rows 52-56
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}52', f"='Steel Supply Model'!${lc}$101", MT)
        w.link(f'{nc}53', f"='Steel Demand Model'!${lc}$116", MT)
        w.link(f'{nc}54', f"='Capacity Utilisation'!${lc}$67", PCT1)
        w.link(f'{nc}55', f"='Steel Price Forecast'!${lc}$103", RS)
        w.link(f'{nc}56', f"='EBITDA Model'!${lc}$101", RS)

    # ---- lagging indicators rows 61-65
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        w.link(f'{nc}61', f"='Capacity Forecast'!${lc}$100", MT)
        w.link(f'{nc}62', f"='Capital Allocation'!${lc}$80", CR)
        if i:
            w.link(f'{nc}64', f"='Capacity Expansion Tracker'!{MA7[i - 1]}$99", MT)
        else:
            w.na(f'{nc}64', 'The dated project tracker covers FY2027E onwards only. Every project with a '
                            'FY2026 commissioning date is already inside the 220.4 Mtpa base.')
        # dividend payout is a deliberate zero; employment now reads the ESG Model headcount
        w.inp(f'{nc}63', 0, PCT0)
        w.link(f'{nc}65', f"='ESG Model'!{nc}51", NUM0)
    rows.append(est_row(
        'Dividend payout and industry employment', 'B63:I63, B65:I65',
        'Dividend payout is zero because this model applies every rupee of unlevered free cash flow to net '
        'debt and takes no view on distribution policy. Employment reads the headcount modelled on 20 ESG '
        'Model row 51, which is indexed to finished steel production net of a 2% annual productivity gain '
        'and anchored on the Ministry of Steel figure of over 6 lakh employed.',
        'Two rows across eight years, blank. Both are now answered by work done elsewhere rather than by new '
        'assumptions: the payout zero is the same convention already stated on 17 Capital Allocation and in '
        'the financing section of 16 Cash Flow Model, and the employment series already exists on the ESG '
        'sheet. Linking rather than re-keying means the lagging-indicator block cannot disagree with either.',
        unit='% and employees', linked='ESG Model, Capital Allocation', value='=I65', numfmt=NUM0,
        method='Deliberate zero for payout; cross-sheet link for employment',
        formula="='ESG Model'!B51",
        primary='Payout: modelling convention. Employment: Ministry of Steel headcount, indexed to '
                'production.',
        secondary='17 Capital Allocation row 91 states the payout convention',
        cross='Employment must equal 20 ESG Model row 51 exactly and must grow more slowly than production',
        conf='High for the payout zero, Low for the employment trajectory', freq='Annually'))

    # ---- scenario analysis rows 70-71
    # SIMPLIFIED. This was one 550-character formula per scenario: eight MEDIAN clamps
    # multiplied by eight weights and concatenated with '+'. The eight indicator scores
    # are now eight ordinary rows inside the scenario engine, so each cell here is a
    # single cross-sheet reference and each indicator can be inspected on its own.
    for col, sc in zip('BCDE', SCEN):
        w.link(f'{col}70', eng.ref(sc, 'cycscore', 7), SCORE)
        w.f(f'{col}71', PHASE.format(c=f'{col}70'), TEXT)
    rows.append(_sup(
        'Four-scenario cycle score, FY2033E', '=B70', 'score 0-100',
        'Cross-sheet reference to the weighted cycle score computed inside the scenario engine, which scores '
        'the same eight indicators on the same bands as the live scorecard above',
        "='Scenario Manager'!$J$<engine cycle score row>",
        'Rebuilt live from the scenario engine on 24 Scenario Manager',
        'The scorecard above, which uses identical bands and weights on the live model',
        'Identical bands and weights to the live scorecard, so the four scenarios and the live model are '
        'measured on exactly the same ruler.',
        'This is the table that answers the only question an investment committee really asks: where in the '
        'cycle does each scenario leave us in the terminal year? The Bull Case reaches Peak, the base case '
        'sits in Expansion, and both Bear and Stress land in Recovery or below - which is the correct '
        'signal that a terminal-year exit multiple should NOT be a peak multiple in those states.',
        'Must use the same weights as row 27 and the same bands as rows 19 to 26; the base-case column must '
        'be close to the live FY2033E score, and identical when the Base Case is selected',
        'Medium', 'Scenario Manager', 'Live', 'FY2033E',
        'SIMPLIFIED. Each of these four cells was a 550-character formula holding eight weighted MEDIAN '
        'clamps. The eight indicator scores now sit in their own rows in the scenario engine, so this cell '
        'is a single reference, the arithmetic is visible one indicator at a time, and the weights are '
        'applied in one place instead of eight.', SCORE))

    # ---- valuation implications rows 76-80
    mult = {76: '$K$113', 77: '$K$112', 78: '$K$110', 79: '$K$111'}
    view = {76: 'Maximum operational and financial stress. Historically the best entry point, but only for '
                'balance sheets that survive it - the Stress Case takes industry margin to -9.2%.',
            77: 'Margin recovering off a trough with capacity discipline improving. Accumulate quality '
                'cost-curve positions.',
            78: 'The base case position for most of the horizon. Selective - returns come from cost '
                'position and product mix, not from the cycle.',
            79: 'Utilisation tight and margin at or above the FY2022 observed peak. Discipline required: '
                'this is when boards approve the capex that creates the next downcycle.',
            80: 'Not reached in any scenario in this model. Retained for completeness.'}
    for r in (76, 77, 78, 79):
        w.link(f'B{r}', f"='Model Assumptions'!{mult[r]}", X1)
    w.na('B80', 'No scenario in this model reaches an Overheating cycle score, so no exit multiple is '
                'asserted for that phase. The four multiples above are the model\'s own scenario exit '
                'multiples, driver D19.')
    # P/B and P/E by cycle phase, scaled off the peer medians on 23 Comparable Valuation in
    # the same proportion as the phase exit multiple bears to the base-case exit multiple.
    for r in (76, 77, 78, 79):
        w.link(f'C{r}', f"=B{r}/$B$78*'Comparable Valuation'!$E$55", X2)
        w.link(f'D{r}', f"=B{r}/$B$78*'Comparable Valuation'!$D$55", X1)
    for col in 'BCD':
        w.na(f'{col}80', 'No scenario in this model reaches an Overheating cycle score, so no multiple is '
                         'asserted for that phase. The four rows above are the model\'s own scenario exit '
                         'multiples, driver D19, and the equity multiples derived from them.')
    for r in range(76, 81):
        c = ws.cell(r, 5, view[r])
        c.font = Font(name='Calibri', sz=8, color='FF404040')
        c.alignment = Alignment(wrap_text=True, vertical='top')
    nas.append(na_row('Multiples for the Overheating cycle phase', 'B80:D80',
                      'No scenario in this model reaches an Overheating cycle score, so no multiple is '
                      'asserted for that phase.', unit='x'))
    rows.append(est_row(
        'Price to book and price to earnings by cycle phase', 'C76:D79',
        'Each phase multiple is the peer median price-to-book or price-to-earnings from 23 Comparable '
        'Valuation, scaled by the ratio of that phase\'s exit multiple to the base-case exit multiple. So '
        'the Peak row carries a 7.0x/6.0x uplift on the observed median and the Deep Downturn row a '
        '4.0x/6.0x discount.',
        'Ten blank cells, previously attributed to the absence of share prices. Share prices for the five '
        'listed majors DO exist on the sheet, and price to book now resolves because book value is derived '
        'by inverting return on equity - so both columns can be filled. Scaling the observed medians by the '
        'phase exit multiple keeps all three multiple columns internally consistent: they move together, in '
        'the same proportion, which is what a cycle-phase table is meant to show. The Overheating row stays '
        'blank because no scenario in this model reaches that score, and inventing a multiple for a state '
        'the model never visits would be worse than leaving it empty.',
        unit='x', linked='Comparable Valuation, Model Assumptions', value='=C78', numfmt=X2,
        method='Peer median multiple scaled by the ratio of the phase exit multiple to the base case',
        formula="=B76/$B$78*'Comparable Valuation'!$E$55",
        primary='Peer medians from 23 Comparable Valuation; exit multiples are driver D19 by scenario',
        secondary='Book value on 23 Comparable Valuation is derived by inverting return on equity',
        cross='The Expansion row must equal the peer median exactly, because Expansion IS the base case',
        conf='Low - inherits both the exit multiple, which is unsourced, and the ROE assumption',
        freq='Quarterly'))
    rows.append(_sup(
        'Exit multiple by cycle phase', '=B78', 'x EV/EBITDA',
        'Cross-sheet reference to the four scenario values of driver D19 on 02 Model Assumptions',
        "='Model Assumptions'!$K$110 (Base 6.0x) .. $K$113 (Stress 4.0x)",
        'NOT SOURCED - market multiples were not obtained in this research cycle',
        'Standing limitation 6 on 25 Audit Checks',
        'The four scenario exit multiples are MAPPED onto the four cycle phases: Stress 4.0x to Deep '
        'Downturn, Bear 5.0x to Recovery, Base 6.0x to Expansion, Bull 7.0x to Peak.',
        'The mapping is not arbitrary: each scenario already implies a cycle position, so re-using its own '
        'exit multiple keeps the valuation internally consistent instead of introducing a fifth set of '
        'numbers. It also makes the circularity explicit - the multiple is an assumption, not an '
        'observation, and until traded multiples are entered on 23 Comparable Valuation it should be '
        'treated as a placeholder for judgement rather than evidence.',
        'CONTRADICTION FLAGGED: the DCF on 23 Comparable Valuation cell D135 hard-codes 7.0x while driver '
        'D19 base case is 6.0x. That cell sits inside the research block and was NOT altered. It must be '
        'repointed to Model Assumptions before transaction use.',
        'Low', 'Model Assumptions', 'Quarterly', 'n/a',
        'Overheating is deliberately left blank - no scenario reaches it.', X1))
    return rows, nas



# ========================================================================= 19 Trade Model
TRADE_HIST = [(19, 'FY2021', 4.752, 10.784), (20, 'FY2022', 4.669, 13.494),
              (21, 'FY2023', 6.022, 6.716), (22, 'FY2024', 8.320, 7.487),
              (23, 'FY2025', 9.551, 4.858), (24, 'FY2026', 6.524, 6.602)]
EXPORT_DEST = {30: ('Europe', 0.344, "Italy 16.2% + Belgium 10.9% + Spain 7.3%"),
               31: ('Middle East', 0.074, 'UAE 7.4%'),
               32: ('Southeast Asia', 0.117, 'Vietnam 11.7%')}
IMPORT_SRC = {30: ('China', 0.235), 31: ('Japan', 0.202), 32: ('South Korea', 0.354)}


def trade(wb, audit, eng):
    ws = wb['Trade Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []

    # ---- executive summary rows 9-14
    for nc, lc in zip(NEW8, LEG):
        w.link(f'{nc}9', f'=${lc}$76', MT)
        w.link(f'{nc}10', f'=${lc}$78', MT)
        w.link(f'{nc}11', f'=${lc}$85', MTS)
        w.link(f'{nc}12', f'=${lc}$83', PCT1)
        w.link(f'{nc}13', f'=${lc}$84', PCT1)
        w.f(f'{nc}14', f"=({lc}78-{lc}76)*'Steel Price Forecast'!{lc}103/10", CR)

    # ---- historical trade rows 19-24
    for r, fy, imp, exp in TRADE_HIST:
        w.db(f'B{r}', imp, NUM3, ref=f"='[Master Industry Database.xlsx]Imports & Exports'!$C${r + 7}")
        w.db(f'C{r}', exp, NUM3, ref=f"='[Master Industry Database.xlsx]Imports & Exports'!$D${r + 7}")
        w.f(f'D{r}', f'=C{r}-B{r}', MTS)
        if r == 19:
            for col in 'EF':
                w.na(f'{col}{r}', 'FY2020 is outside the six-year window imported here, so a FY2021 growth '
                                  'rate cannot be computed. The Master Industry Database carries FY2015 '
                                  'onwards - FY2020 imports were 6.768 Mt and exports 8.355 Mt.')
        else:
            w.f(f'E{r}', f'=IFERROR(B{r}/B{r - 1}-1,"")', PCT1)
            w.f(f'F{r}', f'=IFERROR(C{r}/C{r - 1}-1,"")', PCT1)
    rows.append(_sup(
        'Historical trade FY2021-FY2026', '=D24', 'Mt net',
        'Master Industry Database import (RED - to be reconnected as an external link)',
        "='[Master Industry Database.xlsx]Imports & Exports'!$C$26:$D$31",
        'Joint Plant Committee, published by the Ministry of Steel',
        'DGCIS; Master Industry Database source S02',
        'Finished steel only. Semi-finished exports are reported separately and are excluded.',
        'THE MOST VOLATILE SERIES IN THE ENTIRE MODEL. India swung from a record net export of +8.8 Mt in '
        'FY2022 to a net IMPORT of -4.7 Mt in FY2025, then back to a marginal net export of +0.078 Mt in '
        'FY2026 as the safeguard duty took effect. A 13.5 Mt swing in four years on a 160 Mt market is why '
        'net trade is a scenario driver (D07) rather than a forecast.',
        'FY2026 imports of 6.524 Mt are down 31.7% year on year and exports up 35.9%; both are consistent '
        'with the April-2026 safeguard duty of 12%',
        'High', 'Master Industry Database', 'Monthly, on each Ministry of Steel report',
        'FY2026 (provisional)', 'Six years of history were blank.', MTS))

    # ---- trade forecast rows 19-25 (columns I..L)
    for i, r in enumerate(range(19, 26)):
        lc = FLEG[i]
        w.link(f'I{r}', f'=${lc}$76', MT)
        w.link(f'J{r}', f'=${lc}$78', MT)
        w.f(f'K{r}', f'=I{r}-J{r}', MTS)
        w.f(f'L{r}', f"=(J{r}-I{r})*'Steel Price Forecast'!{lc}103/10", CR)

    # ---- export destination and import source rows 30-35
    for r, (name, share, note) in EXPORT_DEST.items():
        w.db(f'D{r}', share, PCT1, ref=f"='[Master Industry Database.xlsx]Imports & Exports'! section C "
                                      f"({note})")
        w.f(f'B{r}', f'=$C$78*D{r}', NUM3)
        w.f(f'C{r}', f'=$J$78*D{r}', NUM3)
    for r in (33, 34):
        w.narow('BCD', r, 'North American and African destination shares are not separately disclosed by the '
                          'Ministry of Steel. The top-ten destination list for FY2026 comprises Italy, '
                          'Vietnam, Belgium, UAE, Spain, Nepal and Taiwan; neither region appears. Any '
                          'volume to those regions sits inside Others on row 35.')
    w.f('D35', '=1-SUM(D30:D34)', PCT1)
    w.f('B35', '=$C$78*D35', NUM3)
    w.f('C35', '=$J$78*D35', NUM3)
    for r, (name, share) in IMPORT_SRC.items():
        w.db(f'I{r}', share, PCT1, ref=f"='[Master Industry Database.xlsx]Imports & Exports'!$F$52:$F$56 "
                                      f"({name})")
        w.f(f'G{r}', f'=$C$76*I{r}', NUM3)
        w.f(f'H{r}', f'=$J$76*I{r}', NUM3)
    for r in (33, 34):
        w.narow('GHI', r, 'Vietnam was 7.3% of imports in FY2025 but is NOT separately disclosed for '
                          'FY2026; Indonesia is not disclosed in either year. Korea, China and Japan '
                          'together were 79.1% of FY2026 imports, so both sit inside the Others residual '
                          'on row 35.')
    w.f('I35', '=1-SUM(I30:I34)', PCT1)
    w.f('G35', '=$C$76*I35', NUM3)
    w.f('H35', '=$J$76*I35', NUM3)
    nas.append(na_row('Export shares to North America and Africa; import shares from Vietnam and Indonesia',
                      'B33:D34 and G33:I34', NA['trade_region'], unit='% and Mt'))
    rows.append(_sup(
        'Trade shares by country and region', '=D30', '% of exports',
        'Shares imported from the Master Industry Database (RED); volumes computed as total trade multiplied '
        'by share; Others is a residual so the shares always sum to 100%',
        '=$C$78*D30 and =1-SUM(D30:D34)',
        'Ministry of Steel Monthly Economic Report; Master Industry Database source S01',
        'DGCIS trade statistics',
        'FY2026 SHARES ARE HELD CONSTANT to FY2033E. No country-level trade forecast is attempted.',
        'Holding shares flat is disclosed rather than dressed up as a forecast, and it makes the one '
        'genuinely important number visible: EUROPE IS 34.4% OF INDIAN STEEL EXPORTS. That single figure is '
        'what makes the EU CBAM calculation on 20 ESG Model material, and it is why the EU share is also '
        'carried as a live input there rather than re-keyed.',
        'Shares must sum to exactly 100% by construction; the EU share of 34.4% must equal the value used '
        'on 20 ESG Model research block row 106',
        'High for the shares, Low for holding them flat',
        'Steel Price Forecast, ESG Model', 'Monthly for shares', 'FY2026 (provisional)',
        'Was entirely blank.', PCT1))

    # ---- trade drivers rows 40-47
    for i, nc in enumerate(NEW7):
        lc, ma = FLEG[i], MA7[i]
        w.link(f'{nc}40', f"='Steel Demand Model'!${lc}$116", MT)
        w.link(f'{nc}41', f"='Steel Supply Model'!${lc}$101", MT)
        w.link(f'{nc}43', f"='Steel Price Forecast'!${lc}$103", RS)
        w.link(f'{nc}44', f"='Model Assumptions'!{ma}$10", NUM1)
        w.link(f'{nc}46', f'=${lc}$88', PCT1)
        w.na(f'{nc}42', 'No global steel price forecast is made. The model carries point observations for '
                        'China, South East Asia and CIS at 22-Jun-2026 but forecasting regional prices is '
                        'outside its scope - see 09 Steel Price Forecast rows 59 to 63.')
        w.na(f'{nc}45', 'Freight is inside conversion cost, which is indexed to RBI CPI as a single block, '
                        'so no separate freight path exists in the model.')
        w.na(f'{nc}47', 'Anti-dumping duties are product and origin specific and no consolidated forward '
                        'schedule exists. The SAFEGUARD duty, which does have a notified schedule, is on '
                        'row 46 and in the research block row 88.')
    nas.append(na_row('Global steel prices, freight costs and anti-dumping duties',
                      'B42:H42, B45:H45, B47:H47',
                      'No global price or freight path is modelled; anti-dumping duties have no consolidated '
                      'forward schedule. The notified safeguard duty IS modelled on row 46.'))
    rows.append(_sup(
        'Safeguard duty schedule', '=B46', '%',
        'Cross-sheet reference to the notified policy schedule in the research block',
        '=$D$88',
        'Government of India safeguard duty notification; Master Industry Database policy record P03',
        '19 Trade Model research block row 88',
        'A SCHEDULE, NOT A FORECAST. 12% to 20-Apr-2026, 11.5% from 21-Apr-2026, 11% thereafter, expiring '
        '20-Apr-2028. Zero from FY2029E.',
        'THE SINGLE LARGEST IDENTIFIABLE POLICY RISK IN THE MODEL, and the one place where a policy cliff '
        'is explicitly carried. The duty is currently defending a domestic premium of roughly Rs 9,000/t '
        'over Chinese FOB, about 20%. The model does NOT reduce realisation when it expires in FY2029E, '
        'because whether the premium survives depends on whether the duty is renewed - which is a '
        'judgement, not a forecast. That is disclosed, and check A23 on 25 Audit Checks flags the year.',
        'Must match the notified schedule exactly; the FY2029E step to zero is deliberate and is the year '
        'to stress-test',
        'High for the schedule, n/a for its renewal', 'This sheet',
        'On any Government notification', '21-Apr-2026', 'Was blank.', PCT1))

    # ---- price competitiveness rows 52-56
    w.db('B53', 48580, RS, ref="='[Master Industry Database.xlsx]Steel Prices'!$I$27")
    for i, nc in enumerate(NEW8):
        w.link(f'{nc}52', f"='Steel Price Forecast'!${nc}$9", RS)
        if i:
            w.f(f'{nc}53', '=$B$53', RS)
        w.f(f'{nc}54', f'={nc}52-{nc}53', RS)
        w.f(f'{nc}55', f'=IF({nc}54<0,"Competitive - India below China",'
                       f'IF({nc}54<5000,"Broadly at parity","Uncompetitive on price"))', TEXT)
        w.f(f'{nc}56', f'=IF({nc}54>7500,"High - imports attractive",'
                       f'IF({nc}54>3000,"Moderate","Low"))', TEXT)
    rows.append(_sup(
        'India-China HRC price competitiveness', '=B54', 'Rs/t',
        'Computed in-cell from the India HRC series on 09 Steel Price Forecast and a held-flat China print',
        '=B52-B53',
        'IDBI Capital regional price snapshot, week ended 22-Jun-2026: India Rs 58,200/t, China Rs 48,580/t',
        'Master Industry Database KPI K19 records the spread at Rs 9,620/t; import-parity check in Apr-2026',
        'THE CHINA PRICE IS HELD FLAT across the horizon. No Chinese price forecast is attempted.',
        'Holding China flat while India moves with the model means the spread widens mechanically, and that '
        'is the honest way to present the risk rather than hide it: if Indian realisation rises to Rs '
        '69,000/t by FY2033E while China stays where it is, the import-parity gap becomes indefensible '
        'without a duty. Read this table as a stress indicator on the realisation driver, not as a forecast '
        'of Chinese prices.',
        'The FY2026A spread must reproduce the Rs 9,620/t recorded in the database; the classification rows '
        'must flip to "Uncompetitive" as the spread widens',
        'Medium for FY2026A, Low for the held-flat forecast', 'Steel Price Forecast', 'Weekly',
        '22-Jun-2026', 'Was entirely blank.', RS))

    # ---- scenario analysis rows 61-63
    for col, sc in zip('BCDE', SCEN):
        w.link(f'{col}61', eng.ref(sc, 'imp', 7), MT)
        w.link(f'{col}62', eng.ref(sc, 'exp', 7), MT)
        w.f(f'{col}63', f'={col}62-{col}61', MTS)
    rows.append(_sup(
        'Four-scenario trade position, FY2033E', '=B63', 'Mt net',
        'Cross-sheet reference to the scenario engine on 24 Scenario Manager',
        "='Scenario Manager'!$J$<engine row>",
        'Driver D07 net exports and D21 imports, by scenario, on 02 Model Assumptions',
        'Observed history: net +8.8 Mt (FY2022) to net -4.7 Mt (FY2025)',
        'Imports are the driver plus any shortfall routed from the capacity ceiling; exports are imports '
        'plus net exports.',
        'The FY2033E range runs from a net export of about +6.5 Mt in the Bear Case to a net IMPORT of about '
        '-5.0 Mt in the Stress Case. The Bear Case exports MORE because weak domestic demand pushes tonnes '
        'offshore - which is exactly what happened in FY2021 and FY2022. Every scenario value sits inside '
        'the observed FY2015-FY2026 range, which is the test that matters.',
        'FY2033E net trade must equal exports less imports; check A25 on 25 Audit Checks caps import '
        'penetration at 25%',
        'Medium', 'Scenario Manager, Model Assumptions', 'Live', 'Live', 'Was entirely blank.', MTS))
    return rows, nas


# =========================================================================== 20 ESG Model
def esg(wb, audit, eng):
    ws = wb['ESG Model']
    w = SheetWriter(ws, audit)
    rows, nas = [], []
    NO_FY26_CBAM = ('The CBAM calculation runs from FY2027E, which is when the EU definitive regime begins '
                    'charging. FY2026A carries no carbon cost because none was payable.')

    # ---- executive summary rows 9-14
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        w.link(f'{nc}10', f'=${lc}$102', NUM2)
        if i:
            w.link(f'{nc}11', f'=${lc}$112', CR)
        else:
            # A real zero, not a blank: no CBAM was payable in FY2026A. Blue, because it is
            # a modelling statement a user could change, and consistent with the FY2026A
            # ESG capex cell directly below it.
            w.est(f'{nc}11', 0, CR)
            rows.append(est_row(
                'FY2026A carbon cost', 'B11', NO_FY26_CBAM,
                'This was a blank flagged as missing data, which was the wrong classification. Nothing is '
                'unknown about it: the EU CBAM definitive regime begins charging in FY2027E, so the '
                'FY2026A carbon cost is zero as a matter of fact. A deliberate zero belongs in the cell as '
                'a blue input, exactly like the carbon credits row, not as a gap.',
                unit='Rs cr', linked='This sheet', value='=B11', numfmt=CR,
                method='Deliberate zero - no CBAM liability arose in FY2026A',
                formula='(entered zero)',
                primary='EU CBAM definitive regime start date',
                secondary='The CBAM calculation on research block rows 108 to 113 runs from FY2027E',
                cross='Consistent with the carbon cost forecast table, which starts at FY2027E',
                conf='High - this is a fact about the regime, not an estimate',
                freq='n/a', last='FY2026A'))
        w.link(f'{nc}9', '=$C$72', SCORE)
        w.link(f'{nc}13', f'=${nc}$21', PCT1)
        w.f(f'{nc}14', f'=IF({nc}20<$C$98,"Improving - qualifies as green",'
                       f'IF({nc}20<$C$99,"Improving","Stable - above the green threshold"))', TEXT)
        if i:
            w.link(f'{nc}12', f'=${NEW7[i - 1]}$46', CR)
        else:
            w.est(f'{nc}12', 0, CR)
    rows.append(est_row(
        'ESG executive summary - score, capex, renewable share and rating trend', 'B9:I9, B12:I14',
        'Every line here is now a link to the block on this sheet that computes it: the ESG score reads the '
        'scorecard on C72, ESG capex reads the capex total on row 46, renewable share reads row 21, and the '
        'rating trend is a formula that classifies emission intensity against the notified taxonomy '
        'thresholds on rows 96 to 99. FY2026A ESG capex is entered as zero because the capex block runs '
        'from FY2027E, matching the carbon cost line directly above it.',
        'Four executive-summary rows were blank because the four blocks beneath them were blank. Filling '
        'the blocks fixed the summary automatically, which is the point of linking rather than re-keying - '
        'the summary cannot now disagree with the detail.',
        unit='score, Rs cr, %', linked='This sheet', value='=C72', numfmt=SCORE,
        method='Cross-sheet links to the scorecard, capex and environmental blocks on this sheet',
        formula='=$C$72 for the score; =$B$46 for capex; =$B$21 for renewable share',
        primary='Computed from the blocks below on this sheet',
        secondary='Green steel taxonomy notified 23-Dec-2024 for the rating trend thresholds',
        cross='Must equal the blocks they read; the score must equal C72 exactly',
        conf='Low - inherits the confidence of the modelled blocks beneath',
        freq='Annually'))

    # ---- environmental KPIs rows 19-26
    for nc, lc in zip(NEW8, LEG):
        w.f(f'{nc}19', f"={lc}102*'Steel Supply Model'!{lc}101", NUM1)
        w.link(f'{nc}20', f'=${lc}$102', NUM2)
    # rows 21-26 MODELLED as an anchor (FY2026A, column B) and a target (FY2033E, column I)
    # with a linear glide between them. Two visible inputs per row, no hidden constants.
    for r, label, anchor, target, fmt, basis in ESG_ENV:
        w.est(f'B{r}', anchor, fmt)
        w.est(f'I{r}', target, fmt)
        for j, nc in enumerate('CDEFGH', start=1):
            w.f(f'{nc}{r}', f'=$B${r}+($I${r}-$B${r})*{j}/7', fmt)
        rows.append(est_row(
            f'Environmental KPI - {label}', f'B{r}:I{r}', basis,
            'The whole environmental KPI block below the emissions lines was blank because BRSR filings '
            'were not obtained in this research cycle. A blank row here is worse than a modelled one: the '
            'ESG scorecard on row 72 weights an environmental pillar at 0.50, and a pillar built on empty '
            'rows cannot be scored at all. Each row is now an FY2026A anchor in column B and an FY2033E '
            'target in column I, with a straight-line glide between them, so the two numbers worth arguing '
            'about are the only two entered and both are visible on the sheet.',
            unit='per tonne or %', linked='This sheet', value=f'=I{r}', numfmt=fmt,
            method='FY2026A anchor and FY2033E target entered; intervening years interpolated linearly',
            formula=f'=$B${r}+($I${r}-$B${r})*1/7',
            primary='Modelled estimate - see the Assumption column for the basis of each level',
            secondary='Company BRSR filings and sustainability reports publish all of these annually',
            cross='Feeds the environmental pillar score on row 69 and the ESG scorecard on row 72',
            conf='Low - modelled trajectories, not company disclosures',
            freq='Annually, on each BRSR filing cycle'))
    rows.append(_sup(
        'Industry CO2 emissions', '=I19', 'Mt CO2e',
        'Computed in-cell: emission intensity multiplied by finished steel production',
        "=C102*'Steel Supply Model'!C101",
        'Emission intensity is a USER INPUT of 2.55 tCO2e/tfs, explicitly NOT SOURCED (research block row '
        '102); production is from the model',
        'Green steel taxonomy thresholds notified 23-Dec-2024 (research block rows 96 to 99)',
        'Indian average intensity of 2.55 tCO2e/tfs held flat across the horizon - no decarbonisation '
        'trajectory is assumed.',
        'Holding intensity flat is a deliberately conservative and transparent choice. Indian producers '
        'have announced net-zero ambitions but none has published a dated, costed intensity pathway, so '
        'assuming improvement would flatter both the CBAM cost and the green-steel classification. On these '
        'numbers industry emissions rise from about 410 Mt CO2e to roughly 685 Mt by FY2033E purely because '
        'volume grows - which is the honest picture and the reason this is a first-order transition risk.',
        'Intensity of 2.55 sits ABOVE the 2.2 tCO2e/tfs threshold for 3-star green steel, so on these '
        'inputs no Indian average production qualifies as green',
        'Low - the intensity input is not sourced', 'Steel Supply Model', 'Annually, on BRSR filings',
        'n/a - user input', 'Rows 19 and 20 were blank. EDIT ROW 102 BEFORE USE.', NUM1))

    # ---- carbon cost forecast rows 31-35
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        w.f(f'{nc}31', f'={lc}104*{lc}105', NUM0)
        w.f(f'{nc}32', f'={lc}110*{lc}111', NUM2)
        w.link(f'{nc}33', f'=${lc}$112', CR)
        w.inp(f'{nc}34', 0, CR)
        w.f(f'{nc}35', f'={nc}33-{nc}34', CR)
    rows.append(_sup(
        'CBAM carbon cost', '=H35', 'Rs cr',
        'Computed in-cell: EU-destined exports multiplied by the emission intensity excess over the EU '
        'benchmark, multiplied by the EU carbon price converted to rupees',
        '=D110*D111*D104*D105/10 (research block row 112)',
        'EU share of Indian exports 34.4% IS sourced (Master Industry Database). The carbon price, both '
        'emission intensities and EUR/INR are USER INPUTS and NOT SOURCED',
        '19 Trade Model export destination shares; green steel taxonomy notified 23-Dec-2024',
        'CARBON CREDITS ARE SET TO ZERO. India has no operating compliance carbon market that would '
        'generate CBAM-eligible credits, and the Carbon Credit Trading Scheme is not yet in force.',
        'This is the only transition cost the model quantifies, and it is quantified precisely because the '
        'exposure is concentrated and measurable: 34.4% of exports go to the EU, and the intensity gap of '
        '0.70 tCO2e/tfs against the EU benchmark is chargeable. On the user inputs as delivered the cost is '
        'roughly Rs 1,400-2,100 crore a year - about 0.5% of industry EBITDA, material but not existential. '
        'FOUR OF THE FIVE INPUTS ARE UNSOURCED, so treat the magnitude as an order of estimate.',
        'Row 31 multiplied by row 32 divided by ten must equal row 33; row 33 must equal research block row '
        '112',
        'Low', 'Trade Model, Steel Supply Model', 'Quarterly, and on any EU ETS price move', 'n/a',
        'Rows 31 to 35 were blank. Carbon credits are blue because zero is a live user input.', CR))

    # ---- ESG capex rows 40-46.
    # Sized as a SHARE of the total capex the model already forecasts, so it scales with the
    # scenario instead of being an invented rupee schedule. Row 47 carries the share itself.
    for i, nc in enumerate(NEW7):
        lc = FLEG[i]
        w.est(f'{nc}47', ESG_CAPEX_SHARE[i], PCT0)
        w.f(f'{nc}46', f"='Cash Flow Model'!{lc}110*{nc}47", CR)
        for r, label, share, _why in ESG_CAPEX_MIX:
            w.est(f'J{r}', share, PCT0)
            w.f(f'{nc}{r}', f'=${nc}$46*$J${r}', CR)
    ws['A47'] = 'ESG capex as a share of total capex'
    rows.append(est_row(
        'ESG and decarbonisation capex', 'B40:H47', ESG_CAPEX_SHARE_BASIS,
        'This block was blank, and so were the ESG capex lines in the executive summary, the financial '
        'impact table and the scenario table that all read from it - five separate holes from one missing '
        'driver. It is now built the way the rest of the model builds capex: a share of a forecast total, '
        'not an absolute schedule. Total capex comes from 16 Cash Flow Model row 110, the share is on row '
        '47 of this sheet, and the category mix is in column J. Because it is a share, the Bear Case '
        'automatically spends less on decarbonisation than the Bull Case - which is what happens in '
        'practice and could not be represented by a hard-coded schedule.',
        unit='Rs cr', linked='Cash Flow Model', value='=H46', numfmt=CR,
        method='Total capex from 16 Cash Flow Model multiplied by the ESG share on row 47, then allocated '
               'across six categories using the mix in column J',
        formula="='Cash Flow Model'!D110*B47 for the total; =$B$46*$J$40 for each category",
        primary='Modelled: share of the capex the model forecasts',
        secondary='Ministry of Steel net-zero-by-2070 roadmap; EU CBAM definitive regime; green steel '
                  'taxonomy notified 23-Dec-2024',
        cross='Row 46 must equal the sum of rows 40 to 45, and the category shares in column J must sum '
              'to 100%',
        conf='Low - no producer publishes a dated, costed decarbonisation schedule',
        freq='Annually, or on any producer publishing a costed transition plan'))

    # ---- social metrics rows 51-55
    prod, prod_basis = LABOUR_PRODUCTIVITY
    emp, emp_basis = EMPLOYEES_FY26
    csr, csr_basis = CSR_RATE
    w.est('B51', emp, NUM0)
    w.est('J51', prod, PCT1)
    for i, (nc, lc) in enumerate(zip(NEW8, LEG)):
        if i:
            w.f(f'{nc}51', f"=$B$51*'Steel Supply Model'!{lc}101/'Steel Supply Model'!$C$101"
                           f"*(1-$J$51)^{i}", NUM0)
        w.f(f'{nc}55', f"=$J$55*('EBITDA Model'!{lc}103-'Cash Flow Model'!{lc}99)"
                       f"*(1-'Cash Flow Model'!{lc}101)", CR)
    w.est('J55', csr, PCT1)
    for r, label, anchor, target, fmt, basis in ESG_SOCIAL:
        w.est(f'B{r}', anchor, fmt)
        w.est(f'I{r}', target, fmt)
        for j, nc in enumerate('CDEFGH', start=1):
            w.f(f'{nc}{r}', f'=$B${r}+($I${r}-$B${r})*{j}/7', fmt)
        rows.append(est_row(
            f'Social metric - {label}', f'B{r}:I{r}', basis,
            'Modelled as an FY2026A anchor and an FY2033E target with a linear glide. These rows feed the '
            'social pillar score on row 70, which in turn is 0.25 of the overall ESG score - so leaving '
            'them blank left a quarter of the scorecard unstruck.',
            unit='rate or %', linked='This sheet', value=f'=I{r}', numfmt=fmt,
            method='FY2026A anchor and FY2033E target entered; intervening years interpolated linearly',
            formula=f'=$B${r}+($I${r}-$B${r})*1/7',
            primary='Modelled estimate - see the Assumption column',
            secondary='BRSR section A requires all of these disclosures annually',
            cross='Feeds the social pillar score on row 70',
            conf='Low - modelled trajectories, not company disclosures',
            freq='Annually, on each BRSR filing cycle'))
    rows.append(est_row(
        'Social metric - Employees', 'B51:I51', emp_basis + ' ' + prod_basis,
        'Headcount is modelled rather than entered year by year: it is indexed to finished steel '
        'production and then discounted by a productivity gain held in J51. That means it responds to the '
        'scenario - more tonnes means more people - and it cannot drift away from the production forecast, '
        'which a typed series would.',
        unit='employees', linked='Steel Supply Model', value='=I51', numfmt=NUM0,
        method='FY2026A headcount indexed to finished steel production, less the annual productivity gain',
        formula="=$B$51*'Steel Supply Model'!D101/'Steel Supply Model'!$C$101*(1-$J$51)^1",
        primary='Ministry of Steel: the sector employs over 6 lakh people',
        secondary='Company BRSR filings disclose headcount individually',
        cross='Must reproduce roughly 6 lakh in FY2026A and grow more slowly than production',
        conf='Medium for the FY2026A level, Low for the trajectory',
        freq='Annually'))
    rows.append(est_row(
        'Social metric - Community investment', 'B55:I55', csr_basis,
        'The only social line in the block with a statutory basis. Corporate social responsibility spend '
        'is 2% of average net profit under section 135 of the Companies Act, so this is modelled as a live '
        'formula off the profit line rather than entered as a level - it moves with earnings, and in a '
        'Bear Case it falls, which is exactly the real-world behaviour.',
        unit='Rs cr', linked='EBITDA Model, Cash Flow Model', value='=I55', numfmt=CR,
        method='Statutory CSR rate in J55 applied to profit after tax, built as EBITDA less depreciation, '
               'less tax',
        formula="=$J$55*('EBITDA Model'!C103-'Cash Flow Model'!C99)*(1-'Cash Flow Model'!C101)",
        primary='Companies Act 2013 section 135 - statutory 2% of average net profit',
        secondary='Company annual reports disclose actual CSR spend against the obligation',
        cross='Moves with the profit line, so it must fall in the Bear and Stress cases',
        conf='High for the rate, Medium for the profit base it is applied to',
        freq='Annually'))

    # ---- governance metrics rows 60-64
    for r, label, anchor, target, fmt, basis in ESG_GOV:
        w.est(f'B{r}', anchor, fmt)
        w.est(f'I{r}', target, fmt)
        for j, nc in enumerate('CDEFGH', start=1):
            w.f(f'{nc}{r}', f'=$B${r}+($I${r}-$B${r})*{j}/7', fmt)
        rows.append(est_row(
            f'Governance metric - {label}', f'B{r}:I{r}', basis,
            'Modelled as an FY2026A anchor and an FY2033E target with a linear glide. The independent '
            'director and board diversity lines are close to sourced, because they are effectively SEBI '
            'compliance floors for the listed majors; the ESG-linked pay and compliance score lines are '
            'the weakest in the block and are labelled Low confidence.',
            unit='%', linked='This sheet', value=f'=I{r}', numfmt=fmt,
            method='FY2026A anchor and FY2033E target entered; intervening years interpolated linearly',
            formula=f'=$B${r}+($I${r}-$B${r})*1/7',
            primary='Modelled estimate - see the Assumption column',
            secondary='SEBI Listing Obligations and Disclosure Requirements set the floors on rows 60-61',
            cross='Feeds the governance pillar score on row 71',
            conf='Medium for rows 60-61 (regulatory floors), Low for rows 62-63',
            freq='Annually, on each annual report'))
    # row 64 is a composite of the four governance lines above
    for nc in NEW8:
        w.f(f'{nc}64', f'=({nc}60/60%*25+{nc}61/30%*25+{nc}62/50%*25+{nc}63/100*25)', SCORE)
    rows.append(est_row(
        'Governance score', 'B64:I64',
        'A composite of the four governance lines above, each normalised against a stated benchmark and '
        'contributing 25 points: independent directors against 60%, board diversity against 30%, '
        'ESG-linked pay against 50% and the compliance score against 100.',
        'Row 64 was blank even though the four rows it should aggregate were also blank - a score with no '
        'inputs. It is now an explicit formula whose benchmarks are visible in the formula itself, so a '
        'reader can see precisely what "good" was taken to mean.',
        unit='score 0-100', linked='This sheet', value='=I64', numfmt=SCORE,
        method='Four normalised governance lines, 25 points each',
        formula='=(B60/60%*25+B61/30%*25+B62/50%*25+B63/100*25)',
        primary='Computed from the governance metrics above',
        secondary='SEBI LODR sets the regulatory floors the benchmarks are chosen around',
        cross='Feeds the governance pillar score on row 71',
        conf='Low - the benchmarks are a framework', freq='Annually'))

    # ---- ESG scorecard rows 69-72
    w.f('C69', '=MIN(100,MAX(0,($C$99-$C$102)/($C$99-$C$96)*100))', SCORE)
    # social pillar: safety, diversity and training, each normalised and equally weighted
    w.f('C70', '=(MIN(100,MAX(0,(0.5-$B$52)/0.4*100))+MIN(100,MAX(0,$B$54/25%*100))'
               '+MIN(100,MAX(0,$B$53/50*100)))/3', SCORE)
    w.link('C71', '=$B$64', SCORE)
    w.f('C72', '=SUMPRODUCT($B$69:$B$71,C69:C71)', SCORE)
    rows.append(est_row(
        'Social and governance pillar scores', 'C70:C71',
        'The social pillar is an equally weighted composite of three normalised FY2026A metrics: LTIFR '
        'scored inversely on a 0.10 to 0.50 band, female participation against a 25% benchmark, and '
        'training hours against a 50-hour benchmark. The governance pillar reads the composite governance '
        'score on row 64 directly.',
        'Both pillars were blank, which meant the overall ESG score on row 72 could not be struck at all '
        'even though its weights - 0.50 environmental, 0.25 social, 0.25 governance - sum to 1.00. Row 72 '
        'is now a plain SUMPRODUCT with no COUNT guard, because all three pillars resolve. The '
        'environmental pillar still scores zero, and that remains the important answer: at 2.55 tCO2e/tfs '
        'the Indian average sits above the 2.2 threshold at which steel stops qualifying as green under '
        'the taxonomy notified on 23-Dec-2024. A high social and governance score cannot rescue it, '
        'because environmental carries half the weight - which is the disclosure an investment committee '
        'needs.',
        unit='score 0-100', linked='This sheet', value='=C72', numfmt=SCORE,
        method='Normalised composites of the metrics above, clamped to 0-100 with MIN/MAX',
        formula='=(MIN(100,MAX(0,(0.5-$B$52)/0.4*100))+...)/3',
        primary='Computed from the modelled social and governance metrics on this sheet',
        secondary='Green steel taxonomy notified 23-Dec-2024 anchors the environmental pillar',
        cross='Weights in B69:B71 sum to 1.00; row 72 is their SUMPRODUCT with the three pillar scores',
        conf='Low - the underlying metrics are modelled and the benchmarks are a framework',
        freq='Annually'))
    rows.append(_sup(
        'Environmental pillar score', '=C69', 'score 0-100',
        'Computed in-cell: the position of Indian average emission intensity between the "not green" '
        'threshold and the 5-star threshold of the notified taxonomy, clamped to 0-100',
        '=MEDIAN(0,100,($C$99-$C$102)/($C$99-$C$96)*100)',
        'Green steel taxonomy thresholds notified 23-Dec-2024 - these ARE sourced (Master Industry Database '
        'policy record P15)',
        'Research block rows 96 to 99 and row 102',
        'Scored against the Government of India taxonomy rather than against an external ESG rating '
        'methodology.',
        'The score is ZERO, and that is the correct and important answer: at 2.55 tCO2e/tfs the Indian '
        'average is above the 2.2 threshold at which steel ceases to qualify as green at all. Using the '
        "Government's own notified taxonomy as the yardstick is far more defendable than an agency score, "
        "because the thresholds are law. For reference, Tata Steel's 0.75 Mtpa Ludhiana scrap EAF is the "
        'kind of asset that clears the 5-star threshold of 1.6.',
        'The overall score on row 72 stays blank until the social and governance pillars are entered, which '
        'is deliberate - a partial weighted score would be misleading',
        'Medium for the method, Low for the intensity input',
        'This sheet', 'On any taxonomy amendment', '23-Dec-2024', 'Was blank.', SCORE))

    # ---- ESG financial impact rows 77-81
    for i, nc in enumerate(NEW7):
        lc, ma = FLEG[i], MA7[i]
        w.link(f'{nc}77', f'=${nc}$46', CR)
        w.link(f'{nc}78', f'=${lc}$112', CR)
        w.f(f'{nc}79', f'=-{lc}112', CR)
        w.f(f'{nc}80', f'=-{lc}112', CR)
        w.f(f'{nc}81', f"=-{lc}112*'Model Assumptions'!{ma}$26", CR)
    rows.append(_sup(
        'ESG financial impact', '=H81', 'Rs cr',
        'Computed in-cell: the carbon cost flows one-for-one to EBITDA and free cash flow, and is '
        'capitalised at the scenario exit multiple to give the valuation impact',
        "=-D112*'Model Assumptions'!E$26",
        'CBAM calculation in the research block; exit multiple driver D19',
        '23 Comparable Valuation for the enterprise value the impact is measured against',
        'The carbon cost is treated as a permanent operating cost, so capitalising it at the exit multiple '
        'is the correct valuation treatment. ESG CAPEX IS NOT INCLUDED because it is not known.',
        'THE VALUATION IMPACT IS THEREFORE UNDERSTATED, and knowingly so. It captures the carbon cost but '
        'not the capital required to avoid it, which is the larger number: a scrap-EAF or hydrogen-ready '
        'route change costs materially more per tonne than Rs 55,000. The direction of the omission is '
        'disclosed here rather than buried.',
        'EBITDA impact must equal the negative of the CBAM cost; valuation impact must equal that '
        'multiplied by the exit multiple in force for the selected scenario',
        'Low', 'Model Assumptions, Comparable Valuation', 'Quarterly', 'n/a',
        'Rows 78 to 81 were blank; ESG capex remains blank by necessity.', CR))

    # ---- scenario analysis rows 86-89
    for r, sc in zip(range(86, 90), SCEN):
        w.f(f'B{r}', '=$D$104*$D$105', NUM0)
        w.link(f'C{r}', f'={eng.cell(sc, "capex", 7)}*$H$47', CR)
        w.link(f'D{r}', f'=-{eng.cell(sc, "exp", 7)}*$J$106*$J$111*$J$104*$J$105/10', CR)
    rows.append(est_row(
        'ESG capex by scenario, FY2033E', 'C86:C89', ESG_CAPEX_SHARE_BASIS,
        'The scenario ESG capex column was blank. It is now the scenario engine total capex for FY2033E '
        'multiplied by the same ESG share used on the live sheet, so the four scenarios and the live model '
        'are struck on one assumption. This is the payoff from modelling ESG capex as a share rather than '
        'as a schedule: the Bear Case, which builds less capacity, also spends less on decarbonisation, '
        'and no separate scenario assumption had to be invented to make that happen.',
        unit='Rs cr', linked='Scenario Manager', value='=C86', numfmt=CR,
        method='Scenario engine total capex for FY2033E multiplied by the FY2033E ESG capex share (H47)',
        formula="='Scenario Manager'!$J$<engine capex>*$H$47",
        primary='Modelled: share of the capex the scenario engine forecasts',
        secondary='Consistent with the live ESG capex block on rows 40 to 47',
        cross='The base-case column must equal the live FY2033E ESG capex total in H46',
        conf='Low', freq='Live'))
    rows.append(_sup(
        'Four-scenario CBAM impact, FY2033E', '=D86', 'Rs cr',
        'Computed in-cell on the scenario engine export volume, at the same unsourced carbon price and '
        'intensity inputs',
        "=-'Scenario Manager'!$J$<exports>*$J$106*$J$111*$J$104*$J$105/10",
        'Export volumes from the scenario engine; all price and intensity inputs are USER INPUTS',
        '19 Trade Model scenario trade position',
        'The carbon PRICE is identical in all four scenarios, because it is an unsourced user input and the '
        'model does not flex unsourced parameters to manufacture a scenario.',
        'The CBAM cost is therefore scenario-sensitive only through EXPORT VOLUME, which is the honest '
        'result: the Bear Case pushes more tonnes offshore and so carries the highest carbon cost, at '
        'roughly Rs 2,600 crore against Rs 1,400 crore in the base case. A user who wants to stress the '
        'carbon price should edit research block row 104, and every figure on this sheet and its valuation '
        'impact will move together.',
        'Export volumes must match 19 Trade Model row 62; the base-case column must reproduce the live '
        'FY2033E CBAM cost',
        'Low', 'Scenario Manager, Trade Model', 'Live', 'n/a',
        'Was blank. Identical carbon prices across scenarios are deliberate, not a copy error.', CR))
    return rows, nas
