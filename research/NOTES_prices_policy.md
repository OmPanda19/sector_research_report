# Research Notes — Prices & Government Policy

## PART 1 — IRON ORE PRICES

### 1A. World Bank Commodity Price Data ("The Pink Sheet") — PRIMARY, highest confidence
Series: "Iron ore, cfr spot".
DEFINITION (World Bank "Description" sheet, row 91): *Iron ore (any origin) fines, spot price,
c.f.r. China, 62% Fe* beginning December 2008; previously 63.5% Fe.
Row 90 adds: "Iron ore, spot in US dollar/dry ton and contract in US cents/dmtu."
=> IMPORTANT UNIT NOTE: the column header in the Pink Sheet workbook is legacy-labelled "($/dmtu)"
   but the spot series is in **US$ per dry metric tonne (dmt), 62% Fe, CFR China**. The workbook must
   label it US$/dmt, and disclose the Pink Sheet's own legacy header. Getting this wrong by a factor
   of 62 is a classic modelling error.
Files: CMO-Historical-Data-Annual.xlsx (A4 = "Updated on March 03, 2026"),
       CMO-Historical-Data-Monthly.xlsx (A4 = "Updated on July 02, 2026")
URL root: https://thedocs.worldbank.org/en/doc/74e8be41ceb20fa0da750cda2f6b9e4e-0050012026/related/

CALENDAR-YEAR averages (nominal US$/dmt):
2011 167.8 | 2012 128.5 | 2013 135.4 | 2014 97.0 | 2015 55.9 | 2016 58.4 | 2017 71.8 | 2018 69.8
2019 93.8 | 2020 108.9 | 2021 161.7 | 2022 121.3 | 2023 120.6 | 2024 109.4 | 2025 100.2

FISCAL-YEAR (Apr–Mar) averages, computed from the monthly series (n=12 months each):
FY2016 52.18 | FY2017 67.81 | FY2018 68.98 | FY2019 71.99 | FY2020 95.66 | FY2021 128.03
FY2022 155.52 | FY2023 117.20 | FY2024 119.91 | FY2025 103.97 | FY2026 100.52

FY2026 QUARTERLY averages: Q1 95.50 | Q2 100.10 | Q3 103.53 | Q4 102.93
Q1 FY2027 (Apr–Jun 2026): 105.17

MONTHLY series (US$/dmt), Apr-2025 to Jun-2026:
2025M04 97.2 | 2025M05 97.0 | 2025M06 92.3 | 2025M07 97.3 | 2025M08 99.7 | 2025M09 103.3
2025M10 103.5 | 2025M11 102.5 | 2025M12 104.6 | 2026M01 105.5 | 2026M02 98.8 | 2026M03 104.5
2026M04 106.1 | 2026M05 108.6 | 2026M06 100.8

### 1B. NMDC notified domestic prices — PRIMARY (BSE/NSE regulatory filings by NMDC Ltd)
Baila = Bailadila, Chhattisgarh. Prices are ex-mine and EXCLUDE royalty, DMF, NMET, cess and other
levies/taxes unless stated.
| Effective from | Baila Lump (65.5% Fe, 10–40 mm) Rs/t | Baila Fines (64% Fe, −10 mm) Rs/t | Note |
|---|---|---|---|
| 05-Apr-2026 | 5,300 | 4,500 | Hike of up to 11.1% |
| 06-May-2026 | 5,500 | 4,700 | +Rs 200/t on both grades |
| 10-Jul-2026 | 5,450 | 4,700 | Reported as a reduction; lump −Rs 50/t |
Inclusive-of-taxes-and-charges equivalents reported alongside the Jul-2026 revision:
Rs 6,745/t lump and Rs 5,815/t fines.
CAVEAT: the FY26 (Apr-2025 – Mar-2026) NMDC notification series was not fully reconstructed within
the research window. Populate the dated notifications verified above; mark earlier months
"Data Not Publicly Available — not sourced within research window" rather than interpolating.
FY26 NMDC operating context (JPC CPSE data): iron ore production 53.147 Mt (+20.59% YoY),
sales 50.231 Mt (+13.12%) — highest ever, first Indian mining company to reach ~53 Mt.

### 1C. Ministry of Steel domestic iron ore price reference — PRIMARY
MoS Monthly Economic Report for March 2026, para 10(i): "During the month of March'26, prices of
iron ore **lumps** stood at **Rs. 4,800/tonne**, showing a slight increase compared with the
previous month." MoS publishes a monthly chart of iron ore prices from April 2023.
NOTE: MoS Rs 4,800/t (Mar-26) vs NMDC notified lump Rs 5,300/t (from 05-Apr-26) — different
constructs (MoS is a market/average assessment for lumps; NMDC is a single producer's notified
ex-mine price for a specific 65.5% Fe grade). NOT a conflict; label both.

### 1D. Regional Indian spot (broker tracker — SECONDARY)
IDBI Capital / IDBI Direct "Commodity Price Update", 22-Jun-2026:
| Item | Level | 52-wk high | 52-wk low |
|---|---|---|---|
| Iron ore Odisha (INR/t) | 6,900 | 8,100 | 6,900 |
| Pellets ex-Barbil (INR/t) | 8,000 | 9,250 | 7,900 |
| Iron ore 62% import fines (US$/t) | 96 | 107 | 88 |
| Pellets (US$/t) | 144 | 150 | 111 |
IDBI Capital "Metals & Mining" sector update, 09-Apr-2026: in Mar-26 domestic iron ore prices
flat MoM at Rs 7,850/t; international +2.4% MoM to US$102/t. Manganese +2% MoM to Rs 21,125/t.
NOTE the internal inconsistency between the two IDBI publications' domestic iron ore levels
(Rs 7,850/t in Mar-26 vs Rs 6,900/t in Jun-26): consistent with a ~12% correction Apr–Jun 2026,
and the Jun-26 print equals the 52-week low. Plausible, but grade/location definitions differ
between the "domestic iron ore" and "Iron ore Odisha" quotes. Flag Confidence = Medium.

---

## PART 2 — COKING COAL PRICES

### 2A. Premium Hard Coking Coal (HCC), FOB Australia — MoS narrative, PRIMARY (government)
MoS Monthly Economic Report for March 2026, para 10(ii). MoS labels the chart
"Trend in prices of HCC Coking Coal f.o.b. Australia":
- Rose sharply during 2023, peaking at **US$354/t in October 2023**
- Declined steadily through 2024 to a low of **~US$175/t in March 2025**
- Fluctuated in early 2025, then moved up to **~US$246/t in February 2026**
- Declined to **~US$225/t in March 2026**

### 2B. Coking coal FOB Australia — broker tracker, SECONDARY
IDBI Capital / IDBI Direct "Commodity Price Update", 22-Jun-2026:
Coking coal (FOB Australia) **US$200/t**, flat WoW, +14.3% MoM, +16.3% YTD, +48.1% YoY;
52-week high US$217/t, 52-week low US$135/t.
IDBI Capital "Metals & Mining", 09-Apr-2026: coking coal flat MoM at **US$186/t** in Mar-26,
after declining 13% MoM in Feb-26.

### 2C. **CONFLICT C1 — coking coal benchmark level (MATERIAL; must be documented)**
For March 2026, MoS states ~US$225/t while IDBI Capital states US$186/t. For Feb-2026 MoS states
~US$246/t while IDBI implies ~US$214/t (US$186 after a 13% MoM decline). IDBI's 52-week HIGH of
US$217/t (as at Jun-2026) is BELOW MoS's Feb-2026 level of US$246/t — so the two series cannot be
the same assessment.
WHY: "coking coal FOB Australia" is not a single price. Differences arise from
(i) grade — Premium Low-Vol Hard Coking Coal vs. mid-vol / PCI / semi-soft;
(ii) index provider and methodology — Platts PLV HCC FOB Australia, Argus, Fastmarkets, IHS;
(iii) FOB Australia vs CFR India (freight and demurrage);
(iv) spot assessment vs monthly average vs quarterly settlement.
RESOLUTION: carry BOTH series in the Coking Coal Prices sheet as separately labelled rows
("Premium HCC FOB Australia — Ministry of Steel / JPC assessment" and "Coking coal FOB Australia —
IDBI Capital tracker"), do NOT blend them, and do NOT compute a single FY average across them.
Preferred series for modelling Indian integrated-mill cost: the MoS/JPC assessment, because it is
the government's own primary published benchmark for the Indian industry and is the series the
Ministry uses in its own monthly reporting.

### 2D. Australian THERMAL coal (NOT coking coal) — World Bank, PRIMARY, for energy-cost reference
DEFINITION (Pink Sheet Description sheet, row 4): "Coal (Australia), from February 2022, port
thermal, f.o.b. Newcastle, 6000 kcal/kg **futures** price. From 2015 to January 2022, port thermal,
f.o.b. Newcastle, 6000 kcal/kg spot price..."
=> MUST be labelled THERMAL. It is a common and serious error to use this World Bank series as a
   coking coal proxy. Included only as a power/energy input reference for DRI-EAF and captive power.
Calendar-year averages (US$/mt): 2011 121.4 | 2012 96.4 | 2013 84.6 | 2014 70.1 | 2015 58.9 |
2016 66.1 | 2017 88.5 | 2018 107.0 | 2019 77.9 | 2020 60.8 | 2021 138.1 | 2022 344.9 | 2023 172.8 |
2024 136.1 | 2025 108.4
Fiscal-year (Apr–Mar) averages (US$/mt): FY2016 55.30 | FY2017 73.88 | FY2018 93.90 | FY2019 105.19 |
FY2020 70.97 | FY2021 66.15 | FY2022 176.58 | FY2023 343.40 | FY2024 145.10 | FY2025 131.89 |
FY2026 111.50
FY2026 quarterly: Q1 104.00 | Q2 110.47 | Q3 109.27 | Q4 122.27 ; Q1FY27 135.43
Monthly Apr-25 to Jun-26: 98.6, 104.4, 109.0, 112.9, 112.2, 106.3, 107.5, 112.6, 107.7,
109.8, 118.4, 138.6, 130.9, 136.9, 138.5

---

## PART 3 — STEEL PRICES

### 3A. Domestic HRC price path FY2026 — ICRA (SECONDARY, rating agency research)
ICRA research (report IDs 6790 and 6940, icra.in):
- HRC fell sharply in late-2025, touching **Rs 46,000/MT in early December 2025**, amid excess
  supply over demand
- Prices **rebounded from January 2026** following **safeguard duty reinstatement** and higher
  coking coal costs, reaching **~Rs 53,800/MT by early February 2026**
- **Q4 FY2026: +~14% QoQ to Rs 57,700/MT by end-March 2026**, supported by higher coking coal
  and iron ore costs
- At the start of FY2027 (April 2026), domestic HRC traded at a **discount of ~US$29–45/MT to
  landed imports**, keeping imports uncompetitive and supporting domestic realisations
Earlier ICRA vintage: after the Safeguard Duty, prices "corrected to ~Rs 49,500/tonne by
September 2025 and ~Rs 46,000/tonne by November 2025" (an alternative vintage says
~Rs 47,000/tonne by mid-November 2025 — minor vintage difference, disclose).

### 3B. Domestic HRC / rebar — broker trackers (SECONDARY)
IDBI Capital "Metals & Mining", 09-Apr-2026: Mar-26 Indian steel price **+10.2% MoM to
Rs 59,500/t — highest in 45 months**; Chinese steel +7.5% MoM to US$500/t.
IDBI Capital "Commodity Price Update", 02-Feb-2026: Indian HRC +1.9% WoW to **Rs 53,500/t**,
tier-1 mills announcing hikes backed by import dynamics and safeguard duty.
IDBI "Commodity Price Update", 22-Jun-2026 snapshot:
| Item | Level | 52-wk high | 52-wk low |
|---|---|---|---|
| Indian HRC (INR/t) | 58,200 | 59,600 | 45,700 |
| China HRC FOB (INR/t) | 48,580 | 49,052 | 41,977 |
| SE Asia HRC CNF (INR/t) | 50,938 | 51,410 | 44,807 |
| CIS HRC FOB (INR/t) | 51,410 | 51,410 | 41,505 |
ETInfra (Mar-2026), "Domestic steel prices surge 18-25% on extended safeguard duty":
HRC **Rs 55,900/t** vs Rs 47,317/t in the December quarter; primary rebar **Rs 59,800/t**
vs Rs 47,615/t.
BigMint benchmark, 31-Jul-2026: rebar (IS 1786 Fe 550D, 12–32 mm, BF route), ex-Mumbai,
distributor-to-dealer, **excluding 18% GST**: **INR 50,900/t**, up INR 2,200/t WoW from
INR 48,700/t on 24-Jul-2026.
HDFC Securities (Q3FY26 preview, Jan-2026): in Q3FY26 domestic HRC fell 4.5% QoQ (INR 2.3k/MT)
and 1% YoY (INR 0.5k/MT); rebar also corrected.

### 3C. Ministry of Steel retail price series — PRIMARY (chart only)
MoS Monthly Economic Report publishes "Trend in Retail Steel Prices in Mumbai (Rs/t incl. GST)"
for TMT 10 MM, HR COILS 2.50 MM and CR COILS 0.63 MM, month-end from 31-Oct-2023 to 31-Mar-2026.
The values are presented only as a chart in the PDF; the underlying month-end numbers are not
tabulated in the report. => Record the series' existence, definition, frequency and source in the
Steel Prices sheet with values marked "Data Not Publicly Available in tabular form (chart only);
obtain from JPC on request".

### 3D. **CONFLICT C2 — Indian HRC level, March 2026**
ICRA Rs 57,700/t (end-March) vs IDBI Capital Rs 59,500/t (Mar-26 monthly) vs ETInfra Rs 55,900/t.
WHY: differing (i) assessment point (month-end vs monthly average), (ii) location/basis
(ex-works vs ex-Mumbai trade), (iii) GST inclusion, (iv) grade/thickness, (v) primary-mill vs
secondary/trade quotes.
RESOLUTION: carry all three, labelled with provider and basis. For modelling, ICRA's end-March
Rs 57,700/t is the recommended reference because ICRA explicitly states the assessment point
(end-March) and the QoQ construction, making it reproducible. Note that none of these are GST-
inclusive retail prices and none are directly comparable to the MoS Mumbai retail series.

---

## PART 4 — GOVERNMENT POLICY & REGULATION

### 4A. SAFEGUARD DUTY ON FLAT STEEL — the single most important FY26 policy driver
Chronology (verified):
1. **21-Apr-2025** — Ministry of Finance imposed a **12% provisional safeguard duty for 200 days**
   on five categories of steel flat products, including hot-rolled coils, sheets and plates
   (non-alloy and alloy steel flat products). Motivation: surge in imports amid trade diversion
   following steep US duties on Chinese goods. [The Hindu, Indian Express, ETInfra 21-Apr-2025]
   200 days from 21-Apr-2025 expires ~07-Nov-2025.
2. **18-Aug-2025** — **DGTR final findings** recommended a definitive safeguard duty for
   **three years**, tapering: **12% year 1, 11.5% year 2, 11% year 3**. DGTR observed
   "a recent, sudden, sharp and significant increase in imports of the product under
   consideration". Industry groups were divided (primary producers in favour; MSME/user
   industries against). [DGTR final findings; S&P Global 18-Aug-2025; TOI, ET]
3. **~Nov-2025** — provisional duty lapsed on expiry of the 200-day period. Domestic HRC fell to
   ~Rs 46,000/MT in early Dec-2025 on excess supply. [ICRA]
4. **30-Dec-2025** — Ministry of Finance notification imposed the **definitive 3-year safeguard
   duty**: **12% from 21-Apr-2025 to 20-Apr-2026; 11.5% from 21-Apr-2026 to 20-Apr-2027;
   11% from 21-Apr-2027 to 20-Apr-2028**, with a **provision for a mid-term review**.
   [MoF notification 30-Dec-2025; Argus 31-Dec-2025; Livemint 30-Dec-2025; ET]
   Effect: the year-1 window back-fills the provisional period, so the levy is continuous on paper
   while the *collection* gap around Nov–Dec 2025 drove the price dip and subsequent rebound.
5. **Jan-2026 onwards** — prices rebounded on "safeguard duty reinstatement" plus higher coking
   coal, reaching ~Rs 53,800/MT by early Feb-2026 and Rs 57,700/MT by end-Mar-2026. [ICRA]
6. Scope confirmed by Ministry of Steel in **Rajya Sabha USQ 3350, answered 20-Mar-2026**:
   safeguard duties on import of "**Non-Alloy and Alloy Steel Flat Products**" for a period of
   **three years**, on goods falling under **tariff headings 7208, 7209, 7210, 7211, 7212, 7225
   and 7226**.
FY27 rate now in force from 21-Apr-2026: **11.5%**.

### 4B. QUALITY CONTROL ORDERS (QCO) / BIS
- Governing instrument: **Steel and Steel Products (Quality Control) Order, 2024**. MoS describes
  it as "banning sub-standard/defective steel products in domestic market as well as imports".
  [Rajya Sabha USQ 3350, 20-Mar-2026]
- **~20-Nov-2025** — Steel Ministry extended exemption from mandatory quality-norm compliance for
  certain steel and stainless-steel grades **until March 2026**, to preserve availability of
  critical steel products while domestic capacity ramps up. In the same action the Ministry
  decided that **steel grades not covered by any QCO no longer require clarification or NOC from
  the Ministry of Steel**; for QCO-covered grades, importers must ensure the manufacturer holds a
  BIS licence for those grades. [ET 20-Nov-2025; TOI]
- **Feb-2026** — BIS published "Amendments in Quality Control Order".
  https://www.bis.gov.in/wp-content/uploads/2026/02/Amendments-in-Quality-Control-Order.pdf
- **Amendments (2026)** — enforcement of BIS standards **deferred by three years** from date of
  publication for **mild steel used in metal arc welding electrodes** and **steels for die blocks
  used in drop forging**, under the Steel and Steel Products (QCO) Order, 2024. [Livemint]
- **27-Apr-2026 (order date)** — Ministry of Steel **SUSPENDED the QCO** that mandated use of only
  BIS-certified items for **various stainless steel products**, to ease the compliance burden on
  domestic producers, especially MSMEs. Stainless-steel-sector MSMEs subsequently flagged an
  import surge and urged the Government to **reintroduce** the QCO. [The Hindu, ET — Jun/Jul-2026]
  CORROBORATION FROM A PRIMARY COMPANY SOURCE: Jindal Stainless MD Abhyuday Jindal, on the
  exchange-filed 4QFY26 earnings call (05-May-2026), called "the temporary suspension of QCO
  ... a matter of concern and ... a discouraging setback for quality-focused domestic industry
  players", noting inferior imported material continuing to enter India at scale.
- **25-Jun-2026** — Ministry of Commerce & Industry press release announcing the **Transition
  Facilitation (Quality Control) Order, 2026** ("Transition QCO"): an **alternative risk-based
  compliance mechanism**, allowing eligible companies to obtain BIS certification through a
  **simplified registration-based mechanism** instead of the conventional inspection-intensive
  process; eases QCO rules across several sectors for **five years**. [The Hindu; Livemint]

### 4C. PLI SCHEME FOR SPECIALTY STEEL — PRIMARY (Parliament)
Source: **Lok Sabha Unstarred Question No. 5098, answered 24-Mar-2026**, Ministry of Steel
(Minister of State Shri Bhupathiraju Srinivasa Varma).
URL: https://sansad.in/getFile/loksabhaquestions/annex/187/AU5098_BI0By2.pdf?source=pqals
| Round | Projects | Companies |
|---|---|---|
| Round 1 | 44 | 19 |
| Round 2 | 42 | 25 |
| Round 3 | 85 | 60 |
- Committed investment across all three rounds: **Rs 55,993 crore**
- Investment realised so far: **Rs 23,827 crore** (42.6% of committed)
- Specialty steel capacity created: **24 million tonnes** in coated/plated, high-strength,
  alloy steel and electrical steel categories
- Contribution to import substitution: incremental production valued at **Rs 6,000 crore**
Scheme design history:
- Original PLI for Specialty Steel: budgetary outlay **Rs 6,322 crore**; five product categories —
  Coated/Plated Steel Products; High Strength/Wear-Resistant Steel; Specialty Rails; Alloy Steel
  Products & Steel Wires; Electrical Steel. Round 1 outcome as announced: Rs 27,106 crore
  committed investment, 14,760 direct jobs, ~7.90 Mt of specialty steel production.
- **PLI 1.1** launched **06-Jan-2025** (applications 06–31 Jan 2025) by Union Minister
  H.D. Kumaraswamy; same Rs 6,322 crore allocation retained; **reduced investment thresholds**
  for certain categories and provision to **carry forward excess production** for incentive claims;
  production period **FY2026 to FY2030**. Steel Secretary Sandeep Poundrik noted the changes aimed
  to increase participation in sub-categories that received no applications in the previous round.
- **PLI 1.2** — as at Feb-2026, secured **Rs 11,887 crore** of investment commitments, expected to
  add **8.7 Mt** of specialty steel capacity by **FY2031**. [ET 09-Feb-2026]
- MoS expectation: specialty steel production to reach **42 Mt by end of FY2027**
  [MoS "An Overview of Steel Sector", Mar-2026 and Jun-2026]

### 4D. OTHER POLICY INSTRUMENTS IN FORCE (source: Rajya Sabha USQ 3350, 20-Mar-2026 unless noted)
- **National Steel Policy 2017** — roadmap to FY2031; targets **300 Mtpa** crude steel capacity.
  Against this, capacity was 220.4 Mt in FY2026. [MoS Overview]
- **DMI&SP Policy** (Domestically Manufactured Iron & Steel Products) — "Made in India" steel
  preference in Government procurement.
- **Steel Quality Control Order** — see 4B.
- **Countervailing Duty (CVD)** in place on **Welded Stainless Steel Pipes and Tubes from China
  and Vietnam**.
- **Union Budget 2026-27** steel measures:
  (a) **Nil Basic Customs Duty continued on Ferro-Nickel**;
  (b) BCD exemption **extended up to 31-Mar-2028** on **Ferrous Scrap**, **Magnesium Oxide (MgO)
      coated cold-rolled steel coils** for use in manufacture of cold-rolled grain-oriented (CRGO)
      steel, and **specified goods for the manufacture of CRGO steel**.
- **Steel export duty** — 15% export duty on finished steel (imposed May-2022) was **completely
  scrapped with effect from 19-Nov-2022** via **Notification 58/2022-Customs**, returning the
  tariff on finished steel to **NIL (0%)**, following industry pushback and a 15–20% correction in
  domestic prices. [MoS "Development of Indian Steel Sector since 2014-15", Jun-2026 vintage]

### 4E. GREEN STEEL / DECARBONISATION
- **Taxonomy for Green Steel notified 23-Dec-2024** by the Ministry of Steel — described as the
  world's first national green steel taxonomy — to provide standards for defining and categorising
  low-emission steel. [Rajya Sabha USQ 3350, 20-Mar-2026 confirms the 23-Dec-2024 notification date]
- Threshold and star rating (emission intensity, tCO2e per tonne of finished steel, tfs):
  | Rating | Emission intensity |
  |---|---|
  | 5-star green steel | < 1.6 tCO2e/tfs |
  | 4-star green steel | 1.6 – 2.0 tCO2e/tfs |
  | 3-star green steel | 2.0 – 2.2 tCO2e/tfs |
  | Not "green steel" | > 2.2 tCO2e/tfs |
  Headline definition: green steel = emission intensity **below 2.2 tCO2e/tfs**.
  [S&P Global 13-Dec-2024; ET 13/20-Dec-2024]
- A **37% government procurement** share for green steel was proposed. [ET Explains, Dec-2024]
  Treat as PROPOSAL, not enacted mandate — Confidence Medium.
- **Guideline for Green Steel Certification** dated **30-Dec-2025**, referencing the notified
  2.2 tCO2e/tfs threshold (published via NISST).
- **CBAM**: the EU Carbon Border Adjustment Mechanism entered its **definitive regime from
  1 January 2026** — confirmed from a primary corporate source: Tata Steel's 4QFY2026 press
  release notes "In Europe, while import safeguards and roll out of the Carbon Border Adjustment
  Mechanism from 1st January has improved pricing conditions...". CBAM is a live export-competitiveness
  issue for Indian mills; MoS's answer to RS USQ 3350 was specifically framed around CBAM.
- Company net-zero commitments (primary, company disclosures): Tata Steel **2045**;
  Jindal Steel **2047**.

### 4F. OTHER FY26/FY27 TRADE DEVELOPMENTS AFFECTING INDIAN PRODUCERS
- **UK import quotas** — changes announced **March 2026**, expected by Tata Steel to "bring greater
  balance" to the UK market. [Tata Steel 4QFY26 press release]
- **EU import safeguards** plus CBAM improved European pricing conditions. [Tata Steel 4QFY26]
- **Netherlands regulatory action** — Dutch Environment Agency and Province issued a letter on
  **23-Apr-2026** indicating intention to **revoke operating permits** and trigger early closure of
  Tata Steel Netherlands' coke and gas plants; TSN paid **>EUR 20m of penalties in FY2026**;
  **material uncertainty related to going concern** disclosed in TSN's financial statements.
  [Tata Steel 4QFY26 press release — PRIMARY]
- **West Asia conflict** — from Q4FY26, developments in West Asia pressured supply chains, energy,
  oil, freight and currency markets; Tata Steel and Jindal Stainless both flagged continuing
  pressure into FY2027. Jindal Stainless specifically cited constrained availability of propane,
  LPG, natural gas and ammonia, plus shipping-lane diversions. [Company disclosures — PRIMARY]

### 4G. CONFLICT C3 — safeguard duty "reinstatement" vs "continuous" levy
Some secondary coverage describes the 30-Dec-2025 action as an "extension"/"reinstatement" while the
notification's year-1 window runs from 21-Apr-2025, implying continuity.
RESOLUTION: both are correct in different senses — the *provisional* measure lapsed after 200 days
(~07-Nov-2025) and the *definitive* measure was notified on 30-Dec-2025 with a year-1 period dated
from 21-Apr-2025. Present the chronology (4A) rather than a single characterisation, and note that
the observable market effect (HRC to Rs 46,000/MT in early Dec-2025, then a rebound from Jan-2026)
is consistent with an effective gap in the levy.
