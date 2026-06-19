# Dixon Technologies — 3-Statement Financial Model & DCF Valuation

> **Bottom-up equity research project** built as a WiDS 5.0 Financial Modelling Bootcamp portfolio piece, targeting Summer 2027 finance internship applications (FP&A · Corp Finance · Equity Research · IB · VC).
>
> **Analyst:** Om Chaudhari | IIT Bombay, B.Tech (Year 2) | June 2026

---

## Project at a Glance

| Item | Detail |
|------|--------|
| **Company** | Dixon Technologies (India) Ltd. — NSE: DIXON · BSE: 541988 |
| **Sector** | Electronics Manufacturing Services (EMS) · Consumer Electronics · PLI |
| **Model Type** | Fully integrated 3-Statement Model (IS + BS + CF) + FCFF DCF Valuation |
| **Historical Data** | FY22A – FY25A (4 years, Consolidated) · Source: Annual Reports (BSE/NSE) |
| **Forecast Period** | FY26E – FY30E (5-year explicit forecast) |
| **Currency / Units** | ₹ Indian Rupee · Crore (₹ Cr) · Indian FY: April 1 – March 31 |
| **Valuation Method** | FCFF DCF · WACC via CAPM · Terminal Value via Gordon Growth Model |
| **Cross-checks** | EV/EBITDA multiple · 2-way Sensitivity (WACC × TGR) · Peer COMPS |
| **Formula Count** | 634 live Excel formulas · 134 hardcoded inputs · 0 formula errors |
| **Tabs** | 8 (Cover · ASSUMP · IS · BS · CF · DCF · SENSITIVITY · COMPS) |

---

## Repository Structure

```
dixon-dcf-model/
│
├── Dixon_3Statement_DCF_Model_v2.0.xlsx    ← MAIN DELIVERABLE — open in Excel
│
├── README.md                               ← You are here
├── requirements.txt                        ← Python dependencies (openpyxl)
├── .gitignore
├── LICENSE
│
├── scripts/
│   ├── build_model_v2.py                  ← Python builder (WiDS Week 5 automation)
│   └── audit_model.py                     ← Formula integrity audit script
│
├── data/
│   └── data_sources.md                    ← Raw data, source URLs, methodology notes
│
└── docs/
    ├── resume_bullets.md                  ← Copy-ready resume bullets (FP&A / ER / IB / VC)
    └── interview_qa.md                    ← 15 interview Q&As (walkthrough → follow-ups)
```

---

## Why Dixon Technologies?

Dixon is India's largest **Electronics Manufacturing Services (EMS)** company — a pure-play on three simultaneous structural tailwinds:

**1. Production-Linked Incentive (PLI) Scheme**
Dixon is approved under PLI for Mobile Phones (₹6,500 Cr incentive pool, FY22–FY26) and IT Hardware (approved FY24). PLI requires manufacturing above volume thresholds and rewards incremental production — directly visible in Dixon's FY25 revenue jump (+120% YoY to ₹38,860 Cr).

**2. China+1 Supply Chain**
Global OEMs — Samsung, Motorola, Google, OnePlus, HP — are routing more production through India to reduce China concentration. Dixon is first-mover with established OEM relationships and government-backed CAPEX support.

**3. India's Electronics Manufacturing Ambition**
India targets $300 Bn electronics output by 2026. Dixon operates at the nexus of domestic consumption growth and export-oriented manufacturing.

**What makes the modeling interesting:** EMS businesses operate on thin EBITDA margins (4–6%) and high working capital. Understanding how a 3% PAT margin company creates significant enterprise value — through revenue scale and operating leverage — requires a different mental model than a typical FMCG or software company.

---

## Workbook Structure — 8 Tabs

| # | Tab | Contents |
|---|-----|----------|
| 1 | **Cover** | Company overview · key metrics FY22–FY25 · workbook navigation · color legend · data sources |
| 2 | **ASSUMP** | All blue input cells: revenue growth rates FY26–30 · EBITDA% · D&A% · tax rate · pre-tax Kd · capex% · NWC% · shares outstanding · Rf · ERP · β · Ke · WACC · TGR · source URLs |
| 3 | **IS** | Consolidated Income Statement (P&L) · FY22A–FY30E · plus full **Ratio Analysis** section (Profitability · Efficiency · Leverage) |
| 4 | **BS** | Consolidated Balance Sheet · formula-linked to IS and CF |
| 5 | **CF** | Cash Flow Statement (indirect method) · CFO / CFI / CFF · closing cash feeds into BS |
| 6 | **DCF** | WACC derivation (CAPM) at top → FCFF model → PV of cash flows → Terminal Value → EV bridge → **Implied Share Price** → EV/EBITDA sanity check |
| 7 | **SENSITIVITY** | 2-way table: Implied Share Price vs WACC (11%–16%) × TGR (4%–7%) · base case highlighted |
| 8 | **COMPS** | Dixon vs Amber Enterprises · Kaynes Technology · Syrma SGS — Revenue · EBITDA% · EV/EBITDA · P/E · peer median/mean · positioning notes |

### Color Coding Convention

| Color | Meaning |
|-------|---------|
| 🔵 **Blue text** | Hardcoded inputs — source from Annual Reports / Damodaran / RBI. Edit only these. |
| ⚫ **Black text** | Formula cells — calculated automatically. Do not edit. |
| 🟢 **Green text** | Cross-sheet links — pulling values from another tab. |
| 🟡 **Yellow fill** | Key assumption requiring verification before presenting. |

> **Rule:** If you want to change a forecast, change only a blue cell in the **ASSUMP** tab. Every other tab recalculates automatically.

---

## Historical Financial Summary (Real Data)

All historical figures from Dixon Technologies Consolidated Annual Reports (BSE/NSE filings).

| Metric | FY22A | FY23A | FY24A | FY25A | Source |
|--------|-------|-------|-------|-------|--------|
| Revenue (₹ Cr) | 10,697 | 12,192 | 17,691 | 38,860 | AR FY22–FY25 |
| YoY Growth | — | +14% | +45% | +120% | Calculated |
| EBITDA (₹ Cr) | 430 | 510 | 708 | 1,943 | AR |
| EBITDA Margin | 4.0% | 4.2% | 4.0% | 5.0% | Calculated |
| D&A (₹ Cr) | 115 | 140 | 185 | 280 | AR |
| EBIT (₹ Cr) | 315 | 370 | 523 | 1,663 | Calculated |
| Interest (₹ Cr) | 35 | 45 | 85 | 120 | AR |
| PAT (₹ Cr) | 190 | 253 | 365 | 1,215 | AR |
| PAT Margin | 1.8% | 2.1% | 2.1% | 3.1% | Calculated |
| EPS (₹) | 31.7 | 42.2 | 60.9 | 202.8 | Calculated |
| ROE | 17% | 19% | 21% | 42% | Calculated |

---

## Forecast Assumptions (FY26E–FY30E)

All assumptions are in the **ASSUMP tab** (blue cells). Key rationale:

### Revenue Growth
| Year | Growth | Rationale |
|------|--------|-----------|
| FY26E | +30% | PLI mobile ramp, guided ₹50,000 Cr+; new IT hardware vertical |
| FY27E | +25% | Continued EMS expansion; management targets ₹1,00,000 Cr+ by FY27 |
| FY28E | +18% | Growth normalisation; strong base effect kicks in |
| FY29E | +15% | Stable market share gains; new category contributions |
| FY30E | +12% | Mature phase; diversified revenue base |

### Margin Assumptions
- **EBITDA Margin: 5.2%** (all forecast years) — modest expansion from PLI income accrual and operating leverage; conservative vs. sell-side (some analysts model 6–7%)
- **D&A: 0.8% of Revenue** — asset-light EMS model; equipment mostly customer-funded or leased
- **Tax Rate: 25%** — new tax regime effective FY22; stable
- **Capex: 2.0% of Revenue** — light relative to scale; new capacity for IT hardware vertical
- **NWC: 4.5% of Revenue** — lean working capital; Dixon runs tighter cycles than peers

---

## Valuation Summary

### WACC (CAPM Framework)

| Parameter | Value | Source |
|-----------|-------|--------|
| Risk-Free Rate (Rf) | 7.00% | India 10-yr G-Sec yield — RBI (June 2026) |
| Equity Risk Premium (ERP) | 6.00% | Damodaran India ERP — Jan 2026 update |
| Beta (β) — Levered | 1.40 | 2-yr weekly regression vs Nifty 50 |
| **Cost of Equity (Ke)** | **15.4%** | Rf + β × ERP = 7.0% + 1.40 × 6.0% |
| Pre-tax Cost of Debt (Kd) | 8.50% | Repo rate + ~200bps spread |
| After-tax Kd | 6.38% | 8.5% × (1 − 25%) |
| Debt Weight D/(D+E) | 20% | Dixon is equity-heavy (FY25 BS) |
| **WACC** | **≈ 13.6%** | Ke × 80% + Kd(at) × 20% |
| Terminal Growth Rate (g) | 5.50% | Conservative India nominal GDP proxy |

### DCF Formula Chain

```
Revenue × EBITDA% = EBITDA
EBITDA − D&A = EBIT
EBIT × (1 − Tax Rate) = NOPAT

FCFF = NOPAT + D&A − Capex − ΔNWC

EV = Σ [FCFF_t / (1+WACC)^t]  +  [FCFF_5 × (1+g) / (WACC−g)] / (1+WACC)^5
                                         ↑ Terminal Value (Gordon Growth)

Equity Value = EV − Total Debt + Cash − Minority Interest
Implied Price = Equity Value × 100 / Shares Outstanding
```

### Why DCF Implies < Market Price (and Why That's the Right Answer)

The DCF-implied price is below Dixon's market price (~₹15,000 in June 2026). This is not an error — it is the most interview-worthy insight in the project:

1. **FCFF is near-zero in early years**: Dixon is reinvesting heavily — every rupee of new revenue requires more inventory, receivables, and capacity. The business consumes cash during its high-growth phase; it does not generate it.

2. **Terminal value dominates (~85–90% of EV)**: The entire DCF value rests on post-FY30 cash flows. A 5-year model structurally undercaptures the value of a growth company.

3. **Market uses multiples, not DCF, for EMS**: Analysts value Dixon on EV/Revenue or EV/EBITDA against peers — the market prices in optionality (Apple volumes, EV components, defense) that a base-case DCF does not.

> This pattern is identical to how analysts handle Foxconn, Jabil, or Flex globally. Understanding it is what separates a model-builder from a model-understander.

### 2-Way Sensitivity (WACC × TGR)

|  WACC ↓ / TGR → | 4.0% | 4.5% | 5.0% | **5.5%** | 6.0% | 6.5% | 7.0% |
|------------------|------|------|------|----------|------|------|------|
| 11.0% | | | | | | | |
| 12.0% | | | | | | | |
| 13.0% | | | | | | | |
| **13.6%** | | | | **★ Base** | | | |
| 14.0% | | | | | | | |
| 15.0% | | | | | | | |
| 16.0% | | | | | | | |

*See SENSITIVITY tab for exact implied prices. Cells auto-calculate from live DCF formulas.*

---

## Peer Comparables (FY25A, ₹ Crore)

| Company | Revenue | EBITDA | EBITDA % | PAT | MCap | EV | EV/EBITDA | P/E |
|---------|---------|--------|----------|-----|------|----|-----------|-----|
| **Dixon Technologies** | 38,860 | 1,943 | 5.0% | 1,215 | 75,000 | 75,200 | 38.7x | 61.7x |
| Amber Enterprises | 5,200 | 442 | 8.5% | 148 | 20,000 | 20,600 | 46.6x | 135.0x |
| Kaynes Technology | 1,800 | 270 | 15.0% | 175 | 14,000 | 14,150 | 52.4x | 80.0x |
| Syrma SGS Technology | 2,800 | 196 | 7.0% | 100 | 6,500 | 6,850 | 34.9x | 65.0x |

**Key positioning:** Dixon is the only large-cap, high-volume EMS play. Lower margin than Kaynes (PCBA-focused, higher value-add) but 20x larger revenue. Trades at a discount to Kaynes/Amber on EV/EBITDA — partially justified by lower margin, partially a function of the market pricing in revenue scale over margin quality.

---

## Ratio Analysis (IS Tab — WiDS Week 1)

The IS tab includes a full ratio analysis section covering:

**Profitability:** Gross Margin · EBITDA% · EBIT% · PAT% · ROE · ROA · ROCE

**Efficiency (days):** Inventory Days · Receivable Days · Payable Days · Asset Turnover

**Leverage:** Debt/Equity · Interest Coverage · Net Debt/EBITDA

All ratios are formula-linked to IS and BS — change any assumption in ASSUMP and ratios update automatically.

---

## How to Use This Model

1. **Open** `Dixon_3Statement_DCF_Model_v2.0.xlsx` in Microsoft Excel *(Google Sheets works but some formatting may differ)*
2. **Start on ASSUMP tab** — every editable input is here, colored blue or yellow
3. **To change the revenue forecast**: Edit ASSUMP rows 5–9 (FY26–FY30 growth rates)
4. **To change valuation assumptions**: Edit ASSUMP rows 24–26 (Rf, ERP, β) or row 35 (TGR)
5. **WACC auto-updates** at ASSUMP row 31 — no manual entry needed
6. **Check DCF tab** (row 56) for updated implied share price
7. **Check SENSITIVITY tab** to see price range across WACC/TGR assumptions
8. **Do not edit** black or green formula cells — they recalculate from ASSUMP inputs

---

## Python Builder (WiDS Week 5 — Automation)

The Excel file was generated programmatically using `scripts/build_model_v2.py` and the `openpyxl` library. This satisfies the WiDS Week 5 Python automation module.

```bash
# Install dependency
pip install openpyxl

# Regenerate the Excel model from scratch
python scripts/build_model_v2.py

# Audit formula integrity
python scripts/audit_model.py
```

The Python script writes actual Excel formula strings (e.g. `=IS!F18*(1-ASSUMP!B12)`) into cells — the output is a standard `.xlsx` file fully compatible with Excel and Google Sheets.

---

## Skills Demonstrated

- **3-Statement Integration**: IS drives BS (retained earnings, NFA), BS drives CF (WC changes), CF closes BS (cash balance) — all via formula links, zero hardcoding in derived cells
- **DCF Mechanics**: FCFF = NOPAT + D&A − Capex − ΔNWC; discounted at CAPM-derived WACC; Gordon Growth terminal value; EV→equity bridge
- **WACC via CAPM**: India-specific risk parameters (RBI Rf, Damodaran ERP); levered beta; after-tax cost of debt; capital structure weighting
- **Ratio Analysis**: Profitability, efficiency (working capital days), and leverage ratios — all formula-linked
- **Sensitivity Analysis**: 2-way data table sensitising implied price across WACC and TGR
- **Peer Comparables**: EV/EBITDA and P/E benchmarking against Indian EMS peer set
- **Excel Best Practices**: Assumption separation (single ASSUMP tab), color-coding convention (blue/black/green), no circular references
- **Python / openpyxl**: Programmatic Excel generation with cell-level formula writing, styling, and formula audit scripting
- **Sector Analysis**: EMS business model understanding (thin margins, PLI scheme, working capital intensity, China+1 tailwinds)

---

## WiDS 5.0 Curriculum Coverage

| WiDS Week | Topic | Covered in |
|-----------|-------|------------|
| Week 1 | 3-Statement Model + Ratio Analysis | IS, BS, CF tabs + IS Ratio section |
| Week 2 | Revenue Forecasting + Assumption Drivers | ASSUMP tab + IS forecast rows |
| Week 3 | WACC/CAPM + 2-Stage DCF | DCF tab (Section A + B + C + D) |
| Week 4 | Peer Comps + Relative Valuation | COMPS tab + DCF EV/EBITDA check |
| Week 5 | Python Automation | `scripts/build_model_v2.py` |

---

## Data Sources

| Source | Used For | URL |
|--------|----------|-----|
| Dixon AR FY25 (BSE) | Revenue, EBITDA, PAT, BS, CF | [BSE Filing](https://www.bseindia.com/bseplus/AnnualReport/541988/10124541988.pdf) |
| Dixon AR FY24–FY22 | Historical financials | BSE filings archive |
| Screener.in | Consolidated P&L / BS / CF cross-check | [screener.in/company/DIXON](https://www.screener.in/company/DIXON/consolidated/) |
| Trendlyne | Ratios, segment data, management guidance | [trendlyne.com/DIXON](https://trendlyne.com/equity/1624/DIXON/) |
| Damodaran Online | India ERP (6.0%) | [NYU Damodaran](https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html) |
| RBI | India 10-yr G-Sec yield (7.0% Rf) | [RBI](https://www.rbi.org.in/Scripts/BS_NSDPDisplayGenericDetails.aspx?Id=219) |
| NSE India | Shares outstanding, market cap | [NSE DIXON](https://www.nseindia.com/get-quotes/equity?symbol=DIXON) |
| Motilal Oswal / ICICI Direct | Broker estimates (FY26–FY30 guidance context) | Research portal (subscription) |

---

## Important Disclaimer

> This model is an **academic/learning exercise** by a 2nd-year undergraduate student. It is **not investment advice**. All data is sourced from publicly available filings. Forecast assumptions are the author's own and do not represent any broker's views. Always verify figures against actual company disclosures before use.

---

*Questions or feedback: omchaudhari19125@gmail.com*
*Built with Python (openpyxl) following industry color-coding conventions*
