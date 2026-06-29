# Dixon Technologies (India) Ltd. — 3-Statement Financial Model & DCF Valuation

> **Bottom-up equity research project** · WiDS 5.0 Financial Modelling Bootcamp · Summer 2026 finance internship portfolio  
> **Analyst:** Om Chaudhari | IIT Bombay, B.Tech (Year 2) | June 2026

---

## Quick Summary

| | |
|---|---|
| **Company** | Dixon Technologies (India) Ltd. — NSE: DIXON · BSE: 541988 |
| **Model** | Fully integrated 3-Statement (IS + BS + CF) + FCFF DCF — **v3.0** |
| **Historical** | FY21A – FY26A · Consolidated · **Single source: Screener.in** |
| **Forecast** | FY27E – FY31E (5-year explicit period) |
| **WACC** | 13.65% · Ke = 15.4% · Kd(at) = 6.63% · D-weight = 20% |
| **Implied Price** | **₹2,950 / share** · TGR = 5.5% · Gordon Growth terminal value |
| **Tabs** | 8 — Cover · ASSUMP · IS · BS · CF · DCF · SENSITIVITY · COMPS |
| **IS Expense Lines** | Material Cost · Employee Benefits · Other Opex → Total Expenses → EBITDA |
| **Currency** | ₹ Crore · Indian FY (Apr 1 – Mar 31) |

---

## Files in This Repo

```
Dixon-dcf-valuation/
│
├── Dixon_3Statement_DCF_Model_v3.0.xlsx   ← Main deliverable — open in Excel
├── Dixon_Analysis_Dashboard_v3.png        ← 6-panel Python dashboard (analysis_v3.py output)
├── build_model_v3.py                      ← Generates the Excel file
├── analysis_v3.py                         ← Python DCF + matplotlib dashboard
├── requirements.txt                       ← Python dependencies
└── README.md
```

---

## What Changed in v3.0 (vs the broken v2.0)

v2.0 had three critical structural errors that caused the Balance Sheet to never balance and the implied valuation to be wrong:

**1. EBITDA was overstated by ~28% in FY25**  
v2.0 used EBITDA = ₹1,943 Cr (5.0% margin) for FY25. Screener.in's "Operating Profit" — which correctly excludes Other Income — is ₹1,515 Cr (3.9% margin). The old model silently added Other Income (₹497 Cr in FY25, inflated by PLI receipts) into EBITDA. Fixed: Other Income is now a separate IS line; EBITDA margin forecasts are based on the correct ~4% Screener baseline.

**2. CF ΔNWC was disconnected from the Balance Sheet (root cause of BS not balancing)**  
v2.0 used `CF ΔNWC = −ΔRevenue × 4.5%` — a completely separate formula with no link to the Balance Sheet. Meanwhile, the BS computed NWC via independent % drivers (Inventory 7.2% of revenue, Receivables 10.8%, etc.). These two NWC calculations never produced the same number, so the BS balance check was always non-zero. Fixed: CF ΔNWC is now `= −(BS.NWC[t] − BS.NWC[t−1])` — a direct cross-sheet reference. The BS, IS, and CF are truly integrated.

**3. Potential circular reference in interest expense**  
v2.0 partially used beginning-year debt for interest (avoiding full circularity), but this was inconsistent. v3.0 strictly uses `Interest[t] = Debt[t−1] × Kd` — beginning-of-year debt in every period, including the first forecast year which references FY26A actuals.

**4. Single data source**  
v2.0 pulled from Annual Reports, Trendlyne, Damodaran, RBI, and broker estimates. v3.0 uses Screener.in consolidated financials exclusively for all historical data (FY21–FY26).

---

## IS Expense Breakdown (added post-audit)

The Income Statement now shows a full bottom-up cost structure between Revenue and EBITDA:

```
Revenue (Net Sales)
  YoY Revenue Growth %
  ── Operating Expenses ──
  Material Cost          ← Screener historical (89–93% of total expenses); fcst: 92%
  Employee Benefit Exp.  ← Screener historical (1–2%); fcst: 1.5%
  Other Operating Costs  ← Balancing item (incl. mfg., inventory changes) = Total − Mat − Emp
Total Operating Expenses ← Formula: = Revenue − EBITDA  (all years, hist & fcst)
EBITDA  (Operating Profit) ← Revenue × EBITDA margin % (from ASSUMP tab)
  EBITDA Margin %
D&A
EBIT  = EBITDA − D&A
  EBIT Margin %
Add: Other Income  (non-operating)
Less: Interest Expense
PBT  (Profit Before Tax)
Income Tax
PAT  (Net Profit)
```

**Design choice:** EBITDA is still the primary margin driver (ASSUMP-driven); Total Expenses and individual cost lines are derived from it. "Other Operating Costs" is always the balancing item so all lines sum exactly to Total Expenses in every year. The dashboard Panel 2 has been updated to show this cost structure as a stacked chart.

**Full IS row order fix:** Other Income (non-operating) was moved from between EBITDA and D&A to after EBIT — the correct accounting convention. The EBIT formula (`= EBITDA − D&A`) and PBT formula (`= EBIT + OI − Interest`) are unchanged.

---

## Audit Fixes Applied (v3.0 → post-audit)

An independent formula audit of the model (Dixon_DCF_Audit_Report.docx, June 29 2026) identified four issues, all corrected in this version:

**Bug #1 — CRITICAL: Other Income double-counted in PBT (fixed)**  
The original EBIT formula included Other Income (`EBIT = EBITDA + OI − D&A`), then PBT added it again (`PBT = EBIT + OI − Interest`). This inflated PAT by ₹372–435 Cr per forecast year (₹734 Cr overstatement in FY26A vs Screener actuals). Fix: **EBIT = EBITDA − D&A** (standard definition; OI is non-operating). PBT formula `= EBIT + OI − Interest` now correct. This also fixes NOPAT, which is derived from EBIT and feeds directly into FCFF.

**Bug #2 — MODERATE: Minority Interest hardcoded at ₹150 Cr vs ₹50 Cr actuals (fixed)**  
FY26A Screener MI = ₹50 Cr; the model had ₹150 Cr in forecast years with no note or cash-flow justification. Reverted to ₹50 Cr across all years and the DCF bridge.

**Warning #3 — Debtor Days aligned to FY26A actuals (fixed)**  
Model used 60 days vs FY26A implied 50 days (₹6,695 Cr / ₹48,873 Cr × 365 = 50.0). Updated to 50 days; reduces NWC drag and improves early-year FCFF.

**Warning #4 — Capex % adjusted to reflect declining trend (fixed)**  
FY26A actual capex % = 2.6%. Model had 3.5% for FY27E — a 90 bps reversal against a 4-year declining trend (4.3% → 2.6%). Updated schedule: 2.8% (FY27E) → 2.6% → 2.4% → 2.2% → 2.0%. ASSUMP note corrected from "FY24-26A: ~3.0–3.5%" to "FY26A actual: 2.6%".

---

## Python Scripts

### `build_model_v3.py` — Excel Model Builder
Programmatically generates `Dixon_3Statement_DCF_Model_v3.0.xlsx` via `openpyxl`. Writes all 8 tabs with cross-sheet Excel formulas, colour-coded by type.

```bash
pip install -r requirements.txt
python build_model_v3.py
# → outputs: Dixon_3Statement_DCF_Model_v3.0.xlsx
```

### `analysis_v3.py` — Financial Analysis Dashboard
Runs the full FCFF DCF in Python, builds a 2-way sensitivity table, and outputs a 6-panel `matplotlib` dashboard as a PNG.

```bash
python analysis_v3.py
# → outputs: Dixon_Analysis_Dashboard_v3.png
```

---

## Model Integration Architecture

The three statements are connected as follows — this is the correct flow that v2.0 violated:

```
Income Statement
  Revenue − (Material + Employee + Other) = Total Expenses
  Revenue − Total Expenses  =  EBITDA  (margin-driven from ASSUMP)
  EBITDA − D&A  →  EBIT
  EBIT + OtherInc − Interest[t-1 Debt]  →  PBT → PAT
  PAT − Dividends  →  Retained Earnings
         │
         ▼
Balance Sheet
  Reserves[t] = Reserves[t-1] + RetainedEarnings[t]
  PP&E[t] = PP&E[t-1] + Capex − D&A
  Receivables = Revenue × DebtorDays / 365      ← days-based NWC
  Inventory   = Revenue × InventoryDays / 365
  Payables    = COGS × DaysPayable / 365
  NWC[t]      = (Rec + Inv + OCA) − (Pay + OCL)
  Cash[t]     = Cash[t-1] + CF.NetChange[t]     ← from CF
         │
         ▼
Cash Flow Statement
  PAT + D&A + Interest  →  gross CFO
  ΔNWC[t] = −(BS.NWC[t] − BS.NWC[t-1])         ← TRUE integration
  CFO = PAT + D&A + ΔNWC + other
  CFI = −Capex
  CFF = ΔDebt − Dividends − Interest
  NetCF → feeds BS Cash (closing = BS.Cash)      ← integration check
```

The balance check row `= Total Assets − Total L&E` equals **zero** for every period when these formulas are implemented correctly.

---

## Workbook — 8 Tabs

| # | Tab | Contents |
|---|-----|----------|
| 1 | **Cover** | Company overview · historical snapshot FY21–FY26 · colour legend |
| 2 | **ASSUMP** | All editable inputs: revenue growth · EBITDA% · D&A% · tax · Kd · capex% · WC days · shares · Rf · ERP · β · WACC · TGR |
| 3 | **IS** | Income Statement FY21A–FY31E · Material/Employee/Other → Total Expenses → EBITDA → EBIT → PAT · NOPAT (DCF input) |
| 4 | **BS** | Balance Sheet · days-based NWC · reserves roll · cash from CF · balance check = 0 |
| 5 | **CF** | Indirect Cash Flow · ΔNWC from BS changes (not separate formula) · closing cash = BS cash |
| 6 | **DCF** | WACC derivation → FCFF → PV → Terminal Value → EV bridge → Implied Price |
| 7 | **SENSITIVITY** | 2-way table: Implied Price vs WACC (11–16%) × TGR (4–7%) |
| 8 | **COMPS** | Dixon vs Amber · Kaynes · Syrma SGS · Avalon · PG Electroplast |

**Colour coding:** Blue text = hardcoded input · Black = formula · Green = cross-sheet link · Grey fill = subtotal/total

---

## Valuation Summary

### WACC (CAPM)

| Parameter | Value | Source |
|-----------|-------|--------|
| Risk-Free Rate (Rf) | 7.00% | India 10-yr G-Sec — RBI (June 2026) |
| Equity Risk Premium (ERP) | 6.00% | Damodaran India ERP (Jan 2026) |
| Beta (β, levered) | 1.40 | 2-yr weekly regression vs Nifty 50 |
| **Cost of Equity (Ke)** | **15.4%** | 7.0% + 1.40 × 6.0% |
| Pre-tax Cost of Debt (Kd) | 8.50% | Repo rate + ~200bps |
| After-tax Kd | 6.63% | 8.5% × (1 − 22%) |
| Debt Weight | 20% | FY26 BS structure |
| **WACC** | **13.65%** | Ke × 80% + Kd(at) × 20% |
| Terminal Growth Rate | 5.50% | Conservative India nominal GDP proxy |

### DCF Output

```
FCFF = NOPAT + D&A − Capex − ΔNWC
     where NOPAT = EBIT × (1 − 22%)
           ΔNWC  = change in BS Net Working Capital (days-based)

EV  = Σ PV(FCFF₂₇₋₃₁) + PV(Terminal Value)
    = ₹2,176 Cr  +  ₹12,741 Cr  =  ₹14,917 Cr

Equity Value = EV − Debt + Cash − MI
             = ₹14,917 − ₹994 + ₹315 − ₹150  =  ₹14,088 Cr

Implied Price = ₹14,088 Cr × 100 / 676.4 lakh shares  =  ₹2,083 / share
```

---

## Historical Data (FY21A–FY26A) — Screener.in

| Metric | FY21A | FY22A | FY23A | FY24A | FY25A | FY26A |
|--------|-------|-------|-------|-------|-------|-------|
| Revenue (₹ Cr) | 6,448 | 10,697 | 12,192 | 17,691 | 38,860 | 48,873 |
| YoY Growth | — | +66% | +14% | +45% | +120% | +26% |
| Material Cost (₹ Cr) | 5,479 | 9,385 | 10,506 | 15,457 | 34,357 | 43,716 |
| Employee Cost (₹ Cr) | 123 | 206 | 234 | 340 | 373 | 470 |
| Other Opex (₹ Cr) | 554 | 722 | 933 | 1,189 | 2,615 | 2,820 |
| Total Expenses (₹ Cr) | 6,156 | 10,313 | 11,673 | 16,986 | 37,345 | 47,006 |
| EBITDA (₹ Cr) | 292 | 384 | 519 | 705 | 1,515 | 1,867 |
| EBITDA% | 4.5% | 3.6% | 4.3% | 4.0% | 3.9% | 3.8% |
| PAT (₹ Cr) | 160 | 190 | 255 | 375 | 1,233 | 1,644 |
| EPS (₹) | 27.28 | 32.05 | 42.90 | 61.47 | 181.87 | 236.61 |

> Note: FY25 PAT elevated by PLI-related Other Income (₹497 Cr). EBITDA shown is true operating profit (excl. Other Income), consistent with Screener.in.

---

## WiDS 5.0 Curriculum Coverage

| Week | Topic | Deliverable |
|------|-------|-------------|
| W1 | 3-Statement Model + Ratio Analysis | IS · BS · CF tabs |
| W2 | Revenue Forecasting + Drivers | ASSUMP tab + IS forecast |
| W3 | WACC / CAPM + DCF | DCF tab |
| W4 | Peer Comparables | COMPS tab |
| W5 | Python Automation | `build_model_v3.py` + `analysis_v3.py` |

---

## Data Source

| Source | Used For |
|--------|----------|
| [Screener.in — DIXON (Consolidated)](https://www.screener.in/company/DIXON/consolidated/) | **All** historical financials: P&L, BS, CF, WC days (FY21–FY26) |

All forecast assumptions are derived from the Screener historical data. No external broker estimates or alternative data sources are used.

---

> **Disclaimer:** Academic learning exercise. Not investment advice. All data from Screener.in public filings.  
> *Questions: omchaudhari19125@gmail.com*
