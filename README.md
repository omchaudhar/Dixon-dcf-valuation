# Dixon Technologies — 3-Statement Financial Model & DCF Valuation

> **Bottom-up equity research project** · WiDS 5.0 Financial Modelling Bootcamp · Summer 2027 finance internship portfolio  
> **Analyst:** Om Chaudhari | IIT Bombay, B.Tech (Year 2) | June 2026

---

## Quick Summary

| | |
|---|---|
| **Company** | Dixon Technologies (India) Ltd. — NSE: DIXON · BSE: 541988 |
| **Model** | Fully integrated 3-Statement (IS + BS + CF) + FCFF DCF |
| **Historical** | FY22A – FY25A · Consolidated · Source: Annual Reports (BSE/NSE) |
| **Forecast** | FY26E – FY30E (5-year explicit) |
| **WACC** | 13.6% · Ke = 15.4% · Kd(at) = 6.4% · D-weight = 20% |
| **Implied Price** | ₹2,346/share · TGR = 5.5% · Gordon Growth terminal value |
| **Tabs** | 8 — Cover · ASSUMP · IS · BS · CF · DCF · SENSITIVITY · COMPS |
| **Currency** | ₹ Crore · Indian FY (Apr 1 – Mar 31) |

---

## Files in This Repo

```
Dixon-dcf-valuation/
│
├── Dixon_3Statement_DCF_Model_v2.0.xlsx   ← Main deliverable — open in Excel
├── Dixon_Analysis_Dashboard.png           ← 6-panel Python dashboard (analysis.py output)
├── build_model_v2.py                      ← Generates the Excel file (WiDS Week 5)
├── analysis.py                            ← Python DCF + charts (WiDS Week 5)
├── requirements.txt                       ← Python dependencies
└── README.md
```

---

## Python Scripts

### `build_model_v2.py` — Excel Model Builder
Programmatically generates the entire `Dixon_3Statement_DCF_Model_v2.0.xlsx` workbook using `openpyxl`. Writes all 8 tabs, applies colour-coding (blue = hardcoded inputs, black = formulas, green = cross-sheet links), and outputs 634 live Excel formulas. This is the **WiDS Week 5 Python automation deliverable**.

```bash
pip install -r requirements.txt
python build_model_v2.py
# → outputs: Dixon_3Statement_DCF_Model_v2.0.xlsx
```

### `analysis.py` — Financial Analysis Dashboard
Loads Dixon financials into **pandas DataFrames**, computes WACC via CAPM, runs the full FCFF DCF in Python, builds a 2-way sensitivity table, and outputs a 6-panel **matplotlib dashboard** as a PNG.

```bash
python analysis.py
# → outputs: Dixon_Analysis_Dashboard.png
```

---

## Workbook — 8 Tabs

| # | Tab | Contents |
|---|-----|----------|
| 1 | **Cover** | Company overview · key metrics FY22–FY25 · colour legend · data sources |
| 2 | **ASSUMP** | All editable inputs: revenue growth · EBITDA% · D&A% · tax · Kd · capex% · NWC% · shares · Rf · ERP · β · WACC · TGR |
| 3 | **IS** | Income Statement FY22A–FY30E + full Ratio Analysis (profitability · efficiency · leverage) |
| 4 | **BS** | Balance Sheet · formula-linked to IS and CF |
| 5 | **CF** | Cash Flow Statement (indirect) · CFO / CFI / CFF · closing cash feeds into BS |
| 6 | **DCF** | WACC derivation → FCFF model → PV of CFs → Terminal Value → EV bridge → Implied Price |
| 7 | **SENSITIVITY** | 2-way table: Implied Price vs WACC (11–16%) × TGR (4–7%) |
| 8 | **COMPS** | Dixon vs Amber Enterprises · Kaynes Technology · Syrma SGS — EV/EBITDA · P/E · EBITDA% |

**Colour coding:** Blue text = hardcoded input · Black = formula · Green = cross-sheet link · Yellow fill = assumption to verify

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
| After-tax Kd | 6.38% | 8.5% × (1 − 25%) |
| Debt Weight | 20% | Dixon equity-heavy (FY25 BS) |
| **WACC** | **13.6%** | Ke × 80% + Kd(at) × 20% |
| Terminal Growth Rate | 5.50% | Conservative India nominal GDP proxy |

### DCF Output

```
FCFF = NOPAT + D&A − Capex − ΔNWC
     where NOPAT = EBIT × (1 − Tax Rate)

EV  = Σ PV(FCFF) + PV(Terminal Value)
    = ₹3,378 Cr  + ₹10,696 Cr  =  ₹14,073 Cr

Equity Value = EV − Debt + Cash − MI
             = ₹14,073 − ₹1,100 + ₹1,200 − ₹100  =  ₹14,073 Cr

Implied Price = ₹14,073 Cr × 100 / 600 lakh shares  =  ₹2,346/share
```

---

## Historical Data (FY22A–FY25A)

| Metric | FY22A | FY23A | FY24A | FY25A |
|--------|-------|-------|-------|-------|
| Revenue (₹ Cr) | 10,697 | 12,192 | 17,691 | 38,860 |
| YoY Growth | — | +14% | +45% | +120% |
| EBITDA% | 4.0% | 4.2% | 4.0% | 5.0% |
| PAT (₹ Cr) | 190 | 253 | 365 | 1,215 |
| EPS (₹) | 31.7 | 42.2 | 60.9 | 202.5 |

---

## WiDS 5.0 Curriculum Coverage

| Week | Topic | Deliverable |
|------|-------|-------------|
| W1 | 3-Statement Model + Ratio Analysis | IS · BS · CF tabs |
| W2 | Revenue Forecasting + Drivers | ASSUMP tab + IS forecast |
| W3 | WACC / CAPM + DCF | DCF tab |
| W4 | Peer Comparables | COMPS tab |
| W5 | Python Automation | `build_model_v2.py` + `analysis.py` |

---

## Data Sources

| Source | Used For |
|--------|----------|
| Dixon Annual Reports FY22–FY25 (BSE/NSE) | Revenue, EBITDA, PAT, BS, CF |
| [Screener.in — DIXON](https://www.screener.in/company/DIXON/consolidated/) | Consolidated financials cross-check |
| [Damodaran India ERP](https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html) | ERP = 6.0% |
| [RBI](https://www.rbi.org.in/) | 10-yr G-Sec yield (Rf = 7.0%) |
| Motilal Oswal / ICICI Direct | Broker estimates for FY26–FY30 context |

---

> **Disclaimer:** Academic learning exercise. Not investment advice. All data from public filings.  
> *Questions: omchaudhari19125@gmail.com*
