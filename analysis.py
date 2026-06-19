"""
Dixon Technologies (India) Ltd. — Python Financial Analysis
============================================================
WiDS 5.0 Financial Modelling Bootcamp | Week 5 — Python Automation

This script demonstrates how Python supports financial analysis:
  1. Load & structure financials into pandas DataFrames
  2. Compute key ratios (profitability, efficiency, leverage)
  3. Run FCFF-based DCF valuation in Python code
  4. Build a 2-way sensitivity table (WACC × TGR) as a DataFrame
  5. Produce a 6-panel matplotlib dashboard → saved as PNG
  6. Print a clean summary report to console

Run:
    pip install pandas matplotlib seaborn numpy openpyxl
    python analysis.py

Outputs:
    Dixon_Analysis_Dashboard.png  — 6-chart visual summary
    (console)                     — WACC derivation + DCF summary + ratios
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
from matplotlib.gridspec import GridSpec

# ══════════════════════════════════════════════════════════════════════════════
# 0. STYLE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
plt.rcParams.update({
    'font.family':   'DejaVu Sans',
    'font.size':     9,
    'axes.spines.top':   False,
    'axes.spines.right': False,
    'axes.grid':     True,
    'grid.alpha':    0.3,
    'figure.facecolor': '#FAFAFA',
    'axes.facecolor':   '#FAFAFA',
})

NAVY   = '#003566'
TEAL   = '#0077B6'
SKY    = '#48CAE4'
AMBER  = '#F4A261'
GREEN  = '#2DC653'
RED    = '#E63946'
GREY   = '#ADB5BD'
HIST_C = TEAL
FCST_C = AMBER

# ══════════════════════════════════════════════════════════════════════════════
# 1. FINANCIAL DATA  (₹ Crore; Indian FY Apr-Mar)
# ══════════════════════════════════════════════════════════════════════════════

YEARS = ['FY22A', 'FY23A', 'FY24A', 'FY25A', 'FY26E', 'FY27E', 'FY28E', 'FY29E', 'FY30E']
HIST  = ['FY22A', 'FY23A', 'FY24A', 'FY25A']
FCST  = ['FY26E', 'FY27E', 'FY28E', 'FY29E', 'FY30E']

# ── ASSUMPTIONS (mirrors ASSUMP tab in Excel model) ──────────────────────────
REV_GROWTH  = [None, 0.14, 0.45, 1.20, 0.30, 0.25, 0.18, 0.15, 0.12]
EBITDA_PCT  = 0.052    # 5.2% consolidated margin (PLI + scale)
DA_PCT      = 0.008    # 0.8% of revenue
TAX_RATE    = 0.25     # 25% concessional rate
INT_PCT     = 0.003    # ~0.3% of revenue (working capital finance)
CAPEX_PCT   = 0.020    # 2.0% of revenue (asset-light EMS model)
NWC_PCT     = 0.045    # 4.5% of revenue (working capital intensity)
SHARES_LAKH = 600      # 600 lakh = 6 crore shares (NSE: DIXON)

# WACC INPUTS
RF          = 0.070    # Risk-free rate: India 10-yr G-Sec (RBI, June 2026)
ERP         = 0.060    # Equity Risk Premium: Damodaran India (Jan 2026)
BETA        = 1.40     # Levered beta: 2-yr weekly regression vs Nifty 50
KD_PRETAX   = 0.085    # Pre-tax cost of debt (AR FY25 notes to accounts)
D_WEIGHT    = 0.20     # Debt weight in capital structure
TGR         = 0.055    # Terminal Growth Rate (below India nominal GDP ~8%)

# EV BRIDGE INPUTS (FY25A balance sheet, ₹ Crore)
TOTAL_DEBT  = 1100     # LT + ST borrowings
CASH        = 1200     # Cash & equivalents
MIN_INT     = 100      # Minority interest (NCI)

# ── HISTORICAL P&L (Source: Dixon AR FY25, BSE filing) ───────────────────────
hist_revenue  = [10697, 12192, 17691, 38860]
hist_ebitda   = [430,   510,   708,   1943]
hist_da       = [115,   140,   185,   280]
hist_interest = [35,    45,    85,    120]
hist_pat      = [190,   253,   365,   1215]
hist_capex    = [155,   195,   340,   600]   # from CF statement

# ── FORECAST P&L (derived from ASSUMP) ───────────────────────────────────────
def build_forecasts(base_rev, growth_rates):
    rev = []
    r   = base_rev
    for g in growth_rates:
        r = round(r * (1 + g))
        rev.append(r)
    return rev

fcst_rev      = build_forecasts(38860, [0.30, 0.25, 0.18, 0.15, 0.12])
fcst_ebitda   = [round(r * EBITDA_PCT)  for r in fcst_rev]
fcst_da       = [round(r * DA_PCT)      for r in fcst_rev]
fcst_ebit     = [eb - d for eb, d in zip(fcst_ebitda, fcst_da)]
fcst_interest = [round(r * INT_PCT)     for r in fcst_rev]
fcst_ebt      = [e - i for e, i in zip(fcst_ebit, fcst_interest)]
fcst_tax      = [round(ebt * TAX_RATE)  for ebt in fcst_ebt]
fcst_pat      = [ebt - t for ebt, t in zip(fcst_ebt, fcst_tax)]
fcst_capex    = [round(r * CAPEX_PCT)   for r in fcst_rev]

# ── BUILD COMBINED DATAFRAME ──────────────────────────────────────────────────
revenue  = hist_revenue  + fcst_rev
ebitda   = hist_ebitda   + fcst_ebitda
da       = hist_da       + fcst_da
ebit     = [eb - d  for eb, d  in zip(ebitda, da)]
interest = hist_interest + fcst_interest
pat      = hist_pat      + fcst_pat
capex    = hist_capex    + fcst_capex

df = pd.DataFrame({
    'Revenue':   revenue,
    'EBITDA':    ebitda,
    'DA':        da,
    'EBIT':      ebit,
    'Interest':  interest,
    'PAT':       pat,
    'Capex':     capex,
}, index=YEARS)

# ── COMPUTED RATIOS ───────────────────────────────────────────────────────────
df['EBITDA_pct'] = (df['EBITDA'] / df['Revenue'] * 100).round(1)
df['EBIT_pct']   = (df['EBIT']   / df['Revenue'] * 100).round(1)
df['PAT_pct']    = (df['PAT']    / df['Revenue'] * 100).round(1)
df['YoY_Growth'] = df['Revenue'].pct_change() * 100
df['EPS']        = (df['PAT'] * 100 / SHARES_LAKH).round(1)

# ══════════════════════════════════════════════════════════════════════════════
# 2. WACC DERIVATION
# ══════════════════════════════════════════════════════════════════════════════
KE       = RF + BETA * ERP
KD_AT    = KD_PRETAX * (1 - TAX_RATE)
E_WEIGHT = 1 - D_WEIGHT
WACC     = KE * E_WEIGHT + KD_AT * D_WEIGHT

# ══════════════════════════════════════════════════════════════════════════════
# 3. FCFF DCF CALCULATION
# ══════════════════════════════════════════════════════════════════════════════
fcst_df = df.loc[FCST].copy()
prev_rev = [38860] + fcst_rev[:-1]

nopat     = fcst_df['EBIT'] * (1 - TAX_RATE)
delta_nwc = pd.Series(
    [round((r - p) * NWC_PCT) for r, p in zip(fcst_rev, prev_rev)],
    index=FCST
)
fcff = nopat + fcst_df['DA'] - fcst_df['Capex'] - delta_nwc

pv_fcff = pd.Series(
    [f / (1 + WACC)**t for f, t in zip(fcff, range(1, 6))],
    index=FCST
)

# Terminal Value (Gordon Growth Model)
tv       = fcff.iloc[-1] * (1 + TGR) / (WACC - TGR)
pv_tv    = tv / (1 + WACC)**5

sum_pv_fcff = pv_fcff.sum()
ev          = sum_pv_fcff + pv_tv
equity_val  = ev - TOTAL_DEBT + CASH - MIN_INT
implied_px  = equity_val * 100 / SHARES_LAKH

# ══════════════════════════════════════════════════════════════════════════════
# 4. SENSITIVITY TABLE (WACC × TGR)
# ══════════════════════════════════════════════════════════════════════════════
wacc_range = [0.110, 0.120, 0.130, 0.136, 0.140, 0.150, 0.160]
tgr_range  = [0.040, 0.045, 0.050, 0.055, 0.060, 0.065, 0.070]

def calc_price(w, g):
    if g >= w:
        return np.nan
    pv_f = sum(f / (1+w)**t for f, t in zip(fcff, range(1, 6)))
    tv_s = fcff.iloc[-1] * (1+g) / (w - g)
    eq   = pv_f + tv_s/(1+w)**5 - TOTAL_DEBT + CASH - MIN_INT
    return round(eq * 100 / SHARES_LAKH)

sens_data = {
    f'{g:.1%}': {f'{w:.1%}': calc_price(w, g) for w in wacc_range}
    for g in tgr_range
}
sens_df = pd.DataFrame(sens_data)
sens_df.index.name   = 'WACC'
sens_df.columns.name = 'TGR →'

# ══════════════════════════════════════════════════════════════════════════════
# 5. CONSOLE REPORT
# ══════════════════════════════════════════════════════════════════════════════
SEP = '═' * 60

print(f'\n{SEP}')
print('  DIXON TECHNOLOGIES — PYTHON FINANCIAL ANALYSIS')
print('  NSE: DIXON | WiDS 5.0 Week 5 | June 2026')
print(SEP)

print(f'\n── WACC DERIVATION ──────────────────────────────────────')
print(f'  Rf  (India 10-yr G-Sec) :  {RF:.1%}')
print(f'  ERP (Damodaran India)   :  {ERP:.1%}')
print(f'  Beta (levered, 2yr wkly):  {BETA:.2f}')
print(f'  Ke  = {RF:.0%} + {BETA}×{ERP:.0%}   :  {KE:.1%}')
print(f'  Kd  (pre-tax)           :  {KD_PRETAX:.1%}')
print(f'  Kd  (after-tax)         :  {KD_AT:.3%}')
print(f'  D-weight / E-weight     :  {D_WEIGHT:.0%} / {E_WEIGHT:.0%}')
print(f'  ► WACC                  :  {WACC:.2%}')

print(f'\n── FREE CASH FLOW TO FIRM (₹ Cr) ────────────────────────')
print(f'  {"Year":<8} {"EBIT":>7} {"NOPAT":>7} {"D&A":>6} {"Capex":>7} {"ΔNWC":>7} {"FCFF":>7} {"PV(FCFF)":>10}')
for yr in FCST:
    print(f'  {yr:<8} '
          f'{fcst_df.loc[yr,"EBIT"]:>7,.0f} '
          f'{nopat[yr]:>7,.0f} '
          f'{fcst_df.loc[yr,"DA"]:>6,.0f} '
          f'{fcst_df.loc[yr,"Capex"]:>7,.0f} '
          f'{delta_nwc[yr]:>7,.0f} '
          f'{fcff[yr]:>7,.0f} '
          f'{pv_fcff[yr]:>10,.0f}')

print(f'\n── DCF VALUATION BRIDGE ─────────────────────────────────')
print(f'  Sum of PV(FCFF)     :  ₹ {sum_pv_fcff:>10,.0f} Cr')
print(f'  Terminal Value (TV) :  ₹ {tv:>10,.0f} Cr')
print(f'  PV of TV            :  ₹ {pv_tv:>10,.0f} Cr  '
      f'({pv_tv/ev:.0%} of EV — "terminal value dominance")')
print(f'  Enterprise Value    :  ₹ {ev:>10,.0f} Cr')
print(f'  Less: Total Debt    : (₹ {TOTAL_DEBT:>9,.0f} Cr)')
print(f'  Add:  Cash          :  ₹ {CASH:>10,.0f} Cr')
print(f'  Less: Min. Interest : (₹ {MIN_INT:>9,.0f} Cr)')
print(f'  Equity Value        :  ₹ {equity_val:>10,.0f} Cr')
print(f'  Shares (lakh)       :  {SHARES_LAKH}  (= {SHARES_LAKH/100:.1f} crore)')
print(f'  ► Implied Price     :  ₹ {implied_px:>10,.0f} per share')

print(f'\n── KEY RATIOS ───────────────────────────────────────────')
ratio_display = df[['Revenue','EBITDA_pct','PAT_pct','EPS']].copy()
ratio_display.columns = ['Revenue(₹Cr)', 'EBITDA%', 'PAT%', 'EPS(₹)']
print(ratio_display.to_string(float_format='{:,.1f}'.format))

print(f'\n── SENSITIVITY: Implied Share Price ₹ (WACC × TGR) ─────')
print(sens_df.to_string(float_format='{:,.0f}'.format, na_rep='N/A'))
print(f'\n  ★ Base (WACC={WACC:.1%}, TGR={TGR:.1%}) → ₹{implied_px:,.0f}')
print()

# ══════════════════════════════════════════════════════════════════════════════
# 6. DASHBOARD — 6 CHARTS
# ══════════════════════════════════════════════════════════════════════════════
fig = plt.figure(figsize=(16, 10), facecolor='#F8F9FA')
fig.suptitle('Dixon Technologies (India) Ltd. — Financial Analysis Dashboard\n'
             'NSE: DIXON | WiDS 5.0 Week 5 | ₹ Crore unless stated',
             fontsize=13, fontweight='bold', color=NAVY, y=0.98)

gs = GridSpec(2, 3, figure=fig, hspace=0.45, wspace=0.35,
              left=0.07, right=0.97, top=0.92, bottom=0.08)

colors = [HIST_C if y in HIST else FCST_C for y in YEARS]

# ── CHART 1: Revenue + Growth ──────────────────────────────────────────────
ax1 = fig.add_subplot(gs[0, 0])
bars = ax1.bar(YEARS, revenue, color=colors, width=0.6, zorder=2)
ax1.set_title('Revenue (₹ Cr)', fontweight='bold', color=NAVY, pad=6)
ax1.set_ylabel('₹ Crore')
ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x/1000:.0f}K'))
ax2 = ax1.twinx()
ax2.plot(YEARS, df['YoY_Growth'].fillna(0),
         color=RED, marker='o', ms=4, lw=1.5, label='YoY%')
ax2.yaxis.set_major_formatter(mticker.PercentFormatter())
ax2.set_ylabel('YoY Growth %', color=RED, fontsize=8)
ax2.tick_params(axis='y', labelcolor=RED, labelsize=8)
ax2.spines['right'].set_visible(True)
ax2.spines['top'].set_visible(False)
ax1.set_xticklabels(YEARS, rotation=45, ha='right', fontsize=7)
# Add value labels on bars
for bar, val in zip(bars, revenue):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 600,
             f'{val/1000:.0f}K', ha='center', va='bottom', fontsize=6.5, color=NAVY)
ax1.axvline(x=3.5, color=GREY, lw=1, ls='--', alpha=0.7)
ax1.text(3.7, max(revenue)*0.85, 'Forecast →', fontsize=7, color=GREY, style='italic')

# ── CHART 2: EBITDA & PAT Margins ─────────────────────────────────────────
ax3 = fig.add_subplot(gs[0, 1])
x   = np.arange(len(YEARS))
w   = 0.35
ax3.bar(x - w/2, df['EBITDA_pct'], w, color=TEAL,  label='EBITDA%', zorder=2)
ax3.bar(x + w/2, df['PAT_pct'],    w, color=AMBER, label='PAT%',    zorder=2)
ax3.set_title('Margin Profile (%)', fontweight='bold', color=NAVY, pad=6)
ax3.set_ylabel('Margin %')
ax3.yaxis.set_major_formatter(mticker.PercentFormatter())
ax3.set_xticks(x); ax3.set_xticklabels(YEARS, rotation=45, ha='right', fontsize=7)
ax3.legend(fontsize=8, loc='upper left')
ax3.axvline(x=3.5, color=GREY, lw=1, ls='--', alpha=0.7)
# Convert % values for display
for xi, (ep, pp) in enumerate(zip(df['EBITDA_pct'], df['PAT_pct'])):
    ax3.text(xi - w/2, ep + 0.05, f'{ep:.1f}%', ha='center', va='bottom', fontsize=5.5, color=TEAL)
    ax3.text(xi + w/2, pp + 0.05, f'{pp:.1f}%', ha='center', va='bottom', fontsize=5.5, color='#8B4513')

# ── CHART 3: FCFF Waterfall (forecast only) ───────────────────────────────
ax4 = fig.add_subplot(gs[0, 2])
fcff_vals = fcff.values
bar_colors = [GREEN if v > 0 else RED for v in fcff_vals]
ax4.bar(FCST, fcff_vals, color=bar_colors, width=0.5, zorder=2)
ax4.axhline(0, color=NAVY, lw=0.8)
ax4.set_title('FCFF — Free Cash Flow to Firm (₹ Cr)', fontweight='bold', color=NAVY, pad=6)
ax4.set_ylabel('₹ Crore')
ax4.set_xticklabels(FCST, rotation=45, ha='right', fontsize=7)
for yr, val in zip(FCST, fcff_vals):
    va  = 'bottom' if val > 0 else 'top'
    off = 20 if val > 0 else -20
    ax4.text(yr, val + off, f'₹{val:,.0f}', ha='center', va=va, fontsize=7, fontweight='bold')
ax4.text(0.02, 0.02, 'FCFF = NOPAT + D&A − Capex − ΔNWC',
         transform=ax4.transAxes, fontsize=6.5, color=GREY, style='italic')

# ── CHART 4: DCF Value Bridge (waterfall) ─────────────────────────────────
ax5 = fig.add_subplot(gs[1, 0])
labels  = ['ΣPV(FCFF)', 'PV(TV)', 'EV', '−Debt', '+Cash', '−MI', 'Equity']
vals    = [sum_pv_fcff, pv_tv, 0, -TOTAL_DEBT, CASH, -MIN_INT, 0]
running = 0
bottoms = []
heights = []
for i, (l, v) in enumerate(zip(labels, vals)):
    if l in ['EV', 'Equity']:
        bottoms.append(0)
        running_sum = sum_pv_fcff + pv_tv if l == 'EV' else equity_val
        heights.append(running_sum)
        running = running_sum
    else:
        bottoms.append(min(running, running + v))
        heights.append(abs(v))
        running += v

bar_clrs = [TEAL, SKY, NAVY, RED, GREEN, RED, AMBER]
ax5.bar(labels, heights, bottom=bottoms, color=bar_clrs, width=0.55, zorder=2)
ax5.set_title('EV → Equity Bridge (₹ Cr)', fontweight='bold', color=NAVY, pad=6)
ax5.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x/1000:.0f}K'))
ax5.set_ylabel('₹ Crore')
ax5.set_xticklabels(labels, rotation=30, ha='right', fontsize=7)

# ── CHART 5: PV Breakdown pie ──────────────────────────────────────────────
ax6 = fig.add_subplot(gs[1, 1])
pv_labels   = [f'PV yr{i+1}\n₹{v:,.0f}' for i, v in enumerate(pv_fcff)]
pv_labels  += [f'PV(TV)\n₹{pv_tv:,.0f}']
pv_vals     = list(pv_fcff.values) + [pv_tv]
pie_colors  = [SKY, TEAL, '#0096C7', '#0077B6', '#023E8A', NAVY]
wedges, texts, autotexts = ax6.pie(
    pv_vals, labels=pv_labels, colors=pie_colors,
    autopct='%1.0f%%', startangle=90, pctdistance=0.75,
    textprops={'fontsize': 6.5}
)
for at in autotexts:
    at.set_fontsize(6.5); at.set_color('white'); at.set_fontweight('bold')
ax6.set_title(f'EV Composition\n(TV = {pv_tv/ev:.0%} of EV)',
              fontweight='bold', color=NAVY, pad=6)

# ── CHART 6: Sensitivity Heatmap ──────────────────────────────────────────
ax7 = fig.add_subplot(gs[1, 2])
hmap_data = sens_df.T.astype(float)
sns.heatmap(
    hmap_data, ax=ax7, annot=True, fmt='.0f', cmap='RdYlGn',
    linewidths=0.4, linecolor='#E0E0E0',
    annot_kws={'size': 7, 'weight': 'bold'},
    cbar_kws={'label': '₹ / share', 'shrink': 0.8}
)
ax7.set_title(f'Sensitivity: Implied Price ₹\n(WACC rows × TGR cols)',
              fontweight='bold', color=NAVY, pad=6)
ax7.set_xlabel('Terminal Growth Rate', fontsize=8)
ax7.set_ylabel('WACC', fontsize=8)
ax7.tick_params(axis='both', labelsize=7)

# Highlight base case cell
wacc_idx = [f'{w:.1%}' for w in wacc_range].index(f'{WACC:.1%}')
tgr_idx  = [f'{g:.1%}' for g in tgr_range].index(f'{TGR:.1%}')
ax7.add_patch(plt.Rectangle((tgr_idx, wacc_idx), 1, 1,
              fill=False, edgecolor='black', lw=2.5))

# ── FOOTER ────────────────────────────────────────────────────────────────
fig.text(0.5, 0.01,
         f'Base Case: WACC {WACC:.1%} | TGR {TGR:.1%} | '
         f'Implied Price ₹{implied_px:,.0f} | '
         f'Rf {RF:.0%} | ERP {ERP:.0%} | β {BETA} | '
         'Source: Dixon AR FY22-25, Damodaran India ERP, RBI G-Sec | '
         'This is not investment advice.',
         ha='center', fontsize=6.5, color=GREY, style='italic')

# ── SAVE ──────────────────────────────────────────────────────────────────
outfile = 'Dixon_Analysis_Dashboard.png'
plt.savefig(outfile, dpi=150, bbox_inches='tight', facecolor='#F8F9FA')
print(f'✅  Dashboard saved → {outfile}')
plt.show()
