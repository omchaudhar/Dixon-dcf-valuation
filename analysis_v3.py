"""
analysis_v3.py
Dixon Technologies (India) Ltd. — Python Financial Analysis & DCF Dashboard
Version 3.0  |  Source: Screener.in  |  Consistent with build_model_v3.py

Generates: Dixon_Analysis_Dashboard_v3.png  (6-panel matplotlib dashboard)

Author: Om Chaudhari | IIT Bombay
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.patches as mpatches
from matplotlib.colors import LinearSegmentedColormap
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
# 1.  DATA  (from Screener.in — single source)
# ─────────────────────────────────────────────

HIST_YEARS = ["FY21", "FY22", "FY23", "FY24", "FY25", "FY26"]
FCST_YEARS = ["FY27E", "FY28E", "FY29E", "FY30E", "FY31E"]
ALL_YEARS  = HIST_YEARS + FCST_YEARS

# P&L (₹ Crore, Screener.in consolidated)
revenue_hist    = [6448, 10697, 12192, 17691, 38860, 48873]
ebitda_hist     = [ 292,   384,   519,   705,  1515,  1867]  # Operating Profit (excl Other Inc)
other_inc_hist  = [   1,     4,     4,    32,   497,   734]
da_hist         = [  44,    84,   115,   162,   281,   393]
interest_hist   = [  33,    49,    64,    81,   162,   137]
pat_hist        = [ 160,   190,   255,   375,  1233,  1644]

# Expense breakdown (from Screener cost-structure %, consistent with build_model_v3.py)
_total_exp_hist = [r - e for r, e in zip(revenue_hist, ebitda_hist)]
_mat_hist   = [5479, 9385, 10506, 15457, 34357, 43716]  # 89-93% of total expenses
_emp_hist   = [ 123,  206,   234,   340,   373,   470]  # 1-2% of total expenses
_other_hist = [t - m - e for t, m, e in
               zip(_total_exp_hist, _mat_hist, _emp_hist)]   # balancing
# As % of revenue (for stacked chart)
mat_pct_hist  = [m/r*100 for m, r in zip(_mat_hist,   revenue_hist)]
emp_pct_hist  = [e/r*100 for e, r in zip(_emp_hist,   revenue_hist)]
oth_pct_hist  = [o/r*100 for o, r in zip(_other_hist, revenue_hist)]
ebitda_pct_hist = [e/r*100 for e, r in zip(ebitda_hist, revenue_hist)]

# BS (₹ Crore, Screener.in)
debt_hist       = [ 295,  667,  453,  489,  671,  994]
cash_hist_approx = 180   # FY21 starting cash estimate
cf_net_hist     = [ -32,  113,   41,  -17,   30,  424]
# Build running cash
cash_hist = []
c = cash_hist_approx
for net in cf_net_hist:
    cash_hist.append(max(int(c), 10))
    c += net

cfo_hist = [ 170,  273,  726,  584, 1150, 1782]
cfi_hist = [-265, -464, -356, -531,-1093,-1251]
cff_hist = [  63,  304, -330,  -70,  -27, -108]

# WC Days
debtor_days_hist  = [50, 46, 51, 48, 65, 62]
inv_days_hist     = [40, 43, 32, 39, 41, 40]
pay_days_hist     = [90, 86, 81, 92,111,108]

# ─────────────────────────────────────────────
# 2.  FORECAST ASSUMPTIONS
# ─────────────────────────────────────────────

REV_GROWTH    = [0.22, 0.20, 0.18, 0.16, 0.14]
EBITDA_MARGIN = [0.041, 0.043, 0.045, 0.047, 0.048]
OTHER_INC_PCT = [0.008, 0.007, 0.006, 0.005, 0.005]
DA_PCT        = [0.0075, 0.0080, 0.0082, 0.0083, 0.0083]
TAX_RATE      = 0.22
CAPEX_PCT     = [0.028, 0.026, 0.024, 0.022, 0.020]   # FY26A 2.6%; 2.8% for FY27E then declining
DEBT_FCST     = [1100, 1200, 1300, 1350, 1400]
DEBTOR_FCST   = 50    # FY26A actual: 50 days (₹6,695Cr/₹48,873Cr×365)
INV_FCST      = 40
PAY_FCST      = 105
COGS_PCT      = 0.88
MINORITY_INT  = 50    # FY26A actual; no acquisition modelled (audit fix: was 150)
FCST_MAT_PCT  = 0.920  # material as % of total expenses in forecast (consistent with build_model_v3.py)
FCST_EMP_PCT  = 0.015  # employee as % of total expenses

# Forecast cost structure (% of revenue): total_exp = (1 - EBITDA_MARGIN) per year
mat_pct_fcst  = [(1-em)*FCST_MAT_PCT*100       for em in EBITDA_MARGIN]
emp_pct_fcst  = [(1-em)*FCST_EMP_PCT*100        for em in EBITDA_MARGIN]
oth_pct_fcst  = [(1-em)*(1-FCST_MAT_PCT-FCST_EMP_PCT)*100 for em in EBITDA_MARGIN]
ebitda_pct_fcst = [em*100                       for em in EBITDA_MARGIN]

# WACC
RF   = 0.070
ERP  = 0.060
BETA = 1.40
KE   = RF + BETA * ERP        # 15.4%
DEBT_WEIGHT = 0.20
KD_AT = 0.085 * (1 - TAX_RATE)  # 6.63%
WACC = KE * (1 - DEBT_WEIGHT) + KD_AT * DEBT_WEIGHT
TGR  = 0.055
SHARES_LAKH = 676.4  # lakh shares (FY26 diluted)

# ─────────────────────────────────────────────
# 3.  FORECAST BUILD
# ─────────────────────────────────────────────

rev_fcst, ebitda_fcst, da_fcst, ebit_fcst = [], [], [], []
interest_fcst, pat_fcst, nopat_fcst, capex_fcst = [], [], [], []
fcff_fcst, debt_fcst_vals = [], []
nwc_hist, nwc_fcst = [], []

# Compute historical NWC
for i, yr in enumerate(HIST_YEARS):
    rec = round(revenue_hist[i] * debtor_days_hist[i] / 365)
    inv = round(revenue_hist[i] * inv_days_hist[i] / 365)
    pay = round(revenue_hist[i] * COGS_PCT * pay_days_hist[i] / 365)
    nwc_hist.append(rec + inv - pay)

rev = revenue_hist[-1]
prev_debt = debt_hist[-1]
prev_nwc  = nwc_hist[-1]

for i in range(5):
    g = REV_GROWTH[i]
    rev = rev * (1 + g)
    rev_fcst.append(round(rev))

    ebitda = round(rev * EBITDA_MARGIN[i])
    ebitda_fcst.append(ebitda)

    da = round(rev * DA_PCT[i])
    da_fcst.append(da)

    ebit = ebitda - da
    ebit_fcst.append(ebit)

    # Interest on beginning-year debt (no circularity)
    interest = round(prev_debt * 0.085)
    interest_fcst.append(interest)

    pbt = ebit + round(rev * OTHER_INC_PCT[i]) - interest
    tax = round(pbt * TAX_RATE)
    pat = pbt - tax
    pat_fcst.append(pat)

    nopat = round(ebit * (1 - TAX_RATE))
    nopat_fcst.append(nopat)

    capex = round(rev * CAPEX_PCT[i])
    capex_fcst.append(capex)

    # NWC via days-based ratios
    rec_f = round(rev * DEBTOR_FCST / 365)
    inv_f = round(rev * INV_FCST / 365)
    pay_f = round(rev * COGS_PCT * PAY_FCST / 365)
    nwc_curr = rec_f + inv_f - pay_f
    nwc_fcst.append(nwc_curr)

    delta_nwc = -(nwc_curr - prev_nwc)   # negative = NWC increase = outflow
    fcff = nopat + da - capex + delta_nwc
    fcff_fcst.append(fcff)

    prev_debt = DEBT_FCST[i]
    prev_nwc  = nwc_curr
    debt_fcst_vals.append(DEBT_FCST[i])

# ─────────────────────────────────────────────
# 4.  DCF COMPUTATION
# ─────────────────────────────────────────────

pv_fcffs = [fcff / (1 + WACC)**(0.5 + i) for i, fcff in enumerate(fcff_fcst)]
sum_pv_fcff = sum(pv_fcffs)

fcff_terminal = fcff_fcst[-1] * (1 + TGR)
tv = fcff_terminal / (WACC - TGR)
pv_tv = tv / (1 + WACC)**5.0

ev = sum_pv_fcff + pv_tv

debt_fy26 = debt_hist[-1]
cash_fy26 = cash_hist[-1]

equity_val = ev - debt_fy26 + cash_fy26 - MINORITY_INT
implied_price = equity_val * 100 / SHARES_LAKH

print("─" * 55)
print("DIXON TECHNOLOGIES — DCF VALUATION SUMMARY  (v3.0)")
print("─" * 55)
print(f"  WACC:                  {WACC*100:.2f}%")
print(f"  TGR:                   {TGR*100:.1f}%")
print(f"  Sum PV(FCFF):          ₹{sum_pv_fcff:,.0f} Cr")
print(f"  PV(Terminal Value):    ₹{pv_tv:,.0f} Cr")
print(f"  Enterprise Value:      ₹{ev:,.0f} Cr")
print(f"  Less: Debt (FY26A):    ₹{debt_fy26:,.0f} Cr")
print(f"  Add: Cash (FY26A):     ₹{cash_fy26:,.0f} Cr")
print(f"  Less: Minority Int.:   ₹{MINORITY_INT:,} Cr")
print(f"  Equity Value:          ₹{equity_val:,.0f} Cr")
print(f"  Shares:                {SHARES_LAKH:.1f} Lakh")
print(f"  ──────────────────────────────────")
print(f"  IMPLIED PRICE:         ₹{implied_price:,.0f} / share")
print("─" * 55)

# ─────────────────────────────────────────────
# 5.  SENSITIVITY TABLE
# ─────────────────────────────────────────────

wacc_range = np.arange(0.110, 0.165, 0.005)
tgr_range  = np.arange(0.040, 0.075, 0.005)
sens_matrix = np.zeros((len(wacc_range), len(tgr_range)))

for ri, wc in enumerate(wacc_range):
    for ci, tg in enumerate(tgr_range):
        pv_sum = sum(fcff / (1 + wc)**(0.5 + i) for i, fcff in enumerate(fcff_fcst))
        tv_s   = fcff_fcst[-1] * (1 + tg) / (wc - tg)
        pv_tv_s = tv_s / (1 + wc)**5.0
        ev_s   = pv_sum + pv_tv_s
        eq_s   = ev_s - debt_fy26 + cash_fy26 - MINORITY_INT
        sens_matrix[ri, ci] = eq_s * 100 / SHARES_LAKH

# ─────────────────────────────────────────────
# 6.  PLOTTING — 6-PANEL DASHBOARD
# ─────────────────────────────────────────────

BLUE_HIST  = "#2E75B6"
ORANGE_FCST = "#C55A11"
GREEN_EBITDA = "#375623"
GREY_TEXT  = "#404040"

fig, axes = plt.subplots(2, 3, figsize=(18, 11))
fig.patch.set_facecolor('#F5F5F5')
fig.suptitle("Dixon Technologies (India) Ltd. — Financial Analysis Dashboard  v3.0\n"
             "Source: Screener.in  ·  Model: 3-Statement DCF  ·  WiDS 5.0 Financial Modelling Bootcamp",
             fontsize=13, fontweight='bold', color='#1F3864', y=0.98)

all_rev  = revenue_hist + rev_fcst
all_ebitda = ebitda_hist + ebitda_fcst
all_pat  = pat_hist + pat_fcst
all_ebitda_pct = [e/r*100 for e, r in zip(ebitda_hist + ebitda_fcst,
                                            revenue_hist + rev_fcst)]
pat_pct_all = [p/r*100 for p, r in zip(pat_hist + pat_fcst,
                                         revenue_hist + rev_fcst)]

colors_bar = [BLUE_HIST]*6 + [ORANGE_FCST]*5

# ── Panel 1: Revenue + Growth ──────────────────────────────
ax = axes[0, 0]
ax.set_facecolor('#FAFAFA')
bars = ax.bar(ALL_YEARS, [v/1000 for v in all_rev], color=colors_bar,
              edgecolor='white', linewidth=0.8, zorder=3)
ax.set_title("Revenue  (₹ '000 Cr)", fontweight='bold', color=GREY_TEXT, fontsize=10)
ax.set_ylabel("₹ '000 Cr", fontsize=8, color=GREY_TEXT)
ax.tick_params(axis='x', rotation=45, labelsize=8)
ax.tick_params(axis='y', labelsize=8)
ax.grid(axis='y', alpha=0.3, zorder=0)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)

# Add growth % labels
rev_growth_all = [None] + [(all_rev[i]/all_rev[i-1]-1)*100 for i in range(1, len(all_rev))]
for i, (bar, g) in enumerate(zip(bars, rev_growth_all)):
    if g is not None:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                f"+{g:.0f}%", ha='center', va='bottom', fontsize=6.5,
                color=ORANGE_FCST if i >= 6 else BLUE_HIST, fontweight='bold')

# ── Panel 2: Cost Structure  (stacked % of revenue) ─────────
ax = axes[0, 1]
ax.set_facecolor('#FAFAFA')
x = np.arange(len(ALL_YEARS))
all_mat_pct   = mat_pct_hist  + mat_pct_fcst
all_emp_pct   = emp_pct_hist  + emp_pct_fcst
all_oth_pct   = oth_pct_hist  + oth_pct_fcst
all_ebitda_bar = ebitda_pct_hist + ebitda_pct_fcst

ax.bar(x, all_mat_pct, label="Material Cost",  color="#C55A11", edgecolor='white', zorder=3)
ax.bar(x, all_emp_pct,  label="Employee",       color="#ED7D31",
       bottom=all_mat_pct, edgecolor='white', zorder=3)
bot_oth = [m+e for m, e in zip(all_mat_pct, all_emp_pct)]
ax.bar(x, all_oth_pct,  label="Other Opex",    color="#FFC000",
       bottom=bot_oth, edgecolor='white', zorder=3)
bot_ebitda = [m+e+o for m, e, o in zip(all_mat_pct, all_emp_pct, all_oth_pct)]
ax.bar(x, all_ebitda_bar, label="EBITDA",      color="#70AD47",
       bottom=bot_ebitda, edgecolor='white', zorder=3)

# PAT margin as line overlay
ax2 = ax.twinx()
ax2.plot(x, pat_pct_all, 'o-', color=BLUE_HIST, linewidth=1.5,
         markersize=4, label="PAT %", zorder=4)
ax2.set_ylabel("PAT %", fontsize=7, color=BLUE_HIST)
ax2.tick_params(axis='y', labelsize=7, colors=BLUE_HIST)
ax2.set_ylim(0, max(pat_pct_all) * 3.5)
ax2.spines['top'].set_visible(False)

ax.set_title("Revenue Cost Structure  (% of Revenue)", fontweight='bold',
             color=GREY_TEXT, fontsize=10)
ax.set_ylabel("% of Revenue", fontsize=8, color=GREY_TEXT)
ax.set_xticks(x); ax.set_xticklabels(ALL_YEARS, rotation=45, fontsize=8)
ax.tick_params(axis='y', labelsize=8)
ax.set_ylim(0, 105)
lines1, labels1 = ax.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax.legend(lines1 + lines2, labels1 + labels2, fontsize=6.5, framealpha=0.7,
          loc='upper left', ncol=2)
ax.grid(axis='y', alpha=0.3, zorder=0)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
ax.axvline(x=5.5, color='grey', linestyle='--', alpha=0.5, linewidth=1)
ax.text(5.6, 98, 'Forecast →', fontsize=7, color='grey', va='top')

# ── Panel 3: FCFF Waterfall (forecast years) ──────────────
ax = axes[0, 2]
ax.set_facecolor('#FAFAFA')
ax.set_title("FCFF Bridge  FY27E–FY31E  (₹ Cr)", fontweight='bold',
             color=GREY_TEXT, fontsize=10)

fcff_components = {
    "NOPAT":    nopat_fcst,
    "+ D&A":    da_fcst,
    "− Capex":  [-v for v in capex_fcst],
    "= FCFF":   fcff_fcst,
}

bar_positions = np.arange(len(FCST_YEARS))
comp_colors = {'NOPAT': '#2E75B6', '+ D&A': '#70AD47',
               '− Capex': '#FF0000', '= FCFF': '#C55A11'}

bottom = np.zeros(5)
for label, vals in list(fcff_components.items())[:-1]:
    ax.bar(bar_positions, vals, bottom=bottom, label=label,
           color=comp_colors[label], edgecolor='white', alpha=0.85, zorder=3)
    bottom += np.array(vals)

# FCFF as standalone bar
ax.bar(bar_positions + 0.4, fcff_fcst, 0.3, label="= FCFF",
       color=comp_colors["= FCFF"], edgecolor='white', zorder=3)
for i, (p, v) in enumerate(zip(bar_positions + 0.4, fcff_fcst)):
    ax.text(p, v + 20, f"₹{v:,.0f}", ha='center', va='bottom',
            fontsize=6.5, fontweight='bold', color=ORANGE_FCST)

ax.set_xticks(bar_positions + 0.2)
ax.set_xticklabels(FCST_YEARS, rotation=45, fontsize=8)
ax.tick_params(axis='y', labelsize=8)
ax.legend(fontsize=7, framealpha=0.7)
ax.grid(axis='y', alpha=0.3, zorder=0)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)

# ── Panel 4: EV Bridge Waterfall ──────────────────────────
ax = axes[1, 0]
ax.set_facecolor('#FAFAFA')
ax.set_title("EV → Equity Value Bridge  (₹ Cr)", fontweight='bold',
             color=GREY_TEXT, fontsize=10)

bridge_labels = ["PV FCFFs", "PV(TV)", "EV", "−Debt",
                 "+Cash", "−MI", "Equity\nValue"]
bridge_vals   = [sum_pv_fcff, pv_tv, ev, -debt_fy26, cash_fy26,
                 -MINORITY_INT, equity_val]

running = 0
bar_colors = []
bottoms_wb = []
bar_heights = []
for i, v in enumerate(bridge_vals):
    if bridge_labels[i] in ("EV", "Equity\nValue"):
        bar_colors.append('#1F3864')
        bottoms_wb.append(0)
        bar_heights.append(v)
    else:
        bar_colors.append('#70AD47' if v > 0 else '#FF0000')
        if bridge_labels[i] == "PV(TV)":
            bottoms_wb.append(sum_pv_fcff)
        elif bridge_labels[i] == "−Debt":
            bottoms_wb.append(equity_val + MINORITY_INT - cash_fy26)
        elif bridge_labels[i] == "+Cash":
            bottoms_wb.append(equity_val + MINORITY_INT - cash_fy26 + v + debt_fy26)
        elif bridge_labels[i] == "−MI":
            bottoms_wb.append(equity_val)
        else:
            bottoms_wb.append(0)
        bar_heights.append(abs(v))

# Simpler approach: just plot as grouped bars with annotations
bar_pos = np.arange(len(bridge_labels))
bclrs = [('#2E75B6' if v > 0 else '#FF4040') for v in bridge_vals]
bclrs[2] = '#1F3864'   # EV
bclrs[-1] = '#833C00'  # Equity Value
ax.bar(bar_pos, [abs(v)/1000 for v in bridge_vals], color=bclrs,
       edgecolor='white', zorder=3)
for i, (p, v) in enumerate(zip(bar_pos, bridge_vals)):
    ax.text(p, abs(v)/1000 + 100, f"₹{abs(v):,.0f}", ha='center',
            va='bottom', fontsize=6.5, fontweight='bold',
            color='#1F3864' if v > 0 else '#C00000')

ax.set_xticks(bar_pos)
ax.set_xticklabels(bridge_labels, fontsize=8)
ax.set_ylabel("₹ '000 Cr", fontsize=8, color=GREY_TEXT)
ax.tick_params(axis='y', labelsize=8)
ax.grid(axis='y', alpha=0.3, zorder=0)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)

# ── Panel 5: PV Composition Pie ───────────────────────────
ax = axes[1, 1]
ax.set_facecolor('#FAFAFA')
ax.set_title(f"EV Composition  (Total EV = ₹{ev/1000:.1f}k Cr)", fontweight='bold',
             color=GREY_TEXT, fontsize=10)

# Group negative FCFFs into a combined "negative FCFF years" bucket if needed
# Safe: only include positive PV values in the pie; label total for negatives
pos_pv  = [(fy, pv) for fy, pv in zip(FCST_YEARS, pv_fcffs) if pv > 0]
neg_sum = sum(pv for pv in pv_fcffs if pv <= 0)

year_pv_labels = [f"{fy}\n₹{pv:,.0f}" for fy, pv in pos_pv]
pie_vals       = [pv for _, pv in pos_pv]

if neg_sum < 0:
    year_pv_labels.append(f"Near-term\n(net neg)\n₹{neg_sum:,.0f}")
    pie_vals.append(abs(neg_sum))  # show magnitude

year_pv_labels.append(f"Terminal\n₹{pv_tv:,.0f}")
pie_vals.append(pv_tv)

all_colors = ['#2E75B6', '#41719C', '#5B9BD5', '#9DC3E6', '#BDD7EE',
              '#FF8C00', '#FFC000']
pie_colors = all_colors[:len(pie_vals)]
explode = [0] * (len(pie_vals) - 1) + [0.05]

wedges, texts, autotexts = ax.pie(
    pie_vals, labels=year_pv_labels, autopct='%1.1f%%',
    colors=pie_colors, explode=explode, startangle=90,
    textprops={'fontsize': 7},
    wedgeprops={'edgecolor': 'white', 'linewidth': 1.5})
for at in autotexts:
    at.set_fontsize(7); at.set_fontweight('bold')

# ── Panel 6: Sensitivity Heatmap ──────────────────────────
ax = axes[1, 2]
ax.set_facecolor('#FAFAFA')
ax.set_title(f"Implied Price Sensitivity  (₹)  ·  Base = ₹{implied_price:,.0f}",
             fontweight='bold', color=GREY_TEXT, fontsize=10)

green_red = LinearSegmentedColormap.from_list("gr", ["#C00000", "#FFFF00", "#70AD47"])
im = ax.imshow(sens_matrix, cmap=green_red, aspect='auto')

ax.set_xticks(range(len(tgr_range)))
ax.set_yticks(range(len(wacc_range)))
ax.set_xticklabels([f"{t*100:.1f}%" for t in tgr_range], fontsize=8)
ax.set_yticklabels([f"{w*100:.1f}%" for w in wacc_range], fontsize=8)
ax.set_xlabel("Terminal Growth Rate →", fontsize=8, color=GREY_TEXT)
ax.set_ylabel("← WACC", fontsize=8, color=GREY_TEXT)

for ri in range(len(wacc_range)):
    for ci in range(len(tgr_range)):
        val = sens_matrix[ri, ci]
        color = 'white' if val < sens_matrix.mean() else 'black'
        is_base = (abs(wacc_range[ri] - WACC) < 0.003 and
                   abs(tgr_range[ci] - TGR) < 0.003)
        txt = ax.text(ci, ri, f"₹{val:,.0f}", ha='center', va='center',
                      fontsize=7, color=color,
                      fontweight='bold' if is_base else 'normal')
        if is_base:
            rect = plt.Rectangle((ci-0.5, ri-0.5), 1, 1,
                                  fill=False, edgecolor='yellow', linewidth=2)
            ax.add_patch(rect)

fig.colorbar(im, ax=ax, fraction=0.04, pad=0.04,
             label="Implied Price (₹)")

plt.tight_layout(rect=[0, 0, 1, 0.95])

OUTPUT_PNG = "Dixon_Analysis_Dashboard_v3.png"
plt.savefig(OUTPUT_PNG, dpi=150, bbox_inches='tight',
            facecolor='#F5F5F5', edgecolor='none')
plt.close()
print(f"✅  Saved: {OUTPUT_PNG}")
print(f"   Implied Price: ₹{implied_price:,.0f} / share")
print(f"   EV: ₹{ev:,.0f} Cr  (PV FCFFs: ₹{sum_pv_fcff:,.0f}  +  PV TV: ₹{pv_tv:,.0f})")
