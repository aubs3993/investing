# Aggregate US nonfinancial-corporate EV multiples, quarterly 1952Q4-present,
# built entirely from FRED (Z.1 Financial Accounts stocks/flows + NIPA Table
# 1.14 flows): EV/EBIT, EV/EBITDA, EV/unlevered-FCF, MarketCap/levered-FCF,
# plus MarketCap/(dividends + net buybacks) and a Baa-Aaa credit overlay.
#
# INTENTIONAL EXCEPTION to the repo's 2006-01-01 start convention: valuation
# multiples are only meaningful across full cycles (1950s-60s bull, Nifty
# Fifty, 1970s-80s trough, dot-com, GFC, 2021), so primary charts run
# 1952Q4-present. Pre-1952 Z.1 data is annual-frequency / lower quality, so
# the chart window starts 1952Q4 even though the pull buffers from 1945.
# A 2006+ companion chart (ev_multiples_2006.png) is emitted so the headline
# multiples also slot into the comparable-axis macro chart set.
#
# LEVEL-COMPARABILITY WARNING: the NCB (nonfinancial corporate business)
# sector = ALL public AND private US nonfinancial corporates, and the NIPA
# earnings are DOMESTIC-only (national accounts measure domestic production).
# These multiples are NOT comparable to the S&P 500's EV/EBITDA (~15x):
#   - Foreign-earnings wedge: S&P constituents earn ~30-40% of profits abroad;
#     those earnings are in their market caps (numerator here) but NOT in
#     NIPA domestic earnings (denominator). The wedge has GROWN over time as
#     US corporates globalized, biasing this multiple structurally UP across
#     decades — another reason to read percentiles-vs-own-history, not levels.
#   - Debt is at BOOK value in Z.1 (debt securities + loans), not market.
#   - Macro EBITDA != accounting EBITDA: consumption of fixed capital (CFC)
#     is ECONOMIC depreciation at current cost, profits carry IVA/CCAdj, and
#     there is no stock-comp addback.
# Use percentile rank vs. this series' own history ONLY.
#
# FCF legs are structurally depressed: the sector aggregate expenses ALL
# growth capex of every US nonfinancial corporate (including private firms
# investing heavily), so aggregate FCF is far thinner than any listed-company
# FCF — and goes genuinely NEGATIVE (mid-1970s, ~2000). Raw multiples are
# meaningless around zero crossings, so the FCF legs are charted as YIELDS
# (smoothed FCF / EV or MktCap, in %). The xlsx still carries the multiples
# but masks them (NaN) where the smoothed denominator < 1% of EV / MktCap.
# For the same reason MC/LFCF's ~44x full-history median is NOT comparable
# to a company P/FCF.
#
# Units: Z.1 stocks/flows are MILLIONS of $ (flows SAAR); NIPA flows are
# BILLIONS of $ SAAR. Everything is converted to billions (Z.1 / 1000).
# Flows are SAAR (already annualized), so the trailing-year denominator is
# the trailing 4-quarter rolling MEAN (not sum); the noisier FCF legs use an
# 8-quarter rolling mean.

from datetime import datetime
from pathlib import Path
import sys

import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.fred_helpers import (
    get_fred_client,
    get_recession_periods,
    pull_series,
    resolve_output_dir,
    series_stats,
    style_macro_chart,
)

fred = get_fred_client()

end = datetime.today()
chart_start = datetime(1952, 10, 1)   # 1952Q4: true quarterly Z.1 coverage
chart_start_2006 = datetime(2006, 1, 1)
data_start = datetime(1945, 1, 1)     # buffer; Z.1 stocks exist from 1945Q4

quarterly_series = {
    # --- Z.1 stocks, MILLIONS of $, NSA, end of period ---
    "EQ_MM": "NCBCEL",                  # NCB equities, market value (= NCBEILQ027S)
    "DEBT_MM": "BCNSDODNS",             # NCB debt securities + loans (book), SA
    "LIQ_MM": "BOGZ1FL104001005Q",      # NCB liquid assets, broad measure
    # --- NIPA Table 1.14 flows, BILLIONS of $, SAAR ---
    "PROFIT_PRETAX": "A463RC1Q027SBEA",  # NCB pretax profits w/ IVA & CCAdj
    "NET_INTEREST": "B471RC1Q027SBEA",   # NCB net interest & misc payments
    "CFC": "B456RC1Q027SBEA",            # NCB consumption of fixed capital
    "TAXES": "B465RC1Q027SBEA",          # NCB taxes on corporate income
    "PROFIT_AT": "W328RC1Q027SBEA",      # NCB after-tax profits w/ IVA & CCAdj
    "GVA": "A455RC1Q027SBEA",            # NCB gross value added
    # --- Z.1 flows, MILLIONS of $, SAAR ---
    "CAPEX_MM": "BOGZ1FA105050005Q",     # NCB total capital expenditures
    "DIV_MM": "BOGZ1FA106121075Q",       # NCB net dividends paid
    "EQ_ISS_MM": "NCBCEBQ027S",          # NCB net equity issuance (neg = buybacks)
}
monthly_series = {
    "BAA": "BAA",   # Moody's Baa corporate yield, monthly
    "AAA": "AAA",   # Moody's Aaa corporate yield, monthly
}

df = pull_series(fred, quarterly_series, data_start, end)

# Z.1 millions -> billions so everything matches the NIPA series' units.
for mm_col, bil_col in [("EQ_MM", "MKTCAP"), ("DEBT_MM", "DEBT"),
                        ("LIQ_MM", "LIQ"), ("CAPEX_MM", "CAPEX"),
                        ("DIV_MM", "DIV"), ("EQ_ISS_MM", "EQ_ISS")]:
    df[bil_col] = df[mm_col] / 1000.0
df = df.drop(columns=["EQ_MM", "DEBT_MM", "LIQ_MM", "CAPEX_MM",
                      "DIV_MM", "EQ_ISS_MM"])

# --- Levels and earnings flows (all Bil$; flows SAAR) ---
df["EV"] = df["MKTCAP"] + df["DEBT"] - df["LIQ"]
df["EBIT"] = df["PROFIT_PRETAX"] + df["NET_INTEREST"]
df["EBITDA"] = df["EBIT"] + df["CFC"]
df["UFCF"] = df["EBIT"] - df["TAXES"] + df["CFC"] - df["CAPEX"]
df["LFCF"] = df["PROFIT_AT"] + df["CFC"] - df["CAPEX"]
# Cash actually returned to shareholders: dividends + net buybacks (net equity
# issuance is negative when buybacks dominate, so subtracting it adds them).
df["LFCF_DIST"] = df["DIV"] - df["EQ_ISS"]

# Flows are SAAR, so trailing-year smoothing = rolling MEAN (not sum).
# 4q for the earnings legs and distributions; 8q for the noisier FCF legs.
df["EBIT_4Q"] = df["EBIT"].rolling(4).mean()
df["EBITDA_4Q"] = df["EBITDA"].rolling(4).mean()
df["LFCF_DIST_4Q"] = df["LFCF_DIST"].rolling(4).mean()
df["UFCF_8Q"] = df["UFCF"].rolling(8).mean()
df["LFCF_8Q"] = df["LFCF"].rolling(8).mean()

# --- Ratios and yields ---
df["EV_EBIT"] = df["EV"] / df["EBIT_4Q"]
df["EV_EBITDA"] = df["EV"] / df["EBITDA_4Q"]
df["MC_LFCF_DIST"] = df["MKTCAP"] / df["LFCF_DIST_4Q"]
# FCF charted as yields: aggregate FCF crosses zero (mid-1970s, ~2000), where
# a multiple is meaningless but a yield stays interpretable.
df["UFCF_YIELD"] = df["UFCF_8Q"] / df["EV"] * 100
df["LFCF_YIELD"] = df["LFCF_8Q"] / df["MKTCAP"] * 100
df["LFCF_DIST_YIELD"] = df["LFCF_DIST_4Q"] / df["MKTCAP"] * 100
# FCF multiples kept for the xlsx, but masked (NaN) where the smoothed
# denominator < 1% of EV / MktCap — near zero-crossings the multiple explodes
# and flips sign, so those readings carry no information.
df["EV_UFCF"] = (df["EV"] / df["UFCF_8Q"]).where(df["UFCF_8Q"] >= 0.01 * df["EV"])
df["MC_LFCF"] = (df["MKTCAP"] / df["LFCF_8Q"]).where(
    df["LFCF_8Q"] >= 0.01 * df["MKTCAP"])
# Hussman-style market cap / gross value added and EBITDA margin on GVA.
df["MC_GVA"] = df["MKTCAP"] / df["GVA"]
df["EBITDA_MARGIN_GVA"] = df["EBITDA"] / df["GVA"] * 100

# Baa-Aaa credit spread: monthly, averaged to quarters. resample("QS") labels
# quarters by their start date, matching the quarterly FRED timestamps above.
df_credit = pull_series(fred, monthly_series, data_start, end)
df_credit["BAA_AAA_SPREAD"] = df_credit["BAA"] - df_credit["AAA"]
spread_q = (df_credit.set_index("Date")["BAA_AAA_SPREAD"]
            .resample("QS").mean().rename("BAA_AAA_SPREAD").reset_index())
df = df.merge(spread_q, on="Date", how="left")

# Drop the 1945-1952Q3 buffer (annual-frequency / warm-up rows).
df = df[df["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)

RATIO_COLS = [
    "EV_EBIT", "EV_EBITDA", "EV_UFCF", "MC_LFCF", "MC_LFCF_DIST",
    "UFCF_YIELD", "LFCF_YIELD", "LFCF_DIST_YIELD",
    "MC_GVA", "EBITDA_MARGIN_GVA", "BAA_AAA_SPREAD",
]

# Summary: series_stats over the full 1952Q4+ window (the charted window for
# the primary charts) plus the current value's percentile rank within it.
summary_rows = {c: series_stats(df[c]) for c in RATIO_COLS}
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]
summary["current_pctile"] = [
    (df[c].dropna() <= df[c].dropna().iloc[-1]).mean() * 100 for c in RATIO_COLS
]

OUT_DIR = resolve_output_dir(__file__, "ev_multiples")
(OUT_DIR / ".gitkeep").touch()
xlsx_path = OUT_DIR / "ev_multiples.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)
    summary.to_excel(writer, sheet_name="Summary")

recessions = get_recession_periods(fred, chart_start, end)
XLIM = (pd.Timestamp(chart_start), pd.Timestamp(end))

ev_ebit_median = df["EV_EBIT"].median()
ev_ebitda_median = df["EV_EBITDA"].median()
mc_gva_median = df["MC_GVA"].median()
margin_median = df["EBITDA_MARGIN_GVA"].median()

# Chart 1 — EV/EBIT and EV/EBITDA, two stacked panels, 1952Q4-present.
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
ax1.plot(df["Date"], df["EV_EBIT"], color="#1f3b73", linewidth=1.6,
         label="EV / EBIT (4q avg)")
ax1.set_xlim(*XLIM)
style_macro_chart(
    ax1,
    title="US nonfinancial corporate EV/EBIT (Z.1 + NIPA), 1952Q4–present",
    ylabel="Multiple (x)",
    recessions=recessions,
    hlines=[{"y": ev_ebit_median,
             "label": f"Median {ev_ebit_median:.1f}x (1952Q4–present)"}],
)
ax2.plot(df["Date"], df["EV_EBITDA"], color="#1f3b73", linewidth=1.6,
         label="EV / EBITDA (4q avg)")
ax2.set_xlim(*XLIM)
style_macro_chart(
    ax2,
    title="US nonfinancial corporate EV/EBITDA (Z.1 + NIPA), 1952Q4–present",
    ylabel="Multiple (x)",
    recessions=recessions,
    hlines=[{"y": ev_ebitda_median,
             "label": f"Median {ev_ebitda_median:.1f}x (1952Q4–present)"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "ev_ebit_ebitda.png", dpi=150)
plt.close(fig)

# Chart 2 — FCF yields (not multiples: aggregate FCF crosses zero, mid-1970s
# and ~2000, so only the yield stays interpretable through those episodes).
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["UFCF_YIELD"], color="#1f3b73", linewidth=1.6,
        label="Unlevered FCF / EV (8q avg)")
ax.plot(df["Date"], df["LFCF_YIELD"], color="#9ec5e8", linewidth=1.4,
        label="Levered FCF / MktCap (8q avg)")
ax.plot(df["Date"], df["LFCF_DIST_YIELD"], color="#2ca02c", linewidth=1.4,
        label="Dividends + net buybacks / MktCap (4q avg)")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="US nonfinancial corporate FCF yields, 1952Q4–present",
    ylabel="Percent",
    recessions=recessions,
    hlines=[{"y": 0.0}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "fcf_yields.png", dpi=150)
plt.close(fig)

# Chart 3 — MktCap/GVA (Hussman-style sales-proxy multiple) and the EBITDA
# margin on GVA, two stacked panels: together they decompose the EV/EBITDA
# move into valuation-per-unit-of-output vs. profitability-of-output.
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
ax1.plot(df["Date"], df["MC_GVA"], color="#1f3b73", linewidth=1.6,
         label="MktCap / gross value added")
ax1.set_xlim(*XLIM)
style_macro_chart(
    ax1,
    title="US nonfinancial corporate MktCap / gross value added, 1952Q4–present",
    ylabel="Ratio (x)",
    recessions=recessions,
    hlines=[{"y": mc_gva_median,
             "label": f"Median {mc_gva_median:.2f}x (1952Q4–present)"}],
)
ax2.plot(df["Date"], df["EBITDA_MARGIN_GVA"], color="#1f3b73", linewidth=1.6,
         label="EBITDA / gross value added")
ax2.set_xlim(*XLIM)
style_macro_chart(
    ax2,
    title="US nonfinancial corporate EBITDA margin on GVA, 1952Q4–present",
    ylabel="Percent",
    recessions=recessions,
    hlines=[{"y": margin_median,
             "label": f"Median {margin_median:.1f}% (1952Q4–present)"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "marketcap_gva.png", dpi=150)
plt.close(fig)

# Chart 4 — 2006+ companion of the headline multiples so they slot into the
# comparable-axis macro chart set. Medians recomputed over the 2006+ window
# (what's on this chart), per repo convention for reference stats.
df_2006 = df[df["Date"] >= pd.Timestamp(chart_start_2006)].reset_index(drop=True)
recessions_2006 = [r for r in recessions if r[1] >= pd.Timestamp(chart_start_2006)]
ev_ebit_median_06 = df_2006["EV_EBIT"].median()
ev_ebitda_median_06 = df_2006["EV_EBITDA"].median()
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_2006["Date"], df_2006["EV_EBIT"], color="#9ec5e8", linewidth=1.4,
        label="EV / EBIT (4q avg)")
ax.plot(df_2006["Date"], df_2006["EV_EBITDA"], color="#1f3b73", linewidth=1.6,
        label="EV / EBITDA (4q avg)")
ax.set_xlim(pd.Timestamp(chart_start_2006), pd.Timestamp(end))
style_macro_chart(
    ax,
    title="US nonfinancial corporate EV/EBIT and EV/EBITDA, 2006–present",
    ylabel="Multiple (x)",
    recessions=recessions_2006,
    hlines=[
        {"y": ev_ebit_median_06, "color": "#9ec5e8",
         "label": f"EV/EBIT median 2006+ ({ev_ebit_median_06:.1f}x)"},
        {"y": ev_ebitda_median_06, "color": "#1f3b73",
         "label": f"EV/EBITDA median 2006+ ({ev_ebitda_median_06:.1f}x)"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "ev_multiples_2006.png", dpi=150)
plt.close(fig)

# Chart 5 — valuation vs credit: EV/EBITDA on top, Baa-Aaa spread below,
# plotted INVERTED (x -1) so that UP = tight spreads = easy credit = the
# regime that mechanically supports high multiples. The two panels moving up
# together is the froth signature (1999-2000, 2007, 2021).
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
ax1.plot(df["Date"], df["EV_EBITDA"], color="#1f3b73", linewidth=1.6,
         label="EV / EBITDA (4q avg)")
ax1.set_xlim(*XLIM)
style_macro_chart(
    ax1,
    title="US nonfinancial corporate EV/EBITDA vs credit spreads, 1952Q4–present",
    ylabel="Multiple (x)",
    recessions=recessions,
    hlines=[{"y": ev_ebitda_median,
             "label": f"Median {ev_ebitda_median:.1f}x (1952Q4–present)"}],
)
ax2.plot(df["Date"], -df["BAA_AAA_SPREAD"], color="#c0392b", linewidth=1.4,
         label="Baa − Aaa spread, inverted: up = tight spreads / froth")
ax2.set_xlim(*XLIM)
style_macro_chart(
    ax2,
    title="Moody's Baa − Aaa spread (quarterly avg, inverted)",
    ylabel="− Percentage points",
    recessions=recessions,
)
fig.tight_layout(rect=(0, 0.03, 1, 1))
fig.text(0.01, 0.005,
         "Note: 1999–2000, 2007 and 2021 all show the high-multiple + "
         "tight-spread signature — tight credit mechanically supports high "
         "multiples, and both reverse together.",
         fontsize=8, color="0.35")
fig.savefig(OUT_DIR / "valuation_vs_credit.png", dpi=150)
plt.close(fig)

# --- Print summary ---
latest = df.dropna(subset=["EV_EBITDA"]).iloc[-1]
print(f"Chart start:             {chart_start.date()} (1952Q4)")
print(f"End date:                {end.date()}")
print(f"Rows (1952Q4+):          {len(df)}")
print(f"Latest ratio quarter:    {latest['Date'].date()}")
print(f"Latest EV:               ${latest['EV']:,.0f}B "
      f"(equity {latest['MKTCAP']:,.0f} + debt {latest['DEBT']:,.0f} "
      f"- liquid {latest['LIQ']:,.0f})")
print(f"Latest EV/EBITDA (4q):   {latest['EV_EBITDA']:.1f}x "
      f"(median {ev_ebitda_median:.1f}x, "
      f"{summary.loc['EV_EBITDA', 'current_pctile']:.0f}th pctile since 1952Q4)")
print(f"Latest EV/EBIT (4q):     {latest['EV_EBIT']:.1f}x "
      f"(median {ev_ebit_median:.1f}x, "
      f"{summary.loc['EV_EBIT', 'current_pctile']:.0f}th pctile)")
print(f"Latest UFCF yield (8q):  {df['UFCF_YIELD'].dropna().iloc[-1]:.2f}%")
print(f"Latest LFCF yield (8q):  {df['LFCF_YIELD'].dropna().iloc[-1]:.2f}%")
print(f"Latest dist yield (4q):  {df['LFCF_DIST_YIELD'].dropna().iloc[-1]:.2f}%")
print(f"Latest MC/dist (4q):     {df['MC_LFCF_DIST'].dropna().iloc[-1]:.1f}x")
print(f"MC/LFCF median (masked): {df['MC_LFCF'].median():.1f}x "
      f"(NOT comparable to a company P/FCF — see header)")
print(f"Latest MktCap/GVA:       {df['MC_GVA'].dropna().iloc[-1]:.2f}x")
print(f"Latest Baa-Aaa spread:   {df['BAA_AAA_SPREAD'].dropna().iloc[-1]:.2f}pp")
print(f"Wrote {xlsx_path.name}, ev_ebit_ebitda.png, fcf_yields.png, "
      f"marketcap_gva.png, ev_multiples_2006.png, valuation_vs_credit.png "
      f"to {OUT_DIR}")
