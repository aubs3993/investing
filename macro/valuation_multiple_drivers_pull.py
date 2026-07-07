# Valuation-multiple DRIVERS for the US nonfinancial corporate sector, quarterly
# 1947Q1-present: the interpretation/decomposition layer for the aggregate EV
# multiples. Five companion series that explain WHY the multiple sits where it
# does: (1) D&A share of EBITDA (the EV/EBIT-vs-EV/EBITDA wedge), (2) effective
# federal corporate tax rate, (3) EBIT/EBITDA yields vs the real 10y Treasury,
# (4) EBITDA margin on gross value added and EV/GVA (margin vs multiple split:
# EV/EBITDA = (EV/GVA) / (EBITDA/GVA)), (5) intangible-investment share of
# nonresidential fixed investment (structural drift toward asset-light).
#
# INTENTIONAL EXCEPTION to the repo's 2006-01-01 start convention: these are
# structural, decades-scale drivers — the whole point is the 1947-present
# trend. A 2006+ companion of the overview (drivers_overview_2006.png) is
# emitted so it slots into the comparable-axis macro chart set.
#
# Conventions / construction:
# - With-IVA+CCAdj profits convention throughout (A463 family): EBIT = A463
#   (NF corporate profits w/ IVA+CCAdj) + B471 (net interest & misc payments);
#   EBITDA = EBIT + B456 (NF corporate consumption of fixed capital = D&A).
#   ONE exception: the effective-tax-rate DENOMINATOR is A053 (profits before
#   tax WITHOUT IVA+CCAdj) because taxes are levied on book-style profits, and
#   that pair is economy-wide (incl. financials) — the only all-corporate panel.
# - Effective tax rate numerator B075 is FEDERAL taxes on corporate income
#   only — state/local corporate taxes are excluded, so the level understates
#   the total effective rate by a few points throughout.
# - IMPORTANT (verified in the data): the effective rate does NOT step down at
#   TRA86 — it ROSE from ~20% to ~25% over 1986-88 because base-broadening
#   outweighed the 46%->34% statutory cut. Only TCJA (2018) shows a clean
#   step: ~14% in 2016 down to ~9-10% in 2018-19.
# - EV = (NCBCEL equities + BCNSDODNS debt securities & loans − BOGZ1FL104001005Q
#   liquid assets, all NF corporate Z.1 levels in $MM) / 1000 -> $BN, matching
#   the SAAR flow units. Z.1 levels are ANNUAL (Q4-only) before 1952, so
#   EV-based series are sparse pre-1952; plots connect via dropna segments.
# - Real 10y Treasury: REAINTRATREARAT10Y (Cleveland Fed model, monthly 1982+),
#   SPLICED before 1982 with a proxy = GS10 minus trailing-10y annualized
#   CPIAUCSL inflation. The proxy segment starts 1957 (GS10 starts 1953-04 and
#   the CPI lookback needs 10 years from 1947) and is charted dashed/gray and
#   labeled as proxy. Monthly rates are averaged to quarterly.
#   SPLICE GAP: the backward-looking proxy sits ~250bp BELOW the Cleveland
#   forward-looking model at the boundary (Dec-1981 proxy 5.1% vs Jan-1982
#   model 7.6%) because realized 1970s inflation exceeded early-80s expected
#   inflation. Do not read pre- vs post-1982 REAL10Y/EBIT_SPREAD levels as
#   directly comparable; the level shift is a construction artifact.
# - EV/GVA uses 4q-average GVA and the margin uses 4q-average EBITDA / 4q-average
#   GVA, so the identity EV/EBITDA(4q) = (EV/GVA) / margin holds exactly.
#
# MEAN-NONSTATIONARITY WARNING: tax cuts, the rising D&A share, intangible
# intensity, and sector-mix shift all move the JUSTIFIED multiple up across
# decades. Naive reversion of today's EV/EBITDA to its 1975 mean is the single
# biggest interpretive error this chart set exists to prevent.
#
# FUTURE WORK (deliberately deferred — extra external downloads):
# - Damodaran implied-ERP annual overlay on the yield-spread panel.
# - Ken French tech-share-of-market-cap composition panel.

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
chart_start = datetime(1947, 1, 1)      # NIPA flows begin 1947Q1
chart_start_2006 = datetime(2006, 1, 1)
data_start = datetime(1945, 1, 1)       # Z.1 levels exist from 1945Q4
monthly_start = datetime(1947, 1, 1)    # CPI history for the 10y-trailing lookback

quarterly_series = {
    # NIPA flows, $BN SAAR (nonfinancial corporate unless noted).
    "PROFITS_IVA": "A463RC1Q027SBEA",   # profits w/ IVA+CCAdj, domestic NF
    "NET_INTEREST": "B471RC1Q027SBEA",  # net interest & misc payments
    "DA": "B456RC1Q027SBEA",            # consumption of fixed capital (D&A)
    "GVA": "A455RC1Q027SBEA",           # gross value added, NF corporate
    "FED_TAX": "B075RC1Q027SBEA",       # FEDERAL taxes on corporate income (all corp)
    "PBT": "A053RC1Q027SBEA",           # profits before tax w/o IVA+CCAdj (all corp)
    "IPP": "Y001RC1Q027SBEA",           # intellectual property products investment
    "PNFI": "PNFI",                     # private nonresidential fixed investment
    # Z.1 levels, $MM, end of period (annual Q4-only before 1952).
    "EQUITIES_MM": "NCBCEL",            # NF corporate equities, market value
    "DEBT_MM": "BCNSDODNS",             # NF corporate debt securities + loans
    "LIQUID_MM": "BOGZ1FL104001005Q",   # NF corporate liquid assets (broad)
}
monthly_series = {
    "REAL10Y_CLEV": "REAINTRATREARAT10Y",  # Cleveland Fed 10y real rate, 1982+
    "GS10": "GS10",                        # nominal 10y, 1953-04+
    "CPI": "CPIAUCSL",
}

df = pull_series(fred, quarterly_series, data_start, end)

# --- Monthly block: build the spliced real 10y, then average to quarterly ---
df_mo = pull_series(fred, monthly_series, monthly_start, end).set_index("Date")
# Trailing-10y annualized CPI inflation (120 months), %.
infl_10y = ((df_mo["CPI"] / df_mo["CPI"].shift(120)) ** (1 / 10) - 1) * 100
df_mo["REAL10Y_PROXY"] = df_mo["GS10"] - infl_10y  # defined 1957-01 onward
# Splice: Cleveland Fed model where available (1982+), proxy before.
df_mo["REAL10Y"] = df_mo["REAL10Y_CLEV"].combine_first(df_mo["REAL10Y_PROXY"])
df_q_rates = (
    df_mo[["REAL10Y_CLEV", "REAL10Y_PROXY", "REAL10Y"]]
    .resample("QS").mean().reset_index()
)
# Keep the proxy column only where the Cleveland series is absent, so the
# charted proxy segment ends exactly at the 1982 splice point.
df_q_rates.loc[df_q_rates["REAL10Y_CLEV"].notna(), "REAL10Y_PROXY"] = pd.NA
df_q_rates["REAL10Y_PROXY"] = df_q_rates["REAL10Y_PROXY"].astype(float)
df = df.merge(df_q_rates, on="Date", how="left")

# --- Derived columns ---
df["EBIT"] = df["PROFITS_IVA"] + df["NET_INTEREST"]
df["EBITDA"] = df["EBIT"] + df["DA"]
df["EBIT_4Q"] = df["EBIT"].rolling(4).mean()
df["EBITDA_4Q"] = df["EBITDA"].rolling(4).mean()
df["GVA_4Q"] = df["GVA"].rolling(4).mean()

# (1) D&A share of EBITDA — the EBIT-vs-EBITDA wedge.
df["DA_SHARE"] = df["DA"] / df["EBITDA"] * 100
df["DA_SHARE_4Q"] = df["DA_SHARE"].rolling(4).mean()

# (2) Effective FEDERAL corporate tax rate on book-style pretax profits.
df["EFF_TAX"] = df["FED_TAX"] / df["PBT"] * 100
df["EFF_TAX_4Q"] = df["EFF_TAX"].rolling(4).mean()

# (3) EV and earnings yields vs real 10y. $MM -> $BN to match SAAR flows.
df["EV"] = (df["EQUITIES_MM"] + df["DEBT_MM"] - df["LIQUID_MM"]) / 1000.0
df["EBIT_YIELD"] = df["EBIT_4Q"] / df["EV"] * 100
df["EBITDA_YIELD"] = df["EBITDA_4Q"] / df["EV"] * 100
df["EBIT_SPREAD"] = df["EBIT_YIELD"] - df["REAL10Y"]

# (4) Margin-vs-multiple split. Both legs on 4q-average GVA/EBITDA so
# EV/EBITDA(4q) = EV_GVA / (EBITDA_MARGIN/100) holds exactly.
df["EBITDA_MARGIN"] = df["EBITDA_4Q"] / df["GVA_4Q"] * 100
df["EV_GVA"] = df["EV"] / df["GVA_4Q"]

# (5) Intangible-investment share of nonresidential fixed investment.
df["INTANG_SHARE"] = df["IPP"] / df["PNFI"] * 100
df["INTANG_SHARE_4Q"] = df["INTANG_SHARE"].rolling(4).mean()

# Drop the 1945Q4-1946Q4 buffer rows (Z.1 levels only, no NIPA flows).
df = df[df["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)
df = df.drop(columns=["EQUITIES_MM", "DEBT_MM", "LIQUID_MM"])

DERIVED_COLS = [
    "DA_SHARE", "DA_SHARE_4Q",
    "EFF_TAX", "EFF_TAX_4Q",
    "EBIT_YIELD", "EBITDA_YIELD", "REAL10Y", "EBIT_SPREAD",
    "EBITDA_MARGIN", "EV_GVA",
    "INTANG_SHARE", "INTANG_SHARE_4Q",
]

# Summary over the full 1947+ window (the primary charted window), plus the
# current value's percentile rank within that window.
summary_rows = {c: series_stats(df[c]) for c in DERIVED_COLS}
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]
summary["current_pctile"] = [
    (df[c].dropna() <= df[c].dropna().iloc[-1]).mean() * 100 for c in DERIVED_COLS
]

OUT_DIR = resolve_output_dir(__file__, "valuation_multiple_drivers")
(OUT_DIR / ".gitkeep").touch()
xlsx_path = OUT_DIR / "valuation_multiple_drivers.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)
    summary.to_excel(writer, sheet_name="Summary")

recessions = get_recession_periods(fred, chart_start, end)
XLIM = (pd.Timestamp(chart_start), pd.Timestamp(end))
TAX_VLINES = [
    (pd.Timestamp("1987-01-01"), "TRA86 phase-in"),
    (pd.Timestamp("1993-01-01"), "OBRA93"),
    (pd.Timestamp("2018-01-01"), "TCJA"),
]
# Z.1 levels are annual (Q4-only) pre-1952 -> EV-based columns have embedded
# NaNs; dropna frames let matplotlib connect the early points with segments.
# The two real-rate segments (proxy / Cleveland) are charted but summarized
# only via the spliced REAL10Y, so they're added here rather than DERIVED_COLS.
plot_df = {
    c: df[["Date", c]].dropna()
    for c in DERIVED_COLS + ["REAL10Y_PROXY", "REAL10Y_CLEV"]
}

# --- Chart 1: D&A share of EBITDA ---
da_median = df["DA_SHARE_4Q"].median()
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["DA_SHARE"], color="#9ec5e8", linewidth=0.9,
        label="Quarterly")
ax.plot(df["Date"], df["DA_SHARE_4Q"], color="#1f3b73", linewidth=1.8,
        label="4-quarter mean")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="D&A share of EBITDA — US nonfinancial corporate, 1947–present",
    ylabel="Percent of EBITDA",
    recessions=recessions,
    hlines=[{"y": da_median, "label": f"Median {da_median:.0f}% (1947–present)"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "da_share.png", dpi=150)
plt.close(fig)

# --- Chart 2: effective federal corporate tax rate ---
tax_median = df["EFF_TAX_4Q"].median()
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["EFF_TAX"], color="#9ec5e8", linewidth=0.9,
        label="Quarterly")
ax.plot(df["Date"], df["EFF_TAX_4Q"], color="#1f3b73", linewidth=1.8,
        label="4-quarter mean")
for x, lbl in TAX_VLINES:
    ax.axvline(x, color="#c0392b", linestyle=":", linewidth=1.2, label=lbl)
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Effective FEDERAL corporate tax rate — federal receipts / pretax book profits, 1947–present",
    ylabel="Percent",
    recessions=recessions,
    hlines=[{"y": tax_median, "label": f"Median {tax_median:.0f}% (1947–present)"}],
)
fig.text(0.01, 0.01,
         "Federal taxes only (B075) — excludes state/local. Note the rate RISES "
         "~20%→~25% through 1986–88: TRA86 base-broadening outweighed the "
         "statutory cut. Only TCJA (2018) is a clean step down.",
         fontsize=7, color="0.35")
fig.tight_layout(rect=(0, 0.03, 1, 1))
fig.savefig(OUT_DIR / "effective_tax_rate.png", dpi=150)
plt.close(fig)

# --- Chart 3: EBIT/EBITDA yields vs real 10y, plus spread subpanel ---
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
ax1.plot(plot_df["EBITDA_YIELD"]["Date"], plot_df["EBITDA_YIELD"]["EBITDA_YIELD"],
         color="#9ec5e8", linewidth=1.4, label="EBITDA yield (EBITDA 4q / EV)")
ax1.plot(plot_df["EBIT_YIELD"]["Date"], plot_df["EBIT_YIELD"]["EBIT_YIELD"],
         color="#1f3b73", linewidth=1.8, label="EBIT yield (EBIT 4q / EV)")
ax1.plot(plot_df["REAL10Y_PROXY"]["Date"], plot_df["REAL10Y_PROXY"]["REAL10Y_PROXY"],
         color="0.55", linewidth=1.4, linestyle="--",
         label="Real 10y proxy, pre-1982 (GS10 − trailing-10y CPI)")
ax1.plot(plot_df["REAL10Y_CLEV"]["Date"], plot_df["REAL10Y_CLEV"]["REAL10Y_CLEV"],
         color="#c0392b", linewidth=1.6, label="Real 10y (Cleveland Fed, 1982+)")
ax1.set_xlim(*XLIM)
style_macro_chart(
    ax1,
    title="Earnings yields on EV vs real 10y Treasury — US nonfinancial corporate, 1947–present",
    ylabel="Percent",
    recessions=recessions,
)
spread_median = df["EBIT_SPREAD"].median()
ax2.plot(plot_df["EBIT_SPREAD"]["Date"], plot_df["EBIT_SPREAD"]["EBIT_SPREAD"],
         color="#2ca02c", linewidth=1.6, label="EBIT yield − real 10y")
ax2.set_xlim(*XLIM)
style_macro_chart(
    ax2,
    title="Spread: EBIT yield minus real 10y (pre-1982 uses spliced proxy)",
    ylabel="Percentage points",
    recessions=recessions,
    hlines=[
        {"y": 0.0},
        {"y": spread_median, "label": f"Median {spread_median:.1f}pp",
         "color": "#2ca02c"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "earnings_yield_vs_real10y.png", dpi=150)
plt.close(fig)

# --- Chart 4: EBITDA margin on GVA + EV/GVA (margin vs multiple split) ---
margin_median = df["EBITDA_MARGIN"].median()
ev_gva_median = df["EV_GVA"].median()
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
ax1.plot(plot_df["EBITDA_MARGIN"]["Date"], plot_df["EBITDA_MARGIN"]["EBITDA_MARGIN"],
         color="#1f3b73", linewidth=1.8, label="EBITDA / GVA (4q avgs)")
ax1.set_xlim(*XLIM)
style_macro_chart(
    ax1,
    title="EBITDA margin on gross value added — US nonfinancial corporate, 1947–present",
    ylabel="Percent of GVA",
    recessions=recessions,
    hlines=[{"y": margin_median,
             "label": f"Median {margin_median:.0f}% (1947–present)"}],
)
ax2.plot(plot_df["EV_GVA"]["Date"], plot_df["EV_GVA"]["EV_GVA"],
         color="#1f3b73", linewidth=1.8, label="EV / GVA (4q-avg GVA)")
ax2.set_xlim(*XLIM)
style_macro_chart(
    ax2,
    title="EV / gross value added — EV/EBITDA = (EV/GVA) ÷ (EBITDA/GVA)",
    ylabel="Ratio (x)",
    recessions=recessions,
    hlines=[{"y": ev_gva_median,
             "label": f"Median {ev_gva_median:.1f}x (1947–present)"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "ebitda_margin_ev_gva.png", dpi=150)
plt.close(fig)

# --- Chart 5: intangible share of nonresidential fixed investment ---
intang_median = df["INTANG_SHARE_4Q"].median()
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["INTANG_SHARE"], color="#9ec5e8", linewidth=0.9,
        label="Quarterly")
ax.plot(df["Date"], df["INTANG_SHARE_4Q"], color="#1f3b73", linewidth=1.8,
        label="4-quarter mean")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Intangible (IPP) share of nonresidential fixed investment, 1947–present",
    ylabel="Percent of PNFI",
    recessions=recessions,
    hlines=[{"y": intang_median,
             "label": f"Median {intang_median:.0f}% (1947–present)"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "intangible_share.png", dpi=150)
plt.close(fig)


# --- Overview: panels 1-2-4-5 side by side (panel 4 = margin leg; the EV/GVA
# leg lives in the dedicated chart). This is the chart to read next to the
# aggregate EV/EBITDA history. Built for both the full and 2006+ windows. ---
def build_overview(frame, rec_periods, xlim, window_label, out_name):
    """Overview builder — the ONE function in this script, because the full and
    2006+ companions are identical except data window and reference medians
    (recomputed per charted window, per repo convention)."""
    fig, axes = plt.subplots(2, 2, figsize=(11, 9), sharex=True)
    panels = [
        (axes[0, 0], "DA_SHARE_4Q", "D&A share of EBITDA (4q mean)",
         "% of EBITDA"),
        (axes[0, 1], "EFF_TAX_4Q", "Effective federal corp tax rate (4q mean)",
         "%"),
        (axes[1, 0], "EBITDA_MARGIN", "EBITDA margin on GVA (4q avgs)",
         "% of GVA"),
        (axes[1, 1], "INTANG_SHARE_4Q", "Intangible share of PNFI (4q mean)",
         "% of PNFI"),
    ]
    for ax, col, title, ylabel in panels:
        seg = frame[["Date", col]].dropna()
        med = seg[col].median()
        ax.plot(seg["Date"], seg[col], color="#1f3b73", linewidth=1.5,
                label=title.split(" (")[0])
        if col == "EFF_TAX_4Q":
            for x, _lbl in TAX_VLINES:
                if x >= xlim[0]:
                    ax.axvline(x, color="#c0392b", linestyle=":", linewidth=1.0)
        ax.set_xlim(*xlim)
        style_macro_chart(
            ax,
            title=title,
            ylabel=ylabel,
            recessions=rec_periods,
            hlines=[{"y": med, "label": f"Median {med:.0f} ({window_label})"}],
        )
        ax.title.set_fontsize(10)
        ax.legend(loc="best", frameon=False, fontsize=7)
    fig.suptitle(
        f"Why the aggregate EV multiple sits where it does — drivers, {window_label}",
        fontsize=12,
    )
    fig.tight_layout(rect=(0, 0, 1, 0.97))
    fig.savefig(OUT_DIR / out_name, dpi=150)
    plt.close(fig)


build_overview(df, recessions, XLIM, "1947–present", "drivers_overview.png")

df_2006 = df[df["Date"] >= pd.Timestamp(chart_start_2006)].reset_index(drop=True)
recessions_2006 = [r for r in recessions if r[1] >= pd.Timestamp(chart_start_2006)]
build_overview(
    df_2006, recessions_2006,
    (pd.Timestamp(chart_start_2006), pd.Timestamp(end)),
    "2006–present", "drivers_overview_2006.png",
)

# --- Print summary ---
latest = {c: df[c].dropna().iloc[-1] for c in DERIVED_COLS}
latest_q = df.dropna(subset=["EBIT_YIELD"])["Date"].iloc[-1]
mid80s = df[(df["Date"] >= "1984-01-01") & (df["Date"] <= "1986-12-31")]["EFF_TAX_4Q"].mean()
fifties = df[(df["Date"] >= "1950-01-01") & (df["Date"] <= "1959-12-31")]
spread_pos_share = (df["EBIT_SPREAD"].dropna() > 0).mean() * 100
print(f"Start date:                 {chart_start.date()}")
print(f"End date:                   {end.date()}")
print(f"Rows (1947Q1+):             {len(df)}")
print(f"Latest EV quarter:          {latest_q.date()}")
print(f"Latest EV:                  ${df['EV'].dropna().iloc[-1]:,.0f}B")
print(f"D&A share of EBITDA (4q):   {latest['DA_SHARE_4Q']:.1f}% "
      f"(median {da_median:.1f}%, 1947 first: "
      f"{df['DA_SHARE_4Q'].dropna().iloc[0]:.1f}%)")
print(f"Eff fed tax rate (4q):      {latest['EFF_TAX_4Q']:.1f}% "
      f"(1950s avg {fifties['EFF_TAX_4Q'].mean():.1f}%, "
      f"1984-86 avg {mid80s:.1f}%)")
print(f"EBIT yield:                 {latest['EBIT_YIELD']:.2f}%")
print(f"EBITDA yield:               {latest['EBITDA_YIELD']:.2f}%")
print(f"Real 10y (spliced):         {latest['REAL10Y']:.2f}%")
print(f"EBIT spread vs real 10y:    {latest['EBIT_SPREAD']:.2f}pp "
      f"(positive {spread_pos_share:.0f}% of history)")
print(f"EBITDA margin on GVA:       {latest['EBITDA_MARGIN']:.1f}% "
      f"(median {margin_median:.1f}%)")
print(f"EV/GVA:                     {latest['EV_GVA']:.2f}x "
      f"(median {ev_gva_median:.2f}x)")
print(f"Intangible share of PNFI:   {latest['INTANG_SHARE_4Q']:.1f}% "
      f"(1950s avg {fifties['INTANG_SHARE_4Q'].mean():.1f}%)")
print(f"Wrote {xlsx_path.name}, da_share.png, effective_tax_rate.png, "
      f"earnings_yield_vs_real10y.png, ebitda_margin_ev_gva.png, "
      f"intangible_share.png, drivers_overview.png, drivers_overview_2006.png "
      f"to {OUT_DIR}")
